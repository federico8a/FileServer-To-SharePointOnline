#requires -Version 7.4
<#
.SYNOPSIS
  Migrates an NTFS directory tree to a SharePoint Online document library with 10 upload workers.
.DESCRIPTION
  PowerShell 7.4+, Windows, PnP.PowerShell, ThreadJob. Certificate-based authentication.
  Interactive source/target prompts. Preflight assessment, folder ACL mapping, parent-first
  preparation, 10 concurrent independent connections, remote existence/size checks, resume logs.
  Do not use for NTFS Deny, file-level ACL, special permissions, or exact NTFS fidelity.
.EXAMPLE
  .\Migrate-FileShareToSPO.ps1 -ScanOnly -AppId '...' -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint '...'
.EXAMPLE
  .\Migrate-FileShareToSPO.ps1 -ApplyFolderPermissions -AuditFileAcls -AppId '...' -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint '...' -MappingCsv .\samples\identity-map.csv
#>
[CmdletBinding()]
param(
    [string]$SourcePath,
    [string]$TargetUrl,
    [string]$Library,
    [string]$AppId,
    [string]$Tenant,
    [string]$CertificateThumbprint,
    [string]$MappingCsv,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'output'),
    [ValidateRange(1,10)][int]$Workers = 10,
    [switch]$ScanOnly,
    [switch]$ApplyFolderPermissions,
    [switch]$ContentOnly,
    [switch]$AuditFileAcls,
    [switch]$Resume,
    [switch]$SkipConfirmation
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'NTFS ACL migration requires Windows PowerShell 7.4+ on a Windows host.' }
Import-Module PnP.PowerShell -ErrorAction Stop
Import-Module ThreadJob -ErrorAction Stop

function Prompt-Required([string]$Title, [string]$Existing) {
    if (-not [string]::IsNullOrWhiteSpace($Existing)) { return $Existing.Trim() }
    do { $response = Read-Host $Title } until (-not [string]::IsNullOrWhiteSpace($response))
    return $response.Trim()
}
function Normalize-UrlPath([string]$Path) {
    return ([Uri]::UnescapeDataString($Path)).TrimEnd('/')
}
function RelPath([string]$FullPath) {
    return [System.IO.Path]::GetRelativePath($script:SourceRoot, $FullPath).Replace('\','/')
}
function Join-WebPath([string]$Left, [string]$Right) {
    return $Left.TrimEnd('/') + '/' + $Right.TrimStart('/')
}
function Is-ValidName([string]$Name) {
    if ($Name -match '["*:<>?/\\|]' -or $Name -match '[\. ]$') { return $false }
    if ($Name -in @('.','..','_private')) { return $false }
    return $true
}
function Connect-App([string]$WebUrl) {
    return Connect-PnPOnline -Url $WebUrl -ClientId $AppId -Tenant $Tenant -Thumbprint $CertificateThumbprint -ReturnConnection -ErrorAction Stop
}
function Find-Destination([string]$InputUrl) {
    $uri = [Uri]$InputUrl
    if ($uri.Scheme -ne 'https' -or $uri.Host -notlike '*.sharepoint.com') {
        throw 'Use a clean HTTPS SharePoint Online site/subsite/library URL.'
    }
    if ($uri.Query -or $uri.Fragment -or $uri.AbsolutePath -match '(?i)\.aspx$') {
        throw 'Use the clean site/subsite/library URL, not an AllItems.aspx view or sharing link.'
    }
    $decoded = Normalize-UrlPath $uri.AbsolutePath
    $parts = @($decoded.Trim('/').Split('/', [StringSplitOptions]::RemoveEmptyEntries))
    $origin = $uri.GetLeftPart([UriPartial]::Authority)
    $found = $null
    for ($count=$parts.Count; $count -ge 0; $count--) {
        $candidate = if ($count -eq 0) { $origin } else { $origin + '/' + ($parts[0..($count-1)] -join '/') }
        try {
            $connection = Connect-App $candidate
            $web = Get-PnPWeb -Includes Url,ServerRelativeUrl -Connection $connection
            if (([Uri]$web.Url).AbsolutePath.TrimEnd('/') -ieq ([Uri]$candidate).AbsolutePath.TrimEnd('/')) {
                $found = [pscustomobject]@{ Connection=$connection; Web=$web; Url=$candidate }
                break
            }
        } catch { continue }
    }
    if (-not $found) { throw 'Could not authenticate to a matching site/subsite. Check URL, app certificate, Sites.Selected grant and site existence.' }
    $lists = @(Get-PnPList -Includes RootFolder,BaseTemplate,Title,Hidden -Connection $found.Connection | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden })
    $selection = $null
    $destinationPath = $null
    if ($decoded -ne (Normalize-UrlPath (([Uri]$found.Web.Url).AbsolutePath))) {
        $selection = $lists | Where-Object {
            $root = Normalize-UrlPath $_.RootFolder.ServerRelativeUrl
            $decoded -ieq $root -or $decoded.StartsWith($root + '/', [StringComparison]::OrdinalIgnoreCase)
        } | Sort-Object { $_.RootFolder.ServerRelativeUrl.Length } -Descending | Select-Object -First 1
        if (-not $selection) { throw 'URL is inside this site but does not match a document library. Provide a site URL or an existing library/folder URL.' }
        $destinationPath = $decoded
    } else {
        $names = ($lists | Select-Object -ExpandProperty Title) -join ', '
        $chosen = Prompt-Required "Document library title or root URL (available: $names)" $Library
        $selection = $lists | Where-Object {
            $_.Title -ieq $chosen -or (Normalize-UrlPath $_.RootFolder.ServerRelativeUrl) -ieq (Normalize-UrlPath ([Uri]::UnescapeDataString($chosen)))
        } | Select-Object -First 1
        if (-not $selection) { throw "Library '$chosen' was not found in the supplied site/subsite." }
        $destinationPath = Normalize-UrlPath $selection.RootFolder.ServerRelativeUrl
    }
    return [pscustomobject]@{
        WebUrl=$found.Web.Url; WebRelativePath=$found.Web.ServerRelativeUrl;
        Connection=$found.Connection; LibraryTitle=$selection.Title;
        LibraryRoot=(Normalize-UrlPath $selection.RootFolder.ServerRelativeUrl);
        DestinationPath=$destinationPath
    }
}
function Get-AccessMap([string]$Path, [hashtable]$Mapping) {
    $acl = Get-Acl -LiteralPath $Path
    $roles = @{}
    $issues = [System.Collections.Generic.List[string]]::new()
    foreach ($ace in $acl.Access) {
        $who = [string]$ace.IdentityReference.Value
        if ($ace.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Deny) {
            $issues.Add("Deny ACE: $who on $Path"); continue
        }
        if (-not $Mapping.ContainsKey($who.ToLowerInvariant())) {
            $issues.Add("Unmapped identity: $who on $Path"); continue
        }
        $mapped = $Mapping[$who.ToLowerInvariant()]
        if ($mapped.Type -eq 'Skip') { continue } # Only after a security administrator has approved the Skip rows.
        $rights = [System.Security.AccessControl.FileSystemRights]$ace.FileSystemRights
        $rank = 0
        $full = [System.Security.AccessControl.FileSystemRights]::FullControl
        $modify = [System.Security.AccessControl.FileSystemRights]::Modify
        $write = [System.Security.AccessControl.FileSystemRights]::Write
        $read = [System.Security.AccessControl.FileSystemRights]::Read
        if (($rights -band $full) -eq $full) { $rank = 3 }
        elseif (($rights -band $modify) -eq $modify -or ($rights -band $write) -eq $write) { $rank = 2 }
        elseif (($rights -band $read) -eq $read -or ($rights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) -eq [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) { $rank = 1 }
        else { $issues.Add("Special/unmappable NTFS rights '$rights': $who on $Path"); continue }
        $key = ($mapped.Type + '|' + $mapped.Principal).ToLowerInvariant()
        if (-not $roles.ContainsKey($key) -or $roles[$key].Rank -lt $rank) {
            $roles[$key] = [pscustomobject]@{ Type=$mapped.Type; Principal=$mapped.Principal; Rank=$rank; Role=@('','Read','Contribute','Full Control')[$rank] }
        }
    }
    $entries = @($roles.Values | Sort-Object Type,Principal)
    $signature = (@($entries | ForEach-Object { $_.Type + '|' + $_.Principal + '|' + $_.Role }) -join ';')
    return [pscustomobject]@{ Entries=$entries; Signature=$signature; Issues=@($issues) }
}
function Get-Mapping([string]$File) {
    $mapping = @{}
    if (-not (Test-Path -LiteralPath $File -PathType Leaf)) { throw 'ApplyFolderPermissions requires an existing mapping CSV.' }
    foreach ($r in @(Import-Csv -LiteralPath $File)) {
        if ([string]::IsNullOrWhiteSpace($r.SourceIdentity) -or $r.TargetType -notin @('User','Group','Skip')) { throw "Bad identity mapping row: $($r.SourceIdentity)" }
        if ($r.TargetType -ne 'Skip' -and [string]::IsNullOrWhiteSpace($r.TargetPrincipal)) { throw "Missing target principal: $($r.SourceIdentity)" }
        $key = $r.SourceIdentity.Trim().ToLowerInvariant()
        if ($mapping.ContainsKey($key)) { throw "Duplicate source identity mapping: $key" }
        $mapping[$key] = [pscustomobject]@{ Type=$r.TargetType; Principal=$r.TargetPrincipal.Trim() }
    }
    return $mapping
}
function Save-Csv([object[]]$Objects, [string]$Path) {
    if (@($Objects).Count -gt 0) { $Objects | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 }
}

$SourcePath = Prompt-Required 'On-premises source folder (local or UNC path)' $SourcePath
$TargetUrl = Prompt-Required 'SharePoint site, subsite, library or library-folder HTTPS URL' $TargetUrl
$AppId = Prompt-Required 'Entra application (client) ID' $AppId
$Tenant = Prompt-Required 'Tenant domain (e.g., contoso.onmicrosoft.com)' $Tenant
$CertificateThumbprint = Prompt-Required 'Certificate thumbprint (installed in Windows certificate store)' $CertificateThumbprint
$sourceItem = Get-Item -LiteralPath $SourcePath
if (-not $sourceItem.PSIsContainer) { throw 'Source must be a folder.' }
$script:SourceRoot = $sourceItem.FullName.TrimEnd('\','/')
if (-not (Is-ValidName $sourceItem.Name)) { throw "Invalid SharePoint source root folder name: $($sourceItem.Name)" }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$target = Find-Destination $TargetUrl
$rootServerRelative = Join-WebPath $target.DestinationPath $sourceItem.Name
$siteRelativeRoot = $rootServerRelative.Substring($target.WebRelativePath.TrimEnd('/').Length).TrimStart('/')
Write-Host "Source: $script:SourceRoot"
Write-Host "Target web: $($target.WebUrl)" 
Write-Host "Library: $($target.LibraryTitle)" 
Write-Host "Upload root (inside library): $rootServerRelative" 

$scanErrors = @()
$allItems = @(Get-ChildItem -LiteralPath $script:SourceRoot -Recurse -Force -ErrorAction Continue -ErrorVariable +scanErrors)
if ($scanErrors.Count) { throw "Source enumeration had $($scanErrors.Count) errors. Resolve access and long-path issues before proceeding." }
$directories = @($allItems | Where-Object { $_.PSIsContainer -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) } | Sort-Object { (RelPath $_.FullName).Split('/').Count },FullName)
$files = @($allItems | Where-Object { -not $_.PSIsContainer -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) })
$invalid = [System.Collections.Generic.List[object]]::new()
$records = [System.Collections.Generic.List[object]]::new()
foreach ($entry in $allItems) {
    if ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) { $invalid.Add([pscustomobject]@{Path=$entry.FullName; Reason='Reparse point/symlink unsupported'}) }
}
$allFolders = @([pscustomobject]@{ Rel=''; Full=$script:SourceRoot }) + @($directories | ForEach-Object { [pscustomobject]@{ Rel=(RelPath $_.FullName); Full=$_.FullName } })
foreach ($folder in $allFolders) {
    $serverFolder = if ($folder.Rel -eq '') { $rootServerRelative } else { Join-WebPath $rootServerRelative $folder.Rel }
    if ($serverFolder.Length -gt 400) { $invalid.Add([pscustomobject]@{Path=$folder.Full; Reason='SharePoint path exceeds 400 characters'}) }
    if ($folder.Rel -ne '' -and -not (Is-ValidName ([IO.Path]::GetFileName($folder.Full)))) { $invalid.Add([pscustomobject]@{Path=$folder.Full; Reason='Invalid SharePoint folder name'}) }
}
foreach ($file in $files) {
    $rel = RelPath $file.FullName
    $relParent = if ($rel.Contains('/')) { $rel.Substring(0,$rel.LastIndexOf('/')) } else { '' }
    $destFolder = if ($relParent) { Join-WebPath $rootServerRelative $relParent } else { $rootServerRelative }
    $destFile = Join-WebPath $destFolder $file.Name
    if (-not (Is-ValidName $file.Name)) { $invalid.Add([pscustomobject]@{Path=$file.FullName; Reason='Invalid SharePoint file name'}) }
    if ($destFile.Length -gt 400) { $invalid.Add([pscustomobject]@{Path=$file.FullName; Reason='SharePoint path exceeds 400 characters'}) }
    if ($file.Length -gt 250GB) { $invalid.Add([pscustomobject]@{Path=$file.FullName; Reason='File exceeds 250 GB'}) }
    $records.Add([pscustomobject]@{Path=$file.FullName;Relative=$rel;Folder=$destFolder;Destination=$destFile;Size=[long]$file.Length;LastWriteUtcTicks=[long]$file.LastWriteTimeUtc.Ticks})
}
Save-Csv $records.ToArray() (Join-Path $OutputDirectory 'inventory.csv')
Save-Csv $invalid.ToArray() (Join-Path $OutputDirectory 'preflight-issues.csv')
$sizeTotal = ($records | Measure-Object Size -Sum).Sum
Write-Host "Preflight: $($records.Count) files, $($allFolders.Count) folders (including root), $([math]::Round($sizeTotal/1GB,2)) GiB, $($invalid.Count) issues."
if ($invalid.Count) { throw "Preflight blocked: fix $($invalid.Count) invalid paths/items (see preflight-issues.csv). No data was changed." }

$plans = [System.Collections.Generic.List[object]]::new()
$aclIssues = [System.Collections.Generic.List[object]]::new()
if ($ApplyFolderPermissions) {
    $MappingCsv = Prompt-Required 'CSV mapping file (SourceIdentity,TargetType,TargetPrincipal)' $MappingCsv
    $mapping = Get-Mapping $MappingCsv
    $signatures = @{}
    foreach ($folder in $allFolders) {
        $rights = Get-AccessMap $folder.Full $mapping
        foreach ($issue in $rights.Issues) { $aclIssues.Add([pscustomobject]@{Path=$folder.Full;Issue=$issue}) }
        $parentRel = if ($folder.Rel -and $folder.Rel.Contains('/')) { $folder.Rel.Substring(0,$folder.Rel.LastIndexOf('/')) } else { '' }
        $unique = ($folder.Rel -eq '' -or $rights.Signature -ne $signatures[$parentRel])
        $signatures[$folder.Rel] = $rights.Signature
        if ($rights.Entries.Count -eq 0) { $aclIssues.Add([pscustomobject]@{Path=$folder.Full;Issue='No mapped Allow entries; refusing to set empty ACL'}) }
        $targetFolder = if ($folder.Rel -eq '') { $rootServerRelative } else { Join-WebPath $rootServerRelative $folder.Rel }
        $plans.Add([pscustomobject]@{Local=$folder.Full;Remote=$targetFolder;Unique=$unique;Entries=@($rights.Entries)})
    }
    if ($AuditFileAcls) {
        foreach ($file in $files) {
            $fileAcl = Get-AccessMap $file.FullName $mapping
            foreach ($issue in $fileAcl.Issues) { $aclIssues.Add([pscustomobject]@{Path=$file.FullName;Issue=$issue}) }
            $rel = RelPath $file.DirectoryName
            if ($rel -eq '.') { $rel = '' }
            if ($fileAcl.Signature -ne $signatures[$rel]) {
                $aclIssues.Add([pscustomobject]@{Path=$file.FullName;Issue='File-level permissions differ from folder (unsupported by this script)'})
            }
        }
    } else {
        Write-Warning 'File-level NTFS ACLs are NOT scanned or reproduced. Use -AuditFileAcls to block on differing file permissions.'
    }
    Save-Csv $aclIssues.ToArray() (Join-Path $OutputDirectory 'acl-issues.csv')
    $uniqueCount = @($plans | Where-Object Unique).Count
    Write-Host "Folder ACL: $uniqueCount unique security scopes planned, $($aclIssues.Count) blockers."
    if ($uniqueCount -gt 5000) { throw 'Over 5000 unique folder ACLs: redesign SharePoint permissions before migration.' }
    if ($aclIssues.Count) { throw 'Permissions preflight blocked; resolve all unmapped, Deny, special, or differing file ACL problems first.' }
}
if ($ScanOnly) { Write-Host "SCAN ONLY. See $OutputDirectory. Nothing was uploaded."; return }
if (-not $ApplyFolderPermissions -and -not $ContentOnly) {
    throw 'For data security, explicitly choose -ApplyFolderPermissions (recommended) or -ContentOnly (no ACL migration, SharePoint inheritance applies).'
}
if ($ApplyFolderPermissions -and $ContentOnly) { throw 'Choose either -ApplyFolderPermissions or -ContentOnly, not both.' }
if ($ContentOnly) { Write-Warning 'CONTENT ONLY: NTFS permissions are NOT migrated; every uploaded file inherits the target SharePoint folder/library permissions.' }
if (-not $SkipConfirmation) {
    $approval = Read-Host "Type MIGRATE to create folders and upload (existing destination files may be versioned/overwritten)"
    if ($approval -cne 'MIGRATE') { throw 'Cancelled by operator.' }
}

# A fresh migration uses its own root folder, preventing accidental overwrites or permission changes in the library.
$existing = $null
try { $existing = Get-PnPFolder -Url $rootServerRelative -Connection $target.Connection -ErrorAction SilentlyContinue } catch {}
if ($null -ne $existing -and -not $Resume) { throw 'Destination root already exists. Review the destination; pass -Resume ONLY to continue an earlier migration.' }
if ($null -eq $existing -and $Resume) { throw 'Cannot resume: destination root folder is missing.' }
foreach ($folder in $allFolders) {
    $dest = if ($folder.Rel -eq '') { $rootServerRelative } else { Join-WebPath $rootServerRelative $folder.Rel }
    $siteRelative = $dest.Substring($target.WebRelativePath.TrimEnd('/').Length).TrimStart('/')
    $null = Resolve-PnPFolder -SiteRelativePath $siteRelative -Connection $target.Connection -ErrorAction Stop
}
Write-Host "Created/confirmed $($allFolders.Count) source folders under the target root."

# Fail fast on any ACL error.  This phase is NOT parallel, to preserve parent-first inheritance.
if ($ApplyFolderPermissions) {
    foreach ($plan in $plans) {
        if (-not $plan.Unique) { continue }
        $folderObject = Get-PnPFolder -Url $plan.Remote -Connection $target.Connection -ErrorAction Stop
        $first = $true
        foreach ($grant in $plan.Entries) {
            $args = @{ List=$target.LibraryTitle;Identity=$folderObject;AddRole=$grant.Role;Connection=$target.Connection;ErrorAction='Stop' }
            if ($first) { $args.ClearExisting = $true }
            if ($grant.Type -eq 'Group') { $args.Group = $grant.Principal } else { $args.User = $grant.Principal }
            Set-PnPFolderPermission @args | Out-Null
            $first = $false
        }
    }
    Write-Host 'Applied approved folder ACL mappings.'
}

# Resume is scoped to BOTH source and destination; the logs are an optimization, not proof of cloud integrity.
$journalName = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($script:SourceRoot.ToLowerInvariant()+'|'+$rootServerRelative.ToLowerInvariant()))).Substring(0,20)
$done = @{}
if ($Resume) {
    foreach ($log in @(Get-ChildItem -LiteralPath $OutputDirectory -Filter "journal-$journalName-*.jsonl" -ErrorAction SilentlyContinue)) {
        foreach ($line in [IO.File]::ReadLines($log.FullName)) {
            try {
                $value = $line | ConvertFrom-Json
                $key = "$($value.Destination)|$($value.Size)|$($value.LastWriteUtcTicks)"
                if ($value.Status -eq 'Verified') { $done[$key] = $true }
                elseif ($value.Status -eq 'Failed') { $done.Remove($key) | Out-Null }
            } catch { Write-Warning "Ignoring malformed journal line in $($log.Name)" }
        }
    }
}
$remaining = @($records | Where-Object { -not $done.ContainsKey("$($_.Destination)|$($_.Size)|$($_.LastWriteUtcTicks)") })
Write-Host "Uploading $($remaining.Count) files with a maximum of $Workers independent workers (resumed/skipped: $($records.Count-$remaining.Count))."
if ($remaining.Count -eq 0) { Write-Host 'No files remain according to the journal. Verify in SharePoint before cutover.'; return }
$shards = @()
for ($i=0;$i -lt $Workers;$i++) { $shards += ,([System.Collections.Generic.List[object]]::new()) }
for ($i=0;$i -lt $remaining.Count;$i++) { $shards[$i % $Workers].Add($remaining[$i]) }
$jobs = [System.Collections.Generic.List[object]]::new()
for ($worker=0;$worker -lt $Workers;$worker++) {
    if ($shards[$worker].Count -eq 0) { continue }
    $logPath = Join-Path $OutputDirectory "journal-$journalName-$worker.jsonl"
    $workerSettings = [pscustomobject]@{
        Queue=$shards[$worker].ToArray(); WebUrl=$target.WebUrl; ClientId=$AppId;
        Tenant=$Tenant; Thumbprint=$CertificateThumbprint; JournalPath=$logPath
    }
    $jobs.Add((Start-ThreadJob -Name "Upload-$worker" -ThrottleLimit $Workers -ArgumentList $workerSettings -ScriptBlock {
        param($settings)
        $ErrorActionPreference = 'Stop'
        Import-Module PnP.PowerShell -ErrorAction Stop
        $connection = Connect-PnPOnline -Url $settings.WebUrl -ClientId $settings.ClientId -Tenant $settings.Tenant -Thumbprint $settings.Thumbprint -ReturnConnection -ErrorAction Stop
        $summary = [pscustomobject]@{Ok=0;Failed=0;Bytes=0L}
        foreach ($file in $settings.Queue) {
            $result = [ordered]@{Destination=$file.Destination;Path=$file.Path;Size=$file.Size;LastWriteUtcTicks=$file.LastWriteUtcTicks;Status='Failed';Error='';Utc=[DateTime]::UtcNow.ToString('o')}
            for ($attempt=1;$attempt -le 7;$attempt++) {
                try {
                    $null = Add-PnPFile -Path $file.Path -Folder $file.Folder -Connection $connection -ErrorAction Stop
                    $remote = Get-PnPFile -Url $file.Destination -AsFileObject -Connection $connection -ErrorAction Stop
                    if ($null -eq $remote) { throw 'Uploaded file not visible in SharePoint verification.' }
                    if ([long]$remote.Length -ne [long]$file.Size) { throw "Size mismatch: local=$($file.Size) remote=$($remote.Length)" }
                    $result.Status = 'Verified'
                    $summary.Ok++; $summary.Bytes += [long]$file.Size
                    break
                } catch {
                    $message = $_.Exception.ToString()
                    $result.Error = $message.Substring(0,[Math]::Min(1500,$message.Length))
                    # Only retry potentially transient failures; permission/invalid filename failures need remediation.
                    $transient = $message -match '(?i)\b(429|503|502|504)\b|throttl|server busy|timed?\s*out|temporar|connection reset|task cancel'
                    if (-not $transient -or $attempt -eq 7) { break }
                    $delay = [int][Math]::Min(120,[Math]::Pow(2,$attempt) + (Get-Random -Minimum 0 -Maximum 5))
                    # PnP exceptions do not always expose response headers. Honor Retry-After when available.
                    try {
                        $h = $_.Exception.Response.Headers
                        if ($h.RetryAfter -and $h.RetryAfter.Delta) {
                            $delay = [Math]::Max($delay,[int][Math]::Ceiling($h.RetryAfter.Delta.TotalSeconds))
                        } elseif ($h['Retry-After']) {
                            $delay = [Math]::Max($delay,[int]$h['Retry-After'])
                        }
                    } catch {}
                    Start-Sleep -Seconds $delay
                }
            }
            if ($result.Status -ne 'Verified') { $summary.Failed++ }
            # Independent journal per worker prevents concurrent writes and supports restart after a crash.
            [IO.File]::AppendAllText($settings.JournalPath, (($result | ConvertTo-Json -Compress) + [Environment]::NewLine), [Text.Encoding]::UTF8)
        }
        return $summary
    }))
}
$jobs | Wait-Job | Out-Null
$jobErrors = [System.Collections.Generic.List[string]]::new()
foreach ($job in $jobs) {
    if ($job.State -ne 'Completed') { $jobErrors.Add("$($job.Name): $($job.State) $($job.ChildJobs[0].JobStateInfo.Reason)") }
}
$receiveErrors = @()
$results = @($jobs | Receive-Job -ErrorAction Continue -ErrorVariable +receiveErrors)
foreach ($item in $receiveErrors) { $jobErrors.Add([string]$item) }
$jobs | Remove-Job -Force
foreach ($e in $jobErrors) { Write-Warning $e }
$totalOk = ($results | Measure-Object Ok -Sum).Sum
$totalFail = ($results | Measure-Object Failed -Sum).Sum
Write-Host "Completed this pass: verified=$totalOk, failed=$totalFail. Output journals: $OutputDirectory"
if ($jobErrors.Count -gt 0 -or $totalFail -gt 0) {
    throw 'Migration incomplete; investigate journal-*.jsonl failures and run again using -Resume. Do NOT cut over.'
}
Write-Host 'Upload pass complete. A size match is not a cryptographic checksum. Perform content, permission, and business acceptance checks before cutover.'