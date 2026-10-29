# Windows File Server → SharePoint Online | PowerShell + 10 Workers

**Status:** Reference implementation, not yet tested against a live tenant. Test in an isolated non-production site with representative NTFS ACLs before production. Source is read-only; destination is modified. The code is not equivalent to Microsoft's migration engine.

## Functionality

- Prompts for an on-premises local/UNC folder and a **clean** SharePoint Online URL for a site, subsite, library, or folder inside a library; if given a site URL, prompts for its document library.
- Creates a **new root folder named after the source folder** under the selected library/folder; recreates nested folders, including empty ones.
- Scans items and potential SharePoint invalid paths/oversized files before touching destination content.
- Optional, explicit mapping of **effective NTFS folder Allow** rules to SharePoint `Read`, `Contribute` or `Full Control` via CSV, with a parent-first application pass. Changes only source-folder destination objects, not the library root permissions.
- Uses up to **ten real PowerShell thread-job workers**, each connecting to SharePoint using its own PnP session; chunked upload is delegated to PnP `Add-PnPFile`.
- Checks destination file existence/size after upload, retries transient failures with exponential backoff, writes per-worker JSONL journal logs, and supports `-Resume` to skip previously verified files where file size and modification timestamp have not changed.
- Fails on unsupported/unmapped **folder** permissions if `-ApplyFolderPermissions` is enabled. Optional `-AuditFileAcls` detects file-level ACL differences and fails closed.
- Does NOT delete or rename source items or automatically switch users over.

## One-time prerequisites

1. Windows file server or Windows management machine with SMB read access to full source tree and ACLs. Use a **read-only service account** for source access if practical; `Get-Acl` requires reading security descriptors. Confirm all shares, underlying NTFS and effective SMB share permissions.
2. PowerShell 7.4+ (run `pwsh`, not Windows PowerShell 5.1), PnP.PowerShell stable, and ThreadJob:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   Install-Module ThreadJob -Scope CurrentUser
   ```
3. Register a **single-tenant Entra application** in the destination tenant. Add **SharePoint** Application permission `Sites.Selected`, grant tenant admin consent, then separately grant the application **FullControl** on the intended site collection (requires a privileged operator; see [Grant-PnPEntraIDAppSitePermission](https://pnp.github.io/powershell/cmdlets/Grant-PnPEntraIDAppSitePermission.html)). For a proof of concept, evaluate whether narrower rights suffice; setting custom folder permissions generally needs FullControl at the site permission grant.
4. Upload the certificate public key `.cer` to the app registration; install its private key in Windows **CurrentUser\My** or **LocalMachine\My** with access restricted to the operator; supply the thumbprint. **Never commit `.pfx`, private keys, credentials, or journals containing sensitive paths.** Each worker has an independent application connection. Rotate/revoke after project and review Entra sign-in logs.
5. Create a SharePoint site, a document library and sufficient tenant/site storage. Restrict site/library access to your migration/validation team **before uploads**. Review versioning, major/minor drafts, check-out, sensitivity labels, retention/DLP, malware scanning, metadata columns, approval, and external sharing. These can block uploads or expose content despite folder ACL changes.
6. Build and **approve** `samples/identity-map.csv`, ideally mapping source **AD security groups** to preprovisioned **SharePoint groups** or mapped Entra group claims supported by `-User`; this sample uses SharePoint groups. `TargetType=Skip` is a SECURITY EXCEPTION only for identities explicitly approved by the destination security owner. Ensure existing target users/groups have correct memberships; the script DOES NOT create/synchronize groups or members.
7. Place the project **outside** the source directory, so output files do not get inventoried/uploaded.

## Commands (terminal)

Inspect first (no SharePoint changes):

```powershell
pwsh -File .\Migrate-FileShareToSPO.ps1 `
  -SourcePath '\\fileserver\Department' `
  -TargetUrl 'https://contoso.sharepoint.com/sites/Finance' `
  -Library 'Documents' `
  -AppId '<app-client-id>' -Tenant 'contoso.onmicrosoft.com' `
  -CertificateThumbprint '<certificate-thumbprint>' -ScanOnly
```

With 10 workers **and** folder ACL mapping (only after the preflight passes):

```powershell
pwsh -File .\Migrate-FileShareToSPO.ps1 `
  -SourcePath '\\fileserver\Department' `
  -TargetUrl 'https://contoso.sharepoint.com/sites/Finance/Shared%20Documents' `
  -AppId '<app-client-id>' -Tenant 'contoso.onmicrosoft.com' `
  -CertificateThumbprint '<certificate-thumbprint>' `
  -MappingCsv '.\samples\identity-map.csv' -Workers 10 `
  -ApplyFolderPermissions -AuditFileAcls
```

If interrupted or failures are fixed, run exactly the same command with `-Resume`. Avoid mixing multiple runs using the same output directory concurrently. To intentionally upload **without** NTFS ACL mapping, specify `-ContentOnly`. Without `-ApplyFolderPermissions`, files inherit destination library/site permissions: do not select this mode for confidential source data unless explicitly authorized. The script refuses to upload unless one of these two modes is explicitly selected.

## Permission model and non-migrated data

| Source (NTFS) | Target (SharePoint) | Caveat |
|---|---|---|
| Read / Read and execute | Read | Not the same effective NTFS privileges |
| Write / Modify | Contribute | Not identical; SharePoint may allow actions not possible with NTFS |
| Full Control | Full Control | Not identical; carefully limit admins |
| Domain user | User UPN | Source-to-target identity mapping required |
| Domain group | SharePoint group / supported Entra principal | Group and membership must already exist |
| Inherited folder ACL | SharePoint inheritance if normalized effective grants match | May differ with inheritance-only/special ACE flags |
| Explicit Deny / special ACE | BLOCKED | Review / redesign manually; explicit deny is not preserved |
| File-specific ACLs | **NOT migrated** | Run `-AuditFileAcls` to stop if different; use SPMT/advanced tool if needed |
| SMB share-level ACL / audit SACL | **NOT migrated** | Evaluate effective SMB + NTFS access manually |
| Owner, timestamps, versions, ADS, EFS | **NOT fully migrated** | Created/Modified by and exact source metadata not preserved by this script |
| NTFS links, junctions, reparse points | BLOCKED | No symbolic link migration |

`-ApplyFolderPermissions` does **not** assume an identity is equivalent based on name. Users and groups must be explicitly mapped, and source Deny entries stop that operation. **Review source folder security with a security owner.** SharePoint/M365 group-connected Teams sites may have membership-based access superseding the folder-level intent; design before migration. Group nesting, side effects of limited access, shared links, alternate access paths and broad site-admin permissions require review.

## Checks, limitations, and operational notes

- SharePoint file upload: **250 GB** per file; decoded site/library/folder/file path **400 characters**; no automatic rename to avoid collisions. File/folder names containing illegal characters or trailing spaces/dots are blocked by preflight. Some additional SharePoint-specific reserved names, content types or tenant policies may still fail server-side.
- **Unique permission scopes:** recommended <5,000 (hard supported limit 50,000); script blocks >5,000 *planned* folder scopes. Counts existing destination scopes separately.
- **Library size:** 30 million total files and folders supported; 100,000-item inheritance restrictions can block permission changes; 5,000 item list-view threshold affects how views/filters are designed.
- Uploads may create a **new version** for an existing target file. Fresh runs refuse to touch an existing source-root folder; `-Resume` is required and can overwrite/version files not in the valid journal. Existing colliding folder names may still contain unrelated data: inspect before approval.
- A worker confirms remote **existence and byte length**, **not SHA-256 equivalence**. For strict integrity, compare separately downloaded file checksums or use a migration platform with stronger validation. Journals can be stale if someone modifies/deletes files in SharePoint after upload. The operator must verify target counts, totals, critical file open/read, permissions and business owner acceptance before cutover.
- No delta reconciliation for **deletes/renames**: pre-stage plus final pass captures new/changed source files if timestamp/size differ but does not delete destination files that disappeared from source. A changed file retaining both size and timestamp can be missed.
- This is not a VSS snapshot or transactional migration. Open or changing source files can change during the upload. Use a **read-only cutover window** on the source and re-run with `-Resume` to synchronize updated file metadata where possible.
- Large migrations: do **not** assume 10 workers is always optimal; HTTP 429/503 require honoring `Retry-After` (PnP may not surface it); lower `-Workers` to 2–4 on sustained throttling. Tune off-hours. PnP may not always expose HTTP response headers reliably; this implementation uses conservative fallbacks.
- Plan for antivirus/file-open locks, small-file overhead, bandwidth/latency, sparse files, names with `%`/`#`, cross-boundary content, compliance labels, retention and managed metadata.
- **No DNS cutover** is required for SharePoint-hosted document libraries unless you're moving application endpoints, vanity URLs, or legacy SMB referrals. **Drive mappings, UNC references, Office templates, scripts, integrations and Excel external links do not automatically rewrite to SharePoint URLs**; plan replacement and user communications.

## Recommended production migration runbook

1. **Discovery** — inventory folders/files/size/owners, effective NTFS and SMB permissions, principal mappings, special ACLs, stubs, previous versions, downstream workflows and business dependencies. Scan migration eligibility.
2. **Target build** — provision users/Entra groups, destination SharePoint sites/libraries, storage and compliance. Restrict the destination to migration administrators. Decide how to handle sensitive folders and Teams membership.
3. **Pilot** — test representative filenames, nested paths, ACLs and large files; measure upload throughput and compare downloaded hashes for representative data.
4. **Pre-stage** — launch controlled transfer during normal operations, not yet exposed to business users. Source stays read-write. Validate structure and scope.
5. **Cutover notice** — notify users; set the source SMB share **read-only** or take equivalent application-coordinated snapshot/quiesce, freeze edits and capture final inventory. Record the time.
6. **Final pass** — `-Resume` with a frozen source; check newly added/modified files and separately reconcile removed/renamed files; process any skipped/incompatible items; complete independent target data/ACL verification.
7. **Switch** — update links, Teams navigation, user shortcuts, application configurations and SMB mappings. Grant end-user access after validation and owner signoff. Keep original server available read-only as a rollback reference.
8. **Hypercare / rollback** — monitor failed access and broken links; for rollback re-enable original SMB access (if it is still authoritative) and communicate the switch. Cloud-only edits after cutover must be accounted for to avoid loss during rollback.

**Downtime:** Pre-staging can run without downtime but user edits during transfer require delta reconciliation. Read-only final cutover is recommended: **duration is workload dependent**, determined from a pilot, final changed bytes/count, throttling and validation. It is NOT possible to promise zero downtime or a fixed migration speed. Reverting users to SMB is a routing decision, not automatic restoration of cloud-side edits.

## Enterprise-native alternative: Microsoft SPMT / Migration Manager

Microsoft's [free SPMT](https://learn.microsoft.com/sharepointmigration/introducing-the-sharepoint-migration-tool) and Migration Manager are usually **preferable for production file-share migration** because they have migration scanning, reports, identity mapping, bulk migration architecture, delta passes and supported permission handling. See [supported features](https://learn.microsoft.com/sharepointmigration/what-is-supported-spmt), [file-share guide](https://learn.microsoft.com/sharepointmigration/fileshare-to-odsp-migration-guide) and [settings](https://learn.microsoft.com/sharepointmigration/spmt-settings). Even SPMT doesn't preserve Deny, all NTFS advanced rights or complete file metadata. The PowerShell project is useful for custom automation, controlled smaller migrations, and showcasing technical architecture; use a hardened migration platform for high-value/large-scale data or file-level ACL requirements.

## Citations and reference docs

- [PnP installation](https://pnp.github.io/powershell/articles/installation.html) | [certificate authentication](https://pnp.github.io/powershell/articles/authentication.html) | [PnP file upload](https://pnp.github.io/powershell/cmdlets/Add-PnPFile.html) | [PnP folder permissions](https://pnp.github.io/powershell/cmdlets/Set-PnPFolderPermission.html)
- [SharePoint limits](https://learn.microsoft.com/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits) | [storage planning](https://learn.microsoft.com/sharepoint/sharepoint-storage-planning) | [throttling](https://learn.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online)
- [SPMT file-share permission model](https://learn.microsoft.com/sharepointmigration/understanding-permissions-when-migrating) | [Microsoft file-share migration guide](https://learn.microsoft.com/sharepointmigration/fileshare-to-odsp-migration-guide)
- [Current Canadian Microsoft 365 business pricing](https://www.microsoft.com/en-ca/microsoft-365/business/compare-all-microsoft-365-business-products-b?market=ca). Confirm current offer with your Microsoft account/reseller and tax jurisdiction.

## Security

Treat mapping CSV, inventory.csv and output journals as sensitive. `.gitignore` excludes output, credentials and certificates. Never commit client certificates. Never use blanket tenant-wide `Sites.FullControl.All` if a restricted `Sites.Selected` grant works. If the staging site is already accessible to users, **do not run a confidential-content migration into it**.

See [licensing, costs and limits](docs/LICENSING-AND-LIMITS.md) for Canadian pricing, specific tenant quotas, on-prem prerequisites, and cost assumptions.