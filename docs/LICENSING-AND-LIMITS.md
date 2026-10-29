# Licensing, costs, limits and infrastructure prerequisites

Reviewed October 29, 2026. All example prices are CAD before taxes; confirm current billing and reseller quotes.

## Licensing and cost

| Component | Example cost | Licensing scope |
|---|---|---|
| SharePoint Online Plan 1 | CAD $6.80/user/month, annual billing | Per accessing end user |
| Microsoft 365 Business Basic, with Teams | CAD $9.50/user/month, annual billing | Per user; includes SharePoint |
| Microsoft 365 Business Basic, without Teams | CAD $7.30/user/month, annual billing | Per user; where offered |
| Microsoft 365 E3/E5, Office 365 E1/E3/E5 | SharePoint included in eligible assigned plan | Per user |
| Microsoft SharePoint Migration Tool (SPMT) | Free | No separate migration license |
| PnP.PowerShell and ThreadJob | Free | No separate license |
| Entra application with Sites.Selected | No separate registration fee | Per application, with per-site grant |
| SharePoint storage | Normally 1 TB + 10 GB per eligible user license | Tenant-wide pooled quota; excess charged |
| Windows/SMB source | Existing licensing obligations | Server/workstation as licensed |
| Backup, professional support and storage add-ons | Variable | Quote separately |

One licensed Microsoft 365 user does not license every user in the tenant. A migration operator does not need a separate paid migration-tool license when using these free tools. Business plans normally have a 300-user limit.

## Platform and source restrictions

- Maximum file size: 250 GB. Maximum decoded SharePoint URL path: 400 characters.
- Library limit: up to 30 million files and folders; site collection can be limited to 25 TB and tenant storage allocation.
- Unique security scopes: 50,000 is the supported limit, while fewer than 5,000 are recommended; 100,000-item permission inheritance restrictions can apply.
- HTTP 429 and 503 throttling, sometimes with Retry-After, may require reducing the number of concurrent workers.
- Permissions are NOT a perfect NTFS-to-SharePoint translation. NTFS Deny entries, advanced ACLs, file-level ACLs, SMB share permissions, file ownership, auditing, and membership cannot be fully recreated by this script.
- The custom script does not preserve old file versions, historical Created By / Modified By metadata, NTFS alternate streams, or full original timestamps. It does not automatically reconcile deleted or renamed source items.
- Check versioning, retention, DLP, sensitivity labels, required metadata, checkout rules, antivirus and external sharing policies before migration.

## Prerequisites and cutover

Use PowerShell 7.4+ on Windows, PnP.PowerShell, ThreadJob, a certificate-authenticated Entra application with approved SharePoint Sites.Selected site access, and read access to SMB source files and ACLs. Provision the site, library, storage, identities and mapped SharePoint group memberships before upload.

Pre-stage the files while users continue working, then freeze on-premises writes and perform final upload/delta verification. Validate effective permissions and file inventories before changing users to SharePoint URLs. Do not delete the on-premises source until business acceptance and backup obligations are met.

SharePoint Online migrations do not generally require a DNS cutover. Replace UNC paths and legacy SMB integrations as a separate application change.

## Microsoft documentation

- https://learn.microsoft.com/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits
- https://learn.microsoft.com/sharepoint/sharepoint-storage-planning
- https://learn.microsoft.com/sharepointmigration/fileshare-to-odsp-migration-guide
- https://learn.microsoft.com/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online
- https://www.microsoft.com/en-ca/microsoft-365/business/compare-all-microsoft-365-business-products-b
