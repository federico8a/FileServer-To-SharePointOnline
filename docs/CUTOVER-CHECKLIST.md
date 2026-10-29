# Cutover and rollback checklist

| Phase | Required evidence | Go / no-go |
|---|---|---|
| Preflight | `inventory.csv` and `preflight-issues.csv`, source access tested | Zero blockers |
| Identity mapping | CSV approval, user/group existence, effective ACL spot-check, ownership | Security owner signoff |
| Pilot | A representative read/write/ACL/performance pilot | Business owner signs off |
| Storage | Site/tenant capacity, retention effects, versioning and storage headroom | Capacity confirmed |
| Staging | Restricted SharePoint access, no end-user exposure | Security gate |
| Source freeze | Written source read-only time, open files checked, applications coordinated | Cutover starts |
| Last delta | Re-run with `-Resume`; compare source/destination inventory and deletions | Zero unresolved mismatches |
| Permissions | Test user access (positive AND negative tests) | Least privilege confirmed |
| Functional | Critical files open, links and integrations reviewed | Owner acceptance |
| Go-live | Replace UNC references, publish SharePoint links, user documentation | Change approved |
| Rollback | Repoint to original read-only server if needed; manage post-cutover cloud edits | Rollback owner designated |

## Risk register

- High: unresolved source NTFS Deny, file-specific ACLs, SMB share permissions -> wrong access after migration.
- High: staging library available to broad site membership -> accidental unauthorized access.
- High: no final delta/reconciliation after source freeze -> missed changes/deletions.
- Medium: versioning/retention multiplies storage and upload overwrite effects.
- Medium: 429/503 and transient HTTP errors slow final cutover.
- Medium: blocked names, path length, old file locks, malware/DLP policies.
- Medium: legacy apps relying on UNC paths, SMB locking, executable content or custom ACLs cannot run unmodified from SharePoint.

Do not retire or delete the on-premises source until retention/restore obligations and business acceptance have been met.
