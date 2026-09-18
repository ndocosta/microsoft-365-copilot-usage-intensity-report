# Security and privacy

## Least-privilege design

Both deployment modes use application-only authentication and the same runtime
Microsoft Graph permissions:

| Permission | Data or action |
| --- | --- |
| `Reports.Read.All` | Microsoft 365 Copilot aggregate and per-user usage reports |
| `User.Read.All` | Entra user and organizational properties |
| `Lists.SelectedOperations.Selected` | Access to explicitly selected lists or libraries |

The runtime identity also receives `write` on one document library. It is not
granted tenant-wide SharePoint access. Azure Automation uses a system-assigned
managed identity; local execution uses a non-exportable certificate.

Provisioning needs broader delegated administrator permissions to create the
identity, assign Graph roles, and create the resource-specific SharePoint grant.
Those permissions belong to the interactive administrator session and are not
assigned to the runtime identity.

## Data collected

Usage snapshots can include:

- user principal name and report display name;
- prompts submitted and active usage days;
- last-activity dates for Copilot experiences and Microsoft 365 apps;
- adoption segment.

Directory snapshots can include:

- display name and user principal name;
- department, job title, office, city, state, country, and usage location;
- company, account state, and user type.

Treat these files as personal and organizational data.

## Recommended controls

- Limit SharePoint membership to approved report owners and readers.
- Apply an appropriate sensitivity label and retention policy.
- Review whether all exported Entra properties are necessary for your purpose.
- Restrict access to the Power BI workspace and semantic model.
- Monitor failed jobs and unexpected changes in export volume.
- Review managed-identity or app-role assignments periodically.
- Remove old snapshots according to your organization's retention policy.
- Never commit generated CSVs, logs, local configuration, or credentials.

## Concealed names

Microsoft 365 can conceal names in usage reports. When enabled, report
identifiers are pseudonymized and cannot be reliably joined to Entra profiles.
Decide whether identifiable analytics is lawful and appropriate before disabling
that privacy control.

## Authentication material

Azure Automation stores no secret for this solution. Its system-assigned
identity is tied to the Automation account and is removed with that resource.

The local option stores the certificate private key in the current Windows
user's certificate store. The generated configuration contains only identifiers
and paths. Protect the Windows profile and rotate the certificate before expiry.

## Responsibility

The repository provides implementation guidance, not legal or compliance
advice. The deploying organization remains responsible for purpose limitation,
transparency, access control, retention, licensing, and applicable employment
and privacy requirements.
