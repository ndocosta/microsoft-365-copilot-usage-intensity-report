# Local Scheduled Task deployment

This option runs the export from Windows using certificate-based application
authentication. It is suitable for evaluation, demos, and small environments.
For a service that must run with no signed-in user, use
[Azure Automation](azure-automation-deployment.md).

## Prerequisites

- Windows PowerShell 5.1
- A Microsoft 365 tenant with licensed Microsoft 365 Copilot users
- A SharePoint site and document library for the CSV snapshots
- Permission to create app registrations and grant admin consent
- `Microsoft.Graph.Authentication`

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

Run setup as the same Windows user that will own the Scheduled Task.

## 1. Create the runtime identity

From the repository root:

```powershell
Set-Location .\Automation

.\New-CopilotUsageApp.ps1 `
    -TenantId 'contoso.onmicrosoft.com' `
    -SharePointSiteUrl 'https://contoso.sharepoint.com/sites/CopilotReports' `
    -SharePointLibraryName 'UsageData'
```

A browser opens for administrator sign-in. The script creates:

- a non-exportable RSA certificate in `Cert:\CurrentUser\My`;
- a single-tenant app registration and service principal;
- the three runtime Microsoft Graph application permissions;
- a `write` grant on the selected document library only;
- `Automation\automation.config.json`.

The administrator temporarily consents to `Application.ReadWrite.All`,
`AppRoleAssignment.ReadWrite.All`, and `Sites.ReadWrite.All` during setup. These
delegated permissions are not assigned to the runtime app.

The configuration file contains identifiers and local paths, but no client
secret or private key. It is environment-specific and excluded from Git.

## 2. Test the complete pipeline

```powershell
.\Invoke-CopilotUsagePipeline.ps1 `
    -ConfigPath .\automation.config.json
```

Confirm that the three snapshot folders and CSV files appear in SharePoint.

## 3. Register the daily task

```powershell
.\Register-CopilotUsageScheduledTask.ps1 `
    -ConfigPath .\automation.config.json `
    -DailyAt '06:00'
```

`DailyAt` uses the computer's local time. The task runs with limited privileges
and can run while the session is locked, but the same Windows user must remain
signed in. `StartWhenAvailable` runs a missed trigger when the computer becomes
available.

## Certificate maintenance

The bootstrap creates a certificate valid for 24 months. Rotate it before
expiry and update both the app registration credential and
`CertificateThumbprint` in `automation.config.json`.

## Troubleshooting

**The task cannot access the certificate**

Confirm that the task runs as the user who created the certificate and that the
certificate has a private key in `Cert:\CurrentUser\My`.

**The library cannot be found**

`SharePointLibraryName` must be the display name of a document library, not a
folder name or URL segment.

**The report returns no recognizable users**

Confirm that the tenant has licensed Microsoft 365 Copilot users and that
concealed names are disabled in Microsoft 365 report settings.

**The Power BI refresh returns no files**

Complete one successful pipeline run, then check the Power BI parameters,
SharePoint credentials, library name, and snapshot folder names.

## Script reference

| Script | Purpose |
| --- | --- |
| `New-CopilotUsageApp.ps1` | Creates the certificate, app, permissions, and configuration |
| `Export-CopilotReports.ps1` | Exports and normalizes Copilot D28 usage |
| `Export-EntraUsers.ps1` | Exports Entra organizational properties |
| `Invoke-CopilotUsagePipeline.ps1` | Runs exports, logging, and SharePoint upload |
| `Register-CopilotUsageScheduledTask.ps1` | Registers the daily Windows task |

Each script includes comment-based help:

```powershell
Get-Help .\New-CopilotUsageApp.ps1 -Full
Get-Help .\Invoke-CopilotUsagePipeline.ps1 -Examples
```
