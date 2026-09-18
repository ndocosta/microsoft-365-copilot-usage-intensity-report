# Azure Automation deployment

This is the recommended production option. The runbook uses the Automation
account's system-assigned managed identity, so there is no certificate, client
secret, stored credential, or local machine to maintain.

## What the deployment creates

- Resource group, when the requested one does not exist
- Azure Automation account with a system-assigned managed identity
- PowerShell 7.4 runtime environment with the default Az package
- Published Copilot usage runbook
- Non-secret Automation variables
- Daily UTC schedule and runbook association
- Microsoft Graph application-role assignments on the managed identity
- `write` access on only the selected SharePoint document library

Existing resources with the requested names are reused and updated where
possible.

## Prerequisites

- PowerShell 7.4
- An Azure subscription in the same Microsoft Entra tenant as Microsoft 365
- Permission to create and configure the Azure resources
- Permission to grant Microsoft Graph admin consent and SharePoint library
  access
- These PowerShell modules:

```powershell
Install-Module Az.Accounts, Az.Automation, Az.Resources, Microsoft.Graph.Authentication `
    -Scope CurrentUser
```

The setup uses interactive Azure and Microsoft Graph sign-ins. The administrator
temporarily consents to `Application.ReadWrite.All`,
`AppRoleAssignment.ReadWrite.All`, and `Sites.ReadWrite.All`. These delegated
permissions are used only for deployment.

## 1. Deploy

From the repository root:

```powershell
.\Automation\Azure\Deploy-CopilotUsageAzureAutomation.ps1 `
    -TenantId 'contoso.onmicrosoft.com' `
    -SubscriptionId '00000000-0000-0000-0000-000000000000' `
    -ResourceGroupName 'rg-copilot-adoption' `
    -AutomationAccountName 'aa-copilot-adoption-contoso' `
    -Location 'westeurope' `
    -SharePointSiteUrl 'https://contoso.sharepoint.com/sites/CopilotReports' `
    -SharePointLibraryName 'UsageData' `
    -DailyAtUtc '06:00'
```

Automation account names must be globally unique. `DailyAtUtc` is intentionally
UTC so the run time does not move with daylight-saving changes.

The deployment may take a few minutes while the managed identity becomes
visible in Microsoft Entra ID and the PowerShell runtime is prepared.

## 2. Test the runbook

Start a job:

```powershell
$job = Start-AzAutomationRunbook `
    -ResourceGroupName 'rg-copilot-adoption' `
    -AutomationAccountName 'aa-copilot-adoption-contoso' `
    -Name 'Invoke-CopilotUsageExport'

$job | Wait-AzAutomationJob
$job | Get-AzAutomationJobOutput
```

Confirm that the job completes and that all three snapshot folders contain a
CSV file in SharePoint. Azure Automation writes progress and errors to the job
stream; it does not retain local CSV or log files.

## Runtime configuration

The deployment stores these non-secret Automation variables:

| Variable | Purpose |
| --- | --- |
| `CopilotUsageTenantId` | Requests a Graph token without Azure subscription access |
| `CopilotUsageSharePointDriveId` | Identifies the selected document library |
| `CopilotUsagePeriod` | Report period; currently `D28` |
| `CopilotUsageSharePointSiteUrl` | Human-readable deployment reference |
| `CopilotUsageSharePointLibraryName` | Human-readable deployment reference |

The first three are read by the scheduled runbook. Runbook parameters can
override them during a manual test.

## Managed identity permissions

The managed identity receives only:

- `Reports.Read.All`
- `User.Read.All`
- `Lists.SelectedOperations.Selected`
- `write` on the configured SharePoint document library

It receives no Azure subscription role from this solution. Azure Automation can
obtain its own managed-identity token without a subscription RBAC assignment.

## Update the runbook

Run the deployment script again after pulling a newer version. It republishes
the runbook and preserves the existing identity, permissions, variables, and
schedule.

If you need a different daily time, create a new schedule name with
`-ScheduleName`, or update the existing schedule in Azure before redeploying.
The deployment stops rather than silently reusing a schedule at a different
time.

## Remove the deployment

Resource deletion is intentionally not automated. Review retained CSV data,
SharePoint permissions, job history, and any other resources in the resource
group before deleting them through your normal Azure governance process.

## Troubleshooting

**The runbook receives HTTP 403 from Microsoft Graph**

Confirm all three application roles are assigned to the managed identity. Role
changes may take several minutes to appear in a new token; retry with a new job.

**SharePoint upload receives HTTP 403**

Confirm the list-specific `write` grant exists on the destination library and
that `CopilotUsageSharePointDriveId` points to that same library.

**The PowerShell 7.4 runtime is still provisioning**

Wait for the Runtime Environment status in Azure Automation to become ready,
then start a new job.

**The deployment cannot grant application roles**

Use an administrator with permission to grant application permissions and
assign app roles. Privileged Role Administrator is a common choice; tenant
policy may require a different approved role or process.
