#requires -Version 7.4
#requires -Modules Az.Accounts, Az.Automation, Az.Resources, Microsoft.Graph.Authentication

<#
.SYNOPSIS
Deploys the Microsoft 365 Copilot usage pipeline to Azure Automation.

.DESCRIPTION
Creates or updates an Azure Automation account with a system-assigned managed
identity, a PowerShell 7.4 runtime environment, the Copilot usage runbook,
Automation variables, and a daily UTC schedule.

The script assigns these Microsoft Graph application permissions directly to
the managed identity:

- Reports.Read.All
- User.Read.All
- Lists.SelectedOperations.Selected

It then grants that identity write access to one SharePoint document library.
No app registration, certificate, client secret, or stored credential is used
by the runbook.

Azure and Microsoft Graph administrator sign-ins are interactive and are needed
only during deployment.

.PARAMETER TenantId
Microsoft Entra tenant ID or verified domain.

.PARAMETER SubscriptionId
Azure subscription in which the Automation account will be deployed.

.PARAMETER ResourceGroupName
Name of the target resource group. It is created when it does not exist.

.PARAMETER AutomationAccountName
Globally unique name of the Azure Automation account.

.PARAMETER Location
Azure region for a new resource group and Automation account.

.PARAMETER SharePointSiteUrl
Absolute URL of the SharePoint site containing the destination library.

.PARAMETER SharePointLibraryName
Display name of the destination document library.

.PARAMETER DailyAtUtc
Daily schedule time in UTC, in 24-hour HH:mm format.

.PARAMETER RunbookName
Name of the Azure Automation runbook.

.PARAMETER ScheduleName
Name of the Azure Automation schedule.

.PARAMETER RuntimeEnvironmentName
Name of the PowerShell 7.4 runtime environment.

.EXAMPLE
.\Deploy-CopilotUsageAzureAutomation.ps1 `
    -TenantId 'contoso.onmicrosoft.com' `
    -SubscriptionId '00000000-0000-0000-0000-000000000000' `
    -ResourceGroupName 'rg-copilot-adoption' `
    -AutomationAccountName 'aa-copilot-adoption-contoso' `
    -Location 'westeurope' `
    -SharePointSiteUrl 'https://contoso.sharepoint.com/sites/CopilotReports' `
    -SharePointLibraryName 'UsageData' `
    -DailyAtUtc '06:00'

.NOTES
The deploying administrator needs Azure permission to create and configure the
Automation resources, plus permission to consent to
Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All, and
Sites.ReadWrite.All during setup. These delegated permissions are not assigned
to the runtime managed identity.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [guid]$SubscriptionId,

    [Parameter(Mandatory)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory)]
    [string]$AutomationAccountName,

    [string]$Location = 'westeurope',

    [Parameter(Mandatory)]
    [ValidatePattern('^https://')]
    [string]$SharePointSiteUrl,

    [Parameter(Mandatory)]
    [string]$SharePointLibraryName,

    [ValidatePattern('^\d{2}:\d{2}$')]
    [string]$DailyAtUtc = '06:00',

    [string]$RunbookName = 'Invoke-CopilotUsageExport',

    [string]$ScheduleName = 'Daily-CopilotUsageExport',

    [ValidatePattern('^[A-Za-z][A-Za-z0-9_-]*$')]
    [string]$RuntimeEnvironmentName = 'CopilotUsage-PowerShell-7-4'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

function Set-AutomationVariableValue {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Value
    )

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create or update Azure Automation variable')) {
        return
    }

    $existing = Get-AzAutomationVariable `
        -ResourceGroupName $ResourceGroupName `
        -AutomationAccountName $AutomationAccountName `
        -Name $Name `
        -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        $null = New-AzAutomationVariable `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name $Name `
            -Value $Value `
            -Encrypted $false
    }
    else {
        $null = Set-AzAutomationVariable `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name $Name `
            -Value $Value `
            -Encrypted $false
    }
}

function Invoke-ArmWebRequest {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'PUT', 'PATCH', 'POST')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [string]$Body,

        [string]$ContentType = 'application/json'
    )

    $parameters = @{
        Method = $Method
        Uri = $Uri
        Authentication = 'Bearer'
        Token = $script:ArmAccessToken
        SkipHttpErrorCheck = $true
        ErrorAction = 'Stop'
    }
    if ($PSBoundParameters.ContainsKey('Body')) {
        $parameters.Body = $Body
        $parameters.ContentType = $ContentType
    }

    $response = Invoke-WebRequest @parameters
    if ($response.StatusCode -lt 200 -or $response.StatusCode -ge 300) {
        throw "Azure Resource Manager returned HTTP $($response.StatusCode) for $Method $Uri. $($response.Content)"
    }

    if ($response.StatusCode -eq 202) {
        $operationUri = [string]$response.Headers['Azure-AsyncOperation']
        if ([string]::IsNullOrWhiteSpace($operationUri)) {
            $operationUri = [string]$response.Headers.Location
        }

        if (-not [string]::IsNullOrWhiteSpace($operationUri)) {
            for ($attempt = 1; $attempt -le 60; $attempt++) {
                Start-Sleep -Seconds 2
                $operation = Invoke-WebRequest `
                    -Method GET `
                    -Uri $operationUri `
                    -Authentication Bearer `
                    -Token $script:ArmAccessToken `
                    -SkipHttpErrorCheck `
                    -ErrorAction Stop
                if ($operation.StatusCode -lt 200 -or $operation.StatusCode -ge 300) {
                    throw "Azure asynchronous operation returned HTTP $($operation.StatusCode). $($operation.Content)"
                }

                $operationStatus = $null
                if (-not [string]::IsNullOrWhiteSpace($operation.Content)) {
                    $operationBody = $operation.Content | ConvertFrom-Json
                    $statusProperty = $operationBody.PSObject.Properties['status']
                    if ($null -ne $statusProperty) {
                        $operationStatus = [string]$statusProperty.Value
                    }
                }

                if ([string]::IsNullOrWhiteSpace($operationStatus) -or $operationStatus -eq 'Succeeded') {
                    break
                }
                if ($operationStatus -in @('Failed', 'Canceled', 'Cancelled')) {
                    throw "Azure asynchronous operation ended with status '$operationStatus'. $($operation.Content)"
                }
                if ($attempt -eq 60) {
                    throw "Azure asynchronous operation did not finish within two minutes: $operationUri"
                }
            }
        }
    }

    return $response
}

function Get-ManagedIdentityServicePrincipal {
    param(
        [Parameter(Mandatory)]
        [guid]$PrincipalId
    )

    for ($attempt = 1; $attempt -le 12; $attempt++) {
        $statusCode = 0
        $servicePrincipal = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/${PrincipalId}?`$select=id,appId,displayName" `
            -OutputType PSObject `
            -SkipHttpErrorCheck `
            -StatusCodeVariable statusCode
        if ($statusCode -eq 200) {
            return $servicePrincipal
        }
        if ($statusCode -ne 404) {
            throw "Unable to read the managed identity service principal. Microsoft Graph returned HTTP $statusCode."
        }

        Start-Sleep -Seconds 5
    }

    throw 'The managed identity service principal was not visible in Microsoft Graph after 60 seconds.'
}

Write-Information 'Signing in to Azure. Use an account that can deploy the Automation resources.'
$azureContext = (Connect-AzAccount `
    -Tenant $TenantId `
    -Subscription $SubscriptionId `
    -ErrorAction Stop).Context
$azureContext = Set-AzContext `
    -SubscriptionId $SubscriptionId `
    -DefaultProfile $azureContext `
    -ErrorAction Stop

$resourceGroup = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if ($null -eq $resourceGroup) {
    $resourceGroup = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
    Write-Information "Created resource group '$ResourceGroupName'."
}

$automationAccount = Get-AzAutomationAccount `
    -ResourceGroupName $ResourceGroupName `
    -Name $AutomationAccountName `
    -ErrorAction SilentlyContinue
if ($null -eq $automationAccount) {
    $automationAccount = New-AzAutomationAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $AutomationAccountName `
        -Location $Location `
        -Plan Basic `
        -AssignSystemIdentity
    Write-Information "Created Automation account '$AutomationAccountName'."
}
elseif (
    $null -eq $automationAccount.Identity -or
    [string]$automationAccount.Identity.Type -notmatch 'SystemAssigned'
) {
    if (
        $null -ne $automationAccount.Identity -and
        [string]$automationAccount.Identity.Type -match 'UserAssigned'
    ) {
        throw "Automation account '$AutomationAccountName' already uses a user-assigned identity. Enable its system-assigned identity explicitly or use a dedicated account."
    }

    $automationAccount = Set-AzAutomationAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $AutomationAccountName `
        -AssignSystemIdentity
    Write-Information "Enabled the system-assigned identity on '$AutomationAccountName'."
}

$automationAccount = Get-AzAutomationAccount `
    -ResourceGroupName $ResourceGroupName `
    -Name $AutomationAccountName
if ($null -eq $automationAccount.Identity.PrincipalId) {
    throw "Automation account '$AutomationAccountName' does not expose a system-assigned identity principal ID."
}

Write-Information 'A browser sign-in will open for Microsoft Graph administrative consent and configuration.'
Connect-MgGraph `
    -TenantId $TenantId `
    -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Sites.ReadWrite.All') `
    -ContextScope Process `
    -NoWelcome `
    -ErrorAction Stop

try {
    $managedIdentity = Get-ManagedIdentityServicePrincipal `
        -PrincipalId $automationAccount.Identity.PrincipalId
    $graphResponse = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles" `
        -OutputType PSObject `
        -ErrorAction Stop
    $graphServicePrincipal = @($graphResponse.value)[0]
    if ($null -eq $graphServicePrincipal) {
        throw 'Unable to resolve the Microsoft Graph service principal.'
    }

    $requiredRoleNames = @(
        'Reports.Read.All',
        'User.Read.All',
        'Lists.SelectedOperations.Selected'
    )
    $existingAssignments = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($managedIdentity.id)/appRoleAssignments?`$select=appRoleId,resourceId" `
        -OutputType PSObject `
        -ErrorAction Stop

    foreach ($roleName in $requiredRoleNames) {
        $roles = @($graphServicePrincipal.appRoles | Where-Object {
            $_.value -eq $roleName -and $_.allowedMemberTypes -contains 'Application'
        })
        if ($roles.Count -ne 1) {
            throw "Unable to resolve Microsoft Graph application permission '$roleName'."
        }

        $role = $roles[0]
        $assignmentExists = @($existingAssignments.value | Where-Object {
            $_.appRoleId -eq $role.id -and $_.resourceId -eq $graphServicePrincipal.id
        }).Count -gt 0
        if (-not $assignmentExists) {
            $assignmentBody = @{
                principalId = $managedIdentity.id
                resourceId = $graphServicePrincipal.id
                appRoleId = $role.id
            } | ConvertTo-Json
            $null = Invoke-MgGraphRequest `
                -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($managedIdentity.id)/appRoleAssignments" `
                -Body $assignmentBody `
                -ContentType 'application/json' `
                -ErrorAction Stop
            Write-Information "Assigned Microsoft Graph application permission '$roleName'."
        }
    }

    $siteUri = [uri]$SharePointSiteUrl
    $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
    $siteLookupUri = if ([string]::IsNullOrWhiteSpace($sitePath)) {
        "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host)"
    }
    else {
        "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host):$sitePath"
    }
    $site = Invoke-MgGraphRequest `
        -Method GET `
        -Uri $siteLookupUri `
        -OutputType PSObject `
        -ErrorAction Stop
    $drives = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" `
        -OutputType PSObject `
        -ErrorAction Stop
    $matchingDrives = @($drives.value | Where-Object { $_.name -eq $SharePointLibraryName })
    if ($matchingDrives.Count -ne 1) {
        throw "Expected one SharePoint document library named '$SharePointLibraryName', but found $($matchingDrives.Count)."
    }
    $drive = $matchingDrives[0]
    $driveList = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/list?`$select=id" `
        -OutputType PSObject `
        -ErrorAction Stop

    $permissions = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($driveList.id)/permissions" `
        -OutputType PSObject `
        -ErrorAction Stop
    $existingPermission = @($permissions.value | Where-Object {
        $grantedToProperty = $_.PSObject.Properties['grantedToV2']
        if ($null -eq $grantedToProperty) {
            return $false
        }

        $applicationProperty = $grantedToProperty.Value.PSObject.Properties['application']
        if ($null -eq $applicationProperty) {
            return $false
        }

        return $applicationProperty.Value.id -eq $managedIdentity.appId
    })
    if ($existingPermission.Count -gt 1) {
        throw 'Multiple selected-permission grants exist for the managed identity on the target library.'
    }

    if ($existingPermission.Count -eq 0) {
        $permissionBody = @{
            roles = @('write')
            grantedToV2 = @{
                application = @{
                    id = $managedIdentity.appId
                    displayName = $AutomationAccountName
                }
            }
        } | ConvertTo-Json -Depth 10
        $null = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($driveList.id)/permissions" `
            -Body $permissionBody `
            -ContentType 'application/json' `
            -ErrorAction Stop
        Write-Information "Granted write access to SharePoint library '$SharePointLibraryName'."
    }
    elseif ($existingPermission[0].roles -notcontains 'write') {
        $permissionBody = @{ roles = @('write') } | ConvertTo-Json
        $null = Invoke-MgGraphRequest `
            -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($driveList.id)/permissions/$($existingPermission[0].id)" `
            -Body $permissionBody `
            -ContentType 'application/json' `
            -ErrorAction Stop
        Write-Information "Updated the SharePoint library grant to write."
    }

    Set-AutomationVariableValue `
        -Name 'CopilotUsageTenantId' `
        -Value $automationAccount.Identity.TenantId.ToString()
    Set-AutomationVariableValue `
        -Name 'CopilotUsageSharePointDriveId' `
        -Value $drive.id
    Set-AutomationVariableValue `
        -Name 'CopilotUsagePeriod' `
        -Value 'D28'
    Set-AutomationVariableValue `
        -Name 'CopilotUsageSharePointSiteUrl' `
        -Value $SharePointSiteUrl.TrimEnd('/')
    Set-AutomationVariableValue `
        -Name 'CopilotUsageSharePointLibraryName' `
        -Value $SharePointLibraryName

    $script:ArmAccessToken = (
        Get-AzAccessToken `
            -ResourceTypeName Arm `
            -AsSecureString `
            -DefaultProfile $azureContext `
            -ErrorAction Stop
    ).Token
    $armBase = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName"
    $apiVersion = '2024-10-23'
    $runtimeBody = @{
        location = $automationAccount.Location
        properties = @{
            runtime = @{
                language = 'PowerShell'
                version = '7.4'
            }
            defaultPackages = @{
                Az = '12.3.0'
            }
            description = 'PowerShell 7.4 runtime for the Microsoft 365 Copilot usage export.'
        }
    } | ConvertTo-Json -Depth 10
    $runtimeUri = "$armBase/runtimeEnvironments/${RuntimeEnvironmentName}?api-version=$apiVersion"
    $null = Invoke-ArmWebRequest `
        -Method PUT `
        -Uri $runtimeUri `
        -Body $runtimeBody

    for ($attempt = 1; $attempt -le 60; $attempt++) {
        $runtimeResponse = Invoke-ArmWebRequest -Method GET -Uri $runtimeUri
        $runtime = $runtimeResponse.Content | ConvertFrom-Json
        $provisioningState = [string]$runtime.properties.provisioningState
        if ([string]::IsNullOrWhiteSpace($provisioningState) -or $provisioningState -eq 'Succeeded') {
            break
        }
        if ($provisioningState -in @('Failed', 'Canceled', 'Cancelled')) {
            throw "Runtime environment provisioning ended with status '$provisioningState'."
        }
        if ($attempt -eq 60) {
            throw "Runtime environment '$RuntimeEnvironmentName' was not ready after five minutes."
        }
        Start-Sleep -Seconds 5
    }

    $runbookBody = @{
        location = $automationAccount.Location
        properties = @{
            runbookType = 'PowerShell'
            runtimeEnvironment = $RuntimeEnvironmentName
            logProgress = $false
            logVerbose = $false
            description = 'Exports Microsoft 365 Copilot and Entra snapshots to SharePoint.'
        }
    } | ConvertTo-Json -Depth 10
    $null = Invoke-ArmWebRequest `
        -Method PUT `
        -Uri "$armBase/runbooks/${RunbookName}?api-version=$apiVersion" `
        -Body $runbookBody

    $runbookPath = Join-Path $PSScriptRoot 'Invoke-CopilotUsageRunbook.ps1'
    $runbookContent = Get-Content -LiteralPath $runbookPath -Raw
    $null = Invoke-ArmWebRequest `
        -Method PUT `
        -Uri "$armBase/runbooks/${RunbookName}/draft/content?api-version=$apiVersion" `
        -Body $runbookContent `
        -ContentType 'text/plain'
    $null = Invoke-ArmWebRequest `
        -Method POST `
        -Uri "$armBase/runbooks/${RunbookName}/draft/publish?api-version=$apiVersion"
    Write-Information "Published runbook '$RunbookName' with PowerShell 7.4."

    $schedule = Get-AzAutomationSchedule `
        -ResourceGroupName $ResourceGroupName `
        -AutomationAccountName $AutomationAccountName `
        -Name $ScheduleName `
        -ErrorAction SilentlyContinue
    if ($null -eq $schedule) {
        $scheduleTime = [timespan]::ParseExact(
            $DailyAtUtc,
            'hh\:mm',
            [Globalization.CultureInfo]::InvariantCulture
        )
        $startTime = [DateTime]::UtcNow.Date.Add($scheduleTime)
        if ($startTime -lt [DateTime]::UtcNow.AddMinutes(10)) {
            $startTime = $startTime.AddDays(1)
        }
        $schedule = New-AzAutomationSchedule `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name $ScheduleName `
            -StartTime $startTime `
            -DayInterval 1 `
            -TimeZone 'UTC'
        Write-Information "Created daily schedule '$ScheduleName' for $DailyAtUtc UTC."
    }
    elseif ($schedule.StartTime.ToUniversalTime().ToString('HH:mm') -ne $DailyAtUtc) {
        throw "Schedule '$ScheduleName' already starts at $($schedule.StartTime.ToUniversalTime().ToString('HH:mm')) UTC, not $DailyAtUtc UTC. Use a different ScheduleName or update the existing schedule."
    }

    $scheduledRunbooks = @(
        Get-AzAutomationScheduledRunbook `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -RunbookName $RunbookName `
            -ErrorAction SilentlyContinue
    )
    if ($scheduledRunbooks.ScheduleName -notcontains $ScheduleName) {
        $null = Register-AzAutomationScheduledRunbook `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -RunbookName $RunbookName `
            -ScheduleName $ScheduleName
    }

    Write-Information ''
    Write-Information 'Azure Automation deployment completed.'
    Write-Information "Managed identity object ID: $($managedIdentity.id)"
    Write-Information "Managed identity application ID: $($managedIdentity.appId)"
    Write-Information "SharePoint drive ID: $($drive.id)"
    Write-Information "Daily schedule: $DailyAtUtc UTC"
    Write-Information "Test with: Start-AzAutomationRunbook -ResourceGroupName '$ResourceGroupName' -AutomationAccountName '$AutomationAccountName' -Name '$RunbookName'"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
