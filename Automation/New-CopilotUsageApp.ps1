#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
Provisions the identity and SharePoint access used by the export pipeline.

.DESCRIPTION
Performs the one-time administrative setup for the Microsoft 365 Copilot usage
export solution. The script:

1. Creates a non-exportable, 2048-bit RSA certificate in Cert:\CurrentUser\My.
2. Creates a single-tenant app registration and service principal.
3. Assigns the runtime application permissions Reports.Read.All,
   User.Read.All, and Lists.SelectedOperations.Selected.
4. Grants the app the write role on one SharePoint document library only.
5. Writes automation.config.json for the pipeline and Scheduled Task scripts.

The interactive administrator permissions requested by this script are used
only during provisioning. They are not assigned to the runtime application.

The script stops if an app registration with the same display name already
exists. It does not modify or delete an existing application.

.PARAMETER TenantId
Microsoft Entra tenant ID or verified domain to provision.

.PARAMETER DisplayName
Display name for the app registration, service principal, and certificate.

.PARAMETER SharePointSiteUrl
Absolute URL of the SharePoint site containing the destination library, for
example https://contoso.sharepoint.com/sites/CopilotReports.

.PARAMETER SharePointLibraryName
Display name of the destination SharePoint document library. This must be a
document library, not a folder path.

.PARAMETER ConfigPath
Path where the generated JSON configuration is written. The file contains
identifiers and local paths but no client secret or private key.

.OUTPUTS
None. The script writes a provisioning summary to the console.

.EXAMPLE
.\New-CopilotUsageApp.ps1 `
    -TenantId 'contoso.onmicrosoft.com' `
    -SharePointSiteUrl 'https://contoso.sharepoint.com/sites/CopilotReports' `
    -SharePointLibraryName 'UsageData'

Creates the runtime identity, grants access only to UsageData, and writes the
default automation.config.json.

.NOTES
Run this script as the same Windows user that will run the Scheduled Task. The
authentication certificate is non-exportable and stored in that user's
certificate store.

Interactive provisioning requires an administrator able to consent to:
Application.ReadWrite.All, AppRoleAssignment.ReadWrite.All, and
Sites.ReadWrite.All.

The certificate expires after 24 months and must be rotated before expiry.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TenantId,

    [string]$DisplayName = 'Copilot Usage Reports Exporter',

    [Parameter(Mandatory)]
    [ValidatePattern('^https://')]
    [string]$SharePointSiteUrl,

    [Parameter(Mandatory)]
    [string]$SharePointLibraryName,

    [string]$ConfigPath = (Join-Path $PSScriptRoot 'automation.config.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Information 'A browser sign-in will open. Sign in as an administrator of the target tenant.' -InformationAction Continue
Connect-MgGraph `
    -TenantId $TenantId `
    -Scopes @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Sites.ReadWrite.All') `
    -ContextScope Process `
    -NoWelcome `
    -ErrorAction Stop

try {
    $existing = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$($DisplayName.Replace("'", "''"))'&`$select=id,appId,displayName" `
        -OutputType PSObject `
        -ErrorAction Stop
    if (@($existing.value).Count -gt 0) {
        throw "An app registration named '$DisplayName' already exists. Reuse or remove it explicitly before running this setup."
    }

    $certificate = New-SelfSignedCertificate `
        -Subject "CN=$DisplayName" `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -KeyExportPolicy NonExportable `
        -KeySpec Signature `
        -KeyUsage DigitalSignature `
        -NotAfter (Get-Date).AddMonths(24) `
        -HashAlgorithm SHA256

    $graphServicePrincipal = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles" `
        -OutputType PSObject `
        -ErrorAction Stop
    $graphSp = @($graphServicePrincipal.value)[0]

    $requiredRoleNames = @('Reports.Read.All', 'User.Read.All', 'Lists.SelectedOperations.Selected')
    $resourceAccess = foreach ($roleName in $requiredRoleNames) {
        $role = @($graphSp.appRoles | Where-Object { $_.value -eq $roleName -and $_.allowedMemberTypes -contains 'Application' })
        if ($role.Count -ne 1) {
            throw "Unable to resolve Microsoft Graph application permission '$roleName'."
        }
        @{ id = $role[0].id; type = 'Role' }
    }

    $applicationBody = @{
        displayName = $DisplayName
        signInAudience = 'AzureADMyOrg'
        requiredResourceAccess = @(
            @{
                resourceAppId = '00000003-0000-0000-c000-000000000000'
                resourceAccess = @($resourceAccess)
            }
        )
        keyCredentials = @(
            @{
                type = 'AsymmetricX509Cert'
                usage = 'Verify'
                key = [Convert]::ToBase64String($certificate.RawData)
                displayName = 'Scheduled Task certificate'
                startDateTime = $certificate.NotBefore.ToUniversalTime().ToString('o')
                endDateTime = $certificate.NotAfter.ToUniversalTime().ToString('o')
            }
        )
    } | ConvertTo-Json -Depth 10

    $application = Invoke-MgGraphRequest `
        -Method POST `
        -Uri 'https://graph.microsoft.com/v1.0/applications' `
        -Body $applicationBody `
        -ContentType 'application/json' `
        -OutputType PSObject `
        -ErrorAction Stop

    $servicePrincipalBody = @{ appId = $application.appId } | ConvertTo-Json
    $servicePrincipal = Invoke-MgGraphRequest `
        -Method POST `
        -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' `
        -Body $servicePrincipalBody `
        -ContentType 'application/json' `
        -OutputType PSObject `
        -ErrorAction Stop

    foreach ($roleAccess in $resourceAccess) {
        $assignmentBody = @{
            principalId = $servicePrincipal.id
            resourceId = $graphSp.id
            appRoleId = $roleAccess.id
        } | ConvertTo-Json

        $null = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($servicePrincipal.id)/appRoleAssignments" `
            -Body $assignmentBody `
            -ContentType 'application/json' `
            -ErrorAction Stop
    }

    $siteUri = [uri]$SharePointSiteUrl
    $sitePath = $siteUri.AbsolutePath.TrimEnd('/')
    $site = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host):$sitePath" `
        -OutputType PSObject `
        -ErrorAction Stop

    $drives = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" `
        -OutputType PSObject `
        -ErrorAction Stop
    $drive = @($drives.value | Where-Object { $_.name -eq $SharePointLibraryName })
    if ($drive.Count -ne 1) {
        throw "Expected one SharePoint document library named '$SharePointLibraryName', but found $($drive.Count)."
    }

    $driveList = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/drives/$($drive[0].id)/list?`$select=id" `
        -OutputType PSObject `
        -ErrorAction Stop

    # Selected permissions require both Entra consent and a resource-specific grant.
    $listPermissionBody = @{
        roles = @('write')
        grantedToV2 = @{
            application = @{
                id = $application.appId
                displayName = $DisplayName
            }
        }
    } | ConvertTo-Json -Depth 10

    $null = Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($driveList.id)/permissions" `
        -Body $listPermissionBody `
        -ContentType 'application/json' `
        -ErrorAction Stop

    $projectRoot = Split-Path $PSScriptRoot -Parent
    $config = [ordered]@{
        TenantId = $TenantId
        ClientId = $application.appId
        CertificateThumbprint = $certificate.Thumbprint
        SharePointSiteUrl = $SharePointSiteUrl
        SharePointLibraryName = $SharePointLibraryName
        SharePointDriveId = $drive[0].id
        Period = 'D28'
        LocalDataRoot = Join-Path $projectRoot 'Data'
        LogRoot = Join-Path $projectRoot 'Logs'
    }
    $config | ConvertTo-Json | Set-Content -LiteralPath $ConfigPath -Encoding UTF8

    Write-Information "Created app registration: $($application.appId)" -InformationAction Continue
    Write-Information "Granted runtime permissions: $($requiredRoleNames -join ', ')" -InformationAction Continue
    Write-Information 'Granted library-specific SharePoint role: write' -InformationAction Continue
    Write-Information "Created configuration: $ConfigPath" -InformationAction Continue
    Write-Warning "The certificate expires on $($certificate.NotAfter.ToString('yyyy-MM-dd')). Schedule rotation before that date."
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
