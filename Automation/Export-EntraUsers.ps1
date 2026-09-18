#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
Exports Microsoft Entra user profile properties to a daily CSV snapshot.

.DESCRIPTION
Reads users from Microsoft Graph, follows all result pages, and exports the
organizational attributes used to enrich Microsoft 365 Copilot usage data in
Power BI. Exported attributes include department, job title, office location,
geography, company, account state, and user type.

The file is written to the DirectorySnapshots child directory under OutputRoot.
One file name is used per execution date, so retries on the same day replace
the same local snapshot.

.PARAMETER OutputRoot
Root directory for generated data. The script creates a DirectorySnapshots
child directory when it does not exist.

.PARAMETER TenantId
Microsoft Entra tenant ID or verified domain. Required unless SkipConnect is
used.

.PARAMETER ClientId
Application (client) ID of the app registration used for certificate
authentication. Required unless SkipConnect is used.

.PARAMETER CertificateThumbprint
Thumbprint of the authentication certificate in Cert:\CurrentUser\My. Required
unless SkipConnect is used.

.PARAMETER SkipConnect
Reuses an existing Microsoft Graph context. This is intended for the pipeline
orchestrator, which opens one connection for all export and upload operations.

.OUTPUTS
PSCustomObject containing SnapshotDate, DirectoryFile, and UserCount.

.EXAMPLE
.\Export-EntraUsers.ps1 `
    -OutputRoot 'C:\CopilotAdoption\Data' `
    -TenantId 'contoso.onmicrosoft.com' `
    -ClientId '00000000-0000-0000-0000-000000000000' `
    -CertificateThumbprint '0123456789ABCDEF0123456789ABCDEF01234567'

Exports a complete Entra user snapshot using application-only authentication.

.NOTES
Runtime Microsoft Graph application permission: User.Read.All.

The output contains personal and organizational data. Store it in a protected
location and apply appropriate retention and access controls.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputRoot,

    [string]$TenantId,

    [string]$ClientId,

    [string]$CertificateThumbprint,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$connectedHere = $false

try {
    if (-not $SkipConnect) {
        foreach ($value in @($TenantId, $ClientId, $CertificateThumbprint)) {
            if ([string]::IsNullOrWhiteSpace($value)) {
                throw 'TenantId, ClientId, and CertificateThumbprint are required for certificate authentication.'
            }
        }

        Connect-MgGraph `
            -TenantId $TenantId `
            -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint `
            -ContextScope Process `
            -NoWelcome `
            -ErrorAction Stop
        $connectedHere = $true
    }
    elseif (-not (Get-MgContext)) {
        throw 'SkipConnect was specified, but no Microsoft Graph context is active.'
    }

    $directoryFolder = Join-Path $OutputRoot 'DirectorySnapshots'
    [void](New-Item -ItemType Directory -Path $directoryFolder -Force)

    $users = [System.Collections.Generic.List[object]]::new()
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,department,jobTitle,officeLocation,city,state,country,usageLocation,companyName,accountEnabled,userType'

    # Microsoft Graph returns users in pages and supplies the next page as an OData link.
    while ($uri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
        foreach ($user in $response.value) {
            $users.Add($user)
        }
        $nextLink = $response.PSObject.Properties['@odata.nextLink']
        $uri = if ($null -ne $nextLink) { [string]$nextLink.Value } else { $null }
    }

    if ($users.Count -eq 0) {
        throw 'Microsoft Entra ID returned no users.'
    }

    $snapshotDate = (Get-Date).ToString('yyyy-MM-dd')
    $outputFile = Join-Path $directoryFolder "DirectoryUsers-$snapshotDate.csv"

    $users |
        Sort-Object userPrincipalName |
        ForEach-Object {
            [pscustomobject][ordered]@{
                SnapshotDate      = $snapshotDate
                UserId            = $_.id
                UserPrincipalName = $_.userPrincipalName
                DisplayName       = $_.displayName
                Department        = $_.department
                JobTitle          = $_.jobTitle
                OfficeLocation    = $_.officeLocation
                City              = $_.city
                State             = $_.state
                Country           = $_.country
                UsageLocation     = $_.usageLocation
                CompanyName       = $_.companyName
                AccountEnabled    = $_.accountEnabled
                UserType          = $_.userType
            }
        } |
        Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding UTF8

    [pscustomobject]@{
        SnapshotDate = $snapshotDate
        DirectoryFile = $outputFile
        UserCount = $users.Count
    }
}
finally {
    if ($connectedHere) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
