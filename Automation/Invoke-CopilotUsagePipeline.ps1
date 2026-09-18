#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
Runs the Copilot usage export pipeline and uploads snapshots to SharePoint.

.DESCRIPTION
Loads a JSON configuration file, authenticates to Microsoft Graph with a
certificate, exports Microsoft 365 Copilot reports and Entra user properties,
and uploads the resulting CSV files to a single SharePoint document library.

The script creates three folders in the library when needed:
UserSnapshots, TrendSnapshots, and DirectorySnapshots. Uploads use the source
file name and replace an existing file with the same name, making retries
idempotent.

This is the entry point intended for Windows Task Scheduler.

.PARAMETER ConfigPath
Path to automation.config.json. See automation.config.example.json for the
required properties.

.OUTPUTS
None. Progress and errors are written to the console and to a daily log file.
The script throws on export, authentication, configuration, or upload failure.

.EXAMPLE
.\Invoke-CopilotUsagePipeline.ps1 `
    -ConfigPath '.\automation.config.json'

Runs all exports and uploads using the supplied configuration.

.NOTES
Required runtime Microsoft Graph application permissions:
- Reports.Read.All
- User.Read.All
- Lists.SelectedOperations.Selected

Lists.SelectedOperations.Selected grants no access by itself. The app must also
receive the write role on the target document library, which the bootstrap
script configures.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-PipelineLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    $line = '{0:u} {1}' -f (Get-Date), $Message
    Write-Information $line -InformationAction Continue
    Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
}

function Get-RequiredConfigValue {
    param(
        [Parameter(Mandatory)]
        [psobject]$Config,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        throw "Configuration value '$Name' is required."
    }

    return [string]$property.Value
}

function Initialize-DriveFolder {
    param(
        [Parameter(Mandatory)]
        [string]$DriveId,

        [Parameter(Mandatory)]
        [string]$FolderName
    )

    $encodedFolder = [uri]::EscapeDataString($FolderName)
    $statusCode = 0
    $null = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encodedFolder" `
        -OutputType PSObject `
        -SkipHttpErrorCheck `
        -StatusCodeVariable statusCode

    if ($statusCode -eq 200) {
        return
    }
    if ($statusCode -ne 404) {
        throw "Unable to check SharePoint folder '$FolderName'. Microsoft Graph returned HTTP $statusCode."
    }

    $body = @{
        name = $FolderName
        folder = @{}
        '@microsoft.graph.conflictBehavior' = 'fail'
    } | ConvertTo-Json -Depth 5

    $null = Invoke-MgGraphRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children" `
        -Body $body `
        -ContentType 'application/json' `
        -ErrorAction Stop
}

function Send-FileToSharePoint {
    param(
        [Parameter(Mandatory)]
        [string]$DriveId,

        [Parameter(Mandatory)]
        [string]$FolderName,

        [Parameter(Mandatory)]
        [string]$FilePath
    )

    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        throw "Upload source file does not exist: $FilePath"
    }

    Initialize-DriveFolder -DriveId $DriveId -FolderName $FolderName
    $encodedFolder = [uri]::EscapeDataString($FolderName)
    $encodedFileName = [uri]::EscapeDataString([System.IO.Path]::GetFileName($FilePath))
    $uploadUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encodedFolder/${encodedFileName}:/content"

    # PUT to the path-based content endpoint replaces the same logical snapshot on retry.
    $null = Invoke-MgGraphRequest `
        -Method PUT `
        -Uri $uploadUri `
        -InputFilePath $FilePath `
        -ContentType 'text/csv' `
        -ErrorAction Stop
}

$resolvedConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$config = Get-Content -LiteralPath $resolvedConfigPath -Raw | ConvertFrom-Json
$tenantId = Get-RequiredConfigValue -Config $config -Name 'TenantId'
$clientId = Get-RequiredConfigValue -Config $config -Name 'ClientId'
$certificateThumbprint = Get-RequiredConfigValue -Config $config -Name 'CertificateThumbprint'
$driveId = Get-RequiredConfigValue -Config $config -Name 'SharePointDriveId'
$period = Get-RequiredConfigValue -Config $config -Name 'Period'
$localDataRoot = Get-RequiredConfigValue -Config $config -Name 'LocalDataRoot'
$logRoot = Get-RequiredConfigValue -Config $config -Name 'LogRoot'

[void](New-Item -ItemType Directory -Path $localDataRoot -Force)
[void](New-Item -ItemType Directory -Path $logRoot -Force)
$script:LogFile = Join-Path $logRoot ("CopilotUsagePipeline-{0}.log" -f (Get-Date -Format 'yyyyMMdd'))

try {
    Write-PipelineLog 'Starting Microsoft 365 Copilot usage pipeline.'
    Connect-MgGraph `
        -TenantId $tenantId `
        -ClientId $clientId `
        -CertificateThumbprint $certificateThumbprint `
        -ContextScope Process `
        -NoWelcome `
        -ErrorAction Stop

    $reports = & (Join-Path $PSScriptRoot 'Export-CopilotReports.ps1') `
        -OutputRoot $localDataRoot `
        -Period $period `
        -SkipConnect
    Write-PipelineLog ("Exported {0} Copilot users for report refresh date {1}." -f $reports.UserCount, $reports.ReportRefreshDate)

    $directory = & (Join-Path $PSScriptRoot 'Export-EntraUsers.ps1') `
        -OutputRoot $localDataRoot `
        -SkipConnect
    Write-PipelineLog ("Exported {0} Microsoft Entra ID users." -f $directory.UserCount)

    Send-FileToSharePoint -DriveId $driveId -FolderName 'UserSnapshots' -FilePath $reports.UserFile
    Send-FileToSharePoint -DriveId $driveId -FolderName 'TrendSnapshots' -FilePath $reports.TrendFile
    Send-FileToSharePoint -DriveId $driveId -FolderName 'DirectorySnapshots' -FilePath $directory.DirectoryFile

    Write-PipelineLog 'Uploaded all CSV snapshots to SharePoint successfully.'
}
catch {
    Write-PipelineLog ("FAILED: {0}" -f $_.Exception.Message)
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
