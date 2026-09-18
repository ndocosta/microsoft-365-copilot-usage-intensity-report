#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
Exports Microsoft 365 Copilot usage and daily trend snapshots.

.DESCRIPTION
Downloads version 2 of the Microsoft 365 Copilot user-detail and user-count
trend reports from Microsoft Graph. The script normalizes the source columns
into stable CSV contracts used by the accompanying Power BI semantic model.

The export is currently fixed to the rolling 28-day reporting period. Snapshot
file names use the report refresh date rather than the execution date. Running
the script more than once before the source report refreshes therefore replaces
the same local snapshot instead of creating duplicate logical snapshots.

The Microsoft Graph report contains licensed Microsoft 365 Copilot users only.
Unlicensed Copilot Chat usage is not included.

.PARAMETER OutputRoot
Root directory for generated data. The script creates UserSnapshots and
TrendSnapshots child directories when they do not exist.

.PARAMETER Period
Reporting window to request. Only D28 is supported because the Power BI model
and adoption thresholds are designed for a rolling 28-day period.

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
PSCustomObject containing ReportRefreshDate, UserFile, TrendFile, UserCount,
and TrendRowCount.

.EXAMPLE
.\Export-CopilotReports.ps1 `
    -OutputRoot 'C:\CopilotAdoption\Data' `
    -TenantId 'contoso.onmicrosoft.com' `
    -ClientId '00000000-0000-0000-0000-000000000000' `
    -CertificateThumbprint '0123456789ABCDEF0123456789ABCDEF01234567'

Exports the D28 user-detail and trend reports using application-only
authentication.

.NOTES
Runtime Microsoft Graph application permission: Reports.Read.All.

Microsoft 365 usage reports can take up to approximately 48 hours to reflect
activity. If concealed user names are enabled in Microsoft 365 report settings,
the exported identities are pseudonymized and cannot be joined to Entra users.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputRoot,

    [ValidateSet('D28')]
    [string]$Period = 'D28',

    [string]$TenantId,

    [string]$ClientId,

    [string]$CertificateThumbprint,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ReportValue {
    param(
        [Parameter(Mandatory)]
        [psobject]$Row,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $null
}

$connectedHere = $false
$temporaryFolder = $null

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

    $userFolder = Join-Path $OutputRoot 'UserSnapshots'
    $trendFolder = Join-Path $OutputRoot 'TrendSnapshots'
    [void](New-Item -ItemType Directory -Path $userFolder -Force)
    [void](New-Item -ItemType Directory -Path $trendFolder -Force)

    $temporaryFolder = Join-Path ([System.IO.Path]::GetTempPath()) ("CopilotReports-" + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $temporaryFolder)

    $rawUserFile = Join-Path $temporaryFolder 'UserDetail.csv'
    $rawTrendFile = Join-Path $temporaryFolder 'UserCountTrend.csv'
    $userUri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$Period',version='v2')"
    $trendUri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountTrend(period='$Period',version='v2')"

    Invoke-MgGraphRequest -Method GET -Uri $userUri -OutputFilePath $rawUserFile -ErrorAction Stop
    Invoke-MgGraphRequest -Method GET -Uri $trendUri -OutputFilePath $rawTrendFile -ErrorAction Stop

    $rawUsers = @(Import-Csv -LiteralPath $rawUserFile)
    $rawTrend = @(Import-Csv -LiteralPath $rawTrendFile)
    if ($rawUsers.Count -eq 0) {
        throw 'The Microsoft 365 Copilot user-detail report returned no rows.'
    }
    if ($rawTrend.Count -eq 0) {
        throw 'The Microsoft 365 Copilot trend report returned no rows.'
    }

    $header = Get-Content -LiteralPath $rawUserFile -TotalCount 1
    if ($header -notmatch 'Prompts submitted') {
        throw 'The user-detail report does not contain the version 2 prompt columns.'
    }

    # The source refresh date makes retries idempotent while preserving daily history.
    $reportRefreshDate = [datetime](Get-ReportValue -Row $rawUsers[0] -Names @('Report Refresh Date'))
    $snapshotDate = $reportRefreshDate.ToString('yyyy-MM-dd')
    $reportPeriod = [int]($Period.Substring(1))

    $users = foreach ($row in $rawUsers) {
        $promptsAll = [int](Get-ReportValue -Row $row -Names @(
            'Prompts submitted for all apps',
            'Prompts submitted (any app)',
            'Prompts submitted'
        ))

        $segment = if ($promptsAll -ge 30) {
            'High adoption (30+)'
        }
        elseif ($promptsAll -ge 10) {
            'Developing (10-29)'
        }
        else {
            'Low adoption (0-9)'
        }

        [pscustomobject][ordered]@{
            SnapshotDate                        = $snapshotDate
            ReportRefreshDate                   = $snapshotDate
            ReportPeriod                        = $reportPeriod
            UserPrincipalName                   = Get-ReportValue -Row $row -Names @('User Principal Name')
            ReportDisplayName                   = Get-ReportValue -Row $row -Names @('Display Name')
            UsageSegment                        = $segment
            PromptsAllApps                      = $promptsAll
            PromptsCopilotChatWork              = [int](Get-ReportValue -Row $row -Names @('Prompts submitted for Copilot Chat (work)'))
            PromptsCopilotChatWeb               = [int](Get-ReportValue -Row $row -Names @('Prompts submitted for Copilot Chat (web)'))
            ActiveUsageDays                     = [int](Get-ReportValue -Row $row -Names @('Active Usage Days for all apps'))
            LastActivityDate                    = Get-ReportValue -Row $row -Names @('Last Activity Date')
            Microsoft365CopilotLastActivityDate = Get-ReportValue -Row $row -Names @('Microsoft 365 Copilot Last Activity Date')
            CopilotChatLastActivityDate         = Get-ReportValue -Row $row -Names @('Copilot Chat Last Activity Date')
            CopilotChatWorkLastActivityDate     = Get-ReportValue -Row $row -Names @('Copilot Chat (work) Last Activity Date')
            CopilotChatWebLastActivityDate      = Get-ReportValue -Row $row -Names @('Copilot Chat (web) Last Activity Date')
            TeamsLastActivityDate               = Get-ReportValue -Row $row -Names @('Microsoft Teams Copilot Last Activity Date')
            WordLastActivityDate                = Get-ReportValue -Row $row -Names @('Word Copilot Last Activity Date')
            ExcelLastActivityDate               = Get-ReportValue -Row $row -Names @('Excel Copilot Last Activity Date')
            PowerPointLastActivityDate          = Get-ReportValue -Row $row -Names @('PowerPoint Copilot Last Activity Date')
            OutlookLastActivityDate             = Get-ReportValue -Row $row -Names @('Outlook Copilot Last Activity Date')
            OneNoteLastActivityDate             = Get-ReportValue -Row $row -Names @('OneNote Copilot Last Activity Date')
            LoopLastActivityDate                = Get-ReportValue -Row $row -Names @('Loop Copilot Last Activity Date')
            EdgeLastActivityDate                = Get-ReportValue -Row $row -Names @('Edge Last Activity Date')
            CopilotAgentLastActivityDate        = Get-ReportValue -Row $row -Names @('Copilot Agent Last Activity Date')
        }
    }

    $trend = foreach ($row in $rawTrend) {
        [pscustomobject][ordered]@{
            SnapshotDate                   = $snapshotDate
            ReportRefreshDate              = Get-ReportValue -Row $row -Names @('Report Refresh Date')
            ReportDate                     = Get-ReportValue -Row $row -Names @('Report Date')
            ReportPeriod                   = $reportPeriod
            EnabledUsers                   = [int](Get-ReportValue -Row $row -Names @('Any App Enabled Users'))
            ActiveUsers                    = [int](Get-ReportValue -Row $row -Names @('Any App Active Users'))
            Microsoft365CopilotActiveUsers = [int](Get-ReportValue -Row $row -Names @('Microsoft 365 Copilot Active Users'))
            CopilotChatActiveUsers         = [int](Get-ReportValue -Row $row -Names @('Copilot Chat Active Users'))
            CopilotChatWorkActiveUsers     = [int](Get-ReportValue -Row $row -Names @('Copilot Chat (work) Active Users'))
            CopilotChatWebActiveUsers      = [int](Get-ReportValue -Row $row -Names @('Copilot Chat (web) Active Users'))
            TeamsActiveUsers               = [int](Get-ReportValue -Row $row -Names @('Microsoft Teams Active Users'))
            WordActiveUsers                = [int](Get-ReportValue -Row $row -Names @('Word Active Users'))
            ExcelActiveUsers               = [int](Get-ReportValue -Row $row -Names @('Excel Active Users'))
            PowerPointActiveUsers          = [int](Get-ReportValue -Row $row -Names @('PowerPoint Active Users'))
            OutlookActiveUsers             = [int](Get-ReportValue -Row $row -Names @('Outlook Active Users'))
            OneNoteActiveUsers             = [int](Get-ReportValue -Row $row -Names @('OneNote Active Users'))
            LoopActiveUsers                = [int](Get-ReportValue -Row $row -Names @('Loop Active Users'))
            EdgeActiveUsers                = [int](Get-ReportValue -Row $row -Names @('Edge Active Users'))
            PromptsSubmitted               = [int](Get-ReportValue -Row $row -Names @('Prompts submitted'))
        }
    }

    $userOutputFile = Join-Path $userFolder "CopilotUsers-$Period-$snapshotDate.csv"
    $trendOutputFile = Join-Path $trendFolder "CopilotTrend-$Period-$snapshotDate.csv"
    $users | Export-Csv -LiteralPath $userOutputFile -NoTypeInformation -Encoding UTF8
    $trend | Export-Csv -LiteralPath $trendOutputFile -NoTypeInformation -Encoding UTF8

    [pscustomobject]@{
        ReportRefreshDate = $snapshotDate
        UserFile = $userOutputFile
        TrendFile = $trendOutputFile
        UserCount = $users.Count
        TrendRowCount = $trend.Count
    }
}
finally {
    if ($temporaryFolder -and (Test-Path -LiteralPath $temporaryFolder)) {
        Remove-Item -LiteralPath $temporaryFolder -Recurse -Force
    }
    if ($connectedHere) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
