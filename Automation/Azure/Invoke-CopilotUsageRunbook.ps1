#requires -Version 7.4
#requires -Modules Az.Accounts

<#
.SYNOPSIS
Exports Microsoft 365 Copilot usage intensity snapshots from Azure Automation.

.DESCRIPTION
Uses the Azure Automation account's system-assigned managed identity to read
Microsoft 365 Copilot usage reports and Microsoft Entra user properties, then
uploads normalized CSV snapshots to a selected SharePoint document library.

The runbook reads its default settings from these Azure Automation variables:

- CopilotUsageTenantId
- CopilotUsageSharePointDriveId
- CopilotUsagePeriod

Parameters override the corresponding Automation variables, which is useful for
manual tests. No certificate, client secret, or persistent local storage is
required. Temporary files are removed at the end of every job.

.PARAMETER TenantId
Microsoft Entra tenant ID. When omitted, the value is read from
CopilotUsageTenantId.

.PARAMETER SharePointDriveId
Microsoft Graph drive ID of the destination SharePoint document library. When
omitted, the value is read from CopilotUsageSharePointDriveId.

.PARAMETER Period
Microsoft 365 Copilot reporting period. Only D28 is supported by the included
Power BI model. When omitted, the value is read from CopilotUsagePeriod.

.OUTPUTS
PSCustomObject with the report refresh date and exported row counts.

.NOTES
The Automation account managed identity requires these Microsoft Graph
application permissions:

- Reports.Read.All
- User.Read.All
- Lists.SelectedOperations.Selected

It must also have the write role on the destination document library.
#>

[CmdletBinding()]
param(
    [string]$TenantId,

    [string]$SharePointDriveId,

    [ValidateSet('D28')]
    [string]$Period
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

function Get-RequiredAutomationSetting {
    param(
        [string]$ParameterValue,

        [Parameter(Mandatory)]
        [string]$VariableName
    )

    if (-not [string]::IsNullOrWhiteSpace($ParameterValue)) {
        return $ParameterValue
    }

    $value = Get-AutomationVariable -Name $VariableName
    if ([string]::IsNullOrWhiteSpace([string]$value)) {
        throw "Automation variable '$VariableName' is required."
    }

    return [string]$value
}

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

function Invoke-GraphJsonRequest {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [string]$Body
    )

    $parameters = @{
        Method = $Method
        Uri = $Uri
        Authentication = 'Bearer'
        Token = $script:GraphAccessToken
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $parameters.Body = $Body
        $parameters.ContentType = 'application/json'
    }

    Invoke-RestMethod @parameters
}

function Initialize-DriveFolder {
    param(
        [Parameter(Mandatory)]
        [string]$DriveId,

        [Parameter(Mandatory)]
        [string]$FolderName
    )

    $encodedFolder = [uri]::EscapeDataString($FolderName)
    $folderUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encodedFolder"

    try {
        $null = Invoke-GraphJsonRequest -Method GET -Uri $folderUri
        return
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode
        if ($statusCode -ne [System.Net.HttpStatusCode]::NotFound) {
            throw
        }
    }

    $body = @{
        name = $FolderName
        folder = @{}
        '@microsoft.graph.conflictBehavior' = 'fail'
    } | ConvertTo-Json -Depth 5

    $null = Invoke-GraphJsonRequest `
        -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children" `
        -Body $body
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

    $null = Invoke-RestMethod `
        -Method PUT `
        -Uri $uploadUri `
        -Authentication Bearer `
        -Token $script:GraphAccessToken `
        -InFile $FilePath `
        -ContentType 'text/csv' `
        -ErrorAction Stop
}

function Export-CopilotSnapshot {
    param(
        [Parameter(Mandatory)]
        [string]$OutputRoot,

        [Parameter(Mandatory)]
        [ValidateSet('D28')]
        [string]$ReportPeriod
    )

    $userFolder = Join-Path $OutputRoot 'UserSnapshots'
    $trendFolder = Join-Path $OutputRoot 'TrendSnapshots'
    [void](New-Item -ItemType Directory -Path $userFolder -Force)
    [void](New-Item -ItemType Directory -Path $trendFolder -Force)

    $rawUserFile = Join-Path $OutputRoot 'UserDetail.csv'
    $rawTrendFile = Join-Path $OutputRoot 'UserCountTrend.csv'
    $userUri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod',version='v2')"
    $trendUri = "https://graph.microsoft.com/v1.0/copilot/reports/getMicrosoft365CopilotUserCountTrend(period='$ReportPeriod',version='v2')"

    Invoke-WebRequest `
        -Method GET `
        -Uri $userUri `
        -Authentication Bearer `
        -Token $script:GraphAccessToken `
        -OutFile $rawUserFile `
        -ErrorAction Stop
    Invoke-WebRequest `
        -Method GET `
        -Uri $trendUri `
        -Authentication Bearer `
        -Token $script:GraphAccessToken `
        -OutFile $rawTrendFile `
        -ErrorAction Stop

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

    $reportRefreshDate = [datetime](Get-ReportValue -Row $rawUsers[0] -Names @('Report Refresh Date'))
    $snapshotDate = $reportRefreshDate.ToString('yyyy-MM-dd')
    $reportPeriodDays = [int]($ReportPeriod.Substring(1))

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
            ReportPeriod                        = $reportPeriodDays
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
            ReportPeriod                   = $reportPeriodDays
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

    $userOutputFile = Join-Path $userFolder "CopilotUsers-$ReportPeriod-$snapshotDate.csv"
    $trendOutputFile = Join-Path $trendFolder "CopilotTrend-$ReportPeriod-$snapshotDate.csv"
    $users | Export-Csv -LiteralPath $userOutputFile -NoTypeInformation -Encoding utf8
    $trend | Export-Csv -LiteralPath $trendOutputFile -NoTypeInformation -Encoding utf8

    [pscustomobject]@{
        ReportRefreshDate = $snapshotDate
        UserFile = $userOutputFile
        TrendFile = $trendOutputFile
        UserCount = $users.Count
        TrendRowCount = $trend.Count
    }
}

function Export-DirectorySnapshot {
    param(
        [Parameter(Mandatory)]
        [string]$OutputRoot
    )

    $directoryFolder = Join-Path $OutputRoot 'DirectorySnapshots'
    [void](New-Item -ItemType Directory -Path $directoryFolder -Force)
    $users = [System.Collections.Generic.List[object]]::new()
    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName,department,jobTitle,officeLocation,city,state,country,usageLocation,companyName,accountEnabled,userType'

    while ($uri) {
        $response = Invoke-GraphJsonRequest -Method GET -Uri $uri
        foreach ($user in $response.value) {
            $users.Add($user)
        }
        $nextLink = $response.PSObject.Properties['@odata.nextLink']
        $uri = if ($null -ne $nextLink) { [string]$nextLink.Value } else { $null }
    }

    if ($users.Count -eq 0) {
        throw 'Microsoft Entra ID returned no users.'
    }

    $snapshotDate = [DateTime]::UtcNow.ToString('yyyy-MM-dd')
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
        Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding utf8

    [pscustomobject]@{
        SnapshotDate = $snapshotDate
        DirectoryFile = $outputFile
        UserCount = $users.Count
    }
}

$TenantId = Get-RequiredAutomationSetting `
    -ParameterValue $TenantId `
    -VariableName 'CopilotUsageTenantId'
$SharePointDriveId = Get-RequiredAutomationSetting `
    -ParameterValue $SharePointDriveId `
    -VariableName 'CopilotUsageSharePointDriveId'
$Period = Get-RequiredAutomationSetting `
    -ParameterValue $Period `
    -VariableName 'CopilotUsagePeriod'
if ($Period -ne 'D28') {
    throw "Unsupported reporting period '$Period'. This solution supports D28."
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("CopilotUsage-" + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $temporaryRoot)

try {
    Write-Information 'Starting Microsoft 365 Copilot usage export.'
    Disable-AzContextAutosave -Scope Process | Out-Null
    $azureContext = (
        Connect-AzAccount `
            -Identity `
            -Tenant $TenantId `
            -SkipContextPopulation `
            -Scope Process `
            -ErrorAction Stop
    ).Context
    $script:GraphAccessToken = (
        Get-AzAccessToken `
            -ResourceTypeName MSGraph `
            -TenantId $TenantId `
            -AsSecureString `
            -DefaultProfile $azureContext `
            -ErrorAction Stop
    ).Token

    $reports = Export-CopilotSnapshot -OutputRoot $temporaryRoot -ReportPeriod $Period
    Write-Information ("Exported {0} Copilot users for report refresh date {1}." -f $reports.UserCount, $reports.ReportRefreshDate)

    $directory = Export-DirectorySnapshot -OutputRoot $temporaryRoot
    Write-Information ("Exported {0} Microsoft Entra users." -f $directory.UserCount)

    Send-FileToSharePoint `
        -DriveId $SharePointDriveId `
        -FolderName 'UserSnapshots' `
        -FilePath $reports.UserFile
    Send-FileToSharePoint `
        -DriveId $SharePointDriveId `
        -FolderName 'TrendSnapshots' `
        -FilePath $reports.TrendFile
    Send-FileToSharePoint `
        -DriveId $SharePointDriveId `
        -FolderName 'DirectorySnapshots' `
        -FilePath $directory.DirectoryFile

    Write-Information 'Uploaded all CSV snapshots to SharePoint successfully.'
    [pscustomobject]@{
        ReportRefreshDate = $reports.ReportRefreshDate
        CopilotUsers = $reports.UserCount
        TrendRows = $reports.TrendRowCount
        DirectoryUsers = $directory.UserCount
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
