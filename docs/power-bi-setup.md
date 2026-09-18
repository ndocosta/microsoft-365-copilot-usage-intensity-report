# Power BI setup

## Prerequisites

- Power BI Desktop with Power BI Project (PBIP) support
- Read access to the SharePoint site containing the snapshots
- At least one successful export from either deployment option

## Configure the data source

1. Open `CopilotAdoption.pbip`.
2. In **Transform data > Edit parameters**, set:

| Parameter | Example |
| --- | --- |
| `SharePointSiteUrl` | `https://contoso.sharepoint.com/sites/CopilotReports` |
| `SharePointLibraryName` | `UsageData` |

3. Apply the changes.
4. Authenticate to SharePoint with an organizational account that can read the
   site.
5. Refresh the report.

Use the site URL, not the library URL. The library parameter is its display
name, not a folder path.

## How the model selects data

- `Users` loads the newest `CopilotUsers-D28-*.csv` snapshot.
- `Trend` loads the newest Copilot trend snapshot.
- Entra attributes come from the newest directory snapshot.
- Usage and directory rows join on `UserPrincipalName`.
- The header displays the 28-day analysis interval based on the latest Microsoft
  report refresh date.

Repeated exports before Microsoft refreshes the source report replace the same
logical snapshot. This keeps retries idempotent while daily runs retain a longer
history in SharePoint.

## Publish and refresh

After publishing to the Power BI service:

1. Configure SharePoint Online credentials for the semantic model.
2. Confirm the refresh identity has read access to the site.
3. Schedule the Power BI refresh after the export normally completes.
4. Allow additional time for the Azure Automation job or local task to finish.

The Power BI identity does not need the exporter's Graph permissions. The export
identity does not need access to the Power BI workspace.

## Customize safely

The PBIP format keeps report and semantic-model definitions as source files.
Create a branch before changing measures, Power Query, or visual definitions,
and test a full refresh with representative data before publishing.
