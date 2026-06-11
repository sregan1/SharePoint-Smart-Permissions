# SharePoint Smart Permissions — Permissions Auditing Web Part

[![Website](https://img.shields.io/badge/Website-sharepointsmartsolutions.com-blue)](https://sharepointsmartsolutions.com/smart-permissions) [![User Guide](https://img.shields.io/badge/User%20Guide-Read%20Now-green)](USER-GUIDE.md) [![Download](https://img.shields.io/badge/Download-Latest%20Release-CA5010?logo=github&logoColor=white)](../../releases/latest) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A free, open-source SPFx web part that gives SharePoint site owners a real-time, browser-based view of who has access to what — no PowerShell, no third-party software, and no elevated service accounts required.

![SPFx](https://img.shields.io/badge/SPFx-1.21.1-0078D4?logo=microsoft&logoColor=white) ![React](https://img.shields.io/badge/React-17.0.1-61DAFB?logo=react&logoColor=black) ![TypeScript](https://img.shields.io/badge/TypeScript-5.3-3178C6?logo=typescript&logoColor=white) ![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft&logoColor=white) ![ExcelJS](https://img.shields.io/badge/ExcelJS-4.4-107C10)

![Home screen showing three feature cards: Permissions Explorer, Permissions Report, and User Access](docs/screenshots/01_home.png)

---

## Features

> **Requires Site Owner access.** All three tools read SharePoint role assignments, which requires the Manage Permissions right. Anyone without Site Owner access (Members, Visitors, Limited Access users, guests) will see the web part but the feature cards will be disabled with a clear explanation.

### Permissions Explorer

Interactively browse a document library tree and inspect live permissions on any folder or file.

| Feature | Description |
|---|---|
| **Live permission lookup** | Click any folder or file to see its full role assignment list in real time |
| **Inherited vs. unique** | Clearly labels whether permissions are inherited or uniquely assigned |
| **Group member expansion** | Expand SharePoint groups, Security groups, and M365 groups to see individual members |
| **Automatic external user detection** | Background scan flags items where external (#EXT#) accounts have explicit access — no clicking required |
| **Three folder indicators** | Circle arrow = unique permissions below; triangle arrow = external access below; both together = both |
| **External user email decoding** | Decoded email shown beneath the display name in the permissions panel |
| **Tree filters** | Toggle to show only unique-permission items or only items with external user access |

![Permissions Explorer with a folder selected and the permissions panel showing group members](docs/screenshots/05_explorer.png)

---

### Permissions Report

Scan a site's lists, libraries, folders, and files and export a color-coded Excel workbook of every unique permission assignment.

| Feature | Description |
|---|---|
| **Configurable scan depth** | Choose Site only, Libraries, or full depth (folders and files) |
| **Full list coverage** | Scans all visible lists — generic lists, Site Pages, calendars — not just document libraries |
| **Subsite scanning** | Optionally recurse into every subsite below the selected site |
| **In-browser results table** | Browse results directly: sortable, paginated, with expandable per-object permission details |
| **Report compare** | Diff two saved reports to see permissions added, removed, or changed between scans |
| **Folder depth limit** | Cap recursion at 1–5 levels to control scan time on large sites |
| **Group member expansion** | Expand all group types into individual user rows in the export |
| **External user filter** | Scope the scan and export to external accounts only |
| **Hidden-from-search flag** | NoCrawl libraries are included and badged instead of silently skipped |
| **Excel export** | In-browser `.xlsx` workbook with color-coded permission levels and an Access Via column |
| **CSV export** | Plain-text alternative for scripted processing |
| **Progress bar and timer** | Concurrent, throttling-aware scan with elapsed timer and live item count |

![Permissions Report configuration screen](docs/screenshots/02_report_config.png)

![Permissions Report Excel export preview showing color-coded rows](docs/screenshots/04b_report_export.png)

---

### User Access

Look up any user and see every location they can access on a site, with their exact permission level at each location.

| Feature | Description |
|---|---|
| **Searchable user picker** | Combobox with live filtering across all site users |
| **Tenant-wide people search** | Type 3+ characters to also find users not yet in the site's user list ("Not in this site") |
| **Site-level detection** | Immediately identifies when a user has full-site access via M365 Group or SP group membership |
| **M365 Group support** | Detects membership via SharePoint REST, site GroupId, and Graph `transitiveMemberOf` |
| **Sortable results table** | Click any column header to sort |
| **Export to Excel / CSV** | Download the results for any user |
| **Scan history** | Past scans persist in IndexedDB — re-export without re-scanning |

![User Access results showing accessible locations for a selected user](docs/screenshots/07_user_access_complete.png)

---

## Screenshots

| | |
|---|---|
| **Home** | **Permissions Report — Configuration** |
| ![Home screen](docs/screenshots/01_home.png) | ![Permissions Report configuration](docs/screenshots/02_report_config.png) |
| **Permissions Report — Excel Export** | **Permissions Explorer** |
| ![Permissions Report Excel export](docs/screenshots/04b_report_export.png) | ![Permissions Explorer](docs/screenshots/05_explorer.png) |
| **User Access — Results** | **Settings** |
| ![User Access results](docs/screenshots/07_user_access_complete.png) | ![Settings](docs/screenshots/09_settings.png) |

---

## Installation (No Build Required)

1. Go to the [Releases](../../releases/latest) page and download `smart-permissions.sppkg`.
2. Open the **SharePoint Admin Center** → **More features** → **Apps** → **App Catalog**.
3. Upload `smart-permissions.sppkg` to the **Apps for SharePoint** library.
4. When prompted, click **Deploy**. The *"Make this solution available to all sites"* checkbox controls where the web part can be placed — tick it for a tenant-wide deployment.
5. Navigate to the SharePoint page where you want to add the web part, click **Edit**, and add **Smart Permissions** from the web part picker.
6. *(Optional)* Approve the Graph API permissions in **SharePoint Admin Center → Advanced → API access** to enable group member expansion (see [Graph API Permissions](#graph-api-permissions) below).

---

## Prerequisites (for Development Only)

| Requirement | Detail |
|---|---|
| **Node.js** | 18.x (`>=18.17.1 <19.0.0`) |
| **gulp-cli** | Install globally: `npm install -g gulp-cli` |
| **SharePoint** | Online (Microsoft 365) |
| **SPFx** | 1.21.1 |
| **Permissions to deploy** | Site collection administrator or above |

---

## Development Setup

```bash
# Install dependencies
npm install

# Edit config/serve.json and set initialPage to your hosted workbench URL:
# "initialPage": "https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx"

# Start the local dev server (opens hosted workbench)
gulp serve
```

The local workbench at `https://localhost:4321/temp/workbench.html` does not have SharePoint REST API access. Use the **hosted workbench** URL above for full functionality.

---

## Build & Deploy

```bash
# Production bundle (minified, ship mode)
gulp bundle --ship

# Create the .sppkg deployment package
gulp package-solution --ship

# Or run both in one step:
npm run ship
```

The package is written to `sharepoint/solution/smart-permissions.sppkg`.

**Deploy to SharePoint:**

1. Upload `smart-permissions.sppkg` to the tenant or site App Catalog.
2. Click **Deploy** when prompted.
3. Add the web part to a page and optionally approve Graph permissions (see below).

---

## Configuration

To change web part settings, put the page in **Edit** mode, click the web part pencil icon, and use the property pane.

**General**

| Setting | Default | Description |
|---|---|---|
| **Default view on open** | Home | The screen shown when the web part first loads. Options: Home, Permissions Report, Permissions Explorer, User Access |

---

## Graph API Permissions

The package declares one optional `webApiPermissionRequests` entry. The core tools work without it, but the following feature requires approval:

| Permission | Enables | Required? |
|---|---|---|
| `GroupMember.Read.All` | Expanding Security group and M365 group members in the Permissions Report and Explorer | Optional |

**To approve (SharePoint Admin Center):**

1. Go to **SharePoint Admin Center → Advanced → API access** (or navigate directly to `https://<tenant>-admin.sharepoint.com/_layouts/15/online/ManageApiPermissions.aspx`).
2. Find the pending request for Smart Permissions and click **Approve**.

> **What approved SPFx permissions look like:** Once approved, `GroupMember.Read.All` appears in the list under **Microsoft Graph** with a blank app name (`—`) and no date. This is normal for SPFx permissions. If you see the entry with consent granted, the permission is active.

**To approve via PnP PowerShell:**

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive

# List pending requests
Get-PnPTenantServicePrincipalPermissionRequests

# Approve by ID
Approve-PnPTenantServicePrincipalPermissionRequest -RequestId "<guid>"
```

**Without `GroupMember.Read.All`:** SharePoint group members expand normally. Security groups and M365 groups show a warning banner instead of member lists.

---

## Project Structure

```
src/webparts/smartPermissions/
├── components/
│   ├── App.tsx                        # Root component — routing, permission check, theme wiring
│   ├── HomeView.tsx                   # Home screen with feature cards and non-owner warning
│   ├── PermissionsReportView.tsx      # Report configuration, results table, history, compare
│   ├── PermissionsExplorerView.tsx    # Interactive folder/file tree and permission panel
│   ├── UserAccessView.tsx             # Per-user access scan, history, export
│   ├── SettingsView.tsx               # Full-page settings screen
│   └── shared/                        # Shared UI: PermTable, role badge colors, site-owner links
├── services/
│   ├── SharePointService.ts           # Facade over the sp/ modules (stable public API)
│   ├── sp/                            # API client, report scan, explorer, groups, user access
│   └── ExcelExportService.ts          # ExcelJS workbook generation (lazy-loaded chunk)
├── utils/
│   ├── notifications.ts               # Guarded browser-notification helpers
│   └── reportDiff.ts                  # Pure diff between two stored reports
├── models/
│   └── models.ts                      # Shared TypeScript interfaces
└── SmartPermissionsWebPart.ts         # SPFx entry point, property pane, theme wiring

config/
├── package-solution.json              # Solution ID, version, Graph permission requests
└── serve.json                         # Local dev server config — set initialPage here

docs/
├── generate-screenshots.js            # Puppeteer script to regenerate all screenshots
└── screenshots/                       # Auto-generated UI screenshots (git-ignored)

scripts/
└── Provision-SmartPermissions.ps1     # PnP PowerShell provisioning script

sharepoint/solution/
└── smart-permissions.sppkg            # Pre-built deployment package
```

---

## Key Dependencies

| Package | Purpose |
|---|---|
| `@microsoft/sp-webpart-base` | SPFx web part base class and framework integration |
| `@fluentui/react-components` | Fluent UI v9 — components and theme tokens |
| `react` / `react-dom` | UI rendering (v17) |
| `exceljs` | In-browser Excel workbook generation for report exports |

---

## Troubleshooting

**"Feature cards are grayed out and I see a warning banner"** — The signed-in account does not have the Manage Permissions right. Site Owner access or higher is required to read role assignments. Ask a site owner to either grant you Owner access or run the tool on your behalf.

**"Security groups / M365 groups show a warning instead of expanding"** — The `GroupMember.Read.All` Graph permission has not been approved. A SharePoint or Global Administrator must approve it in **SharePoint Admin Center → Advanced → API access**. SharePoint group expansion works without this permission.

**"gulp serve opens but the web part shows no data"** — The local workbench (`localhost:4321`) cannot authenticate to SharePoint REST. Switch to the hosted workbench: edit `config/serve.json` and set `initialPage` to `https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`.

**"npm install fails" or build errors about Node version** — This project requires Node 18.x exactly (`>=18.17.1 <19.0.0`). Run `node --version` to confirm. Use `nvm` or `nvm-windows` to switch versions.

**"The scan takes a very long time"** — User Access and Permissions Report scan time scales with the number of unique permission assignments on the site. Use the scan depth and folder depth settings in the Permissions Report to limit scope, or run the scan against a specific library.

---

## Limitations

- All three tools require the **Manage Permissions** right (Site Owner access). Members, Visitors, Limited Access users, and guests cannot use any feature.
- Runs entirely as the signed-in user — results reflect that user's view. An account with read access to only part of a site will produce an incomplete scan.
- In-browser Excel generation via ExcelJS may slow down or fail on very large sites (tens of thousands of unique permission rows) due to browser memory limits.
- Scan history (User Access) persists in **IndexedDB** in the browser. Clearing browser data removes all saved scan results.
---

## License

[MIT](LICENSE) © 2026 Sean Regan
