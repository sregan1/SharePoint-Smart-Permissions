# SharePoint Smart Permissions — SPFx Web Part

A SharePoint Framework (SPFx) web part that provides browser-based permissions auditing for SharePoint Online. Add it to any modern SharePoint page and it runs as the signed-in user with zero extra configuration.

---

## Features

### Permissions Report
Scan a site, library, folder tree, or all site collections and export a colour-coded Excel workbook showing every unique permission assignment.

- **Scope options:** Site only · Libraries · Folders (with configurable depth) · Files & Folders
- **All-sites scan:** Enumerate the entire tenant via SharePoint Search (root site only)
- **Hidden/system libraries:** Optionally include hidden and NoCrawl lists via global settings
- **Excel export:** Full permissions workbook, colour-coded by object type and role level, with auto-filter and frozen header row
- **Cancellable:** Abort a long-running scan at any time

### Permissions Explorer
Browse a document library interactively and inspect permissions on any folder or file in real time.

- **Folder/file tree:** Lazy-loaded, expand-on-demand tree with Unique badges on items with broken inheritance
- **Down-arrow indicator:** Folders containing unique-permission items deeper in their tree are flagged at load time
- **Permission panel:** Every user/group with access to the selected item, colour-coded by role level
- **Expand SP groups:** Optionally expand SharePoint groups to show individual members inline
- **Show parent permissions:** For inherited items, display the nearest ancestor's unique permissions

### User Access
Select any user on a site and see every location they can access, with their exact permission level at each location.

- **Full site access detection:** Detects owner/full-control accounts and short-circuits the full scan
- **Elapsed timer and cancel:** Long scans show a live timer and can be cancelled at any time
- **Partial results:** Cancelling mid-scan returns whatever has been found so far

---

## Architecture

| Layer | Detail |
|---|---|
| Language | TypeScript |
| UI | React 17 + Fluent UI v9 (`@fluentui/react-components`) |
| SharePoint API | SPFx `SPHttpClient` (REST / OData) |
| Authentication | SPFx context — runs as the signed-in user, no app registration needed |
| Excel export | ExcelJS (MIT, browser-compatible) |
| Distribution | Deploy `.sppkg` to App Catalog |

**Authentication note:** The web part runs as the currently signed-in SharePoint user. It can only see sites and items that user has permission to access.

---

## Project structure

```
src/
├── global.d.ts                             # Module declarations for PNG/JPG/SVG imports
└── webparts/smartPermissions/
    ├── SmartPermissionsWebPart.ts          # SPFx entry point + property pane
    ├── SmartPermissionsWebPart.manifest.json
    ├── models/
    │   └── models.ts                       # Shared TypeScript interfaces
    ├── services/
    │   ├── SharePointService.ts            # All SharePoint REST calls
    │   └── ExcelExportService.ts           # Excel workbook generation (ExcelJS)
    └── components/
        ├── App.tsx                         # Root — FluentProvider, banner, view routing, global settings
        ├── HomeView.tsx                    # Landing page with three feature cards
        ├── PermissionsReportView.tsx       # Report scan, progress, export
        ├── PermissionsExplorerView.tsx     # Interactive folder tree + permission panel
        └── UserAccessView.tsx              # Per-user access lookup

docs/
├── UserGuide.md                            # End-user documentation
├── screenshot.js                           # Generates docs/screenshots/ via puppeteer-core
└── generate-word.js                        # Generates UserGuide.docx via the docx package
```

---

## Prerequisites

| Tool | Version |
|---|---|
| Node.js | 18.x (engines field enforces `>=18.17.1 <19.0.0`) |
| npm | 9+ (bundled with Node 18) |
| gulp-cli | `npm install -g gulp-cli` |

---

## Getting started

### 1. Install dependencies

```powershell
cd c:\Development\SharePointSmartPermissionsWebPart
npm install
```

### 2. Configure the workbench URL

Copy `config/serve.json.template` to `config/serve.json` and set `initialPage` to your SharePoint Online workbench URL:

```json
"initialPage": "https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/workbench.aspx"
```

`config/serve.json` is gitignored — each developer maintains their own copy.

### 3. Run locally

```powershell
gulp serve
```

This starts the HTTPS dev server on port 4321 and opens the SharePoint Workbench. Accept the self-signed certificate if prompted. Add the **Smart Permissions** web part to the workbench canvas to start testing.

---

## Deploying to SharePoint

### Build for production

```powershell
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/smart-permissions.sppkg`.

### Deploy to the App Catalog

1. Open your SharePoint Admin Center → **More features** → **Apps** → **App Catalog**.
2. Upload `smart-permissions.sppkg` to **Apps for SharePoint**.
3. Check **Make this solution available to all sites** if you want tenant-wide deployment (`skipFeatureDeployment: true` is already set in `package-solution.json`).
4. Click **Deploy**.

### Add to a page

1. Navigate to any modern SharePoint page.
2. Edit the page → **+** → search for **Smart Permissions**.
3. Add and publish.

---

## Generating documentation

The `docs/` folder contains two Node.js scripts for regenerating the documentation assets.

### Screenshots

```powershell
node docs/screenshot.js
```

Produces `docs/screenshots/*.png` — one screenshot per screen, populated with dummy data. Requires a local Chrome installation. The script auto-detects Chrome on Windows, macOS, and Linux. Override with the `CHROME_PATH` environment variable if Chrome is installed in a non-standard location:

```powershell
$env:CHROME_PATH = "C:\path\to\chrome.exe"; node docs/screenshot.js
```

### Word user guide

```powershell
node docs/generate-word.js
```

Produces `docs/UserGuide.docx` from the same content as `docs/UserGuide.md`, with screenshots embedded. Both output files are gitignored — regenerate them locally as needed.

---

## Property pane

The web part exposes one property pane setting (accessible in SharePoint edit mode):

| Setting | Description |
|---|---|
| **Default view on open** | Choose which screen loads when the web part first renders: Home, Permissions Report, Permissions Explorer, or User Access |

Global settings visible to all users (the gear icon in the banner) are UI-only and not persisted across sessions:

| Setting | Description |
|---|---|
| **Include system and hidden libraries** | When checked, Permissions Explorer and User Access include hidden/system libraries such as Style Library, Form Templates, and Site Assets |

---

## Key implementation notes

### REST API pattern
All SharePoint calls use `SPHttpClient` from `@microsoft/sp-http`. `odata=nometadata` format (set automatically by `SPHttpClient.configurations.v1`) returns clean JSON without `__metadata` or `results` wrappers. The `valueArray()` helper normalises across both formats for resilience.

### Two-pass fallback
Every method that loads role assignments attempts the full combined query first (`$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`). On failure (permission denied, system folder with no backing list item) it falls back to basic metadata and continues rather than aborting.

### Recursive folder walk
`walkFolder` and `walkFoldersForUser` skip `_*` system folders and the `Forms` folder (no backing list item). They check `HasUniqueRoleAssignments` via `ListItemAllFields` expansion and only fetch full role assignments for items where it is `true`.

### Root-folder pre-fetch
On library load, `PermissionsExplorerView` immediately fires a background fetch of each root folder's direct children. If any child has `HasUniqueRoleAssignments: true`, the parent folder is marked with the down-arrow indicator without requiring the user to expand it first.

### AMD / property pane
`@microsoft/sp-property-pane` is an AMD external in the SPFx build. To avoid a CSP violation from a runtime `require()` call, the property pane dropdown is constructed by directly building the `IPropertyPaneField` descriptor object (type discriminant `6` = `PropertyPaneFieldType.Dropdown`) rather than importing `PropertyPaneDropdown`. This keeps `@microsoft/sp-property-pane` entirely off the static dependency list.

### Excel export
Uses [ExcelJS](https://github.com/exceljs/exceljs) (MIT, browser-compatible). Output is triggered as a browser `Blob` download — no server required.

---

## Known limitations

- **Runs as the signed-in user.** To scan sites the current user cannot read, sign in with an account that has appropriate access.
- **Large tenants.** Search-based all-sites enumeration is paged at 500 results per request. Libraries with thousands of items are slower than a server-side tool due to browser REST pagination.
- **Item scope on large libraries.** Each item with unique permissions requires an additional REST call. Test against a representative library before running tenant-wide.
- **Global settings are session-only.** The "Include system and hidden libraries" toggle is not persisted. It resets to off on each page load.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `HTTP 401` on connect | User lacks read access to the site | Use an account with at least Read permission |
| `HTTP 403` on permissions panel | User lacks Manage Permissions | Sign in as a site owner or admin |
| Empty site list (all-sites scan) | Search index hasn't crawled new sites yet | Wait for crawl, or enter the site URL directly |
| Excel download doesn't start | Browser popup blocker | Allow popups from the SharePoint domain |
| `gulp serve` certificate warning | Self-signed dev cert | Accept the certificate at `https://localhost:4321` first, then reload the workbench |

---

## License

[MIT](LICENSE)
