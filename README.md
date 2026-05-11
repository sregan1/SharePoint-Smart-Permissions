# SharePoint Smart Permissions — SPFx Web Part

A SharePoint Framework (SPFx) web part that replicates the functionality of the SharePoint Smart Permissions desktop app — entirely in the browser, with no installation required. Add it to any SharePoint Online page and it runs as the signed-in user with zero extra configuration.

---

## Features

### 📊 Permissions Report
Scan a site, library, folder tree, or all site collections and export a colour-coded Excel workbook showing every permission break.

- **Scope options:** Site only · Libraries · Folders (with configurable depth) · All files and folders
- **All-sites scan:** Enumerate the entire tenant via SharePoint Search
- **Hidden libraries:** Optionally include hidden/NoCrawl lists
- **Excel export:** Two-sheet workbook — Summary (statistics) + Permissions (full detail), colour-coded by object type and role level, with auto-filter and frozen header row
- **Cancellable:** Abort a long-running scan at any time

### 🔍 Real-time Audit
Browse a document library live and inspect permissions on any folder or file without leaving the page.

- **Folder tree:** Lazy-loaded, expand-on-demand tree with **Unique** badges on items with broken inheritance
- **Permission panel:** See every user/group with access to the selected item, colour-coded by role level (Full Control / Edit / Read)
- **Expand SP groups:** Optionally expand SharePoint groups to show individual members inline
- **User access lookup:** Select any site user and instantly see every location they can access on the site

---

## Architecture

This web part is a full browser-side port of the WPF desktop app (`SharePointSmartPermissions`). No C# code is shared — the runtimes are entirely different — but the business logic maps 1:1.

| Layer | Desktop app | This web part |
|---|---|---|
| Language | C# .NET 9 | TypeScript |
| UI | WPF / XAML | React + Fluent UI v9 |
| SharePoint API | CSOM | SPFx `SPHttpClient` (REST) |
| Authentication | MSAL + Azure AD app reg | SPFx context (automatic) |
| Excel export | ClosedXML | ExcelJS |
| Distribution | Install EXE | Deploy `.sppkg` to App Catalog |

**Authentication note:** The web part runs as the currently signed-in SharePoint user — no Azure AD app registration is needed. It can only see sites and items that user has permission to access.

---

## Project structure

```
src/webparts/smartPermissions/
├── SmartPermissionsWebPart.ts          # SPFx entry point
├── SmartPermissionsWebPart.manifest.json
├── models/
│   └── models.ts                       # TypeScript interfaces (PermissionEntry, FolderFileNode, …)
├── services/
│   ├── SharePointService.ts            # All SharePoint REST calls (port of SharePointService.cs)
│   └── ExcelExportService.ts           # Excel workbook generation (port of ExcelExportService.cs)
└── components/
    ├── App.tsx                         # Root component — FluentProvider + view routing
    ├── HomeView.tsx                    # Landing page with two feature cards
    ├── PermissionsReportView.tsx       # Report scan, progress, export
    └── RealtimeAuditView.tsx           # Folder tree + user access tabs
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

> **Tip:** Test `SharePointService` methods first — open the browser console, confirm the REST calls return data from the site, then test the UI on top.

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
3. Check **Make this solution available to all sites** if you want tenant-wide deployment (recommended — `skipFeatureDeployment: true` is already set in `package-solution.json`).
4. Click **Deploy**.

### Add to a page

1. Navigate to any modern SharePoint page.
2. Edit the page → **+** → search for **Smart Permissions**.
3. Add and publish.

No property pane configuration is required — everything is driven from within the web part UI.

---

## Key implementation notes

### REST API vs CSOM
All SharePoint calls use `SPHttpClient` from `@microsoft/sp-http`. This mirrors the `HttpClient` pattern in the C# service and gives full control over the OData query structure. The `odata=nometadata` format (set automatically by `SPHttpClient.configurations.v1`) returns clean JSON arrays without `__metadata` or `results` wrappers.

### Two-pass fallback pattern
Every service method that loads role assignments uses the same pattern as the C# original: attempt the combined load (with `$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`); if it fails (e.g. permission denied, system folder with no backing list item), fall back to basic metadata only and continue.

### Role assignment normalisation
SPO REST returns `RoleDefinitionBindings` as a direct array with `odata=nometadata`. The `rdbArray()` helper also accepts `{ value: [] }` and `{ results: [] }` shapes so the service works correctly if a tenant's OData format differs.

### Recursive folder walk
`walkFolder` and `walkFoldersForUser` mirror the C# recursion exactly — skip `_*` system folders and the `Forms` folder (which has no backing list item), check `HasUniqueRoleAssignments` via `ListItemAllFields` expansion, and load role assignments only for items where it's `true`.

### Excel export
Uses [ExcelJS](https://github.com/exceljs/exceljs) (MIT, browser-compatible). The same colour scheme as the C# ClosedXML export: blue header, object-type badges (blue/cyan/yellow/grey), green/red/yellow role level highlights. Output is triggered as a browser download — no server required.

---

## Known limitations

- **Runs as the signed-in user.** If you need to scan sites the current user can't read, sign in with an account that has appropriate access.
- **Large tenants.** Search-based all-sites enumeration is capped at 500 results per page (paged automatically). Libraries with thousands of items will be slower than the CSOM desktop app due to browser-based REST pagination.
- **`Item` scope on large libraries.** Each item with unique permissions requires an additional REST call to load role assignments. Test against a representative library before running tenant-wide.
- **No persistent settings.** The desktop app persists tenant config to a local JSON file. This web part has no equivalent — URLs are entered per session.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `HTTP 401` on connect | User lacks read access to the site | Use an account with at least Read permission |
| `HTTP 403` on permissions panel | User lacks "Manage Permissions" | Sign in as a site owner or use an admin account |
| Empty site list (all-sites scan) | Search index hasn't crawled new sites | Wait for crawl, or enter the site URL directly |
| Excel download doesn't start | Browser popup blocker | Allow popups from the SharePoint domain |
| `gulp serve` certificate warning | Self-signed dev cert | Accept the certificate at `https://localhost:4321` first, then reload the workbench |
