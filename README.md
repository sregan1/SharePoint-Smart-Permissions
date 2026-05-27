# SharePoint Smart Permissions

A free, open source browser-based SharePoint Online permissions auditing tool delivered as an SPFx web part. Gives site owners, administrators, and compliance teams a real-time view of who has access to what — with no PowerShell, no third-party software, and no elevated service accounts required.

<p>
  <a href="https://sharepointsmartsolutions.com/smart-permissions"><img src="https://img.shields.io/badge/Website-sharepointsmartsolutions.com-0078d4?style=for-the-badge&logo=microsoftsharepoint&logoColor=white" alt="Website"/></a>
  &nbsp;
  <a href="https://github.com/sregan1/SharePoint-Smart-Permissions/blob/main/docs/UserGuide.md"><img src="https://img.shields.io/badge/User_Guide-View_Documentation-107c10?style=for-the-badge&logo=readthedocs&logoColor=white" alt="User Guide"/></a>
</p>

---

## Features

### Main Tools

| Tool | Description |
|------|-------------|
| **Permissions Report** | Scans a site's libraries, folders, and files and exports a color-coded Excel workbook of every unique permission assignment. Supports full-tenant scans and optional group member expansion. |
| **Permissions Explorer** | Interactively browse a document library tree and inspect live permissions on any folder or file. Shows inherited vs. unique permissions and supports expanding group members (SharePoint groups, Security groups, M365 groups). |
| **User Access** | Look up any user and see every location they can access, with their exact permission level at each location. Detects full-site owner access immediately. |

### More Tools

| Tool | Description |
|------|-------------|
| **Permission Groups** | Shows all SharePoint groups on the site and their members. Quickly find who belongs to Owners, Members, Visitors, and any custom groups. |
| **External Users** | Lists all external (#EXT#) accounts that have been granted access to the site, including which SharePoint groups they belong to. Includes a one-click **Check Access** shortcut per user. |
| **Broken Inheritance Finder** | Scans every library, folder, and file and highlights every item that has had its permissions explicitly changed from its parent — the quickest way to find permission sprawl. |
| **Sharing Links** | Lists all sharing links across all document libraries — internal, external, and anonymous — so you can see at a glance what has been shared outside the site. Requires `Sites.Read.All`. |
| **Anonymous Access Summary** | Enumerates all anonymous and org-wide sharing links using the Microsoft Graph API, grouped by library and showing expiry dates where set. Requires `Sites.Read.All`. |

---

## Screenshots

<table>
<tr>
  <td align="center"><strong>Home</strong></td>
  <td align="center"><strong>Permissions Report — Configuration</strong></td>
</tr>
<tr>
  <td><img src="docs/screenshots/01_home.png" width="460" alt="Home screen"/></td>
  <td><img src="docs/screenshots/02_report_config.png" width="460" alt="Permissions Report configuration"/></td>
</tr>
<tr>
  <td align="center"><strong>Permissions Report — Export to Excel</strong></td>
  <td align="center"><strong>Permissions Explorer</strong></td>
</tr>
<tr>
  <td><img src="docs/screenshots/04b_report_export.png" width="460" alt="Permissions Report complete with Export to Excel"/></td>
  <td><img src="docs/screenshots/05_explorer.png" width="460" alt="Permissions Explorer"/></td>
</tr>
<tr>
  <td align="center"><strong>User Access</strong></td>
  <td align="center"><strong>Settings</strong></td>
</tr>
<tr>
  <td><img src="docs/screenshots/07_user_access_complete.png" width="460" alt="User Access results"/></td>
  <td><img src="docs/screenshots/09_settings.png" width="460" alt="Settings"/></td>
</tr>
</table>

---

## Security & Privacy

- Runs entirely in the user's browser as the signed-in identity — no elevated permissions
- Read-only: never creates, modifies, or deletes any SharePoint content
- No external services: all data stays within the Microsoft 365 tenant
- No data storage: results exist only in the browser session

---

## Technology

- **SPFx 1.21.1** · React 17 · TypeScript
- **Fluent UI v9** (`@fluentui/react-components`) — theme tokens driven by the SharePoint site theme
- **SharePoint REST API** — all permission data is read via standard SPO REST endpoints
- **Microsoft Graph API** — used for group member expansion (`GroupMember.Read.All`, optional) and sharing link enumeration (`Sites.Read.All`, optional)
- **ExcelJS** — in-browser Excel workbook generation

---

## Prerequisites (if building) - EXE available for download

- Node.js 18.x
- `gulp-cli` installed globally (`npm install -g gulp-cli`)
- A SharePoint Online tenant for testing

---

## Getting Started

```bash
# Install dependencies
npm install

# Start the local workbench
gulp serve
```

The local workbench opens at `https://localhost:4321/temp/workbench.html`. For full SharePoint REST API access, use the hosted workbench at `https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`.

---

## Download or Build & Package

If you don't want to build from source, you can download the pre-built `smart-permissions.sppkg` directly from the [Releases](https://github.com/sregan1/SharePoint-Smart-Permissions/releases) page and skip straight to uploading it to your App Catalog.

To build yourself:

```bash
# Production bundle
gulp bundle --ship

# Create the .sppkg deployment package
gulp package-solution --ship
```

The package is written to `sharepoint/solution/smart-permissions.sppkg`. Upload it to the SharePoint App Catalog. Optionally, approve the Graph permission requests in the SharePoint Admin Center to enable group member expansion and sharing link features (see below).

---

## Graph API Permissions (optional)

The package declares two `webApiPermissionRequests` entries:

| Permission | Purpose | Required? |
|------------|---------|-----------|
| `GroupMember.Read.All` | Expand Security group and M365 group members in Permissions Report and Explorer | Optional |
| `Sites.Read.All` | Enumerate sharing links in Sharing Links and Anonymous Access Summary | Optional |

Both permissions are **optional** — all other tools work without them:
- Without `GroupMember.Read.All`: Security groups and M365 groups cannot be expanded to list individual members. SharePoint group expansion works without it.
- Without `Sites.Read.All`: The **Sharing Links** and **Anonymous Access Summary** tools will show a permission error.

To enable these features, a SharePoint or Global Administrator must approve the permission requests after the package is deployed.

**Step 1 — Deploy the package**

Upload the `.sppkg` to the Tenant App Catalog and click **Deploy** when prompted. The *"Make this solution available to all sites"* checkbox controls where the web part can be placed but has no effect on permission approval.

**Step 2 — Check whether the permissions are already approved**

Go to **SharePoint Admin Center → Advanced → API access** (or open `https://<your-tenant>-admin.sharepoint.com/_layouts/15/online/ManageApiPermissions.aspx` directly).

> **What approved SPFx permissions look like:** Once approved, `GroupMember.Read.All` and `Sites.Read.All` appear in the list under **Microsoft Graph** with a blank app name (`—`) and no date — they do *not* show "Smart Permissions" as the app. This is normal. If you see both entries with "Yes" in the consent column, the permissions are active and no further action is needed.

If there are **pending requests** for Smart Permissions, approve them. If you see nothing pending and the permissions described above are already in the list, the web part is fully enabled.

**Option B — PnP PowerShell**

If the Admin Center shows no pending requests and you need to grant the permissions from scratch, use [PnP PowerShell](https://pnp.github.io/powershell/):

```powershell
Connect-PnPOnline -Url "https://<your-tenant>-admin.sharepoint.com" -Interactive

# List pending requests and note their IDs
Get-PnPTenantServicePrincipalPermissionRequests

# Approve each request by its ID
Approve-PnPTenantServicePrincipalPermissionRequest -RequestId "<guid-from-above>"
```

---

## Web Part Properties

The web part exposes one property pane setting:

| Property | Default | Description |
|----------|---------|-------------|
| **Default view on open** | Home | Sets which screen opens when the web part first loads. Options: Home, Permissions Report, Permissions Explorer, User Access. |

To change it: put the page in Edit mode → click the web part pencil → select from the **Default view on open** dropdown → republish the page.

---

## Project Structure

```
src/webparts/smartPermissions/
├── components/
│   ├── App.tsx                        # Root component, banner, navigation, theme wiring
│   ├── HomeView.tsx                   # Home screen — feature cards and More tools section
│   ├── PermissionsReportView.tsx      # Report configuration, progress, export
│   ├── PermissionsExplorerView.tsx    # Interactive folder/file tree + permission panel
│   ├── UserAccessView.tsx             # Per-user access scan
│   ├── SharingLinksView.tsx           # Sharing links browser (requires Sites.Read.All)
│   ├── PermissionGroupsView.tsx       # SharePoint group membership browser
│   ├── ExternalUsersView.tsx          # External (#EXT#) user report
│   ├── BrokenInheritanceView.tsx      # Broken inheritance finder
│   ├── AnonymousLinksView.tsx         # Anonymous and org-wide sharing links
│   └── SettingsView.tsx               # Full-page settings screen
├── services/
│   ├── SharePointService.ts           # All REST + Graph API calls
│   └── ExcelExportService.ts          # ExcelJS workbook generation
├── models/
│   └── models.ts                      # Shared TypeScript interfaces
└── SmartPermissionsWebPart.ts         # SPFx entry point, property pane, theme wiring
```

---

## Documentation

- [User Guide](docs/UserGuide.md) — end-user documentation covering all tools, settings, and web part configuration
- `docs/screenshots/` — auto-generated UI screenshots (regenerate with `node docs/screenshot.js`)
