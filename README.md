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
| **Permissions Explorer** | Interactively browse a document library tree and inspect live permissions on any folder or file. Shows inherited vs. unique permissions, supports expanding group members (SharePoint groups, Security groups, M365 groups). Automatically scans for external users in the background — red warning badges appear at load time on items where external access is explicitly granted; a down-arrow indicator marks folders that contain external-user items below. A tree filter narrows the view to only folders with external access, and external user email addresses are decoded and shown in the permissions panel. |
| **User Access** | Look up any user and see every location they can access, with their exact permission level at each location. Detects full-site owner access immediately. System accounts are automatically excluded from the user dropdown. |

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
- **Microsoft Graph API** — used for group member expansion (`GroupMember.Read.All`, optional)
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

The package is written to `sharepoint/solution/smart-permissions.sppkg`. Upload it to the SharePoint App Catalog. Optionally, approve the Graph permission request in the SharePoint Admin Center to enable group member expansion (see below).

---

## Graph API Permissions (optional)

The package declares one optional `webApiPermissionRequests` entry:

| Permission | Purpose | Required? |
|------------|---------|-----------|
| `GroupMember.Read.All` | Expand Security group and M365 group members in Permissions Report and Explorer | Optional |

This permission is **optional** — all tools work without it:
- Without `GroupMember.Read.All`: Security groups and M365 groups cannot be expanded to list individual members. SharePoint group expansion works without it.

To enable group member expansion, a SharePoint or Global Administrator must approve the permission request after the package is deployed.

**Step 1 — Deploy the package**

Upload the `.sppkg` to the Tenant App Catalog and click **Deploy** when prompted. The *"Make this solution available to all sites"* checkbox controls where the web part can be placed but has no effect on permission approval.

**Step 2 — Check whether the permission is already approved**

Go to **SharePoint Admin Center → Advanced → API access** (or open `https://<your-tenant>-admin.sharepoint.com/_layouts/15/online/ManageApiPermissions.aspx` directly).

> **What approved SPFx permissions look like:** Once approved, `GroupMember.Read.All` appears in the list under **Microsoft Graph** with a blank app name (`—`) and no date — it does *not* show "Smart Permissions" as the app. This is normal. If you see the entry with "Yes" in the consent column, the permission is active and no further action is needed.

If there is a **pending request** for Smart Permissions, approve it. If you see nothing pending and the permission described above is already in the list, the web part is fully enabled.

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
│   ├── HomeView.tsx                   # Home screen — three feature cards
│   ├── PermissionsReportView.tsx      # Report configuration, progress, export
│   ├── PermissionsExplorerView.tsx    # Interactive folder/file tree + permission panel
│   ├── UserAccessView.tsx             # Per-user access scan
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
