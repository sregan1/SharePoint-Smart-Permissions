# SharePoint Smart Permissions

A browser-based SharePoint Online permissions auditing tool delivered as an SPFx web part. Gives site owners, administrators, and compliance teams a real-time view of who has access to what — with no PowerShell, no third-party software, and no elevated service accounts required.

---

## Features

| Tool | Description |
|------|-------------|
| **Permissions Report** | Scans a site's libraries, folders, and files and exports a colour-coded Excel workbook of every unique permission assignment. Supports full-tenant scans and optional group member expansion. |
| **Permissions Explorer** | Interactively browse a document library tree and inspect live permissions on any folder or file. Shows inherited vs. unique permissions and supports expanding group members (SharePoint groups, Security groups, M365 groups). |
| **User Access** | Look up any user and see every location they can access, with their exact permission level at each location. Detects full-site owner access immediately. |

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

## Technology

- **SPFx 1.21.1** · React 17 · TypeScript
- **Fluent UI v9** (`@fluentui/react-components`) — theme tokens driven by the SharePoint site theme
- **SharePoint REST API** — all permission data is read via standard SPO REST endpoints
- **Microsoft Graph API** — used for Security Group and M365 Group member expansion (`GroupMember.Read.All`)
- **ExcelJS** — in-browser Excel workbook generation

---

## Prerequisites

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

## Building & Packaging

```bash
# Production bundle
gulp bundle --ship

# Create the .sppkg deployment package
gulp package-solution --ship
```

The package is written to `sharepoint/solution/smart-permissions.sppkg`. Upload it to the SharePoint App Catalog and approve the `GroupMember.Read.All` Graph permission request in the SharePoint Admin Center.

---

## Graph API Permission

The package declares a `webApiPermissionRequests` entry for `Microsoft Graph / GroupMember.Read.All`. This permission is required for expanding Security Group and M365 Group members in the Permissions Explorer and Permissions Report. A SharePoint or Global Administrator must approve the request in **SharePoint Admin Center → Advanced → API access** after the package is deployed.

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
│   ├── App.tsx                     # Root component, banner, navigation, theme wiring
│   ├── HomeView.tsx                # Home screen with feature cards
│   ├── PermissionsReportView.tsx   # Report configuration, progress, export
│   ├── PermissionsExplorerView.tsx # Interactive folder/file tree + permission panel
│   ├── UserAccessView.tsx          # Per-user access scan
│   └── SettingsView.tsx            # Full-page settings screen
├── services/
│   ├── SharePointService.ts        # All REST + Graph API calls
│   └── ExcelExportService.ts       # ExcelJS workbook generation
├── models/
│   └── models.ts                   # Shared TypeScript interfaces
└── SmartPermissionsWebPart.ts      # SPFx entry point, property pane, theme wiring
```

---

## Documentation

- [User Guide](docs/UserGuide.md) — end-user documentation covering all three tools, settings, and web part configuration
- `docs/screenshots/` — auto-generated UI screenshots (regenerate with `node docs/screenshot.js`)

---

## Security & Privacy

- Runs entirely in the user's browser as the signed-in identity — no elevated permissions
- Read-only: never creates, modifies, or deletes any SharePoint content
- No external services: all data stays within the Microsoft 365 tenant
- No data storage: results exist only in the browser session
