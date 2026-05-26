# Changelog

All notable changes to this project are documented here.

---

## [1.1.0] — 2026-05-26

### Added

- **Five new tools in a collapsible "More tools" section on the home screen**
  The home screen retains its three main feature cards and adds a collapsible
  "More tools" panel with five additional reports:
  - **Permission Groups** — all SharePoint groups on the site with their members
  - **External Users** — all external (#EXT#) accounts with group memberships and
    a per-row "Check Access" shortcut to the User Access report
  - **Broken Inheritance Finder** — every library, folder, and file with unique permissions
  - **Sharing Links** — all sharing links by library (requires `Sites.Read.All`)
  - **Anonymous Access Summary** — anonymous and org-wide links by library (requires `Sites.Read.All`)

- **`Sites.Read.All` Graph permission request**
  Added to `webApiPermissionRequests` in `package-solution.json` to support the
  Sharing Links and Anonymous Access Summary tools. Approval is optional — all
  other tools work without it.

- **Permissions Report — "Access Via" column in Excel export**
  The Excel output now has 8 columns. The new "Access Via" column shows which
  SharePoint group, Security group, or M365 group a user was expanded from, or
  "Direct" if they were assigned permissions directly.

- **Permissions Report — Expand Groups works for all scan depths including Site**
  Group expansion (SharePoint groups, Security groups, M365 groups) now runs
  regardless of the selected scan depth. Previously, scanning at the "Site only"
  depth returned raw group entries and never expanded their members.

- **Settings — dedicated full-page view**
  The settings gear icon now opens a full-page settings screen (with a Back button
  to return) instead of a floating popover.

- **Web Part Configuration section in documentation**
  New section in the user guide explains how site administrators can use the SharePoint
  property pane to set a default view so the web part opens directly on that screen.

### Fixed

- **`GroupMember.Read.All` not approved — graceful error instead of silent failure**
  When the optional `GroupMember.Read.All` Graph permission has not been approved,
  the Permissions Report now shows a warning banner after the scan completes and the
  Permissions Explorer shows an error message when expanding a group — rather than
  silently returning empty member lists.

- **`getExternalUsers` — concurrent group expansion and cancellation**
  Group membership lookups were fired all at once with `Promise.all`, causing
  throttling (HTTP 429) on sites with more than a handful of external users.
  Replaced with `runConcurrent` respecting `scanConcurrency`. The `AbortSignal`
  is now also checked between each lookup so cancellation is immediate.

- **`scanBrokenInheritance` — N+1 API pattern eliminated**
  Previously called `getLibraries` (no `HasUniqueRoleAssignments` field) followed
  by a separate `GetList` call per library to check inheritance status. Replaced
  with a single `lists?$select=...,HasUniqueRoleAssignments` request.

- **`scanBrokenInheritance` — parallel library and subfolder scanning**
  Libraries and subfolder recursion are now scanned concurrently via `runConcurrent`,
  matching the pattern used by the Permissions Report and User Access tools.

- **Permissions Report — "External users only" export now filters permission rows**
  When the "External users only (#ext#)" filter is active and the user exports,
  the Excel and CSV output contain only the external-user rows within each matching
  entry. Previously all principals for those entries were included.

- **Permissions Report — stale results persisting after changing scan options**
  Changing scan depth, folder depth, site scope, or expand-groups while results
  were shown left the previous result set visible. Options changes now clear the
  result panel immediately.

- **Web part property — default view not loading on first render**
  When "Default view on open" was set via the property pane, the web part crashed
  on load. Root cause: `applyTheme()` called `this.render()` before `this._sp` was
  assigned. Fixed by moving service initialization ahead of theme provider setup.

- **Settings — site URL field removed**
  The URL input in settings was redundant (the URL is editable via the banner's
  "Change URL" button).

### Changed

- **Permissions Report — default scan depth changed to "Site only"**
  The radio group now defaults to "Site only" instead of "Libraries".

- **Settings — "Default view on open" picker removed**
  Moved to the SharePoint property pane where it belongs as a deployment-time setting.

- **Documentation updated throughout**
  README and User Guide updated to cover all tools, the revised home screen layout,
  both Graph API permissions, and the expanded project structure. All British
  spellings corrected to American English.

---

## [1.0.0] — 2026-05-12

### Added

- **Permissions Report — progress bar, timer, and live item count**
  Added a deterministic progress bar (advances library-by-library), an elapsed timer
  (M:SS format), and a live "N items scanned · Library X of Y" counter that updates
  in real time as folders and files are processed.

- **Permissions Explorer — inherited-permissions banner**
  Items that inherit permissions now display a visually distinct callout banner
  (blue accent border, chain-link icon, light blue background) instead of plain grey
  text, making the inherited state immediately obvious at a glance.

- **Permissions Explorer — options moved to top of panel**
  "Expand SharePoint group members" and "Show parent permissions" (when applicable)
  are now shown at the top of the permission panel, above the content area, so they
  are always visible without scrolling.

- **Home screen — full-width brand banner**
  The home screen now opens with the same full-width blue banner as all other views,
  providing a consistent header across the entire web part.

- **Documentation — automated screenshots**
  `docs/screenshot.js`: generates nine PNG screenshots using puppeteer-core and a
  local Chrome installation. Output goes to `docs/screenshots/` (git-ignored).

- **Documentation — Word user guide**
  `docs/generate-word.js`: generates `docs/UserGuide.docx` from the markdown source
  with all screenshots embedded. Output is git-ignored; regenerate locally as needed.

- **Documentation — end-user guide**
  `docs/UserGuide.md`: full end-user documentation covering all three features,
  global settings, and troubleshooting tips.

### Fixed

- **System/hidden library filtering — root cause resolved (all three views)**
  The previous client-side `NoCrawl` / `IsSiteAssetsLibrary` checks were silently
  bypassed because SharePoint REST ignores a `$filter` clause when the same request
  also contains a deep `$expand=RoleAssignments/Member`. Fixed by decoupling the two
  concerns: the library list is now fetched with a simple, reliable query; role
  assignments are fetched per-library in a separate targeted call only when
  `HasUniqueRoleAssignments` is true. Applies to Permissions Report, Permissions
  Explorer, and User Access.

- **Form Templates and Style Library explicitly excluded**
  Added URL-suffix checks (`/formservertemplates`, `/style library`) to the shared
  `isSystemLibrary` helper as an additional safeguard, independent of the
  `NoCrawl` / `IsSiteAssetsLibrary` field values returned by SharePoint.

- **Permissions Explorer — "Show parent permissions" deselected on group expansion**
  Toggling "Expand SharePoint group members" previously called `handleSelectNode`,
  which unconditionally reset `showParentPerms` to `false`. The effect now refreshes
  permissions in place: node permissions are re-fetched with the updated expansion
  setting, and if parent permissions were already showing they are re-fetched too
  (also with group members expanded).

- **Home screen — settings gear icon invisible against blue banner**
  The floating settings gear rendered in `App.tsx` for the home view was styled
  `appearance="subtle"` with no explicit colour. Updated to
  `appearance="transparent"` with a white icon to match its appearance on the
  non-home banner.

### Changed

- **Banner background colour — consistent across all screens**
  All four screens (Home, Permissions Report, Permissions Explorer, User Access) use
  `tokens.colorBrandBackground`, which resolves to the current SharePoint site's
  theme colour. The banner will automatically match the site branding everywhere.

- **Permissions Explorer — "Inherited" badge**
  Changed from `appearance="outline"` (barely visible) to
  `appearance="filled" color="informative"` (solid blue chip) so Unique and Inherited
  states are symmetrically distinct in the item title row.

- **Home screen icon alignment**
  The shield icon is now vertically centred with the title text
  (`alignItems: center` instead of `flex-start`).

- **Public repository preparation**
  - Version bumped to `1.0.0` in `package.json`
  - Removed placeholder `mpnId` from `config/package-solution.json`
  - Added generated documentation artifacts to `.gitignore`
    (`docs/UserGuide.docx`, `docs/screenshots/`)
  - README updated: project structure, generating-documentation section,
    MIT licence reference, Chrome path instructions for screenshot script

---

## [0.0.1] — Initial release

First internal build. Core features: Permissions Report, Permissions Explorer,
User Access.
