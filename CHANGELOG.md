# Changelog

All notable changes to this project are documented here.

---

## [1.1.0] — 2026-05-13

### Added

- **Permissions Report — "Access Via" column in Excel export**
  The Excel output now has 8 columns. The new "Access Via" column shows which
  SharePoint group, Security group, or M365 group a user was expanded from, or
  "Direct" if they were assigned permissions directly. This makes it easy to see
  both group-level and individual assignments in the same row.

- **Permissions Report — Expand Groups works for all scan depths including Site**
  Group expansion (SharePoint groups, Security groups, M365 groups) now runs
  regardless of the selected scan depth. Previously, scanning at the "Site only"
  depth returned raw group entries and never expanded their members.

- **Settings — dedicated full-page view**
  The settings gear icon now opens a full-page settings screen (with a Back button
  to return) instead of a floating popover. The settings view contains two sections:
  "Include system and hidden libraries" (checkbox) and "Default view on load"
  (instructions for the web part property pane).

- **Web Part Configuration section in documentation**
  New section in the user guide explains how site administrators can use the SharePoint
  property pane to set a default view (Home, Permissions Report, Permissions Explorer,
  or User Access) so the web part opens directly on that screen.

### Fixed

- **Permissions Report — stale results persisting after changing scan options**
  Changing the scan depth, folder depth, site scope, or expand-groups setting
  while results were already shown left the previous result set visible. Options
  changes now clear the result panel immediately so the display is never stale.

- **Web part property — default view not loading on first render**
  When the "Default view on open" property was set to anything other than Home via
  the SharePoint property pane, the web part crashed on load with "Cannot read
  properties of undefined (reading 'getLibraries')". Root cause: `applyTheme()`
  (inside the theme provider setup) called `this.render()` synchronously before
  `this._sp` had been assigned. Fixed by moving service initialisation ahead of
  theme provider setup in `onInit()`.

- **Settings — site URL field removed**
  The URL input that appeared in settings was redundant (the URL is already
  editable via the banner's "Change URL" button) and could confuse users.

### Changed

- **Permissions Report — default scan depth changed to "Site only"**
  The radio group now defaults to "Site only" instead of "Libraries". Site-only is
  the fastest scope and the most common starting point for an audit; users can
  switch to a deeper scan when needed.

- **Settings — "Default view on open" picker removed**
  The dropdown for selecting the default view has been removed from the settings
  panel. This is a deployment-time configuration that belongs in the SharePoint
  property pane, not a per-session user setting.

- **Documentation — all screenshots refreshed**
  All nine documentation screenshots regenerated to reflect the current UI, including
  the new full-page settings view and the updated report configuration screen.

- **Documentation — UserGuide.docx updated**
  Word document updated to match the current UserGuide.md: Settings section
  rewritten (full-page view description), new "Default View on Load" subsection,
  new "Web Part Configuration" section with property pane walkthrough, Step 4
  (Expand group members) added to the Permissions Report how-to with steps
  renumbered, and the Understanding the Results column list updated to reflect
  the 8-column Excel output (including the new Access Via column).

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
