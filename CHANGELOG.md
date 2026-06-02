# Changelog

All notable changes to this project are documented here.

---

## [1.2.0] — 2026-06-02

### Added

- **User Access — site permission shown for M365 Group members**
  Users who access a site through a Microsoft 365 Group (rather than through a classic
  SharePoint group) now correctly see a site-level permission entry. The scan uses a
  three-tier detection approach: SP group membership via SharePoint REST, M365 Group
  membership via the site's `GroupId` and `AssociatedOwnerGroup`/`AssociatedMemberGroup`
  properties, and Graph `transitiveMemberOf` as a final fallback. This resolves a long-
  standing gap where M365-connected sites showed only folder and file permissions with no
  site-level entry for members.

- **User Access — highest permission shown when in multiple groups**
  When a user is a member of both the Owners and Members M365 Group, the site permission
  now reflects the most permissive role (Full Control) rather than the lower one (Edit).
  Owner status is checked via `GET /groups/{id}/owners/{userId}` before falling back to
  member-level roles.

- **Permissions Explorer — permission-denied banner for member accounts**
  Reading role assignments in SharePoint Online requires the Manage Permissions right,
  which site owners have but regular members typically do not. The Explorer previously
  showed nothing when role-assignment reads failed with HTTP 403 — the view looked
  identical to an empty site. A blue informational banner now appears when this condition
  is detected, explaining that external user indicators and permission tables may be
  incomplete for the current account.

- **User Access — "New scan" replaces "Scan again"**
  The "Scan again" button previously re-ran the scan for the same user. It now clears the
  results and returns to the user picker so a different user can be selected.

### Fixed

- **User Access — "Web-Only Limited Access" no longer shown as site permission**
  SharePoint auto-assigns "Web-Only Limited Access" (and "Limited Access") to users when
  they are granted item-level access without explicit site-level access. Both are now
  treated as system-assigned pass-through roles and filtered from all results, alongside
  any role definition whose name begins with `System.`.

- **User Access — site permission missing for M365-backed SharePoint groups**
  Sites using classic SharePoint groups ("Team Members", "Team Owners") whose members are
  managed via an M365 Group now correctly detect membership. The detection uses
  `sitegroups/getbyid(groupId)/users?$filter=LoginName eq '...'` rather than
  `getbyloginname`, which returned HTTP 404 because `Member.LoginName` in role assignments
  is the group display name, not its internal login name.

- **Scan in-progress text simplified**
  The lengthy "This scan may take several minutes depending on the size of the site."
  message during a User Access scan has been replaced with "Scanning…".

### Changed

- **Permissions Explorer — three distinct folder indicator icons**
  Folders now use three shapes to indicate what lies below: a **circle** arrow-down for
  unique permissions, a **triangle** arrow-down for external user access, and both icons
  side-by-side when a folder contains both. Previously a single icon covered all three
  cases.

---

## [1.1.1] — 2026-05-29

### Added

- **Permissions Explorer — automatic external user detection**
  The Explorer now runs a background scan when a library loads (and again whenever a
  folder is expanded) that checks every unique-permission node for external accounts.
  The red warning-person icon appears automatically — no clicking required — within
  seconds of the library loading. The scan is optimized: inherited nodes are skipped
  entirely (they require no API call), and each unique-permission node requires only
  one direct `RoleAssignments` fetch instead of two calls, reducing API traffic by
  roughly 90% on typical libraries.

- **Permissions Explorer — "External users only" filter now filters the tree**
  The toggle previously filtered only the permissions table for the selected item.
  It now filters the entire folder tree, showing only folders and files where external
  users have access (direct or inherited) and the ancestor path leading to them —
  mirroring how the "Unique permissions only" filter works.

- **Permissions Explorer — external user email shown in permissions panel**
  The permissions table now shows the decoded email address of external users (e.g.
  `john@contoso.com`) beneath their display name. The email is decoded from the
  SharePoint `#EXT#` login-name format and is only shown when it differs from the
  display name already shown.

- **Permissions Explorer — down-arrow indicator extended to external users below**
  The grey down-arrow (↓) icon on a folder previously indicated only that unique
  permissions existed somewhere below. It now also appears when external user access
  exists somewhere below. The tooltip is context-aware: it reads "Contains items with
  unique permissions", "Contains items with external user access", or "Contains items
  with unique permissions and external user access" depending on what is present.
  The legend entry is updated to match.

### Fixed

- **Permissions Explorer — level-2 folder icons not appearing without expanding**
  The library load pre-fetch fetched second-level folder contents for structural
  inspection but never set `parent` references on the pre-fetched nodes and never
  called the external-user scan on them. Icons on second-level unique-permission
  folders now appear at library-load time, not only after the user expands their
  parent folder. Both `parent` assignment and `scanExternalUsers` are now called
  inside the pre-fetch callback.

- **Permissions Explorer — warning icon incorrectly shown on inherited folders**
  The red warning-person icon was appearing on folders that merely inherited external
  access from an ancestor, not on folders where access was explicitly granted. The
  icon is now restricted to items where `hasUniquePermissions` is true — the exact
  location where the external access was assigned — so it serves as a direct action
  point. Folders that pass through inherited external access show only the down-arrow
  indicator instead.

- **Permissions Explorer — external user access propagated to inherited descendants**
  When a unique-permission item is found to have external users during the scan, all
  already-loaded descendant folders that inherit (no unique permissions of their own)
  are now marked immediately. This ensures the "External users only" tree filter and
  the down-arrow indicator are correct even before the user expands those folders.

### Removed

- **Five secondary tools removed from the home screen**
  The collapsible "More tools" panel and its five reports — Permission Groups,
  External Users, Broken Inheritance Finder, Sharing Links, and Anonymous Access
  Summary — have been removed. The home screen returns to the original three-tool
  layout: Permissions Report, Permissions Explorer, and User Access.

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
