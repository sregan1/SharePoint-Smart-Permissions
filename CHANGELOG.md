# Changelog

All notable changes to this project are documented here.

---

## [Unreleased]

### Fixed

- **Provisioning script (`scripts/Provision-SmartPermissions.ps1`)**
  - App install now targets the tenant App Catalog by default instead of a
    site-collection app catalog most sites don't have, and waits (`-Wait`) for
    the (asynchronous) install to actually finish before adding the web part.
  - Publishing now uses `Set-PnPPage -Publish` (the current PnP.PowerShell
    cmdlet) instead of the removed `Publish-PnPPage`, and publishes the same
    in-memory page object that was just edited instead of re-fetching a stale
    copy that overwrote the change.
  - The web part is now resolved via `Get-PnPAvailablePageComponents` before
    being added — passing a raw component ID string previously produced an
    empty, non-rendering placeholder control.
  - Fixed a duplicate-webpart guard that silently never matched due to a
    braced-vs-unbraced GUID string comparison.
  - The pages library is now resolved once by its stable "SitePages" URL
    segment instead of the localized "Site Pages" display title, which broke
    every permission/audience-targeting step on non-English sites.
  - Added optional `$clientId` support for tenants that reject the default PnP
    Management Shell app during interactive sign-in.
  - `$adminUrl` is now only required in tenant-wide mode, not when provisioning
    a single site via `$targetSiteUrl`.
  - Quick Launch navigation node de-duplication now matches on URL as well as
    title.
  - Comments are now disabled on the provisioned page.

- **Audit accuracy (Permissions Report / User Access)**
  - Items whose unique-permission check fails transiently (network blip,
    throttling past retries) are now flagged as scan-incomplete instead of
    being silently presented as inheriting their parent's permissions.
  - `getJson` now retries thrown/rejected requests (not just 429/503
    responses) with backoff, and threads an abort signal through paged
    requests.
  - A Graph profile-lookup permission error in the M365-group ownership check
    no longer gets misreported as a confirmed "Member" role — it's now
    surfaced as "Graph permission required" instead of a false result.
  - `transitiveMemberOf` group-membership checks now follow `@odata.nextLink`
    instead of silently truncating at 999 groups.
  - SharePoint group member expansion now resolves by numeric group ID
    instead of by display name, which 404'd (and silently dropped members)
    for renamed or non-canonically-named groups.
  - `checkCanManagePermissions` now fails closed on an API error instead of
    granting Owner-level UI to an unverified caller.
  - "Everyone" and "Everyone except external users" claims are now detected
    and flagged prominently as tenant-wide access instead of rendering as an
    ordinary, easy-to-miss group row.
  - The "exclude limited access" and "external users only" filters are now
    applied consistently across the Explorer, Report, and User Access views
    and their Excel/CSV exports (shared `applyPermFilters` helper) — the
    Report's export and expanded-row view previously ignored the
    limited-access filter, so the workbook and on-screen filter could
    disagree.
  - The Report export button's enabled state and row count now reflect what
    will actually be exported after filtering, instead of a pre-filter count
    that could show "0 filtered rows" as exportable.
  - A results-pagination reset effect now also fires when "exclude limited
    access" is toggled.
  - Group-member-count fetches no longer request more than Graph's 999-item
    `$top` maximum, even when the configured display cap is set higher.

- **Explorer view reliability/performance**
  - Folder prefetches are now bounded by the configured scan concurrency
    instead of firing one request per folder simultaneously.
  - Background scans and prefetches are now discarded if the user switches
    libraries (or navigates away) while they're still in flight, instead of
    applying stale results to whatever is now on screen.
  - A failed folder listing now shows inline, non-navigable error text
    instead of a synthetic tree node with an empty URL.

### Changed

- User-identifying diagnostic logging in the User Access scan is now gated
  behind `localStorage.setItem('smartPermissionsDebug', '1')` instead of
  always writing to the browser console.
- `.yo-rc.json` now records the actual SPFx generator version (1.21.1).
- `config/package-solution.json` now includes publisher name and website.
- Removed the superseded `docs/screenshot.js` in favor of the current
  `docs/generate-screenshots.js`.
- Documentation (README, USER-GUIDE) updated to match current script/build
  behavior: idempotent re-running, the app auto-install step, the
  `serve.json.template` copy step, and accurate screenshot tracking status.

---

## [1.4.0] — 2026-06-13

### Added

- **Lists and Site Pages included in scans**
  The Permissions Report and User Access scans now cover all visible lists —
  generic lists, Site Pages, calendars, task lists, etc. — not just document
  libraries. Generic lists are reported at list level with a new **List** type
  badge; document-library-like templates (document libraries, picture
  libraries, Site Pages) are still walked for folders and files.

- **Subsite scanning**
  A new "Include subsites" checkbox on the Permissions Report recursively scans
  every subsite below the selected site (and below each site collection in
  all-sites mode). Report history shows "+ subsites" in the Scope column.

- **Results table in the Permissions Report**
  Scan results can now be browsed directly in the browser: a paginated,
  sortable table with type badges, unique/inherited status, and expandable rows
  showing each object's permission assignments. Previously results were only
  visible after exporting.

- **Report compare (diff)**
  Select two saved reports in the History panel and click **Compare selected**
  to see what changed between scans: permission assignments added, removed, or
  role-changed per object; inheritance broken or restored; and objects that
  appeared or disappeared. A warning is shown when the two reports used
  different sites or scan options.

- **Tenant-wide people search in User Access**
  The user picker now searches the whole tenant (via the standard SharePoint
  people-picker API — no extra Graph permissions) once you type 3+ characters.
  Users who have access through an AAD group but have never visited the site
  appear under a "Not in this site" group.

- **NoCrawl libraries surfaced instead of hidden**
  Libraries marked NoCrawl (hidden from search) were previously treated as
  system libraries and excluded by default. They are now included and flagged
  with a "Hidden from search" badge in results, a suffix in the Explorer
  library dropdown, and a marker in Excel/CSV exports — content deliberately
  hidden from search is exactly what an audit should surface.

- **Explorer keyboard accessibility**
  The folder tree is now fully keyboard-navigable (arrow keys, Home/End,
  Enter/Space) with proper `tree`/`treeitem` ARIA roles, expanded/selected
  states, and a roving tabindex for screen readers.

### Changed

- **Permissions Report scan parallelized**
  Lists are now scanned concurrently (bounded by the Settings → Concurrent API
  requests value) instead of one at a time, while preserving the
  parent-before-child ordering of results within each library.

- **Cancelled scans keep partial results**
  Cancelling a Permissions Report scan now shows everything collected so far —
  with filtering and export available — instead of discarding it.

- **Smaller initial bundle**
  The Excel export library (exceljs, ~950 KB) now loads on demand as a
  separate chunk the first time an .xlsx export is run, instead of being part
  of the main bundle downloaded on every page load.

- **Notification permission requested on first scan**
  Browser notification permission is now requested when a scan starts rather
  than when a view opens.

- **SharePointService split into focused modules**
  The 1,800-line service is now a thin facade over `services/sp/` modules
  (core API client, site discovery, report scan, explorer, groups, user
  access). The public API is unchanged.

### Fixed

- **"Scan all site collections" returned no sites**
  The tenant-wide site discovery parsed a response shape the SharePoint search
  API never returns under SPFx's default OData mode (and the search API
  rejects the default OData v4 header outright). Discovery now overrides to
  OData v3, handles both response shapes, and surfaces errors instead of
  silently scanning nothing.

- **Items named with `&`, `#`, or `%` broke scans**
  Folder, file, list, and group API calls now use the
  `*ByServerRelativePath(decodedUrl=...)` endpoints with URI encoding. Items
  in paths containing these characters were previously skipped silently.

- **Results beyond one page were silently dropped**
  All collection fetches (folders, files, lists, site users) now follow
  server-side paging links instead of stopping at the first `$top` page.

- **Owner detection checked the wrong permission bits**
  `checkCanManagePermissions` tested the ManagePermissions/ManageWeb masks
  against the High word of the permission bitmask; they live in the Low word.
  Full Control users were unaffected, but custom roles granting Manage
  Permissions were incorrectly locked out of the home-screen tools.

- **Throttling risk from unbounded concurrency**
  Recursive folder walks previously started a new concurrency pool per
  recursion level, multiplying in-flight requests (up to N^depth). A single
  shared work queue now caps total concurrency at the configured value.

- **Notification crash could mask successful scans**
  On platforms where the `Notification` constructor throws (e.g. Android
  Chrome), a completed scan displayed an error and skipped saving to history.
  Notifications are now best-effort.

- **User Access launched from Explorer showed the login instead of the name**
  The display name is now passed through when cross-navigating, so the status
  line and saved history show the user's name.

### Removed

- **Permission Groups view**
  The unreachable Permission Groups view (hidden from the home screen since
  v1.3.0) and its supporting service code have been removed, along with other
  dead code (`getSharingLinks`, `getExternalUsers`, `scanBrokenInheritance`
  and their models).

- **Debug logging**
  Removed the module-evaluation console marker and the global
  `unhandledrejection` listener left over from debugging.

---

## [1.3.0] — 2026-06-09

### Added

- **Access detection on the home screen**
  The web part now checks the signed-in user's effective permissions at startup
  (and whenever the site URL changes). Anyone without the Manage Permissions right —
  Members, Visitors, Limited Access users, and guests — sees a warning banner
  explaining that Site Owner access is required and prompting them to contact a
  site owner.

- **Feature cards disabled for non-owners**
  When the signed-in user lacks Site Owner access, the three home screen feature
  cards are rendered at very low opacity with a grayscale filter, a `not-allowed`
  cursor, and a "Requires Site Owner access" tooltip. Cards are non-clickable,
  preventing navigation into tools that would show empty or error states.

- **`checkCanManagePermissions()` in SharePointService**
  New method that reads `/_api/web?$select=EffectiveBasePermissions` and checks the
  ManagePermissions and ManageWeb bitmask flags to determine whether the signed-in
  user has the right to read role assignments.

- **Provisioning script — single-site mode**
  `scripts/Provision-SmartPermissions.ps1` now supports a `$targetSiteUrl` variable.
  Set it to provision a single site collection without enumerating the entire tenant.
  Leave it empty for the existing tenant-wide behavior.

- **Provisioning script — Owners-only page permissions**
  After creating and publishing the Permissions page, the script now breaks
  permission inheritance on the page list item and grants Full Control exclusively
  to the site's Associated Owners group, preventing Members and Visitors from
  accessing the page.

- **Provisioning script — Quick Launch navigation**
  The Permissions page is automatically added to the site's Quick Launch navigation.
  On Microsoft 365 Group-connected sites the nav node is scoped to the Owners group
  via `-AudienceIds` so only Owners see the link.

- **Provisioning script — page audience targeting**
  On M365-connected sites, audience targeting is enabled on the Site Pages library
  and the Owners group is set as the page audience via `_ModernAudienceTargetUserField`.

### Changed

- **Home screen — three-column card layout**
  The feature cards are now displayed in a single row of three columns (responsive:
  two columns below 800 px, one column below 500 px). Previously two columns.

- **Home screen — card order**
  Cards are now ordered: Permissions Explorer, Permissions Report, User Access.

- **Home screen — card image height normalized**
  All card screenshots are cropped to a fixed 180 px height (`object-fit: cover`,
  `object-position: top`) so cards align evenly regardless of the source image
  aspect ratio.

- **Home screen — Permission Groups card hidden**
  The Permission Groups feature card has been removed from the home screen. The
  underlying view and backend remain in the codebase.

- **Site owners list removed from Member warning**
  The warning banner no longer lists site owner names. System Account is also
  filtered from the owners query in SharePointService.

### Fixed

- **System Account excluded from site owners list**
  `getSiteOwners()` now filters out any user with the title "System Account" or
  a `SHAREPOINT\system` login name.

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

- **Documentation — end-user guide**
  `USER-GUIDE.md`: full end-user documentation covering all three features,
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
    (`docs/screenshots/`)
  - README updated: project structure, generating-documentation section,
    MIT licence reference, Chrome path instructions for screenshot script

---

## [0.0.1] — Initial release

First internal build. Core features: Permissions Report, Permissions Explorer,
User Access.
