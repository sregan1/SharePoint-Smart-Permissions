import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, MSGraphClientV3 } from '@microsoft/sp-http';
import {
  UserPermissionInfo, PermissionEntry, FolderFileNode, LibraryInfo,
  SiteCollectionInfo, SiteUserInfo, ReportOptions, ReportScope, ObjectType, ScanProgress,
  SharingLinkEntry, PermissionGroup, ExternalUserEntry, BrokenInheritanceEntry,
} from '../models/models';

// Escape single-quotes in OData string literals (SQL-style doubling).
function odata(s: string): string {
  return s.replace(/'/g, "''");
}

// Detect Graph API permission errors (HTTP 401/403 or well-known message patterns).
// Exported so views can use it without duplicating the detection logic.
export function isGraphPermissionError(err: any): boolean {
  const msg = String(err?.message ?? err ?? '').toLowerCase();
  return (
    err?.statusCode === 401 ||
    err?.statusCode === 403 ||
    msg.includes('forbidden') ||
    msg.includes('unauthorized') ||
    msg.includes('accessdenied') ||
    msg.includes('does not represent a site')
  );
}

// Map numeric SPO PrincipalType to label string.
function principalTypeLabel(type: number): string {
  if (type === 4) return 'SecurityGroup';
  if (type === 8) return 'SharePointGroup';
  return 'User';
}

// Normalise role-definition-binding arrays: SPO REST returns a direct array
// with odata=nometadata; legacy verbose mode wraps it in { results: [] }.
function rdbArray(bindings: any): any[] {
  if (Array.isArray(bindings)) return bindings;
  if (Array.isArray(bindings?.value)) return bindings.value;
  if (Array.isArray(bindings?.results)) return bindings.results;
  return [];
}

// Normalise top-level value arrays (same odata format issue).
function valueArray(data: any): any[] {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.value)) return data.value;
  if (Array.isArray(data?.results)) return data.results;
  return [];
}

// Extract the Azure AD object GUID from a SharePoint claims login name.
// M365 Groups:        c:0o.c|federateddirectoryclaimprovider|{GUID}
// M365 Group owners:  c:0o.c|federateddirectoryclaimprovider|{GUID}_o
// Security Groups:    c:0t.c|tenant|{GUID}  /  c:0p.c|s2s|{GUID}
// The _o suffix indicates the owners slice of an M365 group — requires /groups/{id}/owners.
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function extractAadGroupId(loginName: string): string | null {
  const last = loginName.split('|').pop() ?? '';
  return GUID_RE.test(last) ? last : null;
}
function extractAadGroupInfo(loginName: string): { id: string; isOwners: boolean } | null {
  let last = loginName.split('|').pop() ?? '';
  const isOwners = last.endsWith('_o');
  if (isOwners) last = last.slice(0, -2);
  return GUID_RE.test(last) ? { id: last, isOwners } : null;
}

// Returns true when an API error indicates the current user lacks read permission on
// role assignments (HTTP 403 Forbidden or 401 Unauthorized).
function isPermissionDenied(err: any): boolean {
  const msg = String(err?.message ?? '');
  return msg.includes('HTTP 403') || msg.includes('HTTP 401');
}

// Translate an EffectiveBasePermissions {High, Low} bitmask to the highest matching
// SharePoint permission level name. SP REST returns High/Low as strings — callers
// must parseInt before passing in.
function spBitmaskToLevel(high: number, low: number): string {
  void high; // High bits not needed for the levels we distinguish
  const lo = low >>> 0; // treat as unsigned 32-bit
  if (lo & 0x02000000 || lo & 0x40000000) return 'Full Control'; // ManagePermissions | ManageWeb
  if (lo & 0x00000800) return 'Design';                           // ManageLists
  if (lo & 0x00000004) return 'Edit';                             // EditListItems
  if (lo & 0x00000001) return 'Read';                             // ViewListItems
  return 'Limited access';
}

// Returns true for system-assigned pass-through roles that carry no meaningful permission
// of their own and should be hidden from all results ("Limited Access" and its web-scoped
// variant are both auto-assigned by SharePoint and are never explicitly granted).
function isSystemRole(name: string): boolean {
  const l = name.toLowerCase();
  return l === 'limited access' || l === 'web-only limited access' || l.startsWith('system.');
}

// Known system/infrastructure library URL suffixes (lowercased, site-relative).
// Checked as a suffix so they match regardless of site path prefix.
const SYSTEM_LIB_SUFFIXES = [
  '/formservertemplates', // Form Templates
  '/style library',       // Style Library
];

// Returns true if this list entry should be treated as a system/hidden library
// and excluded when includeHidden is false.
function isSystemLibrary(lib: any): boolean {
  if (lib.NoCrawl || lib.IsSiteAssetsLibrary) return true;
  const url = ((lib.RootFolder?.ServerRelativeUrl) ?? '').toLowerCase();
  return SYSTEM_LIB_SUFFIXES.some((s) => url.endsWith(s));
}

export class SharePointService {
  private readonly context: WebPartContext;
  /** Max concurrent API requests during scans. Settable from Settings. */
  public scanConcurrency = 4;
  /** Max group members fetched before capping. Settable from Settings. */
  public groupMemberCap = 500;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Retries on 429/503 using the Retry-After header, with a 3-attempt cap.
  private async getJson(url: string, attempt = 0): Promise<any> {
    const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if ((resp.status === 429 || resp.status === 503) && attempt < 3) {
      const retryAfter = parseInt(resp.headers.get('Retry-After') ?? '10', 10);
      await new Promise((r) => setTimeout(r, (isNaN(retryAfter) ? 10 : retryAfter) * 1000));
      return this.getJson(url, attempt + 1);
    }
    if (!resp.ok) {
      const txt = await resp.text();
      throw new Error(`HTTP ${resp.status} — ${txt.substring(0, 300)}`);
    }
    return resp.json();
  }

  // ── Tenant / Site discovery ───────────────────────────────────────────────

  async getAllSites(tenantUrl: string, signal?: AbortSignal): Promise<SiteCollectionInfo[]> {
    const sites: SiteCollectionInfo[] = [];
    const base = tenantUrl.replace(/\/$/, '');
    let startRow = 0;
    const rowLimit = 500;

    while (!signal?.aborted) {
      const url =
        `${base}/_api/search/query?querytext='contentclass:STS_Site'` +
        `&rowlimit=${rowLimit}&startrow=${startRow}&selectproperties='Title,Path'`;

      try {
        const data = await this.getJson(url);
        const rows: any[] =
          data?.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results ?? [];
        if (rows.length === 0) break;

        for (const row of rows) {
          let title: string | null = null;
          let path: string | null = null;
          for (const cell of row.Cells?.results ?? []) {
            if (cell.Key === 'Title') title = cell.Value;
            if (cell.Key === 'Path') path = cell.Value;
          }
          if (path) sites.push({ url: path, title: title ?? path });
        }

        const total: number =
          data?.d?.query?.PrimaryQueryResult?.RelevantResults?.TotalRows ?? 0;
        startRow += rows.length;
        if (startRow >= total || rows.length < rowLimit) break;
      } catch {
        break;
      }
    }

    return sites;
  }

  async getLibraries(siteUrl: string, signal?: AbortSignal, includeHidden = false): Promise<LibraryInfo[]> {
    // Server-side: BaseTemplate=101 (doc libraries) + Hidden=false (reliably supported).
    // NoCrawl and IsSiteAssetsLibrary are fetched and filtered client-side because
    // NoCrawl is not reliably filterable via OData across all SPO tenants.
    const filter = includeHidden
      ? 'BaseTemplate eq 101'
      : 'BaseTemplate eq 101 and Hidden eq false';
    const url =
      `${siteUrl}/_api/web/lists` +
      `?$filter=${encodeURIComponent(filter)}` +
      `&$select=Title,RootFolder/ServerRelativeUrl,NoCrawl,IsSiteAssetsLibrary` +
      `&$expand=RootFolder&$orderby=Title&$top=500`;
    const data = await this.getJson(url);
    return valueArray(data)
      .filter((l: any) => includeHidden || !isSystemLibrary(l))
      .map((l: any) => ({
        title: l.Title,
        serverRelativeUrl: l.RootFolder?.ServerRelativeUrl ?? '',
      }));
  }

  async getSiteUsers(siteUrl: string, signal?: AbortSignal): Promise<SiteUserInfo[]> {
    const url =
      `${siteUrl}/_api/web/siteusers` +
      `?$filter=IsHiddenInUI eq false and PrincipalType eq 1` +
      `&$select=LoginName,Title,Email&$orderby=Title&$top=2000`;
    const data = await this.getJson(url);
    return valueArray(data)
      .filter(
        (u: any) =>
          !u.LoginName?.includes('_spo_') &&
          !u.LoginName?.includes('app@sharepoint'),
      )
      .map((u: any) => ({ loginName: u.LoginName, displayName: u.Title, email: u.Email || undefined }));
  }

  // ── Permissions Report scan ───────────────────────────────────────────────

  async scanPermissions(
    options: ReportOptions,
    onProgress: (progress: ScanProgress) => void,
    signal?: AbortSignal,
    onEntry?: (entry: PermissionEntry) => void,
  ): Promise<{ entries: PermissionEntry[]; groupPermissionDenied: boolean; roleAssignmentsDenied: boolean }> {
    const entries: PermissionEntry[] = [];
    const flags = { groupPermissionDenied: false, roleAssignmentsDenied: false };
    const emit0 = (message: string): void =>
      onProgress({ message, scanned: entries.length, libsDone: 0, libsTotal: 0 });

    if (options.allSites) {
      emit0('Discovering site collections…');
      const sites = await this.getAllSites(options.siteUrl, signal);
      for (const site of sites) {
        if (signal?.aborted) break;
        emit0(`Scanning: ${site.title}`);
        await this.scanSite(site.url, options, entries, onProgress, signal, onEntry, flags);
      }
    } else {
      await this.scanSite(options.siteUrl, options, entries, onProgress, signal, onEntry, flags);
    }

    return { entries, groupPermissionDenied: flags.groupPermissionDenied, roleAssignmentsDenied: flags.roleAssignmentsDenied };
  }

  private async scanSite(
    siteUrl: string,
    options: ReportOptions,
    entries: PermissionEntry[],
    onProgress: (progress: ScanProgress) => void,
    signal?: AbortSignal,
    onEntry?: (entry: PermissionEntry) => void,
    flags?: { groupPermissionDenied: boolean; roleAssignmentsDenied: boolean },
  ): Promise<void> {
    const startIndex = entries.length;
    let libsDone = 0;
    let libsTotal = 0;
    const emit = (message: string): void =>
      onProgress({ message, scanned: entries.length, libsDone, libsTotal });

    // ── Site-level permissions ────────────────────────────────────────────
    emit(`Loading site permissions: ${siteUrl}`);
    let sitePerms: UserPermissionInfo[] = [];

    try {
      const webData = await this.getJson(
        `${siteUrl}/_api/web` +
          `?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
      );
      sitePerms = this.toPermissionInfoList(valueArray(webData.RoleAssignments));
      const siteEntry: PermissionEntry = {
        objectType: ObjectType.Site,
        name: webData.Title ?? siteUrl,
        serverRelativeUrl: webData.ServerRelativeUrl ?? '',
        siteUrl,
        hasUniquePermissions: true,
        depth: 0,
        uniquePermissions: sitePerms,
      };
      entries.push(siteEntry);
      onEntry?.(siteEntry);
    } catch (err: any) {
      if (isPermissionDenied(err) && flags) flags.roleAssignmentsDenied = true;
      // Fallback: no role assignments, just record the site.
      try {
        const webData = await this.getJson(
          `${siteUrl}/_api/web?$select=Title,Url,ServerRelativeUrl`,
        );
        const siteEntry: PermissionEntry = {
          objectType: ObjectType.Site,
          name: webData.Title ?? siteUrl,
          serverRelativeUrl: webData.ServerRelativeUrl ?? '',
          siteUrl,
          hasUniquePermissions: true,
          depth: 0,
          uniquePermissions: [],
        };
        entries.push(siteEntry);
        onEntry?.(siteEntry);
      } catch { /* skip */ }
    }

    if (options.scope !== ReportScope.Site) {
      // ── Libraries ───────────────────────────────────────────────────────
      // Fetch the library list with a simple query so $filter is applied reliably.
      // Combining $filter with a deep $expand=RoleAssignments can cause SPO to silently
      // ignore the filter clause, letting hidden/system libraries slip through.
      emit('Loading document libraries…');
      let libs: any[] = [];

      const hiddenFilter = encodeURIComponent(
        options.includeHidden
          ? 'BaseTemplate eq 101'
          : 'BaseTemplate eq 101 and Hidden eq false',
      );

      try {
        const listsData = await this.getJson(
          `${siteUrl}/_api/web/lists?$filter=${hiddenFilter}` +
            `&$select=Title,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=200`,
        );
        libs = valueArray(listsData);
      } catch { /* no libraries available */ }

      if (!options.includeHidden) {
        libs = libs.filter((l: any) => !isSystemLibrary(l));
      }
      if (options.libraryUrls) {
        const allowed = new Set(options.libraryUrls);
        libs = libs.filter((l: any) => allowed.has(l.RootFolder?.ServerRelativeUrl ?? ''));
      }

      libsTotal = libs.length;
      emit('Starting library scan…');

      for (const lib of libs) {
        if (signal?.aborted) break;
        emit(`Scanning library: ${lib.Title}`);

        let libPerms = sitePerms;
        if (lib.HasUniqueRoleAssignments) {
          try {
            const raData = await this.getJson(
              `${siteUrl}/_api/web/GetList('${odata(lib.RootFolder.ServerRelativeUrl)}')/RoleAssignments` +
                `?$expand=Member,RoleDefinitionBindings` +
                `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
            );
            libPerms = this.toPermissionInfoList(valueArray(raData));
          } catch (err: any) {
            if (isPermissionDenied(err) && flags) flags.roleAssignmentsDenied = true;
          }
        }

        const libEntry: PermissionEntry = {
          objectType: ObjectType.Library,
          name: lib.Title,
          serverRelativeUrl: lib.RootFolder?.ServerRelativeUrl ?? '',
          siteUrl,
          hasUniquePermissions: !!lib.HasUniqueRoleAssignments,
          depth: 1,
          uniquePermissions: libPerms,
        };
        entries.push(libEntry);
        onEntry?.(libEntry);

        if (
          options.scope === ReportScope.Folder ||
          options.scope === ReportScope.Item
        ) {
          try {
            await this.walkFolder(
              siteUrl,
              lib.RootFolder.ServerRelativeUrl,
              2,
              1,
              options,
              entries,
              libPerms,
              emit,
              signal,
              onEntry,
            );
          } catch { /* partial results OK */ }
        }

        libsDone++;
        emit(`Scanning library: ${lib.Title}`);
      }
    }

    if (options.expandGroups && !signal?.aborted) {
      // memberCache: group login/name → expanded UserPermissionInfo[] (avoids re-fetching same group)
      const memberCache = new Map<string, UserPermissionInfo[]>();
      // expandedCache: original perms array reference → expanded array reference.
      // Inherited entries share the SAME parentPerms object reference, so this eliminates
      // duplicate expanded arrays for the thousands of entries that inherit from the same source.
      const expandedCache = new Map<UserPermissionInfo[], UserPermissionInfo[]>();

      for (const entry of entries.slice(startIndex)) {
        if (signal?.aborted) break;

        if (expandedCache.has(entry.uniquePermissions)) {
          entry.uniquePermissions = expandedCache.get(entry.uniquePermissions)!;
          continue;
        }

        const originalRef = entry.uniquePermissions;
        const expanded: UserPermissionInfo[] = [];
        for (const up of originalRef) {
          expanded.push(up);
          if (up.principalType === 'SharePointGroup' || up.principalType === 'SecurityGroup') {
            const cacheKey = up.loginName || up.displayName;
            let members: UserPermissionInfo[];
            if (memberCache.has(cacheKey)) {
              members = memberCache.get(cacheKey)!;
            } else {
              try {
                members = await this.getGroupMembers(siteUrl, up.displayName, up.loginName, up.principalType, signal);
              } catch (err: any) {
                if (err?.isGraphPermissionError && flags) flags.groupPermissionDenied = true;
                members = [];
              }
              memberCache.set(cacheKey, members);
            }
            members.forEach((m) => expanded.push({ ...m, roles: [...up.roles], sourceGroup: up.displayName }));
          }
        }
        entry.uniquePermissions = expanded;
        expandedCache.set(originalRef, expanded);
      }
    }
  }

  private async walkFolder(
    siteUrl: string,
    folderUrl: string,
    depth: number,
    currentLevel: number,
    options: ReportOptions,
    results: PermissionEntry[],
    parentPerms: UserPermissionInfo[],
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
    onEntry?: (entry: PermissionEntry) => void,
  ): Promise<void> {
    if (signal?.aborted) return;

    const enc = odata(folderUrl);
    let subFolders: any[] = [];
    let files: any[] = [];
    let uniquePermsLoaded = false;

    try {
      const [fData, fiData] = await Promise.all([
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
        ),
        options.scope === ReportScope.Item
          ? this.getJson(
              `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Files` +
                `?$select=Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments` +
                `&$expand=ListItemAllFields&$top=2000`,
            )
          : Promise.resolve({ value: [] }),
      ]);
      subFolders = valueArray(fData);
      files = valueArray(fiData);
      uniquePermsLoaded = true;
    } catch {
      try {
        const [fData, fiData] = await Promise.all([
          this.getJson(
            `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
              `?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
          ),
          options.scope === ReportScope.Item
            ? this.getJson(
                `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Files` +
                  `?$select=Name,ServerRelativeUrl&$top=2000`,
              )
            : Promise.resolve({ value: [] }),
        ]);
        subFolders = valueArray(fData);
        files = valueArray(fiData);
      } catch { return; }
    }

    const visibleFolders = subFolders.filter(
      (f: any) =>
        !f.Name.startsWith('_') &&
        f.Name.toLowerCase() !== 'forms',
    );

    for (const subfolder of visibleFolders) {
      if (signal?.aborted) break;

      let folderPerms = parentPerms;
      let hasUnique = false;

      if (uniquePermsLoaded && subfolder.ListItemAllFields?.HasUniqueRoleAssignments) {
        try {
          const raData = await this.getJson(
            `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${odata(subfolder.ServerRelativeUrl)}')/ListItemAllFields/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
          );
          folderPerms = this.toPermissionInfoList(valueArray(raData));
          hasUnique = true;
        } catch { /* inherit parent */ }
      }

      const folderEntry: PermissionEntry = {
        objectType: ObjectType.Folder,
        name: subfolder.Name,
        serverRelativeUrl: subfolder.ServerRelativeUrl,
        siteUrl,
        hasUniquePermissions: hasUnique,
        depth,
        uniquePermissions: folderPerms,
      };
      results.push(folderEntry);
      onEntry?.(folderEntry);
      onProgress(subfolder.Name);

      const shouldRecurse =
        options.scope === ReportScope.Item || currentLevel < options.folderDepth;
      if (shouldRecurse) {
        try {
          await this.walkFolder(
            siteUrl,
            subfolder.ServerRelativeUrl,
            depth + 1,
            currentLevel + 1,
            options,
            results,
            folderPerms,
            onProgress,
            signal,
            onEntry,
          );
        } catch { /* continue */ }
      }
    }

    if (options.scope !== ReportScope.Item) return;

    for (const file of files) {
      if (signal?.aborted) break;

      let filePerms = parentPerms;
      let hasUnique = false;

      if (uniquePermsLoaded && file.ListItemAllFields?.HasUniqueRoleAssignments) {
        try {
          const raData = await this.getJson(
            `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${odata(file.ServerRelativeUrl)}')/ListItemAllFields/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
          );
          filePerms = this.toPermissionInfoList(valueArray(raData));
          hasUnique = true;
        } catch { /* inherit parent */ }
      }

      const fileEntry: PermissionEntry = {
        objectType: ObjectType.File,
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        siteUrl,
        hasUniquePermissions: hasUnique,
        depth,
        uniquePermissions: filePerms,
      };
      results.push(fileEntry);
      onEntry?.(fileEntry);
      onProgress(file.Name);
    }
  }

  // ── Real-time audit ───────────────────────────────────────────────────────

  async getFolderContents(
    siteUrl: string,
    folderUrl: string,
    signal?: AbortSignal,
  ): Promise<FolderFileNode[]> {
    const enc = odata(folderUrl);

    const toNodes = (folders: any[], filesArr: any[], withUnique: boolean): FolderFileNode[] => {
      const nodes: FolderFileNode[] = [];
      for (const sub of folders.filter((f: any) => !f.Name.startsWith('_'))) {
        nodes.push({
          name: sub.Name,
          serverRelativeUrl: sub.ServerRelativeUrl,
          isFolder: true,
          hasChildren: (sub.ItemCount ?? 0) > 0,
          hasUniquePermissions: withUnique
            ? !!sub.ListItemAllFields?.HasUniqueRoleAssignments
            : undefined,
          children: [],
        });
      }
      for (const file of filesArr) {
        nodes.push({
          name: file.Name,
          serverRelativeUrl: file.ServerRelativeUrl,
          isFolder: false,
          hasChildren: false,
          hasUniquePermissions: withUnique
            ? !!file.ListItemAllFields?.HasUniqueRoleAssignments
            : undefined,
          children: [],
        });
      }
      return nodes.sort((a, b) => {
        if (a.isFolder !== b.isFolder) return a.isFolder ? -1 : 1;
        return a.name.localeCompare(b.name);
      });
    };

    try {
      const [fData, fiData] = await Promise.all([
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
        ),
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Files` +
            `?$select=Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
        ),
      ]);
      return toNodes(valueArray(fData), valueArray(fiData), true);
    } catch {
      // Fallback: no HasUniqueRoleAssignments, fresh requests.
      const [fData, fiData] = await Promise.all([
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
        ),
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Files` +
            `?$select=Name,ServerRelativeUrl&$top=2000`,
        ),
      ]);
      return toNodes(valueArray(fData), valueArray(fiData), false);
    }
  }

  async getItemPermissions(
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<{ hasUnique: boolean; users: UserPermissionInfo[]; permissionDenied?: boolean }> {
    const enc = odata(node.serverRelativeUrl);
    const base = node.isFolder
      ? `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/ListItemAllFields`
      : `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${enc}')/ListItemAllFields`;

    try {
      const itemData = await this.getJson(`${base}?$select=HasUniqueRoleAssignments`);
      if (!itemData.HasUniqueRoleAssignments) return { hasUnique: false, users: [] };

      const raData = await this.getJson(
        `${base}/RoleAssignments` +
          `?$expand=Member,RoleDefinitionBindings` +
          `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
      );
      return { hasUnique: true, users: this.toPermissionInfoList(valueArray(raData)) };
    } catch (err) {
      return { hasUnique: false, users: [], permissionDenied: isPermissionDenied(err) };
    }
  }

  // Checks a single tree node for external users without re-fetching HasUniqueRoleAssignments
  // (already known from the tree). Returns false immediately for inherited nodes.
  // Returns 'denied' when the current user lacks ManagePermissions on this item.
  async scanNodeForExternalUsers(
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<boolean | 'denied'> {
    if (!node.hasUniquePermissions || signal?.aborted) return false;
    const enc = odata(node.serverRelativeUrl);
    const base = node.isFolder
      ? `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/ListItemAllFields`
      : `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${enc}')/ListItemAllFields`;
    try {
      const raData = await this.getJson(
        `${base}/RoleAssignments?$expand=Member&$select=Member/LoginName`,
      );
      return valueArray(raData).some((ra: any) =>
        (ra.Member?.LoginName ?? '').toLowerCase().includes('#ext#'),
      );
    } catch (err) {
      return isPermissionDenied(err) ? 'denied' : false;
    }
  }

  async getParentPermissions(
    siteUrl: string,
    serverRelativeUrl: string,
    signal?: AbortSignal,
  ): Promise<{ name: string; serverRelativeUrl: string; users: UserPermissionInfo[] } | null> {
    const siteServerRelUrl = new URL(siteUrl).pathname.replace(/\/$/, '');
    let currentUrl = serverRelativeUrl.replace(/\/$/, '');

    while (!signal?.aborted) {
      const lastSlash = currentUrl.lastIndexOf('/');
      if (lastSlash <= 0) return null;
      currentUrl = currentUrl.substring(0, lastSlash);

      if (currentUrl === siteServerRelUrl) {
        try {
          const webData = await this.getJson(
            `${siteUrl}/_api/web` +
              `?$select=Title,ServerRelativeUrl` +
              `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
          );
          return {
            name: webData.Title ?? siteUrl,
            serverRelativeUrl: webData.ServerRelativeUrl ?? currentUrl,
            users: this.toPermissionInfoList(valueArray(webData.RoleAssignments)),
          };
        } catch {
          return null;
        }
      }

      try {
        const enc = odata(currentUrl);
        const base =
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/ListItemAllFields`;
        const itemData = await this.getJson(`${base}?$select=HasUniqueRoleAssignments`);

        if (itemData.HasUniqueRoleAssignments) {
          const raData = await this.getJson(
            `${base}/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
          );
          return {
            name: currentUrl.substring(currentUrl.lastIndexOf('/') + 1),
            serverRelativeUrl: currentUrl,
            users: this.toPermissionInfoList(valueArray(raData)),
          };
        }
      } catch {
        // continue walking up
      }
    }

    return null;
  }

  async getEffectivePermissions(
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<string> {
    const enc = odata(node.serverRelativeUrl);
    const base = node.isFolder
      ? `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/ListItemAllFields`
      : `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${enc}')/ListItemAllFields`;
    try {
      const data = await this.getJson(`${base}?$select=EffectiveBasePermissions`);
      if (signal?.aborted) return '';
      return spBitmaskToLevel(
        parseInt(data.EffectiveBasePermissions?.High ?? '0', 10),
        parseInt(data.EffectiveBasePermissions?.Low ?? '0', 10),
      );
    } catch {
      return '';
    }
  }

  async getSiteGroups(
    siteUrl: string,
    signal?: AbortSignal,
  ): Promise<{ id: number; title: string; loginName: string; description: string }[]> {
    if (signal?.aborted) return [];
    const data = await this.getJson(
      `${siteUrl}/_api/web/sitegroups?$select=Id,Title,LoginName,Description&$orderby=Title`,
    );
    return valueArray(data).map((g: any) => ({
      id: g.Id as number,
      title: g.Title ?? '',
      loginName: g.LoginName ?? '',
      description: g.Description ?? '',
    }));
  }

  async getSiteOwners(
    siteUrl: string,
    signal?: AbortSignal,
  ): Promise<{ title: string; email: string }[]> {
    try {
      if (signal?.aborted) return [];
      const data = await this.getJson(
        `${siteUrl}/_api/web/AssociatedOwnerGroup/users?$select=Title,Email,IsHiddenInUI&$top=10`,
      );
      return valueArray(data)
        .filter((u: any) => !u.IsHiddenInUI &&
          u.Title !== 'System Account' &&
          (u.LoginName ?? '').toLowerCase().indexOf('sharepoint\\system') === -1)
        .map((u: any) => ({ title: u.Title ?? '', email: u.Email ?? '' }));
    } catch {
      return [];
    }
  }

  async checkCanManagePermissions(siteUrl: string): Promise<boolean> {
    try {
      const data = await this.getJson(`${siteUrl}/_api/web?$select=EffectiveBasePermissions`);
      const high: number = data?.EffectiveBasePermissions?.High ?? 0;
      return !!(high & 0x02000000 || high & 0x40000000);
    } catch {
      return true; // fail open — don't block owners on API error
    }
  }

  async getGroupMembers(
    siteUrl: string,
    groupName: string,
    loginName: string,
    principalType: string,
    signal?: AbortSignal,
  ): Promise<UserPermissionInfo[]> {
    if (principalType === 'SharePointGroup') {
      try {
        const data = await this.getJson(
          `${siteUrl}/_api/web/sitegroups/getbyname('${odata(groupName)}')/users` +
            // Fetch one extra to detect truncation without a separate count call
            `?$select=LoginName,Title,IsHiddenInUI,PrincipalType&$top=${this.groupMemberCap + 1}`,
        );
        const rawMembers: any[] = valueArray(data).filter((u: any) => !u.IsHiddenInUI);
        const expanded: UserPermissionInfo[] = [];
        for (const u of rawMembers) {
          const pt: number = u.PrincipalType ?? 1;
          if (pt !== 1) {
            // Member is a nested group (M365 group or security group) — try to expand via Graph.
            // M365-connected SP groups commonly contain the backing M365 group as their only member.
            const nestedInfo = extractAadGroupInfo(u.LoginName ?? '');
            if (nestedInfo) {
              const nestedMembers = await this._getAadGroupMembers(nestedInfo.id, nestedInfo.isOwners ? 'owners' : 'members');
              if (nestedMembers !== null) {
                nestedMembers.forEach((m) => expanded.push({ ...m, isGroupMember: true }));
                continue;
              }
            }
            // Can't expand — show the nested group as a named entry rather than hiding it
            expanded.push({
              loginName: u.LoginName ?? '',
              displayName: u.Title ?? '',
              principalType: principalTypeLabel(pt),
              roles: [],
              isGroupMember: true,
            });
          } else {
            expanded.push({
              loginName: u.LoginName ?? '',
              displayName: u.Title ?? '',
              principalType: 'User',
              roles: [],
              isGroupMember: true,
            });
          }
        }
        expanded.sort((a, b) => a.displayName.localeCompare(b.displayName));
        if (expanded.length > this.groupMemberCap) {
          const capped = expanded.slice(0, this.groupMemberCap);
          capped.push({
            loginName: '',
            displayName: `(group has more than ${this.groupMemberCap} members — only first ${this.groupMemberCap} shown)`,
            principalType: 'User',
            roles: [],
            isGroupMember: true,
          });
          return capped;
        }
        return expanded;
      } catch {
        return [];
      }
    }

    if (principalType === 'SecurityGroup') {
      const groupInfo = extractAadGroupInfo(loginName);
      if (!groupInfo) return [];
      const result = await this._getAadGroupMembers(groupInfo.id, groupInfo.isOwners ? 'owners' : 'members');
      if (result === null) {
        const e: any = new Error(
          'Group member expansion requires the GroupMember.Read.All Graph permission. ' +
          'A tenant admin must approve it in SharePoint Admin Center → Advanced → API access.',
        );
        e.isGraphPermissionError = true;
        throw e;
      }
      return result;
    }

    return [];
  }

  private async _getAadGroupMembers(groupId: string, endpoint: 'members' | 'owners' = 'members'): Promise<UserPermissionInfo[] | null> {
    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const result = await client
        .api(`/groups/${groupId}/${endpoint}`)
        .select('displayName,userPrincipalName,mail,id')
        .top(this.groupMemberCap + 1)
        .get();
      const all = (result?.value ?? [])
        .map((m: any): UserPermissionInfo => ({
          loginName: m.userPrincipalName ?? m.mail ?? m.id ?? '',
          displayName: m.displayName ?? m.userPrincipalName ?? m.id ?? '',
          principalType: 'User',
          roles: [],
          isGroupMember: true,
        }))
        .sort((a: UserPermissionInfo, b: UserPermissionInfo) =>
          a.displayName.localeCompare(b.displayName),
        );
      if (all.length > this.groupMemberCap) {
        const capped = all.slice(0, this.groupMemberCap);
        capped.push({
          loginName: '',
          displayName: `(group has more than ${this.groupMemberCap} members — only first ${this.groupMemberCap} shown)`,
          principalType: 'User',
          roles: [],
          isGroupMember: true,
        });
        return capped;
      }
      return all;
    } catch (err: any) {
      if (isGraphPermissionError(err)) return null;
      return [];
    }
  }

  public async runConcurrent<T>(
    tasks: (() => Promise<T | undefined>)[],
    concurrency = 5,
  ): Promise<(T | undefined)[]> {
    if (tasks.length === 0) return [];
    const results: (T | undefined)[] = new Array(tasks.length);
    let idx = 0;
    const worker = async () => {
      while (idx < tasks.length) {
        const i = idx++;
        try { results[i] = await tasks[i](); }
        catch { results[i] = undefined; }
      }
    };
    await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
    return results;
  }


  async getUserAccess(
    siteUrl: string,
    userLoginName: string,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
    includeHidden = false,
  ): Promise<{ fullSiteAccess: boolean; items: PermissionEntry[]; graphPermissionRequired: boolean; roleAssignmentsDenied: boolean }> {
    onProgress('Loading user info…');

    let userTitle = userLoginName;
    let userEmail = '';
    let groupLogins = new Set<string>();

    try {
      const userData = await this.getJson(
        `${siteUrl}/_api/web/siteusers/getbyloginname('${encodeURIComponent(odata(userLoginName))}')` +
          `?$expand=Groups&$select=Title,LoginName,Email,Groups/LoginName`,
      );
      userTitle = userData.Title ?? userLoginName;
      userEmail = userData.Email ?? '';
      const groups = valueArray(userData.Groups);
      // Normalize to lowercase — role-assignment member logins can differ in case
      // from user-groups logins depending on the SharePoint REST endpoint used.
      groupLogins = new Set(groups.map((g: any) => (g.LoginName as string).toLowerCase()));
      console.debug('[SmartPermissions] getUserAccess: user=%s email=%s SP-groups=%o',
        userLoginName, userEmail, Array.from(groupLogins));
    } catch (err) {
      console.debug('[SmartPermissions] getUserAccess: siteusers fetch failed', err);
    }

    // ── Site roles ──
    onProgress('Loading site permissions…');
    let siteRoles: string[] = [];
    let webData: any = null;
    let roleAssignmentsDenied = false;

    try {
      webData = await this.getJson(
        `${siteUrl}/_api/web` +
          `?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
      );
      const allRa = valueArray(webData.RoleAssignments);
      console.debug('[SmartPermissions] getUserAccess: site RoleAssignments=%o',
        allRa.map((ra: any) => ({
          login: ra.Member?.LoginName,
          principalType: ra.Member?.PrincipalType,
          roles: rdbArray(ra.RoleDefinitionBindings).map((r: any) => r.Name),
        })));
      siteRoles = this.extractRoles(allRa, userLoginName, groupLogins)
        .filter((r) => !isSystemRole(r));
      console.debug('[SmartPermissions] getUserAccess: siteRoles after extractRoles=%o', siteRoles);
    } catch (err) {
      console.debug('[SmartPermissions] getUserAccess: web RoleAssignments fetch failed', err);
      if (isPermissionDenied(err)) roleAssignmentsDenied = true;
    }

    // Member-access fallback: RoleAssignments are not readable, but the user's SP group
    // memberships (fetched above without elevated rights) can be checked against the site's
    // three default associated groups, which are always readable. This gives a site-level
    // result without needing Manage Permissions.
    if (roleAssignmentsDenied) {
      onProgress('Checking site group membership…');
      try {
        const [web, ownerGrp, memberGrp, visitorGrp] = await Promise.all([
          this.getJson(`${siteUrl}/_api/web?$select=Title,Url,ServerRelativeUrl`).catch(() => null),
          this.getJson(`${siteUrl}/_api/web/AssociatedOwnerGroup?$select=LoginName,Title`).catch(() => null),
          this.getJson(`${siteUrl}/_api/web/AssociatedMemberGroup?$select=LoginName,Title`).catch(() => null),
          this.getJson(`${siteUrl}/_api/web/AssociatedVisitorGroup?$select=LoginName,Title`).catch(() => null),
        ]);

        const matchedRoles: string[] = [];
        if (ownerGrp?.LoginName && groupLogins.has((ownerGrp.LoginName as string).toLowerCase())) {
          matchedRoles.push('Full Control');
        }
        if (memberGrp?.LoginName && groupLogins.has((memberGrp.LoginName as string).toLowerCase())) {
          matchedRoles.push('Edit');
        }
        if (visitorGrp?.LoginName && groupLogins.has((visitorGrp.LoginName as string).toLowerCase())) {
          matchedRoles.push('Read');
        }

        if (matchedRoles.length > 0) {
          const siteEntry: PermissionEntry = {
            objectType: ObjectType.Site,
            name: web?.Title ?? siteUrl,
            serverRelativeUrl: web?.ServerRelativeUrl ?? '',
            siteUrl,
            hasUniquePermissions: true,
            depth: 0,
            uniquePermissions: [{ loginName: userLoginName, displayName: userTitle, principalType: 'User', roles: matchedRoles }],
          };
          return { fullSiteAccess: false, items: [siteEntry], graphPermissionRequired: false, roleAssignmentsDenied: true };
        }
      } catch { /* ignore fallback errors */ }

      return { fullSiteAccess: false, items: [], graphPermissionRequired: false, roleAssignmentsDenied: true };
    }

    // Fallback for M365 Group-connected sites: extractRoles only matches SP group logins.
    // M365 Groups and AAD Security Groups appear directly in RoleAssignments but their
    // members are not returned by the siteusers Groups expansion. Use Graph to check
    // transitive membership in any such groups found in the site's role assignments.
    let graphPermissionRequired = false;
    if (siteRoles.length === 0 && webData) {
      console.debug('[SmartPermissions] getUserAccess: extractRoles returned empty — trying AAD group fallback');
      const aadResult = await this.getAadGroupSiteRoles(
        siteUrl,
        userLoginName,
        userEmail,
        valueArray(webData.RoleAssignments),
      );
      siteRoles = aadResult.roles;
      graphPermissionRequired = aadResult.graphUnavailable;
      console.debug('[SmartPermissions] getUserAccess: siteRoles after AAD fallback=%o graphPermissionRequired=%s',
        siteRoles, graphPermissionRequired);
    }

    // ── Libraries ──
    onProgress('Loading libraries…');
    let libs: any[] = [];
    let libsHaveRoles = false;
    const libFilter = encodeURIComponent(
      includeHidden
        ? 'BaseTemplate eq 101'
        : 'BaseTemplate eq 101 and Hidden eq false',
    );

    try {
      const listsData = await this.getJson(
        `${siteUrl}/_api/web/lists` +
          `?$filter=${libFilter}` +
          `&$select=Title,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
          `,RoleAssignments/Member/LoginName,RoleAssignments/Member/PrincipalType` +
          `,RoleAssignments/RoleDefinitionBindings/Name` +
          `&$expand=RootFolder,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` +
          `&$top=200`,
      );
      libs = valueArray(listsData);
      libsHaveRoles = true;
    } catch {
      try {
        const listsData = await this.getJson(
          `${siteUrl}/_api/web/lists` +
            `?$filter=${libFilter}` +
            `&$select=Title,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=200`,
        );
        libs = valueArray(listsData);
      } catch { /* no libraries */ }
    }

    if (!includeHidden) {
      libs = libs.filter((l: any) => !isSystemLibrary(l));
    }

    let siteEntry: PermissionEntry | undefined;
    if (siteRoles.length > 0 && webData) {
      siteEntry = {
        objectType: ObjectType.Site,
        name: webData.Title ?? siteUrl,
        serverRelativeUrl: webData.ServerRelativeUrl ?? '',
        siteUrl,
        hasUniquePermissions: true,
        depth: 0,
        uniquePermissions: [
          {
            loginName: userLoginName,
            displayName: userTitle,
            principalType: 'User',
            roles: siteRoles,
          },
        ],
      };
    }

    const isOwner = siteRoles.some(
      (r) =>
        r.toLowerCase().includes('full control') ||
        r.toLowerCase().includes('owner'),
    );

    // Full Site Access banner — owners have Full Control everywhere when no
    // library breaks inheritance. If some do, scan to verify actual access.
    if (isOwner && libs.every((l) => !l.HasUniqueRoleAssignments)) {
      return { fullSiteAccess: true, items: siteEntry ? [siteEntry] : [], graphPermissionRequired: false, roleAssignmentsDenied: false };
    }

    // Member/Visitor with site-level access and no broken-inheritance libraries —
    // all content is accessible via site inheritance; no scan needed.
    const hasSiteAccess = siteRoles.length > 0;
    if (hasSiteAccess && !isOwner && libs.every((l) => !l.HasUniqueRoleAssignments)) {
      return { fullSiteAccess: false, items: siteEntry ? [siteEntry] : [], graphPermissionRequired: false, roleAssignmentsDenied: false };
    }

    const items: PermissionEntry[] = siteEntry ? [siteEntry] : [];

    await this.runConcurrent(libs.map((lib: any) => async () => {
      if (signal?.aborted) return;
      onProgress(`Scanning library: ${lib.Title}`);

      if (lib.HasUniqueRoleAssignments && libsHaveRoles) {
        const libRoles = this.extractRoles(
          valueArray(lib.RoleAssignments),
          userLoginName,
          groupLogins,
        ).filter((r) => !isSystemRole(r));

        if (libRoles.length > 0) {
          items.push({
            objectType: ObjectType.Library,
            name: lib.Title,
            serverRelativeUrl: lib.RootFolder?.ServerRelativeUrl ?? '',
            siteUrl,
            hasUniquePermissions: true,
            depth: 1,
            uniquePermissions: [
              {
                loginName: userLoginName,
                displayName: userTitle,
                principalType: 'User',
                roles: libRoles,
              },
            ],
          });
        }
      }

      await this.walkFoldersForUser(
        siteUrl,
        lib.RootFolder?.ServerRelativeUrl ?? '',
        userLoginName,
        userTitle,
        groupLogins,
        items,
        2,
        onProgress,
        signal,
      );
    }), this.scanConcurrency);

    return { fullSiteAccess: false, items, graphPermissionRequired, roleAssignmentsDenied };
  }

  private async walkFoldersForUser(
    siteUrl: string,
    folderUrl: string,
    userLogin: string,
    userDisplayName: string,
    groupLogins: Set<string>,
    results: PermissionEntry[],
    depth: number,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
  ): Promise<void> {
    if (signal?.aborted) return;

    const enc = odata(folderUrl);
    let subFolders: any[] = [];
    let files: any[] = [];
    let uniquePermsLoaded = false;

    try {
      const [fData, fiData] = await Promise.all([
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
        ),
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Files` +
            `?$select=Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
        ),
      ]);
      subFolders = valueArray(fData);
      files = valueArray(fiData);
      uniquePermsLoaded = true;
    } catch {
      try {
        const fData = await this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
        );
        subFolders = valueArray(fData);
      } catch { return; }
    }

    const visibleFolders = subFolders.filter(
      (f: any) => !f.Name.startsWith('_') && f.Name.toLowerCase() !== 'forms',
    );

    if (uniquePermsLoaded) {
      const uniqueFolders = visibleFolders.filter(
        (f: any) => f.ListItemAllFields?.HasUniqueRoleAssignments,
      );
      const uniqueFiles = files.filter(
        (f: any) => f.ListItemAllFields?.HasUniqueRoleAssignments,
      );

      await this.runConcurrent(uniqueFolders.map((subfolder: any) => async () => {
        if (signal?.aborted) return;
        const raData = await this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${odata(subfolder.ServerRelativeUrl)}')/ListItemAllFields/RoleAssignments` +
            `?$expand=Member,RoleDefinitionBindings` +
            `&$select=Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name`,
        );
        const roles = this.extractRoles(valueArray(raData), userLogin, groupLogins).filter(
          (r) => !isSystemRole(r),
        );
        if (roles.length > 0) {
          results.push({
            objectType: ObjectType.Folder,
            name: subfolder.Name,
            serverRelativeUrl: subfolder.ServerRelativeUrl,
            siteUrl,
            hasUniquePermissions: true,
            depth,
            uniquePermissions: [
              { loginName: userLogin, displayName: userDisplayName, principalType: 'User', roles },
            ],
          });
        }
      }));

      await this.runConcurrent(uniqueFiles.map((file: any) => async () => {
        if (signal?.aborted) return;
        const raData = await this.getJson(
          `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${odata(file.ServerRelativeUrl)}')/ListItemAllFields/RoleAssignments` +
            `?$expand=Member,RoleDefinitionBindings` +
            `&$select=Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name`,
        );
        const roles = this.extractRoles(valueArray(raData), userLogin, groupLogins).filter(
          (r) => !isSystemRole(r),
        );
        if (roles.length > 0) {
          results.push({
            objectType: ObjectType.File,
            name: file.Name,
            serverRelativeUrl: file.ServerRelativeUrl,
            siteUrl,
            hasUniquePermissions: true,
            depth,
            uniquePermissions: [
              { loginName: userLogin, displayName: userDisplayName, principalType: 'User', roles },
            ],
          });
        }
      }));
    }

    // Recurse into all visible subfolders concurrently.
    await this.runConcurrent(
      visibleFolders
        .filter((subfolder: any) => !signal?.aborted && (subfolder.ItemCount ?? 0) > 0)
        .map((subfolder: any) => () =>
          this.walkFoldersForUser(
            siteUrl,
            subfolder.ServerRelativeUrl,
            userLogin,
            userDisplayName,
            groupLogins,
            results,
            depth + 1,
            onProgress,
            signal,
          ),
        ),
      this.scanConcurrency,
    );
  }

  // For M365 Group-connected sites the SP groups (Team Members, Team Owners, etc.) only
  // store static members. Actual M365 Group members are resolved dynamically by SharePoint
  // and never appear in sitegroups/users. This method:
  //   1. Gets the site's connected M365 Group GUID from _api/web?$select=GroupId
  //   2. Checks whether the user is in that M365 Group via Graph transitiveMemberOf
  //   3. Checks whether they are an Owner (→ AssociatedOwnerGroup roles) or a Member
  //      (→ AssociatedMemberGroup roles) and returns the corresponding permission roles.
  private async checkM365ConnectedSiteRoles(
    siteUrl: string,
    userLoginName: string,
    userEmail: string,
    actionableGroupRas: any[],
  ): Promise<{ roles: string[]; graphUnavailable: boolean }> {
    try {
      // GroupId lives on the site collection (_api/site), not the web.
      // AssociatedOwnerGroup/MemberGroup live on the web (_api/web).
      const [siteProps, webProps] = await Promise.all([
        this.getJson(`${siteUrl}/_api/site?$select=GroupId`),
        this.getJson(
          `${siteUrl}/_api/web?$expand=AssociatedOwnerGroup,AssociatedMemberGroup` +
          `&$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id`,
        ),
      ]);
      const m365GroupId: string = siteProps?.GroupId ?? '';
      console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: site GroupId=%s ownerGroupId=%s memberGroupId=%s',
        m365GroupId || '(none)', webProps?.AssociatedOwnerGroup?.Id, webProps?.AssociatedMemberGroup?.Id);

      if (!m365GroupId || m365GroupId === '00000000-0000-0000-0000-000000000000') {
        console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: site not M365-connected');
        return { roles: [], graphUnavailable: false };
      }

      const identifier = userEmail ||
        (userLoginName.includes('|') ? userLoginName.split('|').pop() : null) ||
        userLoginName;
      if (!identifier) return { roles: [], graphUnavailable: true };

      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

      // Fetch transitive group memberships and the user's AAD object ID in parallel.
      // The AAD ID is needed for the ownership check below.
      const [memberOfData, userProfileData] = await Promise.all([
        client.api(`/users/${encodeURIComponent(identifier)}/transitiveMemberOf`).select('id').top(999).get(),
        client.api(`/users/${encodeURIComponent(identifier)}`).select('id').get().catch(() => null),
      ]);
      const userGroupIds = new Set<string>(
        (memberOfData?.value ?? []).map((g: any) => (g.id as string).toLowerCase()),
      );
      const userAadId: string = userProfileData?.id ?? '';
      console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: user in %d group(s), site M365 Group present=%s, aadId=%s',
        userGroupIds.size, userGroupIds.has(m365GroupId.toLowerCase()), userAadId || '(unknown)');

      if (!userGroupIds.has(m365GroupId.toLowerCase())) {
        return { roles: [], graphUnavailable: false };
      }

      // User is in the M365 Group. Check if they're an Owner (→ AssociatedOwnerGroup →
      // Full Control) or just a Member (→ AssociatedMemberGroup → Edit).
      // Strategy: try GET /groups/{groupId}/owners/{userId} first (direct lookup, no filter
      // needed). Fall back to a filtered list with ConsistencyLevel:eventual if ID unknown.
      let isOwner = false;
      try {
        if (userAadId) {
          // Direct lookup: 200 = owner, 404 = not owner (no filter, always supported)
          await client.api(`/groups/${m365GroupId}/owners/${userAadId}`).select('id').get();
          isOwner = true;
        } else {
          // Fallback: filter requires ConsistencyLevel:eventual
          const ownersResp = await client
            .api(`/groups/${m365GroupId}/owners`)
            .header('ConsistencyLevel', 'eventual')
            .count(true)
            .filter(`userPrincipalName eq '${identifier}'`)
            .select('id').top(1).get();
          isOwner = (ownersResp?.value?.length ?? 0) > 0;
        }
      } catch { /* 404 = not an owner; other errors → treat as member */ }

      console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: isOwner=%s', isOwner);

      const targetSpGroupId: number | undefined = isOwner
        ? webProps?.AssociatedOwnerGroup?.Id as number | undefined
        : webProps?.AssociatedMemberGroup?.Id as number | undefined;
      console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: mapped to SP group id=%s', targetSpGroupId);

      const roles: string[] = [];
      if (targetSpGroupId) {
        const matchedRa = actionableGroupRas.find((ra: any) => ra.Member?.Id === targetSpGroupId);
        if (matchedRa) {
          rdbArray(matchedRa.RoleDefinitionBindings)
            .map((r: any) => r.Name as string)
            .filter((r) => !isSystemRole(r))
            .forEach((r) => { if (roles.indexOf(r) === -1) roles.push(r); });
        }
      }
      console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: resolved roles=%o', roles);
      return { roles, graphUnavailable: false };
    } catch (err) {
      console.debug('[SmartPermissions] checkM365ConnectedSiteRoles: failed', err);
      return { roles: [], graphUnavailable: true };
    }
  }

  // Checks whether the user is a member of any AAD group (M365 Group or Security Group)
  // that appears directly in the site's role assignments. These groups have claims login
  // names starting with c:0o.c| or c:0t.c| and are never returned by the siteusers
  // Groups expansion, so extractRoles misses them.
  //
  // Phase 1 — SharePoint REST: for each AAD group, ask SharePoint directly whether the
  // user is a member via sitegroups/getbyloginname/users?$filter=LoginName eq '...'.
  // This is a yes/no membership check that works without Graph when SharePoint can
  // resolve the group (returns 200). A 404 or error means SharePoint can't resolve it
  // that way, so those groups fall through to Phase 2.
  //
  // Phase 2 — Graph fallback: for groups REST couldn't resolve, use Graph
  // /users/{id}/transitiveMemberOf. Requires GroupMember.Read.All to be approved.
  // graphUnavailable is only true when REST failed AND Graph also failed.
  private async getAadGroupSiteRoles(
    siteUrl: string,
    userLoginName: string,
    userEmail: string,
    roleAssignments: any[],
  ): Promise<{ roles: string[]; graphUnavailable: boolean }> {
    // Include all non-user principals: SharePoint groups (pt=8) AND AAD/M365 groups (pt=4).
    // Direct user assignments (pt=1) are already handled by extractRoles.
    // We extend beyond AAD-only because M365-connected sites often use classic SP groups
    // ("IT Members") whose membership is backed by M365 Group but not lazily synced into
    // the user's siteusers Groups expansion — so extractRoles misses them too.
    const groupRas = roleAssignments.filter((ra: any) => {
      const pt: number = ra.Member?.PrincipalType ?? 1;
      return pt === 4 || pt === 8;
    });
    // Skip groups whose roles are all system roles (e.g. Limited Access System Group).
    const actionableGroupRas = groupRas.filter((ra: any) =>
      rdbArray(ra.RoleDefinitionBindings).some((r: any) => !isSystemRole(r.Name as string)),
    );
    console.debug('[SmartPermissions] getAadGroupSiteRoles: actionable groups=%o',
      actionableGroupRas.map((ra: any) => ({
        id: ra.Member?.Id,
        login: ra.Member?.LoginName,
        title: ra.Member?.Title,
        principalType: ra.Member?.PrincipalType,
        roles: rdbArray(ra.RoleDefinitionBindings).map((r: any) => r.Name),
      })));
    if (actionableGroupRas.length === 0) {
      console.debug('[SmartPermissions] getAadGroupSiteRoles: no actionable groups — nothing to check');
      return { roles: [], graphUnavailable: false };
    }

    // Helper: is this group an AAD/M365/Security group (vs a classic SP group)?
    // AAD groups get a Graph fallback when REST can't confirm; SP groups do not.
    const isAadGroup = (login: string): boolean => {
      const l = login.toLowerCase();
      return l.startsWith('c:0o.c|') || l.startsWith('c:0t.c|') || l.startsWith('c:0p.c|');
    };

    // ── Phase 1: SharePoint REST membership check ─────────────────────────
    const roles: string[] = [];
    const needsGraph: any[] = [];
    const encUser = encodeURIComponent(`'${odata(userLoginName)}'`);

    await Promise.all(actionableGroupRas.map(async (ra: any) => {
      const groupLogin: string = ra.Member?.LoginName ?? '';
      const groupId: number | undefined = ra.Member?.Id;
      const pt: number = ra.Member?.PrincipalType ?? 0;
      const label: string = ra.Member?.Title || groupLogin;

      // SP groups (pt=8): look up by numeric ID — reliable because Member.LoginName in
      // role assignments is the group's display name, NOT its internal login name, so
      // getbyloginname('Team Members') returns 404. getbyid is unambiguous.
      // AAD groups (pt=4): use getbyloginname with the claims-format login name.
      const url = (pt === 8 && groupId)
        ? `${siteUrl}/_api/web/sitegroups/getbyid(${groupId})/users?$filter=LoginName eq ${encUser}&$top=1&$select=LoginName`
        : `${siteUrl}/_api/web/sitegroups/getbyloginname(${encodeURIComponent(`'${odata(groupLogin)}'`)})/users?$filter=LoginName eq ${encUser}&$top=1&$select=LoginName`;

      try {
        const data = await this.getJson(url);
        const found = valueArray(data).length > 0;
        console.debug('[SmartPermissions] getAadGroupSiteRoles: REST "%s" (id=%s pt=%s) → found=%s', label, groupId, pt, found);
        if (found) {
          rdbArray(ra.RoleDefinitionBindings)
            .map((r: any) => r.Name as string)
            .filter((r) => !isSystemRole(r))
            .forEach((r) => { if (roles.indexOf(r) === -1) roles.push(r); });
        } else if (isAadGroup(groupLogin)) {
          console.debug('[SmartPermissions] getAadGroupSiteRoles: "%s" → empty (AAD, ambiguous), queuing for Graph', label);
          needsGraph.push(ra);
        } else {
          console.debug('[SmartPermissions] getAadGroupSiteRoles: "%s" → empty (SP group, not a member)', label);
        }
      } catch (err) {
        console.debug('[SmartPermissions] getAadGroupSiteRoles: "%s" → REST error=%o', label, err);
        if (isAadGroup(groupLogin)) {
          needsGraph.push(ra);
        }
      }
    }));

    if (roles.length > 0) {
      console.debug('[SmartPermissions] getAadGroupSiteRoles: resolved via REST=%o', roles);
      return { roles, graphUnavailable: false };
    }
    if (needsGraph.length === 0) {
      // No AAD groups need Graph, but SP group member lists only contain static entries.
      // For M365 Group-connected sites, members are resolved dynamically — the user
      // won't appear in sitegroups/users even though they have access. Check the site's
      // connected M365 Group via Graph and map back to the SP group's roles.
      console.debug('[SmartPermissions] getAadGroupSiteRoles: no direct AAD groups — checking M365-connected site');
      return await this.checkM365ConnectedSiteRoles(siteUrl, userLoginName, userEmail, actionableGroupRas);
    }

    // ── Phase 2: Graph fallback for groups REST couldn't resolve ──────────
    const identifier = userEmail ||
      (userLoginName.includes('|') ? userLoginName.split('|').pop() : null) ||
      userLoginName;
    console.debug('[SmartPermissions] getAadGroupSiteRoles: %d group(s) need Graph, identifier=%s',
      needsGraph.length, identifier);
    if (!identifier) return { roles: [], graphUnavailable: true };

    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const memberOfData = await client
        .api(`/users/${encodeURIComponent(identifier)}/transitiveMemberOf`)
        .select('id')
        .top(999)
        .get();

      const userGroupIds = new Set<string>(
        (memberOfData?.value ?? []).map((g: any) => (g.id as string).toLowerCase()),
      );
      console.debug('[SmartPermissions] getAadGroupSiteRoles: user in %d AAD group(s) via Graph', userGroupIds.size);

      for (const ra of needsGraph) {
        const guid = ((ra.Member?.LoginName ?? '').split('|').pop() ?? '').toLowerCase();
        const matched = guid && userGroupIds.has(guid);
        console.debug('[SmartPermissions] getAadGroupSiteRoles: Graph guid=%s matched=%s', guid, matched);
        if (matched) {
          rdbArray(ra.RoleDefinitionBindings)
            .map((r: any) => r.Name as string)
            .filter((r) => !isSystemRole(r))
            .forEach((r) => { if (roles.indexOf(r) === -1) roles.push(r); });
        }
      }
      console.debug('[SmartPermissions] getAadGroupSiteRoles: resolved via Graph=%o', roles);
      return { roles, graphUnavailable: false };
    } catch (err) {
      console.debug('[SmartPermissions] getAadGroupSiteRoles: Graph failed', err);
      return { roles: [], graphUnavailable: true };
    }
  }

  // ── Sharing links ─────────────────────────────────────────────────────────

  async getSharingLinks(
    siteUrl: string,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
  ): Promise<SharingLinkEntry[]> {
    const { hostname, pathname } = new URL(siteUrl);
    const sitePath = pathname.replace(/\/$/, ''); // '' for root site, '/sites/mysite' for subsite
    // Path notation needs a trailing colon before sub-resources: hostname:/sites/foo:/drives
    // Root sites use plain hostname notation: hostname/drives
    const graphSite = sitePath ? `${hostname}:${sitePath}:` : hostname;
    const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

    onProgress('Loading document libraries…');
    const drivesResp = await graphClient
      .api(`/sites/${graphSite}/drives`)
      .select('id,name')
      .get();
    const drives: any[] = drivesResp?.value ?? [];

    if (drives.length === 0) return [];

    // ── Phase 1: scan all drives in parallel, collect shared items ────────────
    // Each drive independently paginates its delta and builds a Map<itemId, item>.
    type DriveItems = { drive: any; items: Map<string, any> };

    const driveResults = await this.runConcurrent(
      drives.map((drive) => async (): Promise<DriveItems | undefined> => {
        const sharedItems: Map<string, any> = new Map();
        let apiPath: string =
          `/drives/${drive.id}/root/delta?$select=id,name,webUrl,shared,parentReference`;

        while (apiPath && !signal?.aborted) {
          const resp = await graphClient.api(apiPath).get();
          for (const item of (resp?.value ?? []) as any[]) {
            if (item.shared) sharedItems.set(item.id, item);
          }
          const next: string | undefined = resp['@odata.nextLink'];
          apiPath = next ? next.replace('https://graph.microsoft.com/v1.0', '') : '';
        }

        onProgress(`${drive.name}: ${sharedItems.size} shared item(s) found`);
        return { drive, items: sharedItems };
      }),
      this.scanConcurrency,
    );

    // ── Phase 2: fetch permissions for all shared items across all drives ─────
    // One unified task list so the concurrency budget is shared globally,
    // not reset per library. O(1) item lookup via the Map from Phase 1.
    const results: SharingLinkEntry[] = [];

    const permTasks = (driveResults as (DriveItems | undefined)[])
      .filter((r): r is DriveItems => r !== undefined)
      .reduce<(() => Promise<undefined>)[]>((acc, { drive, items }) =>
        acc.concat(Array.from(items.entries()).map(([itemId, item]) => async (): Promise<undefined> => {
          if (signal?.aborted) return undefined;
          try {
            const permResp = await graphClient
              .api(`/drives/${drive.id}/items/${itemId}/permissions`)
              .get();
            for (const perm of (permResp?.value ?? []) as any[]) {
              if (!perm.link) continue;
              results.push({
                name: item.name,
                webUrl: item.webUrl,
                libraryName: drive.name,
                linkScope: perm.link.scope ?? 'unknown',
                linkType: perm.link.type ?? 'view',
                linkUrl: perm.link.webUrl ?? '',
                sharedWith: ((perm.grantedToIdentitiesV2 ?? []) as any[])
                  .map((g: any) => g.user?.displayName ?? g.group?.displayName ?? '')
                  .filter(Boolean)
                  .join(', '),
                expiresAt: perm.expirationDateTime,
              });
            }
          } catch { /* skip inaccessible items */ }
          return undefined;
        })),
      []);

    if (permTasks.length > 0) {
      onProgress(`Fetching permissions for ${permTasks.length} shared item(s)…`);
      await this.runConcurrent(permTasks, this.scanConcurrency);
    }

    return results;
  }

  // ── Permission groups ──────────────────────────────────────────────────────

  async getPermissionGroups(
    siteUrl: string,
    signal?: AbortSignal,
  ): Promise<PermissionGroup[]> {
    // Fetch all site groups with their role assignments
    const [groupsData, raData] = await Promise.all([
      this.getJson(
        `${siteUrl}/_api/web/sitegroups` +
          `?$select=Id,Title,LoginName,Description,Users/LoginName,Users/Title` +
          `&$expand=Users&$top=200`,
      ),
      this.getJson(
        `${siteUrl}/_api/web/RoleAssignments` +
          `?$expand=Member,RoleDefinitionBindings` +
          `&$select=Member/Id,Member/LoginName,RoleDefinitionBindings/Name`,
      ),
    ]);

    // Build a map of group login → roles from site-level role assignments
    const rolesByGroupId = new Map<number, string[]>();
    for (const ra of valueArray(raData)) {
      const memberId: number = ra.Member?.Id;
      const roles = rdbArray(ra.RoleDefinitionBindings)
        .map((r: any) => r.Name as string)
        .filter((r) => !isSystemRole(r));
      if (memberId && roles.length > 0) {
        rolesByGroupId.set(memberId, roles);
      }
    }

    return valueArray(groupsData).map((g: any): PermissionGroup => {
      const members: { loginName: string; displayName: string }[] = valueArray(g.Users)
        .filter((u: any) => !u.LoginName?.includes('_spo_') && !u.LoginName?.includes('app@sharepoint'))
        .map((u: any) => ({ loginName: u.LoginName, displayName: u.Title }));
      return {
        id: g.Id,
        title: g.Title,
        loginName: g.LoginName,
        description: g.Description ?? '',
        roles: rolesByGroupId.get(g.Id) ?? [],
        memberCount: members.length,
        members,
      };
    });
  }

  // ── External users ────────────────────────────────────────────────────────

  async getExternalUsers(siteUrl: string, signal?: AbortSignal): Promise<ExternalUserEntry[]> {
    const data = await this.getJson(
      `${siteUrl}/_api/web/siteusers?$filter=IsHiddenInUI eq false` +
      `&$select=LoginName,Title,Email,IsSiteAdmin&$orderby=Title&$top=2000`,
    );
    if (signal?.aborted) return [];
    const external = valueArray(data).filter(
      (u: any) => (u.LoginName ?? '').toLowerCase().indexOf('#ext#') !== -1,
    );
    const raw = await this.runConcurrent(
      external.map((u: any) => async () => {
        if (signal?.aborted) return undefined;
        try {
          const ud = await this.getJson(
            `${siteUrl}/_api/web/siteusers/getbyloginname('${encodeURIComponent(odata(u.LoginName))}')` +
            `?$expand=Groups&$select=LoginName,Groups/Title`,
          );
          return {
            loginName: u.LoginName,
            displayName: u.Title,
            email: u.Email ?? '',
            isSiteAdmin: !!u.IsSiteAdmin,
            groups: valueArray(ud.Groups).map((g: any) => g.Title as string),
          } as ExternalUserEntry;
        } catch {
          return { loginName: u.LoginName, displayName: u.Title, email: u.Email ?? '', isSiteAdmin: !!u.IsSiteAdmin, groups: [] } as ExternalUserEntry;
        }
      }),
      this.scanConcurrency,
    );
    return raw.filter((r): r is ExternalUserEntry => r !== undefined);
  }

  // ── Broken inheritance finder ──────────────────────────────────────────────

  async scanBrokenInheritance(
    siteUrl: string,
    includeHidden: boolean,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
  ): Promise<BrokenInheritanceEntry[]> {
    const results: BrokenInheritanceEntry[] = [];
    onProgress('Loading document libraries…');

    // Single request that includes HasUniqueRoleAssignments, avoiding a separate call per library.
    const filter = encodeURIComponent(
      includeHidden ? 'BaseTemplate eq 101' : 'BaseTemplate eq 101 and Hidden eq false',
    );
    let libs: any[] = [];
    try {
      const listsData = await this.getJson(
        `${siteUrl}/_api/web/lists?$filter=${filter}` +
        `&$select=Title,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
        `&$expand=RootFolder&$top=200`,
      );
      libs = valueArray(listsData);
    } catch { return results; }

    if (!includeHidden) libs = libs.filter((l: any) => !isSystemLibrary(l));
    if (signal?.aborted) return results;

    onProgress(`Scanning ${libs.length} librar${libs.length !== 1 ? 'ies' : 'y'}…`);

    await this.runConcurrent(
      libs.map((lib: any) => async () => {
        if (signal?.aborted) return;
        onProgress(`Scanning library: ${lib.Title}`);
        const libUrl = lib.RootFolder?.ServerRelativeUrl ?? '';
        if (lib.HasUniqueRoleAssignments) {
          results.push({ objectType: 'Library', name: lib.Title, serverRelativeUrl: libUrl, depth: 1 });
        }
        try {
          await this.walkForUniquePerms(siteUrl, libUrl, 2, results, onProgress, signal);
        } catch { /* skip inaccessible library */ }
      }),
      this.scanConcurrency,
    );

    return results;
  }

  private async walkForUniquePerms(
    siteUrl: string,
    folderUrl: string,
    depth: number,
    results: BrokenInheritanceEntry[],
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
  ): Promise<void> {
    if (signal?.aborted) return;
    const enc = odata(folderUrl);

    let subFolders: any[] = [];
    let files: any[] = [];
    try {
      const [fData, fiData] = await Promise.all([
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Folders` +
          `?$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/HasUniqueRoleAssignments` +
          `&$expand=ListItemAllFields&$top=2000`,
        ),
        this.getJson(
          `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${enc}')/Files` +
          `?$select=Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments` +
          `&$expand=ListItemAllFields&$top=2000`,
        ),
      ]);
      subFolders = valueArray(fData);
      files = valueArray(fiData);
    } catch { return; }

    const visibleFolders = subFolders.filter(
      (f: any) => !f.Name.startsWith('_') && f.Name.toLowerCase() !== 'forms',
    );

    // Record all unique-perms items at this level first (synchronous, no concurrency issue).
    for (const folder of visibleFolders) {
      if (signal?.aborted) return;
      onProgress(folder.Name);
      if (folder.ListItemAllFields?.HasUniqueRoleAssignments) {
        results.push({ objectType: 'Folder', name: folder.Name, serverRelativeUrl: folder.ServerRelativeUrl, depth });
      }
    }
    for (const file of files) {
      if (signal?.aborted) return;
      if (file.ListItemAllFields?.HasUniqueRoleAssignments) {
        results.push({ objectType: 'File', name: file.Name, serverRelativeUrl: file.ServerRelativeUrl, depth });
      }
    }

    // Recurse into non-empty subfolders concurrently.
    await this.runConcurrent(
      visibleFolders
        .filter((f: any) => !signal?.aborted && (f.ItemCount ?? 0) > 0)
        .map((folder: any) => () =>
          this.walkForUniquePerms(siteUrl, folder.ServerRelativeUrl, depth + 1, results, onProgress, signal),
        ),
      this.scanConcurrency,
    );
  }

  // ── Helpers ───────────────────────────────────────────────────────────────

  private toPermissionInfoList(roleAssignments: any[]): UserPermissionInfo[] {
    const result: UserPermissionInfo[] = [];
    for (const ra of roleAssignments) {
      const roles = rdbArray(ra.RoleDefinitionBindings)
        .map((r: any) => r.Name as string)
        .filter((r) => !isSystemRole(r));
      if (roles.length > 0) {
        result.push({
          loginName: ra.Member?.LoginName ?? '',
          displayName: ra.Member?.Title ?? '',
          principalType: principalTypeLabel(ra.Member?.PrincipalType ?? 1),
          roles,
        });
      }
    }
    return result;
  }

  private extractRoles(
    roleAssignments: any[],
    userLogin: string,
    groupLogins: Set<string>,
  ): string[] {
    const roles = new Set<string>();
    for (const ra of roleAssignments) {
      const memberLogin: string = ra.Member?.LoginName ?? '';
      const match =
        memberLogin.toLowerCase() === userLogin.toLowerCase() ||
        groupLogins.has(memberLogin.toLowerCase());
      if (match) {
        for (const rb of rdbArray(ra.RoleDefinitionBindings)) {
          roles.add(rb.Name);
        }
      }
    }
    return Array.from(roles);
  }
}
