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
// M365 Groups:      c:0o.c|federateddirectoryclaimprovider|{GUID}
// Security Groups:  c:0t.c|tenant|{GUID}  /  c:0p.c|s2s|{GUID}
function extractAadGroupId(loginName: string): string | null {
  const last = loginName.split('|').pop() ?? '';
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(last) ? last : null;
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
  ): Promise<{ entries: PermissionEntry[]; groupPermissionDenied: boolean }> {
    const entries: PermissionEntry[] = [];
    const flags = { groupPermissionDenied: false };
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

    return { entries, groupPermissionDenied: flags.groupPermissionDenied };
  }

  private async scanSite(
    siteUrl: string,
    options: ReportOptions,
    entries: PermissionEntry[],
    onProgress: (progress: ScanProgress) => void,
    signal?: AbortSignal,
    onEntry?: (entry: PermissionEntry) => void,
    flags?: { groupPermissionDenied: boolean },
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
    } catch {
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
          } catch { /* fall back to site perms */ }
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
  ): Promise<{ hasUnique: boolean; users: UserPermissionInfo[] }> {
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
    } catch {
      return { hasUnique: false, users: [] };
    }
  }

  // Checks a single tree node for external users without re-fetching HasUniqueRoleAssignments
  // (already known from the tree). Returns false immediately for inherited nodes.
  async scanNodeForExternalUsers(
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<boolean> {
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
    } catch {
      return false;
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
            `?$select=LoginName,Title,IsHiddenInUI&$top=${this.groupMemberCap + 1}`,
        );
        const all = valueArray(data)
          .filter((u: any) => !u.IsHiddenInUI)
          .map(
            (u: any): UserPermissionInfo => ({
              loginName: u.LoginName,
              displayName: u.Title,
              principalType: 'User',
              roles: [],
              isGroupMember: true,
            }),
          )
          .sort((a, b) => a.displayName.localeCompare(b.displayName));
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
      } catch {
        return [];
      }
    }

    if (principalType === 'SecurityGroup') {
      const groupId = extractAadGroupId(loginName);
      if (!groupId) return [];
      const result = await this._getAadGroupMembers(groupId);
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

  private async _getAadGroupMembers(groupId: string): Promise<UserPermissionInfo[] | null> {
    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const result = await client
        .api(`/groups/${groupId}/members`)
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
  ): Promise<{ fullSiteAccess: boolean; items: PermissionEntry[] }> {
    onProgress('Loading user info…');

    let userTitle = userLoginName;
    let groupLogins = new Set<string>();

    try {
      const userData = await this.getJson(
        `${siteUrl}/_api/web/siteusers/getbyloginname('${encodeURIComponent(odata(userLoginName))}')` +
          `?$expand=Groups&$select=Title,LoginName,Groups/LoginName`,
      );
      userTitle = userData.Title ?? userLoginName;
      const groups = valueArray(userData.Groups);
      // Normalize to lowercase — role-assignment member logins can differ in case
      // from user-groups logins depending on the SharePoint REST endpoint used.
      groupLogins = new Set(groups.map((g: any) => (g.LoginName as string).toLowerCase()));
    } catch { /* proceed without groups */ }

    // ── Site roles ──
    onProgress('Loading site permissions…');
    let siteRoles: string[] = [];
    let webData: any = null;

    try {
      webData = await this.getJson(
        `${siteUrl}/_api/web` +
          `?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
      );
      siteRoles = this.extractRoles(
        valueArray(webData.RoleAssignments),
        userLoginName,
        groupLogins,
      ).filter((r) => r.toLowerCase() !== 'limited access');
    } catch { /* no site roles */ }

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
      return { fullSiteAccess: true, items: siteEntry ? [siteEntry] : [] };
    }

    // Member/Visitor with site-level access and no broken-inheritance libraries —
    // all content is accessible via site inheritance; no scan needed.
    const hasSiteAccess = siteRoles.length > 0;
    if (hasSiteAccess && !isOwner && libs.every((l) => !l.HasUniqueRoleAssignments)) {
      return { fullSiteAccess: false, items: siteEntry ? [siteEntry] : [] };
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
        ).filter((r) => r.toLowerCase() !== 'limited access');

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

    return { fullSiteAccess: false, items };
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
          (r) => r.toLowerCase() !== 'limited access',
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
          (r) => r.toLowerCase() !== 'limited access',
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
        .filter((r) => r.toLowerCase() !== 'limited access');
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
        .filter((r) => r.toLowerCase() !== 'limited access');
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
