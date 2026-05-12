import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  UserPermissionInfo, PermissionEntry, FolderFileNode, LibraryInfo,
  SiteCollectionInfo, SiteUserInfo, ReportOptions, ReportScope, ObjectType,
} from '../models/models';

// Escape single-quotes in OData string literals (SQL-style doubling).
function odata(s: string): string {
  return s.replace(/'/g, "''");
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

export class SharePointService {
  private readonly context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private async getJson(url: string): Promise<any> {
    const resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
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
      .filter((l: any) => includeHidden || (!l.NoCrawl && !l.IsSiteAssetsLibrary))
      .map((l: any) => ({
        title: l.Title,
        serverRelativeUrl: l.RootFolder?.ServerRelativeUrl ?? '',
      }));
  }

  async getSiteUsers(siteUrl: string, signal?: AbortSignal): Promise<SiteUserInfo[]> {
    const url =
      `${siteUrl}/_api/web/siteusers` +
      `?$filter=IsHiddenInUI eq false and PrincipalType eq 1` +
      `&$select=LoginName,Title&$orderby=Title&$top=2000`;
    const data = await this.getJson(url);
    return valueArray(data)
      .filter(
        (u: any) =>
          !u.LoginName?.includes('_spo_') &&
          !u.LoginName?.includes('app@sharepoint'),
      )
      .map((u: any) => ({ loginName: u.LoginName, displayName: u.Title }));
  }

  // ── Permissions Report scan ───────────────────────────────────────────────

  async scanPermissions(
    options: ReportOptions,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
  ): Promise<PermissionEntry[]> {
    const entries: PermissionEntry[] = [];

    if (options.allSites) {
      onProgress('Discovering site collections…');
      const sites = await this.getAllSites(options.siteUrl, signal);
      for (const site of sites) {
        if (signal?.aborted) break;
        onProgress(`Scanning: ${site.title}`);
        entries.push(...(await this.scanSite(site.url, options, onProgress, signal)));
      }
    } else {
      entries.push(...(await this.scanSite(options.siteUrl, options, onProgress, signal)));
    }

    return entries;
  }

  private async scanSite(
    siteUrl: string,
    options: ReportOptions,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
  ): Promise<PermissionEntry[]> {
    const entries: PermissionEntry[] = [];

    // ── Site-level permissions ────────────────────────────────────────────
    onProgress(`Loading site permissions: ${siteUrl}`);
    let sitePerms: UserPermissionInfo[] = [];

    try {
      const webData = await this.getJson(
        `${siteUrl}/_api/web` +
          `?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
      );
      sitePerms = this.toPermissionInfoList(valueArray(webData.RoleAssignments));
      entries.push({
        objectType: ObjectType.Site,
        name: webData.Title ?? siteUrl,
        serverRelativeUrl: webData.ServerRelativeUrl ?? '',
        siteUrl,
        hasUniquePermissions: true,
        depth: 0,
        uniquePermissions: sitePerms,
      });
    } catch {
      // Fallback: no role assignments, just record the site.
      try {
        const webData = await this.getJson(
          `${siteUrl}/_api/web?$select=Title,Url,ServerRelativeUrl`,
        );
        entries.push({
          objectType: ObjectType.Site,
          name: webData.Title ?? siteUrl,
          serverRelativeUrl: webData.ServerRelativeUrl ?? '',
          siteUrl,
          hasUniquePermissions: true,
          depth: 0,
          uniquePermissions: [],
        });
      } catch { /* skip */ }
    }

    if (options.scope === ReportScope.Site) return entries;

    // ── Libraries ─────────────────────────────────────────────────────────
    onProgress('Loading document libraries…');
    let libs: any[] = [];
    let libsHaveRoles = false;

    const hiddenFilter = options.includeHidden
      ? 'BaseTemplate eq 101'
      : 'BaseTemplate eq 101 and Hidden eq false and NoCrawl eq false';

    try {
      const listsData = await this.getJson(
        `${siteUrl}/_api/web/lists?$filter=${hiddenFilter}` +
          `&$select=Title,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl` +
          `,RoleAssignments/Member/LoginName,RoleAssignments/Member/Title` +
          `,RoleAssignments/Member/PrincipalType,RoleAssignments/RoleDefinitionBindings/Name` +
          `&$expand=RootFolder,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` +
          `&$top=200`,
      );
      libs = valueArray(listsData);
      libsHaveRoles = true;
    } catch {
      try {
        const listsData = await this.getJson(
          `${siteUrl}/_api/web/lists?$filter=${hiddenFilter}` +
            `&$select=Title,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=200`,
        );
        libs = valueArray(listsData);
      } catch { /* no libraries available */ }
    }

    for (const lib of libs) {
      if (signal?.aborted) break;
      onProgress(`Scanning library: ${lib.Title}`);

      const libPerms =
        lib.HasUniqueRoleAssignments && libsHaveRoles
          ? this.toPermissionInfoList(valueArray(lib.RoleAssignments))
          : sitePerms;

      entries.push({
        objectType: ObjectType.Library,
        name: lib.Title,
        serverRelativeUrl: lib.RootFolder?.ServerRelativeUrl ?? '',
        siteUrl,
        hasUniquePermissions: !!lib.HasUniqueRoleAssignments,
        depth: 1,
        uniquePermissions: libPerms,
      });

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
            onProgress,
            signal,
          );
        } catch { /* partial results OK */ }
      }
    }

    return entries;
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

      results.push({
        objectType: ObjectType.Folder,
        name: subfolder.Name,
        serverRelativeUrl: subfolder.ServerRelativeUrl,
        siteUrl,
        hasUniquePermissions: hasUnique,
        depth,
        uniquePermissions: folderPerms,
      });

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

      results.push({
        objectType: ObjectType.File,
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        siteUrl,
        hasUniquePermissions: hasUnique,
        depth,
        uniquePermissions: filePerms,
      });
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
    principalType: string,
    signal?: AbortSignal,
  ): Promise<UserPermissionInfo[]> {
    if (principalType !== 'SharePointGroup') return [];

    try {
      const data = await this.getJson(
        `${siteUrl}/_api/web/sitegroups/getbyname('${odata(groupName)}')/users` +
          `?$select=LoginName,Title,IsHiddenInUI&$top=2000`,
      );
      return valueArray(data)
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
    } catch {
      return [];
    }
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
      groupLogins = new Set(groups.map((g: any) => g.LoginName as string));
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
    const libFilter = includeHidden
      ? 'BaseTemplate eq 101'
      : 'BaseTemplate eq 101 and Hidden eq false and NoCrawl eq false';

    try {
      const listsData = await this.getJson(
        `${siteUrl}/_api/web/lists` +
          `?$filter=${libFilter}` +
          `&$select=Title,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl` +
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
            `&$select=Title,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=200`,
        );
        libs = valueArray(listsData);
      } catch { /* no libraries */ }
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

    if (isOwner && libs.every((l) => !l.HasUniqueRoleAssignments)) {
      return { fullSiteAccess: true, items: siteEntry ? [siteEntry] : [] };
    }

    const items: PermissionEntry[] = siteEntry ? [siteEntry] : [];

    for (const lib of libs) {
      if (signal?.aborted) break;
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

      try {
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
      } catch { /* partial results OK */ }
    }

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

      for (const subfolder of uniqueFolders) {
        if (signal?.aborted) break;
        try {
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
        } catch { /* skip this folder */ }
      }

      for (const file of uniqueFiles) {
        if (signal?.aborted) break;
        try {
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
        } catch { /* skip this file */ }
      }
    }

    // Recurse into all visible subfolders that have content.
    for (const subfolder of visibleFolders) {
      if (signal?.aborted) break;
      if ((subfolder.ItemCount ?? 0) > 0) {
        try {
          await this.walkFoldersForUser(
            siteUrl,
            subfolder.ServerRelativeUrl,
            userLogin,
            userDisplayName,
            groupLogins,
            results,
            depth + 1,
            onProgress,
            signal,
          );
        } catch { /* continue to next folder */ }
      }
    }
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
        groupLogins.has(memberLogin);
      if (match) {
        for (const rb of rdbArray(ra.RoleDefinitionBindings)) {
          roles.add(rb.Name);
        }
      }
    }
    return Array.from(roles);
  }
}
