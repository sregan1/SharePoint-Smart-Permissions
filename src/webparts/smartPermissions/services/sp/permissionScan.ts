import {
  UserPermissionInfo, PermissionEntry, ReportOptions, ReportScope, ObjectType, ScanProgress,
} from '../../models/models';
import {
  SpApiClient, TaskQueue, valueArray, isPermissionDenied, isSystemLibrary, isLibraryTemplate,
  folderApi, fileApi, listApi, toPermissionInfoList,
} from './spCore';
import { getAllSites } from './siteDiscovery';
import { getGroupMembers } from './groups';

// ── Permissions Report scan ──────────────────────────────────────────────────

export async function scanPermissions(client: SpApiClient, 
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
      const sites = await getAllSites(client, options.siteUrl, signal);
      for (const site of sites) {
        if (signal?.aborted) break;
        emit0(`Scanning: ${site.title}`);
        await scanSiteTree(client, site.url, options, entries, onProgress, signal, onEntry, flags);
      }
    } else {
      await scanSiteTree(client, options.siteUrl, options, entries, onProgress, signal, onEntry, flags);
    }

    return { entries, groupPermissionDenied: flags.groupPermissionDenied, roleAssignmentsDenied: flags.roleAssignmentsDenied };
  }

  // Scans a web and, when includeSubsites is set, recurses depth-first into
  // its subwebs (each subweb re-fetches its own /webs).
async function scanSiteTree(client: SpApiClient, 
    siteUrl: string,
    options: ReportOptions,
    entries: PermissionEntry[],
    onProgress: (progress: ScanProgress) => void,
    signal?: AbortSignal,
    onEntry?: (entry: PermissionEntry) => void,
    flags?: { groupPermissionDenied: boolean; roleAssignmentsDenied: boolean },
  ): Promise<void> {
    await scanSite(client, siteUrl, options, entries, onProgress, signal, onEntry, flags);
    if (!options.includeSubsites || signal?.aborted) return;

    let subwebs: any[] = [];
    try {
      subwebs = await client.getJsonPaged(
        `${siteUrl}/_api/web/webs?$select=Title,Url&$top=500`,
        signal,
      );
    } catch { return; /* no access to subwebs — skip silently */ }

    for (const web of subwebs) {
      if (signal?.aborted) break;
      if (!web.Url) continue;
      onProgress({ message: `Scanning subsite: ${web.Title ?? web.Url}`, scanned: entries.length, libsDone: 0, libsTotal: 0 });
      await scanSiteTree(client, web.Url, options, entries, onProgress, signal, onEntry, flags);
    }
  }

async function scanSite(client: SpApiClient, 
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
      const webData = await client.getJson(
        `${siteUrl}/_api/web` +
          `?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
      );
      sitePerms = toPermissionInfoList(valueArray(webData.RoleAssignments));
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
      const denied = isPermissionDenied(err);
      if (denied && flags) flags.roleAssignmentsDenied = true;
      // Fallback: no role assignments, just record the site.
      try {
        const webData = await client.getJson(
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
          scanIncomplete: denied ? undefined : true,
        };
        entries.push(siteEntry);
        onEntry?.(siteEntry);
      } catch { /* skip */ }
    }

    if (options.scope !== ReportScope.Site) {
      // ── Lists & libraries ─────────────────────────────────────────────────
      // All visible lists are included (generic lists carry permissions too);
      // only library-like templates are walked for folders/files. The query is
      // kept simple ($filter without deep $expand) so SPO applies it reliably.
      emit('Loading lists and libraries…');
      let libs: any[] = [];

      const filterClause = options.includeHidden
        ? ''
        : `$filter=${encodeURIComponent('Hidden eq false')}&`;

      try {
        libs = await client.getJsonPaged(
          `${siteUrl}/_api/web/lists?${filterClause}` +
            `$select=Title,BaseTemplate,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=500`,
          signal,
        );
      } catch { /* no lists available */ }

      if (!options.includeHidden) {
        libs = libs.filter((l: any) => !isSystemLibrary(l));
      }
      if (options.libraryUrls) {
        const allowed = new Set(options.libraryUrls);
        libs = libs.filter((l: any) => allowed.has(l.RootFolder?.ServerRelativeUrl ?? ''));
      }

      libsTotal = libs.length;
      emit('Starting scan…');

      // Lists are scanned concurrently (capped at scanConcurrency); each list
      // writes into its own local array so concurrent walks don't interleave
      // entries. Arrays are concatenated in original list order after the
      // drain, preserving the parent-before-child ordering the export relies
      // on. Within one library the folder walk stays a sequential DFS.
      const libResults: PermissionEntry[][] = libs.map(() => []);
      const queue = new TaskQueue(client.scanConcurrency);

      libs.forEach((lib: any, libIndex: number) => {
        queue.add(async () => {
          if (signal?.aborted) return;
          emit(`Scanning: ${lib.Title}`);
          const local = libResults[libIndex];
          const libUrl: string = lib.RootFolder?.ServerRelativeUrl ?? '';
          const isLib = isLibraryTemplate(lib.BaseTemplate ?? 0);

          let libPerms = sitePerms;
          let libIncomplete = false;
          if (lib.HasUniqueRoleAssignments && libUrl) {
            try {
              const raData = await client.getJson(
                `${listApi(siteUrl, libUrl)}/RoleAssignments` +
                  `?$expand=Member,RoleDefinitionBindings` +
                  `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind`,
              );
              libPerms = toPermissionInfoList(valueArray(raData));
            } catch (err: any) {
              if (isPermissionDenied(err)) {
                if (flags) flags.roleAssignmentsDenied = true;
              } else {
                // Transient failure, not "no access" — this item genuinely has
                // unique permissions we couldn't read; don't silently present
                // the parent's (site) permissions as if they were confirmed.
                libIncomplete = true;
              }
            }
          }

          const libEntry: PermissionEntry = {
            objectType: isLib ? ObjectType.Library : ObjectType.List,
            name: lib.Title,
            serverRelativeUrl: libUrl,
            siteUrl,
            hasUniquePermissions: !!lib.HasUniqueRoleAssignments,
            depth: 1,
            uniquePermissions: libPerms,
            noCrawl: lib.NoCrawl ? true : undefined,
            scanIncomplete: libIncomplete || undefined,
          };
          local.push(libEntry);
          onEntry?.(libEntry);

          if (
            isLib && libUrl &&
            (options.scope === ReportScope.Folder || options.scope === ReportScope.Item)
          ) {
            try {
              await walkFolder(client, 
                siteUrl,
                libUrl,
                2,
                1,
                options,
                local,
                libPerms,
                emit,
                signal,
                onEntry,
              );
            } catch { /* partial results OK */ }
          }

          libsDone++;
          emit(`Scanned: ${lib.Title}`);
        });
      });

      await queue.drain();
      for (const arr of libResults) entries.push(...arr);
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
                members = await getGroupMembers(client, siteUrl, up.displayName, up.loginName, up.principalType, signal, up.groupId);
              } catch (err: any) {
                if (err?.isGraphPermissionError && flags) flags.groupPermissionDenied = true;
                members = [];
              }
              memberCache.set(cacheKey, members);
            }
            members.forEach((m) => expanded.push({ ...m, roles: [...up.roles], roleTypeKinds: up.roleTypeKinds, sourceGroup: up.displayName }));
          }
        }
        entry.uniquePermissions = expanded;
        expandedCache.set(originalRef, expanded);
      }
    }
  }

async function walkFolder(client: SpApiClient, 
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

    const apiBase = folderApi(siteUrl, folderUrl);
    let subFolders: any[] = [];
    let files: any[] = [];
    let uniquePermsLoaded = false;

    const fetchEnriched = (): Promise<[any[], any[]]> =>
      Promise.all([
        client.getJsonPaged(
          `${apiBase}/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
          signal,
        ),
        options.scope === ReportScope.Item
          ? client.getJsonPaged(
              `${apiBase}/Files` +
                `?$select=Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments` +
                `&$expand=ListItemAllFields&$top=2000`,
              signal,
            )
          : Promise.resolve([]),
      ]);

    try {
      [subFolders, files] = await fetchEnriched();
      uniquePermsLoaded = true;
    } catch {
      // One retry of the enriched (unique-flags) query before giving up
      // richness — a single transient failure here previously dropped every
      // unique-permission flag for the whole subtree silently.
      try {
        [subFolders, files] = await fetchEnriched();
        uniquePermsLoaded = true;
      } catch {
        try {
          [subFolders, files] = await Promise.all([
            client.getJsonPaged(
              `${apiBase}/Folders?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
              signal,
            ),
            options.scope === ReportScope.Item
              ? client.getJsonPaged(
                  `${apiBase}/Files?$select=Name,ServerRelativeUrl&$top=2000`,
                  signal,
                )
              : Promise.resolve([]),
          ]);
        } catch { return; }
      }
    }
    // When the enriched query never succeeded, every entry below inherits the
    // parent's permissions as a best-effort fallback — flag them as
    // unknown/incomplete rather than presenting that as a confirmed result.
    const branchIncomplete = !uniquePermsLoaded;

    const visibleFolders = subFolders.filter(
      (f: any) =>
        !f.Name.startsWith('_') &&
        f.Name.toLowerCase() !== 'forms',
    );

    for (const subfolder of visibleFolders) {
      if (signal?.aborted) break;

      let folderPerms = parentPerms;
      let hasUnique = false;
      let folderIncomplete = branchIncomplete;

      if (uniquePermsLoaded && subfolder.ListItemAllFields?.HasUniqueRoleAssignments) {
        try {
          const raData = await client.getJson(
            `${folderApi(siteUrl, subfolder.ServerRelativeUrl)}/ListItemAllFields/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind`,
          );
          folderPerms = toPermissionInfoList(valueArray(raData));
          hasUnique = true;
        } catch (err: any) {
          // Permission-denied means we genuinely can't read this item's ACL — not
          // the same as a transient failure that should be flagged as incomplete.
          if (!isPermissionDenied(err)) folderIncomplete = true;
        }
      }

      const folderEntry: PermissionEntry = {
        objectType: ObjectType.Folder,
        name: subfolder.Name,
        serverRelativeUrl: subfolder.ServerRelativeUrl,
        siteUrl,
        hasUniquePermissions: hasUnique,
        depth,
        uniquePermissions: folderPerms,
        scanIncomplete: folderIncomplete || undefined,
      };
      results.push(folderEntry);
      onEntry?.(folderEntry);
      onProgress(subfolder.Name);

      const shouldRecurse =
        options.scope === ReportScope.Item || currentLevel < options.folderDepth;
      if (shouldRecurse) {
        try {
          await walkFolder(client, 
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
      let fileIncomplete = branchIncomplete;

      if (uniquePermsLoaded && file.ListItemAllFields?.HasUniqueRoleAssignments) {
        try {
          const raData = await client.getJson(
            `${fileApi(siteUrl, file.ServerRelativeUrl)}/ListItemAllFields/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind`,
          );
          filePerms = toPermissionInfoList(valueArray(raData));
          hasUnique = true;
        } catch (err: any) {
          if (!isPermissionDenied(err)) fileIncomplete = true;
        }
      }

      const fileEntry: PermissionEntry = {
        objectType: ObjectType.File,
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        siteUrl,
        hasUniquePermissions: hasUnique,
        depth,
        uniquePermissions: filePerms,
        scanIncomplete: fileIncomplete || undefined,
      };
      results.push(fileEntry);
      onEntry?.(fileEntry);
      onProgress(file.Name);
    }
  }
