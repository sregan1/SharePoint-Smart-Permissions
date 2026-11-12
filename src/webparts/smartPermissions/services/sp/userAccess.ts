import { MSGraphClientV3 } from '@microsoft/sp-http';
import { UserPermissionInfo, PermissionEntry, ObjectType } from '../../models/models';
import {
  SpApiClient, TaskQueue, odata, valueArray, rdbArray, isPermissionDenied, isGraphPermissionError,
  isSystemRole, isSystemLibrary, folderApi, fileApi, listApi, isLibraryTemplate, roleAccessTier,
  extractAadGroupInfo, isAbortError, debugLog, isSystemFolder,
} from './spCore';

export interface UserAccessResult {
  fullSiteAccess: boolean;
  items: PermissionEntry[];
  graphPermissionRequired: boolean;
  roleAssignmentsDenied: boolean;
  /**
   * True when part of the scan failed for a reason other than "access denied" (transient
   * errors, an unreadable group, a listing that fell back to less data). The result may
   * under-report access and is never reported as full site access.
   */
  incomplete?: boolean;
  /** Server-relative URLs (or site URLs) whose role assignments could not be read. */
  failedPaths?: string[];
}

// Follows @odata.nextLink so a membership check is exhaustive — a single .top(999)
// page silently truncated for users in more than 999 groups, turning a genuine member
// into a false "not a member" result.
async function fetchAllGraphIds(graph: MSGraphClientV3, path: string): Promise<Set<string>> {
  const ids = new Set<string>();
  let resp = await graph.api(path).select('id').top(999).get();
  for (;;) {
    (resp?.value ?? []).forEach((g: any) => { if (g?.id) ids.add((g.id as string).toLowerCase()); });
    const nextLink: string | undefined = resp?.['@odata.nextLink'];
    if (!nextLink) break;
    // nextLink is an absolute URL; the Graph client's .api() expects a path
    // relative to its configured base, so strip the version prefix.
    resp = await graph.api(nextLink.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, '')).get();
  }
  return ids;
}

async function fetchAllTransitiveGroupIds(graph: MSGraphClientV3, identifier: string): Promise<Set<string>> {
  return fetchAllGraphIds(graph, `/users/${encodeURIComponent(identifier)}/transitiveMemberOf`);
}

// Candidate Graph identifiers for a user, best first. The UPN embedded in the claims
// login (i:0#.f|membership|user@contoso.com) is what /users/{id} resolves reliably; the
// SharePoint Email property is the *mail* attribute, which can differ from the UPN.
function graphIdentifiers(userLoginName: string, userEmail: string): string[] {
  const out: string[] = [];
  const loginPart = userLoginName.includes('|') ? (userLoginName.split('|').pop() ?? '') : userLoginName;
  if (loginPart.includes('@')) out.push(loginPart);
  if (userEmail && !out.some((x) => x.toLowerCase() === userEmail.toLowerCase())) out.push(userEmail);
  return out;
}

// ── Access context ───────────────────────────────────────────────────────────
// Everything needed to decide whether a role assignment's principal covers the user,
// resolved once per scan and shared by the site, library, folder and file levels.

interface GraphGroups {
  /** Lowercased ids of every group the user belongs to (transitive). */
  member: Set<string>;
  /** Lowercased ids of every group the user owns. */
  owned: Set<string>;
}

interface AccessFlags {
  graphPermissionRequired: boolean;
  incomplete: boolean;
  roleAssignmentsDenied: boolean;
  failedPaths: string[];
  /** Libraries/folders/files seen with broken inheritance (whether or not they cover the user). */
  uniqueBelow: number;
}

interface AccessContext {
  userLogin: string;
  userLower: string;
  displayName: string;
  /** False for guests; "Everyone except external users" does not cover them. */
  isInternal: boolean;
  /** Lowercased login names of the SharePoint groups the user is a direct member of. */
  groupLogins: Set<string>;
  spGroupIds: Set<number>;
  flags: AccessFlags;
  getGraphGroups: () => Promise<GraphGroups | null>;
  getSpGroupMembers: (groupId: number) => Promise<any[]>;
}

const RA_SELECT =
  'Member/Id,Member/LoginName,Member/Title,Member/PrincipalType,' +
  'RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind';

function recordFailure(ctx: AccessContext, err: unknown, path: string): void {
  if (isAbortError(err)) return;
  if (isPermissionDenied(err)) ctx.flags.roleAssignmentsDenied = true;
  else ctx.flags.incomplete = true;
  ctx.flags.failedPaths.push(path);
}

function createAccessContext(
  client: SpApiClient,
  siteUrl: string,
  userLogin: string,
  userEmail: string,
  displayName: string,
  groupLogins: Set<string>,
  spGroupIds: Set<number>,
): AccessContext {
  const flags: AccessFlags = {
    graphPermissionRequired: false,
    incomplete: false,
    roleAssignmentsDenied: false,
    failedPaths: [],
    uniqueBelow: 0,
  };

  // Graph is only queried when a role assignment actually needs it, and at most once.
  let graphPromise: Promise<GraphGroups | null> | undefined;
  const loadGraphGroups = async (): Promise<GraphGroups | null> => {
    try {
      const graph: MSGraphClientV3 = await client.context.msGraphClientFactory.getClient('3');
      let userId = '';
      let sawOtherError = false;
      for (const identifier of graphIdentifiers(userLogin, userEmail)) {
        try {
          const profile = await graph.api(`/users/${encodeURIComponent(identifier)}`).select('id').get();
          userId = profile?.id ?? '';
          if (userId) break;
        } catch (err) {
          if (isGraphPermissionError(err)) throw err;
          // 404 just means this identifier isn't the AAD user; try the next one.
          if ((err as any)?.statusCode !== 404) sawOtherError = true;
        }
      }
      if (!userId) {
        // A SharePoint-only account has no AAD object and legitimately belongs to no AAD
        // group. Only an unexplained lookup failure makes the result uncertain.
        if (sawOtherError) flags.incomplete = true;
        return null;
      }
      const member = await fetchAllTransitiveGroupIds(graph, userId);
      let owned = new Set<string>();
      try {
        owned = await fetchAllGraphIds(graph, `/users/${encodeURIComponent(userId)}/ownedObjects/microsoft.graph.group`);
      } catch (err) {
        // Owner claims (…_o) can't be verified without this; treat them as not matching.
        debugLog('[SmartPermissions] ownedObjects lookup failed', err);
      }
      debugLog('[SmartPermissions] access context: in %d group(s), owns %d', member.size, owned.size);
      return { member, owned };
    } catch (err) {
      if (isGraphPermissionError(err)) flags.graphPermissionRequired = true;
      else flags.incomplete = true;
      debugLog('[SmartPermissions] access context: Graph group lookup failed', err);
      return null;
    }
  };

  const spMemberCache = new Map<number, Promise<any[]>>();
  const getSpGroupMembers = (groupId: number): Promise<any[]> => {
    let p = spMemberCache.get(groupId);
    if (!p) {
      p = client
        .getJson(`${siteUrl}/_api/web/sitegroups/getbyid(${groupId})/users?$select=LoginName,PrincipalType&$top=5000`)
        .then((data) => valueArray(data))
        .catch((err) => {
          // Can't see who is in this group, so we can't rule it out.
          if (!isAbortError(err)) flags.incomplete = true;
          return [];
        });
      spMemberCache.set(groupId, p);
    }
    return p;
  };

  return {
    userLogin,
    userLower: userLogin.toLowerCase(),
    displayName,
    isInternal: !/#ext#/i.test(userLogin),
    groupLogins,
    spGroupIds,
    flags,
    getGraphGroups: () => {
      if (!graphPromise) graphPromise = loadGraphGroups();
      return graphPromise;
    },
    getSpGroupMembers,
  };
}

// Does a non-SharePoint-group principal (identified by claims login) cover the user?
async function claimCoversUser(ctx: AccessContext, login: string): Promise<boolean> {
  const l = login.toLowerCase();
  if (!l) return false;
  if (l === ctx.userLower) return true;
  // "Everyone": every authenticated user. "Everyone except external users": members only.
  if (l.startsWith('c:0(.s|true')) return true;
  if (l.startsWith('c:0-.f|rolemanager|spo-grid-all-users')) return ctx.isInternal;
  // AAD security group / M365 group (…|guid) or the group's owners (…|guid_o)
  if (l.startsWith('c:0t.c|') || l.startsWith('c:0o.c|') || l.startsWith('c:0p.c|')) {
    const aad = extractAadGroupInfo(login);
    if (!aad) return false;
    const groups = await ctx.getGraphGroups();
    if (!groups) return false;
    return (aad.isOwners ? groups.owned : groups.member).has(aad.id.toLowerCase());
  }
  return false;
}

async function memberCoversUser(ctx: AccessContext, member: any): Promise<boolean> {
  const login: string = member?.LoginName ?? '';
  const l = login.toLowerCase();
  if (!l) return false;
  if (ctx.groupLogins.has(l)) return true;
  const groupId: number | undefined = typeof member?.Id === 'number' ? member.Id : undefined;
  const isSpGroup = (member?.PrincipalType ?? 1) === 8;
  if (isSpGroup && groupId !== undefined && ctx.spGroupIds.has(groupId)) return true;
  if (await claimCoversUser(ctx, login)) return true;

  // A SharePoint group can contain AAD/M365 groups (the usual shape on M365-connected
  // sites) that the user's own siteusers Groups expansion never reports.
  if (isSpGroup && groupId !== undefined) {
    for (const m of await ctx.getSpGroupMembers(groupId)) {
      if (await claimCoversUser(ctx, m?.LoginName ?? '')) return true;
    }
  }
  return false;
}

// Merges every role-assignment source that covers the user into one role set.
async function resolveRoles(
  ctx: AccessContext,
  roleAssignments: any[],
): Promise<{ roles: string[]; roleTypeKinds: Record<string, number>; via: string[] }> {
  const roles = new Set<string>();
  const roleTypeKinds: Record<string, number> = {};
  const via: string[] = [];
  for (const ra of roleAssignments) {
    const bindings = rdbArray(ra.RoleDefinitionBindings).filter((r: any) => !isSystemRole(r));
    if (bindings.length === 0) continue; // e.g. Limited Access only — skip the lookups
    if (!(await memberCoversUser(ctx, ra.Member))) continue;
    for (const rb of bindings) {
      roles.add(rb.Name);
      if (typeof rb.RoleTypeKind === 'number') roleTypeKinds[rb.Name] = rb.RoleTypeKind;
    }
    const label: string = ra.Member?.Title || ra.Member?.LoginName || '';
    if (label && via.indexOf(label) === -1) via.push(label);
  }
  return { roles: Array.from(roles), roleTypeKinds, via };
}

function accessEntryPerms(
  ctx: AccessContext,
  resolved: { roles: string[]; roleTypeKinds: Record<string, number>; via: string[] },
): UserPermissionInfo[] {
  return [{
    loginName: ctx.userLogin,
    displayName: ctx.displayName,
    principalType: 'User',
    roles: resolved.roles,
    roleTypeKinds: resolved.roleTypeKinds,
    accessVia: resolved.via.length > 0 ? resolved.via : undefined,
  }];
}

// ── User Access scan ─────────────────────────────────────────────────────────

export async function getUserAccess(client: SpApiClient,
    siteUrl: string,
    userLoginName: string,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
    includeHidden = false,
  ): Promise<UserAccessResult> {
    onProgress('Loading user info…');

    let userTitle = userLoginName;
    let userEmail = '';
    let groupLogins = new Set<string>();
    let spGroupIds = new Set<number>();
    let userInfoFailed = false;

    try {
      const userData = await client.getJson(
        `${siteUrl}/_api/web/siteusers/getbyloginname('${encodeURIComponent(odata(userLoginName))}')` +
          `?$expand=Groups&$select=Title,LoginName,Email,Groups/LoginName,Groups/Id`,
      );
      userTitle = userData.Title ?? userLoginName;
      userEmail = userData.Email ?? '';
      const groups = valueArray(userData.Groups);
      // Normalize to lowercase — role-assignment member logins can differ in case
      // from user-groups logins depending on the SharePoint REST endpoint used.
      groupLogins = new Set(groups.map((g: any) => (g.LoginName as string).toLowerCase()));
      spGroupIds = new Set(groups.filter((g: any) => typeof g.Id === 'number').map((g: any) => g.Id as number));
      debugLog('[SmartPermissions] getUserAccess: user=%s email=%s SP-groups=%o',
        userLoginName, userEmail, Array.from(groupLogins));
    } catch (err) {
      // A user who has never visited the site has no siteusers entry; that is not a
      // failure. Anything else leaves us unable to see their SharePoint group memberships.
      if (isPermissionDenied(err) || /HTTP 5\d\d/.test(String((err as any)?.message ?? ''))) userInfoFailed = true;
      debugLog('[SmartPermissions] getUserAccess: siteusers fetch failed', err);
    }

    const ctx = createAccessContext(client, siteUrl, userLoginName, userEmail, userTitle, groupLogins, spGroupIds);
    if (userInfoFailed) ctx.flags.incomplete = true;

    // ── Site roles ──
    onProgress('Loading site permissions…');
    let siteRoles: string[] = [];
    let siteRoleTypeKinds: Record<string, number> = {};
    let siteVia: string[] = [];
    let webData: any = null;
    let roleAssignmentsDenied = false;

    try {
      webData = await client.getJson(
        `${siteUrl}/_api/web` +
          `?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments` +
          `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
      );
      const allRa = valueArray(webData.RoleAssignments);
      debugLog('[SmartPermissions] getUserAccess: site RoleAssignments=%o',
        allRa.map((ra: any) => ({
          login: ra.Member?.LoginName,
          principalType: ra.Member?.PrincipalType,
          roles: rdbArray(ra.RoleDefinitionBindings).map((r: any) => r.Name),
        })));
      const siteResolved = await resolveRoles(ctx, allRa);
      siteRoles = siteResolved.roles;
      siteRoleTypeKinds = siteResolved.roleTypeKinds;
      siteVia = siteResolved.via;
      debugLog('[SmartPermissions] getUserAccess: siteRoles after resolveRoles=%o', siteRoles);
    } catch (err) {
      debugLog('[SmartPermissions] getUserAccess: web RoleAssignments fetch failed', err);
      if (isPermissionDenied(err)) roleAssignmentsDenied = true;
      else if (!isAbortError(err)) ctx.flags.incomplete = true;
    }

    // Member-access fallback: RoleAssignments are not readable, but the user's SP group
    // memberships (fetched above without elevated rights) can be checked against the site's
    // three default associated groups, which are always readable. This gives a site-level
    // result without needing Manage Permissions.
    if (roleAssignmentsDenied) {
      onProgress('Checking site group membership…');
      try {
        const [web, ownerGrp, memberGrp, visitorGrp] = await Promise.all([
          client.getJson(`${siteUrl}/_api/web?$select=Title,Url,ServerRelativeUrl`).catch(() => null),
          client.getJson(`${siteUrl}/_api/web/AssociatedOwnerGroup?$select=Id,LoginName,Title`).catch(() => null),
          client.getJson(`${siteUrl}/_api/web/AssociatedMemberGroup?$select=Id,LoginName,Title`).catch(() => null),
          client.getJson(`${siteUrl}/_api/web/AssociatedVisitorGroup?$select=Id,LoginName,Title`).catch(() => null),
        ]);

        const matchedRoles: string[] = [];
        const via: string[] = [];
        // Same coverage rules as the full scan, so nested AAD/M365 membership counts here too.
        const check = async (grp: any, role: string): Promise<void> => {
          if (!grp?.LoginName) return;
          const covers = await memberCoversUser(ctx, {
            Id: grp.Id, LoginName: grp.LoginName, Title: grp.Title, PrincipalType: 8,
          });
          if (covers) {
            matchedRoles.push(role);
            via.push(grp.Title || grp.LoginName);
          }
        };
        await check(ownerGrp, 'Full Control');
        await check(memberGrp, 'Edit');
        await check(visitorGrp, 'Read');

        if (matchedRoles.length > 0) {
          const siteEntry: PermissionEntry = {
            objectType: ObjectType.Site,
            name: web?.Title ?? siteUrl,
            serverRelativeUrl: web?.ServerRelativeUrl ?? '',
            siteUrl,
            hasUniquePermissions: true,
            depth: 0,
            uniquePermissions: accessEntryPerms(ctx, { roles: matchedRoles, roleTypeKinds: {}, via }),
          };
          return {
            fullSiteAccess: false, items: [siteEntry],
            graphPermissionRequired: ctx.flags.graphPermissionRequired, roleAssignmentsDenied: true,
            incomplete: ctx.flags.incomplete || undefined,
          };
        }
      } catch { /* ignore fallback errors */ }

      return {
        fullSiteAccess: false, items: [],
        graphPermissionRequired: ctx.flags.graphPermissionRequired, roleAssignmentsDenied: true,
        incomplete: ctx.flags.incomplete || undefined,
      };
    }

    // Last-resort fallback for M365 Group-connected sites: the SP groups only hold static
    // members, and the resolver above already follows nested AAD/M365 groups, but keep the
    // older Graph-based check for shapes it can't see (e.g. a site-level owners lookup).
    let graphPermissionRequired = ctx.flags.graphPermissionRequired;
    if (siteRoles.length === 0 && webData) {
      debugLog('[SmartPermissions] getUserAccess: resolver returned empty — trying AAD group fallback');
      const aadResult = await getAadGroupSiteRoles(client,
        siteUrl,
        userLoginName,
        userEmail,
        valueArray(webData.RoleAssignments),
      );
      siteRoles = aadResult.roles;
      siteRoleTypeKinds = aadResult.roleTypeKinds;
      graphPermissionRequired = graphPermissionRequired || aadResult.graphUnavailable;
      debugLog('[SmartPermissions] getUserAccess: siteRoles after AAD fallback=%o graphPermissionRequired=%s',
        siteRoles, graphPermissionRequired);
    }

    // ── Lists & libraries ──
    onProgress('Loading lists and libraries…');
    let libs: any[] = [];
    let libsHaveRoles = false;
    let libsLoaded = false;
    const libFilterClause = includeHidden
      ? ''
      : `$filter=${encodeURIComponent('Hidden eq false')}&`;

    try {
      libs = await client.getJsonPaged(
        `${siteUrl}/_api/web/lists` +
          `?${libFilterClause}` +
          `$select=Title,BaseTemplate,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
          `,RoleAssignments/Member/Id,RoleAssignments/Member/LoginName,RoleAssignments/Member/Title,RoleAssignments/Member/PrincipalType` +
          `,RoleAssignments/RoleDefinitionBindings/Name,RoleAssignments/RoleDefinitionBindings/RoleTypeKind` +
          `&$expand=RootFolder,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` +
          `&$top=500`,
        signal,
      );
      libsHaveRoles = true;
      libsLoaded = true;
    } catch {
      try {
        libs = await client.getJsonPaged(
          `${siteUrl}/_api/web/lists` +
            `?${libFilterClause}` +
            `$select=Title,BaseTemplate,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=500`,
          signal,
        );
        libsLoaded = true;
      } catch {
        // No list data at all: nothing below the site level could be checked.
        ctx.flags.incomplete = true;
      }
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
        // A subsite can inherit from its parent web; the root web always reports true.
        hasUniquePermissions: webData.HasUniqueRoleAssignments !== false,
        depth: 0,
        uniquePermissions: accessEntryPerms(ctx, { roles: siteRoles, roleTypeKinds: siteRoleTypeKinds, via: siteVia }),
      };
    }

    const isOwner = siteRoles.some((r) => roleAccessTier(r, siteRoleTypeKinds[r]) === 'admin');

    const items: PermissionEntry[] = siteEntry ? [siteEntry] : [];

    // One shared queue caps total concurrent requests across all libraries
    // and all recursion levels at scanConcurrency. There is deliberately no early
    // "owner/member of the site, so all content is accessible" shortcut: a single folder
    // or file with broken inheritance anywhere below can remove that access, and the only
    // way to know is to look.
    const queue = new TaskQueue(client.scanConcurrency);

    for (const lib of libs) {
      queue.add(async () => {
        if (signal?.aborted) return;
        onProgress(`Scanning: ${lib.Title}`);
        const isLib = isLibraryTemplate(lib.BaseTemplate ?? 0);
        const libUrl: string = lib.RootFolder?.ServerRelativeUrl ?? '';

        try {
          if (lib.HasUniqueRoleAssignments) {
            ctx.flags.uniqueBelow++;
            let libRas: any[];
            if (libsHaveRoles) {
              libRas = valueArray(lib.RoleAssignments);
            } else if (libUrl) {
              // The expanded list query fell back, so this list's assignments weren't
              // loaded — fetch them directly instead of skipping the list.
              const raData = await client.getJson(
                `${listApi(siteUrl, libUrl)}/RoleAssignments?$expand=Member,RoleDefinitionBindings&$select=${RA_SELECT}`,
              );
              libRas = valueArray(raData);
            } else {
              libRas = [];
              ctx.flags.incomplete = true;
            }
            const resolved = await resolveRoles(ctx, libRas);
            if (resolved.roles.length > 0) {
              items.push({
                objectType: isLib ? ObjectType.Library : ObjectType.List,
                name: lib.Title,
                serverRelativeUrl: libUrl,
                siteUrl,
                hasUniquePermissions: true,
                depth: 1,
                uniquePermissions: accessEntryPerms(ctx, resolved),
                noCrawl: lib.NoCrawl ? true : undefined,
              });
            }
          }
        } catch (err) {
          recordFailure(ctx, err, `${siteUrl}${libUrl}`);
        }

        // Only library-like templates have Files/Folders to walk.
        if (isLib && libUrl) {
          await walkFoldersForUser(client, siteUrl, libUrl, ctx, items, 2, onProgress, queue, signal);
        }
      });
    }
    await queue.drain();

    // Full site access is only a safe claim when every check above completed cleanly and
    // nothing below the site breaks inheritance.
    const aborted = !!signal?.aborted;
    const clean = !aborted && libsLoaded && !ctx.flags.incomplete && !ctx.flags.roleAssignmentsDenied;
    const noUniqueBelow = ctx.flags.uniqueBelow === 0;
    if (siteEntry && ctx.flags.incomplete) siteEntry.scanIncomplete = true;

    return {
      fullSiteAccess: isOwner && clean && noUniqueBelow,
      items,
      graphPermissionRequired,
      roleAssignmentsDenied: roleAssignmentsDenied || ctx.flags.roleAssignmentsDenied,
      incomplete: ctx.flags.incomplete || undefined,
      failedPaths: ctx.flags.failedPaths.length > 0 ? ctx.flags.failedPaths : undefined,
    };
  }

async function walkFoldersForUser(client: SpApiClient,
    siteUrl: string,
    folderUrl: string,
    ctx: AccessContext,
    results: PermissionEntry[],
    depth: number,
    onProgress: (msg: string) => void,
    queue: TaskQueue,
    signal?: AbortSignal,
  ): Promise<void> {
    if (signal?.aborted) return;

    const apiBase = folderApi(siteUrl, folderUrl);
    let subFolders: any[] = [];
    let files: any[] = [];
    let uniquePermsLoaded = false;

    try {
      [subFolders, files] = await Promise.all([
        client.getJsonPaged(
          `${apiBase}/Folders` +
            `?$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
          signal,
        ),
        client.getJsonPaged(
          `${apiBase}/Files` +
            `?$select=Name,ServerRelativeUrl,ListItemAllFields/HasUniqueRoleAssignments` +
            `&$expand=ListItemAllFields&$top=2000`,
          signal,
        ),
      ]);
      uniquePermsLoaded = true;
    } catch {
      // Without the unique-permission flags we can't tell which children break
      // inheritance, and the files weren't listed at all — the result is incomplete.
      ctx.flags.incomplete = true;
      try {
        subFolders = await client.getJsonPaged(
          `${apiBase}/Folders?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
          signal,
        );
      } catch (err) {
        recordFailure(ctx, err, folderUrl);
        return;
      }
    }

    const visibleFolders = subFolders.filter((f: any) => !isSystemFolder(f.Name));

    if (uniquePermsLoaded) {
      const uniqueFolders = visibleFolders.filter(
        (f: any) => f.ListItemAllFields?.HasUniqueRoleAssignments,
      );
      const uniqueFiles = files.filter(
        (f: any) => f.ListItemAllFields?.HasUniqueRoleAssignments,
      );
      ctx.flags.uniqueBelow += uniqueFolders.length + uniqueFiles.length;

      const addUniqueItem = (
        objectType: ObjectType, name: string, serverRelativeUrl: string, api: string,
      ): void => {
        queue.add(async () => {
          if (signal?.aborted) return;
          try {
            const raData = await client.getJson(
              `${api}/ListItemAllFields/RoleAssignments?$expand=Member,RoleDefinitionBindings&$select=${RA_SELECT}`,
            );
            const resolved = await resolveRoles(ctx, valueArray(raData));
            if (resolved.roles.length > 0) {
              results.push({
                objectType,
                name,
                serverRelativeUrl,
                siteUrl,
                hasUniquePermissions: true,
                depth,
                uniquePermissions: accessEntryPerms(ctx, resolved),
              });
            }
          } catch (err) {
            // Previously swallowed by the queue, leaving a clean-looking but incomplete result.
            recordFailure(ctx, err, serverRelativeUrl);
          }
        });
      };

      for (const subfolder of uniqueFolders) {
        addUniqueItem(ObjectType.Folder, subfolder.Name, subfolder.ServerRelativeUrl,
          folderApi(siteUrl, subfolder.ServerRelativeUrl));
      }
      for (const file of uniqueFiles) {
        addUniqueItem(ObjectType.File, file.Name, file.ServerRelativeUrl,
          fileApi(siteUrl, file.ServerRelativeUrl));
      }
    }

    // Recurse into non-empty subfolders via the shared queue.
    for (const subfolder of visibleFolders) {
      if (signal?.aborted) break;
      if ((subfolder.ItemCount ?? 0) === 0) continue;
      queue.add(() =>
        walkFoldersForUser(client,
          siteUrl,
          subfolder.ServerRelativeUrl,
          ctx,
          results,
          depth + 1,
          onProgress,
          queue,
          signal,
        ),
      );
    }
  }

  // For M365 Group-connected sites the SP groups (Team Members, Team Owners, etc.) only
  // store static members. Actual M365 Group members are resolved dynamically by SharePoint
  // and never appear in sitegroups/users. This method:
  //   1. Gets the site's connected M365 Group GUID from _api/web?$select=GroupId
  //   2. Checks whether the user is in that M365 Group via Graph transitiveMemberOf
  //   3. Checks whether they are an Owner (→ AssociatedOwnerGroup roles) or a Member
  //      (→ AssociatedMemberGroup roles) and returns the corresponding permission roles.
async function checkM365ConnectedSiteRoles(client: SpApiClient, 
    siteUrl: string,
    userLoginName: string,
    userEmail: string,
    actionableGroupRas: any[],
  ): Promise<{ roles: string[]; roleTypeKinds: Record<string, number>; graphUnavailable: boolean }> {
    try {
      // GroupId lives on the site collection (_api/site), not the web.
      // AssociatedOwnerGroup/MemberGroup live on the web (_api/web).
      const [siteProps, webProps] = await Promise.all([
        client.getJson(`${siteUrl}/_api/site?$select=GroupId`),
        client.getJson(
          `${siteUrl}/_api/web?$expand=AssociatedOwnerGroup,AssociatedMemberGroup` +
          `&$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id`,
        ),
      ]);
      const m365GroupId: string = siteProps?.GroupId ?? '';
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: site GroupId=%s ownerGroupId=%s memberGroupId=%s',
        m365GroupId || '(none)', webProps?.AssociatedOwnerGroup?.Id, webProps?.AssociatedMemberGroup?.Id);

      if (!m365GroupId || m365GroupId === '00000000-0000-0000-0000-000000000000') {
        debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: site not M365-connected');
        return { roles: [], roleTypeKinds: {}, graphUnavailable: false };
      }

      const identifier = graphIdentifiers(userLoginName, userEmail)[0] ?? userLoginName;
      if (!identifier) return { roles: [], roleTypeKinds: {}, graphUnavailable: false };

      const graph: MSGraphClientV3 = await client.context.msGraphClientFactory.getClient('3');

      // Fetch transitive group memberships and the user's AAD object ID in parallel.
      // The AAD ID is needed for the ownership check below. A permission error on the
      // profile lookup (fetching another user's profile can need a wider Graph scope
      // than transitiveMemberOf) is tracked separately from "profile not found" so it
      // isn't silently treated the same as a resolved-but-empty result below.
      let userProfileError: any = null;
      const [userGroupIds, userProfileData] = await Promise.all([
        fetchAllTransitiveGroupIds(graph, identifier),
        graph.api(`/users/${encodeURIComponent(identifier)}`).select('id').get()
          .catch((e) => { userProfileError = e; return null; }),
      ]);
      const userAadId: string = userProfileData?.id ?? '';
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: user in %d group(s), site M365 Group present=%s, aadId=%s',
        userGroupIds.size, userGroupIds.has(m365GroupId.toLowerCase()), userAadId || '(unknown)');

      if (!userGroupIds.has(m365GroupId.toLowerCase())) {
        return { roles: [], roleTypeKinds: {}, graphUnavailable: false };
      }

      // User is in the M365 Group. Check if they're an Owner (→ AssociatedOwnerGroup →
      // Full Control) or just a Member (→ AssociatedMemberGroup → Edit).
      // Strategy: try GET /groups/{groupId}/owners/{userId} first (direct lookup, no filter
      // needed). Fall back to a filtered list with ConsistencyLevel:eventual if ID unknown.
      let isOwner = false;
      let ownerCheckFailed = false;
      try {
        if (userAadId) {
          // Direct lookup: 200 = owner, 404 = not owner (no filter, always supported)
          await graph.api(`/groups/${m365GroupId}/owners/${userAadId}`).select('id').get();
          isOwner = true;
        } else {
          // No AAD id resolved (profile lookup failed or lacked permission) — fall back
          // to a filtered list, which itself requires ConsistencyLevel:eventual.
          const ownersResp = await graph
            .api(`/groups/${m365GroupId}/owners`)
            .header('ConsistencyLevel', 'eventual')
            .count(true)
            .filter(`userPrincipalName eq '${identifier.replace(/'/g, "''")}'`)
            .select('id').top(1).get();
          isOwner = (ownersResp?.value?.length ?? 0) > 0;
        }
      } catch (err) {
        if (userAadId) {
          // Direct lookup 404 = genuinely not an owner — a safe, confirmed negative.
          isOwner = false;
        } else if (isGraphPermissionError(err) || (userProfileError && isGraphPermissionError(userProfileError))) {
          // Couldn't confirm the AAD id AND the fallback query also failed on
          // permissions — we have no reliable signal either way. Treating this as
          // "member" (as before) risked under-reporting an Owner's actual access.
          ownerCheckFailed = true;
        } else {
          isOwner = false; // filtered query resolved but found nothing = not an owner
        }
      }

      if (ownerCheckFailed) {
        debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: owner check inconclusive — Graph permission required');
        return { roles: [], roleTypeKinds: {}, graphUnavailable: true };
      }

      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: isOwner=%s', isOwner);

      const targetSpGroupId: number | undefined = isOwner
        ? webProps?.AssociatedOwnerGroup?.Id as number | undefined
        : webProps?.AssociatedMemberGroup?.Id as number | undefined;
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: mapped to SP group id=%s', targetSpGroupId);

      const roles: string[] = [];
      const roleTypeKinds: Record<string, number> = {};
      if (targetSpGroupId) {
        const matchedRa = actionableGroupRas.find((ra: any) => ra.Member?.Id === targetSpGroupId);
        if (matchedRa) {
          rdbArray(matchedRa.RoleDefinitionBindings)
            .filter((r: any) => !isSystemRole(r))
            .forEach((r: any) => {
              if (roles.indexOf(r.Name) === -1) roles.push(r.Name);
              if (typeof r.RoleTypeKind === 'number') roleTypeKinds[r.Name] = r.RoleTypeKind;
            });
        }
      }
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: resolved roles=%o', roles);
      return { roles, roleTypeKinds, graphUnavailable: false };
    } catch (err) {
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: failed', err);
      // Only a permission failure means "ask the admin to approve Graph access"; a
      // missing user or a transient error is not that.
      return { roles: [], roleTypeKinds: {}, graphUnavailable: isGraphPermissionError(err) };
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
async function getAadGroupSiteRoles(client: SpApiClient, 
    siteUrl: string,
    userLoginName: string,
    userEmail: string,
    roleAssignments: any[],
  ): Promise<{ roles: string[]; roleTypeKinds: Record<string, number>; graphUnavailable: boolean }> {
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
      rdbArray(ra.RoleDefinitionBindings).some((r: any) => !isSystemRole(r)),
    );
    debugLog('[SmartPermissions] getAadGroupSiteRoles: actionable groups=%o',
      actionableGroupRas.map((ra: any) => ({
        id: ra.Member?.Id,
        login: ra.Member?.LoginName,
        title: ra.Member?.Title,
        principalType: ra.Member?.PrincipalType,
        roles: rdbArray(ra.RoleDefinitionBindings).map((r: any) => r.Name),
      })));
    if (actionableGroupRas.length === 0) {
      debugLog('[SmartPermissions] getAadGroupSiteRoles: no actionable groups — nothing to check');
      return { roles: [], roleTypeKinds: {}, graphUnavailable: false };
    }

    // Helper: is this group an AAD/M365/Security group (vs a classic SP group)?
    // AAD groups get a Graph fallback when REST can't confirm; SP groups do not.
    const isAadGroup = (login: string): boolean => {
      const l = login.toLowerCase();
      return l.startsWith('c:0o.c|') || l.startsWith('c:0t.c|') || l.startsWith('c:0p.c|');
    };

    // ── Phase 1: SharePoint REST membership check ─────────────────────────
    const roles: string[] = [];
    const roleTypeKinds: Record<string, number> = {};
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
        const data = await client.getJson(url);
        const found = valueArray(data).length > 0;
        debugLog('[SmartPermissions] getAadGroupSiteRoles: REST "%s" (id=%s pt=%s) → found=%s', label, groupId, pt, found);
        if (found) {
          rdbArray(ra.RoleDefinitionBindings)
            .filter((r: any) => !isSystemRole(r))
            .forEach((r: any) => {
              if (roles.indexOf(r.Name) === -1) roles.push(r.Name);
              if (typeof r.RoleTypeKind === 'number') roleTypeKinds[r.Name] = r.RoleTypeKind;
            });
        } else if (isAadGroup(groupLogin)) {
          debugLog('[SmartPermissions] getAadGroupSiteRoles: "%s" → empty (AAD, ambiguous), queuing for Graph', label);
          needsGraph.push(ra);
        } else {
          debugLog('[SmartPermissions] getAadGroupSiteRoles: "%s" → empty (SP group, not a member)', label);
        }
      } catch (err) {
        debugLog('[SmartPermissions] getAadGroupSiteRoles: "%s" → REST error=%o', label, err);
        if (isAadGroup(groupLogin)) {
          needsGraph.push(ra);
        }
      }
    }));

    if (roles.length > 0) {
      debugLog('[SmartPermissions] getAadGroupSiteRoles: resolved via REST=%o', roles);
      return { roles, roleTypeKinds, graphUnavailable: false };
    }
    if (needsGraph.length === 0) {
      // No AAD groups need Graph, but SP group member lists only contain static entries.
      // For M365 Group-connected sites, members are resolved dynamically — the user
      // won't appear in sitegroups/users even though they have access. Check the site's
      // connected M365 Group via Graph and map back to the SP group's roles.
      debugLog('[SmartPermissions] getAadGroupSiteRoles: no direct AAD groups — checking M365-connected site');
      return await checkM365ConnectedSiteRoles(client, siteUrl, userLoginName, userEmail, actionableGroupRas);
    }

    // ── Phase 2: Graph fallback for groups REST couldn't resolve ──────────
    const identifier = graphIdentifiers(userLoginName, userEmail)[0] ?? userLoginName;
    debugLog('[SmartPermissions] getAadGroupSiteRoles: %d group(s) need Graph, identifier=%s',
      needsGraph.length, identifier);
    if (!identifier) return { roles: [], roleTypeKinds: {}, graphUnavailable: false };

    try {
      const graph: MSGraphClientV3 = await client.context.msGraphClientFactory.getClient('3');
      const userGroupIds = await fetchAllTransitiveGroupIds(graph, identifier);
      debugLog('[SmartPermissions] getAadGroupSiteRoles: user in %d AAD group(s) via Graph', userGroupIds.size);

      for (const ra of needsGraph) {
        const guid = ((ra.Member?.LoginName ?? '').split('|').pop() ?? '').toLowerCase();
        const matched = guid && userGroupIds.has(guid);
        debugLog('[SmartPermissions] getAadGroupSiteRoles: Graph guid=%s matched=%s', guid, matched);
        if (matched) {
          rdbArray(ra.RoleDefinitionBindings)
            .filter((r: any) => !isSystemRole(r))
            .forEach((r: any) => {
              if (roles.indexOf(r.Name) === -1) roles.push(r.Name);
              if (typeof r.RoleTypeKind === 'number') roleTypeKinds[r.Name] = r.RoleTypeKind;
            });
        }
      }
      debugLog('[SmartPermissions] getAadGroupSiteRoles: resolved via Graph=%o', roles);
      return { roles, roleTypeKinds, graphUnavailable: false };
    } catch (err) {
      debugLog('[SmartPermissions] getAadGroupSiteRoles: Graph failed', err);
      return { roles: [], roleTypeKinds: {}, graphUnavailable: isGraphPermissionError(err) };
    }
  }
