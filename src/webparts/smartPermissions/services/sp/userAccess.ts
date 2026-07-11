import { MSGraphClientV3 } from '@microsoft/sp-http';
import { UserPermissionInfo, PermissionEntry, ObjectType } from '../../models/models';
import {
  SpApiClient, TaskQueue, odata, valueArray, rdbArray, isPermissionDenied, isGraphPermissionError,
  isSystemRole, isSystemLibrary, folderApi, fileApi, isLibraryTemplate, extractRoles,
} from './spCore';

// Verbose diagnostic logging for the group/role resolution logic below. This
// includes PII about the audited user (email, login, AAD object id) and group
// topology, which shouldn't sit in the browser console by default — opt in
// via localStorage.setItem('smartPermissionsDebug', '1') when troubleshooting.
function debugLog(...args: unknown[]): void {
  try {
    if (window.localStorage?.getItem('smartPermissionsDebug') === '1') {
      // eslint-disable-next-line no-console
      console.debug(...args);
    }
  } catch { /* localStorage unavailable (e.g. private browsing) */ }
}

// A single .top(999) page silently truncated for users in more than 999 groups,
// turning a genuine member into a false "not a member" result. Follow
// @odata.nextLink so the membership check is exhaustive.
async function fetchAllTransitiveGroupIds(graph: MSGraphClientV3, identifier: string): Promise<Set<string>> {
  const ids = new Set<string>();
  let resp = await graph.api(`/users/${encodeURIComponent(identifier)}/transitiveMemberOf`).select('id').top(999).get();
  for (;;) {
    (resp?.value ?? []).forEach((g: any) => ids.add((g.id as string).toLowerCase()));
    const nextLink: string | undefined = resp?.['@odata.nextLink'];
    if (!nextLink) break;
    // nextLink is an absolute URL; the Graph client's .api() expects a path
    // relative to its configured base, so strip the version prefix.
    resp = await graph.api(nextLink.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/, '')).get();
  }
  return ids;
}

// ── User Access scan ─────────────────────────────────────────────────────────

export async function getUserAccess(client: SpApiClient, 
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
      const userData = await client.getJson(
        `${siteUrl}/_api/web/siteusers/getbyloginname('${encodeURIComponent(odata(userLoginName))}')` +
          `?$expand=Groups&$select=Title,LoginName,Email,Groups/LoginName`,
      );
      userTitle = userData.Title ?? userLoginName;
      userEmail = userData.Email ?? '';
      const groups = valueArray(userData.Groups);
      // Normalize to lowercase — role-assignment member logins can differ in case
      // from user-groups logins depending on the SharePoint REST endpoint used.
      groupLogins = new Set(groups.map((g: any) => (g.LoginName as string).toLowerCase()));
      debugLog('[SmartPermissions] getUserAccess: user=%s email=%s SP-groups=%o',
        userLoginName, userEmail, Array.from(groupLogins));
    } catch (err) {
      debugLog('[SmartPermissions] getUserAccess: siteusers fetch failed', err);
    }

    // ── Site roles ──
    onProgress('Loading site permissions…');
    let siteRoles: string[] = [];
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
      siteRoles = extractRoles(allRa, userLoginName, groupLogins)
        .filter((r) => !isSystemRole(r));
      debugLog('[SmartPermissions] getUserAccess: siteRoles after extractRoles=%o', siteRoles);
    } catch (err) {
      debugLog('[SmartPermissions] getUserAccess: web RoleAssignments fetch failed', err);
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
          client.getJson(`${siteUrl}/_api/web?$select=Title,Url,ServerRelativeUrl`).catch(() => null),
          client.getJson(`${siteUrl}/_api/web/AssociatedOwnerGroup?$select=LoginName,Title`).catch(() => null),
          client.getJson(`${siteUrl}/_api/web/AssociatedMemberGroup?$select=LoginName,Title`).catch(() => null),
          client.getJson(`${siteUrl}/_api/web/AssociatedVisitorGroup?$select=LoginName,Title`).catch(() => null),
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
      debugLog('[SmartPermissions] getUserAccess: extractRoles returned empty — trying AAD group fallback');
      const aadResult = await getAadGroupSiteRoles(client, 
        siteUrl,
        userLoginName,
        userEmail,
        valueArray(webData.RoleAssignments),
      );
      siteRoles = aadResult.roles;
      graphPermissionRequired = aadResult.graphUnavailable;
      debugLog('[SmartPermissions] getUserAccess: siteRoles after AAD fallback=%o graphPermissionRequired=%s',
        siteRoles, graphPermissionRequired);
    }

    // ── Lists & libraries ──
    onProgress('Loading lists and libraries…');
    let libs: any[] = [];
    let libsHaveRoles = false;
    const libFilterClause = includeHidden
      ? ''
      : `$filter=${encodeURIComponent('Hidden eq false')}&`;

    try {
      libs = await client.getJsonPaged(
        `${siteUrl}/_api/web/lists` +
          `?${libFilterClause}` +
          `$select=Title,BaseTemplate,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
          `,RoleAssignments/Member/LoginName,RoleAssignments/Member/PrincipalType` +
          `,RoleAssignments/RoleDefinitionBindings/Name` +
          `&$expand=RootFolder,RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings` +
          `&$top=500`,
        signal,
      );
      libsHaveRoles = true;
    } catch {
      try {
        libs = await client.getJsonPaged(
          `${siteUrl}/_api/web/lists` +
            `?${libFilterClause}` +
            `$select=Title,BaseTemplate,HasUniqueRoleAssignments,NoCrawl,IsSiteAssetsLibrary,RootFolder/ServerRelativeUrl` +
            `&$expand=RootFolder&$top=500`,
          signal,
        );
      } catch { /* no lists */ }
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

    // One shared queue caps total concurrent requests across all libraries
    // and all recursion levels at scanConcurrency.
    const queue = new TaskQueue(client.scanConcurrency);

    for (const lib of libs) {
      queue.add(async () => {
        if (signal?.aborted) return;
        onProgress(`Scanning: ${lib.Title}`);
        const isLib = isLibraryTemplate(lib.BaseTemplate ?? 0);

        if (lib.HasUniqueRoleAssignments && libsHaveRoles) {
          const libRoles = extractRoles(
            valueArray(lib.RoleAssignments),
            userLoginName,
            groupLogins,
          ).filter((r) => !isSystemRole(r));

          if (libRoles.length > 0) {
            items.push({
              objectType: isLib ? ObjectType.Library : ObjectType.List,
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
              noCrawl: lib.NoCrawl ? true : undefined,
            });
          }
        }

        // Only library-like templates have Files/Folders to walk.
        if (isLib && lib.RootFolder?.ServerRelativeUrl) {
          await walkFoldersForUser(client, 
            siteUrl,
            lib.RootFolder.ServerRelativeUrl,
            userLoginName,
            userTitle,
            groupLogins,
            items,
            2,
            onProgress,
            queue,
            signal,
          );
        }
      });
    }
    await queue.drain();

    return { fullSiteAccess: false, items, graphPermissionRequired, roleAssignmentsDenied };
  }

async function walkFoldersForUser(client: SpApiClient, 
    siteUrl: string,
    folderUrl: string,
    userLogin: string,
    userDisplayName: string,
    groupLogins: Set<string>,
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
      try {
        subFolders = await client.getJsonPaged(
          `${apiBase}/Folders?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
          signal,
        );
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
        queue.add(async () => {
          if (signal?.aborted) return;
          const raData = await client.getJson(
            `${folderApi(siteUrl, subfolder.ServerRelativeUrl)}/ListItemAllFields/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name`,
          );
          const roles = extractRoles(valueArray(raData), userLogin, groupLogins).filter(
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
        });
      }

      for (const file of uniqueFiles) {
        queue.add(async () => {
          if (signal?.aborted) return;
          const raData = await client.getJson(
            `${fileApi(siteUrl, file.ServerRelativeUrl)}/ListItemAllFields/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name`,
          );
          const roles = extractRoles(valueArray(raData), userLogin, groupLogins).filter(
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
        });
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
          userLogin,
          userDisplayName,
          groupLogins,
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
  ): Promise<{ roles: string[]; graphUnavailable: boolean }> {
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
        return { roles: [], graphUnavailable: false };
      }

      const identifier = userEmail ||
        (userLoginName.includes('|') ? userLoginName.split('|').pop() : null) ||
        userLoginName;
      if (!identifier) return { roles: [], graphUnavailable: true };

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
        return { roles: [], graphUnavailable: false };
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
            .filter(`userPrincipalName eq '${identifier}'`)
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
        return { roles: [], graphUnavailable: true };
      }

      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: isOwner=%s', isOwner);

      const targetSpGroupId: number | undefined = isOwner
        ? webProps?.AssociatedOwnerGroup?.Id as number | undefined
        : webProps?.AssociatedMemberGroup?.Id as number | undefined;
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: mapped to SP group id=%s', targetSpGroupId);

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
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: resolved roles=%o', roles);
      return { roles, graphUnavailable: false };
    } catch (err) {
      debugLog('[SmartPermissions] checkM365ConnectedSiteRoles: failed', err);
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
async function getAadGroupSiteRoles(client: SpApiClient, 
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
        const data = await client.getJson(url);
        const found = valueArray(data).length > 0;
        debugLog('[SmartPermissions] getAadGroupSiteRoles: REST "%s" (id=%s pt=%s) → found=%s', label, groupId, pt, found);
        if (found) {
          rdbArray(ra.RoleDefinitionBindings)
            .map((r: any) => r.Name as string)
            .filter((r) => !isSystemRole(r))
            .forEach((r) => { if (roles.indexOf(r) === -1) roles.push(r); });
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
      return { roles, graphUnavailable: false };
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
    const identifier = userEmail ||
      (userLoginName.includes('|') ? userLoginName.split('|').pop() : null) ||
      userLoginName;
    debugLog('[SmartPermissions] getAadGroupSiteRoles: %d group(s) need Graph, identifier=%s',
      needsGraph.length, identifier);
    if (!identifier) return { roles: [], graphUnavailable: true };

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
            .map((r: any) => r.Name as string)
            .filter((r) => !isSystemRole(r))
            .forEach((r) => { if (roles.indexOf(r) === -1) roles.push(r); });
        }
      }
      debugLog('[SmartPermissions] getAadGroupSiteRoles: resolved via Graph=%o', roles);
      return { roles, graphUnavailable: false };
    } catch (err) {
      debugLog('[SmartPermissions] getAadGroupSiteRoles: Graph failed', err);
      return { roles: [], graphUnavailable: true };
    }
  }
