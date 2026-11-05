import { MSGraphClientV3 } from '@microsoft/sp-http';
import { UserPermissionInfo } from '../../models/models';
import {
  SpApiClient, odata, valueArray, principalTypeLabel, extractAadGroupInfo, isGraphPermissionError,
} from './spCore';

// ── Group membership expansion (SharePoint groups + AAD/M365 via Graph) ─────

export async function getGroupMembers(client: SpApiClient,
    siteUrl: string,
    groupName: string,
    loginName: string,
    principalType: string,
    signal?: AbortSignal,
    groupId?: number,
  ): Promise<UserPermissionInfo[]> {
    if (principalType === 'SharePointGroup') {
      try {
        // Prefer the numeric id — Member.Title/LoginName from a role assignment is the
        // group's display name, and getbyname(displayName) 404s on renamed or
        // non-canonically-named groups, silently dropping their members from results.
        const groupRef = groupId
          ? `getbyid(${groupId})`
          : `getbyname('${encodeURIComponent(odata(groupName))}')`;
        const data = await client.getJson(
          `${siteUrl}/_api/web/sitegroups/${groupRef}/users` +
            // Fetch one extra to detect truncation without a separate count call
            `?$select=LoginName,Title,IsHiddenInUI,PrincipalType&$top=${client.groupMemberCap + 1}`,
        );
        // Sorted before expanding (not after) so a nested group's members stay under it.
        const rawMembers: any[] = valueArray(data)
          .filter((u: any) => !u.IsHiddenInUI)
          .sort((a: any, b: any) => String(a.Title ?? '').localeCompare(String(b.Title ?? '')));
        const expanded: UserPermissionInfo[] = [];
        for (const u of rawMembers) {
          const pt: number = u.PrincipalType ?? 1;
          if (pt !== 1) {
            // Member is a nested group (M365 group or security group) — try to expand via Graph.
            // M365-connected SP groups commonly contain the backing M365 group as their only member.
            const nestedInfo = extractAadGroupInfo(u.LoginName ?? '');
            if (nestedInfo) {
              const nestedMembers = await getAadGroupMembers(client, nestedInfo.id, nestedInfo.isOwners ? 'owners' : 'members');
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
        if (expanded.length > client.groupMemberCap) {
          const capped = expanded.slice(0, client.groupMemberCap);
          capped.push({
            loginName: '',
            displayName: `(group has more than ${client.groupMemberCap} members — only first ${client.groupMemberCap} shown)`,
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
      const result = await getAadGroupMembers(client, groupInfo.id, groupInfo.isOwners ? 'owners' : 'members');
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

// Deepest level of nested groups that is expanded; deeper groups are listed unexpanded.
const MAX_NESTED_GROUP_DEPTH = 5;

async function getAadGroupMembers(client: SpApiClient, groupId: string, endpoint: 'members' | 'owners' = 'members'): Promise<UserPermissionInfo[] | null> {
    try {
      const graph: MSGraphClientV3 = await client.context.msGraphClientFactory.getClient('3');
      const cap = client.groupMemberCap;
      const out: UserPermissionInfo[] = [];
      // Groups already expanded in this call — guards against groups that contain each other.
      const visited = new Set<string>([groupId.toLowerCase()]);
      let truncated = false;

      // Lists one group's direct members. A member that is itself a group is added as a
      // group row, followed by its own members one level deeper.
      const collect = async (id: string, kind: 'members' | 'owners', depth: number): Promise<void> => {
        // Graph rejects $top > 999, but the Settings UI allows a cap up to 5000 — clamp
        // the request and instead rely on @odata.nextLink to detect true truncation.
        const result = await graph
          .api(`/groups/${id}/${kind}`)
          .select('displayName,userPrincipalName,mail,id')
          .top(Math.min(cap + 1, 999))
          .get();
        if (result?.['@odata.nextLink']) truncated = true;
        const members: any[] = (result?.value ?? []).slice().sort((a: any, b: any) =>
          String(a.displayName ?? '').localeCompare(String(b.displayName ?? '')),
        );
        for (const m of members) {
          if (out.length >= cap) { truncated = true; return; }
          const displayName = m.displayName ?? m.userPrincipalName ?? m.id ?? '';
          if (m['@odata.type'] === '#microsoft.graph.group') {
            out.push({
              loginName: m.id ?? '',
              displayName,
              principalType: 'SecurityGroup',
              roles: [],
              isGroupMember: true,
              depth,
            });
            const key = String(m.id ?? '').toLowerCase();
            if (key && !visited.has(key) && depth < MAX_NESTED_GROUP_DEPTH) {
              visited.add(key);
              try {
                await collect(m.id, 'members', depth + 1);
              } catch (err: any) {
                // Leave this one group unexpanded; a missing Graph permission still fails the whole call.
                if (isGraphPermissionError(err)) throw err;
              }
            }
          } else {
            out.push({
              loginName: m.userPrincipalName ?? m.mail ?? m.id ?? '',
              displayName,
              principalType: 'User',
              roles: [],
              isGroupMember: true,
              depth,
            });
          }
        }
      };

      await collect(groupId, endpoint, 0);
      if (truncated) {
        out.push({
          loginName: '',
          displayName: `(group has more than ${cap} members — only first ${cap} shown)`,
          principalType: 'User',
          roles: [],
          isGroupMember: true,
        });
      }
      return out;
    } catch (err: any) {
      if (isGraphPermissionError(err)) return null;
      return [];
    }
  }
