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
  ): Promise<UserPermissionInfo[]> {
    if (principalType === 'SharePointGroup') {
      try {
        const data = await client.getJson(
          `${siteUrl}/_api/web/sitegroups/getbyname('${encodeURIComponent(odata(groupName))}')/users` +
            // Fetch one extra to detect truncation without a separate count call
            `?$select=LoginName,Title,IsHiddenInUI,PrincipalType&$top=${client.groupMemberCap + 1}`,
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
        expanded.sort((a, b) => a.displayName.localeCompare(b.displayName));
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

async function getAadGroupMembers(client: SpApiClient, groupId: string, endpoint: 'members' | 'owners' = 'members'): Promise<UserPermissionInfo[] | null> {
    try {
      const graph: MSGraphClientV3 = await client.context.msGraphClientFactory.getClient('3');
      const result = await graph
        .api(`/groups/${groupId}/${endpoint}`)
        .select('displayName,userPrincipalName,mail,id')
        .top(client.groupMemberCap + 1)
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
      if (all.length > client.groupMemberCap) {
        const capped = all.slice(0, client.groupMemberCap);
        capped.push({
          loginName: '',
          displayName: `(group has more than ${client.groupMemberCap} members — only first ${client.groupMemberCap} shown)`,
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
