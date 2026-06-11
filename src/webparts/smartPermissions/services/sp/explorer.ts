import { FolderFileNode, UserPermissionInfo } from '../../models/models';
import {
  SpApiClient, valueArray, folderApi, fileApi, isPermissionDenied, spBitmaskToLevel,
  toPermissionInfoList,
} from './spCore';

// ── Real-time audit (Permissions Explorer) ──────────────────────────────────

export async function getFolderContents(client: SpApiClient, 
    siteUrl: string,
    folderUrl: string,
    signal?: AbortSignal,
  ): Promise<FolderFileNode[]> {
    const apiBase = folderApi(siteUrl, folderUrl);

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
      const [folders, filesArr] = await Promise.all([
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
      return toNodes(folders, filesArr, true);
    } catch {
      // Fallback: no HasUniqueRoleAssignments, fresh requests.
      const [folders, filesArr] = await Promise.all([
        client.getJsonPaged(
          `${apiBase}/Folders?$select=Name,ServerRelativeUrl,ItemCount&$top=2000`,
          signal,
        ),
        client.getJsonPaged(
          `${apiBase}/Files?$select=Name,ServerRelativeUrl&$top=2000`,
          signal,
        ),
      ]);
      return toNodes(folders, filesArr, false);
    }
  }

export async function getItemPermissions(client: SpApiClient, 
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<{ hasUnique: boolean; users: UserPermissionInfo[]; permissionDenied?: boolean }> {
    const base = node.isFolder
      ? `${folderApi(siteUrl, node.serverRelativeUrl)}/ListItemAllFields`
      : `${fileApi(siteUrl, node.serverRelativeUrl)}/ListItemAllFields`;

    try {
      const itemData = await client.getJson(`${base}?$select=HasUniqueRoleAssignments`);
      if (!itemData.HasUniqueRoleAssignments) return { hasUnique: false, users: [] };

      const raData = await client.getJson(
        `${base}/RoleAssignments` +
          `?$expand=Member,RoleDefinitionBindings` +
          `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
      );
      return { hasUnique: true, users: toPermissionInfoList(valueArray(raData)) };
    } catch (err) {
      return { hasUnique: false, users: [], permissionDenied: isPermissionDenied(err) };
    }
  }

  // Checks a single tree node for external users without re-fetching HasUniqueRoleAssignments
  // (already known from the tree). Returns false immediately for inherited nodes.
  // Returns 'denied' when the current user lacks ManagePermissions on this item.
export async function scanNodeForExternalUsers(client: SpApiClient, 
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<boolean | 'denied'> {
    if (!node.hasUniquePermissions || signal?.aborted) return false;
    const base = node.isFolder
      ? `${folderApi(siteUrl, node.serverRelativeUrl)}/ListItemAllFields`
      : `${fileApi(siteUrl, node.serverRelativeUrl)}/ListItemAllFields`;
    try {
      const raData = await client.getJson(
        `${base}/RoleAssignments?$expand=Member&$select=Member/LoginName`,
      );
      return valueArray(raData).some((ra: any) =>
        (ra.Member?.LoginName ?? '').toLowerCase().includes('#ext#'),
      );
    } catch (err) {
      return isPermissionDenied(err) ? 'denied' : false;
    }
  }

export async function getParentPermissions(client: SpApiClient, 
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
          const webData = await client.getJson(
            `${siteUrl}/_api/web` +
              `?$select=Title,ServerRelativeUrl` +
              `&$expand=RoleAssignments/Member,RoleAssignments/RoleDefinitionBindings`,
          );
          return {
            name: webData.Title ?? siteUrl,
            serverRelativeUrl: webData.ServerRelativeUrl ?? currentUrl,
            users: toPermissionInfoList(valueArray(webData.RoleAssignments)),
          };
        } catch {
          return null;
        }
      }

      try {
        const base = `${folderApi(siteUrl, currentUrl)}/ListItemAllFields`;
        const itemData = await client.getJson(`${base}?$select=HasUniqueRoleAssignments`);

        if (itemData.HasUniqueRoleAssignments) {
          const raData = await client.getJson(
            `${base}/RoleAssignments` +
              `?$expand=Member,RoleDefinitionBindings` +
              `&$select=Member/LoginName,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`,
          );
          return {
            name: currentUrl.substring(currentUrl.lastIndexOf('/') + 1),
            serverRelativeUrl: currentUrl,
            users: toPermissionInfoList(valueArray(raData)),
          };
        }
      } catch {
        // continue walking up
      }
    }

    return null;
  }

export async function getEffectivePermissions(client: SpApiClient, 
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<string> {
    const base = node.isFolder
      ? `${folderApi(siteUrl, node.serverRelativeUrl)}/ListItemAllFields`
      : `${fileApi(siteUrl, node.serverRelativeUrl)}/ListItemAllFields`;
    try {
      const data = await client.getJson(`${base}?$select=EffectiveBasePermissions`);
      if (signal?.aborted) return '';
      return spBitmaskToLevel(
        parseInt(data.EffectiveBasePermissions?.High ?? '0', 10),
        parseInt(data.EffectiveBasePermissions?.Low ?? '0', 10),
      );
    } catch {
      return '';
    }
  }
