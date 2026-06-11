import { SPHttpClient, ODataVersion } from '@microsoft/sp-http';
import { LibraryInfo, SiteCollectionInfo, SiteUserInfo } from '../../models/models';
import { SpApiClient, valueArray, isSystemLibrary } from './spCore';

// ── Tenant / site discovery, users, and access checks ───────────────────────

export async function getAllSites(client: SpApiClient, tenantUrl: string, signal?: AbortSignal): Promise<SiteCollectionInfo[]> {
    const sites: SiteCollectionInfo[] = [];
    const base = tenantUrl.replace(/\/$/, '');
    let startRow = 0;
    const rowLimit = 500;
    // The search REST API rejects the default OData v4 header — override to v3.
    const searchConfig = SPHttpClient.configurations.v1.overrideWith({
      defaultODataVersion: ODataVersion.v3,
    });

    while (!signal?.aborted) {
      const url =
        `${base}/_api/search/query?querytext='contentclass:STS_Site'` +
        `&rowlimit=${rowLimit}&startrow=${startRow}&selectproperties='Title,Path'`;

      const resp = await client.context.spHttpClient.get(url, searchConfig);
      if (!resp.ok) {
        const txt = await resp.text();
        throw new Error(`Site discovery failed: HTTP ${resp.status} — ${txt.substring(0, 300)}`);
      }
      const data = await resp.json();
      // Response shape depends on OData mode: verbose wraps everything in
      // d.query and arrays in { results: [] }; nometadata/minimal is top-level
      // with plain arrays. Handle both.
      const relevant =
        data?.PrimaryQueryResult?.RelevantResults ??
        data?.d?.query?.PrimaryQueryResult?.RelevantResults;
      const rows: any[] = relevant?.Table?.Rows?.results ?? relevant?.Table?.Rows ?? [];
      if (rows.length === 0) break;

      for (const row of rows) {
        let title: string | null = null;
        let path: string | null = null;
        for (const cell of (row.Cells?.results ?? row.Cells ?? []) as any[]) {
          if (cell.Key === 'Title') title = cell.Value;
          if (cell.Key === 'Path') path = cell.Value;
        }
        if (path) sites.push({ url: path, title: title ?? path });
      }

      const total: number = relevant?.TotalRows ?? 0;
      startRow += rows.length;
      if (startRow >= total || rows.length < rowLimit) break;
    }

    return sites;
  }

export async function getLibraries(client: SpApiClient, siteUrl: string, signal?: AbortSignal, includeHidden = false): Promise<LibraryInfo[]> {
    // Library-like templates only (the Explorer tree needs Files/Folders
    // semantics). IsSiteAssetsLibrary is filtered client-side because it is
    // not reliably filterable via OData across all SPO tenants.
    const baseFilter = '(BaseTemplate eq 101 or BaseTemplate eq 109 or BaseTemplate eq 119)';
    const filter = includeHidden ? baseFilter : `${baseFilter} and Hidden eq false`;
    const url =
      `${siteUrl}/_api/web/lists` +
      `?$filter=${encodeURIComponent(filter)}` +
      `&$select=Title,RootFolder/ServerRelativeUrl,NoCrawl,IsSiteAssetsLibrary` +
      `&$expand=RootFolder&$orderby=Title&$top=500`;
    const libs = await client.getJsonPaged(url, signal);
    return libs
      .filter((l: any) => includeHidden || !isSystemLibrary(l))
      .map((l: any) => ({
        title: l.Title,
        serverRelativeUrl: l.RootFolder?.ServerRelativeUrl ?? '',
        noCrawl: !!l.NoCrawl || undefined,
      }));
  }

export async function getSiteUsers(client: SpApiClient, siteUrl: string, signal?: AbortSignal): Promise<SiteUserInfo[]> {
    const url =
      `${siteUrl}/_api/web/siteusers` +
      `?$filter=IsHiddenInUI eq false and PrincipalType eq 1` +
      `&$select=LoginName,Title,Email&$orderby=Title&$top=2000`;
    const users = await client.getJsonPaged(url, signal);
    return users
      .filter(
        (u: any) =>
          !u.LoginName?.includes('_spo_') &&
          !u.LoginName?.includes('app@sharepoint'),
      )
      .map((u: any) => ({ loginName: u.LoginName, displayName: u.Title, email: u.Email || undefined }));
  }

  // Tenant-wide people search via the standard SharePoint people-picker
  // endpoint. Runs with the current user's permissions — no Graph scopes or
  // admin approval needed. Returns users who may not yet be in the site's
  // user information list.
export async function searchTenantUsers(client: SpApiClient, siteUrl: string, query: string, signal?: AbortSignal): Promise<SiteUserInfo[]> {
    if (!query.trim() || signal?.aborted) return [];
    const url =
      `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;
    const resp = await client.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify({
        queryParams: {
          QueryString: query,
          MaximumEntitySuggestions: 10,
          AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false,
          PrincipalType: 1,    // users only
          PrincipalSource: 15, // all sources (AAD, SharePoint, etc.)
        },
      }),
    });
    if (!resp.ok) return [];
    const data = await resp.json();
    // The endpoint returns its result as a JSON *string*: in 'value' with
    // nometadata, or d.ClientPeoplePickerSearchUser in verbose mode.
    const raw: string = data?.value ?? data?.d?.ClientPeoplePickerSearchUser ?? '[]';
    let parsed: any[] = [];
    try { parsed = JSON.parse(raw); } catch { return []; }
    return parsed
      .filter((p: any) => p.Key)
      .map((p: any) => ({
        loginName: p.Key as string,
        displayName: (p.DisplayText as string) || (p.Key as string),
        email: p.EntityData?.Email || undefined,
      }));
  }

export async function getSiteOwners(client: SpApiClient, 
    siteUrl: string,
    signal?: AbortSignal,
  ): Promise<{ title: string; email: string }[]> {
    try {
      if (signal?.aborted) return [];
      const data = await client.getJson(
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

export async function checkCanManagePermissions(client: SpApiClient, siteUrl: string): Promise<boolean> {
    try {
      const data = await client.getJson(`${siteUrl}/_api/web?$select=EffectiveBasePermissions`);
      // High/Low come back as strings. ManagePermissions (0x02000000) and
      // ManageWeb (0x40000000) both live in the Low 32 bits.
      const low = parseInt(data?.EffectiveBasePermissions?.Low ?? '0', 10) >>> 0;
      return !!(low & 0x02000000 || low & 0x40000000);
    } catch {
      return true; // fail open — don't block owners on API error
    }
  }
