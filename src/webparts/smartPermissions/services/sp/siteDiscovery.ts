import { SPHttpClient, ODataVersion, MSGraphClientV3 } from '@microsoft/sp-http';
import { LibraryInfo, SiteCollectionInfo, SiteUserInfo } from '../../models/models';
import {
  SpApiClient, valueArray, isSystemLibrary, isGraphPermissionError, isDebugEnabled,
} from './spCore';

// ── Tenant / site discovery, users, and access checks ───────────────────────

// Site discovery is tried in this order:
//  1. Microsoft Graph's tenant site directory (/sites/getAllSites) — authoritative,
//     independent of the SharePoint search index's crawl freshness or a tenant's
//     default result source. Requires the Sites.Read.All Graph permission.
//  2. SharePoint search (contentclass:STS_Site) — used only if Graph is
//     unavailable/not approved, since it can under-report sites depending on
//     crawl state and search configuration.
export async function getAllSites(client: SpApiClient, tenantUrl: string, signal?: AbortSignal): Promise<SiteCollectionInfo[]> {
    let graphNote: string;
    try {
      const graphSites = await getAllSitesViaGraph(client, tenantUrl, signal);
      if (graphSites.length > 0) {
        publishDiagnostics({
          tenantUrl,
          source: 'graph',
          pages: [],
          stoppedBecause: 'Graph returned the full site directory',
          finalCount: graphSites.length,
          discoveredUrls: graphSites.map((s) => s.url).sort((a, b) => a.localeCompare(b)),
          notes: [],
        });
        return graphSites;
      }
      graphNote = 'Graph site directory returned zero sites; used search instead.';
    } catch (err: any) {
      // Any Graph failure — permission not approved, transient error, endpoint
      // unavailable — falls back to search rather than failing the scan. Graph
      // is the better source when available but must never be a hard dependency,
      // since Sites.Read.All requires tenant-admin approval many orgs won't grant.
      graphNote = isGraphPermissionError(err)
        ? 'Graph site directory unavailable (Sites.Read.All not approved); used search instead.'
        : `Graph site discovery failed (${String(err?.message ?? err)}); used search instead.`;
    }
    return getAllSitesViaSearch(client, tenantUrl, signal, graphNote);
  }

// Enumerates every SharePoint site collection in the tenant via Microsoft
// Graph's tenant site directory. Unlike search-index discovery, this isn't
// affected by crawl freshness or a reassigned default search result source —
// it's a direct listing of registered site collections.
async function getAllSitesViaGraph(client: SpApiClient, tenantUrl: string, signal?: AbortSignal): Promise<SiteCollectionInfo[]> {
    const graph: MSGraphClientV3 = await client.context.msGraphClientFactory.getClient('3');
    const tenantHost = new URL(tenantUrl).hostname.toLowerCase();
    const sites: SiteCollectionInfo[] = [];
    const seen = new Set<string>();

    let nextUrl: string | null = `/sites/getAllSites?$select=webUrl,displayName,name&$top=200`;
    while (nextUrl && !signal?.aborted) {
      const result: any = await graph.api(nextUrl).version('v1.0').get();
      const rows: any[] = result?.value ?? [];
      for (const site of rows) {
        const webUrl: string | undefined = site?.webUrl;
        if (!webUrl) continue;
        let hostname: string;
        try {
          hostname = new URL(webUrl).hostname.toLowerCase();
        } catch {
          continue;
        }
        // Exclude OneDrive personal sites (…-my.sharepoint.com) and any site
        // collection on a different hostname (e.g. a different geo tenant),
        // matching what the search-based query would have scoped to.
        if (hostname !== tenantHost) continue;
        if (seen.has(webUrl)) continue;
        seen.add(webUrl);
        sites.push({ url: webUrl, title: site.displayName ?? site.name ?? webUrl });
      }
      nextUrl = result?.['@odata.nextLink'] ?? null;
    }

    return sites;
  }

// Well-known, built-in result source present in every SPO tenant by default.
// Pinning to it explicitly stops the query from silently inheriting whatever
// a tenant/site has reassigned as its *default* result source (a common
// enterprise-search customization) — without this, a query issued from one
// site can return a curated/scoped subset of site collections while the same
// querytext run through a web part configured with an explicit source (or
// from a site with a different default) returns the full set.
const LOCAL_SHAREPOINT_RESULTS_SOURCE_ID = '8413cd39-2156-4e00-b54d-11efd9430b0a';

const SITE_QUERY_TEXT = 'contentclass:STS_Site';

// The search REST API's GET form rejects the default OData v4 header.
const SEARCH_GET_CONFIG = SPHttpClient.configurations.v1.overrideWith({
  defaultODataVersion: ODataVersion.v3,
});

// Pulls the RelevantResults block out of a search response. Verbose OData wraps
// everything in d.query and arrays in { results: [] }; nometadata/minimal is
// top-level with plain arrays — handle both.
function readSearchPage(data: any): {
  rows: any[];
  totalRows: number | null;
  totalRowsIncludingDuplicates: number | null;
  queryModification: string | null;
} {
  const root = data?.d?.query ?? data;
  const relevant = root?.PrimaryQueryResult?.RelevantResults;
  return {
    rows: relevant?.Table?.Rows?.results ?? relevant?.Table?.Rows ?? [],
    totalRows: relevant?.TotalRows ?? null,
    totalRowsIncludingDuplicates: relevant?.TotalRowsIncludingDuplicates ?? null,
    // Non-null when SharePoint rewrote the submitted query (query rules,
    // result-source transforms). A rewritten query is a prime suspect when a
    // result set is unexpectedly small.
    queryModification: root?.SpellingSuggestion || relevant?.QueryModification || null,
  };
}

// Per-scan record of what site discovery actually asked for and got back, so a
// short result set can be diagnosed from a real response instead of guesswork.
// Always collected (site URLs only — no PII) and attached to
// window.__smartPermissionsSiteDiscovery for copying out of DevTools.
export interface SiteDiscoveryDiagnostics {
  tenantUrl: string;
  source: 'graph' | 'search';
  querytext?: string;
  sourceId?: string;
  pages: {
    page: number;
    method: string;
    startRow: number;
    rowLimit: number;
    rowsReturned: number;
    totalRows: number | null;
    totalRowsIncludingDuplicates: number | null;
    newUniqueSites: number;
    cumulative: number;
  }[];
  queryModification?: string | null;
  stoppedBecause: string;
  /** Results dropped for being on another host (OneDrive personal sites etc.). */
  excludedOtherHost?: number;
  finalCount: number;
  discoveredUrls: string[];
  notes: string[];
}

function publishDiagnostics(diag: SiteDiscoveryDiagnostics): void {
  try {
    (window as any).__smartPermissionsSiteDiscovery = diag;
  } catch { /* ignore */ }
  // eslint-disable-next-line no-console
  console.log(
    `[Smart Permissions] Site discovery via ${diag.source}: found ${diag.finalCount} site(s) ` +
    `in ${diag.pages.length} request(s). Stopped because: ${diag.stoppedBecause}. ` +
    `Full detail: copy(window.__smartPermissionsSiteDiscovery)`,
  );
  if (diag.pages.length > 0) {
    // eslint-disable-next-line no-console
    console.table(diag.pages);
  }
  diag.notes.forEach((n) => {
    // eslint-disable-next-line no-console
    console.warn(`[Smart Permissions] ${n}`);
  });
}

// One-off comparison of alternative site-discovery queries, run only when debug
// mode is on. When the primary query under-reports, this identifies which
// dimension is responsible — the querytext, the result source, or the client
// type — in a single pass rather than one deploy-and-retest cycle per guess.
async function probeSiteQueries(client: SpApiClient, base: string, signal?: AbortSignal): Promise<void> {
  const variants: { label: string; request: Record<string, unknown> }[] = [
    { label: 'A: STS_Site + LocalSharePointResults + ContentSearchRegular (current)',
      request: { Querytext: SITE_QUERY_TEXT, SourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID, ClientType: 'ContentSearchRegular' } },
    { label: 'B: STS_Site, no SourceId (tenant default source)',
      request: { Querytext: SITE_QUERY_TEXT, ClientType: 'ContentSearchRegular' } },
    { label: 'C: STS_Site, no ClientType',
      request: { Querytext: SITE_QUERY_TEXT, SourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID } },
    { label: 'D: STS_Site, no SourceId, no ClientType (barest form)',
      request: { Querytext: SITE_QUERY_TEXT } },
    { label: 'E: STS_Site OR STS_Web',
      request: { Querytext: 'contentclass:STS_Site OR contentclass:STS_Web', SourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID } },
    { label: 'F: SPSiteUrl:* (any indexed item, distinct site count)',
      request: { Querytext: 'SPSiteUrl:*', SourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID } },
    { label: 'G: STS_Site with TrimDuplicates on (default)',
      request: { Querytext: SITE_QUERY_TEXT, SourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID, TrimDuplicates: true } },
  ];

  const results: Record<string, unknown>[] = [];
  for (const v of variants) {
    if (signal?.aborted) return;
    try {
      const data = await client.postJson(
        `${base}/_api/search/postquery`,
        {
          request: {
            RowLimit: 500,
            StartRow: 0,
            SelectProperties: ['Path'],
            TrimDuplicates: false,
            ...v.request,
          },
        },
        0,
        signal,
      );
      const { rows, totalRows, totalRowsIncludingDuplicates, queryModification } = readSearchPage(data);
      const distinct = new Set(
        rows.map((r: any) =>
          ((r.Cells?.results ?? r.Cells ?? []) as any[]).find((c) => c.Key === 'Path')?.Value,
        ).filter(Boolean),
      );
      results.push({
        variant: v.label,
        rowsReturned: rows.length,
        distinctPaths: distinct.size,
        totalRows,
        totalRowsIncludingDuplicates,
        queryModification: queryModification ?? '',
      });
    } catch (err) {
      results.push({ variant: v.label, error: String(err).substring(0, 120) });
    }
  }

  try {
    (window as any).__smartPermissionsSiteQueryProbe = results;
  } catch { /* ignore */ }
  // eslint-disable-next-line no-console
  console.log('[Smart Permissions] Site-discovery query probe — copy(window.__smartPermissionsSiteQueryProbe)');
  // eslint-disable-next-line no-console
  console.table(results);
}

async function getAllSitesViaSearch(
    client: SpApiClient,
    tenantUrl: string,
    signal?: AbortSignal,
    graphNote?: string,
  ): Promise<SiteCollectionInfo[]> {
    const sites: SiteCollectionInfo[] = [];
    const seen = new Set<string>();
    const base = tenantUrl.replace(/\/$/, '');
    // contentclass:STS_Site also matches every OneDrive personal site, which
    // live on the "<tenant>-my.sharepoint.com" host. Those aren't sites an
    // admin is auditing here, reading their role assignments 403s for everyone
    // but their owner, and including them both inflates the site count and
    // fills the report with denied paths. Restricting to the tenant's own
    // hostname drops them, and matches what the Graph path already does.
    const tenantHost = new URL(base).hostname.toLowerCase();
    let excludedOtherHost = 0;
    const rowLimit = 500;
    // Safety valve against a server that ignores StartRow and keeps returning
    // page 1 forever. Generous enough for any realistic tenant.
    const maxPages = 60;
    let startRow = 0;
    let reportedTotal = 0;
    // POST /postquery is preferred (structured body, no URL-encoding quirks,
    // and it's what SharePoint's own content-search web parts use). If the
    // tenant rejects it outright, drop to the GET form for the remaining pages
    // rather than returning nothing.
    let usePost = true;

    const diag: SiteDiscoveryDiagnostics = {
      tenantUrl: base,
      source: 'search',
      querytext: SITE_QUERY_TEXT,
      sourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID,
      pages: [],
      stoppedBecause: 'reached maxPages safety cap',
      finalCount: 0,
      discoveredUrls: [],
      notes: graphNote ? [graphNote] : [],
    };

    for (let page = 0; page < maxPages && !signal?.aborted; page++) {
      let data: any;
      if (usePost) {
        try {
          data = await client.postJson(
            `${base}/_api/search/postquery`,
            {
              request: {
                Querytext: SITE_QUERY_TEXT,
                RowLimit: rowLimit,
                StartRow: startRow,
                SelectProperties: ['Title', 'Path'],
                // Without this, near-identical site entries get collapsed away.
                TrimDuplicates: false,
                SourceId: LOCAL_SHAREPOINT_RESULTS_SOURCE_ID,
                // Content-search client type: SharePoint applies different
                // query-side result limiting per client type, and this is the
                // one its own content rollup web parts identify as.
                ClientType: 'ContentSearchRegular',
              },
            },
            0,
            signal,
          );
        } catch (err) {
          diag.notes.push(`search/postquery failed on page ${page}; retrying via GET. ${String(err)}`);
          usePost = false;
        }
      }
      if (!usePost) {
        data = await client.getJson(
          `${base}/_api/search/query?querytext='${SITE_QUERY_TEXT}'` +
            `&rowlimit=${rowLimit}&startrow=${startRow}&selectproperties='Title,Path'` +
            `&trimduplicates=false&clienttype='ContentSearchRegular'` +
            `&sourceid='${LOCAL_SHAREPOINT_RESULTS_SOURCE_ID}'`,
          0,
          signal,
          SEARCH_GET_CONFIG,
        );
      }

      const { rows, totalRows, totalRowsIncludingDuplicates, queryModification } = readSearchPage(data);
      if (queryModification) diag.queryModification = queryModification;
      reportedTotal = Math.max(reportedTotal, totalRows ?? 0, totalRowsIncludingDuplicates ?? 0);

      const countBefore = sites.length;
      // Counts rows not seen on an earlier page, whether or not they survive the
      // host filter. Paging progress must be judged on this rather than on
      // `sites` — a page made up entirely of OneDrive sites is filtered down to
      // nothing but is still forward progress, and treating it as "no new
      // results" would stop the walk with pages still to go.
      let newRowsThisPage = 0;
      for (const row of rows) {
        let title: string | null = null;
        let path: string | null = null;
        for (const cell of (row.Cells?.results ?? row.Cells ?? []) as any[]) {
          if (cell.Key === 'Title') title = cell.Value;
          if (cell.Key === 'Path') path = cell.Value;
        }
        if (!path || seen.has(path)) continue;
        seen.add(path);
        newRowsThisPage++;
        let hostname: string;
        try {
          hostname = new URL(path).hostname.toLowerCase();
        } catch {
          continue;
        }
        if (hostname !== tenantHost) {
          excludedOtherHost++;
          continue;
        }
        sites.push({ url: path, title: title ?? path });
      }

      diag.pages.push({
        page,
        method: usePost ? 'POST postquery' : 'GET query',
        startRow,
        rowLimit,
        rowsReturned: rows.length,
        totalRows,
        totalRowsIncludingDuplicates,
        newUniqueSites: sites.length - countBefore,
        cumulative: sites.length,
      });

      if (rows.length === 0) {
        diag.stoppedBecause = 'response contained zero rows';
        break;
      }
      // Deliberately NOT stopping on `rows.length < rowLimit` or on
      // `startRow >= TotalRows`. SharePoint security-trims each page after the
      // index returns candidates, so a short page does not mean the result set
      // is exhausted, and TotalRows is a relevance estimate that routinely
      // under-reports. Either check terminates paging early and silently — the
      // cause of tenant-wide scans stopping at a fraction of the real site
      // count. Instead we page until a request yields no new site at all.
      if (newRowsThisPage === 0) {
        diag.stoppedBecause = `page ${page} returned ${rows.length} row(s), none of them new`;
        break;
      }
      startRow += rows.length;
    }

    const totalSeen = sites.length + excludedOtherHost;
    if (reportedTotal > totalSeen) {
      diag.notes.push(
        `Search reported up to ${reportedTotal} matching site(s) but only ${totalSeen} ` +
        `distinct site URL(s) were returned — some site collections may be missing from this scan.`,
      );
    }
    if (excludedOtherHost > 0) {
      diag.notes.push(
        `Excluded ${excludedOtherHost} result(s) not on ${tenantHost} — these are OneDrive ` +
        `personal sites and/or sites on another host, which aren't part of a tenant site audit.`,
      );
    }
    if (diag.queryModification) {
      diag.notes.push(
        `SharePoint rewrote the submitted query (${diag.queryModification}) — a query rule or ` +
        `result-source transform may be narrowing these results.`,
      );
    }

    diag.excludedOtherHost = excludedOtherHost;
    diag.finalCount = sites.length;
    diag.discoveredUrls = sites.map((s) => s.url).sort((a, b) => a.localeCompare(b));
    publishDiagnostics(diag);

    if (isDebugEnabled() && !signal?.aborted) {
      await probeSiteQueries(client, base, signal);
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
      // Fail closed: an API error here must not grant the "can manage
      // permissions" UI/behavior to someone we couldn't actually confirm has
      // it. A real owner sees the (harmless) Member-access affordances until
      // the transient error clears; that's a better failure mode than
      // showing privileged actions to a caller we couldn't verify.
      return false;
    }
  }
