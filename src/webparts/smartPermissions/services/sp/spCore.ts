import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { UserPermissionInfo } from '../../models/models';

// Escape single-quotes in OData string literals (SQL-style doubling).
export function odata(s: string): string {
  return s.replace(/'/g, "''");
}

// API bases for a folder/file/list addressed by server-relative path. The
// *ByServerRelativePath(decodedUrl=...) forms plus URI-encoding handle names
// containing &, #, % and other characters that break the legacy
// GetFolderByServerRelativeUrl('...') form (& truncates the query string,
// # and % break URL parsing — items in such paths silently vanished).
export function folderApi(siteUrl: string, serverRelativeUrl: string): string {
  return `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeURIComponent(odata(serverRelativeUrl))}')`;
}
export function fileApi(siteUrl: string, serverRelativeUrl: string): string {
  return `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeURIComponent(odata(serverRelativeUrl))}')`;
}
export function listApi(siteUrl: string, serverRelativeUrl: string): string {
  return `${siteUrl}/_api/web/GetListUsingPath(decodedUrl='${encodeURIComponent(odata(serverRelativeUrl))}')`;
}

// Single shared work queue with a global concurrency cap. Unlike nested
// runConcurrent pools (which multiply: N workers each spawning N more per
// recursion level), tasks here can enqueue follow-up tasks — e.g. recursive
// folder walks — while total in-flight work stays capped at `concurrency`.
export class TaskQueue {
  private active = 0;
  private pending: (() => Promise<void>)[] = [];
  private idleResolvers: (() => void)[] = [];

  constructor(private readonly concurrency: number) {}

  add(task: () => Promise<void>): void {
    this.pending.push(task);
    this.pump();
  }

  private pump(): void {
    while (this.active < this.concurrency && this.pending.length > 0) {
      const task = this.pending.shift()!;
      this.active++;
      const onDone = (): void => {
        this.active--;
        this.pump();
        if (this.active === 0 && this.pending.length === 0) {
          this.idleResolvers.splice(0).forEach((resolve) => resolve());
        }
      };
      // Individual task errors don't stop the queue.
      task().then(onDone, onDone);
    }
  }

  drain(): Promise<void> {
    if (this.active === 0 && this.pending.length === 0) return Promise.resolve();
    return new Promise((resolve) => this.idleResolvers.push(resolve));
  }
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
export function principalTypeLabel(type: number): string {
  if (type === 4) return 'SecurityGroup';
  if (type === 8) return 'SharePointGroup';
  return 'User';
}

// Normalise role-definition-binding arrays: SPO REST returns a direct array
// with odata=nometadata; legacy verbose mode wraps it in { results: [] }.
export function rdbArray(bindings: any): any[] {
  if (Array.isArray(bindings)) return bindings;
  if (Array.isArray(bindings?.value)) return bindings.value;
  if (Array.isArray(bindings?.results)) return bindings.results;
  return [];
}

// Normalise top-level value arrays (same odata format issue).
export function valueArray(data: any): any[] {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.value)) return data.value;
  if (Array.isArray(data?.results)) return data.results;
  return [];
}

// Extract the Azure AD object GUID from a SharePoint claims login name.
// M365 Groups:        c:0o.c|federateddirectoryclaimprovider|{GUID}
// M365 Group owners:  c:0o.c|federateddirectoryclaimprovider|{GUID}_o
// Security Groups:    c:0t.c|tenant|{GUID}  /  c:0p.c|s2s|{GUID}
// The _o suffix indicates the owners slice of an M365 group — requires /groups/{id}/owners.
export const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export function extractAadGroupInfo(loginName: string): { id: string; isOwners: boolean } | null {
  let last = loginName.split('|').pop() ?? '';
  const isOwners = last.endsWith('_o');
  if (isOwners) last = last.slice(0, -2);
  return GUID_RE.test(last) ? { id: last, isOwners } : null;
}

// Returns true when an API error indicates the current user lacks read permission on
// role assignments (HTTP 403 Forbidden or 401 Unauthorized).
export function isPermissionDenied(err: any): boolean {
  const msg = String(err?.message ?? '');
  return msg.includes('HTTP 403') || msg.includes('HTTP 401');
}

// Translate an EffectiveBasePermissions {High, Low} bitmask to the highest matching
// SharePoint permission level name. SP REST returns High/Low as strings — callers
// must parseInt before passing in.
export function spBitmaskToLevel(high: number, low: number): string {
  void high; // High bits not needed for the levels we distinguish
  const lo = low >>> 0; // treat as unsigned 32-bit
  if (lo & 0x02000000 || lo & 0x40000000) return 'Full Control'; // ManagePermissions | ManageWeb
  if (lo & 0x00000800) return 'Design';                           // ManageLists
  if (lo & 0x00000004) return 'Edit';                             // EditListItems
  if (lo & 0x00000001) return 'Read';                             // ViewListItems
  return 'Limited access';
}

// SharePoint RoleTypeKind values for the two auto-assigned system roles that
// carry no meaningful permission of their own: 1 = Guest ("Limited Access"),
// 7 = System ("Web-Only Limited Access"). Language-invariant — unlike Name,
// these don't change on localized tenants (DE "Eingeschränkter Zugriff", FR
// "Accès limité", ...). Falls back to the name check when RoleTypeKind wasn't
// requested/returned.
const SYSTEM_ROLE_TYPE_KINDS = new Set([1, 7]);

// Returns true for system-assigned pass-through roles that carry no meaningful permission
// of their own and should be hidden from all results ("Limited Access" and its web-scoped
// variant are both auto-assigned by SharePoint and are never explicitly granted).
export function isSystemRole(role: { Name?: string; RoleTypeKind?: number } | string): boolean {
  const roleTypeKind = typeof role === 'string' ? undefined : role.RoleTypeKind;
  if (typeof roleTypeKind === 'number') return SYSTEM_ROLE_TYPE_KINDS.has(roleTypeKind);
  const name = typeof role === 'string' ? role : role.Name ?? '';
  const l = name.toLowerCase();
  return l === 'limited access' || l === 'web-only limited access' || l.startsWith('system.');
}

// Classifies a role by access tier using its language-invariant RoleTypeKind
// when known (5 = Administrator/Full Control, 3/4/6 = Contributor/
// WebDesigner/Editor "edit-level", 2 = Reader "read-level"), falling back to
// the English name for custom permission levels or when RoleTypeKind wasn't
// fetched — same defensive pattern as isSystemRole.
export function roleAccessTier(name: string, roleTypeKind?: number): 'admin' | 'edit' | 'read' | 'other' {
  switch (roleTypeKind) {
    case 5: return 'admin';
    case 3: case 4: case 6: return 'edit';
    case 2: return 'read';
  }
  const l = name.toLowerCase();
  if (l.includes('full control')) return 'admin';
  if (l.includes('edit') || l.includes('contribute') || l.includes('design')) return 'edit';
  if (l.includes('read') || l.includes('view')) return 'read';
  return 'other';
}

// Known system/infrastructure library URL suffixes (lowercased, site-relative).
// Checked as a suffix so they match regardless of site path prefix.
export const SYSTEM_LIB_SUFFIXES = [
  '/formservertemplates', // Form Templates
  '/style library',       // Style Library
];

// Returns true if this list entry should be treated as a system/hidden library
// and excluded when includeHidden is false. NoCrawl is deliberately NOT
// treated as system: admins sometimes mark sensitive libraries NoCrawl to
// hide them from search, which is exactly what an auditor needs to see —
// those are included and flagged via PermissionEntry.noCrawl instead.
export function isSystemLibrary(lib: any): boolean {
  if (lib.IsSiteAssetsLibrary) return true;
  const url = ((lib.RootFolder?.ServerRelativeUrl) ?? '').toLowerCase();
  return SYSTEM_LIB_SUFFIXES.some((s) => url.endsWith(s));
}

// Base templates that behave like document libraries (have Files/Folders and
// can be walked): 101 = Document Library, 109 = Picture Library, 119 = Site
// Pages. Everything else is a generic list, reported at list level only.
export const LIBRARY_TEMPLATES = [101, 109, 119];
export function isLibraryTemplate(baseTemplate: number): boolean {
  return LIBRARY_TEMPLATES.indexOf(baseTemplate) !== -1;
}

// Shared API client: SPFx context plus the throttling-aware fetch helpers and
// user-tunable scan settings. All sp/ modules take this as their first argument.
export class SpApiClient {
  public readonly context: WebPartContext;
  /** Max concurrent API requests during scans. Settable from Settings. */
  public scanConcurrency = 4;
  /** Max group members fetched before capping. Settable from Settings. */
  public groupMemberCap = 500;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Retries on 429/503 using the Retry-After header, and on thrown/rejected
  // errors (network blips) using capped exponential backoff with jitter —
  // both share the same 3-attempt cap. A rejected fetch previously got no
  // retry at all, which fed spurious "permission denied"/"inherited" results
  // into callers that treat any thrown error as a permission failure.
  public async getJson(url: string, attempt = 0, signal?: AbortSignal): Promise<any> {
    let resp: SPHttpClientResponse;
    try {
      resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    } catch (err) {
      if (signal?.aborted || attempt >= 3) throw err;
      const backoff = Math.min(1000 * 2 ** attempt, 8000) + Math.random() * 500;
      await new Promise((r) => setTimeout(r, backoff));
      return this.getJson(url, attempt + 1, signal);
    }
    if ((resp.status === 429 || resp.status === 503) && attempt < 3) {
      const retryAfter = parseInt(resp.headers.get('Retry-After') ?? '10', 10);
      await new Promise((r) => setTimeout(r, (isNaN(retryAfter) ? 10 : retryAfter) * 1000));
      return this.getJson(url, attempt + 1, signal);
    }
    if (!resp.ok) {
      const txt = await resp.text();
      throw new Error(`HTTP ${resp.status} — ${txt.substring(0, 300)}`);
    }
    return resp.json();
  }

  // Fetches a collection endpoint and follows server-side paging links so
  // results beyond the $top page size are not silently dropped. maxPages is a
  // safety valve against runaway loops on enormous collections.
  public async getJsonPaged(url: string, signal?: AbortSignal, maxPages = 50): Promise<any[]> {
    const all: any[] = [];
    let next: string | undefined = url;
    for (let page = 0; next && page < maxPages && !signal?.aborted; page++) {
      const data = await this.getJson(next, 0, signal);
      all.push(...valueArray(data));
      next = data?.['odata.nextLink'] ?? data?.['@odata.nextLink'] ?? data?.d?.__next;
    }
    return all;
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

}

export function toPermissionInfoList(roleAssignments: any[]): UserPermissionInfo[] {
    const result: UserPermissionInfo[] = [];
    for (const ra of roleAssignments) {
      const bindings = rdbArray(ra.RoleDefinitionBindings).filter((r: any) => !isSystemRole(r));
      const roles = bindings.map((r: any) => r.Name as string);
      const roleTypeKinds: Record<string, number> = {};
      for (const b of bindings) {
        if (typeof b.RoleTypeKind === 'number') roleTypeKinds[b.Name] = b.RoleTypeKind;
      }
      if (roles.length > 0) {
        const principalType = principalTypeLabel(ra.Member?.PrincipalType ?? 1);
        result.push({
          loginName: ra.Member?.LoginName ?? '',
          displayName: ra.Member?.Title ?? '',
          principalType,
          roles,
          roleTypeKinds: Object.keys(roleTypeKinds).length > 0 ? roleTypeKinds : undefined,
          groupId: principalType === 'SharePointGroup' ? (ra.Member?.Id as number | undefined) : undefined,
        });
      }
    }
    return result;
  }

export function extractRoles(
    roleAssignments: any[],
    userLogin: string,
    groupLogins: Set<string>,
  ): { roles: string[]; roleTypeKinds: Record<string, number> } {
    const roles = new Set<string>();
    const roleTypeKinds: Record<string, number> = {};
    for (const ra of roleAssignments) {
      const memberLogin: string = ra.Member?.LoginName ?? '';
      const match =
        memberLogin.toLowerCase() === userLogin.toLowerCase() ||
        groupLogins.has(memberLogin.toLowerCase());
      if (match) {
        for (const rb of rdbArray(ra.RoleDefinitionBindings)) {
          if (!isSystemRole(rb)) {
            roles.add(rb.Name);
            if (typeof rb.RoleTypeKind === 'number') roleTypeKinds[rb.Name] = rb.RoleTypeKind;
          }
        }
      }
    }
    return { roles: Array.from(roles), roleTypeKinds };
  }
