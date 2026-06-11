import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  UserPermissionInfo, PermissionEntry, FolderFileNode, LibraryInfo,
  SiteCollectionInfo, SiteUserInfo, ReportOptions, ScanProgress,
} from '../models/models';
import { SpApiClient } from './sp/spCore';
import * as siteDiscovery from './sp/siteDiscovery';
import * as permissionScan from './sp/permissionScan';
import * as explorer from './sp/explorer';
import * as groups from './sp/groups';
import * as userAccess from './sp/userAccess';

export { isGraphPermissionError } from './sp/spCore';

// Facade over the sp/ modules so views keep a single dependency with a stable
// API. Implementation lives in:
//   sp/spCore.ts        — API client, throttling/paging, shared helpers
//   sp/siteDiscovery.ts — sites, lists/libraries, users, access checks
//   sp/permissionScan.ts— Permissions Report scan
//   sp/explorer.ts      — Permissions Explorer (real-time audit)
//   sp/groups.ts        — group membership expansion
//   sp/userAccess.ts    — User Access scan
export class SharePointService {
  private readonly client: SpApiClient;

  constructor(context: WebPartContext) {
    this.client = new SpApiClient(context);
  }

  /** Max concurrent API requests during scans. Settable from Settings. */
  get scanConcurrency(): number { return this.client.scanConcurrency; }
  set scanConcurrency(value: number) { this.client.scanConcurrency = value; }

  /** Max group members fetched before capping. Settable from Settings. */
  get groupMemberCap(): number { return this.client.groupMemberCap; }
  set groupMemberCap(value: number) { this.client.groupMemberCap = value; }

  runConcurrent<T>(
    tasks: (() => Promise<T | undefined>)[],
    concurrency = 5,
  ): Promise<(T | undefined)[]> {
    return this.client.runConcurrent(tasks, concurrency);
  }

  // ── Tenant / site discovery ───────────────────────────────────────────────

  getAllSites(tenantUrl: string, signal?: AbortSignal): Promise<SiteCollectionInfo[]> {
    return siteDiscovery.getAllSites(this.client, tenantUrl, signal);
  }

  getLibraries(siteUrl: string, signal?: AbortSignal, includeHidden = false): Promise<LibraryInfo[]> {
    return siteDiscovery.getLibraries(this.client, siteUrl, signal, includeHidden);
  }

  getSiteUsers(siteUrl: string, signal?: AbortSignal): Promise<SiteUserInfo[]> {
    return siteDiscovery.getSiteUsers(this.client, siteUrl, signal);
  }

  searchTenantUsers(siteUrl: string, query: string, signal?: AbortSignal): Promise<SiteUserInfo[]> {
    return siteDiscovery.searchTenantUsers(this.client, siteUrl, query, signal);
  }

  getSiteOwners(siteUrl: string, signal?: AbortSignal): Promise<{ title: string; email: string }[]> {
    return siteDiscovery.getSiteOwners(this.client, siteUrl, signal);
  }

  checkCanManagePermissions(siteUrl: string): Promise<boolean> {
    return siteDiscovery.checkCanManagePermissions(this.client, siteUrl);
  }

  // ── Permissions Report scan ───────────────────────────────────────────────

  scanPermissions(
    options: ReportOptions,
    onProgress: (progress: ScanProgress) => void,
    signal?: AbortSignal,
    onEntry?: (entry: PermissionEntry) => void,
  ): Promise<{ entries: PermissionEntry[]; groupPermissionDenied: boolean; roleAssignmentsDenied: boolean }> {
    return permissionScan.scanPermissions(this.client, options, onProgress, signal, onEntry);
  }

  // ── Real-time audit (Permissions Explorer) ────────────────────────────────

  getFolderContents(siteUrl: string, folderUrl: string, signal?: AbortSignal): Promise<FolderFileNode[]> {
    return explorer.getFolderContents(this.client, siteUrl, folderUrl, signal);
  }

  getItemPermissions(
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<{ hasUnique: boolean; users: UserPermissionInfo[]; permissionDenied?: boolean }> {
    return explorer.getItemPermissions(this.client, siteUrl, node, signal);
  }

  scanNodeForExternalUsers(
    siteUrl: string,
    node: FolderFileNode,
    signal?: AbortSignal,
  ): Promise<boolean | 'denied'> {
    return explorer.scanNodeForExternalUsers(this.client, siteUrl, node, signal);
  }

  getParentPermissions(
    siteUrl: string,
    serverRelativeUrl: string,
    signal?: AbortSignal,
  ): Promise<{ name: string; serverRelativeUrl: string; users: UserPermissionInfo[] } | null> {
    return explorer.getParentPermissions(this.client, siteUrl, serverRelativeUrl, signal);
  }

  getEffectivePermissions(siteUrl: string, node: FolderFileNode, signal?: AbortSignal): Promise<string> {
    return explorer.getEffectivePermissions(this.client, siteUrl, node, signal);
  }

  // ── Groups ────────────────────────────────────────────────────────────────

  getGroupMembers(
    siteUrl: string,
    groupName: string,
    loginName: string,
    principalType: string,
    signal?: AbortSignal,
  ): Promise<UserPermissionInfo[]> {
    return groups.getGroupMembers(this.client, siteUrl, groupName, loginName, principalType, signal);
  }

  // ── User Access scan ──────────────────────────────────────────────────────

  getUserAccess(
    siteUrl: string,
    userLoginName: string,
    onProgress: (msg: string) => void,
    signal?: AbortSignal,
    includeHidden = false,
  ): Promise<{ fullSiteAccess: boolean; items: PermissionEntry[]; graphPermissionRequired: boolean; roleAssignmentsDenied: boolean }> {
    return userAccess.getUserAccess(this.client, siteUrl, userLoginName, onProgress, signal, includeHidden);
  }
}
