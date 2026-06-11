export enum ObjectType {
  Site = 'Site',
  Library = 'Library',
  List = 'List',
  Folder = 'Folder',
  File = 'File',
}

export enum ReportScope {
  Site = 'Site',
  Library = 'Library',
  Folder = 'Folder',
  Item = 'Item',
}

export interface UserPermissionInfo {
  loginName: string;
  displayName: string;
  principalType: string; // 'User', 'SecurityGroup', 'SharePointGroup'
  roles: string[];
  isGroupMember?: boolean;
  sourceGroup?: string; // group name if this user was expanded from a group; undefined = direct assignment
}

export interface PermissionEntry {
  objectType: ObjectType;
  name: string;
  serverRelativeUrl: string;
  siteUrl: string;
  hasUniquePermissions: boolean;
  depth: number;
  uniquePermissions: UserPermissionInfo[];
  /** True when the list/library is marked NoCrawl (hidden from search). */
  noCrawl?: boolean;
}

export interface FolderFileNode {
  name: string;
  serverRelativeUrl: string;
  isFolder: boolean;
  hasChildren: boolean;
  hasUniquePermissions?: boolean;
  hasUniquePermissionsBelow?: boolean;
  hasExternalUsers?: boolean;
  hasExternalUsersBelow?: boolean;
  isLoading?: boolean;
  parent?: FolderFileNode;
  children: FolderFileNode[];
}

export interface LibraryInfo {
  title: string;
  serverRelativeUrl: string;
  /** True when the library is marked NoCrawl (hidden from search). */
  noCrawl?: boolean;
}

export interface SiteCollectionInfo {
  url: string;
  title: string;
}

export interface SiteUserInfo {
  loginName: string;
  displayName: string;
  email?: string;
}

export interface ReportOptions {
  siteUrl: string;
  allSites: boolean;
  /** Recurse into subsites (subwebs) of each scanned site. */
  includeSubsites: boolean;
  scope: ReportScope;
  folderDepth: number;
  includeHidden: boolean;
  expandGroups: boolean;
  /** If set, only scan libraries whose serverRelativeUrl appears in this list. */
  libraryUrls?: string[];
}

export interface ScanProgress {
  message: string;
  scanned: number;    // total items added to results so far
  libsDone: number;   // libraries fully processed
  libsTotal: number;  // total libraries to scan (known after initial fetch)
}

export interface StoredReport {
  id: string;           // Date.now().toString() — unique and chronologically sortable
  timestamp: string;    // ISO 8601 string for display
  siteUrl: string;
  options: {
    allSites: boolean;
    /** Optional for backward compatibility with reports stored before v1.4. */
    includeSubsites?: boolean;
    scope: ReportScope;
    folderDepth: number;
    expandGroups: boolean;
  };
  summary: {
    totalObjects: number;
    uniqueCount: number;
    inheritedCount: number;
    durationSeconds: number;
  };
  entries: PermissionEntry[];
}

export interface StoredUserAccessReport {
  id: string;
  timestamp: string;
  siteUrl: string;
  userLoginName: string;
  userDisplayName: string;
  summary: {
    accessibleLocations: number;
    fullSiteAccess: boolean;
    durationSeconds: number;
  };
  entries: PermissionEntry[];
}

