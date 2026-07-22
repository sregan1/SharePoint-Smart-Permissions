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
  /** Maps each entry in `roles` to its SharePoint RoleTypeKind, when known. */
  roleTypeKinds?: Record<string, number>;
  isGroupMember?: boolean;
  sourceGroup?: string; // group name if this user was expanded from a group; undefined = direct assignment
  /**
   * Numeric SharePoint group id (RoleAssignment.Member.Id), present when
   * principalType is 'SharePointGroup'. Preferred over displayName for group
   * expansion lookups — role-assignment Member.Title is the group's display
   * name, which can 404 against sitegroups/getbyname if the group was renamed
   * or contains characters getbyname mishandles; the id is unambiguous.
   */
  groupId?: number;
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
  /**
   * True when a unique-permissions check for this entry (or its children, for
   * a folder whose enriched listing failed) could not complete due to a
   * transient error — NOT a permission-denied response. The entry is reported
   * with its parent's permissions as a best-effort fallback, but that may be
   * inaccurate: it should be shown as "unknown/incomplete", never as a
   * confirmed inherited result.
   */
  scanIncomplete?: boolean;
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
  /**
   * Set when this node's children failed to load. Rendered as inline,
   * non-navigable text — previously represented as a synthetic child node
   * with an empty serverRelativeUrl, which collided on "" as a React key and
   * was reachable via arrow-key tree navigation like a real item.
   */
  loadError?: string;
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

