export enum ObjectType {
  Site = 'Site',
  Library = 'Library',
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
}

export interface PermissionEntry {
  objectType: ObjectType;
  name: string;
  serverRelativeUrl: string;
  siteUrl: string;
  hasUniquePermissions: boolean;
  depth: number;
  uniquePermissions: UserPermissionInfo[];
}

export interface FolderFileNode {
  name: string;
  serverRelativeUrl: string;
  isFolder: boolean;
  hasChildren: boolean;
  hasUniquePermissions?: boolean;
  hasUniquePermissionsBelow?: boolean;
  isLoading?: boolean;
  parent?: FolderFileNode;
  children: FolderFileNode[];
}

export interface LibraryInfo {
  title: string;
  serverRelativeUrl: string;
}

export interface SiteCollectionInfo {
  url: string;
  title: string;
}

export interface SiteUserInfo {
  loginName: string;
  displayName: string;
}

export interface ReportOptions {
  siteUrl: string;
  allSites: boolean;
  scope: ReportScope;
  folderDepth: number;
  includeHidden: boolean;
}
