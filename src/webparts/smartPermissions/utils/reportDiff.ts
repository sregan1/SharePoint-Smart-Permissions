import { PermissionEntry, StoredReport, UserPermissionInfo } from '../models/models';

// Diff between two stored permission reports. Only DIRECT assignments are
// compared (sourceGroup === undefined): group-expanded members are derived
// data and would double-report every group membership change as a change on
// every object the group touches.

export interface PrincipalChange {
  displayName: string;
  loginName: string;
  /** Roles in the older report ('' when the principal is new). */
  oldRoles: string;
  /** Roles in the newer report ('' when the principal was removed). */
  newRoles: string;
}

export interface ObjectDiff {
  objectType: string;
  name: string;
  serverRelativeUrl: string;
  added: PrincipalChange[];
  removed: PrincipalChange[];
  changed: PrincipalChange[];
}

export interface ReportDiff {
  /** Objects present only in the newer report. */
  addedObjects: PermissionEntry[];
  /** Objects present only in the older report. */
  removedObjects: PermissionEntry[];
  /** Objects whose hasUniquePermissions flag flipped. */
  inheritanceChanged: { entry: PermissionEntry; nowUnique: boolean }[];
  /** Objects whose direct assignments differ. */
  permissionChanges: ObjectDiff[];
  /** True when there is nothing to report. */
  isEmpty: boolean;
}

function entryKey(e: PermissionEntry): string {
  return `${e.objectType}|${e.siteUrl}|${e.serverRelativeUrl}`;
}

// Map of direct principals on an entry: loginName (or displayName when the
// login is empty) → sorted, comma-joined role list.
function directAssignments(e: PermissionEntry): Map<string, { user: UserPermissionInfo; roles: string }> {
  const map = new Map<string, { user: UserPermissionInfo; roles: string }>();
  for (const u of e.uniquePermissions) {
    if (u.sourceGroup !== undefined || u.isGroupMember) continue;
    const key = (u.loginName || u.displayName).toLowerCase();
    if (!key) continue;
    map.set(key, { user: u, roles: [...u.roles].sort().join(', ') });
  }
  return map;
}

export function diffReports(older: StoredReport, newer: StoredReport): ReportDiff {
  const oldByKey = new Map(older.entries.map((e) => [entryKey(e), e]));
  const newByKey = new Map(newer.entries.map((e) => [entryKey(e), e]));

  const addedObjects: PermissionEntry[] = [];
  const removedObjects: PermissionEntry[] = [];
  const inheritanceChanged: { entry: PermissionEntry; nowUnique: boolean }[] = [];
  const permissionChanges: ObjectDiff[] = [];

  newByKey.forEach((newEntry, key) => {
    const oldEntry = oldByKey.get(key);
    if (!oldEntry) {
      addedObjects.push(newEntry);
      return;
    }

    if (oldEntry.hasUniquePermissions !== newEntry.hasUniquePermissions) {
      inheritanceChanged.push({ entry: newEntry, nowUnique: newEntry.hasUniquePermissions });
    }

    // Permission comparison is only meaningful where assignments are the
    // object's own. Inherited entries repeat their parent's list — comparing
    // those would duplicate every site-level change onto thousands of rows.
    if (!oldEntry.hasUniquePermissions || !newEntry.hasUniquePermissions) return;

    const oldPerms = directAssignments(oldEntry);
    const newPerms = directAssignments(newEntry);
    const added: PrincipalChange[] = [];
    const removed: PrincipalChange[] = [];
    const changed: PrincipalChange[] = [];

    newPerms.forEach((np, pKey) => {
      const op = oldPerms.get(pKey);
      if (!op) {
        added.push({ displayName: np.user.displayName, loginName: np.user.loginName, oldRoles: '', newRoles: np.roles });
      } else if (op.roles !== np.roles) {
        changed.push({ displayName: np.user.displayName, loginName: np.user.loginName, oldRoles: op.roles, newRoles: np.roles });
      }
    });
    oldPerms.forEach((op, pKey) => {
      if (!newPerms.has(pKey)) {
        removed.push({ displayName: op.user.displayName, loginName: op.user.loginName, oldRoles: op.roles, newRoles: '' });
      }
    });

    if (added.length || removed.length || changed.length) {
      permissionChanges.push({
        objectType: newEntry.objectType,
        name: newEntry.name,
        serverRelativeUrl: newEntry.serverRelativeUrl,
        added,
        removed,
        changed,
      });
    }
  });

  oldByKey.forEach((oldEntry, key) => {
    if (!newByKey.has(key)) removedObjects.push(oldEntry);
  });

  return {
    addedObjects,
    removedObjects,
    inheritanceChanged,
    permissionChanges,
    isEmpty:
      addedObjects.length === 0 &&
      removedObjects.length === 0 &&
      inheritanceChanged.length === 0 &&
      permissionChanges.length === 0,
  };
}
