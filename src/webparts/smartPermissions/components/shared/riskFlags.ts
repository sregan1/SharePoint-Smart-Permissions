import { PermissionEntry, UserPermissionInfo } from '../../models/models';
import { isSharingLinkPrincipal } from './sharingLinks';

export type RiskKind = 'broad' | 'external' | 'fullControl' | 'incomplete';

// RoleTypeKind.Administrator — the "Full Control" permission level.
const ROLE_TYPE_ADMINISTRATOR = 5;

/** Everyone, Everyone except external users, all authenticated users, or an anonymous link. */
export function isBroadPrincipal(u: UserPermissionInfo): boolean {
  const login = u.loginName.toLowerCase();
  if (login.indexOf('spo-grid-all-users') !== -1) return true;     // Everyone except external users
  if (login.indexOf('c:0(.s|true') !== -1) return true;            // Everyone
  if (login.indexOf('c:0!.s|windows') !== -1) return true;         // All authenticated users
  if (login.indexOf('anonymous') !== -1) return true;              // Anyone-with-the-link
  return /^everyone\b/i.test(u.displayName);
}

/** Full Control granted to something other than a site Owners group. */
export function hasNonOwnerFullControl(u: UserPermissionInfo): boolean {
  if (isSharingLinkPrincipal(u)) return false;
  const full = u.roleTypeKinds
    ? Object.keys(u.roleTypeKinds).some((r) => u.roleTypeKinds![r] === ROLE_TYPE_ADMINISTRATOR)
    : u.roles.some((r) => r.toLowerCase() === 'full control');
  if (!full) return false;
  return !(u.principalType === 'SharePointGroup' && /owners/i.test(u.displayName));
}

export function entryHasRisk(e: PermissionEntry, kind: RiskKind): boolean {
  if (kind === 'incomplete') return !!e.scanIncomplete;
  if (kind === 'external') return e.uniquePermissions.some(isExternalPrincipal);
  return e.uniquePermissions.some(kind === 'broad' ? isBroadPrincipal : hasNonOwnerFullControl);
}

export function isExternalPrincipal(u: UserPermissionInfo): boolean {
  return u.loginName.toLowerCase().indexOf('#ext#') !== -1;
}

/** Absolute browser URL for an entry's server-relative path. */
export function entryAbsoluteUrl(e: PermissionEntry): string {
  try {
    return new URL(e.serverRelativeUrl, new URL(e.siteUrl).origin).toString();
  } catch {
    return e.serverRelativeUrl;
  }
}
