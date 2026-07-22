import { roleAccessTier } from '../../services/sp/spCore';

// Maps a set of SharePoint role names to a Fluent badge color reflecting the
// highest level of access: Full Control → danger, write-level → warning,
// read-level → success, anything else → informative. Classification prefers
// each role's language-invariant RoleTypeKind (via roleAccessTier) when
// known, falling back to the English name on localized tenants.
export function roleBadgeColor(
  roles: string[],
  roleTypeKinds?: Record<string, number>,
): 'brand' | 'danger' | 'warning' | 'success' | 'informative' {
  const tiers = roles.map((r) => roleAccessTier(r, roleTypeKinds?.[r]));
  if (tiers.indexOf('admin') !== -1) return 'danger';
  if (tiers.indexOf('edit') !== -1) return 'warning';
  if (tiers.indexOf('read') !== -1) return 'success';
  return 'informative';
}
