import { UserPermissionInfo } from '../../models/models';
import { isExternalUser } from './externalUsers';

// Shared by the Explorer, Report, and User Access views (on-screen display and
// Excel/CSV export) so "exclude limited access" / "external users only"
// produce the same result everywhere they're applied, instead of each view
// reimplementing (and risking diverging from) the same filtering logic.
export function applyPermFilters(
  users: UserPermissionInfo[],
  excludeLimited: boolean,
  extOnly: boolean,
): UserPermissionInfo[] {
  let result = users;
  if (excludeLimited) result = result.filter((u) => u.roles.length > 0);
  if (extOnly) result = result.filter(isExternalUser);
  return result;
}
