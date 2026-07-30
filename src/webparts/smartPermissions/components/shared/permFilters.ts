import { UserPermissionInfo } from '../../models/models';
import { isExternalUser } from './externalUsers';
import { isSharingLinkPrincipal } from './sharingLinks';

// Shared by the Explorer, Report, and User Access views (on-screen display and
// Excel/CSV export) so "exclude limited access" / "exclude sharing links" /
// "external users only" produce the same result everywhere they're applied,
// instead of each view reimplementing (and risking diverging from) the same
// filtering logic.
export function applyPermFilters(
  users: UserPermissionInfo[],
  excludeLimited: boolean,
  extOnly: boolean,
  excludeSharingLinks?: boolean,
): UserPermissionInfo[] {
  let result = users;
  if (excludeLimited) result = result.filter((u) => u.roles.length > 0);
  if (excludeSharingLinks) result = result.filter((u) => !isSharingLinkPrincipal(u));
  if (extOnly) result = result.filter(isExternalUser);
  return result;
}
