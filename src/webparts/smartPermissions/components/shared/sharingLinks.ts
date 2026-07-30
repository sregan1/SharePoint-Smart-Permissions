import { UserPermissionInfo } from '../../models/models';

// SharePoint auto-creates a hidden SharePointGroup for every "Anyone with the
// link" / "Specific people" sharing link, named "SharingLinks.<guid>.<Type>"
// (e.g. "SharingLinks.a1b2c3d4-....OrganizationEdit"). These are internal
// plumbing, not a person or a named group someone consciously granted access
// to, so they're offered as a filterable noise source alongside Limited
// Access — matched by loginName since Title/displayName use the same prefix.
//
// When "Expand group members" is on, the group row itself is kept AND its
// members are exploded into individual entries carrying the group's display
// name as `sourceGroup` (see permissionScan.ts) rather than the group's own
// loginName — so an exploded member's loginName is the *person's* UPN, not
// "SharingLinks.*". Checking sourceGroup too is what catches those rows.
export function isSharingLinkPrincipal(u: UserPermissionInfo): boolean {
  const loginMatch = u.loginName.toLowerCase().indexOf('sharinglinks.') !== -1;
  const sourceMatch = (u.sourceGroup ?? '').toLowerCase().indexOf('sharinglinks.') !== -1;
  return loginMatch || sourceMatch;
}
