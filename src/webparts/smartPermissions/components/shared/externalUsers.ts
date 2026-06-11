import { UserPermissionInfo } from '../../models/models';

export function isExternalUser(u: UserPermissionInfo): boolean {
  return u.loginName.toLowerCase().indexOf('#ext#') !== -1;
}

// Extract the email from a SharePoint #EXT# login name.
// Format: [claims-prefix|]localpart_domain.tld#EXT#@tenant.onmicrosoft.com
// The last underscore before #EXT# is the @ in the original email address.
export function externalUserEmail(loginName: string): string {
  const extIdx = loginName.toLowerCase().indexOf('#ext#');
  if (extIdx === -1) return '';
  let local = loginName.substring(0, extIdx);
  const pipeIdx = local.lastIndexOf('|');
  if (pipeIdx >= 0) local = local.substring(pipeIdx + 1);
  const lastUnderscore = local.lastIndexOf('_');
  if (lastUnderscore === -1) return local;
  return `${local.substring(0, lastUnderscore)}@${local.substring(lastUnderscore + 1)}`;
}
