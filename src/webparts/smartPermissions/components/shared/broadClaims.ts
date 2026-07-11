import { UserPermissionInfo } from '../../models/models';

// These two claims are the two ways a single grant hands access to
// (effectively) everyone in the tenant — "Everyone" includes anonymous/guest
// access, "Everyone except external users" is every authenticated member.
// They carry SecurityGroup-shaped role assignments but aren't real AAD
// groups, so they can't be expanded via Graph and otherwise render as an
// ordinary, easy-to-miss row. For a permissions audit, this is the single
// most important thing to surface prominently rather than blend in.
export function broadClaimLabel(u: UserPermissionInfo): string | null {
  const ln = u.loginName.toLowerCase();
  if (ln.startsWith('c:0(.s|true')) return 'Everyone';
  if (ln.startsWith('c:0-.f|rolemanager|spo-grid-all-users')) return 'Everyone except external users';
  return null;
}
