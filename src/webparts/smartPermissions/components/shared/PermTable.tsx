import * as React from 'react';
import { Badge, Button, Text, makeStyles, tokens } from '@fluentui/react-components';
import { Person24Regular, People24Regular, PersonSearch16Regular } from '@fluentui/react-icons';

import { UserPermissionInfo } from '../../models/models';
import { roleBadgeColor } from './roleBadge';
import { isExternalUser, externalUserEmail } from './externalUsers';

const useStyles = makeStyles({
  permTable: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: tokens.fontSizeBase200,
  },
  permTh: {
    textAlign: 'left',
    padding: '6px 8px',
    borderBottom: `2px solid ${tokens.colorNeutralStroke1}`,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    position: 'sticky',
    top: 0,
    background: tokens.colorNeutralBackground1,
    zIndex: 1,
  },
  permTd: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'top',
  },
});

export interface PermTableProps {
  users: UserPermissionInfo[];
  onCheckAccess?: (loginName: string) => void;
}

// Table of permission assignments: principal, type, permission-level badges,
// and an optional "check access" action per user row.
export const PermTable: React.FC<PermTableProps> = ({ users, onCheckAccess }) => {
  const styles = useStyles();
  return (
    <table className={styles.permTable} aria-label="Permission assignments">
      <thead>
        <tr>
          <th className={styles.permTh}>User / Group</th>
          <th className={styles.permTh}>Type</th>
          <th className={styles.permTh}>Permission Level</th>
          {onCheckAccess && <th className={styles.permTh} />}
        </tr>
      </thead>
      <tbody>
        {users.map((u, i) => (
          <tr key={i}>
            <td className={styles.permTd}>
              {u.isGroupMember ? (
                <span style={{ paddingLeft: '16px', color: tokens.colorNeutralForeground3 }}>
                  ↳ {u.displayName}
                </span>
              ) : (
                <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  {u.principalType === 'User' ? (
                    <Person24Regular style={{ fontSize: '14px', flexShrink: 0 }} />
                  ) : (
                    <People24Regular style={{ fontSize: '14px', flexShrink: 0 }} />
                  )}
                  <span>
                    <span>{u.displayName || u.loginName}</span>
                    {isExternalUser(u) && (() => {
                      const email = externalUserEmail(u.loginName);
                      return email && email !== u.displayName ? (
                        <div style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 }}>
                          {email}
                        </div>
                      ) : null;
                    })()}
                  </span>
                </span>
              )}
            </td>
            <td className={styles.permTd}>
              <Text style={{ fontSize: tokens.fontSizeBase200 }}>
                {u.principalType === 'SecurityGroup'
                  ? 'Security Group'
                  : u.principalType === 'SharePointGroup'
                  ? 'SP Group'
                  : 'User'}
              </Text>
            </td>
            <td className={styles.permTd}>
              {u.roles.map((r, ri) => (
                <Badge
                  key={ri}
                  appearance="filled"
                  color={roleBadgeColor([r])}
                  size="small"
                  style={{ marginRight: '4px', marginBottom: '2px' }}
                >
                  {r}
                </Badge>
              ))}
            </td>
            {onCheckAccess && (
              <td className={styles.permTd}>
                {u.principalType === 'User' && u.loginName && !u.isGroupMember && (
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<PersonSearch16Regular />}
                    onClick={() => onCheckAccess(u.loginName)}
                    title={`Check access for ${u.displayName || u.loginName}`}
                    aria-label={`Check access for ${u.displayName || u.loginName}`}
                  />
                )}
              </td>
            )}
          </tr>
        ))}
      </tbody>
    </table>
  );
};
