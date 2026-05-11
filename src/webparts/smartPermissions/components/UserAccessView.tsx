import * as React from 'react';
import {
  Button,
  Field,
  Badge,
  Text,
  Title3,
  Body1,
  Spinner,
  MessageBar,
  MessageBarBody,
  ProgressBar,
  Select,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ArrowLeft24Regular } from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { SiteUserInfo, PermissionEntry, ObjectType } from '../models/models';

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '900px',
    margin: '0 auto',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  scanArea: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    background: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: tokens.spacingVerticalM,
  },
  scanRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  accessTable: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: tokens.fontSizeBase200,
  },
  accessTh: {
    textAlign: 'left',
    padding: '6px 8px',
    borderBottom: `2px solid ${tokens.colorNeutralStroke1}`,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  accessTd: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'top',
  },
});

// ── Role badge colour ─────────────────────────────────────────────────────────

function roleBadgeColor(
  roles: string[],
): 'brand' | 'danger' | 'warning' | 'success' | 'informative' {
  if (roles.some((r) => r.toLowerCase().includes('full control'))) return 'danger';
  if (
    roles.some(
      (r) =>
        r.toLowerCase().includes('edit') ||
        r.toLowerCase().includes('contribute') ||
        r.toLowerCase().includes('design'),
    )
  )
    return 'warning';
  if (
    roles.some(
      (r) => r.toLowerCase().includes('read') || r.toLowerCase().includes('view'),
    )
  )
    return 'success';
  return 'informative';
}

function formatElapsed(seconds: number): string {
  const m = Math.floor(seconds / 60);
  const sec = seconds % 60;
  const s = sec < 10 ? `0${sec}` : String(sec);
  return `${m}:${s}`;
}

// ── Component ─────────────────────────────────────────────────────────────────

export interface UserAccessViewProps {
  sp: SharePointService;
  siteUrl: string;
  onBack: () => void;
}

export const UserAccessView: React.FC<UserAccessViewProps> = ({ sp, siteUrl, onBack }) => {
  const styles = useStyles();

  // ── Connection ──
  const [isConnecting, setIsConnecting] = React.useState(false);
  const [connectStatus, setConnectStatus] = React.useState('');
  const [connectError, setConnectError] = React.useState('');
  const [isConnected, setIsConnected] = React.useState(false);
  const [siteUsers, setSiteUsers] = React.useState<SiteUserInfo[]>([]);

  // ── User access ──
  const [selectedUser, setSelectedUser] = React.useState('');
  const [userAccessBusy, setUserAccessBusy] = React.useState(false);
  const [userAccessStatus, setUserAccessStatus] = React.useState('');
  const [userAccessItems, setUserAccessItems] = React.useState<PermissionEntry[]>([]);
  const [isFullSiteAccess, setIsFullSiteAccess] = React.useState(false);
  const [userAccessError, setUserAccessError] = React.useState('');

  // ── Scan timer ──
  const [scanElapsed, setScanElapsed] = React.useState(0);
  const timerRef = React.useRef<ReturnType<typeof setInterval> | null>(null);

  const abortRef = React.useRef<AbortController | null>(null);

  React.useEffect(() => {
    if (userAccessBusy) {
      setScanElapsed(0);
      timerRef.current = setInterval(() => setScanElapsed((s) => s + 1), 1000);
    }
    return () => {
      if (timerRef.current) {
        clearInterval(timerRef.current);
        timerRef.current = null;
      }
    };
  }, [userAccessBusy]);

  // ── Connect ──────────────────────────────────────────────────────────────

  const handleConnect = async (): Promise<void> => {
    abortRef.current?.abort();
    abortRef.current = new AbortController();

    setIsConnecting(true);
    setConnectError('');
    setConnectStatus('Loading users…');
    setIsConnected(false);
    setSiteUsers([]);
    setSelectedUser('');
    setUserAccessItems([]);
    setUserAccessStatus('');

    try {
      const users = await sp.getSiteUsers(siteUrl.trim(), abortRef.current.signal);
      setSiteUsers(users);
      setIsConnected(true);
      setConnectStatus(`Connected — ${users.length} user${users.length === 1 ? '' : 's'} found`);
    } catch (err: any) {
      setConnectError(`Connection failed: ${err?.message ?? String(err)}`);
      setConnectStatus('');
    } finally {
      setIsConnecting(false);
    }
  };

  // Auto-connect on mount
  React.useEffect(() => {
    handleConnect().catch((e) => console.error('[SmartPermissions] UserAccess handleConnect failed:', e));
  }, []);

  // ── User selection ───────────────────────────────────────────────────────

  const handleUserChange = async (e: React.ChangeEvent<HTMLSelectElement>): Promise<void> => {
    const login = e.target.value;
    setSelectedUser(login);
    if (!login) return;

    abortRef.current?.abort();
    abortRef.current = new AbortController();

    setUserAccessBusy(true);
    setIsFullSiteAccess(false);
    setUserAccessItems([]);
    setUserAccessError('');
    const user = siteUsers.find((u) => u.loginName === login);
    setUserAccessStatus(`Checking access for ${user?.displayName ?? login}…`);

    try {
      const { fullSiteAccess, items } = await sp.getUserAccess(
        siteUrl.trim(),
        login,
        (msg) => setUserAccessStatus(msg),
        abortRef.current.signal,
      );
      setIsFullSiteAccess(fullSiteAccess);
      setUserAccessItems(items);
      setUserAccessStatus(
        fullSiteAccess
          ? ''
          : items.length > 0
          ? `${items.length} accessible location(s) found.`
          : 'No accessible locations found.',
      );
    } catch (err: any) {
      setUserAccessError(err?.message ?? String(err));
    } finally {
      setUserAccessBusy(false);
    }
  };

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={onBack}
          disabled={isConnecting}
        >
          Back
        </Button>
        <Title3>User Access</Title3>
      </div>

      {connectError && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{connectError}</MessageBarBody>
        </MessageBar>
      )}
      {connectStatus && !connectError && (
        <Body1
          style={{
            color: isConnected
              ? tokens.colorStatusSuccessForeground1
              : tokens.colorNeutralForeground3,
            marginBottom: tokens.spacingVerticalM,
          }}
        >
          {connectStatus}
        </Body1>
      )}

      {isConnected && (
        <>
          <div style={{ marginBottom: tokens.spacingVerticalM }}>
            <Field label="Select a user">
              <Select
                value={selectedUser}
                onChange={handleUserChange}
                style={{ maxWidth: '400px' }}
                disabled={userAccessBusy}
              >
                <option value="">— pick a user —</option>
                {siteUsers.map((u) => (
                  <option key={u.loginName} value={u.loginName}>
                    {u.displayName}
                  </option>
                ))}
              </Select>
            </Field>
          </div>

          {userAccessBusy && (
            <div className={styles.scanArea}>
              <ProgressBar />
              <div className={styles.scanRow}>
                <Spinner size="tiny" />
                <Text>{userAccessStatus}</Text>
                <Text style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
                  {formatElapsed(scanElapsed)}
                </Text>
              </div>
              <Body1 style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
                This scan may take several minutes depending on the size of the site.
              </Body1>
            </div>
          )}

          {userAccessError && (
            <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>{userAccessError}</MessageBarBody>
            </MessageBar>
          )}

          {!userAccessBusy && isFullSiteAccess && (
            <MessageBar intent="success" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                This user has <strong>Full Control</strong> or Owner-level access to the
                entire site — all libraries and folders are accessible.
              </MessageBarBody>
            </MessageBar>
          )}

          {!userAccessBusy && userAccessStatus && !userAccessError && !isFullSiteAccess && (
            <Body1
              style={{
                color: tokens.colorNeutralForeground3,
                marginBottom: tokens.spacingVerticalM,
              }}
            >
              {userAccessStatus}
            </Body1>
          )}

          {!userAccessBusy && userAccessItems.length > 0 && (
            <table className={styles.accessTable}>
              <thead>
                <tr>
                  <th className={styles.accessTh}>Type</th>
                  <th className={styles.accessTh}>Name</th>
                  <th className={styles.accessTh}>Path</th>
                  <th className={styles.accessTh}>Permission Level</th>
                </tr>
              </thead>
              <tbody>
                {userAccessItems.map((item, i) => (
                  <tr key={i}>
                    <td className={styles.accessTd}>
                      <Badge
                        appearance="filled"
                        color={
                          item.objectType === ObjectType.Site
                            ? 'brand'
                            : item.objectType === ObjectType.Library
                            ? 'informative'
                            : item.objectType === ObjectType.Folder
                            ? 'warning'
                            : undefined
                        }
                        size="small"
                      >
                        {item.objectType}
                      </Badge>
                    </td>
                    <td
                      className={styles.accessTd}
                      style={{ paddingLeft: `${item.depth * 12}px` }}
                    >
                      {item.name}
                    </td>
                    <td className={styles.accessTd}>
                      <Text
                        style={{
                          fontSize: tokens.fontSizeBase100,
                          color: tokens.colorNeutralForeground3,
                        }}
                      >
                        {item.serverRelativeUrl}
                      </Text>
                    </td>
                    <td className={styles.accessTd}>
                      {item.uniquePermissions[0]?.roles.map((r, ri) => (
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
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </>
      )}
    </div>
  );
};
