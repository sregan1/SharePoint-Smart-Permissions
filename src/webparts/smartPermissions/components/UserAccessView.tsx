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
  Combobox,
  Option,
  OptionGroup,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ArrowLeft24Regular, ArrowDownload24Regular, History24Regular, Delete24Regular } from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import { SiteUserInfo, PermissionEntry, ObjectType, StoredUserAccessReport } from '../models/models';
import { requestNotificationPermission, showNotification } from '../utils/notifications';
import { roleBadgeColor } from './shared/roleBadge';
import { SiteOwnersLinks } from './shared/SiteOwnersLinks';

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '900px',
    margin: '0 auto',
    minHeight: '500px',
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
    position: 'sticky',
    top: 0,
    background: tokens.colorNeutralBackground1,
    zIndex: 1,
  },
  accessTd: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'top',
  },
  historyTable: {
    width: '100%',
    borderCollapse: 'collapse' as const,
    fontSize: tokens.fontSizeBase200,
  },
  historyTh: {
    textAlign: 'left' as const,
    padding: '8px',
    borderBottom: `2px solid ${tokens.colorNeutralStroke1}`,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap' as const,
    position: 'sticky' as const,
    top: 0,
    background: tokens.colorNeutralBackground1,
    zIndex: 1,
  },
  historyTd: {
    padding: '8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'middle' as const,
  },
});

// ── System account filter ─────────────────────────────────────────────────────

function isSystemAccount(u: SiteUserInfo): boolean {
  const ln = u.loginName.toLowerCase();
  return (
    ln.startsWith('sharepoint\\') ||
    ln.startsWith('nt authority\\') ||
    ln.startsWith('c:0(.s|') ||
    ln.startsWith('c:0!.s|') ||
    ln.indexOf('spsearch') !== -1 ||
    ln.indexOf('spapp') !== -1 ||
    u.displayName === 'System Account'
  );
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
  excel: ExcelExportService;
  siteUrl: string;
  includeHidden: boolean;
  excludeLimitedAccess: boolean;
  prefillLogin?: string;
  onPrefillUsed?: () => void;
  onBack: () => void;
}

export const UserAccessView: React.FC<UserAccessViewProps> = ({ sp, excel, siteUrl, includeHidden, excludeLimitedAccess, prefillLogin, onPrefillUsed, onBack }) => {
  const styles = useStyles();

  // ── Connection ──
  const [isConnecting, setIsConnecting] = React.useState(false);
  const [connectStatus, setConnectStatus] = React.useState('');
  const [connectError, setConnectError] = React.useState('');
  const [isConnected, setIsConnected] = React.useState(false);
  const [siteUsers, setSiteUsers] = React.useState<SiteUserInfo[]>([]);

  // ── User access ──
  const [selectedUser, setSelectedUser] = React.useState('');
  const [userFilter, setUserFilter] = React.useState('');
  // Tenant-wide people-picker suggestions for users not in the site's list
  const [tenantSuggestions, setTenantSuggestions] = React.useState<SiteUserInfo[]>([]);
  const [userAccessBusy, setUserAccessBusy] = React.useState(false);
  const [userAccessStatus, setUserAccessStatus] = React.useState('');
  const [userAccessItems, setUserAccessItems] = React.useState<PermissionEntry[]>([]);
  const [isFullSiteAccess, setIsFullSiteAccess] = React.useState(false);
  const [userAccessError, setUserAccessError] = React.useState('');
  const [roleAssignmentsDenied, setRoleAssignmentsDenied] = React.useState(false);
  const [siteOwners, setSiteOwners] = React.useState<{ title: string; email: string }[]>([]);

  // ── Pagination ──
  const PAGE_SIZE = 200;
  const [visibleCount, setVisibleCount] = React.useState(PAGE_SIZE);

  // Reset pagination when results change
  React.useEffect(() => { setVisibleCount(PAGE_SIZE); }, [userAccessItems]);

  React.useEffect(() => {
    if (!roleAssignmentsDenied || !siteUrl) return;
    sp.getSiteOwners(siteUrl.trim()).then(setSiteOwners).catch(() => {});
  }, [roleAssignmentsDenied, siteUrl]);

  // ── Sort ──
  const [sortCol, setSortCol] = React.useState<'type' | 'name' | 'path' | 'permission'>('type');
  const [sortAsc, setSortAsc] = React.useState(true);

  const displayAccessItems = React.useMemo(() => {
    if (!excludeLimitedAccess) return userAccessItems;
    return userAccessItems.filter((e) => e.uniquePermissions.some((p) => p.roles.length > 0));
  }, [userAccessItems, excludeLimitedAccess]);

  // Derived from the filtered list (displayAccessItems), not the raw scan
  // results — otherwise the "site-level access" banner below could reference
  // a site row that excludeLimitedAccess has actually filtered out of view.
  const hasSiteEntry = displayAccessItems.some((i) => i.objectType === ObjectType.Site);

  const TYPE_ORDER: Record<string, number> = {
    [ObjectType.Site]: 0,
    [ObjectType.Library]: 1,
    [ObjectType.List]: 2,
    [ObjectType.Folder]: 3,
    [ObjectType.File]: 4,
  };

  const sortedAccessItems = React.useMemo(() => {
    return [...displayAccessItems].sort((a, b) => {
      let diff = 0;
      if (sortCol === 'type') {
        diff = (TYPE_ORDER[a.objectType] ?? 5) - (TYPE_ORDER[b.objectType] ?? 5);
      } else if (sortCol === 'name') {
        const va = a.name, vb = b.name;
        diff = va < vb ? -1 : va > vb ? 1 : 0;
      } else if (sortCol === 'path') {
        const va = a.serverRelativeUrl, vb = b.serverRelativeUrl;
        diff = va < vb ? -1 : va > vb ? 1 : 0;
      } else {
        const va = (a.uniquePermissions[0]?.roles ?? []).join(',');
        const vb = (b.uniquePermissions[0]?.roles ?? []).join(',');
        diff = va < vb ? -1 : va > vb ? 1 : 0;
      }
      if (diff !== 0) return sortAsc ? diff : -diff;
      // Secondary sort: Site → Library → List → Folder → File, then by path
      const typeDiff = (TYPE_ORDER[a.objectType] ?? 5) - (TYPE_ORDER[b.objectType] ?? 5);
      if (typeDiff !== 0) return typeDiff;
      return a.serverRelativeUrl.localeCompare(b.serverRelativeUrl);
    });
  }, [displayAccessItems, sortCol, sortAsc]);

  const handleSort = (col: typeof sortCol): void => {
    if (sortCol === col) { setSortAsc((v) => !v); } else { setSortCol(col); setSortAsc(true); }
  };

  const sortInd = (col: typeof sortCol): string =>
    sortCol !== col ? '' : sortAsc ? ' ▲' : ' ▼';

  // ── History ──
  const historyService = React.useRef(new ReportHistoryService());
  const [showHistory, setShowHistory] = React.useState(false);
  const [historyItems, setHistoryItems] = React.useState<StoredUserAccessReport[]>([]);
  const [exportingHistoryId, setExportingHistoryId] = React.useState<string | null>(null);
  const scanStartRef = React.useRef<number>(0);

  React.useEffect(() => {
    historyService.current.getAllUserAccess()
      .then(setHistoryItems)
      .catch(() => { /* IndexedDB unavailable */ });
  }, []);

  const handleHistoryExport = async (item: StoredUserAccessReport): Promise<void> => {
    setExportingHistoryId(item.id);
    setExportError('');
    try {
      await excel.exportUserAccess(item.entries, item.siteUrl, item.userDisplayName);
    } catch (err: any) {
      setExportError(`Export error: ${err?.message ?? String(err)}`);
    } finally {
      setExportingHistoryId(null);
    }
  };

  const handleHistoryDelete = async (id: string): Promise<void> => {
    await historyService.current.deleteUserAccess(id).catch(() => { /* ignore */ });
    setHistoryItems((prev) => prev.filter((r) => r.id !== id));
  };

  // ── Export ──
  const [isExporting, setIsExporting] = React.useState(false);
  const [graphPermissionRequired, setGraphPermissionRequired] = React.useState(false);
  const [exportError, setExportError] = React.useState('');

  const handleExport = async (): Promise<void> => {
    const user = siteUsers.find((u) => u.loginName === selectedUser);
    setIsExporting(true);
    setExportError('');
    try {
      await excel.exportUserAccess(sortedAccessItems, siteUrl.trim(), user?.displayName ?? selectedUser);
    } catch (err: any) {
      setExportError(`Export error: ${err?.message ?? String(err)}`);
    } finally {
      setIsExporting(false);
    }
  };

  // ── Scan timer ──
  const [scanElapsed, setScanElapsed] = React.useState(0);
  const timerRef = React.useRef<ReturnType<typeof setInterval> | null>(null);

  const abortRef = React.useRef<AbortController | null>(null);
  const prefillHandledRef = React.useRef(false);

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
      const rawUsers = await sp.getSiteUsers(siteUrl.trim(), abortRef.current.signal);
      const users = rawUsers.filter((u) => !isSystemAccount(u));
      setSiteUsers(users);
      setIsConnected(true);
      setConnectStatus(`Connected — ${users.length} user${users.length === 1 ? '' : 's'} found`);

      // Auto-scan for prefill user (from Explorer cross-navigation). The
      // display name is passed explicitly because the siteUsers state update
      // hasn't committed yet when handleUserSelect runs.
      if (prefillLogin && !prefillHandledRef.current) {
        prefillHandledRef.current = true;
        onPrefillUsed?.();
        const match = users.find((u) => u.loginName === prefillLogin);
        setUserFilter(match?.displayName ?? prefillLogin);
        handleUserSelect(prefillLogin, match?.displayName).catch((e) =>
          console.error('[SmartPermissions] prefill handleUserSelect failed:', e),
        );
      }
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

  // Debounced tenant-wide people search: users who have access via an AAD
  // group but never visited the site are missing from siteusers — offer them
  // as "Not in this site" suggestions once the filter has 3+ characters.
  React.useEffect(() => {
    const q = userFilter.trim();
    if (q.length < 3 || !isConnected || userAccessBusy) {
      setTenantSuggestions([]);
      return;
    }
    const timer = setTimeout(() => {
      sp.searchTenantUsers(siteUrl.trim(), q)
        .then((results) => {
          const existing = new Set(siteUsers.map((u) => u.loginName.toLowerCase()));
          setTenantSuggestions(results.filter((r) => !existing.has(r.loginName.toLowerCase())));
        })
        .catch(() => setTenantSuggestions([]));
    }, 300);
    return () => clearTimeout(timer);
  }, [userFilter, isConnected, userAccessBusy, siteUsers]);

  // ── User selection ───────────────────────────────────────────────────────

  const handleUserSelect = async (login: string, knownDisplayName?: string): Promise<void> => {
    setSelectedUser(login);
    if (!login) return;

    requestNotificationPermission();
    abortRef.current?.abort();
    abortRef.current = new AbortController();
    scanStartRef.current = Date.now();

    setUserAccessBusy(true);
    setIsFullSiteAccess(false);
    setUserAccessItems([]);
    setUserAccessError('');
    setGraphPermissionRequired(false);
    setRoleAssignmentsDenied(false);
    setSiteOwners([]);
    const displayName =
      knownDisplayName ?? siteUsers.find((u) => u.loginName === login)?.displayName ?? login;
    setUserAccessStatus(`Checking access for ${displayName}…`);

    try {
      const { fullSiteAccess, items, graphPermissionRequired: graphPerm, roleAssignmentsDenied: raDenied } = await sp.getUserAccess(
        siteUrl.trim(),
        login,
        (msg) => setUserAccessStatus(msg),
        abortRef.current.signal,
        includeHidden,
      );
      setIsFullSiteAccess(fullSiteAccess);
      setUserAccessItems(items);
      setGraphPermissionRequired(graphPerm);
      setRoleAssignmentsDenied(raDenied);
      const statusMsg = fullSiteAccess
        ? 'Full site access detected.'
        : items.length > 0
        ? `${items.length} accessible location(s) found.`
        : raDenied
        ? ''
        : 'No accessible locations found.';
      setUserAccessStatus(fullSiteAccess ? '' : statusMsg);

      showNotification(
        'Smart Permissions — User Access scan complete',
        `${displayName}: ${statusMsg}`,
      );

      // Save to history (errors swallowed — never block the user)
      const storedReport: StoredUserAccessReport = {
        id: Date.now().toString(),
        timestamp: new Date().toISOString(),
        siteUrl: siteUrl.trim(),
        userLoginName: login,
        userDisplayName: displayName,
        summary: {
          accessibleLocations: items.length,
          fullSiteAccess,
          durationSeconds: Math.round((Date.now() - scanStartRef.current) / 1000),
        },
        entries: items,
      };
      historyService.current.addUserAccess(storedReport)
        .then(() => historyService.current.getAllUserAccess())
        .then(setHistoryItems)
        .catch(() => { /* storage unavailable */ });
    } catch (err: any) {
      setUserAccessError(err?.message ?? String(err));
    } finally {
      setUserAccessBusy(false);
    }
  };

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className={styles.root}>
      {/* Screen reader announcements */}
      <div role="status" aria-live="polite" style={{ position: 'absolute', width: '1px', height: '1px', overflow: 'hidden', clip: 'rect(0,0,0,0)' }}>
        {!userAccessBusy && userAccessStatus ? userAccessStatus : ''}
      </div>

      {/* Header */}
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={() => {
            if (userAccessBusy && !window.confirm('A scan is in progress. Leave and cancel?')) return;
            onBack();
          }}
          disabled={isConnecting}
          aria-label="Back to home"
        >
          Back
        </Button>
        <Title3 style={{ flex: 1 }}>User Access</Title3>
        {selectedUser && !userAccessBusy && isConnected && (
          <Button
            appearance="secondary"
            onClick={() => {
              setSelectedUser('');
              setUserFilter('');
              setUserAccessItems([]);
              setUserAccessStatus('');
              setIsFullSiteAccess(false);
              setUserAccessError('');
              setGraphPermissionRequired(false);
              setVisibleCount(PAGE_SIZE);
            }}
          >
            New scan
          </Button>
        )}
        <Button
          appearance="subtle"
          icon={<History24Regular />}
          onClick={() => setShowHistory((v) => !v)}
          disabled={isConnecting}
        >
          History{historyItems.length > 0 ? ` (${historyItems.length})` : ''}
        </Button>
      </div>

      {/* ── History panel ── */}
      {showHistory && (
        <div>
          <div style={{ marginBottom: tokens.spacingVerticalM }}>
            <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={() => setShowHistory(false)}>
              Back to scan
            </Button>
          </div>
          {exportError && (
            <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>{exportError}</MessageBarBody>
            </MessageBar>
          )}
          {historyItems.length === 0 ? (
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No user access scans saved yet.</Body1>
          ) : (
            <table className={styles.historyTable}>
              <thead>
                <tr>
                  <th className={styles.historyTh}>Date / Time</th>
                  <th className={styles.historyTh}>User</th>
                  <th className={styles.historyTh}>Site</th>
                  <th className={styles.historyTh}>Locations</th>
                  <th className={styles.historyTh}></th>
                </tr>
              </thead>
              <tbody>
                {historyItems.map((item) => (
                  <tr key={item.id}>
                    <td className={styles.historyTd} style={{ whiteSpace: 'nowrap' }}>
                      {new Date(item.timestamp).toLocaleString()}
                    </td>
                    <td className={styles.historyTd}>
                      {item.userDisplayName}
                      {item.summary.fullSiteAccess && (
                        <Badge appearance="filled" color="danger" size="small" style={{ marginLeft: '6px' }}>Full Control</Badge>
                      )}
                    </td>
                    <td className={styles.historyTd}>
                      <Text style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>
                        {item.siteUrl}
                      </Text>
                    </td>
                    <td className={styles.historyTd}>{item.summary.accessibleLocations}</td>
                    <td className={styles.historyTd}>
                      <div style={{ display: 'flex', gap: tokens.spacingHorizontalS }}>
                        <Button
                          size="small"
                          appearance="primary"
                          icon={<ArrowDownload24Regular />}
                          onClick={() => handleHistoryExport(item)}
                          disabled={exportingHistoryId === item.id || item.entries.length === 0}
                        >
                          {exportingHistoryId === item.id ? 'Exporting…' : 'Export'}
                        </Button>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<Delete24Regular />}
                          onClick={() => handleHistoryDelete(item.id)}
                          title="Delete this record"
                        />
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}

      {!showHistory && connectError && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{connectError}</MessageBarBody>
        </MessageBar>
      )}
      {!showHistory && connectStatus && !connectError && (
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

      {!showHistory && isConnected && (
        <>
          <div style={{ marginBottom: tokens.spacingVerticalM }}>
            <Field label="Select a user">
              <Combobox
                placeholder="Search users…"
                value={userFilter}
                onInput={(e) => setUserFilter((e.target as HTMLInputElement).value)}
                onOptionSelect={(_, d) => {
                  setUserFilter(d.optionText ?? '');
                  handleUserSelect(d.optionValue ?? '', d.optionText).catch((err) =>
                    console.error('[SmartPermissions] handleUserSelect failed:', err),
                  );
                }}
                disabled={userAccessBusy}
                style={{ maxWidth: '400px' }}
                aria-label="Select a user to check access"
              >
                {siteUsers
                  .filter(
                    (u) =>
                      u.displayName.toLowerCase().includes(userFilter.toLowerCase()) ||
                      u.loginName.toLowerCase().includes(userFilter.toLowerCase()) ||
                      (u.email != null && u.email.toLowerCase().includes(userFilter.toLowerCase())),
                  )
                  .map((u) => (
                    <Option key={u.loginName} value={u.loginName} text={u.displayName}>
                      {u.email ? `${u.displayName} (${u.email})` : u.displayName}
                    </Option>
                  ))}
                {tenantSuggestions.length > 0 && (
                  <OptionGroup label="Not in this site">
                    {tenantSuggestions.map((u) => (
                      <Option key={u.loginName} value={u.loginName} text={u.displayName}>
                        {u.email ? `${u.displayName} (${u.email})` : u.displayName}
                      </Option>
                    ))}
                  </OptionGroup>
                )}
              </Combobox>
            </Field>
          </div>

          {userAccessBusy && (
            <div className={styles.scanArea} role="status" aria-label="Scan in progress">
              <ProgressBar aria-label="Scanning user access" />
              <div className={styles.scanRow}>
                <Spinner size="tiny" />
                <Text>{userAccessStatus}</Text>
                <Text style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
                  {formatElapsed(scanElapsed)}
                </Text>
              </div>
              <div className={styles.scanRow}>
                <Button appearance="secondary" size="small" onClick={() => abortRef.current?.abort()}>
                  Cancel
                </Button>
                <Body1 style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
                  Scanning…
                </Body1>
              </div>
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

          {!userAccessBusy && graphPermissionRequired && (
            <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                Site-level permission could not be determined — this site likely uses Microsoft 365 Group access.
                To show it, a SharePoint Administrator must approve the <strong>GroupMember.Read.All</strong> permission
                in <strong>SharePoint Admin Center → Advanced → API access</strong>.
              </MessageBarBody>
            </MessageBar>
          )}

          {!userAccessBusy && roleAssignmentsDenied && (
            <MessageBar
              intent={userAccessItems.length > 0 ? 'info' : 'warning'}
              style={{ marginBottom: tokens.spacingVerticalM }}
            >
              <MessageBarBody>
                {userAccessItems.length > 0 ? (
                  <>
                    Showing <strong>site-level access only</strong>, based on default SharePoint group
                    membership. Unique permissions on libraries, folders, and files could not be read
                    and are not shown. For complete results, run this scan as a{' '}
                    <strong>Site Owner</strong>.
                  </>
                ) : (
                  <>
                    This user{"'"}s access could not be determined. Reading permission assignments
                    requires the <strong>Manage Permissions</strong> right (Site Owner or higher),
                    and this user does not appear to be a member of the site{"'"}s default Owner,
                    Member, or Visitor groups.
                  </>
                )}
                <SiteOwnersLinks owners={siteOwners} />
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
            <div style={{ display: 'flex', gap: tokens.spacingHorizontalS, justifyContent: 'flex-end', marginBottom: tokens.spacingVerticalS }}>
              <Button
                appearance="secondary"
                icon={<ArrowDownload24Regular />}
                onClick={() => handleExport().catch((e) => console.error('[SmartPermissions] handleExport failed:', e))}
                disabled={isExporting}
              >
                {isExporting ? 'Exporting…' : 'Export to Excel'}
              </Button>
              <Button
                appearance="secondary"
                icon={<ArrowDownload24Regular />}
                onClick={() => {
                  const user = siteUsers.find((u) => u.loginName === selectedUser);
                  excel.exportUserAccessCsv(sortedAccessItems, siteUrl.trim(), user?.displayName ?? selectedUser);
                }}
                disabled={isExporting}
              >
                Export to CSV
              </Button>
            </div>
          )}
          {exportError && (
            <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalS }}>
              <MessageBarBody>{exportError}</MessageBarBody>
            </MessageBar>
          )}
          {!userAccessBusy && hasSiteEntry && !isFullSiteAccess && (
            <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                This user has site-level access. Only locations with{' '}
                <strong>unique permission assignments</strong> appear below — all other
                content is accessible through the site-level permission shown at the top
                of the list.
              </MessageBarBody>
            </MessageBar>
          )}
          {!userAccessBusy && userAccessItems.length > 0 && (
            <table className={styles.accessTable} aria-label="User access results">
              <thead>
                <tr>
                  {(
                    [
                      { col: 'type', label: 'Type' },
                      { col: 'name', label: 'Name' },
                      { col: 'path', label: 'Path' },
                      { col: 'permission', label: 'Permission Level' },
                    ] as { col: typeof sortCol; label: string }[]
                  ).map(({ col, label }) => (
                    <th
                      key={col}
                      className={styles.accessTh}
                      onClick={() => handleSort(col)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleSort(col); }
                      }}
                      tabIndex={0}
                      role="columnheader"
                      aria-sort={sortCol !== col ? 'none' : sortAsc ? 'ascending' : 'descending'}
                      style={{ cursor: 'pointer', userSelect: 'none' }}
                    >
                      {label}{sortInd(col)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedAccessItems.slice(0, visibleCount).map((item, i) => (
                  <tr key={`${item.serverRelativeUrl}|${item.objectType}|${i}`}>
                    <td className={styles.accessTd}>
                      <Badge
                        appearance="filled"
                        color={
                          item.objectType === ObjectType.Site
                            ? 'brand'
                            : item.objectType === ObjectType.Library
                            ? 'informative'
                            : item.objectType === ObjectType.List
                            ? 'success'
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
                      {item.uniquePermissions[0]?.roles.map((r) => (
                        <Badge
                          key={r}
                          appearance="filled"
                          color={roleBadgeColor([r], item.uniquePermissions[0]?.roleTypeKinds)}
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
          {!userAccessBusy && visibleCount < sortedAccessItems.length && (
            <div style={{ textAlign: 'center', marginTop: tokens.spacingVerticalM }}>
              <Button
                appearance="secondary"
                onClick={() => setVisibleCount((c) => c + PAGE_SIZE)}
              >
                Load more ({(sortedAccessItems.length - visibleCount).toLocaleString()} remaining)
              </Button>
            </div>
          )}
        </>
      )}
    </div>
  );
};
