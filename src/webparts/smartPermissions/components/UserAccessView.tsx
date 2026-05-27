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
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ArrowLeft24Regular, ArrowDownload24Regular, History24Regular, Delete24Regular } from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import { SiteUserInfo, PermissionEntry, ObjectType, StoredUserAccessReport } from '../models/models';

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

// ── Role badge color ──────────────────────────────────────────────────────────

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
  excel: ExcelExportService;
  siteUrl: string;
  includeHidden: boolean;
  prefillLogin?: string;
  onPrefillUsed?: () => void;
  onBack: () => void;
}

export const UserAccessView: React.FC<UserAccessViewProps> = ({ sp, excel, siteUrl, includeHidden, prefillLogin, onPrefillUsed, onBack }) => {
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
  const [userAccessBusy, setUserAccessBusy] = React.useState(false);
  const [userAccessStatus, setUserAccessStatus] = React.useState('');
  const [userAccessItems, setUserAccessItems] = React.useState<PermissionEntry[]>([]);
  const [isFullSiteAccess, setIsFullSiteAccess] = React.useState(false);
  const [userAccessError, setUserAccessError] = React.useState('');

  // ── Pagination ──
  const PAGE_SIZE = 200;
  const [visibleCount, setVisibleCount] = React.useState(PAGE_SIZE);

  // Reset pagination when results change
  React.useEffect(() => { setVisibleCount(PAGE_SIZE); }, [userAccessItems]);

  // ── Sort ──
  const [sortCol, setSortCol] = React.useState<'type' | 'name' | 'path' | 'permission'>('type');
  const [sortAsc, setSortAsc] = React.useState(true);

  const sortedAccessItems = React.useMemo(() => {
    return [...userAccessItems].sort((a, b) => {
      let va: string, vb: string;
      if (sortCol === 'type') { va = a.objectType; vb = b.objectType; }
      else if (sortCol === 'name') { va = a.name; vb = b.name; }
      else if (sortCol === 'path') { va = a.serverRelativeUrl; vb = b.serverRelativeUrl; }
      else { va = (a.uniquePermissions[0]?.roles ?? []).join(','); vb = (b.uniquePermissions[0]?.roles ?? []).join(','); }
      if (va < vb) return sortAsc ? -1 : 1;
      if (va > vb) return sortAsc ? 1 : -1;
      return 0;
    });
  }, [userAccessItems, sortCol, sortAsc]);

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
    try {
      await excel.exportUserAccess(item.entries, item.siteUrl, item.userDisplayName);
    } catch (err: any) {
      console.error('[SmartPermissions] history export failed:', err);
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

  const handleExport = async (): Promise<void> => {
    const user = siteUsers.find((u) => u.loginName === selectedUser);
    setIsExporting(true);
    try {
      await excel.exportUserAccess(sortedAccessItems, siteUrl.trim(), user?.displayName ?? selectedUser);
    } catch (err: any) {
      console.error('[SmartPermissions] exportUserAccess failed:', err);
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

  // Request notification permission on mount
  React.useEffect(() => {
    if ('Notification' in window && Notification.permission === 'default') {
      Notification.requestPermission().catch(() => { /* ignore */ });
    }
  }, []);

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

      // Auto-scan for prefill user (from Explorer cross-navigation)
      if (prefillLogin && !prefillHandledRef.current) {
        prefillHandledRef.current = true;
        onPrefillUsed?.();
        const match = users.find((u) => u.loginName === prefillLogin);
        setUserFilter(match?.displayName ?? prefillLogin);
        handleUserSelect(prefillLogin).catch((e) =>
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

  // ── User selection ───────────────────────────────────────────────────────

  const handleUserSelect = async (login: string): Promise<void> => {
    setSelectedUser(login);
    if (!login) return;

    abortRef.current?.abort();
    abortRef.current = new AbortController();
    scanStartRef.current = Date.now();

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
        includeHidden,
      );
      setIsFullSiteAccess(fullSiteAccess);
      setUserAccessItems(items);
      const statusMsg = fullSiteAccess
        ? 'Full site access detected.'
        : items.length > 0
        ? `${items.length} accessible location(s) found.`
        : 'No accessible locations found.';
      setUserAccessStatus(fullSiteAccess ? '' : statusMsg);

      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('Smart Permissions — User Access scan complete', {
          body: `${user?.displayName ?? login}: ${statusMsg}`,
        });
      }

      // Save to history (errors swallowed — never block the user)
      const storedReport: StoredUserAccessReport = {
        id: Date.now().toString(),
        timestamp: new Date().toISOString(),
        siteUrl: siteUrl.trim(),
        userLoginName: login,
        userDisplayName: user?.displayName ?? login,
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
            appearance="primary"
            onClick={() =>
              handleUserSelect(selectedUser).catch((e) =>
                console.error('[SmartPermissions] scan again failed:', e),
              )
            }
          >
            Scan again
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
                  handleUserSelect(d.optionValue ?? '').catch((err) =>
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
                    <Option key={u.loginName} value={u.loginName}>
                      {u.email ? `${u.displayName} (${u.email})` : u.displayName}
                    </Option>
                  ))}
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
                  This scan may take several minutes depending on the size of the site.
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
                      style={{ cursor: 'pointer', userSelect: 'none' }}
                    >
                      {label}{sortInd(col)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedAccessItems.slice(0, visibleCount).map((item, i) => (
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
