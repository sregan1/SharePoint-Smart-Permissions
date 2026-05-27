import * as React from 'react';
import {
  Button,
  Badge,
  Text,
  Title3,
  Body1,
  Input,
  Label,
  ProgressBar,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  ArrowDownload24Regular,
  Person24Regular,
  PersonSearch24Regular,
  Search20Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { ExternalUserEntry } from '../models/models';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '1100px',
    margin: '0 auto',
    minHeight: '500px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  progressArea: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    background: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: tokens.spacingVerticalM,
  },
  filterRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalM,
    flexWrap: 'wrap',
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: tokens.fontSizeBase200,
  },
  th: {
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
  td: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'top',
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalXXL,
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
  },
});

export interface ExternalUsersViewProps {
  sp: SharePointService;
  siteUrl: string;
  onBack: () => void;
  onNavigateToUserAccess: (loginName: string) => void;
}

export const ExternalUsersView: React.FC<ExternalUsersViewProps> = ({ sp, siteUrl, onBack, onNavigateToUserAccess }) => {
  const styles = useStyles();

  const [isBusy, setIsBusy] = React.useState(false);
  const [progressMsg, setProgressMsg] = React.useState('');
  const [error, setError] = React.useState('');
  const [users, setUsers] = React.useState<ExternalUserEntry[] | null>(null);
  const [filter, setFilter] = React.useState('');
  const abortRef = React.useRef<AbortController | null>(null);

  const handleScan = async (): Promise<void> => {
    abortRef.current?.abort();
    abortRef.current = new AbortController();
    setIsBusy(true);
    setError('');
    setUsers(null);
    setProgressMsg('Loading external users…');
    try {
      const result = await sp.getExternalUsers(siteUrl.trim(), abortRef.current.signal);
      if (abortRef.current.signal.aborted) { setProgressMsg('Cancelled.'); return; }
      setUsers(result);
      setProgressMsg(`Found ${result.length} external user${result.length !== 1 ? 's' : ''}.`);
    } catch (err: any) {
      if (err?.name === 'AbortError') { setProgressMsg('Cancelled.'); return; }
      setError(`Error: ${err?.message ?? String(err)}`);
      setProgressMsg('');
    } finally {
      setIsBusy(false);
    }
  };

  const filteredUsers = React.useMemo((): ExternalUserEntry[] => {
    if (!users) return [];
    if (!filter.trim()) return users;
    const q = filter.toLowerCase();
    return users.filter(
      (u) =>
        u.displayName.toLowerCase().indexOf(q) !== -1 ||
        u.email.toLowerCase().indexOf(q) !== -1 ||
        u.loginName.toLowerCase().indexOf(q) !== -1,
    );
  }, [users, filter]);

  const handleExportCsv = (): void => {
    if (!users || users.length === 0) return;
    const rows = filter.trim() ? filteredUsers : users;
    const header = ['Display Name', 'Email', 'Login Name', 'Site Admin', 'Groups'];
    const csvEscape = (v: string): string => `"${v.replace(/"/g, '""')}"`;
    const csv = [
      header.map(csvEscape).join(','),
      ...rows.map((u) =>
        [u.displayName, u.email, u.loginName, u.isSiteAdmin ? 'Yes' : 'No', u.groups.join('; ')]
          .map(csvEscape)
          .join(','),
      ),
    ].join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `external-users-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  React.useEffect(() => { handleScan().catch(console.error); }, []);

  const hasResults = users !== null && users.length > 0;
  const firstScan = users === null && !isBusy && !error && !progressMsg;

  return (
    <div className={styles.root}>
      <div role="status" aria-live="polite" style={{ position: 'absolute', width: '1px', height: '1px', overflow: 'hidden', clip: 'rect(0,0,0,0)' }}>
        {!isBusy && progressMsg ? progressMsg : ''}
      </div>

      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={() => {
            if (isBusy && !window.confirm('A scan is in progress. Leave and cancel?')) return;
            onBack();
          }}
          aria-label="Back"
        >
          Back
        </Button>
        <Title3 style={{ flex: 1 }}>External Users</Title3>
        {hasResults && !isBusy && (
          <Button appearance="secondary" icon={<ArrowDownload24Regular />} onClick={handleExportCsv}>
            Export to CSV
          </Button>
        )}
        <Button appearance="primary" onClick={handleScan} disabled={isBusy}>
          {isBusy ? 'Scanning…' : firstScan ? 'Scan' : 'Scan again'}
        </Button>
      </div>

      <Body1 style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: tokens.spacingVerticalM }}>
        Lists all external accounts on this site — users whose login name contains <strong>#EXT#</strong>, indicating they
        are guests from outside the organization. No additional API permissions are required beyond read access to the site.
      </Body1>

      {(isBusy || progressMsg) && !error && (
        <div className={styles.progressArea}>
          {isBusy && <ProgressBar aria-label="Loading external users" />}
          <Body1>{progressMsg}</Body1>
          {isBusy && (
            <Button appearance="secondary" size="small" onClick={() => abortRef.current?.abort()}>
              Cancel
            </Button>
          )}
        </div>
      )}

      {error && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!isBusy && users !== null && users.length === 0 && (
        <div className={styles.emptyState}>
          <Person24Regular style={{ fontSize: '48px', opacity: 0.4 }} />
          <Text weight="semibold">No external users found</Text>
          <Body1>This site has no external (#EXT#) user accounts.</Body1>
        </div>
      )}

      {hasResults && (
        <>
          <div className={styles.filterRow}>
            <Label htmlFor="extUserFilter">Filter:</Label>
            <Input
              id="extUserFilter"
              contentBefore={<Search20Regular />}
              placeholder="Name, email, or login…"
              value={filter}
              onChange={(_, d) => setFilter(d.value)}
              style={{ minWidth: '260px' }}
            />
            <Body1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
              {filter.trim()
                ? `Showing ${filteredUsers.length} of ${users.length}`
                : `${users.length} external user${users.length !== 1 ? 's' : ''}`}
            </Body1>
          </div>

          <table className={styles.table} aria-label="External users">
            <thead>
              <tr>
                <th className={styles.th}>Display Name</th>
                <th className={styles.th}>Email</th>
                <th className={styles.th}>Site Admin</th>
                <th className={styles.th}>Groups</th>
                <th className={styles.th}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredUsers.map((u, i) => (
                <tr key={i}>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>{u.displayName || u.loginName}</Text>
                  </td>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>{u.email || '—'}</Text>
                  </td>
                  <td className={styles.td}>
                    {u.isSiteAdmin ? (
                      <Badge appearance="filled" color="danger" size="small">Admin</Badge>
                    ) : (
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>No</Text>
                    )}
                  </td>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>
                      {u.groups.length > 0 ? u.groups.join(', ') : '—'}
                    </Text>
                  </td>
                  <td className={styles.td}>
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<PersonSearch24Regular />}
                      onClick={() => onNavigateToUserAccess(u.loginName)}
                      aria-label={`Check access for ${u.displayName || u.loginName}`}
                    >
                      Check Access
                    </Button>
                  </td>
                </tr>
              ))}
              {filteredUsers.length === 0 && filter.trim() && (
                <tr>
                  <td className={styles.td} colSpan={5} style={{ textAlign: 'center', color: tokens.colorNeutralForeground3, padding: tokens.spacingVerticalL }}>
                    No users match the filter.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </>
      )}
    </div>
  );
};
