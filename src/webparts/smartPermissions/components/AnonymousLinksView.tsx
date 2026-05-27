import * as React from 'react';
import {
  Button,
  Badge,
  Text,
  Title3,
  Body1,
  ProgressBar,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  ArrowDownload24Regular,
  Globe24Regular,
  LockOpen24Regular,
} from '@fluentui/react-icons';

import { SharePointService, isGraphPermissionError } from '../services/SharePointService';
import { SharingLinkEntry } from '../models/models';

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
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    justifyContent: 'space-between',
  },
  summaryGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: tokens.spacingHorizontalL,
    marginBottom: tokens.spacingVerticalL,
    '@media (max-width: 600px)': {
      gridTemplateColumns: '1fr',
    },
  },
  statCard: {
    padding: tokens.spacingVerticalM,
    background: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    textAlign: 'center',
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
  sectionHeading: {
    fontWeight: tokens.fontWeightSemibold,
    display: 'block',
    marginBottom: tokens.spacingVerticalS,
    marginTop: tokens.spacingVerticalL,
  },
});

interface LibrarySummary {
  libraryName: string;
  anonymousCount: number;
  organizationCount: number;
  usersCount: number;
  totalCount: number;
}

function scopeLabel(scope: string): string {
  if (scope === 'anonymous') return 'Anyone';
  if (scope === 'organization') return 'Organization';
  if (scope === 'users') return 'Specific people';
  return scope;
}

function typeLabel(t: string): string {
  if (t === 'view') return 'View';
  if (t === 'edit') return 'Edit';
  if (t === 'review') return 'Review';
  return t;
}

function formatElapsed(s: number): string {
  return `${Math.floor(s / 60)}:${s % 60 < 10 ? '0' : ''}${s % 60}`;
}

export interface AnonymousLinksViewProps {
  sp: SharePointService;
  siteUrl: string;
  onBack: () => void;
}

export const AnonymousLinksView: React.FC<AnonymousLinksViewProps> = ({ sp, siteUrl, onBack }) => {
  const styles = useStyles();

  const [isBusy, setIsBusy] = React.useState(false);
  const [progressMsg, setProgressMsg] = React.useState('');
  const [error, setError] = React.useState('');
  const [links, setLinks] = React.useState<SharingLinkEntry[] | null>(null);
  const [elapsed, setElapsed] = React.useState(0);
  const abortRef = React.useRef<AbortController | null>(null);

  React.useEffect(() => {
    if (!isBusy) { setElapsed(0); return; }
    const start = Date.now();
    const id = setInterval(() => setElapsed(Math.floor((Date.now() - start) / 1000)), 500);
    return () => clearInterval(id);
  }, [isBusy]);

  const handleScan = async (): Promise<void> => {
    abortRef.current?.abort();
    abortRef.current = new AbortController();
    setIsBusy(true);
    setError('');
    setLinks(null);
    setProgressMsg('Starting…');
    try {
      const result = await sp.getSharingLinks(
        siteUrl.trim(),
        (msg) => setProgressMsg(msg),
        abortRef.current.signal,
      );
      if (abortRef.current.signal.aborted) { setProgressMsg('Cancelled.'); return; }
      setLinks(result);
      setProgressMsg(`Scan complete — ${result.length} sharing link${result.length !== 1 ? 's' : ''} found.`);
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        setProgressMsg('Cancelled.');
      } else {
        const msg = err?.message ?? String(err);
        setError(
          isGraphPermissionError(err)
            ? 'Graph API access denied. The Sites.Read.All permission must be approved in the ' +
              'SharePoint Admin Center → Advanced → API access before this report can be used. ' +
              `(Detail: ${msg})`
            : `Error: ${msg}`,
        );
        setProgressMsg('');
      }
    } finally {
      setIsBusy(false);
    }
  };

  React.useEffect(() => { handleScan().catch(console.error); }, []);

  const summaryByLibrary = React.useMemo((): LibrarySummary[] => {
    if (!links) return [];
    const map: { [key: string]: LibrarySummary } = {};
    for (const link of links) {
      if (!map[link.libraryName]) {
        map[link.libraryName] = { libraryName: link.libraryName, anonymousCount: 0, organizationCount: 0, usersCount: 0, totalCount: 0 };
      }
      const entry = map[link.libraryName];
      entry.totalCount++;
      if (link.linkScope === 'anonymous') entry.anonymousCount++;
      else if (link.linkScope === 'organization') entry.organizationCount++;
      else if (link.linkScope === 'users') entry.usersCount++;
    }
    const rows = Object.keys(map).map((k) => map[k]);
    rows.sort((a, b) => b.anonymousCount - a.anonymousCount || b.totalCount - a.totalCount);
    return rows;
  }, [links]);

  // Broad-access links: anonymous + organization (most security-sensitive)
  const broadLinks = React.useMemo((): SharingLinkEntry[] => {
    if (!links) return [];
    return links.filter((l) => l.linkScope === 'anonymous' || l.linkScope === 'organization');
  }, [links]);

  const totalAnonymous    = links ? links.filter((l) => l.linkScope === 'anonymous').length : 0;
  const totalOrganization = links ? links.filter((l) => l.linkScope === 'organization').length : 0;
  const totalSpecific     = links ? links.filter((l) => l.linkScope === 'users').length : 0;

  // CSV exports the full link detail (all links, all scopes)
  const handleExportCsv = (): void => {
    if (!links || links.length === 0) return;
    const header = ['Library', 'Item', 'Scope', 'Type', 'Shared With', 'Link URL', 'Expires'];
    const csvEscape = (v: string): string => `"${v.replace(/"/g, '""')}"`;
    const csv = [
      header.map(csvEscape).join(','),
      ...links.map((l) =>
        [
          l.libraryName,
          l.name,
          scopeLabel(l.linkScope),
          typeLabel(l.linkType),
          l.sharedWith,
          l.linkUrl,
          l.expiresAt ? new Date(l.expiresAt).toLocaleDateString() : '',
        ].map(csvEscape).join(','),
      ),
    ].join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `anonymous-access-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const firstScan = links === null && !isBusy && !error && !progressMsg;

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
        <Title3 style={{ flex: 1 }}>Anonymous Access Summary</Title3>
        {links !== null && !isBusy && links.length > 0 && (
          <Button appearance="secondary" icon={<ArrowDownload24Regular />} onClick={handleExportCsv}>
            Export all links to CSV
          </Button>
        )}
        <Button appearance="primary" onClick={handleScan} disabled={isBusy}>
          {isBusy ? 'Scanning…' : firstScan ? 'Scan' : 'Scan again'}
        </Button>
      </div>

      {/* Permission notice — only shown before any scan has returned results */}
      {links === null && !error && !isBusy && (
        <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>
            <strong>Requires Sites.Read.All</strong> — This report uses the Microsoft Graph API.
            Before first use, a SharePoint Admin must approve the <em>Sites.Read.All</em> permission in the
            SharePoint Admin Center → Advanced → API access.
          </MessageBarBody>
        </MessageBar>
      )}

      {(isBusy || progressMsg) && !error && (
        <div className={styles.progressArea}>
          {isBusy && <ProgressBar aria-label="Scanning sharing links" />}
          <div className={styles.row}>
            <Body1>{progressMsg}</Body1>
            {isBusy && elapsed > 0 && (
              <Body1 style={{ color: tokens.colorNeutralForeground3, whiteSpace: 'nowrap' }}>
                {formatElapsed(elapsed)}
              </Body1>
            )}
          </div>
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

      {!isBusy && links !== null && links.length === 0 && (
        <div className={styles.emptyState}>
          <Globe24Regular style={{ fontSize: '48px', opacity: 0.4 }} />
          <Text weight="semibold">No sharing links found</Text>
          <Body1>
            This site has no active sharing links. The scan completed successfully with Sites.Read.All access confirmed.
          </Body1>
        </div>
      )}

      {links !== null && links.length > 0 && (
        <>
          {/* Top-level stat cards */}
          <div className={styles.summaryGrid}>
            <div className={styles.statCard}>
              <Text style={{ fontSize: tokens.fontSizeBase600, fontWeight: tokens.fontWeightBold, color: tokens.colorStatusDangerForeground1 }}>
                {totalAnonymous}
              </Text>
              <Body1 style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
                Anonymous links
              </Body1>
              <Body1 style={{ display: 'block', fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 }}>
                Anyone with the URL can access
              </Body1>
            </div>
            <div className={styles.statCard}>
              <Text style={{ fontSize: tokens.fontSizeBase600, fontWeight: tokens.fontWeightBold, color: tokens.colorStatusWarningForeground1 }}>
                {totalOrganization}
              </Text>
              <Body1 style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
                Organization-wide links
              </Body1>
              <Body1 style={{ display: 'block', fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 }}>
                All signed-in tenant members
              </Body1>
            </div>
            <div className={styles.statCard}>
              <Text style={{ fontSize: tokens.fontSizeBase600, fontWeight: tokens.fontWeightBold, color: tokens.colorBrandForeground1 }}>
                {totalSpecific}
              </Text>
              <Body1 style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
                Specific-people links
              </Body1>
              <Body1 style={{ display: 'block', fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 }}>
                Named individuals only
              </Body1>
            </div>
          </div>

          {/* Per-library summary */}
          <Text className={styles.sectionHeading}>By library</Text>
          <table className={styles.table} aria-label="Sharing links by library" style={{ marginBottom: tokens.spacingVerticalL }}>
            <thead>
              <tr>
                <th className={styles.th}>Library</th>
                <th className={styles.th} style={{ textAlign: 'center' }}>Anonymous</th>
                <th className={styles.th} style={{ textAlign: 'center' }}>Organization</th>
                <th className={styles.th} style={{ textAlign: 'center' }}>Specific People</th>
                <th className={styles.th} style={{ textAlign: 'center' }}>Total</th>
              </tr>
            </thead>
            <tbody>
              {summaryByLibrary.map((row, i) => (
                <tr key={i}>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>{row.libraryName}</Text>
                  </td>
                  <td className={styles.td} style={{ textAlign: 'center' }}>
                    {row.anonymousCount > 0 ? (
                      <Badge appearance="filled" color="danger" size="small">{row.anonymousCount}</Badge>
                    ) : (
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>—</Text>
                    )}
                  </td>
                  <td className={styles.td} style={{ textAlign: 'center' }}>
                    {row.organizationCount > 0 ? (
                      <Badge appearance="filled" color="warning" size="small">{row.organizationCount}</Badge>
                    ) : (
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>—</Text>
                    )}
                  </td>
                  <td className={styles.td} style={{ textAlign: 'center' }}>
                    {row.usersCount > 0 ? (
                      <Badge appearance="filled" color="informative" size="small">{row.usersCount}</Badge>
                    ) : (
                      <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>—</Text>
                    )}
                  </td>
                  <td className={styles.td} style={{ textAlign: 'center' }}>
                    <Text style={{ fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold }}>
                      {row.totalCount}
                    </Text>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* Broad-access links detail (anonymous + org) */}
          {broadLinks.length > 0 && (
            <>
              <Text className={styles.sectionHeading}>
                <LockOpen24Regular style={{ verticalAlign: 'middle', marginRight: '6px' }} />
                Anonymous and organization-wide links ({broadLinks.length})
              </Text>
              <table className={styles.table} aria-label="Broad-access links">
                <thead>
                  <tr>
                    <th className={styles.th}>Library</th>
                    <th className={styles.th}>Item</th>
                    <th className={styles.th}>Scope</th>
                    <th className={styles.th}>Type</th>
                    <th className={styles.th}>Expires</th>
                  </tr>
                </thead>
                <tbody>
                  {broadLinks.map((link, i) => (
                    <tr key={i}>
                      <td className={styles.td}>
                        <Text style={{ fontSize: tokens.fontSizeBase200 }}>{link.libraryName}</Text>
                      </td>
                      <td className={styles.td}>
                        <a
                          href={link.linkUrl || link.webUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{ color: tokens.colorBrandForegroundLink, textDecoration: 'none', fontSize: tokens.fontSizeBase200 }}
                        >
                          {link.name}
                        </a>
                      </td>
                      <td className={styles.td}>
                        <Badge
                          appearance="filled"
                          color={link.linkScope === 'anonymous' ? 'danger' : 'warning'}
                          size="small"
                        >
                          {link.linkScope === 'anonymous' ? 'Anyone' : 'Organization'}
                        </Badge>
                      </td>
                      <td className={styles.td}>
                        <Text style={{ fontSize: tokens.fontSizeBase200 }}>{typeLabel(link.linkType)}</Text>
                      </td>
                      <td className={styles.td}>
                        <Text style={{
                          fontSize: tokens.fontSizeBase200,
                          color: link.expiresAt
                            ? (new Date(link.expiresAt) < new Date()
                              ? tokens.colorStatusDangerForeground1
                              : tokens.colorStatusWarningForeground1)
                            : tokens.colorNeutralForeground3,
                        }}>
                          {link.expiresAt ? new Date(link.expiresAt).toLocaleDateString() : '—'}
                        </Text>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          )}
        </>
      )}
    </div>
  );
};
