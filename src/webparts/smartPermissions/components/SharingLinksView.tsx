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
  Link24Regular,
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

function scopeBadgeColor(scope: string): 'danger' | 'warning' | 'success' | 'informative' {
  if (scope === 'anonymous') return 'danger';
  if (scope === 'organization') return 'warning';
  if (scope === 'users') return 'success';
  return 'informative';
}

function scopeLabel(scope: string): string {
  if (scope === 'anonymous') return 'Anyone';
  if (scope === 'organization') return 'Organization';
  if (scope === 'users') return 'Specific people';
  return scope;
}

function typeLabel(type: string): string {
  if (type === 'view') return 'View';
  if (type === 'edit') return 'Edit';
  if (type === 'review') return 'Review';
  return type;
}

function formatElapsed(s: number): string {
  return `${Math.floor(s / 60)}:${s % 60 < 10 ? '0' : ''}${s % 60}`;
}

export interface SharingLinksViewProps {
  sp: SharePointService;
  siteUrl: string;
  onBack: () => void;
}

export const SharingLinksView: React.FC<SharingLinksViewProps> = ({ sp, siteUrl, onBack }) => {
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

      if (abortRef.current.signal.aborted) {
        setProgressMsg('Cancelled.');
        return;
      }

      setLinks(result);
      setProgressMsg(`Scan complete — ${result.length} sharing link(s) found.`);

      if (result.length > 0 && 'Notification' in window && Notification.permission === 'granted') {
        new Notification('Smart Permissions — Sharing Links scan complete', {
          body: `${result.length} sharing link(s) found on ${siteUrl}.`,
        });
      }
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        setProgressMsg('Cancelled.');
      } else {
        const msg = err?.message ?? String(err);
        setError(
          isGraphPermissionError(err)
            ? 'Graph API access denied. The Sites.Read.All permission must be approved in the ' +
              'SharePoint Admin Center → Advanced → API access before Sharing Links can be used. ' +
              `(Detail: ${msg})`
            : `Error: ${msg}`,
        );
        setProgressMsg('');
      }
    } finally {
      setIsBusy(false);
    }
  };

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
        ]
          .map(csvEscape)
          .join(','),
      ),
    ].join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `sharing-links-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Auto-scan and request notification permission on mount
  React.useEffect(() => {
    if ('Notification' in window && Notification.permission === 'default') {
      Notification.requestPermission().catch(() => { /* ignore */ });
    }
    handleScan().catch(console.error);
  }, []);

  return (
    <div className={styles.root}>
      {/* Screen reader status */}
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
          aria-label="Back to home"
        >
          Back
        </Button>
        <Title3 style={{ flex: 1 }}>Sharing Links</Title3>
        {links !== null && !isBusy && links.length > 0 && (
          <Button
            appearance="secondary"
            icon={<ArrowDownload24Regular />}
            onClick={handleExportCsv}
          >
            Export to CSV
          </Button>
        )}
        <Button appearance="primary" onClick={handleScan} disabled={isBusy}>
          {isBusy ? 'Scanning…' : 'Scan again'}
        </Button>
      </div>

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
            <Button
              appearance="secondary"
              size="small"
              onClick={() => abortRef.current?.abort()}
            >
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
          <Link24Regular style={{ fontSize: '48px', opacity: 0.4 }} />
          <Text weight="semibold">No sharing links found</Text>
          <Body1>
            This site has no active sharing links. The scan completed successfully with Sites.Read.All access confirmed.
          </Body1>
        </div>
      )}

      {links !== null && links.length > 0 && (
        <>
          <Body1
            style={{
              color: tokens.colorNeutralForeground3,
              marginBottom: tokens.spacingVerticalS,
              display: 'block',
            }}
          >
            {links.length} sharing link{links.length !== 1 ? 's' : ''} found
          </Body1>
          <table className={styles.table} aria-label="Sharing links">
            <thead>
              <tr>
                <th className={styles.th}>Library</th>
                <th className={styles.th}>Item</th>
                <th className={styles.th}>Scope</th>
                <th className={styles.th}>Type</th>
                <th className={styles.th}>Shared With</th>
                <th className={styles.th}>Expires</th>
              </tr>
            </thead>
            <tbody>
              {links.map((link, i) => (
                <tr key={i}>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>{link.libraryName}</Text>
                  </td>
                  <td className={styles.td}>
                    <a
                      href={link.linkUrl || link.webUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{
                        color: tokens.colorBrandForegroundLink,
                        textDecoration: 'none',
                        fontSize: tokens.fontSizeBase200,
                      }}
                    >
                      {link.name}
                    </a>
                  </td>
                  <td className={styles.td}>
                    <Badge
                      appearance="filled"
                      color={scopeBadgeColor(link.linkScope)}
                      size="small"
                    >
                      {scopeLabel(link.linkScope)}
                    </Badge>
                  </td>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200 }}>
                      {typeLabel(link.linkType)}
                    </Text>
                  </td>
                  <td className={styles.td}>
                    <Text
                      style={{
                        fontSize: tokens.fontSizeBase200,
                        color: tokens.colorNeutralForeground3,
                      }}
                    >
                      {link.sharedWith || '—'}
                    </Text>
                  </td>
                  <td className={styles.td}>
                    <Text
                      style={{
                        fontSize: tokens.fontSizeBase200,
                        color: link.expiresAt
                          ? (new Date(link.expiresAt) < new Date()
                            ? tokens.colorStatusDangerForeground1
                            : tokens.colorStatusWarningForeground1)
                          : tokens.colorNeutralForeground3,
                      }}
                    >
                      {link.expiresAt
                        ? new Date(link.expiresAt).toLocaleDateString()
                        : '—'}
                    </Text>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
    </div>
  );
};
