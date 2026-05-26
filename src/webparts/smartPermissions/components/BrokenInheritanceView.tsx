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
  Select,
  Label,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  ArrowDownload24Regular,
  BranchFork24Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { BrokenInheritanceEntry } from '../models/models';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '1100px',
    margin: '0 auto',
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
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    justifyContent: 'space-between',
  },
});

function typeBadgeColor(t: string): 'important' | 'informative' | 'subtle' {
  if (t === 'Library') return 'important';
  if (t === 'Folder') return 'informative';
  return 'subtle';
}

function formatElapsed(s: number): string {
  return `${Math.floor(s / 60)}:${s % 60 < 10 ? '0' : ''}${s % 60}`;
}

export interface BrokenInheritanceViewProps {
  sp: SharePointService;
  siteUrl: string;
  includeHidden: boolean;
  onBack: () => void;
}

export const BrokenInheritanceView: React.FC<BrokenInheritanceViewProps> = ({
  sp,
  siteUrl,
  includeHidden,
  onBack,
}) => {
  const styles = useStyles();
  const siteOrigin = React.useMemo(() => { try { return new URL(siteUrl).origin; } catch { return ''; } }, [siteUrl]);

  const [isBusy, setIsBusy] = React.useState(false);
  const [progressMsg, setProgressMsg] = React.useState('');
  const [error, setError] = React.useState('');
  const [entries, setEntries] = React.useState<BrokenInheritanceEntry[] | null>(null);
  const [typeFilter, setTypeFilter] = React.useState<'All' | 'Library' | 'Folder' | 'File'>('All');
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
    setEntries(null);
    setProgressMsg('Starting…');
    try {
      const result = await sp.scanBrokenInheritance(
        siteUrl.trim(),
        includeHidden,
        (msg) => setProgressMsg(msg),
        abortRef.current.signal,
      );
      if (abortRef.current.signal.aborted) { setProgressMsg('Cancelled.'); return; }
      setEntries(result);
      setProgressMsg(`Scan complete — ${result.length} item${result.length !== 1 ? 's' : ''} with unique permissions found.`);
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        setProgressMsg('Cancelled.');
      } else {
        setError(`Error: ${err?.message ?? String(err)}`);
        setProgressMsg('');
      }
    } finally {
      setIsBusy(false);
    }
  };

  const filtered = React.useMemo((): BrokenInheritanceEntry[] => {
    if (!entries) return [];
    if (typeFilter === 'All') return entries;
    return entries.filter((e) => e.objectType === typeFilter);
  }, [entries, typeFilter]);

  const handleExportCsv = (): void => {
    if (!entries || entries.length === 0) return;
    const rows = filtered;
    const header = ['Type', 'Name', 'Full URL', 'Server-relative URL', 'Depth'];
    const csvEscape = (v: string): string => `"${v.replace(/"/g, '""')}"`;
    const csv = [
      header.map(csvEscape).join(','),
      ...rows.map((e) =>
        [
          e.objectType,
          e.name,
          siteOrigin + e.serverRelativeUrl,
          e.serverRelativeUrl,
          String(e.depth),
        ].map(csvEscape).join(','),
      ),
    ].join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `broken-inheritance-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  React.useEffect(() => { handleScan().catch(console.error); }, []);

  const libCount    = entries ? entries.filter((e) => e.objectType === 'Library').length : 0;
  const folderCount = entries ? entries.filter((e) => e.objectType === 'Folder').length : 0;
  const fileCount   = entries ? entries.filter((e) => e.objectType === 'File').length : 0;

  const firstScan = entries === null && !isBusy && !error && !progressMsg;

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
        <Title3 style={{ flex: 1 }}>Broken Inheritance Finder</Title3>
        {entries !== null && !isBusy && entries.length > 0 && (
          <Button appearance="secondary" icon={<ArrowDownload24Regular />} onClick={handleExportCsv}>
            Export to CSV
          </Button>
        )}
        <Button appearance="primary" onClick={handleScan} disabled={isBusy}>
          {isBusy ? 'Scanning…' : firstScan ? 'Scan' : 'Scan again'}
        </Button>
      </div>

      <Body1 style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: tokens.spacingVerticalM }}>
        Finds all libraries, folders, and files that have <strong>unique (broken) permissions</strong> — meaning their
        permission inheritance has been stopped and they have their own role assignments instead of inheriting from their
        parent. No additional API permissions are required beyond read access to the site.
      </Body1>

      {(isBusy || progressMsg) && !error && (
        <div className={styles.progressArea}>
          {isBusy && <ProgressBar aria-label="Scanning for broken inheritance" />}
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

      {!isBusy && entries !== null && entries.length === 0 && (
        <div className={styles.emptyState}>
          <BranchFork24Regular style={{ fontSize: '48px', opacity: 0.4 }} />
          <Text weight="semibold">No broken inheritance found</Text>
          <Body1>All libraries, folders, and files inherit permissions from their parent.</Body1>
        </div>
      )}

      {entries !== null && entries.length > 0 && (
        <>
          <div style={{ marginBottom: tokens.spacingVerticalM }}>
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
              {libCount} librar{libCount !== 1 ? 'ies' : 'y'} &nbsp;·&nbsp;
              {folderCount} folder{folderCount !== 1 ? 's' : ''} &nbsp;·&nbsp;
              {fileCount} file{fileCount !== 1 ? 's' : ''} with unique permissions
            </Body1>
          </div>

          <div className={styles.filterRow}>
            <Label htmlFor="typeFilter">Filter by type:</Label>
            <Select
              id="typeFilter"
              value={typeFilter}
              onChange={(_, d) => setTypeFilter(d.value as typeof typeFilter)}
              style={{ width: '160px' }}
            >
              <option value="All">All ({entries.length})</option>
              <option value="Library">Library ({libCount})</option>
              <option value="Folder">Folder ({folderCount})</option>
              <option value="File">File ({fileCount})</option>
            </Select>
            {typeFilter !== 'All' && (
              <Body1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
                Showing {filtered.length} of {entries.length}
              </Body1>
            )}
          </div>

          <table className={styles.table} aria-label="Items with unique permissions">
            <thead>
              <tr>
                <th className={styles.th}>Type</th>
                <th className={styles.th}>Name</th>
                <th className={styles.th}>Path</th>
                <th className={styles.th}>Depth</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((e, i) => (
                <tr key={i}>
                  <td className={styles.td}>
                    <Badge appearance="tint" color={typeBadgeColor(e.objectType)} size="small">
                      {e.objectType}
                    </Badge>
                  </td>
                  <td className={styles.td}>
                    {siteOrigin ? (
                      <a
                        href={siteOrigin + e.serverRelativeUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ color: tokens.colorBrandForegroundLink, textDecoration: 'none', fontSize: tokens.fontSizeBase200 }}
                      >
                        {e.name}
                      </a>
                    ) : (
                      <Text style={{ fontSize: tokens.fontSizeBase200 }}>{e.name}</Text>
                    )}
                  </td>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                      {e.serverRelativeUrl}
                    </Text>
                  </td>
                  <td className={styles.td}>
                    <Text style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                      {e.depth}
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
