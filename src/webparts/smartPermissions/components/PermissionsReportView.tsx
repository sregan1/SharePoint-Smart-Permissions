import * as React from 'react';
import {
  Button,
  Checkbox,
  Input,
  Label,
  Field,
  RadioGroup,
  Radio,
  SpinButton,
  ProgressBar,
  Text,
  Title3,
  Body1,
  Badge,
  Divider,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  DocumentArrowDown24Regular,
  Globe24Regular,
  BookDatabase24Regular,
  Folder24Regular,
  FolderOpen24Regular,
  History24Regular,
  Delete24Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import { ReportOptions, ReportScope, PermissionEntry, ObjectType, ScanProgress, StoredReport, LibraryInfo } from '../models/models';


const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '760px',
    margin: '0 auto',
    minHeight: '500px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    flexWrap: 'wrap',
  },
  progressArea: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    background: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  resultArea: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    background: tokens.colorStatusSuccessBackground1,
    borderRadius: tokens.borderRadiusMedium,
  },
  radioBox: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    cursor: 'pointer',
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

export interface PermissionsReportViewProps {
  sp: SharePointService;
  excel: ExcelExportService;
  siteUrl: string;
  includeHidden: boolean;
  excludeLimitedAccess: boolean;
  onBack: () => void;
}

export const PermissionsReportView: React.FC<PermissionsReportViewProps> = ({
  sp,
  excel,
  siteUrl,
  includeHidden,
  excludeLimitedAccess,
  onBack,
}) => {
  const styles = useStyles();

  // ── Form state ──
  const [allSites, setAllSites] = React.useState(false);
  const [scope, setScope] = React.useState<string>('Site');
  const [folderDepth, setFolderDepth] = React.useState(2);
  const [expandGroups, setExpandGroups] = React.useState(true);

  // ── Run state ──
  const [isBusy, setIsBusy] = React.useState(false);
  const [scanProgress, setScanProgress] = React.useState<ScanProgress>({ message: '', scanned: 0, libsDone: 0, libsTotal: 0 });
  const [elapsed, setElapsed] = React.useState(0);
  const [error, setError] = React.useState('');
  const [entries, setEntries] = React.useState<PermissionEntry[] | null>(null);
  const [groupPermissionDenied, setGroupPermissionDenied] = React.useState(false);
  const [isExporting, setIsExporting] = React.useState(false);
  const [liveCount, setLiveCount] = React.useState(0);
  const liveCountRef = React.useRef(0);

  // ── Filter state ──
  const [filterText, setFilterText] = React.useState('');
  const [filterExternalOnly, setFilterExternalOnly] = React.useState(false);
  const [filterUniqueOnly, setFilterUniqueOnly] = React.useState(false);

  const filteredEntries = React.useMemo(() => {
    if (!entries) return null;
    const lc = filterText.toLowerCase();
    return entries.filter((e) => {
      if (filterUniqueOnly && !e.hasUniquePermissions) return false;
      if (filterExternalOnly && !e.uniquePermissions.some((u) => u.loginName.toLowerCase().indexOf('#ext#') !== -1)) return false;
      if (excludeLimitedAccess && !e.uniquePermissions.some((u) => u.roles.length > 0)) return false;
      if (!lc) return true;
      if (e.name.toLowerCase().includes(lc)) return true;
      if (e.serverRelativeUrl.toLowerCase().includes(lc)) return true;
      if (e.uniquePermissions.some((u) => u.displayName.toLowerCase().includes(lc))) return true;
      return false;
    });
  }, [entries, filterText, filterExternalOnly, filterUniqueOnly, excludeLimitedAccess]);

  // ── Library picker state ──
  const [availableLibraries, setAvailableLibraries] = React.useState<LibraryInfo[]>([]);
  const [selectedLibraryUrls, setSelectedLibraryUrls] = React.useState<Set<string>>(new Set());
  const [librariesLoading, setLibrariesLoading] = React.useState(false);

  // Auto-load library list whenever scope changes to something that uses libraries
  React.useEffect(() => {
    if (scope === 'Site') { setAvailableLibraries([]); setSelectedLibraryUrls(new Set()); return; }
    setLibrariesLoading(true);
    sp.getLibraries(siteUrl.trim(), undefined, includeHidden)
      .then((libs) => {
        setAvailableLibraries(libs);
        setSelectedLibraryUrls(new Set(libs.map((l) => l.serverRelativeUrl)));
        setLibrariesLoading(false);
      })
      .catch(() => { setLibrariesLoading(false); /* silent — library picker simply won't show */ });
  }, [scope, siteUrl, includeHidden]);

  const toggleLibrary = (url: string): void => {
    setSelectedLibraryUrls((prev) => {
      const next = new Set(prev);
      if (next.has(url)) { next.delete(url); } else { next.add(url); }
      return next;
    });
  };

  // ── History state ──
  const [showHistory, setShowHistory] = React.useState(false);
  const [historyItems, setHistoryItems] = React.useState<StoredReport[]>([]);
  const [exportingHistoryId, setExportingHistoryId] = React.useState<string | null>(null);
  const [historySortCol, setHistorySortCol] = React.useState<'timestamp' | 'siteUrl' | 'scope' | 'total' | 'unique'>('timestamp');
  const [historySortAsc, setHistorySortAsc] = React.useState(false);

  const sortedHistoryItems = React.useMemo(() => {
    const col = historySortCol;
    return [...historyItems].sort((a, b) => {
      let va: string | number, vb: string | number;
      if (col === 'timestamp') { va = a.timestamp; vb = b.timestamp; }
      else if (col === 'siteUrl') { va = a.siteUrl; vb = b.siteUrl; }
      else if (col === 'scope') { va = (a.options.allSites ? 'All · ' : '') + a.options.scope; vb = (b.options.allSites ? 'All · ' : '') + b.options.scope; }
      else if (col === 'total') { va = a.summary.totalObjects; vb = b.summary.totalObjects; }
      else { va = a.summary.uniqueCount; vb = b.summary.uniqueCount; }
      if (va < vb) return historySortAsc ? -1 : 1;
      if (va > vb) return historySortAsc ? 1 : -1;
      return 0;
    });
  }, [historyItems, historySortCol, historySortAsc]);

  const handleHistorySort = (col: typeof historySortCol): void => {
    if (historySortCol === col) { setHistorySortAsc((v) => !v); } else { setHistorySortCol(col); setHistorySortAsc(true); }
  };

  const sortIndicator = (col: typeof historySortCol): string =>
    historySortCol !== col ? '' : historySortAsc ? ' ▲' : ' ▼';

  const abortRef = React.useRef<AbortController | null>(null);
  const scanStartRef = React.useRef<number>(0);
  const historyService = React.useRef(new ReportHistoryService());

  React.useEffect(() => {
    if (!isBusy) { setElapsed(0); return; }
    const start = Date.now();
    const id = setInterval(() => setElapsed(Math.floor((Date.now() - start) / 1000)), 500);
    return () => clearInterval(id);
  }, [isBusy]);

  // Load history on mount; request notification permission
  React.useEffect(() => {
    historyService.current.getAll()
      .then(setHistoryItems)
      .catch(() => { /* IndexedDB unavailable — history simply won't show */ });
    if ('Notification' in window && Notification.permission === 'default') {
      Notification.requestPermission().catch(() => { /* ignore */ });
    }
  }, []);

  // Clear stale results when any option that affects output changes
  React.useEffect(() => {
    setEntries(null);
  }, [allSites, scope, folderDepth, expandGroups]);

  const formatElapsed = (secs: number): string => {
    const m = Math.floor(secs / 60);
    const s = secs % 60;
    return `${m}:${s < 10 ? '0' : ''}${s}`;
  };

  const isRootSite = React.useMemo(() => {
    try {
      return new URL(siteUrl).pathname.replace(/\/$/, '') === '';
    } catch {
      return false;
    }
  }, [siteUrl]);

  const handleRun = async (): Promise<void> => {
    abortRef.current?.abort();
    abortRef.current = new AbortController();
    scanStartRef.current = Date.now();
    setIsBusy(true);
    setError('');
    setEntries(null);
    setFilterText('');
    setFilterExternalOnly(false);
    setFilterUniqueOnly(false);
    setLiveCount(0);
    liveCountRef.current = 0;
    setScanProgress({ message: 'Starting scan…', scanned: 0, libsDone: 0, libsTotal: 0 });

    // Flush the live item count to state every 500ms so React batches renders
    const flushTimer = setInterval(() => setLiveCount(liveCountRef.current), 500);

    try {
      const allSelected = selectedLibraryUrls.size === 0 || selectedLibraryUrls.size === availableLibraries.length;
      const options: ReportOptions = {
        siteUrl: siteUrl.trim(),
        allSites,
        scope: scope as ReportScope,
        folderDepth,
        includeHidden,
        expandGroups,
        libraryUrls: allSelected ? undefined : Array.from(selectedLibraryUrls),
      };

      const { entries: scannedEntries, groupPermissionDenied: permDenied } = await sp.scanPermissions(
        options,
        (progress) => setScanProgress(progress),
        abortRef.current.signal,
        () => { liveCountRef.current += 1; },
      );

      if (abortRef.current.signal.aborted) {
        setScanProgress((prev) => ({ ...prev, message: 'Cancelled.' }));
        return;
      }

      setEntries(scannedEntries);
      setGroupPermissionDenied(permDenied);
      const uniqueCount = scannedEntries.filter((e) => e.hasUniquePermissions).length;
      setScanProgress((prev) => ({
        ...prev,
        message:
          `Scan complete — ${scannedEntries.length} object(s) found, ` +
          `${uniqueCount} with unique permissions.`,
      }));

      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('Smart Permissions — Scan complete', {
          body: `${scannedEntries.length} objects, ${uniqueCount} with unique permissions.`,
        });
      }

      // Save to history (errors are swallowed — never block the user from seeing results)
      const storedReport: StoredReport = {
        id: Date.now().toString(),
        timestamp: new Date().toISOString(),
        siteUrl: siteUrl.trim(),
        options: { allSites, scope: scope as ReportScope, folderDepth, expandGroups },
        summary: {
          totalObjects: scannedEntries.length,
          uniqueCount,
          inheritedCount: scannedEntries.length - uniqueCount,
          durationSeconds: Math.round((Date.now() - scanStartRef.current) / 1000),
        },
        entries: scannedEntries,
      };
      historyService.current.add(storedReport)
        .then(() => historyService.current.getAll())
        .then(setHistoryItems)
        .catch(() => { /* storage unavailable */ });
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        setScanProgress((prev) => ({ ...prev, message: 'Cancelled.' }));
      } else {
        setError(`Error: ${err?.message ?? String(err)}`);
        setScanProgress((prev) => ({ ...prev, message: '' }));
      }
    } finally {
      clearInterval(flushTimer);
      setIsBusy(false);
    }
  };

  const handleCancel = (): void => {
    abortRef.current?.abort();
  };

  const applyExportFilters = (entriesToFilter: PermissionEntry[]): PermissionEntry[] => {
    let result = filterUniqueOnly
      ? entriesToFilter.filter((e) => e.hasUniquePermissions)
      : entriesToFilter;
    if (!filterExternalOnly) return result;
    return result.map((e) => ({
      ...e,
      uniquePermissions: e.uniquePermissions.filter((u) =>
        u.loginName.toLowerCase().indexOf('#ext#') !== -1,
      ),
    }));
  };

  const handleExport = async (): Promise<void> => {
    const toExport = applyExportFilters(filteredEntries ?? entries ?? []);
    if (toExport.length === 0) return;
    setIsExporting(true);
    try {
      await excel.export(toExport, siteUrl.trim());
    } catch (err: any) {
      setError(`Export error: ${err?.message ?? String(err)}`);
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportCsv = (): void => {
    const toExport = applyExportFilters(filteredEntries ?? entries ?? []);
    excel.exportPermissionsCsv(toExport, siteUrl.trim());
  };

  const handleHistoryExport = async (item: StoredReport): Promise<void> => {
    setExportingHistoryId(item.id);
    try {
      await excel.export(item.entries, item.siteUrl);
    } catch (err: any) {
      setError(`Export error: ${err?.message ?? String(err)}`);
    } finally {
      setExportingHistoryId(null);
    }
  };

  const handleHistoryDelete = async (id: string): Promise<void> => {
    await historyService.current.delete(id).catch(() => { /* ignore */ });
    setHistoryItems((prev) => prev.filter((r) => r.id !== id));
  };

  const scopeLabel = (s: string): string =>
    ({ Site: 'Site only', Library: 'Libraries', Folder: 'Folders', Item: 'Files & Folders' } as Record<string, string>)[s] ?? s;

  return (
    <div className={styles.root}>
      {/* Screen reader announcements */}
      <div role="status" aria-live="polite" style={{ position: 'absolute', width: '1px', height: '1px', overflow: 'hidden', clip: 'rect(0,0,0,0)' }}>
        {!isBusy && scanProgress.message ? scanProgress.message : ''}
      </div>

      {/* Header */}
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={() => {
            if (isBusy && !window.confirm('A scan is in progress. Leave and cancel?')) return;
            onBack();
          }}
          disabled={false}
          aria-label="Back to home"
        >
          Back
        </Button>
        <Title3 style={{ flex: 1 }}>Permissions Report</Title3>
        <Button
          appearance="subtle"
          icon={<History24Regular />}
          onClick={() => setShowHistory((v) => !v)}
          disabled={isBusy}
        >
          History{historyItems.length > 0 ? ` (${historyItems.length})` : ''}
        </Button>
      </div>

      {/* ── History panel ── */}
      {showHistory && (
        <div>
          <div className={styles.row} style={{ marginBottom: tokens.spacingVerticalM }}>
            <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={() => setShowHistory(false)}>
              Back to scan
            </Button>
          </div>

          {historyItems.length === 0 ? (
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No reports saved yet.</Body1>
          ) : (
            <table className={styles.historyTable}>
              <thead>
                <tr>
                  {(
                    [
                      { col: 'timestamp', label: 'Date / Time' },
                      { col: 'siteUrl', label: 'Site' },
                      { col: 'scope', label: 'Scope' },
                      { col: 'total', label: 'Objects' },
                      { col: 'unique', label: 'Unique' },
                    ] as { col: typeof historySortCol; label: string }[]
                  ).map(({ col, label }) => (
                    <th
                      key={col}
                      className={styles.historyTh}
                      onClick={() => handleHistorySort(col)}
                      style={{ cursor: 'pointer', userSelect: 'none' }}
                    >
                      {label}{sortIndicator(col)}
                    </th>
                  ))}
                  <th className={styles.historyTh} />
                </tr>
              </thead>
              <tbody>
                {sortedHistoryItems.map((item) => (
                  <tr key={item.id}>
                    <td className={styles.historyTd} style={{ whiteSpace: 'nowrap' }}>
                      {new Date(item.timestamp).toLocaleString()}
                    </td>
                    <td className={styles.historyTd}>
                      <Text style={{ fontSize: tokens.fontSizeBase200, wordBreak: 'break-all' }}>
                        {item.siteUrl}
                      </Text>
                    </td>
                    <td className={styles.historyTd} style={{ whiteSpace: 'nowrap' }}>
                      {item.options.allSites ? 'All sites · ' : ''}{scopeLabel(item.options.scope)}
                    </td>
                    <td className={styles.historyTd}>{item.summary.totalObjects}</td>
                    <td className={styles.historyTd}>{item.summary.uniqueCount}</td>
                    <td className={styles.historyTd}>
                      <div style={{ display: 'flex', gap: tokens.spacingHorizontalS }}>
                        <Button
                          size="small"
                          appearance="primary"
                          icon={<DocumentArrowDown24Regular />}
                          onClick={() => handleHistoryExport(item)}
                          disabled={exportingHistoryId === item.id}
                        >
                          {exportingHistoryId === item.id ? 'Exporting…' : 'Export'}
                        </Button>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<Delete24Regular />}
                          onClick={() => handleHistoryDelete(item.id)}
                          title="Delete this report"
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

      {/* ── Scan form (hidden when history is open) ── */}
      {!showHistory && <div className={styles.form}>
        {/* All-sites toggle */}
        <Checkbox
          label="Scan all site collections in this tenant (only available in root site)"
          checked={allSites}
          onChange={(_, d) => setAllSites(!!d.checked)}
          disabled={!isRootSite || isBusy}
        />

        <Divider />

        {/* Scope */}
        <Field label="Scan depth">
          <RadioGroup
            value={scope}
            onChange={(_, d) => setScope(d.value)}
            layout="horizontal"
            disabled={isBusy}
            style={{ flexWrap: 'wrap', gap: tokens.spacingHorizontalS }}
          >
            {([
              { value: 'Site', Icon: Globe24Regular, label: 'Site only' },
              { value: 'Library', Icon: BookDatabase24Regular, label: 'Libraries' },
              { value: 'Folder', Icon: Folder24Regular, label: 'Folders' },
              { value: 'Item', Icon: FolderOpen24Regular, label: 'Files & Folders' },
            ] as const).map(({ value, Icon, label }) => (
              <div
                key={value}
                className={styles.radioBox}
                style={scope === value ? {
                  borderWidth: '2px',
                  borderColor: tokens.colorBrandForeground1,
                  background: tokens.colorBrandBackground2,
                } : undefined}
                onClick={() => { if (!isBusy) setScope(value); }}
              >
                <Radio
                  value={value}
                  label={
                    <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      <Icon style={{ fontSize: '16px' }} />
                      {label}
                    </span>
                  }
                />
              </div>
            ))}
          </RadioGroup>
        </Field>

        {/* Folder depth — always rendered so it doesn't shift the Run Report button */}
        <div className={styles.row} style={{ visibility: scope === 'Folder' ? 'visible' : 'hidden' }}>
          <Label>Folder depth limit:</Label>
          <SpinButton
            value={folderDepth}
            min={1}
            max={10}
            onChange={(_, d) =>
              setFolderDepth(
                d.value !== undefined ? d.value : parseInt(d.displayValue ?? '2', 10),
              )
            }
            style={{ width: '80px' }}
            disabled={isBusy}
          />
        </div>

        <Checkbox
          label="Expand group members in report (SharePoint groups, Security groups, and M365 groups)"
          checked={expandGroups}
          onChange={(_, d) => setExpandGroups(!!d.checked)}
          disabled={isBusy}
        />

        {/* ── Library picker (only when scope includes libraries) ── */}
        {scope !== 'Site' && availableLibraries.length > 0 && (
          <div
            style={{
              border: `1px solid ${tokens.colorNeutralStroke1}`,
              borderRadius: tokens.borderRadiusMedium,
              padding: tokens.spacingVerticalS,
            }}
          >
            <div className={styles.row} style={{ marginBottom: tokens.spacingVerticalXS }}>
              <Label weight="semibold">
                Libraries to scan
                {selectedLibraryUrls.size < availableLibraries.length && (
                  <span style={{ color: tokens.colorNeutralForeground3, fontWeight: 'normal', marginLeft: '6px' }}>
                    ({selectedLibraryUrls.size} of {availableLibraries.length} selected)
                  </span>
                )}
              </Label>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => setSelectedLibraryUrls(new Set(availableLibraries.map((l) => l.serverRelativeUrl)))}
                disabled={isBusy || selectedLibraryUrls.size === availableLibraries.length}
              >
                All
              </Button>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => setSelectedLibraryUrls(new Set())}
                disabled={isBusy || selectedLibraryUrls.size === 0}
              >
                None
              </Button>
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalM}` }}>
              {availableLibraries.map((lib) => (
                <Checkbox
                  key={lib.serverRelativeUrl}
                  label={lib.title}
                  checked={selectedLibraryUrls.has(lib.serverRelativeUrl)}
                  onChange={() => toggleLibrary(lib.serverRelativeUrl)}
                  disabled={isBusy}
                />
              ))}
            </div>
          </div>
        )}
        {scope !== 'Site' && librariesLoading && (
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>Loading libraries…</Body1>
        )}

        <Divider />

        {/* Action buttons */}
        <div className={styles.row}>
          <Button
            appearance="primary"
            onClick={handleRun}
            disabled={isBusy || (scope !== 'Site' && availableLibraries.length > 0 && selectedLibraryUrls.size === 0)}
          >
            Run Report
          </Button>
          {isBusy && (
            <Button appearance="secondary" onClick={handleCancel}>
              Cancel
            </Button>
          )}
        </div>

        {/* Progress */}
        {(isBusy || scanProgress.message) && !error && (
          <div className={styles.progressArea}>
            {isBusy && (
              <ProgressBar
                value={scanProgress.libsTotal > 0 ? scanProgress.libsDone / scanProgress.libsTotal : undefined}
              />
            )}
            <div className={styles.row} style={{ justifyContent: 'space-between' }}>
              <Body1>{scanProgress.message}</Body1>
              {isBusy && elapsed > 0 && (
                <Body1 style={{ color: tokens.colorNeutralForeground3, whiteSpace: 'nowrap' }}>
                  {formatElapsed(elapsed)}
                </Body1>
              )}
            </div>
            {isBusy && liveCount > 0 && (
              <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                {liveCount.toLocaleString()} items found so far
                {scanProgress.libsTotal > 0 && ` · Library ${scanProgress.libsDone} of ${scanProgress.libsTotal}`}
              </Body1>
            )}
          </div>
        )}

        {/* Error */}
        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}

        {/* Results */}
        {entries && !isBusy && (
          <div className={styles.resultArea}>
            <div className={styles.row}>
              <Text weight="semibold">Scan complete</Text>
              <Badge appearance="filled" color="success">{entries.length} objects</Badge>
              <Badge appearance="filled" color="warning">
                {entries.filter((e) => e.hasUniquePermissions).length} unique
              </Badge>
              <Badge appearance="outline">
                {entries.filter((e) => !e.hasUniquePermissions).length} inherited
              </Badge>
            </div>

            {groupPermissionDenied && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  Group member expansion was skipped — the <strong>GroupMember.Read.All</strong> Graph
                  permission has not been approved in this tenant. A SharePoint or Global Administrator
                  can approve it in <strong>SharePoint Admin Center → Advanced → API access</strong>.
                </MessageBarBody>
              </MessageBar>
            )}

            <Divider />

            {/* Filter bar */}
            <Input
              placeholder="Filter by name, path, or user/group…"
              value={filterText}
              onChange={(_, d) => setFilterText(d.value)}
              style={{ width: '100%' }}
              aria-label="Filter results"
            />
            <div className={styles.row}>
              <Checkbox
                label="Unique permissions only"
                checked={filterUniqueOnly}
                onChange={(_, d) => setFilterUniqueOnly(!!d.checked)}
              />
              <Checkbox
                label="External users only (#ext#)"
                checked={filterExternalOnly}
                onChange={(_, d) => setFilterExternalOnly(!!d.checked)}
              />
              {(filterText || filterExternalOnly || filterUniqueOnly) && (
                <Body1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
                  Showing {filteredEntries?.length ?? 0} of {entries.length}
                </Body1>
              )}
            </div>

            <div className={styles.row}>
              <Button
                appearance="primary"
                icon={<DocumentArrowDown24Regular />}
                onClick={handleExport}
                disabled={isExporting || (filteredEntries?.length ?? 0) === 0}
              >
                {isExporting
                  ? 'Generating Excel…'
                  : filteredEntries && filteredEntries.length < entries.length
                  ? `Export ${filteredEntries.length} filtered rows`
                  : 'Export to Excel'}
              </Button>
              <Button
                appearance="secondary"
                icon={<DocumentArrowDown24Regular />}
                onClick={handleExportCsv}
                disabled={isExporting || (filteredEntries?.length ?? 0) === 0}
              >
                Export to CSV
              </Button>
            </div>
          </div>
        )}
      </div>}
    </div>
  );
};
