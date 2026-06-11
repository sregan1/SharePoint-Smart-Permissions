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
  ChevronRight16Regular,
  ChevronDown16Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import { ReportOptions, ReportScope, PermissionEntry, ObjectType, ScanProgress, StoredReport, LibraryInfo } from '../models/models';
import { requestNotificationPermission, showNotification } from '../utils/notifications';
import { SiteOwnersLinks } from './shared/SiteOwnersLinks';
import { PermTable } from './shared/PermTable';
import { diffReports, ReportDiff } from '../utils/reportDiff';

// Badge color per object type (matches the User Access view's mapping).
function typeBadgeColor(t: ObjectType): 'brand' | 'informative' | 'success' | 'warning' | undefined {
  switch (t) {
    case ObjectType.Site:    return 'brand';
    case ObjectType.Library: return 'informative';
    case ObjectType.List:    return 'success';
    case ObjectType.Folder:  return 'warning';
    default:                 return undefined;
  }
}

const TYPE_ORDER: Record<string, number> = {
  [ObjectType.Site]: 0,
  [ObjectType.Library]: 1,
  [ObjectType.List]: 2,
  [ObjectType.Folder]: 3,
  [ObjectType.File]: 4,
};


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
  const [includeSubsites, setIncludeSubsites] = React.useState(false);
  const [scope, setScope] = React.useState<string>('Site');
  const [folderDepth, setFolderDepth] = React.useState(2);
  const [expandGroups, setExpandGroups] = React.useState(true);

  // ── Run state ──
  const [isBusy, setIsBusy] = React.useState(false);
  const [scanProgress, setScanProgress] = React.useState<ScanProgress>({ message: '', scanned: 0, libsDone: 0, libsTotal: 0 });
  const [elapsed, setElapsed] = React.useState(0);
  const [error, setError] = React.useState('');
  const [entries, setEntries] = React.useState<PermissionEntry[] | null>(null);
  const [cancelled, setCancelled] = React.useState(false);
  const [groupPermissionDenied, setGroupPermissionDenied] = React.useState(false);
  const [roleAssignmentsDenied, setRoleAssignmentsDenied] = React.useState(false);
  const [siteOwners, setSiteOwners] = React.useState<{ title: string; email: string }[]>([]);
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

  // ── Results table state ──
  const RESULTS_PAGE_SIZE = 200;
  const [resultsVisible, setResultsVisible] = React.useState(RESULTS_PAGE_SIZE);
  // 'scan' keeps the natural scan order (site → library → folders, DFS).
  const [resultSortCol, setResultSortCol] = React.useState<'scan' | 'type' | 'name' | 'path' | 'source'>('scan');
  const [resultSortAsc, setResultSortAsc] = React.useState(true);
  const [expandedKeys, setExpandedKeys] = React.useState<Set<string>>(new Set());

  React.useEffect(() => {
    setResultsVisible(RESULTS_PAGE_SIZE);
    setExpandedKeys(new Set());
  }, [entries, filterText, filterExternalOnly, filterUniqueOnly]);

  const sortedResults = React.useMemo(() => {
    if (!filteredEntries) return [];
    if (resultSortCol === 'scan') return filteredEntries;
    return [...filteredEntries].sort((a, b) => {
      let diff = 0;
      if (resultSortCol === 'type') diff = (TYPE_ORDER[a.objectType] ?? 5) - (TYPE_ORDER[b.objectType] ?? 5);
      else if (resultSortCol === 'name') diff = a.name.localeCompare(b.name);
      else if (resultSortCol === 'path') diff = a.serverRelativeUrl.localeCompare(b.serverRelativeUrl);
      else diff = Number(b.hasUniquePermissions) - Number(a.hasUniquePermissions);
      if (diff !== 0) return resultSortAsc ? diff : -diff;
      return a.serverRelativeUrl.localeCompare(b.serverRelativeUrl);
    });
  }, [filteredEntries, resultSortCol, resultSortAsc]);

  const handleResultSort = (col: 'type' | 'name' | 'path' | 'source'): void => {
    if (resultSortCol === col) { setResultSortAsc((v) => !v); } else { setResultSortCol(col); setResultSortAsc(true); }
  };

  const resultSortInd = (col: string): string =>
    resultSortCol !== col ? '' : resultSortAsc ? ' ▲' : ' ▼';

  const entryKey = (e: PermissionEntry): string => `${e.objectType}|${e.siteUrl}|${e.serverRelativeUrl}`;

  const toggleExpanded = (key: string): void => {
    setExpandedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) { next.delete(key); } else { next.add(key); }
      return next;
    });
  };

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

  // ── Compare state ──
  const [compareSelection, setCompareSelection] = React.useState<Set<string>>(new Set());
  const [compareResult, setCompareResult] = React.useState<{ older: StoredReport; newer: StoredReport; diff: ReportDiff } | null>(null);

  const toggleCompareSelection = (id: string): void => {
    setCompareSelection((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else if (next.size < 2) {
        next.add(id);
      }
      return next;
    });
  };

  const handleCompare = (): void => {
    const selected = historyItems.filter((r) => compareSelection.has(r.id));
    if (selected.length !== 2) return;
    // Older report first — ids are Date.now() strings, but compare timestamps
    // to be safe.
    const [older, newer] = selected[0].timestamp <= selected[1].timestamp
      ? [selected[0], selected[1]]
      : [selected[1], selected[0]];
    setCompareResult({ older, newer, diff: diffReports(older, newer) });
  };

  const compareMismatch = compareResult !== null && (
    compareResult.older.siteUrl !== compareResult.newer.siteUrl ||
    compareResult.older.options.scope !== compareResult.newer.options.scope ||
    !!compareResult.older.options.allSites !== !!compareResult.newer.options.allSites ||
    !!compareResult.older.options.includeSubsites !== !!compareResult.newer.options.includeSubsites
  );

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

  // Load history on mount
  React.useEffect(() => {
    historyService.current.getAll()
      .then(setHistoryItems)
      .catch(() => { /* IndexedDB unavailable — history simply won't show */ });
  }, []);

  // Clear stale results when any option that affects output changes
  React.useEffect(() => {
    setEntries(null);
    setRoleAssignmentsDenied(false);
    setSiteOwners([]);
  }, [allSites, includeSubsites, scope, folderDepth, expandGroups]);

  React.useEffect(() => {
    if (!roleAssignmentsDenied || !siteUrl) return;
    sp.getSiteOwners(siteUrl.trim()).then(setSiteOwners).catch(() => {});
  }, [roleAssignmentsDenied, siteUrl]);

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
    requestNotificationPermission();
    abortRef.current?.abort();
    abortRef.current = new AbortController();
    scanStartRef.current = Date.now();
    setIsBusy(true);
    setError('');
    setEntries(null);
    setCancelled(false);
    setRoleAssignmentsDenied(false);
    setSiteOwners([]);
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
        includeSubsites,
        scope: scope as ReportScope,
        folderDepth,
        includeHidden,
        expandGroups,
        libraryUrls: allSelected ? undefined : Array.from(selectedLibraryUrls),
      };

      const { entries: scannedEntries, groupPermissionDenied: permDenied, roleAssignmentsDenied: raDenied } = await sp.scanPermissions(
        options,
        (progress) => setScanProgress(progress),
        abortRef.current.signal,
        () => { liveCountRef.current += 1; },
      );

      // On cancel, keep whatever was collected so it can be reviewed/exported.
      const wasCancelled = abortRef.current.signal.aborted;
      setCancelled(wasCancelled);
      setEntries(scannedEntries);
      setGroupPermissionDenied(permDenied);
      setRoleAssignmentsDenied(raDenied);
      const uniqueCount = scannedEntries.filter((e) => e.hasUniquePermissions).length;
      setScanProgress((prev) => ({
        ...prev,
        message: wasCancelled
          ? `Scan cancelled — ${scannedEntries.length} object(s) collected before cancelling.`
          : `Scan complete — ${scannedEntries.length} object(s) found, ` +
            `${uniqueCount} with unique permissions.`,
      }));

      if (wasCancelled) return;

      showNotification(
        'Smart Permissions — Scan complete',
        `${scannedEntries.length} objects, ${uniqueCount} with unique permissions.`,
      );

      // Save to history (errors are swallowed — never block the user from seeing results)
      const storedReport: StoredReport = {
        id: Date.now().toString(),
        timestamp: new Date().toISOString(),
        siteUrl: siteUrl.trim(),
        options: { allSites, includeSubsites, scope: scope as ReportScope, folderDepth, expandGroups },
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

      {/* ── Compare (diff) panel ── */}
      {showHistory && compareResult && (
        <div>
          <div className={styles.row} style={{ marginBottom: tokens.spacingVerticalM }}>
            <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={() => setCompareResult(null)}>
              Back to history
            </Button>
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
              Comparing {new Date(compareResult.older.timestamp).toLocaleString()} →{' '}
              {new Date(compareResult.newer.timestamp).toLocaleString()}
            </Body1>
          </div>

          {compareMismatch && (
            <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                These reports were taken with different sites or scan options — differences below
                may reflect the changed scan settings rather than actual permission changes.
              </MessageBarBody>
            </MessageBar>
          )}

          {compareResult.diff.isEmpty ? (
            <MessageBar intent="success">
              <MessageBarBody>No permission differences found between these two reports.</MessageBarBody>
            </MessageBar>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL }}>
              {compareResult.diff.permissionChanges.length > 0 && (
                <div>
                  <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalS }}>
                    Permission changes ({compareResult.diff.permissionChanges.length} object{compareResult.diff.permissionChanges.length !== 1 ? 's' : ''})
                  </Text>
                  {compareResult.diff.permissionChanges.map((obj) => (
                    <div key={`${obj.objectType}|${obj.serverRelativeUrl}`} style={{ marginBottom: tokens.spacingVerticalM }}>
                      <div className={styles.row}>
                        <Badge appearance="filled" color={typeBadgeColor(obj.objectType as ObjectType)} size="small">{obj.objectType}</Badge>
                        <Text weight="semibold">{obj.name}</Text>
                        <Text style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, wordBreak: 'break-all' }}>
                          {obj.serverRelativeUrl}
                        </Text>
                      </div>
                      <table className={styles.historyTable}>
                        <thead>
                          <tr>
                            <th className={styles.historyTh}>Change</th>
                            <th className={styles.historyTh}>User / Group</th>
                            <th className={styles.historyTh}>Before</th>
                            <th className={styles.historyTh}>After</th>
                          </tr>
                        </thead>
                        <tbody>
                          {obj.added.map((c, i) => (
                            <tr key={`a${i}`}>
                              <td className={styles.historyTd}><Badge appearance="filled" color="success" size="small">Added</Badge></td>
                              <td className={styles.historyTd}>{c.displayName || c.loginName}</td>
                              <td className={styles.historyTd}>—</td>
                              <td className={styles.historyTd}>{c.newRoles}</td>
                            </tr>
                          ))}
                          {obj.removed.map((c, i) => (
                            <tr key={`r${i}`}>
                              <td className={styles.historyTd}><Badge appearance="filled" color="danger" size="small">Removed</Badge></td>
                              <td className={styles.historyTd}>{c.displayName || c.loginName}</td>
                              <td className={styles.historyTd}>{c.oldRoles}</td>
                              <td className={styles.historyTd}>—</td>
                            </tr>
                          ))}
                          {obj.changed.map((c, i) => (
                            <tr key={`c${i}`}>
                              <td className={styles.historyTd}><Badge appearance="filled" color="warning" size="small">Changed</Badge></td>
                              <td className={styles.historyTd}>{c.displayName || c.loginName}</td>
                              <td className={styles.historyTd}>{c.oldRoles}</td>
                              <td className={styles.historyTd}>{c.newRoles}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ))}
                </div>
              )}

              {compareResult.diff.inheritanceChanged.length > 0 && (
                <div>
                  <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalS }}>
                    Inheritance changed ({compareResult.diff.inheritanceChanged.length})
                  </Text>
                  <table className={styles.historyTable}>
                    <thead>
                      <tr>
                        <th className={styles.historyTh}>Type</th>
                        <th className={styles.historyTh}>Name</th>
                        <th className={styles.historyTh}>Path</th>
                        <th className={styles.historyTh}>Now</th>
                      </tr>
                    </thead>
                    <tbody>
                      {compareResult.diff.inheritanceChanged.map(({ entry, nowUnique }) => (
                        <tr key={`${entry.objectType}|${entry.serverRelativeUrl}`}>
                          <td className={styles.historyTd}>
                            <Badge appearance="filled" color={typeBadgeColor(entry.objectType)} size="small">{entry.objectType}</Badge>
                          </td>
                          <td className={styles.historyTd}>{entry.name}</td>
                          <td className={styles.historyTd}>
                            <Text style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, wordBreak: 'break-all' }}>
                              {entry.serverRelativeUrl}
                            </Text>
                          </td>
                          <td className={styles.historyTd}>
                            {nowUnique
                              ? <Badge appearance="filled" color="warning" size="small">Unique (inheritance broken)</Badge>
                              : <Badge appearance="outline" size="small">Inherited (restored)</Badge>}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {(['addedObjects', 'removedObjects'] as const).map((bucket) => (
                compareResult.diff[bucket].length > 0 && (
                  <div key={bucket}>
                    <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalS }}>
                      {bucket === 'addedObjects' ? 'New objects' : 'Removed objects'} ({compareResult.diff[bucket].length})
                    </Text>
                    <table className={styles.historyTable}>
                      <thead>
                        <tr>
                          <th className={styles.historyTh}>Type</th>
                          <th className={styles.historyTh}>Name</th>
                          <th className={styles.historyTh}>Path</th>
                        </tr>
                      </thead>
                      <tbody>
                        {compareResult.diff[bucket].map((entry) => (
                          <tr key={`${entry.objectType}|${entry.serverRelativeUrl}`}>
                            <td className={styles.historyTd}>
                              <Badge appearance="filled" color={typeBadgeColor(entry.objectType)} size="small">{entry.objectType}</Badge>
                            </td>
                            <td className={styles.historyTd}>{entry.name}</td>
                            <td className={styles.historyTd}>
                              <Text style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, wordBreak: 'break-all' }}>
                                {entry.serverRelativeUrl}
                              </Text>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )
              ))}
            </div>
          )}
        </div>
      )}

      {/* ── History panel ── */}
      {showHistory && !compareResult && (
        <div>
          <div className={styles.row} style={{ marginBottom: tokens.spacingVerticalM }}>
            <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={() => setShowHistory(false)}>
              Back to scan
            </Button>
            <Button
              appearance="primary"
              onClick={handleCompare}
              disabled={compareSelection.size !== 2}
              style={{ marginLeft: 'auto' }}
            >
              Compare selected{compareSelection.size > 0 ? ` (${compareSelection.size}/2)` : ''}
            </Button>
          </div>

          {historyItems.length === 0 ? (
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No reports saved yet.</Body1>
          ) : (
            <table className={styles.historyTable}>
              <thead>
                <tr>
                  <th className={styles.historyTh} style={{ width: '32px' }} aria-label="Select for compare" />
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
                    <td className={styles.historyTd}>
                      <Checkbox
                        checked={compareSelection.has(item.id)}
                        onChange={() => toggleCompareSelection(item.id)}
                        disabled={!compareSelection.has(item.id) && compareSelection.size >= 2}
                        aria-label={`Select report from ${new Date(item.timestamp).toLocaleString()} for compare`}
                      />
                    </td>
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
                      {item.options.includeSubsites ? ' + subsites' : ''}
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

        <Checkbox
          label="Include subsites (scans every subsite below this site)"
          checked={includeSubsites}
          onChange={(_, d) => setIncludeSubsites(!!d.checked)}
          disabled={isBusy}
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
              <Text weight="semibold">{cancelled ? 'Scan cancelled — partial results' : 'Scan complete'}</Text>
              <Badge appearance="filled" color="success">{entries.length} objects</Badge>
              <Badge appearance="filled" color="warning">
                {entries.filter((e) => e.hasUniquePermissions).length} unique
              </Badge>
              <Badge appearance="outline">
                {entries.filter((e) => !e.hasUniquePermissions).length} inherited
              </Badge>
            </div>

            {roleAssignmentsDenied && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  This scan ran with Member access — permission assignments could not be read.
                  Only items with unique permissions are shown; who has access to each item
                  is not visible. Run the scan as a <strong>Site Owner</strong> to see full permission details.
                  <SiteOwnersLinks owners={siteOwners} />
                </MessageBarBody>
              </MessageBar>
            )}

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

            {/* ── Results table ── */}
            {sortedResults.length > 0 && (
              <>
                <table className={styles.historyTable} aria-label="Scan results">
                  <thead>
                    <tr>
                      <th className={styles.historyTh} style={{ width: '28px' }} />
                      {(
                        [
                          { col: 'type', label: 'Type' },
                          { col: 'name', label: 'Name' },
                          { col: 'path', label: 'Path' },
                          { col: 'source', label: 'Permissions' },
                        ] as { col: 'type' | 'name' | 'path' | 'source'; label: string }[]
                      ).map(({ col, label }) => (
                        <th
                          key={col}
                          className={styles.historyTh}
                          onClick={() => handleResultSort(col)}
                          style={{ cursor: 'pointer', userSelect: 'none' }}
                        >
                          {label}{resultSortInd(col)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {sortedResults.slice(0, resultsVisible).map((entry) => {
                      const key = entryKey(entry);
                      const isExpanded = expandedKeys.has(key);
                      const expandable = entry.uniquePermissions.length > 0;
                      return (
                        <React.Fragment key={key}>
                          <tr
                            onClick={expandable ? () => toggleExpanded(key) : undefined}
                            style={expandable ? { cursor: 'pointer' } : undefined}
                          >
                            <td className={styles.historyTd}>
                              {expandable && (
                                <Button
                                  appearance="transparent"
                                  size="small"
                                  icon={isExpanded ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
                                  aria-expanded={isExpanded}
                                  aria-label={isExpanded ? 'Collapse permissions' : 'Expand permissions'}
                                  onClick={(e) => { e.stopPropagation(); toggleExpanded(key); }}
                                />
                              )}
                            </td>
                            <td className={styles.historyTd}>
                              <Badge appearance="filled" color={typeBadgeColor(entry.objectType)} size="small">
                                {entry.objectType}
                              </Badge>
                            </td>
                            <td className={styles.historyTd} style={{ paddingLeft: `${8 + entry.depth * 12}px` }}>
                              {entry.name}
                              {entry.noCrawl && (
                                <Badge appearance="outline" size="small" style={{ marginLeft: '6px' }}>
                                  Hidden from search
                                </Badge>
                              )}
                            </td>
                            <td className={styles.historyTd}>
                              <Text style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, wordBreak: 'break-all' }}>
                                {entry.serverRelativeUrl}
                              </Text>
                            </td>
                            <td className={styles.historyTd} style={{ whiteSpace: 'nowrap' }}>
                              {entry.hasUniquePermissions ? (
                                <Badge appearance="filled" color="warning" size="small">Unique</Badge>
                              ) : (
                                <Badge appearance="outline" size="small">Inherited</Badge>
                              )}
                              {expandable && (
                                <Text style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, marginLeft: '6px' }}>
                                  {entry.uniquePermissions.length} assignment{entry.uniquePermissions.length !== 1 ? 's' : ''}
                                </Text>
                              )}
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr>
                              <td className={styles.historyTd} />
                              <td className={styles.historyTd} colSpan={4} style={{ background: tokens.colorNeutralBackground2 }}>
                                <PermTable users={entry.uniquePermissions} />
                              </td>
                            </tr>
                          )}
                        </React.Fragment>
                      );
                    })}
                  </tbody>
                </table>
                {resultsVisible < sortedResults.length && (
                  <div style={{ textAlign: 'center' }}>
                    <Button appearance="secondary" onClick={() => setResultsVisible((c) => c + RESULTS_PAGE_SIZE)}>
                      Load more ({(sortedResults.length - resultsVisible).toLocaleString()} remaining)
                    </Button>
                  </div>
                )}
              </>
            )}

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
