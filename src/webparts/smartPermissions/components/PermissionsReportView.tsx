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

import * as strings from 'SmartPermissionsWebPartStrings';
import { formatString } from '../utils/localeUtils';
import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportHistoryService } from '../services/ReportHistoryService';
import { ReportOptions, ReportScope, PermissionEntry, ObjectType, ScanProgress, StoredReport, LibraryInfo } from '../models/models';
import { requestNotificationPermission, showNotification } from '../utils/notifications';
import { SiteOwnersLinks } from './shared/SiteOwnersLinks';
import { PermTable } from './shared/PermTable';
import { applyPermFilters } from './shared/permFilters';
import { isSharingLinkPrincipal } from './shared/sharingLinks';
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
  onExcludeLimitedAccessChange: (val: boolean) => void;
  excludeSharingLinks: boolean;
  onExcludeSharingLinksChange: (val: boolean) => void;
  onBack: () => void;
}

export const PermissionsReportView: React.FC<PermissionsReportViewProps> = ({
  sp,
  excel,
  siteUrl,
  includeHidden,
  excludeLimitedAccess,
  onExcludeLimitedAccessChange,
  excludeSharingLinks,
  onExcludeSharingLinksChange,
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
  const [deniedPaths, setDeniedPaths] = React.useState<string[]>([]);
  const [throttleEvents, setThrottleEvents] = React.useState(0);
  const [throttleAborted, setThrottleAborted] = React.useState(false);
  const [siteProgress, setSiteProgress] = React.useState<{ scanned: number; total: number } | null>(null);
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
      if (excludeSharingLinks && !e.uniquePermissions.some((u) => !isSharingLinkPrincipal(u))) return false;
      if (!lc) return true;
      if (e.name.toLowerCase().includes(lc)) return true;
      if (e.serverRelativeUrl.toLowerCase().includes(lc)) return true;
      if (e.uniquePermissions.some((u) => u.displayName.toLowerCase().includes(lc))) return true;
      return false;
    });
  }, [entries, filterText, filterExternalOnly, filterUniqueOnly, excludeLimitedAccess, excludeSharingLinks]);

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
  }, [entries, filterText, filterExternalOnly, filterUniqueOnly, excludeLimitedAccess, excludeSharingLinks]);

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
    setThrottleEvents(0);
    setThrottleAborted(false);
    setSiteProgress(null);
    // A previous run may have left the circuit breaker open; a fresh scan is an
    // explicit "try again", so clear it rather than failing every request.
    sp.resetThrottleState();
    liveCountRef.current = 0;
    setScanProgress({ message: strings.StartingScanStatus, scanned: 0, libsDone: 0, libsTotal: 0 });

    // Flush the live item count to state every 500ms so React batches renders
    const flushTimer = setInterval(() => setLiveCount(liveCountRef.current), 500);
    // The client's throttle counter is cumulative for the session; baseline it
    // so the banner reports throttling from this scan only.
    const throttleAtStart = sp.throttleEventCount;

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

      const {
        entries: scannedEntries,
        groupPermissionDenied: permDenied,
        roleAssignmentsDenied: raDenied,
        deniedPaths: raDeniedPaths,
        throttleAborted: wasThrottleAborted,
        sitesTotal,
        sitesScanned,
      } = await sp.scanPermissions(
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
      setDeniedPaths(raDeniedPaths);
      setThrottleEvents(sp.throttleEventCount - throttleAtStart);
      setThrottleAborted(wasThrottleAborted);
      setSiteProgress(allSites ? { scanned: sitesScanned, total: sitesTotal } : null);
      const uniqueCount = scannedEntries.filter((e) => e.hasUniquePermissions).length;
      setScanProgress((prev) => ({
        ...prev,
        message: wasCancelled
          ? formatString(strings.ScanCancelledStatus, scannedEntries.length)
          : formatString(strings.ScanCompleteStatus, scannedEntries.length, uniqueCount),
      }));

      if (wasCancelled) return;

      showNotification(
        strings.ScanCompleteNotificationTitle,
        formatString(strings.ScanCompleteNotificationBody, scannedEntries.length, uniqueCount),
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
        setScanProgress((prev) => ({ ...prev, message: strings.CancelledStatus }));
      } else {
        setError(formatString(strings.GenericErrorPrefix, err?.message ?? String(err)));
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

  // Mirrors the same excludeLimitedAccess/filterExternalOnly filtering the Explorer
  // and User Access views apply at the user level (shared applyPermFilters) —
  // previously this only trimmed external users, so the exported workbook still
  // listed limited-access principals the on-screen filter had hidden. Rows whose
  // principals are entirely filtered out by these trims are dropped rather than
  // exported as an empty-looking row.
  const applyExportFilters = (entriesToFilter: PermissionEntry[]): PermissionEntry[] => {
    let result = filterUniqueOnly
      ? entriesToFilter.filter((e) => e.hasUniquePermissions)
      : entriesToFilter;
    const trimming = filterExternalOnly || excludeLimitedAccess || excludeSharingLinks;
    if (trimming) {
      result = result
        .map((e) => ({
          ...e,
          uniquePermissions: applyPermFilters(e.uniquePermissions, excludeLimitedAccess, filterExternalOnly, excludeSharingLinks),
        }))
        .filter((e) => e.uniquePermissions.length > 0);
    }
    return result;
  };

  // Memoized so the export button's enabled state and label reflect exactly
  // what handleExport/handleExportCsv will actually produce (see A13) — the
  // disabled-guard previously checked filteredEntries.length, which could be
  // non-zero even when applyExportFilters trimmed every row's principals away.
  const exportableEntries = React.useMemo(
    () => applyExportFilters(filteredEntries ?? entries ?? []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [filteredEntries, entries, filterUniqueOnly, filterExternalOnly, excludeLimitedAccess, excludeSharingLinks],
  );

  // Count of entries whose permissions couldn't be confirmed after retries
  // (scanIncomplete), so the throttling banner can name an exact number
  // instead of a vague "some items" — the user needs to know how much of the
  // report to distrust, not just that throttling happened at all.
  const incompleteCount = React.useMemo(
    () => (entries ?? []).filter((e) => e.scanIncomplete).length,
    [entries],
  );

  const handleExport = async (): Promise<void> => {
    if (exportableEntries.length === 0) return;
    setIsExporting(true);
    try {
      await excel.export(exportableEntries, siteUrl.trim(), allSites);
    } catch (err: any) {
      setError(formatString(strings.ExportErrorPrefix, err?.message ?? String(err)));
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportCsv = (): void => {
    excel.exportPermissionsCsv(exportableEntries, siteUrl.trim(), allSites);
  };

  const handleHistoryExport = async (item: StoredReport): Promise<void> => {
    setExportingHistoryId(item.id);
    try {
      await excel.export(item.entries, item.siteUrl, item.options.allSites);
    } catch (err: any) {
      setError(formatString(strings.ExportErrorPrefix, err?.message ?? String(err)));
    } finally {
      setExportingHistoryId(null);
    }
  };

  const handleHistoryDelete = async (id: string): Promise<void> => {
    await historyService.current.delete(id).catch(() => { /* ignore */ });
    setHistoryItems((prev) => prev.filter((r) => r.id !== id));
  };

  const scopeLabel = (s: string): string =>
    ({
      Site: strings.ScopeSiteOnly,
      Library: strings.ScopeLibraries,
      Folder: strings.ScopeFolders,
      Item: strings.ScopeFilesAndFolders,
    } as Record<string, string>)[s] ?? s;

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
            if (isBusy && !window.confirm(strings.ScanInProgressConfirm)) return;
            onBack();
          }}
          disabled={false}
          aria-label={strings.BackToHomeLabel}
        >
          {strings.BackButton}
        </Button>
        <Title3 style={{ flex: 1 }}>{strings.ReportCardTitle}</Title3>
        <Button
          appearance="subtle"
          icon={<History24Regular />}
          onClick={() => setShowHistory((v) => !v)}
          disabled={isBusy}
        >
          {strings.HistoryButton}{historyItems.length > 0 ? ` (${historyItems.length})` : ''}
        </Button>
      </div>

      {/* ── Compare (diff) panel ── */}
      {showHistory && compareResult && (
        <div>
          <div className={styles.row} style={{ marginBottom: tokens.spacingVerticalM }}>
            <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={() => setCompareResult(null)}>
              {strings.BackToHistoryButton}
            </Button>
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
              {strings.ComparingLabel} {new Date(compareResult.older.timestamp).toLocaleString()} →{' '}
              {new Date(compareResult.newer.timestamp).toLocaleString()}
            </Body1>
          </div>

          {compareMismatch && (
            <MessageBar intent="warning" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                {strings.CompareMismatchWarning}
              </MessageBarBody>
            </MessageBar>
          )}

          {compareResult.diff.isEmpty ? (
            <MessageBar intent="success">
              <MessageBarBody>{strings.NoPermissionDifferencesFound}</MessageBarBody>
            </MessageBar>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL }}>
              {compareResult.diff.permissionChanges.length > 0 && (
                <div>
                  <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalS }}>
                    {formatString(
                      compareResult.diff.permissionChanges.length === 1 ? strings.PermissionChangesHeaderSingular : strings.PermissionChangesHeaderPlural,
                      compareResult.diff.permissionChanges.length,
                    )}
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
                            <th className={styles.historyTh}>{strings.ChangeColumnHeader}</th>
                            <th className={styles.historyTh}>{strings.UserGroupColumnHeader}</th>
                            <th className={styles.historyTh}>{strings.BeforeColumnHeader}</th>
                            <th className={styles.historyTh}>{strings.AfterColumnHeader}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {obj.added.map((c, i) => (
                            <tr key={`a-${c.loginName || c.displayName}-${i}`}>
                              <td className={styles.historyTd}><Badge appearance="filled" color="success" size="small">{strings.AddedBadge}</Badge></td>
                              <td className={styles.historyTd}>{c.displayName || c.loginName}</td>
                              <td className={styles.historyTd}>—</td>
                              <td className={styles.historyTd}>{c.newRoles}</td>
                            </tr>
                          ))}
                          {obj.removed.map((c, i) => (
                            <tr key={`r-${c.loginName || c.displayName}-${i}`}>
                              <td className={styles.historyTd}><Badge appearance="filled" color="danger" size="small">{strings.RemovedBadge}</Badge></td>
                              <td className={styles.historyTd}>{c.displayName || c.loginName}</td>
                              <td className={styles.historyTd}>{c.oldRoles}</td>
                              <td className={styles.historyTd}>—</td>
                            </tr>
                          ))}
                          {obj.changed.map((c, i) => (
                            <tr key={`c-${c.loginName || c.displayName}-${i}`}>
                              <td className={styles.historyTd}><Badge appearance="filled" color="warning" size="small">{strings.ChangedBadge}</Badge></td>
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
                    {formatString(strings.InheritanceChangedHeader, compareResult.diff.inheritanceChanged.length)}
                  </Text>
                  <table className={styles.historyTable}>
                    <thead>
                      <tr>
                        <th className={styles.historyTh}>{strings.TypeColumnHeader}</th>
                        <th className={styles.historyTh}>{strings.NameColumnHeader}</th>
                        <th className={styles.historyTh}>{strings.PathColumnHeader}</th>
                        <th className={styles.historyTh}>{strings.NowColumnHeader}</th>
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
                              ? <Badge appearance="filled" color="warning" size="small">{strings.UniqueInheritanceBrokenBadge}</Badge>
                              : <Badge appearance="outline" size="small">{strings.InheritedRestoredBadge}</Badge>}
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
                      {bucket === 'addedObjects' ? strings.NewObjectsHeader : strings.RemovedObjectsHeader} ({compareResult.diff[bucket].length})
                    </Text>
                    <table className={styles.historyTable}>
                      <thead>
                        <tr>
                          <th className={styles.historyTh}>{strings.TypeColumnHeader}</th>
                          <th className={styles.historyTh}>{strings.NameColumnHeader}</th>
                          <th className={styles.historyTh}>{strings.PathColumnHeader}</th>
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
              {strings.BackToScanButton}
            </Button>
            <Button
              appearance="primary"
              onClick={handleCompare}
              disabled={compareSelection.size !== 2}
              style={{ marginLeft: 'auto' }}
            >
              {strings.CompareSelectedButton}{compareSelection.size > 0 ? ` (${compareSelection.size}/2)` : ''}
            </Button>
          </div>

          {historyItems.length === 0 ? (
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>{strings.NoReportsSavedYet}</Body1>
          ) : (
            <table className={styles.historyTable}>
              <thead>
                <tr>
                  <th className={styles.historyTh} style={{ width: '32px' }} aria-label={strings.SelectForCompareLabel} />
                  {(
                    [
                      { col: 'timestamp', label: strings.DateTimeColumnHeader },
                      { col: 'siteUrl', label: strings.SiteColumnHeader },
                      { col: 'scope', label: strings.ScopeColumnHeader },
                      { col: 'total', label: strings.ObjectsColumnHeader },
                      { col: 'unique', label: strings.UniqueColumnHeader },
                    ] as { col: typeof historySortCol; label: string }[]
                  ).map(({ col, label }) => (
                    <th
                      key={col}
                      className={styles.historyTh}
                      onClick={() => handleHistorySort(col)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleHistorySort(col); }
                      }}
                      tabIndex={0}
                      role="columnheader"
                      aria-sort={historySortCol !== col ? 'none' : historySortAsc ? 'ascending' : 'descending'}
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
                        aria-label={formatString(strings.SelectReportForCompareLabel, new Date(item.timestamp).toLocaleString())}
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
                      {item.options.allSites ? strings.AllSitesPrefix : ''}{scopeLabel(item.options.scope)}
                      {item.options.includeSubsites ? strings.PlusSubsitesSuffix : ''}
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
                          {exportingHistoryId === item.id ? strings.ExportingStatus : strings.ExportButton}
                        </Button>
                        <Button
                          size="small"
                          appearance="subtle"
                          icon={<Delete24Regular />}
                          onClick={() => handleHistoryDelete(item.id)}
                          title={strings.DeleteThisReportTitle}
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
          label={strings.ScanAllSiteCollectionsLabel}
          checked={allSites}
          onChange={(_, d) => setAllSites(!!d.checked)}
          disabled={!isRootSite || isBusy}
        />
        {allSites && (
          <Text size={200} style={{ color: tokens.colorNeutralForeground3, marginLeft: '28px' }}>
            {strings.AllSitesSearchCaveat}
          </Text>
        )}

        <Checkbox
          label={strings.IncludeSubsitesLabel}
          checked={includeSubsites}
          onChange={(_, d) => setIncludeSubsites(!!d.checked)}
          disabled={isBusy}
        />

        <Divider />

        {/* Scope */}
        <Field label={strings.ScanDepthFieldLabel}>
          <RadioGroup
            value={scope}
            onChange={(_, d) => setScope(d.value)}
            layout="horizontal"
            disabled={isBusy}
            style={{ flexWrap: 'wrap', gap: tokens.spacingHorizontalS }}
          >
            {([
              { value: 'Site', Icon: Globe24Regular, label: strings.ScopeSiteOnly },
              { value: 'Library', Icon: BookDatabase24Regular, label: strings.ScopeLibraries },
              { value: 'Folder', Icon: Folder24Regular, label: strings.ScopeFolders },
              { value: 'Item', Icon: FolderOpen24Regular, label: strings.ScopeFilesAndFolders },
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
          <Label>{strings.FolderDepthLimitLabel}</Label>
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
          label={strings.ExpandGroupMembersInReportLabel}
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
                {strings.LibrariesToScanLabel}
                {selectedLibraryUrls.size < availableLibraries.length && (
                  <span style={{ color: tokens.colorNeutralForeground3, fontWeight: 'normal', marginLeft: '6px' }}>
                    ({formatString(strings.OfSelectedCount, selectedLibraryUrls.size, availableLibraries.length)})
                  </span>
                )}
              </Label>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => setSelectedLibraryUrls(new Set(availableLibraries.map((l) => l.serverRelativeUrl)))}
                disabled={isBusy || selectedLibraryUrls.size === availableLibraries.length}
              >
                {strings.AllButton}
              </Button>
              <Button
                size="small"
                appearance="subtle"
                onClick={() => setSelectedLibraryUrls(new Set())}
                disabled={isBusy || selectedLibraryUrls.size === 0}
              >
                {strings.NoneButton}
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
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>{strings.LoadingLibrariesStatus}</Body1>
        )}

        <Divider />

        {/* Action buttons */}
        <div className={styles.row}>
          <Button
            appearance="primary"
            onClick={handleRun}
            disabled={isBusy || (scope !== 'Site' && availableLibraries.length > 0 && selectedLibraryUrls.size === 0)}
          >
            {strings.RunReportButton}
          </Button>
          {isBusy && (
            <Button appearance="secondary" onClick={handleCancel}>
              {strings.CancelButton}
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
                {formatString(strings.ItemsFoundSoFar, liveCount.toLocaleString())}
                {scanProgress.libsTotal > 0 && ` · ${formatString(strings.LibraryOfProgress, scanProgress.libsDone, scanProgress.libsTotal)}`}
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
              <Text weight="semibold">{cancelled ? strings.ScanCancelledPartialResults : strings.ScanCompleteLabel}</Text>
              <Badge appearance="filled" color="success">{formatString(strings.ObjectsCountBadge, entries.length)}</Badge>
              <Badge appearance="filled" color="warning">
                {formatString(strings.UniqueCountBadge, entries.filter((e) => e.hasUniquePermissions).length)}
              </Badge>
              <Badge appearance="outline">
                {formatString(strings.InheritedCountBadge, entries.filter((e) => !e.hasUniquePermissions).length)}
              </Badge>
            </div>

            {/* Export buttons live here (top of the results) as well as at the
                bottom of the table — with potentially thousands of rows, requiring
                a full scroll just to export defeats the point of a quick export. */}
            <div className={styles.row}>
              <Button
                appearance="primary"
                icon={<DocumentArrowDown24Regular />}
                onClick={handleExport}
                disabled={isExporting || exportableEntries.length === 0}
              >
                {isExporting
                  ? strings.GeneratingExcelStatus
                  : exportableEntries.length < entries.length
                  ? formatString(strings.ExportFilteredRowsButton, exportableEntries.length)
                  : strings.ExportToExcelButton}
              </Button>
              <Button
                appearance="secondary"
                icon={<DocumentArrowDown24Regular />}
                onClick={handleExportCsv}
                disabled={isExporting || exportableEntries.length === 0}
              >
                {strings.ExportToCsvButton}
              </Button>
            </div>

            {throttleAborted && (
              <MessageBar intent="error">
                <MessageBarBody>
                  <strong>{strings.ScanStoppedEarlyTitle}</strong>{' '}
                  {siteProgress
                    ? formatString(strings.SiteCollectionsScannedBeforeStopping, siteProgress.scanned, siteProgress.total)
                    : ''}
                  {strings.ResultsBelowCoverOnlyPre} <strong>{strings.NotWord}</strong> {strings.ResultsBelowCoverOnlyPost}{' '}
                  <strong>{strings.ConcurrentRequestsWord}</strong> {strings.SetTo1Or2InSettings}
                </MessageBarBody>
              </MessageBar>
            )}

            {throttleEvents > 0 && !throttleAborted && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  {formatString(strings.ThrottledScanMessage, throttleEvents)}{' '}
                  {incompleteCount > 0 ? (
                    <>
                      {strings.OfEverythingScannedPre} <strong>{formatString(incompleteCount === 1 ? strings.ItemCountSingular : strings.ItemCountPlural, incompleteCount)}</strong>{' '}
                      {strings.CouldNotBeConfirmedText}
                    </>
                  ) : (
                    <>
                      {strings.DespiteThrottlingText}
                    </>
                  )}{' '}
                  {strings.ForCleanRunPre}{' '}
                  <strong>{strings.ConcurrentRequestsWord}</strong> {strings.ValueInSettingsPost}
                </MessageBarBody>
              </MessageBar>
            )}

            {roleAssignmentsDenied && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  {deniedPaths.length === 1
                    ? strings.OneItemWord
                    : formatString(strings.NItemsWord, deniedPaths.length)} {formatString(strings.OutOfScannedCouldNotBeRead, entries.length)}
                  {deniedPaths.length === 1 ? strings.PermissionAssignmentsForItAreNotShown : strings.PermissionAssignmentsForThoseAreNotShown}
                  {strings.RestOfReportReadSuccessfully}
                  {strings.RunScanAsSiteOwnerPre} <strong>{strings.SiteOwnerWord}</strong> {strings.RunScanAsSiteOwnerPost}
                  <SiteOwnersLinks owners={siteOwners} />
                  {deniedPaths.length > 0 && (
                    <ul style={{ margin: '4px 0 0', paddingLeft: '20px' }}>
                      {deniedPaths.slice(0, 10).map((path) => (
                        <li key={path}>{path}</li>
                      ))}
                      {deniedPaths.length > 10 && <li>{formatString(strings.AndNMoreText, deniedPaths.length - 10)}</li>}
                    </ul>
                  )}
                </MessageBarBody>
              </MessageBar>
            )}

            {groupPermissionDenied && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  {strings.GroupExpansionSkippedPre} <strong>{strings.GroupMemberReadAllWord}</strong> {strings.GroupExpansionSkippedPost}{' '}
                  <strong>{strings.SharePointAdminCenterPath}</strong>.
                </MessageBarBody>
              </MessageBar>
            )}

            {entries.some((e) => e.scanIncomplete) && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  {formatString(strings.ScanIncompleteItemsMessage, entries.filter((e) => e.scanIncomplete).length)}
                </MessageBarBody>
              </MessageBar>
            )}

            <Divider />

            {/* Filter bar */}
            <Input
              placeholder={strings.FilterByNamePathUserPlaceholder}
              value={filterText}
              onChange={(_, d) => setFilterText(d.value)}
              style={{ width: '100%' }}
              aria-label={strings.FilterResultsLabel}
            />
            <div className={styles.row}>
              <Checkbox
                label={strings.UniquePermissionsOnlyToggle}
                checked={filterUniqueOnly}
                onChange={(_, d) => setFilterUniqueOnly(!!d.checked)}
              />
              <Checkbox
                label={strings.ExternalUsersOnlyExtLabel}
                checked={filterExternalOnly}
                onChange={(_, d) => setFilterExternalOnly(!!d.checked)}
              />
              <Checkbox
                label={strings.ExcludeLimitedAccessLabel}
                checked={excludeLimitedAccess}
                onChange={(_, d) => onExcludeLimitedAccessChange(!!d.checked)}
              />
              <Checkbox
                label={strings.ExcludeSharingLinksLabel}
                checked={excludeSharingLinks}
                onChange={(_, d) => onExcludeSharingLinksChange(!!d.checked)}
              />
              {(filterText || filterExternalOnly || filterUniqueOnly) && (
                <Body1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
                  {formatString(strings.ShowingOfCount, filteredEntries?.length ?? 0, entries.length)}
                </Body1>
              )}
            </div>

            {/* ── Results table ── */}
            {sortedResults.length > 0 && (
              <>
                <table className={styles.historyTable} aria-label={strings.ScanResultsLabel}>
                  <thead>
                    <tr>
                      <th className={styles.historyTh} style={{ width: '28px' }} />
                      {(
                        [
                          { col: 'type', label: strings.TypeColumnHeader },
                          { col: 'name', label: strings.NameColumnHeader },
                          { col: 'path', label: strings.PathColumnHeader },
                          { col: 'source', label: strings.PermissionsColumnHeader },
                        ] as { col: 'type' | 'name' | 'path' | 'source'; label: string }[]
                      ).map(({ col, label }) => (
                        <th
                          key={col}
                          className={styles.historyTh}
                          onClick={() => handleResultSort(col)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleResultSort(col); }
                          }}
                          tabIndex={0}
                          role="columnheader"
                          aria-sort={resultSortCol !== col ? 'none' : resultSortAsc ? 'ascending' : 'descending'}
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
                      // Filtered the same way as the export (shared applyPermFilters) so the
                      // "N assignments" count and the expanded table agree with what
                      // excludeLimitedAccess/filterExternalOnly/excludeSharingLinks are actually showing.
                      const rowPerms = applyPermFilters(entry.uniquePermissions, excludeLimitedAccess, filterExternalOnly, excludeSharingLinks);
                      const expandable = rowPerms.length > 0;
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
                                  aria-label={isExpanded ? strings.CollapsePermissionsLabel : strings.ExpandPermissionsLabel}
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
                                  {strings.HiddenFromSearchBadge}
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
                                <Badge appearance="filled" color="warning" size="small">{strings.UniqueBadge}</Badge>
                              ) : (
                                <Badge appearance="outline" size="small">{strings.InheritedBadge}</Badge>
                              )}
                              {expandable && (
                                <Text style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, marginLeft: '6px' }}>
                                  {formatString(rowPerms.length === 1 ? strings.AssignmentCountSingular : strings.AssignmentCountPlural, rowPerms.length)}
                                </Text>
                              )}
                            </td>
                          </tr>
                          {isExpanded && (
                            <tr>
                              <td className={styles.historyTd} />
                              <td className={styles.historyTd} colSpan={4} style={{ background: tokens.colorNeutralBackground2 }}>
                                <PermTable users={rowPerms} />
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
                      {formatString(strings.LoadMoreRemainingButton, (sortedResults.length - resultsVisible).toLocaleString())}
                    </Button>
                  </div>
                )}
              </>
            )}
          </div>
        )}
      </div>}
    </div>
  );
};
