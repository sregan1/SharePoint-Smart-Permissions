import * as React from 'react';
import {
  Button,
  Checkbox,
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
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { ReportOptions, ReportScope, PermissionEntry, ScanProgress } from '../models/models';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '760px',
    margin: '0 auto',
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
});

export interface PermissionsReportViewProps {
  sp: SharePointService;
  excel: ExcelExportService;
  siteUrl: string;
  includeHidden: boolean;
  onBack: () => void;
}

export const PermissionsReportView: React.FC<PermissionsReportViewProps> = ({
  sp,
  excel,
  siteUrl,
  includeHidden,
  onBack,
}) => {
  const styles = useStyles();

  // ── Form state ──
  const [allSites, setAllSites] = React.useState(false);
  const [scope, setScope] = React.useState<string>('Library');
  const [folderDepth, setFolderDepth] = React.useState(2);

  // ── Run state ──
  const [isBusy, setIsBusy] = React.useState(false);
  const [scanProgress, setScanProgress] = React.useState<ScanProgress>({ message: '', scanned: 0, libsDone: 0, libsTotal: 0 });
  const [elapsed, setElapsed] = React.useState(0);
  const [error, setError] = React.useState('');
  const [entries, setEntries] = React.useState<PermissionEntry[] | null>(null);
  const [isExporting, setIsExporting] = React.useState(false);

  const abortRef = React.useRef<AbortController | null>(null);

  React.useEffect(() => {
    if (!isBusy) { setElapsed(0); return; }
    const start = Date.now();
    const id = setInterval(() => setElapsed(Math.floor((Date.now() - start) / 1000)), 500);
    return () => clearInterval(id);
  }, [isBusy]);

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
    abortRef.current = new AbortController();
    setIsBusy(true);
    setError('');
    setEntries(null);
    setScanProgress({ message: 'Starting scan…', scanned: 0, libsDone: 0, libsTotal: 0 });

    try {
      const options: ReportOptions = {
        siteUrl: siteUrl.trim(),
        allSites,
        scope: scope as ReportScope,
        folderDepth,
        includeHidden,
      };

      const result = await sp.scanPermissions(
        options,
        (progress) => setScanProgress(progress),
        abortRef.current.signal,
      );

      if (abortRef.current.signal.aborted) {
        setScanProgress((prev) => ({ ...prev, message: 'Cancelled.' }));
        return;
      }

      setEntries(result);
      setScanProgress((prev) => ({
        ...prev,
        message:
          `Scan complete — ${result.length} object(s) found, ` +
          `${result.filter((e) => e.hasUniquePermissions).length} with unique permissions.`,
      }));
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        setScanProgress((prev) => ({ ...prev, message: 'Cancelled.' }));
      } else {
        setError(`Error: ${err?.message ?? String(err)}`);
        setScanProgress((prev) => ({ ...prev, message: '' }));
      }
    } finally {
      setIsBusy(false);
    }
  };

  const handleCancel = (): void => {
    abortRef.current?.abort();
  };

  const handleExport = async (): Promise<void> => {
    if (!entries) return;
    setIsExporting(true);
    try {
      await excel.export(entries, siteUrl.trim());
    } catch (err: any) {
      setError(`Export error: ${err?.message ?? String(err)}`);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={onBack}
          disabled={isBusy}
        >
          Back
        </Button>
        <Title3>Permissions Report</Title3>
      </div>

      <div className={styles.form}>
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

        <Divider />

        {/* Action buttons */}
        <div className={styles.row}>
          <Button
            appearance="primary"
            onClick={handleRun}
            disabled={isBusy}
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
            {isBusy && scanProgress.scanned > 0 && (
              <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                {scanProgress.scanned} items scanned
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
              <Badge appearance="filled" color="success">
                {entries.length} objects
              </Badge>
              <Badge appearance="filled" color="warning">
                {entries.filter((e) => e.hasUniquePermissions).length} unique
              </Badge>
              <Badge appearance="outline">
                {entries.filter((e) => !e.hasUniquePermissions).length} inherited
              </Badge>
            </div>

            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
              {scanProgress.message}
            </Body1>

            <div className={styles.row}>
              <Button
                appearance="primary"
                icon={<DocumentArrowDown24Regular />}
                onClick={handleExport}
                disabled={isExporting}
              >
                {isExporting ? 'Generating Excel…' : 'Export to Excel'}
              </Button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
