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
import { ReportOptions, ReportScope, PermissionEntry } from '../models/models';

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
  const [statusText, setStatusText] = React.useState('');
  const [error, setError] = React.useState('');
  const [entries, setEntries] = React.useState<PermissionEntry[] | null>(null);
  const [isExporting, setIsExporting] = React.useState(false);

  const abortRef = React.useRef<AbortController | null>(null);

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
    setStatusText('Starting scan…');

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
        (msg) => setStatusText(msg),
        abortRef.current.signal,
      );

      if (abortRef.current.signal.aborted) {
        setStatusText('Cancelled.');
        return;
      }

      setEntries(result);
      setStatusText(
        `Scan complete — ${result.length} object(s) found, ` +
          `${result.filter((e) => e.hasUniquePermissions).length} with unique permissions.`,
      );
    } catch (err: any) {
      if (err?.name === 'AbortError') {
        setStatusText('Cancelled.');
      } else {
        setError(`Error: ${err?.message ?? String(err)}`);
        setStatusText('');
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
          >
            <Radio value="Site" label={
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Globe24Regular style={{ fontSize: '16px' }} />Site only
              </span>
            } />
            <Radio value="Library" label={
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <BookDatabase24Regular style={{ fontSize: '16px' }} />Libraries
              </span>
            } />
            <Radio value="Folder" label={
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Folder24Regular style={{ fontSize: '16px' }} />Folders
              </span>
            } />
            <Radio value="Item" label={
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <FolderOpen24Regular style={{ fontSize: '16px' }} />Files &amp; Folders
              </span>
            } />
          </RadioGroup>
        </Field>

        {/* Folder depth (only shown when scope = Folder) */}
        {scope === 'Folder' && (
          <div className={styles.row}>
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
        )}

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
        {(isBusy || statusText) && !error && (
          <div className={styles.progressArea}>
            {isBusy && <ProgressBar />}
            <Body1>{statusText}</Body1>
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
              {statusText}
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
