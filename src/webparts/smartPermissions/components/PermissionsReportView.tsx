import * as React from 'react';
import {
  Button,
  Input,
  Label,
  Field,
  RadioGroup,
  Radio,
  Checkbox,
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
import { ArrowLeft24Regular, DocumentArrowDown24Regular } from '@fluentui/react-icons';

import { WebPartContext } from '@microsoft/sp-webpart-base';
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
  statsRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalL,
    flexWrap: 'wrap',
    marginTop: tokens.spacingVerticalS,
  },
  statItem: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    minWidth: '100px',
  },
});

export interface PermissionsReportViewProps {
  context: WebPartContext;
  sp: SharePointService;
  excel: ExcelExportService;
  onBack: () => void;
}

export const PermissionsReportView: React.FC<PermissionsReportViewProps> = ({
  context,
  sp,
  excel,
  onBack,
}) => {
  const styles = useStyles();

  // ── Form state ──
  const [siteUrl, setSiteUrl] = React.useState(context.pageContext.web.absoluteUrl);
  const [allSites, setAllSites] = React.useState(false);
  const [scope, setScope] = React.useState<string>('Library');
  const [folderDepth, setFolderDepth] = React.useState(2);
  const [includeHidden, setIncludeHidden] = React.useState(false);

  // ── Run state ──
  const [isBusy, setIsBusy] = React.useState(false);
  const [statusText, setStatusText] = React.useState('');
  const [error, setError] = React.useState('');
  const [entries, setEntries] = React.useState<PermissionEntry[] | null>(null);
  const [isExporting, setIsExporting] = React.useState(false);

  const abortRef = React.useRef<AbortController | null>(null);

  const canRun = siteUrl.trim().length > 0 && !isBusy;

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
        {/* Site URL */}
        <Field label="Site URL (or tenant root URL for all-sites scan)">
          <Input
            value={siteUrl}
            onChange={(_, d) => setSiteUrl(d.value)}
            placeholder="https://contoso.sharepoint.com/sites/mysite"
            style={{ width: '100%' }}
            disabled={isBusy}
          />
        </Field>

        {/* All-sites toggle */}
        <Checkbox
          label="Scan all site collections (enter tenant root URL above)"
          checked={allSites}
          onChange={(_, d) => setAllSites(!!d.checked)}
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
          >
            <Radio value="Site" label="Site only" />
            <Radio value="Library" label="Libraries" />
            <Radio value="Folder" label="Folders" />
            <Radio value="Item" label="Files & Folders" />
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

        {/* Include hidden */}
        <Checkbox
          label="Include hidden libraries"
          checked={includeHidden}
          onChange={(_, d) => setIncludeHidden(!!d.checked)}
          disabled={isBusy}
        />

        <Divider />

        {/* Action buttons */}
        <div className={styles.row}>
          <Button
            appearance="primary"
            onClick={handleRun}
            disabled={!canRun}
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
