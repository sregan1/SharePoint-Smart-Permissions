import * as React from 'react';
import {
  Button,
  Checkbox,
  Text,
  Title3,
  Label,
  Divider,
  Tooltip,
  SpinButton,
  tokens,
  makeStyles,
} from '@fluentui/react-components';
import { ArrowLeft24Regular, Info16Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '540px',
    margin: '0 auto',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  hint: {
    display: 'block',
    color: tokens.colorNeutralForeground3,
    marginLeft: '24px',
    lineHeight: '1.5',
  },
  instructionList: {
    margin: '8px 0 0 0',
    paddingLeft: '20px',
    lineHeight: '1.8',
  },
});

export interface SettingsViewProps {
  includeHidden: boolean;
  onIncludeHiddenChange: (val: boolean) => void;
  excludeLimitedAccess: boolean;
  onExcludeLimitedAccessChange: (val: boolean) => void;
  scanConcurrency: number;
  onScanConcurrencyChange: (val: number) => void;
  groupMemberCap: number;
  onGroupMemberCapChange: (val: number) => void;
  onBack: () => void;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  includeHidden,
  onIncludeHiddenChange,
  excludeLimitedAccess,
  onExcludeLimitedAccessChange,
  scanConcurrency,
  onScanConcurrencyChange,
  groupMemberCap,
  onGroupMemberCapChange,
  onBack,
}) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack}>
          Back
        </Button>
        <Title3>Settings</Title3>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXL }}>

        {/* ── Libraries ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Libraries</Text>
          <div className={styles.row}>
            <Checkbox
              label="Include system and hidden libraries"
              checked={includeHidden}
              onChange={(_, d) => onIncludeHiddenChange(!!d.checked)}
            />
            <Tooltip
              content="When checked, includes Style Library, Form Templates, Site Assets, and other libraries hidden from default views. Applies to all tools."
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label="More info about hidden libraries"
              />
            </Tooltip>
          </div>
          <div className={styles.row}>
            <Checkbox
              label="Exclude Limited Access entries"
              checked={excludeLimitedAccess}
              onChange={(_, d) => onExcludeLimitedAccessChange(!!d.checked)}
            />
            <Tooltip
              content="Hides users and groups whose only SharePoint permission is Limited Access — automatically assigned when files are shared via links. Applies to Permissions Report, Explorer, and User Access."
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label="More info about Limited Access"
              />
            </Tooltip>
          </div>
        </div>

        <Divider />

        {/* ── Performance ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Performance</Text>

          <div className={styles.row}>
            <Label>Concurrent API requests:</Label>
            <SpinButton
              value={scanConcurrency}
              min={1}
              max={10}
              onChange={(_, d) =>
                onScanConcurrencyChange(
                  d.value !== undefined ? d.value : parseInt(d.displayValue ?? '4', 10),
                )
              }
              style={{ width: '80px' }}
            />
            <Tooltip
              content="How many SharePoint API requests run in parallel during scans. Higher values are faster but more likely to trigger throttling (HTTP 429). 3–5 is recommended."
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label="More info about concurrency"
              />
            </Tooltip>
          </div>
          <Text size={200} className={styles.hint}>
            Higher values scan faster but may trigger SharePoint throttling. Recommended: 3–5.
          </Text>

          <div className={styles.row} style={{ marginTop: tokens.spacingVerticalS }}>
            <Label>Group member display cap:</Label>
            <SpinButton
              value={groupMemberCap}
              min={50}
              max={5000}
              step={50}
              onChange={(_, d) =>
                onGroupMemberCapChange(
                  d.value !== undefined ? d.value : parseInt(d.displayValue ?? '500', 10),
                )
              }
              style={{ width: '100px' }}
            />
            <Tooltip
              content="Maximum members shown per group when 'Expand group members' is enabled. Larger groups are capped and a notice is shown. Increasing this uses more memory."
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label="More info about group member cap"
              />
            </Tooltip>
          </div>
          <Text size={200} className={styles.hint}>
            Groups larger than this limit show a truncation notice. Default: 500.
          </Text>
        </div>

        <Divider />

        {/* ── Default view instructions ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>Default view on load</Text>
          <Text size={300} style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
            To change which screen opens when the web part first loads, edit the web part properties:
          </Text>
          <ol className={styles.instructionList}>
            {[
              <>Put the SharePoint page into <strong>Edit</strong> mode.</>,
              <>Click the <strong>pencil (edit)</strong> icon on the Smart Permissions web part.</>,
              <>In the property panel, choose a view from the <strong>Default view on open</strong> dropdown.</>,
              <><strong>Republish</strong> the page to save the change.</>,
            ].map((step, i) => (
              <li key={i}>
                <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>{step}</Text>
              </li>
            ))}
          </ol>
        </div>

      </div>
    </div>
  );
};
