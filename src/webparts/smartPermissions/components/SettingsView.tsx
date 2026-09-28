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
import * as strings from 'SmartPermissionsWebPartStrings';

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
  excludeSharingLinks: boolean;
  onExcludeSharingLinksChange: (val: boolean) => void;
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
  excludeSharingLinks,
  onExcludeSharingLinksChange,
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
          {strings.BackButton}
        </Button>
        <Title3>{strings.SettingsTitle}</Title3>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXL }}>

        {/* ── Libraries ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.LibrariesSectionTitle}</Text>
          <div className={styles.row}>
            <Checkbox
              label={strings.IncludeHiddenLabel}
              checked={includeHidden}
              onChange={(_, d) => onIncludeHiddenChange(!!d.checked)}
            />
            <Tooltip
              content={strings.IncludeHiddenTooltip}
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label={strings.MoreInfoHiddenLibrariesLabel}
              />
            </Tooltip>
          </div>
          <div className={styles.row}>
            <Checkbox
              label={strings.ExcludeLimitedAccessLabel}
              checked={excludeLimitedAccess}
              onChange={(_, d) => onExcludeLimitedAccessChange(!!d.checked)}
            />
            <Tooltip
              content={strings.ExcludeLimitedAccessTooltip}
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label={strings.MoreInfoLimitedAccessLabel}
              />
            </Tooltip>
          </div>
          <div className={styles.row}>
            <Checkbox
              label={strings.ExcludeSharingLinksLabel}
              checked={excludeSharingLinks}
              onChange={(_, d) => onExcludeSharingLinksChange(!!d.checked)}
            />
            <Tooltip
              content={strings.ExcludeSharingLinksTooltip}
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label={strings.MoreInfoSharingLinksLabel}
              />
            </Tooltip>
          </div>
        </div>

        <Divider />

        {/* ── Performance ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.PerformanceSectionTitle}</Text>

          <div className={styles.row}>
            <Label>{strings.ConcurrentRequestsLabel}</Label>
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
              content={strings.ConcurrencyTooltip}
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label={strings.MoreInfoConcurrencyLabel}
              />
            </Tooltip>
          </div>
          <Text size={200} className={styles.hint}>
            {strings.ConcurrencyHint}
          </Text>

          <div className={styles.row} style={{ marginTop: tokens.spacingVerticalS }}>
            <Label>{strings.GroupMemberCapLabel}</Label>
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
              content={strings.GroupMemberCapTooltip}
              relationship="description"
              withArrow
            >
              <Button
                appearance="transparent"
                icon={<Info16Regular />}
                size="small"
                style={{ minWidth: 'unset', padding: '2px' }}
                aria-label={strings.MoreInfoGroupMemberCapLabel}
              />
            </Tooltip>
          </div>
          <Text size={200} className={styles.hint}>
            {strings.GroupMemberCapHint}
          </Text>
        </div>

        <Divider />

        {/* ── Default view instructions ── */}
        <div className={styles.section}>
          <Text weight="semibold" style={{ display: 'block' }}>{strings.DefaultViewSectionTitle}</Text>
          <Text size={300} style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
            {strings.DefaultViewInstructionsIntro}
          </Text>
          <ol className={styles.instructionList}>
            {[
              <>{strings.DefaultViewStep1Pre} <strong>{strings.EditModeWord}</strong> {strings.DefaultViewStep1Post}</>,
              <>{strings.DefaultViewStep2Pre} <strong>{strings.PencilEditWord}</strong> {strings.DefaultViewStep2Post}</>,
              <>{strings.DefaultViewStep3Pre} <strong>{strings.DefaultViewOnOpenWord}</strong> {strings.DefaultViewStep3Post}</>,
              <><strong>{strings.RepublishWord}</strong> {strings.DefaultViewStep4Post}</>,
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
