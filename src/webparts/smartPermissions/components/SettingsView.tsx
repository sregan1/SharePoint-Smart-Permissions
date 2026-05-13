import * as React from 'react';
import {
  Button,
  Checkbox,
  Text,
  Title3,
  Divider,
  Tooltip,
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
  instructionList: {
    margin: '8px 0 0 0',
    paddingLeft: '20px',
    lineHeight: '1.8',
  },
});

export interface SettingsViewProps {
  includeHidden: boolean;
  onIncludeHiddenChange: (val: boolean) => void;
  onBack: () => void;
}

export const SettingsView: React.FC<SettingsViewProps> = ({
  includeHidden,
  onIncludeHiddenChange,
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
        <div>
          <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalS }}>
            Libraries
          </Text>
          <div className={styles.row}>
            <Checkbox
              label="Include system and hidden libraries"
              checked={includeHidden}
              onChange={(_, d) => onIncludeHiddenChange(!!d.checked)}
            />
            <Tooltip
              content="When checked, includes system and hidden libraries such as Style Library, Form Templates, Site Assets, and others not shown in default views. Applies to Permissions Explorer and User Access."
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
          <Text
            size={200}
            style={{
              display: 'block',
              color: tokens.colorNeutralForeground3,
              marginTop: tokens.spacingVerticalXS,
              marginLeft: '24px',
              lineHeight: '1.5',
            }}
          >
            When checked, includes Style Library, Form Templates, Site Assets, and other libraries
            hidden from default views. Applies to Permissions Explorer and User Access.
          </Text>
        </div>

        <Divider />

        {/* ── Default view instructions ── */}
        <div>
          <Text weight="semibold" style={{ display: 'block', marginBottom: tokens.spacingVerticalS }}>
            Default view on load
          </Text>
          <Text size={300} style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
            To change which screen opens when the web part first loads, edit the web part properties:
          </Text>
          <ol className={styles.instructionList}>
            <li>
              <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>
                Put the SharePoint page into <strong>Edit</strong> mode.
              </Text>
            </li>
            <li>
              <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>
                Click the <strong>pencil (edit)</strong> icon on the Smart Permissions web part.
              </Text>
            </li>
            <li>
              <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>
                In the property panel, choose a view from the <strong>Default view on open</strong> dropdown.
              </Text>
            </li>
            <li>
              <Text size={300} style={{ color: tokens.colorNeutralForeground2 }}>
                <strong>Republish</strong> the page to save the change.
              </Text>
            </li>
          </ol>
        </div>

      </div>
    </div>
  );
};
