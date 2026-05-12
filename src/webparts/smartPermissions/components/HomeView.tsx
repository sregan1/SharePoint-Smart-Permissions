import * as React from 'react';
import {
  Button,
  Card,
  CardHeader,
  Text,
  LargeTitle,
  Body1,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ShieldLock24Regular,
  DataBarVertical24Regular,
  FolderSearch24Regular,
  PersonSearch24Regular,
  Checkmark16Filled,
} from '@fluentui/react-icons';
import { AppView } from './App';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalXL,
    maxWidth: '1100px',
    margin: '0 auto',
  },
  header: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalS,
  },
  subtitle: {
    marginBottom: tokens.spacingVerticalXL,
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: tokens.spacingHorizontalL,
    '@media (max-width: 800px)': {
      gridTemplateColumns: '1fr 1fr',
    },
    '@media (max-width: 500px)': {
      gridTemplateColumns: '1fr',
    },
  },
  card: {
    cursor: 'pointer',
    padding: tokens.spacingVerticalL,
    minHeight: '180px',
    display: 'flex',
    flexDirection: 'column',
    justifyContent: 'space-between',
    ':hover': {
      boxShadow: tokens.shadow16,
    },
  },
  cardTitleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  cardTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: tokens.spacingVerticalS,
  },
  cardDesc: {
    flexGrow: 1,
    marginBottom: tokens.spacingVerticalM,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
  },
  featureItem: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalXS,
  },
  navButton: {
    width: '100%',
    minHeight: '52px',
    whiteSpace: 'normal',
    height: 'auto',
  },
});

export interface HomeViewProps {
  onNavigate: (view: AppView) => void;
}

export const HomeView: React.FC<HomeViewProps> = ({ onNavigate }) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <ShieldLock24Regular style={{ fontSize: '36px', flexShrink: 0, color: tokens.colorBrandForeground1 }} />
        <LargeTitle style={{ color: tokens.colorBrandForeground1 }}>SharePoint Smart Permissions</LargeTitle>
      </div>

      <div className={styles.subtitle}>
        <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
          Audit and understand SharePoint permissions in real time, directly from your browser — no PowerShell or admin tools required.
        </Body1>
      </div>

      <div className={styles.grid}>
        <Card className={styles.card}>
          <CardHeader
            header={
              <div className={styles.cardTitleRow}>
                <DataBarVertical24Regular style={{ flexShrink: 0 }} />
                <Text className={styles.cardTitle}>Permissions Report</Text>
              </div>
            }
          />
          <div className={styles.cardDesc}>
            <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, marginBottom: tokens.spacingVerticalS }}>
              Generate a comprehensive Excel report showing every unique permission assignment across your site. Only highlights differences from inherited permissions.
            </Body1>
            {[
              'Site, Library, Folder, or Item level',
              'Configurable folder depth',
              'Color-coded Excel export',
              'Scan all sites or a single site',
            ].map((f) => (
              <div key={f} className={styles.featureItem}>
                <Checkmark16Filled style={{ color: tokens.colorBrandForeground1, flexShrink: 0, marginTop: '2px' }} />
                <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>{f}</Body1>
              </div>
            ))}
          </div>
          <Button
            appearance="primary"
            className={styles.navButton}
            onClick={() => onNavigate('report')}
          >
            Run Permissions Report
          </Button>
        </Card>

        <Card className={styles.card}>
          <CardHeader
            header={
              <div className={styles.cardTitleRow}>
                <FolderSearch24Regular style={{ flexShrink: 0 }} />
                <Text className={styles.cardTitle}>Permissions Explorer</Text>
              </div>
            }
          />
          <div className={styles.cardDesc}>
            <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, marginBottom: tokens.spacingVerticalS }}>
              Browse permissions interactively in real time. Select any folder or file to instantly see who has access and what permission levels are assigned.
            </Body1>
            {[
              'Interactive folder/file tree',
              'Instant permission lookup',
              'Unique vs. inherited permissions',
              'Expand SharePoint group members',
            ].map((f) => (
              <div key={f} className={styles.featureItem}>
                <Checkmark16Filled style={{ color: tokens.colorBrandForeground1, flexShrink: 0, marginTop: '2px' }} />
                <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>{f}</Body1>
              </div>
            ))}
          </div>
          <Button
            appearance="primary"
            className={styles.navButton}
            onClick={() => onNavigate('explorer')}
          >
            Open Permissions Explorer
          </Button>
        </Card>

        <Card className={styles.card}>
          <CardHeader
            header={
              <div className={styles.cardTitleRow}>
                <PersonSearch24Regular style={{ flexShrink: 0 }} />
                <Text className={styles.cardTitle}>User Access</Text>
              </div>
            }
          />
          <div className={styles.cardDesc}>
            <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2, marginBottom: tokens.spacingVerticalS }}>
              Look up a specific user to see every location they can access on a site. Quickly identify over-privileged accounts or verify that access is correctly scoped.
            </Body1>
            {[
              'Per-user access analysis',
              'Full Site Access detection',
              'Shows path and permission level',
              'Search users by name',
            ].map((f) => (
              <div key={f} className={styles.featureItem}>
                <Checkmark16Filled style={{ color: tokens.colorBrandForeground1, flexShrink: 0, marginTop: '2px' }} />
                <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground2 }}>{f}</Body1>
              </div>
            ))}
          </div>
          <Button
            appearance="primary"
            className={styles.navButton}
            onClick={() => onNavigate('userAccess')}
          >
            Check User Access
          </Button>
        </Card>
      </div>

      <div
        style={{
          marginTop: tokens.spacingVerticalXXL,
          padding: tokens.spacingVerticalM,
          background: tokens.colorNeutralBackground3,
          borderRadius: tokens.borderRadiusMedium,
        }}
      >
        <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
          <strong>Note:</strong> This web part runs as the currently signed-in
          user. It can only see sites and items that user has permission to
          view. For a full tenant scan, use an account with appropriate
          read access across all sites.
        </Body1>
      </div>
    </div>
  );
};
