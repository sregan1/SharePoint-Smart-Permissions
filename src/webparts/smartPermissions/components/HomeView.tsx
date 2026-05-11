import * as React from 'react';
import {
  Button,
  Card,
  CardHeader,
  Text,
  Title2,
  Body1,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { AppView } from './App';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalXL,
    maxWidth: '1100px',
    margin: '0 auto',
  },
  header: {
    marginBottom: tokens.spacingVerticalXL,
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
  },
  icon: {
    fontSize: '32px',
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
  cardTitle: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: tokens.spacingVerticalS,
  },
  cardDesc: {
    color: tokens.colorNeutralForeground3,
    flexGrow: 1,
    marginBottom: tokens.spacingVerticalM,
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
        <span className={styles.icon}>🔐</span>
        <div>
          <Title2>SharePoint Smart Permissions</Title2>
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
            Browser-based permissions reporting and auditing for SharePoint Online
          </Body1>
        </div>
      </div>

      <div className={styles.grid}>
        <Card className={styles.card}>
          <CardHeader
            header={
              <Text className={styles.cardTitle}>📊 Permissions Report</Text>
            }
          />
          <Body1 className={styles.cardDesc}>
            Scan a site, library, or entire tenant. Export a colour-coded Excel
            workbook showing every unique permission break — who has access to
            what, and where they got it.
          </Body1>
          <Button
            appearance="primary"
            onClick={() => onNavigate('report')}
          >
            Run Permissions Report
          </Button>
        </Card>

        <Card className={styles.card}>
          <CardHeader
            header={
              <Text className={styles.cardTitle}>🔍 Permissions Explorer</Text>
            }
          />
          <Body1 className={styles.cardDesc}>
            Browse a document library folder-by-folder and see live permissions
            on any item. Instantly see whether permissions are unique or
            inherited, and trace them back to their source.
          </Body1>
          <Button
            appearance="primary"
            onClick={() => onNavigate('explorer')}
          >
            Open Permissions Explorer
          </Button>
        </Card>

        <Card className={styles.card}>
          <CardHeader
            header={
              <Text className={styles.cardTitle}>👤 User Access</Text>
            }
          />
          <Body1 className={styles.cardDesc}>
            Look up any user on a site and see every library, folder, and file
            they can access — along with their permission level at each
            location.
          </Body1>
          <Button
            appearance="primary"
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
