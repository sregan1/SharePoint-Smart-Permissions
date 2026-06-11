import * as React from 'react';
import {
  Button,
  Card,
  Text,
  Body1,
  Tooltip,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ShieldLock24Regular,
  DataBarVertical24Regular,
  FolderSearch24Regular,
  PersonSearch24Regular,
  LockClosed16Regular,
} from '@fluentui/react-icons';
import { AppView } from './App';
import screenshotReport from '../assets/screenshot_report.png';
import screenshotExplorer from '../assets/screenshot_explorer.png';
import screenshotUserAccess from '../assets/screenshot_user_access.png';

const useStyles = makeStyles({
  banner: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
  },
  root: {
    padding: tokens.spacingVerticalXL,
    maxWidth: '1100px',
    margin: '0 auto',
  },
  subtitle: {
    marginBottom: tokens.spacingVerticalXL,
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(3, 1fr)',
    gap: tokens.spacingHorizontalL,
    '@media (max-width: 800px)': {
      gridTemplateColumns: 'repeat(2, 1fr)',
    },
    '@media (max-width: 500px)': {
      gridTemplateColumns: '1fr',
    },
  },
  card: {
    cursor: 'pointer',
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    padding: '0',
    ':hover': {
      boxShadow: tokens.shadow16,
    },
  },
  cardImage: {
    width: '100%',
    height: '180px',
    objectFit: 'cover',
    objectPosition: 'top',
    display: 'block',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  cardBody: {
    display: 'flex',
    flexDirection: 'column',
    flexGrow: 1,
    padding: tokens.spacingVerticalM,
    gap: tokens.spacingVerticalS,
  },
  cardTitleRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
  },
  cardTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  cardDesc: {
    flexGrow: 1,
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  navButton: {
    width: '100%',
    minHeight: '36px',
  },
  cardDisabled: {
    opacity: '0.15',
    pointerEvents: 'none',
  },
});

export interface HomeViewProps {
  onNavigate: (view: AppView) => void;
  primaryColor: string;
  canManagePermissions: boolean | null;
  siteOwners: { title: string; email: string }[];
}

export const HomeView: React.FC<HomeViewProps> = ({ onNavigate, primaryColor, canManagePermissions, siteOwners }) => {
  const styles = useStyles();

  const cards = [
    {
      view: 'explorer' as AppView,
      icon: <FolderSearch24Regular style={{ flexShrink: 0 }} />,
      title: 'Permissions Explorer',
      screenshot: screenshotExplorer,
      alt: 'Permissions Explorer with folder tree and permissions panel',
      desc: 'Browse any folder or file and instantly see who has access — with live, real-time permission lookups.',
      buttonLabel: 'Open Permissions Explorer',
    },
    {
      view: 'report' as AppView,
      icon: <DataBarVertical24Regular style={{ flexShrink: 0 }} />,
      title: 'Permissions Report',
      screenshot: screenshotReport,
      alt: 'Permissions Report configuration screen',
      desc: 'Generate a color-coded Excel report of every unique permission assignment across your site.',
      buttonLabel: 'Run Permissions Report',
    },
    {
      view: 'userAccess' as AppView,
      icon: <PersonSearch24Regular style={{ flexShrink: 0 }} />,
      title: 'User Access',
      screenshot: screenshotUserAccess,
      alt: 'User Access screen showing accessible locations for a selected user',
      desc: 'Look up any user to see every location they can access on a site, with their exact permission level.',
      buttonLabel: 'Check User Access',
    },
  ];

  return (
    <div>
      <div className={styles.banner} style={{ background: primaryColor }}>
        <ShieldLock24Regular style={{ color: 'white', fontSize: '20px', flexShrink: 0 }} />
        <Text style={{ color: 'white', fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
          SharePoint Smart Permissions
        </Text>
      </div>

      <div className={styles.root}>
        {canManagePermissions === false && (
          <div style={{
            display: 'flex',
            alignItems: 'flex-start',
            gap: tokens.spacingHorizontalS,
            padding: tokens.spacingVerticalM,
            marginBottom: tokens.spacingVerticalL,
            background: tokens.colorPaletteYellowBackground1,
            border: `1px solid ${tokens.colorPaletteYellowBorder1}`,
            borderRadius: tokens.borderRadiusMedium,
            color: tokens.colorNeutralForeground1,
            fontSize: tokens.fontSizeBase300,
            lineHeight: tokens.lineHeightBase300,
          }}>
            <span style={{ flexShrink: 0, fontSize: '16px' }}>⚠️</span>
            <span>
              <strong>Site Owner access required — </strong>
              These tools require Site Owner access. Contact a site owner if you have questions about permissions.
            </span>
          </div>
        )}

        <div className={styles.subtitle}>
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
            Audit and understand SharePoint permissions — no PowerShell required.
          </Body1>
        </div>

        <div className={styles.grid}>
          {cards.map(({ view, icon, title, screenshot, alt, desc, buttonLabel }) => {
            const disabled = canManagePermissions === false;
            const card = (
              <Card
                key={view}
                className={`${styles.card}${disabled ? ` ${styles.cardDisabled}` : ''}`}
                style={disabled ? { filter: 'grayscale(1)' } : undefined}
                onClick={disabled ? undefined : () => onNavigate(view)}
                role="button"
                tabIndex={disabled ? -1 : 0}
                aria-disabled={disabled}
                aria-label={buttonLabel}
                onKeyDown={disabled ? undefined : (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onNavigate(view); } }}
              >
                <img src={screenshot} alt={alt} className={styles.cardImage} />
                <div className={styles.cardBody}>
                  <div className={styles.cardTitleRow}>
                    {icon}
                    <Text className={styles.cardTitle}>{title}</Text>
                  </div>
                  <Body1 className={styles.cardDesc}>{desc}</Body1>
                  <Button
                    appearance="primary"
                    className={styles.navButton}
                    tabIndex={-1}
                    disabled={disabled}
                    icon={disabled ? <LockClosed16Regular /> : undefined}
                    iconPosition="after"
                  >
                    {buttonLabel}
                  </Button>
                </div>
              </Card>
            );
            return disabled ? (
              <Tooltip key={view} content="Requires Site Owner access" relationship="description">
                <div style={{ cursor: 'not-allowed' }}>{card}</div>
              </Tooltip>
            ) : card;
          })}
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
    </div>
  );
};
