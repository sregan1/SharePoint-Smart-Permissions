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
  DataBarVertical24Regular,
  FolderSearch24Regular,
  PersonSearch24Regular,
  LockClosed16Regular,
} from '@fluentui/react-icons';
import * as strings from 'SmartPermissionsWebPartStrings';
import { AppView } from './App';
import screenshotReport from '../assets/screenshot_report.png';
import screenshotExplorer from '../assets/screenshot_explorer.png';
import screenshotUserAccess from '../assets/screenshot_user_access.png';

const useStyles = makeStyles({
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
    flexShrink: 0,
    objectFit: 'cover',
    objectPosition: 'left top',
    display: 'block',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
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
  cardFooter: {
    padding: tokens.spacingVerticalM,
  },
  navButton: {
    width: '100%',
    minHeight: '36px',
  },
  cardDisabled: {
    // Dimmed but still legible: the description and the lock explain why it's unavailable.
    opacity: '0.6',
    cursor: 'not-allowed',
  },
});

export interface HomeViewProps {
  onNavigate: (view: AppView) => void;
  canManagePermissions: boolean | null;
  hiddenViews?: AppView[];
}

export const HomeView: React.FC<HomeViewProps> = ({ onNavigate, canManagePermissions, hiddenViews }) => {
  const styles = useStyles();

  const cards = [
    {
      view: 'explorer' as AppView,
      icon: <FolderSearch24Regular style={{ flexShrink: 0 }} />,
      title: strings.ExplorerCardTitle,
      screenshot: screenshotExplorer,
      alt: strings.ExplorerCardAlt,
      desc: strings.ExplorerCardDesc,
      buttonLabel: strings.ExplorerCardButton,
    },
    {
      view: 'report' as AppView,
      icon: <DataBarVertical24Regular style={{ flexShrink: 0 }} />,
      title: strings.ReportCardTitle,
      screenshot: screenshotReport,
      alt: strings.ReportCardAlt,
      desc: strings.ReportCardDesc,
      buttonLabel: strings.ReportCardButton,
    },
    {
      view: 'userAccess' as AppView,
      icon: <PersonSearch24Regular style={{ flexShrink: 0 }} />,
      title: strings.UserAccessCardTitle,
      screenshot: screenshotUserAccess,
      alt: strings.UserAccessCardAlt,
      desc: strings.UserAccessCardDesc,
      buttonLabel: strings.UserAccessCardButton,
    },
  ].filter((c) => !hiddenViews || hiddenViews.indexOf(c.view) === -1);

  return (
    <div>
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
            <span aria-hidden="true" style={{ flexShrink: 0, fontSize: '16px' }}>⚠️</span>
            <span>
              <strong>{strings.SiteOwnerRequiredTitle} </strong>
              {strings.SiteOwnerRequiredDesc}
            </span>
          </div>
        )}

        <div className={styles.subtitle}>
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
            {strings.HomeSubtitle}
          </Body1>
        </div>

        <div className={styles.grid}>
          {cards.map(({ view, icon, title, screenshot, alt, desc, buttonLabel }) => {
            const disabled = canManagePermissions === false;
            // The inner Button is the one real control (focusable, announced with its own
            // label); the card itself is only a larger mouse target, so it carries no
            // role/tabIndex. A disabled card keeps a focusable button so keyboard and
            // screen-reader users can reach the "Site Owner required" explanation.
            const button = (
              <Button
                appearance="primary"
                className={styles.navButton}
                disabledFocusable={disabled}
                icon={disabled ? <LockClosed16Regular /> : undefined}
                iconPosition="after"
                onClick={disabled ? undefined : (e) => { e.stopPropagation(); onNavigate(view); }}
              >
                {buttonLabel}
              </Button>
            );
            return (
              <Card
                key={view}
                className={`${styles.card}${disabled ? ` ${styles.cardDisabled}` : ''}`}
                onClick={disabled ? undefined : () => onNavigate(view)}
              >
                <div className={styles.cardBody}>
                  <div className={styles.cardTitleRow}>
                    {icon}
                    <Text as="h2" className={styles.cardTitle}>{title}</Text>
                  </div>
                  <Body1 className={styles.cardDesc}>{desc}</Body1>
                </div>
                <img src={screenshot} alt={alt} className={styles.cardImage} />
                <div className={styles.cardFooter}>
                  {disabled ? (
                    <Tooltip content={strings.RequiresSiteOwnerTooltip} relationship="description">
                      {button}
                    </Tooltip>
                  ) : button}
                </div>
              </Card>
            );
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
            <strong>{strings.HomeNoteLabel}</strong> {strings.HomeNoteText}
          </Body1>
        </div>
      </div>
    </div>
  );
};
