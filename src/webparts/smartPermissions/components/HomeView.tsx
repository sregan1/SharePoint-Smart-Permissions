import * as React from 'react';
import {
  Button,
  Card,
  CardHeader,
  Text,
  Body1,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ShieldLock24Regular,
  DataBarVertical24Regular,
  FolderSearch24Regular,
  PersonSearch24Regular,
  Link24Regular,
  PeopleTeam24Regular,
  Person24Regular,
  BranchFork24Regular,
  Globe24Regular,
  Checkmark16Filled,
  ChevronDown20Regular,
  ChevronUp20Regular,
} from '@fluentui/react-icons';
import { AppView } from './App';

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
  moreToolsSection: {
    marginTop: tokens.spacingVerticalXL,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  moreToolsHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    padding: `${tokens.spacingVerticalM} ${tokens.spacingHorizontalL}`,
    cursor: 'pointer',
    background: tokens.colorNeutralBackground2,
    ':hover': {
      background: tokens.colorNeutralBackground3,
    },
  },
  toolRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalL}`,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    ':hover': {
      background: tokens.colorNeutralBackground2,
    },
  },
  toolIcon: {
    flexShrink: 0,
    color: tokens.colorBrandForeground1,
  },
  toolText: {
    flexGrow: 1,
    minWidth: 0,
  },
});

const MORE_TOOLS: Array<{
  view: AppView;
  icon: React.ReactElement;
  title: string;
  description: string;
}> = [
  {
    view: 'groups',
    icon: <PeopleTeam24Regular />,
    title: 'Permission Groups',
    description: 'List all SharePoint permission groups, their site roles, and members.',
  },
  {
    view: 'externalUsers',
    icon: <Person24Regular />,
    title: 'External Users',
    description: 'List external (#ext#) accounts with their group memberships and check their full access with one click.',
  },
  {
    view: 'brokenInheritance',
    icon: <BranchFork24Regular />,
    title: 'Broken Inheritance Finder',
    description: 'Find all libraries, folders, and files that have unique permissions set.',
  },
  {
    view: 'sharingLinks',
    icon: <Link24Regular />,
    title: 'Sharing Links',
    description: 'Enumerate all active sharing links — anonymous, org-wide, and user-specific.',
  },
  {
    view: 'anonymousLinks',
    icon: <Globe24Regular />,
    title: 'Anonymous Access Summary',
    description: 'Summarize anonymous and org-wide sharing links by library.',
  },
];

export interface HomeViewProps {
  onNavigate: (view: AppView) => void;
  primaryColor: string;
}

export const HomeView: React.FC<HomeViewProps> = ({ onNavigate, primaryColor }) => {
  const styles = useStyles();
  const [moreExpanded, setMoreExpanded] = React.useState(false);

  return (
    <div>
      {/* Full-width banner — matches the header on all other views */}
      <div className={styles.banner} style={{ background: primaryColor }}>
        <ShieldLock24Regular style={{ color: 'white', fontSize: '20px', flexShrink: 0 }} />
        <Text style={{ color: 'white', fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
          SharePoint Smart Permissions
        </Text>
      </div>

      <div className={styles.root}>
        <div className={styles.subtitle}>
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
            Audit and understand SharePoint permissions in real time, directly from your browser — no PowerShell or admin tools required.
          </Body1>
        </div>

        <div className={styles.grid}>
          <Card
            className={styles.card}
            onClick={() => onNavigate('report')}
            role="button"
            tabIndex={0}
            aria-label="Run Permissions Report"
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onNavigate('report'); } }}
          >
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
                Generate a comprehensive Excel report showing every unique permission assignment across your site.
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
            <Button appearance="primary" className={styles.navButton} tabIndex={-1}>
              Run Permissions Report
            </Button>
          </Card>

          <Card
            className={styles.card}
            onClick={() => onNavigate('explorer')}
            role="button"
            tabIndex={0}
            aria-label="Open Permissions Explorer"
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onNavigate('explorer'); } }}
          >
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
                Browse permissions interactively in real time. Select any folder or file to instantly see who has access.
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
            <Button appearance="primary" className={styles.navButton} tabIndex={-1}>
              Open Permissions Explorer
            </Button>
          </Card>

          <Card
            className={styles.card}
            onClick={() => onNavigate('userAccess')}
            role="button"
            tabIndex={0}
            aria-label="Check User Access"
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onNavigate('userAccess'); } }}
          >
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
                Look up a specific user to see every location they can access on a site. Quickly identify over-privileged accounts.
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
            <Button appearance="primary" className={styles.navButton} tabIndex={-1}>
              Check User Access
            </Button>
          </Card>
        </div>

        {/* More Tools — collapsible */}
        <div className={styles.moreToolsSection}>
          <div
            className={styles.moreToolsHeader}
            onClick={() => setMoreExpanded(!moreExpanded)}
            role="button"
            tabIndex={0}
            aria-expanded={moreExpanded}
            aria-controls="more-tools-list"
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setMoreExpanded(!moreExpanded); } }}
          >
            <Text style={{ fontWeight: tokens.fontWeightSemibold }}>
              More tools ({MORE_TOOLS.length})
            </Text>
            {moreExpanded ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
          </div>

          {moreExpanded && (
            <div id="more-tools-list">
              {MORE_TOOLS.map((tool) => (
                <div
                  key={tool.view}
                  className={styles.toolRow}
                  onClick={() => onNavigate(tool.view)}
                  role="button"
                  tabIndex={0}
                  aria-label={tool.title}
                  onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onNavigate(tool.view); } }}
                >
                  <span className={styles.toolIcon}>{tool.icon}</span>
                  <div className={styles.toolText}>
                    <Text style={{ fontWeight: tokens.fontWeightSemibold, display: 'block' }}>
                      {tool.title}
                    </Text>
                    <Body1 style={{ fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 }}>
                      {tool.description}
                    </Body1>
                  </div>
                  <Button appearance="subtle" size="small" tabIndex={-1}>
                    Open
                  </Button>
                </div>
              ))}
            </div>
          )}
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
