import * as React from 'react';
import {
  Button,
  Badge,
  Text,
  Title3,
  Body1,
  Spinner,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  ChevronRight16Regular,
  ChevronDown16Regular,
  People24Regular,
  Person16Regular,
  PeopleTeam16Regular,
  Shield16Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { UserPermissionInfo } from '../models/models';

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '900px',
    margin: '0 auto',
    minHeight: '400px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalL,
  },
  groupList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  groupRow: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
  },
  groupHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    cursor: 'pointer',
    background: tokens.colorNeutralBackground1,
    ':hover': {
      background: tokens.colorNeutralBackground1Hover,
    },
  },
  groupTitle: {
    fontWeight: tokens.fontWeightSemibold,
    flex: 1,
  },
  groupDesc: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  memberList: {
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    background: tokens.colorNeutralBackground2,
    padding: tokens.spacingVerticalS,
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
  },
  memberRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `4px ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusSmall,
  },
  memberName: {
    flex: 1,
    fontSize: tokens.fontSizeBase300,
  },
  loginName: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
    fontFamily: 'monospace',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    maxWidth: '300px',
  },
});

// ── Types ─────────────────────────────────────────────────────────────────────

interface GroupInfo {
  id: number;
  title: string;
  loginName: string;
  description: string;
}

type MemberState = UserPermissionInfo[] | 'loading' | 'error';

// ── Props ─────────────────────────────────────────────────────────────────────

export interface PermissionGroupsViewProps {
  sp: SharePointService;
  siteUrl: string;
  onBack: () => void;
}

// ── Helper ────────────────────────────────────────────────────────────────────

function principalIcon(pt: string): React.ReactElement {
  if (pt === 'SharePointGroup') return <PeopleTeam16Regular style={{ flexShrink: 0 }} />;
  if (pt === 'SecurityGroup') return <Shield16Regular style={{ flexShrink: 0 }} />;
  return <Person16Regular style={{ flexShrink: 0 }} />;
}

function principalColor(pt: string): string {
  if (pt === 'SharePointGroup') return 'brand';
  if (pt === 'SecurityGroup') return 'informative';
  return 'subtle';
}

// ── Component ─────────────────────────────────────────────────────────────────

export const PermissionGroupsView: React.FC<PermissionGroupsViewProps> = ({ sp, siteUrl, onBack }) => {
  const styles = useStyles();

  const [groups, setGroups] = React.useState<GroupInfo[] | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [expandedIds, setExpandedIds] = React.useState<Set<number>>(new Set());
  const [memberCache, setMemberCache] = React.useState<Map<number, MemberState>>(new Map());

  const abortRef = React.useRef<AbortController | null>(null);

  React.useEffect(() => {
    if (!siteUrl) return;

    abortRef.current?.abort();
    abortRef.current = new AbortController();

    setGroups(null);
    setError('');
    setExpandedIds(new Set());
    setMemberCache(new Map());
    setLoading(true);

    sp.getSiteGroups(siteUrl.trim(), abortRef.current.signal)
      .then((result) => {
        setGroups(result);
        setLoading(false);
      })
      .catch((err: any) => {
        setError(err?.message ?? String(err));
        setLoading(false);
      });

    return () => { abortRef.current?.abort(); };
  }, [siteUrl]);

  const handleToggle = async (group: GroupInfo): Promise<void> => {
    const isExpanded = expandedIds.has(group.id);

    const newExpanded = new Set(expandedIds);
    if (isExpanded) {
      newExpanded.delete(group.id);
    } else {
      newExpanded.add(group.id);
    }
    setExpandedIds(newExpanded);

    if (!isExpanded && !memberCache.has(group.id)) {
      setMemberCache((prev) => new Map(prev).set(group.id, 'loading'));
      try {
        const members = await sp.getGroupMembers(
          siteUrl.trim(),
          group.title,
          group.loginName,
          'SharePointGroup',
        );
        setMemberCache((prev) => new Map(prev).set(group.id, members));
      } catch {
        setMemberCache((prev) => new Map(prev).set(group.id, 'error'));
      }
    }
  };

  if (!siteUrl) {
    return (
      <div className={styles.root}>
        <div className={styles.header}>
          <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack}>Back</Button>
          <Title3 style={{ flex: 1 }}>Permission Groups</Title3>
        </div>
        <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
          Connect to a SharePoint site first, then open this tool to browse its permission groups.
        </Body1>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onBack} aria-label="Back to home">
          Back
        </Button>
        <Title3 style={{ flex: 1 }}>Permission Groups</Title3>
      </div>

      {loading && (
        <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS }}>
          <Spinner size="small" />
          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>Loading groups…</Body1>
        </div>
      )}

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {!loading && !error && groups !== null && groups.length === 0 && (
        <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No SharePoint groups found on this site.</Body1>
      )}

      {!loading && !error && groups !== null && groups.length > 0 && (
        <>
          <Body1 style={{ color: tokens.colorNeutralForeground3, marginBottom: tokens.spacingVerticalM, display: 'block' }}>
            {groups.length} group{groups.length !== 1 ? 's' : ''} — click any group to see its members.
          </Body1>

          <div className={styles.groupList}>
            {groups.map((group) => {
              const isExpanded = expandedIds.has(group.id);
              const members = memberCache.get(group.id);

              return (
                <div key={group.id} className={styles.groupRow}>
                  <div
                    className={styles.groupHeader}
                    onClick={() => void handleToggle(group)}
                    role="button"
                    tabIndex={0}
                    aria-expanded={isExpanded}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); void handleToggle(group); } }}
                  >
                    {isExpanded
                      ? <ChevronDown16Regular style={{ flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
                      : <ChevronRight16Regular style={{ flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
                    }
                    <People24Regular style={{ flexShrink: 0, color: tokens.colorBrandForeground1 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <Text className={styles.groupTitle}>{group.title}</Text>
                      {group.description && (
                        <Text className={styles.groupDesc} block>{group.description}</Text>
                      )}
                    </div>
                    {members !== undefined && members !== 'loading' && members !== 'error' && (
                      <Badge appearance="outline" size="small">
                        {members.length} member{members.length !== 1 ? 's' : ''}
                      </Badge>
                    )}
                    {members === 'loading' && <Spinner size="tiny" />}
                  </div>

                  {isExpanded && (
                    <div className={styles.memberList}>
                      {members === 'loading' && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, padding: '4px' }}>
                          <Spinner size="tiny" />
                          <Body1 style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 }}>
                            Loading members…
                          </Body1>
                        </div>
                      )}
                      {members === 'error' && (
                        <Body1 style={{ color: tokens.colorPaletteRedForeground1, fontSize: tokens.fontSizeBase200, padding: '4px' }}>
                          Could not load members for this group.
                        </Body1>
                      )}
                      {members !== undefined && members !== 'loading' && members !== 'error' && members.length === 0 && (
                        <Body1 style={{ color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200, padding: '4px' }}>
                          This group has no members.
                        </Body1>
                      )}
                      {members !== undefined && members !== 'loading' && members !== 'error' && members.length > 0 && (
                        members.map((m, i) => (
                          <div key={`${m.loginName}-${i}`} className={styles.memberRow}>
                            {principalIcon(m.principalType)}
                            <Text className={styles.memberName}>{m.displayName}</Text>
                            <Badge
                              appearance="tint"
                              color={principalColor(m.principalType) as any}
                              size="small"
                            >
                              {m.principalType}
                            </Badge>
                            <Text className={styles.loginName} title={m.loginName}>{m.loginName}</Text>
                          </div>
                        ))
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
};
