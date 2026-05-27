import * as React from 'react';
import {
  Button,
  Badge,
  Text,
  Title3,
  Body1,
  Spinner,
  Input,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  People24Regular,
  Person24Regular,
  ChevronRight16Regular,
  ChevronDown16Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import { PermissionGroup } from '../models/models';

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '900px',
    margin: '0 auto',
    minHeight: '500px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  groupCard: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: tokens.spacingVerticalS,
    overflow: 'hidden',
  },
  groupHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    cursor: 'pointer',
    background: tokens.colorNeutralBackground2,
    ':hover': {
      background: tokens.colorNeutralBackground2Hover,
    },
  },
  memberRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalL}`,
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalXXL,
    color: tokens.colorNeutralForeground3,
    textAlign: 'center',
  },
});

function roleBadgeColor(roles: string[]): 'danger' | 'warning' | 'success' | 'informative' {
  if (roles.some((r) => r.toLowerCase().includes('full control'))) return 'danger';
  if (
    roles.some(
      (r) =>
        r.toLowerCase().includes('edit') ||
        r.toLowerCase().includes('contribute') ||
        r.toLowerCase().includes('design'),
    )
  )
    return 'warning';
  if (roles.some((r) => r.toLowerCase().includes('read') || r.toLowerCase().includes('view')))
    return 'success';
  return 'informative';
}

// ── Single group row (expandable) ─────────────────────────────────────────────

interface GroupRowProps {
  group: PermissionGroup;
  styles: ReturnType<typeof useStyles>;
  expanded: boolean;
  onToggle: () => void;
}

const GroupRow: React.FC<GroupRowProps> = ({ group, styles, expanded, onToggle }) => {
  return (
    <div className={styles.groupCard}>
      <div
        className={styles.groupHeader}
        onClick={onToggle}
        role="button"
        tabIndex={0}
        aria-expanded={expanded}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            onToggle();
          }
        }}
      >
        {expanded ? (
          <ChevronDown16Regular style={{ flexShrink: 0 }} />
        ) : (
          <ChevronRight16Regular style={{ flexShrink: 0 }} />
        )}
        <People24Regular style={{ flexShrink: 0, fontSize: '18px' }} />
        <Text weight="semibold" style={{ flex: 1 }}>
          {group.title}
        </Text>
        <Text
          size={200}
          style={{ color: tokens.colorNeutralForeground3, marginRight: tokens.spacingHorizontalS }}
        >
          {group.memberCount} member{group.memberCount !== 1 ? 's' : ''}
        </Text>
        {group.roles.length > 0
          ? group.roles.map((r) => (
              <Badge
                key={r}
                appearance="filled"
                color={roleBadgeColor([r])}
                size="small"
                style={{ marginLeft: '2px' }}
              >
                {r}
              </Badge>
            ))
          : (
            <Badge appearance="outline" color="subtle" size="small">
              No site role
            </Badge>
          )}
      </div>

      {expanded && (
        <div>
          {group.description && (
            <div
              style={{
                padding: `${tokens.spacingVerticalXS} ${tokens.spacingHorizontalL}`,
                borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
              }}
            >
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                {group.description}
              </Text>
            </div>
          )}
          {group.members.length === 0 ? (
            <div className={styles.memberRow}>
              <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
                No members
              </Text>
            </div>
          ) : (
            group.members.map((m) => (
              <div key={m.loginName} className={styles.memberRow}>
                <Person24Regular
                  style={{
                    flexShrink: 0,
                    fontSize: '14px',
                    color: tokens.colorNeutralForeground3,
                  }}
                />
                <Text size={200} style={{ flex: 1 }}>
                  {m.displayName}
                </Text>
                <Text
                  size={200}
                  style={{
                    color: tokens.colorNeutralForeground3,
                    fontSize: tokens.fontSizeBase100,
                    wordBreak: 'break-all',
                    maxWidth: '40%',
                  }}
                >
                  {m.loginName}
                </Text>
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
};

// ── Main view ─────────────────────────────────────────────────────────────────

export interface PermissionGroupsViewProps {
  sp: SharePointService;
  siteUrl: string;
  onBack: () => void;
}

export const PermissionGroupsView: React.FC<PermissionGroupsViewProps> = ({
  sp,
  siteUrl,
  onBack,
}) => {
  const styles = useStyles();

  const [isLoading, setIsLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [groups, setGroups] = React.useState<PermissionGroup[] | null>(null);
  const [filter, setFilter] = React.useState('');
  const [expandedIds, setExpandedIds] = React.useState<{ [id: number]: boolean }>({});
  const abortRef = React.useRef<AbortController | null>(null);

  const handleLoad = async (): Promise<void> => {
    abortRef.current?.abort();
    abortRef.current = new AbortController();
    setIsLoading(true);
    setError('');
    setGroups(null);

    try {
      const result = await sp.getPermissionGroups(siteUrl.trim(), abortRef.current.signal);
      setGroups(result);
    } catch (err: any) {
      setError(`Failed to load groups: ${err?.message ?? String(err)}`);
    } finally {
      setIsLoading(false);
    }
  };

  React.useEffect(() => {
    handleLoad().catch(console.error);
  }, []);

  const filteredGroups = React.useMemo(() => {
    if (!groups) return [];
    if (!filter.trim()) return groups;
    const lc = filter.toLowerCase();
    return groups.filter(
      (g) =>
        g.title.toLowerCase().includes(lc) ||
        g.description.toLowerCase().includes(lc) ||
        g.roles.some((r) => r.toLowerCase().includes(lc)) ||
        g.members.some(
          (m) =>
            m.displayName.toLowerCase().includes(lc) ||
            m.loginName.toLowerCase().includes(lc),
        ),
    );
  }, [groups, filter]);

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={onBack}
          aria-label="Back to home"
        >
          Back
        </Button>
        <Title3 style={{ flex: 1 }}>Permission Groups</Title3>
        <Button appearance="primary" onClick={handleLoad} disabled={isLoading}>
          {isLoading ? 'Loading…' : 'Reload'}
        </Button>
      </div>

      {error && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {isLoading && (
        <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM }}>
          <Spinner size="small" />
          <Body1>Loading permission groups…</Body1>
        </div>
      )}

      {!isLoading && groups !== null && (
        <>
          <div style={{ marginBottom: tokens.spacingVerticalM }}>
            <Input
              placeholder="Filter groups, members, or roles…"
              value={filter}
              onChange={(_, d) => setFilter(d.value)}
              style={{ width: '100%', maxWidth: '400px' }}
              aria-label="Filter groups"
            />
          </div>

          {filteredGroups.length === 0 ? (
            <div className={styles.emptyState}>
              <People24Regular style={{ fontSize: '48px', opacity: 0.4 }} />
              <Text weight="semibold">
                {groups.length === 0 ? 'No groups found' : 'No groups match the filter'}
              </Text>
              {groups.length === 0 && (
                <Body1>
                  This site has no SharePoint permission groups, or the current account does not
                  have permission to list them.
                </Body1>
              )}
            </div>
          ) : (
            <>
              <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalM, marginBottom: tokens.spacingVerticalS }}>
                <Body1 style={{ color: tokens.colorNeutralForeground3, flex: 1 }}>
                  {filteredGroups.length} group{filteredGroups.length !== 1 ? 's' : ''}
                  {filter && ` matching "${filter}"`}
                </Body1>
                <Button
                  appearance="subtle"
                  size="small"
                  onClick={() => {
                    const allExpanded = filteredGroups.every((g) => expandedIds[g.id]);
                    if (allExpanded) {
                      setExpandedIds({});
                    } else {
                      const next: { [id: number]: boolean } = {};
                      filteredGroups.forEach((g) => { next[g.id] = true; });
                      setExpandedIds(next);
                    }
                  }}
                >
                  {filteredGroups.every((g) => expandedIds[g.id]) ? 'Collapse all' : 'Expand all'}
                </Button>
              </div>
              {filteredGroups.map((g) => (
                <GroupRow
                  key={g.id}
                  group={g}
                  styles={styles}
                  expanded={!!expandedIds[g.id]}
                  onToggle={() => setExpandedIds((prev) => ({ ...prev, [g.id]: !prev[g.id] }))}
                />
              ))}
            </>
          )}
        </>
      )}
    </div>
  );
};
