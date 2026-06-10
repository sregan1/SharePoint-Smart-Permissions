import * as React from 'react';
import {
  Button,
  Field,
  Badge,
  Link,
  Text,
  Title3,
  Body1,
  Spinner,
  Checkbox,
  MessageBar,
  MessageBarBody,
  Select,
  Tooltip,
  ToggleButton,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  Folder24Regular,
  Document24Regular,
  Person24Regular,
  People24Regular,
  ChevronRight16Regular,
  ChevronDown16Regular,
  ArrowCircleDown16Regular,
  Link16Regular,
  PersonSearch16Regular,
  PersonWarning16Regular,
  Filter16Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import {
  LibraryInfo,
  FolderFileNode,
  UserPermissionInfo,
} from '../models/models';

// ── Styles ────────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    padding: tokens.spacingVerticalL,
    maxWidth: '1100px',
    margin: '0 auto',
    minHeight: '500px',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  twoCol: {
    display: 'grid',
    gridTemplateColumns: '280px 1fr',
    gap: tokens.spacingHorizontalM,
    minHeight: '400px',
    '@media (max-width: 700px)': {
      gridTemplateColumns: '1fr',
    },
  },
  treePanel: {
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingRight: tokens.spacingHorizontalM,
    overflowY: 'auto',
    maxHeight: '600px',
  },
  permPanel: {
    paddingLeft: tokens.spacingHorizontalM,
    overflowY: 'auto',
    maxHeight: '600px',
  },
  treeNode: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    padding: '3px 4px',
    borderRadius: tokens.borderRadiusSmall,
    cursor: 'pointer',
    userSelect: 'none',
    ':hover': {
      background: tokens.colorNeutralBackground1Hover,
    },
  },
  treeNodeSelected: {
    background: tokens.colorNeutralBackground1Selected,
  },
  permTable: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: tokens.fontSizeBase200,
  },
  permTh: {
    textAlign: 'left',
    padding: '6px 8px',
    borderBottom: `2px solid ${tokens.colorNeutralStroke1}`,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    whiteSpace: 'nowrap',
    position: 'sticky',
    top: 0,
    background: tokens.colorNeutralBackground1,
    zIndex: 1,
  },
  permTd: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'top',
  },
  optionsBar: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS,
    marginBottom: tokens.spacingVerticalM,
  },
  inheritedBanner: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    background: tokens.colorBrandBackground2,
    borderLeft: `3px solid ${tokens.colorBrandForeground1}`,
    borderRadius: tokens.borderRadiusSmall,
    marginBottom: tokens.spacingVerticalM,
  },
});

// ── Custom indicator icons ────────────────────────────────────────────────────

// Arrow inside a triangle — indicates external user access below
function ArrowTriangleDown({ style }: { style?: React.CSSProperties }): React.ReactElement {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" style={style}>
      <path d="M8 2L14.5 13.5H1.5L8 2Z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round"/>
      <path d="M8 6v4M6 8.5L8 10.5L10 8.5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  );
}

// Circle and triangle icons side by side — indicates both unique permissions and external
// user access below; self-explanatory from the two individual legend entries
function ArrowCircleAndTriangleDown({ style }: { style?: React.CSSProperties }): React.ReactElement {
  return (
    <svg width="28" height="14" viewBox="0 0 30 16" fill="none" style={style}>
      <circle cx="7" cy="8" r="5.5" stroke="currentColor" strokeWidth="1.5"/>
      <path d="M7 5v5M5.5 8.5L7 10L8.5 8.5" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M22 2L29 14H15L22 2Z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round"/>
      <path d="M22 6v5M20.5 9L22 10.5L23.5 9" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round"/>
    </svg>
  );
}


// ── External user helpers ─────────────────────────────────────────────────────

function isExternalUser(u: UserPermissionInfo): boolean {
  return u.loginName.toLowerCase().indexOf('#ext#') !== -1;
}

// Extract the email from a SharePoint #EXT# login name.
// Format: [claims-prefix|]localpart_domain.tld#EXT#@tenant.onmicrosoft.com
// The last underscore before #EXT# is the @ in the original email address.
function externalUserEmail(loginName: string): string {
  const extIdx = loginName.toLowerCase().indexOf('#ext#');
  if (extIdx === -1) return '';
  let local = loginName.substring(0, extIdx);
  const pipeIdx = local.lastIndexOf('|');
  if (pipeIdx >= 0) local = local.substring(pipeIdx + 1);
  const lastUnderscore = local.lastIndexOf('_');
  if (lastUnderscore === -1) return local;
  return `${local.substring(0, lastUnderscore)}@${local.substring(lastUnderscore + 1)}`;
}

function applyPermFilters(
  users: UserPermissionInfo[],
  excludeLimited: boolean,
  extOnly: boolean,
): UserPermissionInfo[] {
  let result = users;
  if (excludeLimited) result = result.filter((u) => u.roles.length > 0);
  if (extOnly) result = result.filter(isExternalUser);
  return result;
}

// ── Role badge color ──────────────────────────────────────────────────────────

function roleBadgeColor(
  roles: string[],
): 'brand' | 'danger' | 'warning' | 'success' | 'informative' {
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
  if (
    roles.some(
      (r) => r.toLowerCase().includes('read') || r.toLowerCase().includes('view'),
    )
  )
    return 'success';
  return 'informative';
}

// ── Permission table ──────────────────────────────────────────────────────────

interface PermTableProps {
  users: UserPermissionInfo[];
  styles: ReturnType<typeof useStyles>;
  onCheckAccess?: (loginName: string) => void;
}

const PermTable: React.FC<PermTableProps> = ({ users, styles, onCheckAccess }) => (
  <table className={styles.permTable} aria-label="Permission assignments">
    <thead>
      <tr>
        <th className={styles.permTh}>User / Group</th>
        <th className={styles.permTh}>Type</th>
        <th className={styles.permTh}>Permission Level</th>
        {onCheckAccess && <th className={styles.permTh} />}
      </tr>
    </thead>
    <tbody>
      {users.map((u, i) => (
        <tr key={i}>
          <td className={styles.permTd}>
            {u.isGroupMember ? (
              <span style={{ paddingLeft: '16px', color: tokens.colorNeutralForeground3 }}>
                ↳ {u.displayName}
              </span>
            ) : (
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                {u.principalType === 'User' ? (
                  <Person24Regular style={{ fontSize: '14px', flexShrink: 0 }} />
                ) : (
                  <People24Regular style={{ fontSize: '14px', flexShrink: 0 }} />
                )}
                <span>
                  <span>{u.displayName || u.loginName}</span>
                  {isExternalUser(u) && (() => {
                    const email = externalUserEmail(u.loginName);
                    return email && email !== u.displayName ? (
                      <div style={{ fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 }}>
                        {email}
                      </div>
                    ) : null;
                  })()}
                </span>
              </span>
            )}
          </td>
          <td className={styles.permTd}>
            <Text style={{ fontSize: tokens.fontSizeBase200 }}>
              {u.principalType === 'SecurityGroup'
                ? 'Security Group'
                : u.principalType === 'SharePointGroup'
                ? 'SP Group'
                : 'User'}
            </Text>
          </td>
          <td className={styles.permTd}>
            {u.roles.map((r, ri) => (
              <Badge
                key={ri}
                appearance="filled"
                color={roleBadgeColor([r])}
                size="small"
                style={{ marginRight: '4px', marginBottom: '2px' }}
              >
                {r}
              </Badge>
            ))}
          </td>
          {onCheckAccess && (
            <td className={styles.permTd}>
              {u.principalType === 'User' && u.loginName && !u.isGroupMember && (
                <Button
                  appearance="subtle"
                  size="small"
                  icon={<PersonSearch16Regular />}
                  onClick={() => onCheckAccess(u.loginName)}
                  title={`Check access for ${u.displayName || u.loginName}`}
                  aria-label={`Check access for ${u.displayName || u.loginName}`}
                />
              )}
            </td>
          )}
        </tr>
      ))}
    </tbody>
  </table>
);

// ── Folder tree node ──────────────────────────────────────────────────────────

interface TreeNodeProps {
  node: FolderFileNode;
  depth: number;
  selectedUrl: string;
  onSelect: (node: FolderFileNode) => void;
  onLoadChildren: (node: FolderFileNode) => void;
  showUniqueOnly: boolean;
  filterExternalOnly: boolean;
  externalAccessUrls: Set<string>;
}

const TreeNode: React.FC<TreeNodeProps> = ({
  node,
  depth,
  selectedUrl,
  onSelect,
  onLoadChildren,
  showUniqueOnly,
  filterExternalOnly,
  externalAccessUrls,
}) => {
  const styles = useStyles();
  const [expanded, setExpanded] = React.useState(false);

  const isSelected = node.serverRelativeUrl === selectedUrl;

  const handleToggle = (e: React.MouseEvent): void => {
    e.stopPropagation();
    if (node.isFolder && node.hasChildren) {
      if (!expanded && node.children.length === 0) {
        onLoadChildren(node);
      }
      setExpanded(!expanded);
    }
    onSelect(node);
  };

  return (
    <div>
      <div
        className={`${styles.treeNode} ${isSelected ? styles.treeNodeSelected : ''}`}
        style={{ paddingLeft: `${depth * 16 + 4}px` }}
        onClick={handleToggle}
      >
        {node.isFolder && node.hasChildren ? (
          expanded ? (
            <ChevronDown16Regular style={{ flexShrink: 0 }} />
          ) : (
            <ChevronRight16Regular style={{ flexShrink: 0 }} />
          )
        ) : (
          <span style={{ width: '16px', flexShrink: 0 }} />
        )}

        {node.isFolder ? (
          <Folder24Regular style={{ fontSize: '16px', flexShrink: 0 }} />
        ) : (
          <Document24Regular style={{ fontSize: '16px', flexShrink: 0 }} />
        )}

        <Text style={{ flexGrow: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {node.name}
        </Text>

        {node.hasUniquePermissionsBelow && node.hasExternalUsersBelow && (
          <Tooltip content="Contains items with unique permissions and external user access" relationship="label">
            <ArrowCircleAndTriangleDown style={{ flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
          </Tooltip>
        )}
        {node.hasUniquePermissionsBelow && !node.hasExternalUsersBelow && (
          <Tooltip content="Contains items with unique permissions" relationship="label">
            <ArrowCircleDown16Regular style={{ flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
          </Tooltip>
        )}
        {node.hasExternalUsersBelow && !node.hasUniquePermissionsBelow && (
          <Tooltip content="Contains items with external user access" relationship="label">
            <ArrowTriangleDown style={{ flexShrink: 0, color: tokens.colorNeutralForeground3 }} />
          </Tooltip>
        )}
        {node.hasUniquePermissions && (
          <Badge appearance="filled" color="warning" size="small" style={{ flexShrink: 0 }}>
            Unique
          </Badge>
        )}
        {node.hasUniquePermissions && (node.hasExternalUsers || externalAccessUrls.has(node.serverRelativeUrl)) && (
          <Tooltip content="External user access detected" relationship="label">
            <PersonWarning16Regular
              style={{ flexShrink: 0, color: tokens.colorPaletteRedForeground1 }}
            />
          </Tooltip>
        )}
        {node.isLoading && (
          <Spinner size="extra-tiny" style={{ flexShrink: 0 }} />
        )}
      </div>

      {expanded && node.children
        .filter((c) =>
          (!showUniqueOnly || c.hasUniquePermissions || c.hasUniquePermissionsBelow) &&
          (!filterExternalOnly || c.hasExternalUsers || c.hasExternalUsersBelow)
        )
        .map((child) => (
          <TreeNode
            key={child.serverRelativeUrl}
            node={child}
            depth={depth + 1}
            selectedUrl={selectedUrl}
            onSelect={onSelect}
            onLoadChildren={onLoadChildren}
            showUniqueOnly={showUniqueOnly}
            filterExternalOnly={filterExternalOnly}
            externalAccessUrls={externalAccessUrls}
          />
        ))}
    </div>
  );
};

// ── Main view ─────────────────────────────────────────────────────────────────

export interface PermissionsExplorerViewProps {
  sp: SharePointService;
  siteUrl: string;
  includeHidden: boolean;
  excludeLimitedAccess: boolean;
  onBack: () => void;
  onNavigateToUserAccess?: (loginName: string) => void;
}

export const PermissionsExplorerView: React.FC<PermissionsExplorerViewProps> = ({ sp, siteUrl, includeHidden, excludeLimitedAccess, onBack, onNavigateToUserAccess }) => {
  const styles = useStyles();

  // ── Connection ──
  const [isConnecting, setIsConnecting] = React.useState(false);
  const [connectStatus, setConnectStatus] = React.useState('');
  const [connectError, setConnectError] = React.useState('');
  const [isConnected, setIsConnected] = React.useState(false);

  // ── Libraries ──
  const [libraries, setLibraries] = React.useState<LibraryInfo[]>([]);

  // ── Browse tab ──
  const [selectedLibrary, setSelectedLibrary] = React.useState('');
  const [rootNodes, setRootNodes] = React.useState<FolderFileNode[]>([]);
  const [showUniqueOnly, setShowUniqueOnly] = React.useState(false);
  const [filterExternalOnly, setFilterExternalOnly] = React.useState(false);
  const [externalAccessUrls, setExternalAccessUrls] = React.useState<Set<string>>(new Set());
  const [permissionsDenied, setPermissionsDenied] = React.useState(false);
  const [myPermLevel, setMyPermLevel] = React.useState('');
  const [siteOwners, setSiteOwners] = React.useState<{ title: string; email: string }[]>([]);
  const [treeStatus, setTreeStatus] = React.useState('');
  const [selectedNode, setSelectedNode] = React.useState<FolderFileNode | null>(null);
  const [nodePerms, setNodePerms] = React.useState<UserPermissionInfo[]>([]);
  const [nodeHasUnique, setNodeHasUnique] = React.useState(false);
  const [nodeLoading, setNodeLoading] = React.useState(false);
  const [nodeError, setNodeError] = React.useState('');
  const [expandGroups, setExpandGroups] = React.useState(false);
  const [showParentPerms, setShowParentPerms] = React.useState(false);
  const [parentPermsLoading, setParentPermsLoading] = React.useState(false);
  const [parentPerms, setParentPerms] = React.useState<UserPermissionInfo[] | null>(null);
  const [parentPermsName, setParentPermsName] = React.useState('');
  const [parentPermsError, setParentPermsError] = React.useState('');

  const abortRef = React.useRef<AbortController | null>(null);
  const rawNodePermsRef = React.useRef<UserPermissionInfo[]>([]);
  const rawParentPermsRef = React.useRef<UserPermissionInfo[] | null>(null);
  const groupMemberCacheRef = React.useRef<Map<string, UserPermissionInfo[]>>(new Map());
  const folderCacheRef = React.useRef<Map<string, FolderFileNode[]>>(new Map());

  // ── Connect ──────────────────────────────────────────────────────────────

  const handleConnect = async (): Promise<void> => {
    abortRef.current?.abort();
    abortRef.current = new AbortController();

    setIsConnecting(true);
    setConnectError('');
    setConnectStatus('Connecting…');
    setIsConnected(false);
    setLibraries([]);
    setRootNodes([]);
    setTreeStatus('');

    try {
      const libs = await sp.getLibraries(siteUrl.trim(), abortRef.current.signal, includeHidden);

      setLibraries(libs);
      setIsConnected(true);
      setConnectStatus(
        `Connected — ${libs.length} librar${libs.length === 1 ? 'y' : 'ies'} found`,
      );

      if (libs.length > 0) {
        setSelectedLibrary(libs[0].serverRelativeUrl);
        loadLibrary(libs[0]).catch((e) => console.error('[SmartPermissions] loadLibrary failed:', e));
      }
    } catch (err: any) {
      setConnectError(`Connection failed: ${err?.message ?? String(err)}`);
      setConnectStatus('');
    } finally {
      setIsConnecting(false);
    }
  };

  // ── Load library root ────────────────────────────────────────────────────

  const loadLibrary = async (lib: LibraryInfo): Promise<void> => {
    setRootNodes([]);
    setTreeStatus('');
    setSelectedNode(null);
    setNodePerms([]);
    setExternalAccessUrls(new Set());
    setPermissionsDenied(false);
    setMyPermLevel('');
    setSiteOwners([]);
    rawNodePermsRef.current = [];
    rawParentPermsRef.current = null;
    groupMemberCacheRef.current.clear();
    folderCacheRef.current.clear();

    try {
      const nodes = await sp.getFolderContents(
        siteUrl.trim(),
        lib.serverRelativeUrl,
        abortRef.current?.signal,
      );
      setRootNodes(nodes);
      if (nodes.length === 0) setTreeStatus('This library is empty.');

      // Proactively check whether role-assignment reads are permitted for this library.
      sp.getItemPermissions(
        siteUrl.trim(),
        { name: lib.title, serverRelativeUrl: lib.serverRelativeUrl, isFolder: true, hasChildren: false, children: [] },
        abortRef.current?.signal,
      ).then((probe) => {
        if (probe.permissionDenied) setPermissionsDenied(true);
      }).catch(() => { /* ignore probe errors */ });

      // Background scan for external users on root-level nodes.
      scanExternalUsers(nodes);

      // Pre-fetch one level deep for each root folder to detect hasUniquePermissionsBelow
      // at load time rather than requiring the user to expand each folder first.
      // Results are stored in folderCacheRef so expanding a folder skips the API call.
      const signal = abortRef.current?.signal;
      nodes.filter((n) => n.isFolder).forEach((folder) => {
        sp.getFolderContents(siteUrl.trim(), folder.serverRelativeUrl, signal)
          .then((children) => {
            // Set parent now so propagation works correctly when the external-user scan fires.
            children.forEach((c) => { c.parent = folder; });
            folderCacheRef.current.set(folder.serverRelativeUrl, children);
            if (children.some((c) => c.hasUniquePermissions)) {
              folder.hasUniquePermissionsBelow = true;
              setRootNodes((prev) => [...prev]);
            }
            // Scan pre-fetched children so icons appear at library-load time,
            // not only after the user expands the folder.
            scanExternalUsers(children);
          })
          .catch(() => { /* ignore prefetch errors */ });
      });
    } catch (err: any) {
      setTreeStatus(`Error loading library: ${err?.message ?? String(err)}`);
    }
  };

  const handleLibraryChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    const url = e.target.value;
    setSelectedLibrary(url);
    const lib = libraries.find((l) => l.serverRelativeUrl === url);
    if (lib) void loadLibrary(lib);
  };

  // ── Load folder children ─────────────────────────────────────────────────

  const loadChildren = React.useCallback(async (node: FolderFileNode): Promise<void> => {
    node.isLoading = true;
    setRootNodes((prev) => [...prev]);

    try {
      const cached = folderCacheRef.current.get(node.serverRelativeUrl);
      const children = cached ?? await sp.getFolderContents(
        siteUrl.trim(),
        node.serverRelativeUrl,
        abortRef.current?.signal,
      );
      if (!cached) folderCacheRef.current.set(node.serverRelativeUrl, children);
      node.children = children.map((c) => {
        c.parent = node;
        return c;
      });

      if (children.some((c) => c.hasUniquePermissions)) {
        propagateUniqueBelow(node);
      }

      // If this folder has external user access, inherited children share that access.
      if (node.hasExternalUsers) {
        for (const child of children) {
          if (!child.hasUniquePermissions) child.hasExternalUsers = true;
        }
      }

      // Background scan for external users on newly loaded children.
      scanExternalUsers(children);

      // Pre-fetch one level deeper for each folder child so their icons are ready
      // before the user expands them — same pattern used in loadLibrary for depth-0.
      const prefetchSignal = abortRef.current?.signal;
      children.filter((c) => c.isFolder && c.hasChildren).forEach((folder) => {
        if (folderCacheRef.current.has(folder.serverRelativeUrl)) return;
        sp.getFolderContents(siteUrl.trim(), folder.serverRelativeUrl, prefetchSignal)
          .then((grandchildren) => {
            grandchildren.forEach((gc) => { gc.parent = folder; });
            folderCacheRef.current.set(folder.serverRelativeUrl, grandchildren);
            if (grandchildren.some((gc) => gc.hasUniquePermissions)) {
              propagateUniqueBelow(folder);
              setRootNodes((prev) => [...prev]);
            }
            scanExternalUsers(grandchildren);
          })
          .catch(() => { /* ignore */ });
      });
    } catch {
      node.children = [
        {
          name: 'Error loading contents',
          serverRelativeUrl: '',
          isFolder: false,
          hasChildren: false,
          children: [],
          isLoading: false,
        },
      ];
    } finally {
      node.isLoading = false;
      setRootNodes((prev) => [...prev]);
    }
  }, [siteUrl]);

  const propagateUniqueBelow = (node: FolderFileNode): void => {
    node.hasUniquePermissionsBelow = true;
    if (node.parent) propagateUniqueBelow(node.parent);
  };

  const propagateExternalBelow = (node: FolderFileNode): void => {
    node.hasExternalUsersBelow = true;
    if (node.parent) propagateExternalBelow(node.parent);
  };

  // Propagate external-user access DOWN through already-loaded descendants.
  // Stops at unique-permission boundaries — those have their own scope and are scanned separately.
  const propagateExternalDown = (node: FolderFileNode): void => {
    for (const child of node.children) {
      if (!child.hasUniquePermissions) {
        child.hasExternalUsers = true;
        propagateExternalDown(child);
      }
    }
  };

  // Background scan: checks only unique-permission nodes (inherited ones are skipped — no API
  // call needed). Uses one direct RoleAssignments fetch per node instead of two calls.
  const scanExternalUsers = (nodes: FolderFileNode[]): void => {
    const uniqueNodes = nodes.filter((n) => n.hasUniquePermissions);
    if (!uniqueNodes.length) return;
    const tasks = uniqueNodes.map((node) => async (): Promise<undefined> => {
      if (abortRef.current?.signal.aborted) return undefined;
      const hasExt = await sp.scanNodeForExternalUsers(siteUrl, node, abortRef.current?.signal);
      if (hasExt === 'denied') {
        setPermissionsDenied(true);
      } else if (hasExt) {
        node.hasExternalUsers = true;
        if (node.parent) propagateExternalBelow(node.parent);
        propagateExternalDown(node);  // mark already-loaded descendants that inherit
        setExternalAccessUrls((prev) => { const next = new Set(Array.from(prev)); next.add(node.serverRelativeUrl); return next; });
        setRootNodes((prev) => [...prev]);
      }
      return undefined;
    });
    sp.runConcurrent(tasks, sp.scanConcurrency).catch(() => { /* ignore — background scan */ });
  };

  // ── Expand groups helper ──────────────────────────────────────────────────

  const withGroupExpansion = React.useCallback(async (users: UserPermissionInfo[]): Promise<UserPermissionInfo[]> => {
    // Fetch all group members in parallel, using a session cache to avoid repeat calls
    await Promise.all(
      users
        .filter((u) => expandGroups && (u.principalType === 'SharePointGroup' || u.principalType === 'SecurityGroup'))
        .map(async (u) => {
          const key = u.loginName || u.displayName;
          if (!groupMemberCacheRef.current.has(key)) {
            const members = await sp.getGroupMembers(
              siteUrl.trim(),
              u.displayName,
              u.loginName,
              u.principalType,
              abortRef.current?.signal,
            );
            groupMemberCacheRef.current.set(key, members);
          }
        }),
    );

    const expanded: UserPermissionInfo[] = [];
    for (const u of users) {
      expanded.push(u);
      if (expandGroups && (u.principalType === 'SharePointGroup' || u.principalType === 'SecurityGroup')) {
        const key = u.loginName || u.displayName;
        const members = groupMemberCacheRef.current.get(key) ?? [];
        members.forEach((m) => expanded.push({ ...m, roles: [...u.roles] }));
      }
    }
    return expanded;
  }, [expandGroups, siteUrl]);

  // ── Select node ──────────────────────────────────────────────────────────

  const handleSelectNode = async (node: FolderFileNode): Promise<void> => {
    const keepParentPerms = showParentPerms;
    setSelectedNode(node);
    setNodeLoading(true);
    setNodePerms([]);
    setNodeHasUnique(false);
    setNodeError('');
    setMyPermLevel('');
    setParentPerms(null);
    setParentPermsName('');
    setParentPermsError('');
    rawNodePermsRef.current = [];
    rawParentPermsRef.current = null;

    try {
      const { hasUnique, users, permissionDenied } = await sp.getItemPermissions(
        siteUrl.trim(),
        node,
        abortRef.current?.signal,
      );
      if (permissionDenied) {
        setPermissionsDenied(true);
        sp.getEffectivePermissions(siteUrl.trim(), node, abortRef.current?.signal)
          .then(setMyPermLevel)
          .catch(() => {});
      }
      rawNodePermsRef.current = users;
      if (users.some(isExternalUser)) {
        setExternalAccessUrls((prev) => {
          const next = new Set(Array.from(prev));
          next.add(node.serverRelativeUrl);
          return next;
        });
      }
      const resolvedHasUnique = permissionDenied ? (node.hasUniquePermissions ?? false) : hasUnique;
      setNodeHasUnique(resolvedHasUnique);
      setNodePerms(await withGroupExpansion(users));
      if (!resolvedHasUnique && keepParentPerms) {
        handleShowParentPerms(node).catch((e) => console.error('[SmartPermissions] handleShowParentPerms failed:', e));
      }
    } catch (err: any) {
      setNodeError(err?.message ?? String(err));
    } finally {
      setNodeLoading(false);
    }
  };

  React.useEffect(() => {
    if (!selectedNode) return;
    // Re-expand node permissions (only relevant when the node has unique permissions).
    if (rawNodePermsRef.current.length > 0) {
      setNodeLoading(true);
      setNodePerms([]);
      setNodeError('');
      withGroupExpansion(rawNodePermsRef.current)
        .then((expanded) => {
          setNodePerms(expanded);
          setNodeLoading(false);
        })
        .catch((e: any) => {
          setNodeError(e?.message ?? String(e));
          setNodeLoading(false);
        });
    }
    // Always re-expand parent permissions — inherited items rely solely on this path.
    if (showParentPerms && rawParentPermsRef.current !== null) {
      setParentPerms(null);
      withGroupExpansion(rawParentPermsRef.current)
        .then((expanded) => setParentPerms(expanded))
        .catch((e: any) => setParentPermsError(e?.message ?? String(e)));
    }
  }, [expandGroups]);

  React.useEffect(() => {
    if (!permissionsDenied || !siteUrl) return;
    sp.getSiteOwners(siteUrl.trim()).then(setSiteOwners).catch(() => {});
  }, [permissionsDenied, siteUrl]);

  // ── Show parent permissions ──────────────────────────────────────────────

  const handleShowParentPerms = async (nodeOverride?: FolderFileNode): Promise<void> => {
    const target = nodeOverride ?? selectedNode;
    if (!target) return;
    setParentPermsLoading(true);
    setParentPermsError('');
    try {
      const result = await sp.getParentPermissions(
        siteUrl.trim(),
        target.serverRelativeUrl,
        abortRef.current?.signal,
      );
      if (result) {
        rawParentPermsRef.current = result.users;
        setParentPermsName(result.name);
        setParentPerms(await withGroupExpansion(result.users));
      } else {
        rawParentPermsRef.current = [];
        setParentPerms([]);
        setParentPermsName('');
      }
    } catch (err: any) {
      setParentPermsError(err?.message ?? String(err));
    } finally {
      setParentPermsLoading(false);
    }
  };

  const handleParentPermsCheckbox = (_: unknown, d: { checked: boolean | 'mixed' }): void => {
    const checked = !!d.checked;
    setShowParentPerms(checked);
    if (checked) {
      setParentPerms(null);
      setParentPermsName('');
      setParentPermsError('');
      handleShowParentPerms().catch((e) => console.error('[SmartPermissions] handleShowParentPerms failed:', e));
    } else {
      setParentPerms(null);
      setParentPermsName('');
      setParentPermsError('');
    }
  };

  // Auto-connect on mount
  React.useEffect(() => {
    handleConnect().catch((e) => console.error('[SmartPermissions] handleConnect failed:', e));
  }, []);

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.header}>
        <Button
          appearance="subtle"
          icon={<ArrowLeft24Regular />}
          onClick={onBack}
          disabled={isConnecting}
          aria-label="Back to home"
        >
          Back
        </Button>
        <Title3>Permissions Explorer</Title3>
      </div>

      {connectError && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM }}>
          <MessageBarBody>{connectError}</MessageBarBody>
        </MessageBar>
      )}
      {connectStatus && !connectError && (
        <Body1
          style={{
            color: isConnected
              ? tokens.colorStatusSuccessForeground1
              : tokens.colorNeutralForeground3,
            marginBottom: tokens.spacingVerticalM,
          }}
        >
          {connectStatus}
        </Body1>
      )}

      {isConnected && (
        <>
          {/* Library picker */}
          <div style={{ marginBottom: tokens.spacingVerticalM }}>
            <Field label="Library">
              <Select
                value={selectedLibrary}
                onChange={handleLibraryChange}
                style={{ maxWidth: '400px' }}
              >
                {libraries.map((lib) => (
                  <option key={lib.serverRelativeUrl} value={lib.serverRelativeUrl}>
                    {lib.title}
                  </option>
                ))}
              </Select>
            </Field>
          </div>

          {/* Toggles + icon legend */}
          <div style={{ marginBottom: tokens.spacingVerticalM, display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }}>
            <div style={{ display: 'flex', gap: tokens.spacingHorizontalS, flexWrap: 'wrap' }}>
              <ToggleButton
                size="small"
                checked={showUniqueOnly}
                onClick={() => setShowUniqueOnly((prev) => !prev)}
                icon={<Filter16Regular />}
              >
                Unique permissions only
              </ToggleButton>
              <ToggleButton
                size="small"
                checked={filterExternalOnly}
                onClick={() => setFilterExternalOnly((prev) => !prev)}
                icon={<PersonWarning16Regular />}
              >
                External users only
              </ToggleButton>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalL, flexWrap: 'wrap' }}>
              <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Legend:</Text>
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Badge appearance="filled" color="warning" size="small">Unique</Badge>
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Item has unique permissions</Text>
              </span>
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <ArrowCircleDown16Regular style={{ color: tokens.colorNeutralForeground3 }} />
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Contains unique permissions below</Text>
              </span>
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <ArrowTriangleDown style={{ color: tokens.colorNeutralForeground3 }} />
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>Contains external user access below</Text>
              </span>
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <PersonWarning16Regular style={{ color: tokens.colorPaletteRedForeground1 }} />
                <Text size={100} style={{ color: tokens.colorNeutralForeground3 }}>External user access on this item</Text>
              </span>
            </div>
          </div>

          {permissionsDenied && (
            <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalM }}>
              <MessageBarBody>
                <strong>Viewing with Member access</strong> — permission assignments are not visible.
                Reading who has access requires the <strong>Manage Permissions</strong> right (Site
                Owner or higher). You can still see which items have broken inheritance using the ↓
                indicators. To see locations you can access, use the <strong>User Access</strong> tool.
                {siteOwners.length > 0 && (
                  <> Site Owners: {siteOwners.map((o, i) => (
                    <React.Fragment key={o.email || o.title}>
                      {i > 0 && ', '}
                      {o.email
                        ? <Link href={`mailto:${o.email}`}>{o.title}</Link>
                        : o.title}
                    </React.Fragment>
                  ))}.</>
                )}
              </MessageBarBody>
            </MessageBar>
          )}

          <div className={styles.twoCol}>
            {/* Tree panel */}
            <div className={styles.treePanel}>
              {treeStatus && (
                <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                  {treeStatus}
                </Body1>
              )}
              {rootNodes
                .filter((node) =>
                  (!showUniqueOnly || node.hasUniquePermissions || node.hasUniquePermissionsBelow) &&
                  (!filterExternalOnly || node.hasExternalUsers || node.hasExternalUsersBelow)
                )
                .map((node) => (
                  <TreeNode
                    key={node.serverRelativeUrl}
                    node={node}
                    depth={0}
                    selectedUrl={selectedNode?.serverRelativeUrl ?? ''}
                    onSelect={handleSelectNode}
                    onLoadChildren={loadChildren}
                    showUniqueOnly={showUniqueOnly}
                    filterExternalOnly={filterExternalOnly}
                    externalAccessUrls={externalAccessUrls}
                  />
                ))}
            </div>

            {/* Permissions panel */}
            <div className={styles.permPanel}>
              {!selectedNode && (
                <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Select a folder or file to view its permissions.
                </Body1>
              )}

              {selectedNode && (
                <>
                  {/* Item name + badge */}
                  <div
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: tokens.spacingHorizontalS,
                      marginBottom: tokens.spacingVerticalM,
                      flexWrap: 'wrap',
                    }}
                  >
                    <Text weight="semibold">{selectedNode.name}</Text>
                    {nodeHasUnique ? (
                      <Badge appearance="filled" color="warning">Unique permissions</Badge>
                    ) : !nodeLoading && !nodeError ? (
                      <Badge appearance="filled" color="informative">Inherited</Badge>
                    ) : null}
                  </div>

                  {/* Options — always at top */}
                  <div className={styles.optionsBar}>
                    {!permissionsDenied && (
                      <Checkbox
                        label="Expand group members"
                        checked={expandGroups}
                        onChange={(_, d) => setExpandGroups(!!d.checked)}
                      />
                    )}
                    {!permissionsDenied && !nodeLoading && !nodeError && !nodeHasUnique && (
                      <Checkbox
                        label="Show parent permissions"
                        checked={showParentPerms}
                        onChange={handleParentPermsCheckbox}
                      />
                    )}
                  </div>

                  {nodeLoading && <Spinner size="small" />}

                  {nodeError && (
                    <MessageBar intent="error">
                      <MessageBarBody>{nodeError}</MessageBarBody>
                    </MessageBar>
                  )}

                  {/* Inherited banner */}
                  {!nodeLoading && !nodeError && !nodeHasUnique && (
                    <>
                      <div className={styles.inheritedBanner}>
                        <Link16Regular style={{ flexShrink: 0, color: tokens.colorBrandForeground1 }} />
                        <Body1>This item inherits permissions from its parent.</Body1>
                      </div>
                      {permissionsDenied && myPermLevel && (
                        <MessageBar intent="success" style={{ marginTop: tokens.spacingVerticalS }}>
                          <MessageBarBody>
                            <strong>Your access:</strong> {myPermLevel}
                          </MessageBarBody>
                        </MessageBar>
                      )}
                    </>
                  )}

                  {/* Unique permissions table */}
                  {!nodeLoading && !nodeError && nodeHasUnique && (
                    permissionsDenied && applyPermFilters(nodePerms, excludeLimitedAccess, filterExternalOnly).length === 0
                      ? (
                        <>
                          {myPermLevel && (
                            <MessageBar intent="success" style={{ marginBottom: tokens.spacingVerticalS }}>
                              <MessageBarBody>
                                <strong>Your access:</strong> {myPermLevel}
                              </MessageBarBody>
                            </MessageBar>
                          )}
                          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                            Permission assignments are not visible — the Manage Permissions right
                            (Site Owner or higher) is required to read who has access to this item.
                          </Body1>
                        </>
                      )
                      : applyPermFilters(nodePerms, excludeLimitedAccess, filterExternalOnly).length === 0 && filterExternalOnly
                        ? (
                          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                            No external users have direct access to this item.
                          </Body1>
                        ) : (
                          <PermTable
                            users={applyPermFilters(nodePerms, excludeLimitedAccess, filterExternalOnly)}
                            styles={styles}
                            onCheckAccess={onNavigateToUserAccess}
                          />
                        )
                  )}

                  {/* Parent permissions */}
                  {!nodeLoading && !nodeError && !nodeHasUnique && showParentPerms && (
                    <>
                      {parentPermsLoading && (
                        <Spinner size="small" style={{ marginTop: tokens.spacingVerticalS }} />
                      )}
                      {parentPermsError && (
                        <MessageBar intent="error" style={{ marginTop: tokens.spacingVerticalS }}>
                          <MessageBarBody>{parentPermsError}</MessageBarBody>
                        </MessageBar>
                      )}
                      {!parentPermsLoading && !parentPermsError && parentPerms && (
                        <>
                          <Text
                            size={200}
                            style={{
                              color: tokens.colorNeutralForeground3,
                              marginBottom: tokens.spacingVerticalS,
                              display: 'block',
                            }}
                          >
                            Inherited from: <strong>{parentPermsName}</strong>
                          </Text>
                          {parentPerms.length === 0 ? (
                            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                              {permissionsDenied
                                ? 'Parent permission assignments are not visible — requires the Manage Permissions right (Site Owner or higher).'
                                : 'No permissions found on parent.'}
                            </Body1>
                          ) : (
                            <PermTable
                            users={applyPermFilters(parentPerms, excludeLimitedAccess, filterExternalOnly)}
                            styles={styles}
                            onCheckAccess={onNavigateToUserAccess}
                          />
                          )}
                        </>
                      )}
                    </>
                  )}
                </>
              )}
            </div>
          </div>
        </>
      )}
    </div>
  );
};
