import * as React from 'react';
import {
  Button,
  Field,
  Badge,
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
  ChevronRight16Regular,
  ChevronDown16Regular,
  ArrowCircleDown16Regular,
  Link16Regular,
  PersonWarning16Regular,
  Filter16Regular,
} from '@fluentui/react-icons';

import { SharePointService } from '../services/SharePointService';
import {
  LibraryInfo,
  FolderFileNode,
  UserPermissionInfo,
} from '../models/models';
import { PermTable } from './shared/PermTable';
import { SiteOwnersLinks } from './shared/SiteOwnersLinks';
import { isExternalUser } from './shared/externalUsers';
import { applyPermFilters } from './shared/permFilters';

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


// ── Folder tree node ──────────────────────────────────────────────────────────

interface TreeNodeProps {
  node: FolderFileNode;
  depth: number;
  selectedUrl: string;
  focusedUrl: string;
  expandedUrls: Set<string>;
  onRowClick: (node: FolderFileNode) => void;
  registerRef: (url: string, el: HTMLDivElement | null) => void;
  showUniqueOnly: boolean;
  filterExternalOnly: boolean;
  externalAccessUrls: Set<string>;
  /**
   * Bumped by the parent on every in-place node mutation. Not read directly —
   * its purpose is to be a prop that changes on every tree mutation, so that
   * if TreeNode is ever wrapped in React.memo, a mutation is still guaranteed
   * to trigger a re-render even though `node`'s own reference doesn't change.
   */
  structureVersion: number;
}

// Expansion state lives in the parent (expandedUrls) so the view can compute
// the flat list of visible nodes for keyboard navigation; keyboard events are
// handled on the role="tree" container via a roving tabindex.
const TreeNode: React.FC<TreeNodeProps> = ({
  node,
  depth,
  selectedUrl,
  focusedUrl,
  expandedUrls,
  onRowClick,
  registerRef,
  showUniqueOnly,
  filterExternalOnly,
  externalAccessUrls,
  structureVersion,
}) => {
  const styles = useStyles();
  const expanded = expandedUrls.has(node.serverRelativeUrl);

  const isSelected = node.serverRelativeUrl === selectedUrl;

  return (
    <div role="none">
      <div
        role="treeitem"
        aria-expanded={node.isFolder && node.hasChildren ? expanded : undefined}
        aria-selected={isSelected}
        aria-level={depth + 1}
        tabIndex={node.serverRelativeUrl === focusedUrl ? 0 : -1}
        ref={(el) => registerRef(node.serverRelativeUrl, el)}
        className={`${styles.treeNode} ${isSelected ? styles.treeNodeSelected : ''}`}
        style={{ paddingLeft: `${depth * 16 + 4}px` }}
        onClick={(e) => { e.stopPropagation(); onRowClick(node); }}
        data-structure-version={structureVersion}
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

      {expanded && node.loadError && (
        <div
          role="none"
          style={{
            paddingLeft: `${(depth + 1) * 16 + 24}px`,
            color: tokens.colorPaletteRedForeground1,
            fontSize: tokens.fontSizeBase200,
          }}
        >
          {node.loadError}
        </div>
      )}

      {expanded && node.children.length > 0 && (
        <div role="group">
          {node.children
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
                focusedUrl={focusedUrl}
                expandedUrls={expandedUrls}
                onRowClick={onRowClick}
                registerRef={registerRef}
                showUniqueOnly={showUniqueOnly}
                filterExternalOnly={filterExternalOnly}
                externalAccessUrls={externalAccessUrls}
                structureVersion={structureVersion}
              />
            ))}
        </div>
      )}
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
  // Tree expansion + keyboard focus (roving tabindex) state
  const [expandedUrls, setExpandedUrls] = React.useState<Set<string>>(new Set());
  const [focusedUrl, setFocusedUrl] = React.useState('');
  const nodeRefs = React.useRef<Map<string, HTMLDivElement>>(new Map());
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
  // Bumped at the start of every loadLibrary call. Background work (probes,
  // prefetches, external-user scans) captures the generation active when it
  // started and checks it before applying state — so a library switch that
  // happens mid-flight can't have a stale result bleed into the new view.
  const loadGenerationRef = React.useRef(0);
  // Bumped alongside every tree mutation so the render can be forced even if
  // TreeNode is ever wrapped in React.memo in the future — plain node
  // mutation + array-copy (below) works today only because TreeNode isn't
  // memoized; a memoized TreeNode wouldn't see a changed `node` prop
  // reference on mutation-in-place.
  const [structureVersion, setStructureVersion] = React.useState(0);
  const touchTree = React.useCallback((): void => {
    setRootNodes((prev) => [...prev]);
    setStructureVersion((v) => v + 1);
  }, []);

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
    const myGeneration = ++loadGenerationRef.current;
    const isStale = (): boolean => loadGenerationRef.current !== myGeneration;

    setRootNodes([]);
    setTreeStatus('');
    setSelectedNode(null);
    setNodePerms([]);
    setExternalAccessUrls(new Set());
    setExpandedUrls(new Set());
    setFocusedUrl('');
    nodeRefs.current.clear();
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
      if (isStale()) return; // user switched libraries again while this fetch was in flight
      setRootNodes(nodes);
      setFocusedUrl(nodes[0]?.serverRelativeUrl ?? '');
      if (nodes.length === 0) setTreeStatus('This library is empty.');

      // Proactively check whether role-assignment reads are permitted for this library.
      sp.getItemPermissions(
        siteUrl.trim(),
        { name: lib.title, serverRelativeUrl: lib.serverRelativeUrl, isFolder: true, hasChildren: false, children: [] },
        abortRef.current?.signal,
      ).then((probe) => {
        // A slow probe resolving after the user has already switched to another
        // library must not flag the (unrelated) library now on screen as denied.
        if (isStale()) return;
        if (probe.permissionDenied) setPermissionsDenied(true);
      }).catch(() => { /* ignore probe errors */ });

      // Background scan for external users on root-level nodes.
      scanExternalUsers(nodes, myGeneration);

      // Pre-fetch one level deep for each root folder to detect hasUniquePermissionsBelow
      // at load time rather than requiring the user to expand each folder first.
      // Results are stored in folderCacheRef so expanding a folder skips the API call.
      // Bounded by scanConcurrency — previously fired one request per root folder
      // simultaneously with no cap, risking throttling on wide libraries.
      const signal = abortRef.current?.signal;
      const prefetchTasks = nodes.filter((n) => n.isFolder).map((folder) => async (): Promise<undefined> => {
        try {
          const children = await sp.getFolderContents(siteUrl.trim(), folder.serverRelativeUrl, signal);
          if (isStale()) return undefined;
          // Set parent now so propagation works correctly when the external-user scan fires.
          children.forEach((c) => { c.parent = folder; });
          folderCacheRef.current.set(folder.serverRelativeUrl, children);
          if (children.some((c) => c.hasUniquePermissions)) {
            folder.hasUniquePermissionsBelow = true;
            touchTree();
          }
          // Scan pre-fetched children so icons appear at library-load time,
          // not only after the user expands the folder.
          scanExternalUsers(children, myGeneration);
        } catch { /* ignore prefetch errors */ }
        return undefined;
      });
      sp.runConcurrent(prefetchTasks, sp.scanConcurrency).catch(() => { /* ignore — background prefetch */ });
    } catch (err: any) {
      if (isStale()) return;
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
    // Snapshot (don't bump) — this call rides on whichever library load is
    // current; if the user switches libraries while it's in flight, the
    // generation changes and the guards below discard the stale result.
    const myGeneration = loadGenerationRef.current;
    node.isLoading = true;
    node.loadError = undefined;
    touchTree();

    try {
      const cached = folderCacheRef.current.get(node.serverRelativeUrl);
      const children = cached ?? await sp.getFolderContents(
        siteUrl.trim(),
        node.serverRelativeUrl,
        abortRef.current?.signal,
      );
      if (loadGenerationRef.current !== myGeneration) return;
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
      scanExternalUsers(children, myGeneration);

      // Pre-fetch one level deeper for each folder child so their icons are ready
      // before the user expands them — same pattern used in loadLibrary for depth-0.
      // Bounded by scanConcurrency — previously fired one request per child folder
      // simultaneously with no cap, and each expansion multiplied it further.
      const prefetchSignal = abortRef.current?.signal;
      const prefetchTasks = children
        .filter((c) => c.isFolder && c.hasChildren && !folderCacheRef.current.has(c.serverRelativeUrl))
        .map((folder) => async (): Promise<undefined> => {
          try {
            const grandchildren = await sp.getFolderContents(siteUrl.trim(), folder.serverRelativeUrl, prefetchSignal);
            if (loadGenerationRef.current !== myGeneration) return undefined;
            grandchildren.forEach((gc) => { gc.parent = folder; });
            folderCacheRef.current.set(folder.serverRelativeUrl, grandchildren);
            if (grandchildren.some((gc) => gc.hasUniquePermissions)) {
              propagateUniqueBelow(folder);
              touchTree();
            }
            scanExternalUsers(grandchildren, myGeneration);
          } catch { /* ignore */ }
          return undefined;
        });
      sp.runConcurrent(prefetchTasks, sp.scanConcurrency).catch(() => { /* ignore — background prefetch */ });
    } catch {
      // Rendered as inline, non-navigable text (see FolderFileNode.loadError) —
      // a synthetic child node here previously collided on an empty-string key
      // and was reachable via arrow-key tree navigation like a real item.
      node.loadError = 'Error loading contents';
      node.children = [];
    } finally {
      node.isLoading = false;
      touchTree();
    }
  }, [siteUrl, touchTree]);

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

  // ── Tree expansion + keyboard navigation ─────────────────────────────────

  const registerNodeRef = React.useCallback((url: string, el: HTMLDivElement | null): void => {
    if (el) { nodeRefs.current.set(url, el); } else { nodeRefs.current.delete(url); }
  }, []);

  const nodeVisible = (n: FolderFileNode): boolean =>
    (!showUniqueOnly || !!n.hasUniquePermissions || !!n.hasUniquePermissionsBelow) &&
    (!filterExternalOnly || !!n.hasExternalUsers || !!n.hasExternalUsersBelow);

  // Flat list of currently rendered nodes, mirroring the render filters —
  // this is the keyboard navigation order. Memoized: this previously walked
  // the entire (possibly large) tree on every Arrow/Home/End keypress.
  const flattenVisible = React.useMemo((): FolderFileNode[] => {
    const out: FolderFileNode[] = [];
    const visit = (nodes: FolderFileNode[]): void => {
      for (const n of nodes.filter(nodeVisible)) {
        out.push(n);
        if (n.isFolder && expandedUrls.has(n.serverRelativeUrl)) visit(n.children);
      }
    };
    visit(rootNodes);
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rootNodes, structureVersion, expandedUrls, showUniqueOnly, filterExternalOnly]);

  const toggleExpand = (node: FolderFileNode): void => {
    if (!node.isFolder || !node.hasChildren) return;
    const isOpen = expandedUrls.has(node.serverRelativeUrl);
    if (!isOpen && node.children.length === 0) {
      loadChildren(node).catch((e) => console.error('[SmartPermissions] loadChildren failed:', e));
    }
    setExpandedUrls((prev) => {
      const next = new Set(prev);
      if (next.has(node.serverRelativeUrl)) { next.delete(node.serverRelativeUrl); } else { next.add(node.serverRelativeUrl); }
      return next;
    });
  };

  const handleRowClick = (node: FolderFileNode): void => {
    setFocusedUrl(node.serverRelativeUrl);
    if (node.isFolder && node.hasChildren) toggleExpand(node);
    handleSelectNode(node).catch((e) => console.error('[SmartPermissions] handleSelectNode failed:', e));
  };

  const focusNode = (node: FolderFileNode | undefined): void => {
    if (!node) return;
    setFocusedUrl(node.serverRelativeUrl);
    nodeRefs.current.get(node.serverRelativeUrl)?.focus();
  };

  const handleTreeKeyDown = (e: React.KeyboardEvent): void => {
    const visible = flattenVisible;
    if (visible.length === 0) return;
    const idx = visible.findIndex((n) => n.serverRelativeUrl === focusedUrl);
    const current = idx >= 0 ? visible[idx] : visible[0];

    switch (e.key) {
      case 'ArrowDown':
        focusNode(visible[Math.min(idx < 0 ? 0 : idx + 1, visible.length - 1)]);
        break;
      case 'ArrowUp':
        focusNode(visible[Math.max(idx - 1, 0)]);
        break;
      case 'ArrowRight':
        if (current.isFolder && current.hasChildren && !expandedUrls.has(current.serverRelativeUrl)) {
          toggleExpand(current);
        } else if (current.isFolder && expandedUrls.has(current.serverRelativeUrl) && idx + 1 < visible.length) {
          focusNode(visible[idx + 1]);
        }
        break;
      case 'ArrowLeft':
        if (current.isFolder && expandedUrls.has(current.serverRelativeUrl)) {
          toggleExpand(current);
        } else if (current.parent) {
          focusNode(current.parent);
        }
        break;
      case 'Home':
        focusNode(visible[0]);
        break;
      case 'End':
        focusNode(visible[visible.length - 1]);
        break;
      case 'Enter':
      case ' ':
        handleRowClick(current);
        break;
      default:
        return; // don't preventDefault on unhandled keys
    }
    e.preventDefault();
  };

  // Background scan: checks only unique-permission nodes (inherited ones are skipped — no API
  // call needed). Uses one direct RoleAssignments fetch per node instead of two calls.
  // `generation` is the loadLibrary/loadChildren generation active when this scan was
  // started — if the user has since switched libraries, results are discarded instead
  // of being applied to whatever is now on screen.
  const scanExternalUsers = (nodes: FolderFileNode[], generation: number): void => {
    const uniqueNodes = nodes.filter((n) => n.hasUniquePermissions);
    if (!uniqueNodes.length) return;
    const tasks = uniqueNodes.map((node) => async (): Promise<undefined> => {
      if (abortRef.current?.signal.aborted || loadGenerationRef.current !== generation) return undefined;
      const hasExt = await sp.scanNodeForExternalUsers(siteUrl.trim(), node, abortRef.current?.signal);
      if (loadGenerationRef.current !== generation) return undefined;
      if (hasExt === 'denied') {
        setPermissionsDenied(true);
      } else if (hasExt) {
        node.hasExternalUsers = true;
        if (node.parent) propagateExternalBelow(node.parent);
        propagateExternalDown(node);  // mark already-loaded descendants that inherit
        setExternalAccessUrls((prev) => { const next = new Set(Array.from(prev)); next.add(node.serverRelativeUrl); return next; });
        touchTree();
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
              u.groupId,
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

  // Auto-connect on mount. Cleanup aborts in-flight work on unmount (e.g. the
  // user navigates away from Explorer) — without this, background scans kept
  // running and calling setState on an unmounted component.
  React.useEffect(() => {
    handleConnect().catch((e) => console.error('[SmartPermissions] handleConnect failed:', e));
    return () => {
      loadGenerationRef.current++; // invalidate any in-flight generation-guarded callbacks
      abortRef.current?.abort();
    };
  }, []);

  // Computed once per render instead of re-filtering nodePerms up to 3× in the
  // permission panel below (each call re-filters the same, potentially large,
  // expanded-group list).
  const filteredNodePerms = React.useMemo(
    () => applyPermFilters(nodePerms, excludeLimitedAccess, filterExternalOnly),
    [nodePerms, excludeLimitedAccess, filterExternalOnly],
  );

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
                    {lib.noCrawl ? `${lib.title} (hidden from search)` : lib.title}
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
                <SiteOwnersLinks owners={siteOwners} />
              </MessageBarBody>
            </MessageBar>
          )}

          <div className={styles.twoCol}>
            {/* Tree panel */}
            <div
              className={styles.treePanel}
              role="tree"
              aria-label="Folders and files"
              onKeyDown={handleTreeKeyDown}
            >
              {treeStatus && (
                <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                  {treeStatus}
                </Body1>
              )}
              {rootNodes
                .filter(nodeVisible)
                .map((node) => (
                  <TreeNode
                    key={node.serverRelativeUrl}
                    node={node}
                    depth={0}
                    selectedUrl={selectedNode?.serverRelativeUrl ?? ''}
                    focusedUrl={focusedUrl}
                    expandedUrls={expandedUrls}
                    onRowClick={handleRowClick}
                    registerRef={registerNodeRef}
                    showUniqueOnly={showUniqueOnly}
                    filterExternalOnly={filterExternalOnly}
                    externalAccessUrls={externalAccessUrls}
                    structureVersion={structureVersion}
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
                    permissionsDenied && filteredNodePerms.length === 0
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
                      : filteredNodePerms.length === 0 && filterExternalOnly
                        ? (
                          <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                            No external users have direct access to this item.
                          </Body1>
                        ) : (
                          <PermTable
                            users={filteredNodePerms}
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
