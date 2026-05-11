import * as React from 'react';
import {
  Button,
  Input,
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
} from '@fluentui/react-icons';

import { WebPartContext } from '@microsoft/sp-webpart-base';
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
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  connectRow: {
    display: 'flex',
    gap: tokens.spacingHorizontalM,
    alignItems: 'flex-end',
    flexWrap: 'wrap',
    marginBottom: tokens.spacingVerticalM,
  },
  urlField: {
    flexGrow: 1,
    minWidth: '300px',
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
  },
  permTd: {
    padding: '5px 8px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    verticalAlign: 'top',
  },
});

// ── Role badge colour ─────────────────────────────────────────────────────────

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

// ── Folder tree node ──────────────────────────────────────────────────────────

interface TreeNodeProps {
  node: FolderFileNode;
  depth: number;
  selectedUrl: string;
  onSelect: (node: FolderFileNode) => void;
  onLoadChildren: (node: FolderFileNode) => void;
}

const TreeNode: React.FC<TreeNodeProps> = ({
  node,
  depth,
  selectedUrl,
  onSelect,
  onLoadChildren,
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
          <Folder24Regular style={{ fontSize: '16px', color: tokens.colorPaletteYellowForeground1, flexShrink: 0 }} />
        ) : (
          <Document24Regular style={{ fontSize: '16px', color: tokens.colorNeutralForeground3, flexShrink: 0 }} />
        )}

        <Text style={{ flexGrow: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
          {node.name}
        </Text>

        {node.hasUniquePermissions && (
          <Badge appearance="filled" color="warning" size="small" style={{ flexShrink: 0 }}>
            Unique
          </Badge>
        )}
        {node.isLoading && (
          <Spinner size="extra-tiny" style={{ flexShrink: 0 }} />
        )}
      </div>

      {expanded && node.children.map((child) => (
        <TreeNode
          key={child.serverRelativeUrl}
          node={child}
          depth={depth + 1}
          selectedUrl={selectedUrl}
          onSelect={onSelect}
          onLoadChildren={onLoadChildren}
        />
      ))}
    </div>
  );
};

// ── Main view ─────────────────────────────────────────────────────────────────

export interface PermissionsExplorerViewProps {
  context: WebPartContext;
  sp: SharePointService;
  onBack: () => void;
}

export const PermissionsExplorerView: React.FC<PermissionsExplorerViewProps> = ({ context, sp, onBack }) => {
  const styles = useStyles();

  // ── Connection ──
  const [siteUrl, setSiteUrl] = React.useState(context.pageContext.web.absoluteUrl);
  const [isConnecting, setIsConnecting] = React.useState(false);
  const [connectStatus, setConnectStatus] = React.useState('');
  const [connectError, setConnectError] = React.useState('');
  const [isConnected, setIsConnected] = React.useState(false);

  // ── Libraries ──
  const [libraries, setLibraries] = React.useState<LibraryInfo[]>([]);

  // ── Browse tab ──
  const [selectedLibrary, setSelectedLibrary] = React.useState('');
  const [rootNodes, setRootNodes] = React.useState<FolderFileNode[]>([]);
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
      const libs = await sp.getLibraries(siteUrl.trim(), abortRef.current.signal);

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

    try {
      const nodes = await sp.getFolderContents(
        siteUrl.trim(),
        lib.serverRelativeUrl,
        abortRef.current?.signal,
      );
      setRootNodes(nodes);
      if (nodes.length === 0) setTreeStatus('This library is empty.');
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

  const loadChildren = async (node: FolderFileNode): Promise<void> => {
    node.isLoading = true;
    setRootNodes((prev) => [...prev]);

    try {
      const children = await sp.getFolderContents(
        siteUrl.trim(),
        node.serverRelativeUrl,
        abortRef.current?.signal,
      );
      node.children = children.map((c) => {
        c.parent = node;
        return c;
      });

      if (children.some((c) => c.hasUniquePermissions)) {
        propagateUniqueBelow(node);
      }
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
  };

  const propagateUniqueBelow = (node: FolderFileNode): void => {
    node.hasUniquePermissionsBelow = true;
    if (node.parent) propagateUniqueBelow(node.parent);
  };

  // ── Select node ──────────────────────────────────────────────────────────

  const handleSelectNode = async (node: FolderFileNode): Promise<void> => {
    setSelectedNode(node);
    setNodeLoading(true);
    setNodePerms([]);
    setNodeHasUnique(false);
    setNodeError('');
    setShowParentPerms(false);
    setParentPerms(null);
    setParentPermsName('');
    setParentPermsError('');

    try {
      const { hasUnique, users } = await sp.getItemPermissions(
        siteUrl.trim(),
        node,
        abortRef.current?.signal,
      );
      setNodeHasUnique(hasUnique);

      const expanded: UserPermissionInfo[] = [];
      for (const u of users) {
        expanded.push(u);
        if (hasUnique && expandGroups && u.principalType === 'SharePointGroup') {
          const members = await sp.getGroupMembers(
            siteUrl.trim(),
            u.displayName,
            u.principalType,
            abortRef.current?.signal,
          );
          members.forEach((m) => {
            m.roles = [...u.roles];
            expanded.push(m);
          });
        }
      }
      setNodePerms(expanded);
    } catch (err: any) {
      setNodeError(err?.message ?? String(err));
    } finally {
      setNodeLoading(false);
    }
  };

  React.useEffect(() => {
    if (selectedNode) {
      handleSelectNode(selectedNode).catch((e) => console.error('[SmartPermissions] handleSelectNode failed:', e));
    }
  }, [expandGroups]);

  // ── Show parent permissions ──────────────────────────────────────────────

  const handleShowParentPerms = async (): Promise<void> => {
    if (!selectedNode) return;
    setShowParentPerms(true);
    setParentPermsLoading(true);
    setParentPermsError('');
    try {
      const result = await sp.getParentPermissions(
        siteUrl.trim(),
        selectedNode.serverRelativeUrl,
        abortRef.current?.signal,
      );
      if (result) {
        setParentPermsName(result.name);
        setParentPerms(result.users);
      } else {
        setParentPerms([]);
        setParentPermsName('');
      }
    } catch (err: any) {
      setParentPermsError(err?.message ?? String(err));
    } finally {
      setParentPermsLoading(false);
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
        >
          Back
        </Button>
        <Title3>Permissions Explorer</Title3>
      </div>

      {/* Connect row */}
      <div className={styles.connectRow}>
        <Field label="Site URL" className={styles.urlField}>
          <Input
            value={siteUrl}
            onChange={(_, d) => setSiteUrl(d.value)}
            placeholder="https://contoso.sharepoint.com/sites/mysite"
            disabled={isConnecting}
          />
        </Field>
        <Button
          appearance="primary"
          onClick={handleConnect}
          disabled={!siteUrl.trim() || isConnecting}
        >
          {isConnecting ? <><Spinner size="tiny" /> Connecting…</> : 'Connect'}
        </Button>
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

          <div className={styles.twoCol}>
            {/* Tree panel */}
            <div className={styles.treePanel}>
              {treeStatus && (
                <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                  {treeStatus}
                </Body1>
              )}
              {rootNodes.map((node) => (
                <TreeNode
                  key={node.serverRelativeUrl}
                  node={node}
                  depth={0}
                  selectedUrl={selectedNode?.serverRelativeUrl ?? ''}
                  onSelect={handleSelectNode}
                  onLoadChildren={loadChildren}
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
                      <Badge appearance="filled" color="warning">
                        Unique permissions
                      </Badge>
                    ) : !nodeLoading && !nodeError ? (
                      <Badge appearance="outline">Inherited permissions</Badge>
                    ) : null}
                  </div>

                  <Checkbox
                    label="Expand SharePoint group members"
                    checked={expandGroups}
                    onChange={(_, d) => setExpandGroups(!!d.checked)}
                    style={{ marginBottom: tokens.spacingVerticalS }}
                  />

                  {nodeLoading && <Spinner size="small" />}

                  {nodeError && (
                    <MessageBar intent="error">
                      <MessageBarBody>{nodeError}</MessageBarBody>
                    </MessageBar>
                  )}

                  {!nodeLoading && !nodeError && nodeHasUnique && (
                    <table className={styles.permTable}>
                      <thead>
                        <tr>
                          <th className={styles.permTh}>User / Group</th>
                          <th className={styles.permTh}>Type</th>
                          <th className={styles.permTh}>Permission Level</th>
                        </tr>
                      </thead>
                      <tbody>
                        {nodePerms.map((u, i) => (
                          <tr key={i}>
                            <td className={styles.permTd}>
                              {u.isGroupMember ? (
                                <span style={{ paddingLeft: '16px', color: tokens.colorNeutralForeground3 }}>
                                  ↳ {u.displayName}
                                </span>
                              ) : (
                                <span
                                  style={{
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '4px',
                                  }}
                                >
                                  {u.principalType === 'User' ? (
                                    <Person24Regular style={{ fontSize: '14px' }} />
                                  ) : (
                                    <People24Regular style={{ fontSize: '14px' }} />
                                  )}
                                  {u.displayName || u.loginName}
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
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}

                  {!nodeLoading && !nodeError && !nodeHasUnique && (
                    <>
                      <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                        This item inherits permissions from its parent.
                      </Body1>
                      {!showParentPerms && (
                        <Button
                          appearance="transparent"
                          size="small"
                          style={{ paddingLeft: 0, marginTop: tokens.spacingVerticalXS }}
                          onClick={handleShowParentPerms}
                        >
                          Show parent permissions
                        </Button>
                      )}
                      {showParentPerms && parentPermsLoading && (
                        <Spinner size="small" style={{ marginTop: tokens.spacingVerticalS }} />
                      )}
                      {showParentPerms && parentPermsError && (
                        <MessageBar intent="error" style={{ marginTop: tokens.spacingVerticalS }}>
                          <MessageBarBody>{parentPermsError}</MessageBarBody>
                        </MessageBar>
                      )}
                      {showParentPerms && !parentPermsLoading && !parentPermsError && parentPerms && (
                        <>
                          <Text
                            size={200}
                            style={{
                              color: tokens.colorNeutralForeground3,
                              marginTop: tokens.spacingVerticalS,
                              display: 'block',
                            }}
                          >
                            Inherited from: <strong>{parentPermsName}</strong>
                          </Text>
                          {parentPerms.length === 0 ? (
                            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>
                              No permissions found on parent.
                            </Body1>
                          ) : (
                            <table className={styles.permTable} style={{ marginTop: tokens.spacingVerticalS }}>
                              <thead>
                                <tr>
                                  <th className={styles.permTh}>User / Group</th>
                                  <th className={styles.permTh}>Type</th>
                                  <th className={styles.permTh}>Permission Level</th>
                                </tr>
                              </thead>
                              <tbody>
                                {parentPerms.map((u, i) => (
                                  <tr key={i}>
                                    <td className={styles.permTd}>
                                      <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                                        {u.principalType === 'User' ? (
                                          <Person24Regular style={{ fontSize: '14px' }} />
                                        ) : (
                                          <People24Regular style={{ fontSize: '14px' }} />
                                        )}
                                        {u.displayName || u.loginName}
                                      </span>
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
                                  </tr>
                                ))}
                              </tbody>
                            </table>
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
