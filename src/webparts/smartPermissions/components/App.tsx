import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  createDOMRenderer,
  RendererProvider,
  Button,
  Input,
  Text,
  tokens,
  Theme,
} from '@fluentui/react-components';
import { Settings24Regular, ShieldLock24Regular } from '@fluentui/react-icons';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { HomeView } from './HomeView';
import { PermissionsReportView } from './PermissionsReportView';
import { PermissionsExplorerView } from './PermissionsExplorerView';
import { UserAccessView } from './UserAccessView';
import { SharingLinksView } from './SharingLinksView';
import { PermissionGroupsView } from './PermissionGroupsView';
import { ExternalUsersView } from './ExternalUsersView';
import { BrokenInheritanceView } from './BrokenInheritanceView';
import { AnonymousLinksView } from './AnonymousLinksView';
import { SettingsView } from './SettingsView';

export type AppView = 'home' | 'report' | 'explorer' | 'userAccess' | 'sharingLinks' | 'groups' | 'externalUsers' | 'brokenInheritance' | 'anonymousLinks' | 'settings';

const LS_SITE_URL   = 'sp-smart-perms-siteUrl';
const LS_CONCURRENCY = 'sp-smart-perms-concurrency';
const LS_GROUP_CAP   = 'sp-smart-perms-groupCap';
const LS_HIDDEN      = 'sp-smart-perms-includeHidden';

export interface IBrandColors {
  primary: string;
  darkAlt: string;
  dark: string;
  darker: string;
  light: string;
  lighter: string;
}

function buildTheme(b: IBrandColors): Theme {
  return {
    ...webLightTheme,
    // Background tokens (primary buttons, selected states)
    colorBrandBackground:                    b.primary,
    colorBrandBackgroundHover:               b.darkAlt,
    colorBrandBackgroundPressed:             b.dark,
    colorBrandBackgroundSelected:            b.darkAlt,
    colorBrandBackgroundStatic:              b.primary,
    colorBrandBackground2:                   b.lighter,
    colorBrandBackground2Hover:              b.light,
    colorBrandBackground2Pressed:            b.light,
    colorBrandBackground3Static:             b.dark,
    colorBrandBackground4Static:             b.darker,
    // Compound brand (checkboxes, radio buttons, sliders)
    colorCompoundBrandBackground:            b.primary,
    colorCompoundBrandBackgroundHover:       b.darkAlt,
    colorCompoundBrandBackgroundPressed:     b.dark,
    // Foreground tokens (icons, text, checkmarks)
    colorBrandForeground1:                   b.primary,
    colorBrandForeground2:                   b.darkAlt,
    colorBrandForeground2Hover:              b.dark,
    colorBrandForeground2Pressed:            b.darker,
    colorCompoundBrandForeground1:           b.primary,
    colorCompoundBrandForeground1Hover:      b.darkAlt,
    colorCompoundBrandForeground1Pressed:    b.dark,
    // Link foreground tokens
    colorBrandForegroundLink:                b.primary,
    colorBrandForegroundLinkHover:           b.darkAlt,
    colorBrandForegroundLinkPressed:         b.dark,
    colorBrandForegroundLinkSelected:        b.primary,
    // Stroke tokens (focus rings, borders)
    colorBrandStroke1:                       b.primary,
    colorBrandStroke2:                       b.light,
    colorBrandStroke2Hover:                  b.primary,
    colorBrandStroke2Pressed:                b.darkAlt,
    colorCompoundBrandStroke:                b.primary,
    colorCompoundBrandStrokeHover:           b.darkAlt,
    colorCompoundBrandStrokePressed:         b.dark,
  };
}

export interface AppProps {
  context: WebPartContext;
  sp: SharePointService;
  excel: ExcelExportService;
  defaultView?: AppView;
  brandColors: IBrandColors;
}

class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { error: Error | null }
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error: Error): { error: Error } {
    return { error };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error('[SmartPermissions] Render error:', error, info.componentStack);
  }

  render(): React.ReactNode {
    const { error } = this.state;
    if (error) {
      return (
        <div style={{
          padding: '16px', fontFamily: 'Consolas, monospace', fontSize: '13px',
          background: '#fff3f3', border: '1px solid #c00', borderRadius: '4px', margin: '8px',
        }}>
          <strong style={{ color: '#c00', fontSize: '14px' }}>Smart Permissions — Render Error</strong>
          <br /><br />
          <strong>Message:</strong> {error.message || String(error)}
          <br /><br />
          <strong>Stack:</strong>
          <pre style={{
            fontSize: '11px', whiteSpace: 'pre-wrap', overflowWrap: 'anywhere',
            background: '#f5f5f5', padding: '8px', margin: '4px 0', borderRadius: '2px',
          }}>
            {error.stack ?? '(no stack available)'}
          </pre>
        </div>
      );
    }
    return this.props.children;
  }
}

let renderer: ReturnType<typeof createDOMRenderer>;
try {
  renderer = createDOMRenderer(document);
} catch (e: any) {
  console.error('[SmartPermissions] createDOMRenderer failed:', e);
  throw e;
}

export const App: React.FC<AppProps> = ({ context, sp, excel, defaultView, brandColors }) => {
  const theme = React.useMemo(() => buildTheme(brandColors), [brandColors.primary]);

  const [view, setView] = React.useState<AppView>(defaultView ?? 'home');
  const [prevView, setPrevView] = React.useState<AppView>('home');

  // Restore last-used site URL from localStorage, fall back to current web
  const defaultUrl = context.pageContext.web.absoluteUrl;
  const [siteUrl, setSiteUrl] = React.useState(
    () => localStorage.getItem(LS_SITE_URL) ?? defaultUrl,
  );
  const [editUrl, setEditUrl] = React.useState(siteUrl);
  const [isEditing, setIsEditing] = React.useState(false);

  // Persist settings to localStorage and keep service in sync
  const [includeHidden, setIncludeHidden] = React.useState(
    () => localStorage.getItem(LS_HIDDEN) === 'true',
  );
  const [scanConcurrency, setScanConcurrency] = React.useState(
    () => parseInt(localStorage.getItem(LS_CONCURRENCY) ?? '4', 10),
  );
  const [groupMemberCap, setGroupMemberCap] = React.useState(
    () => parseInt(localStorage.getItem(LS_GROUP_CAP) ?? '500', 10),
  );

  React.useEffect(() => { localStorage.setItem(LS_SITE_URL, siteUrl); }, [siteUrl]);
  React.useEffect(() => { localStorage.setItem(LS_HIDDEN, String(includeHidden)); }, [includeHidden]);
  React.useEffect(() => {
    localStorage.setItem(LS_CONCURRENCY, String(scanConcurrency));
    sp.scanConcurrency = scanConcurrency;
  }, [scanConcurrency]);
  React.useEffect(() => {
    localStorage.setItem(LS_GROUP_CAP, String(groupMemberCap));
    sp.groupMemberCap = groupMemberCap;
  }, [groupMemberCap]);

  // Pre-fill login for the User Access view when launched from Explorer
  const [userAccessPrefill, setUserAccessPrefill] = React.useState<string | undefined>();

  // Tracks whether any view has a scan in progress (set via callback)
  const [scanBusy, setScanBusy] = React.useState(false);

  const handleConnect = (): void => {
    if (editUrl.trim()) {
      setSiteUrl(editUrl.trim());
    }
    setIsEditing(false);
  };

  const handleStartEdit = (): void => {
    setEditUrl(siteUrl);
    setIsEditing(true);
  };

  const handleCancelEdit = (): void => {
    setEditUrl(siteUrl);
    setIsEditing(false);
  };

  const handleOpenSettings = (): void => {
    setPrevView(view === 'settings' ? prevView : view);
    setView('settings');
  };

  const handleNavigateToUserAccess = (loginName: string): void => {
    setUserAccessPrefill(loginName);
    setView('userAccess');
  };

  return (
    <ErrorBoundary>
    <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={theme} style={{ minHeight: '400px', position: 'relative' }}>

      {/* Banner — shown on inner views (not home or settings) */}
      {view !== 'home' && view !== 'settings' && (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'auto 1fr auto',
            alignItems: 'center',
            paddingTop: tokens.spacingVerticalS,
            paddingBottom: tokens.spacingVerticalS,
            paddingLeft: tokens.spacingHorizontalM,
            paddingRight: tokens.spacingHorizontalS,
            background: brandColors.primary,
            gap: tokens.spacingHorizontalM,
          }}
        >
          {/* Left: branding */}
          <div style={{ display: 'flex', alignItems: 'center', gap: tokens.spacingHorizontalS, flexShrink: 0 }}>
            <ShieldLock24Regular style={{ color: 'white', fontSize: '20px' }} />
            <Text style={{ color: 'white', fontWeight: tokens.fontWeightSemibold, whiteSpace: 'nowrap' }}>
              SharePoint Smart Permissions
            </Text>
          </div>

          {/* Center: URL or edit input */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: tokens.spacingHorizontalS, minWidth: 0, overflow: 'hidden' }}>
            {isEditing ? (
              <>
                <Input
                  value={editUrl}
                  onChange={(_, d) => setEditUrl(d.value)}
                  placeholder="https://contoso.sharepoint.com/sites/mysite"
                  style={{ minWidth: '200px', maxWidth: '400px', flexGrow: 1 }}
                  onKeyDown={(e) => { if (e.key === 'Enter') handleConnect(); }}
                />
                <Button appearance="secondary" onClick={handleConnect} disabled={!editUrl.trim()}>
                  Connect
                </Button>
                <Button appearance="transparent" style={{ color: 'white', flexShrink: 0 }} onClick={handleCancelEdit}>
                  Cancel
                </Button>
              </>
            ) : (
              <>
                <span
                  style={{
                    width: '8px',
                    height: '8px',
                    borderRadius: '50%',
                    background: 'rgba(255,255,255,0.75)',
                    flexShrink: 0,
                    display: 'inline-block',
                  }}
                />
                <Text
                  style={{
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                    color: 'white',
                  }}
                >
                  {siteUrl}
                </Text>
                <Button
                  appearance="transparent"
                  size="small"
                  style={{ color: 'white', flexShrink: 0 }}
                  onClick={handleStartEdit}
                >
                  Change URL
                </Button>
              </>
            )}
          </div>

          {/* Right: settings */}
          <Button
            appearance="transparent"
            icon={<Settings24Regular style={{ color: 'white' }} />}
            aria-label="Settings"
            title="Settings"
            onClick={handleOpenSettings}
          />
        </div>
      )}

      {/* Gear icon on home view */}
      {view === 'home' && (
        <div style={{ position: 'absolute', top: '4px', right: '8px', zIndex: 10 }}>
          <Button
            appearance="transparent"
            icon={<Settings24Regular style={{ color: 'white' }} />}
            aria-label="Settings"
            title="Settings"
            onClick={handleOpenSettings}
          />
        </div>
      )}

      {view === 'home' && (
        <HomeView onNavigate={setView} primaryColor={brandColors.primary} />
      )}
      {view === 'report' && (
        <PermissionsReportView
          key={siteUrl}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          onBack={() => setView('home')}
        />
      )}
      {view === 'explorer' && (
        <PermissionsExplorerView
          key={siteUrl + String(includeHidden)}
          sp={sp}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          onBack={() => setView('home')}
          onNavigateToUserAccess={handleNavigateToUserAccess}
        />
      )}
      {view === 'userAccess' && (
        <UserAccessView
          key={siteUrl + String(includeHidden)}
          sp={sp}
          excel={excel}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          prefillLogin={userAccessPrefill}
          onPrefillUsed={() => setUserAccessPrefill(undefined)}
          onBack={() => setView('home')}
        />
      )}
      {view === 'sharingLinks' && (
        <SharingLinksView
          key={siteUrl}
          sp={sp}
          siteUrl={siteUrl}
          onBack={() => setView('home')}
        />
      )}
      {view === 'groups' && (
        <PermissionGroupsView
          key={siteUrl}
          sp={sp}
          siteUrl={siteUrl}
          onBack={() => setView('home')}
        />
      )}
      {view === 'externalUsers' && (
        <ExternalUsersView
          key={siteUrl}
          sp={sp}
          siteUrl={siteUrl}
          onBack={() => setView('home')}
          onNavigateToUserAccess={handleNavigateToUserAccess}
        />
      )}
      {view === 'brokenInheritance' && (
        <BrokenInheritanceView
          key={siteUrl + String(includeHidden)}
          sp={sp}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          onBack={() => setView('home')}
        />
      )}
      {view === 'anonymousLinks' && (
        <AnonymousLinksView
          key={siteUrl}
          sp={sp}
          siteUrl={siteUrl}
          onBack={() => setView('home')}
        />
      )}
      {view === 'settings' && (
        <SettingsView
          includeHidden={includeHidden}
          onIncludeHiddenChange={setIncludeHidden}
          scanConcurrency={scanConcurrency}
          onScanConcurrencyChange={setScanConcurrency}
          groupMemberCap={groupMemberCap}
          onGroupMemberCapChange={setGroupMemberCap}
          onBack={() => setView(prevView)}
        />
      )}

    </FluentProvider>
    </RendererProvider>
    </ErrorBoundary>
  );
};
