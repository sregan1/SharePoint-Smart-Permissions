import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  createDOMRenderer,
  RendererProvider,
  Button,
  Input,
  Text,
  Checkbox,
  Tooltip,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  tokens,
} from '@fluentui/react-components';
import { Settings24Regular, Info16Regular, ShieldLock24Regular } from '@fluentui/react-icons';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { SharePointService } from '../services/SharePointService';
import { ExcelExportService } from '../services/ExcelExportService';
import { HomeView } from './HomeView';
import { PermissionsReportView } from './PermissionsReportView';
import { PermissionsExplorerView } from './PermissionsExplorerView';
import { UserAccessView } from './UserAccessView';

export type AppView = 'home' | 'report' | 'explorer' | 'userAccess';

export interface AppProps {
  context: WebPartContext;
  sp: SharePointService;
  excel: ExcelExportService;
  defaultView?: AppView;
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

export const App: React.FC<AppProps> = ({ context, sp, excel, defaultView }) => {
  const [view, setView] = React.useState<AppView>(defaultView ?? 'home');
  const [siteUrl, setSiteUrl] = React.useState(context.pageContext.web.absoluteUrl);
  const [editUrl, setEditUrl] = React.useState(context.pageContext.web.absoluteUrl);
  const [isEditing, setIsEditing] = React.useState(false);
  const [includeHidden, setIncludeHidden] = React.useState(false);
  const [settingsOpen, setSettingsOpen] = React.useState(false);

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

  const renderSettingsSurface = () => (
    <PopoverSurface>
      <div style={{ padding: '12px 16px', minWidth: '280px' }}>
        <Text weight="semibold" style={{ display: 'block', marginBottom: '12px' }}>
          Settings
        </Text>
        <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
          <Checkbox
            label="Include system and hidden libraries"
            checked={includeHidden}
            onChange={(_, d) => setIncludeHidden(!!d.checked)}
          />
          <Tooltip
            content="When checked, includes system and hidden libraries such as Style Library, Form Templates, Site Assets, and others not shown in default views. Applies to Permissions Explorer and User Access."
            relationship="description"
            withArrow
          >
            <Button
              appearance="transparent"
              icon={<Info16Regular />}
              size="small"
              style={{ minWidth: 'unset', padding: '2px' }}
              aria-label="More info about hidden libraries"
            />
          </Tooltip>
        </div>
      </div>
    </PopoverSurface>
  );

  return (
    <ErrorBoundary>
    <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={webLightTheme} style={{ minHeight: '400px', position: 'relative' }}>
      {view !== 'home' ? (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'auto 1fr auto',
            alignItems: 'center',
            paddingTop: tokens.spacingVerticalS,
            paddingBottom: tokens.spacingVerticalS,
            paddingLeft: tokens.spacingHorizontalM,
            paddingRight: tokens.spacingHorizontalS,
            background: tokens.colorBrandBackground,
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
          <Popover open={settingsOpen} onOpenChange={(_, d) => setSettingsOpen(d.open)}>
            <PopoverTrigger disableButtonEnhancement>
              <Button
                appearance="transparent"
                icon={<Settings24Regular style={{ color: 'white' }} />}
                aria-label="Settings"
                title="Settings"
              />
            </PopoverTrigger>
            {renderSettingsSurface()}
          </Popover>
        </div>
      ) : (
        <div style={{ position: 'absolute', top: '4px', right: '8px', zIndex: 10 }}>
          <Popover open={settingsOpen} onOpenChange={(_, d) => setSettingsOpen(d.open)}>
            <PopoverTrigger disableButtonEnhancement>
              <Button
                appearance="transparent"
                icon={<Settings24Regular style={{ color: 'white' }} />}
                aria-label="Settings"
                title="Settings"
              />
            </PopoverTrigger>
            {renderSettingsSurface()}
          </Popover>
        </div>
      )}

      {view === 'home' && (
        <HomeView onNavigate={setView} />
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
        />
      )}
      {view === 'userAccess' && (
        <UserAccessView
          key={siteUrl + String(includeHidden)}
          sp={sp}
          siteUrl={siteUrl}
          includeHidden={includeHidden}
          onBack={() => setView('home')}
        />
      )}
    </FluentProvider>
    </RendererProvider>
    </ErrorBoundary>
  );
};
