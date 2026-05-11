import * as React from 'react';
import { FluentProvider, webLightTheme, createDOMRenderer, RendererProvider } from '@fluentui/react-components';
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

// Catches rendering errors and shows them inline with full detail instead of
// letting SPFx swallow the error and show "[object Object]".
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

  return (
    <ErrorBoundary>
    <RendererProvider renderer={renderer} targetDocument={document}>
    <FluentProvider theme={webLightTheme} style={{ minHeight: '400px' }}>
      {view === 'home' && (
        <HomeView onNavigate={setView} />
      )}
      {view === 'report' && (
        <PermissionsReportView
          context={context}
          sp={sp}
          excel={excel}
          onBack={() => setView('home')}
        />
      )}
      {view === 'explorer' && (
        <PermissionsExplorerView
          context={context}
          sp={sp}
          onBack={() => setView('home')}
        />
      )}
      {view === 'userAccess' && (
        <UserAccessView
          context={context}
          sp={sp}
          onBack={() => setView('home')}
        />
      )}
    </FluentProvider>
    </RendererProvider>
    </ErrorBoundary>
  );
};
