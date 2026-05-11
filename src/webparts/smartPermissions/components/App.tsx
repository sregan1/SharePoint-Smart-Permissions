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

const renderer = createDOMRenderer(document);

export const App: React.FC<AppProps> = ({ context, sp, excel, defaultView }) => {
  const [view, setView] = React.useState<AppView>(defaultView ?? 'home');

  return (
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
  );
};
