import { Version } from '@microsoft/sp-core-library';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider, IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { App, AppView, IBrandColors } from './components/App';
import { SharePointService } from './services/SharePointService';
import { ExcelExportService } from './services/ExcelExportService';

// This line runs the instant the module is evaluated — if you see it in the
// browser console, the new bundle is definitely loading.
console.error('[SmartPermissions] *** MODULE EVALUATED ***', new Date().toISOString());

export interface ISmartPermissionsWebPartProps {
  defaultView: AppView;
}

export default class SmartPermissionsWebPart extends BaseClientSideWebPart<ISmartPermissionsWebPartProps> {
  private _sp: SharePointService;
  private _excel: ExcelExportService;
  private _brandColors: IBrandColors = {
    primary:  '#0078d4',
    darkAlt:  '#106ebe',
    dark:     '#005a9e',
    darker:   '#004578',
    light:    '#c7e0f4',
    lighter:  '#deecf9',
  };

  protected onInit(): Promise<void> {
    // Log every unhandled rejection so the real error appears in the browser
    // console even when SPFx swallows it and shows "[object Object]".
    window.addEventListener('unhandledrejection', (ev: PromiseRejectionEvent) => {
      const r = ev.reason;
      console.error(
        '[SmartPermissions] unhandledrejection — open this to see the real error:',
        r,
        'JSON:', (() => { try { return JSON.stringify(r, null, 2); } catch { return '(circular)'; } })(),
        'message:', r?.message,
        'stack:', r?.stack,
        'string:', String(r),
      );
    });

    // Initialise services first so this._sp is defined before any render() call.
    try {
      this._sp = new SharePointService(this.context);
      this._excel = new ExcelExportService();
    } catch (err: any) {
      return Promise.reject(
        new Error(`[SmartPermissions] Service init failed: ${err?.message ?? String(err)}\n${err?.stack ?? ''}`)
      );
    }

    // Read the current SharePoint site theme colour and re-render when it changes.
    try {
      const themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      const applyTheme = (theme: IReadonlyTheme | undefined): void => {
        const p = theme?.palette;
        if (p?.themePrimary) {
          this._brandColors = {
            primary:  p.themePrimary,
            darkAlt:  p.themeDarkAlt  ?? p.themePrimary,
            dark:     p.themeDark     ?? p.themePrimary,
            darker:   p.themeDarker   ?? p.themeDark ?? p.themePrimary,
            light:    p.themeLight    ?? '#c7e0f4',
            lighter:  p.themeLighter  ?? '#deecf9',
          };
        }
        this.render();
      };
      applyTheme(themeProvider.tryGetTheme());
      themeProvider.themeChangedEvent.add(this, (args) => applyTheme(args.theme));
    } catch { /* theme unavailable — keep default blue */ }

    return super.onInit().catch((err: any) => {
      const detail =
        `type=${typeof err} | ` +
        `string=${String(err)} | ` +
        `message=${err?.message} | ` +
        `json=${(() => { try { return JSON.stringify(err, null, 2); } catch { return '(circular)'; } })()}`;
      console.error('[SmartPermissions] super.onInit() rejected:', detail, err);
      throw new Error(`[SmartPermissions] onInit failed — ${detail}`);
    });
  }

  public render(): void {
    try {
      const element = React.createElement(App, {
        context: this.context,
        sp: this._sp,
        excel: this._excel,
        defaultView: this.properties.defaultView ?? 'home',
        brandColors: this._brandColors,
      });
      ReactDom.render(element, this.domElement);
    } catch (err: any) {
      const msg = err?.message ?? String(err);
      const stack = err?.stack ?? '(no stack)';
      this.domElement.innerHTML =
        `<div style="padding:16px;font-family:Consolas,monospace;font-size:13px;` +
        `background:#fff3f3;border:1px solid #c00;border-radius:4px;margin:8px">` +
        `<strong style="color:#c00;font-size:14px">Smart Permissions — Startup Error</strong><br><br>` +
        `<strong>Message:</strong> ${this._escHtml(msg)}<br><br>` +
        `<strong>Stack:</strong><pre style="font-size:11px;white-space:pre-wrap;` +
        `background:#f5f5f5;padding:8px;margin:4px 0;border-radius:2px">` +
        `${this._escHtml(stack)}</pre></div>`;
    }
  }

  private _escHtml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Construct the dropdown field descriptor directly to avoid any runtime
    // import of @microsoft/sp-property-pane (which is an AMD external and
    // cannot be require()'d dynamically in the workbench without a CSP error).
    // PropertyPaneFieldType.Dropdown = 6 (stable since SPFx 1.x).
    const dropdownField: any = {
      type: 6,
      targetProperty: 'defaultView',
      properties: {
        label: 'Default view on open',
        options: [
          { key: 'home', text: 'Home' },
          { key: 'report', text: 'Permissions Report' },
          { key: 'explorer', text: 'Permissions Explorer' },
          { key: 'userAccess', text: 'User Access' },
        ],
        selectedKey: this.properties.defaultView ?? 'home',
      },
    };
    return {
      pages: [{
        header: { description: 'Smart Permissions configuration' },
        groups: [{
          groupName: 'General',
          groupFields: [dropdownField],
        }],
      }],
    };
  }
}
