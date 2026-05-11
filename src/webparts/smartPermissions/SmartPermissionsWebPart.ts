import { Version } from '@microsoft/sp-core-library';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { App, AppView } from './components/App';
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

    try {
      this._sp = new SharePointService(this.context);
      this._excel = new ExcelExportService();
    } catch (err: any) {
      return Promise.reject(
        new Error(`[SmartPermissions] Service init failed: ${err?.message ?? String(err)}\n${err?.stack ?? ''}`)
      );
    }

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
    return { pages: [] };
  }
}
