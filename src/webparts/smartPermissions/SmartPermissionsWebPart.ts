import { Version } from '@microsoft/sp-core-library';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider, IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { App, AppView, IBrandColors } from './components/App';
import { SharePointService } from './services/SharePointService';
import { ExcelExportService } from './services/ExcelExportService';
import * as strings from 'SmartPermissionsWebPartStrings';

export interface ISmartPermissionsWebPartProps {
  defaultView: AppView;
  hideReport?: boolean;
  hideExplorer?: boolean;
  hideUserAccess?: boolean;
  /** Site the part opens on instead of the page's own site. */
  defaultSiteUrl?: string;
  /** When true, Change URL is hidden so viewers can't point the part at another site. */
  lockSiteUrl?: boolean;
}

export default class SmartPermissionsWebPart extends BaseClientSideWebPart<ISmartPermissionsWebPartProps> {
  private _sp: SharePointService;
  private _excel: ExcelExportService;
  private _isInverted = false;
  private _brandColors: IBrandColors = {
    primary:  '#0078d4',
    darkAlt:  '#106ebe',
    dark:     '#005a9e',
    darker:   '#004578',
    light:    '#c7e0f4',
    lighter:  '#deecf9',
  };

  protected onInit(): Promise<void> {
    // Initialize services first so this._sp is defined before any render() call.
    try {
      this._sp = new SharePointService(this.context);
      this._excel = new ExcelExportService();
    } catch (err: any) {
      return Promise.reject(
        new Error(`[SmartPermissions] Service init failed: ${err?.message ?? String(err)}\n${err?.stack ?? ''}`)
      );
    }

    // Read the current SharePoint site theme color and re-render when it changes.
    try {
      const themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      const applyTheme = (theme: IReadonlyTheme | undefined): void => {
        const p = theme?.palette;
        this._isInverted = !!theme?.isInverted;
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

    return super.onInit();
  }

  public render(): void {
    try {
      const element = React.createElement(App, {
        context: this.context,
        sp: this._sp,
        excel: this._excel,
        defaultView: this.properties.defaultView ?? 'home',
        hiddenViews: [
          this.properties.hideReport ? 'report' : undefined,
          this.properties.hideExplorer ? 'explorer' : undefined,
          this.properties.hideUserAccess ? 'userAccess' : undefined,
        ].filter((v): v is AppView => !!v),
        defaultSiteUrl: this.properties.defaultSiteUrl,
        lockSiteUrl: !!this.properties.lockSiteUrl,
        brandColors: this._brandColors,
        isInverted: this._isInverted,
      });
      ReactDom.render(element, this.domElement);
    } catch (err: any) {
      const msg = err?.message ?? String(err);
      const stack = err?.stack ?? '(no stack)';
      this.domElement.innerHTML =
        `<div style="padding:16px;font-family:Consolas,monospace;font-size:13px;` +
        `background:#fff3f3;color:#242424;border:1px solid #c00;border-radius:4px;margin:8px">` +
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
        label: strings.PropertyPane_DefaultViewLabel,
        options: [
          { key: 'home', text: strings.PropertyPane_View_Home },
          { key: 'report', text: strings.PropertyPane_View_Report },
          { key: 'explorer', text: strings.PropertyPane_View_Explorer },
          { key: 'userAccess', text: strings.PropertyPane_View_UserAccess },
        ],
        selectedKey: this.properties.defaultView ?? 'home',
      },
    };
    // PropertyPaneFieldType.TextField = 3, Toggle = 5 (same stability note as above).
    const toggle = (targetProperty: string, label: string): any => ({
      type: 5,
      targetProperty,
      properties: { label, checked: !!(this.properties as any)[targetProperty] },
    });
    const siteUrlField: any = {
      type: 3,
      targetProperty: 'defaultSiteUrl',
      properties: {
        label: strings.PropertyPane_DefaultSiteUrlLabel,
        description: strings.PropertyPane_DefaultSiteUrlDescription,
        value: this.properties.defaultSiteUrl ?? '',
      },
    };
    return {
      pages: [{
        header: { description: strings.PropertyPane_HeaderDescription },
        groups: [
          {
            groupName: strings.PropertyPane_GroupName,
            groupFields: [dropdownField],
          },
          {
            groupName: strings.PropertyPane_ToolsGroupName,
            groupFields: [
              toggle('hideReport', strings.PropertyPane_HideReport),
              toggle('hideExplorer', strings.PropertyPane_HideExplorer),
              toggle('hideUserAccess', strings.PropertyPane_HideUserAccess),
            ],
          },
          {
            groupName: strings.PropertyPane_SiteGroupName,
            groupFields: [siteUrlField, toggle('lockSiteUrl', strings.PropertyPane_LockSiteUrl)],
          },
        ],
      }],
    };
  }
}
