import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { App, AppView } from './components/App';
import { SharePointService } from './services/SharePointService';
import { ExcelExportService } from './services/ExcelExportService';

export interface ISmartPermissionsWebPartProps {
  defaultView: AppView;
}

export default class SmartPermissionsWebPart extends BaseClientSideWebPart<ISmartPermissionsWebPartProps> {
  private _sp: SharePointService;
  private _excel: ExcelExportService;

  protected onInit(): Promise<void> {
    this._sp = new SharePointService(this.context);
    this._excel = new ExcelExportService();
    return super.onInit();
  }

  public render(): void {
    const element = React.createElement(App, {
      context: this.context,
      sp: this._sp,
      excel: this._excel,
      defaultView: this.properties.defaultView ?? 'home',
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configure Smart Permissions' },
          groups: [
            {
              groupName: 'Default View',
              groupFields: [
                PropertyPaneDropdown('defaultView', {
                  label: 'Open to',
                  options: [
                    { key: 'home', text: 'Home Screen' },
                    { key: 'report', text: 'Permissions Report' },
                    { key: 'explorer', text: 'Permissions Explorer' },
                    { key: 'userAccess', text: 'User Access' },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
