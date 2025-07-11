import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'LeaveAdministrationWebPartStrings';
import styles from './LeaveAdministration.module.scss';

import * as React from 'react';
import * as ReactDom from 'react-dom';

import LeaveAdministration from './components/LeaveAdministration';
import { ILeaveAdministrationProps } from './components/ILeaveAdministrationProps';

export interface ILeaveAdministrationWebPartProps {
  title: string;
  description: string;
  defaultView: string;
  showPendingOnly: boolean;
  itemsPerPage: number;
  allowBulkActions: boolean;
}

export default class LeaveAdministrationWebPart extends BaseClientSideWebPart<ILeaveAdministrationWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ILeaveAdministrationProps> = React.createElement(
          LeaveAdministration,
          {
            context: this.context,
            title: this.properties.title,
            description: this.properties.description,
            defaultView: this.properties.defaultView,
            showPendingOnly: this.properties.showPendingOnly,
            itemsPerPage: this.properties.itemsPerPage,
            allowBulkActions: this.properties.allowBulkActions,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName
          }
        );

    ReactDom.render(element, this.domElement);
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('defaultView', {
                  label: 'Default View',
                  options: [
                    { key: 'pending', text: 'Pending Requests' },
                    { key: 'all', text: 'All Requests' },
                    { key: 'approved', text: 'Approved Requests' },
                    { key: 'rejected', text: 'Rejected Requests' }
                  ]
                }),
                PropertyPaneToggle('showPendingOnly', {
                  label: 'Show Pending Only by Default'
                }),
                PropertyPaneSlider('itemsPerPage', {
                  label: 'Items Per Page',
                  min: 5,
                  max: 50,
                  value: 10,
                  showValue: true,
                  step: 5
                }),
                PropertyPaneToggle('allowBulkActions', {
                  label: 'Allow Bulk Actions'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
