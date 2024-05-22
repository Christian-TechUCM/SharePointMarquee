import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'MarqueeWebPartStrings';
import MarqueeComponent from './components/Marquee';
import { IMarqueeProps } from './components/IMarqueeProps';

export interface IMarqueeWebPartProps {
  description: string;
  randomize: boolean;
  customMessage: string;
  selectedList: string;
  showFieldLabels: boolean;
}

export default class MarqueeWebPart extends BaseClientSideWebPart<IMarqueeWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _lists: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IMarqueeProps> = React.createElement(
      MarqueeComponent,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        randomize: this.properties.randomize,
        customMessage: this.properties.customMessage,
        selectedList: this.properties.selectedList,
        showFieldLabels: this.properties.showFieldLabels
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      return this._getLists();
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  private _getLists(): Promise<void> {
    const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 100 and Hidden eq false`;
    return this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: any) => {
        this._lists = data.value.map((list: { Id: any; Title: any; }) => ({
          key: list.Id,
          text: list.Title
        }));
        this.context.propertyPane.refresh();
      });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneToggle('randomize', {
                  label: 'Randomize Order',
                  onText: 'Randomize',
                  offText: 'In Order'
                }),
                PropertyPaneTextField('customMessage', {
                  label: 'Custom Message'
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select a list',
                  options: this._lists
                }),
                PropertyPaneToggle('showFieldLabels', {
                  label: 'Show Field Labels',
                  onText: 'Show',
                  offText: 'Hide'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
