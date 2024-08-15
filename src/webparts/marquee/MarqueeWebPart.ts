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

import * as strings from 'MarqueeWebPartStrings'; // Import localization strings
import MarqueeComponent from './components/Marquee'; // Import the React component
import { IMarqueeProps } from './components/IMarqueeProps'; // Import the props interface

// Define the properties interface for the web part
export interface IMarqueeWebPartProps {
  description: string;
  randomize: boolean;
  customMessage: string;
  selectedList: string;
  showFieldLabels: boolean;
  showCustomMessage: boolean;
  headerColor: string;
  customMessageColor: string;
  customMessageBold: boolean;
  imageUrl: string; // Add this line
}

// Main class for the web part, extending BaseClientSideWebPart
export default class MarqueeWebPart extends BaseClientSideWebPart<IMarqueeWebPartProps> {
  private _isDarkTheme: boolean = false; // Tracks if the current theme is dark
  private _environmentMessage: string = ''; // Message specific to the environment
  private _lists: IPropertyPaneDropdownOption[] = []; // Holds SharePoint lists for dropdown selection

  // Render method to create and render the React component
  public render(): void {
    const element: React.ReactElement<IMarqueeProps> = React.createElement(
      MarqueeComponent,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams, // Check if in a Teams context
        userDisplayName: this.context.pageContext.user.displayName, // Get current user's display name
        spHttpClient: this.context.spHttpClient, // SharePoint HTTP client for API calls
        siteUrl: this.context.pageContext.web.absoluteUrl, // Current site URL
        randomize: this.properties.randomize,
        customMessage: this.properties.customMessage,
        selectedList: this.properties.selectedList,
        showFieldLabels: this.properties.showFieldLabels,
        showCustomMessage: this.properties.showCustomMessage,
        headerColor: this.properties.headerColor,
        customMessageColor: this.properties.customMessageColor,
        customMessageBold: this.properties.customMessageBold,
        imageUrl: this.properties.imageUrl // Add this line
      }
    );

    ReactDom.render(element, this.domElement); // Render the React component in the DOM
  }

  // onInit method to initialize the web part, including fetching environment message and lists
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message; // Set the environment message
      return this._getLists(); // Fetch available SharePoint lists
    });
  }

  // Method to get a message specific to the environment (e.g., SharePoint, Teams)
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // Check if running in Teams
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

    // Default to SharePoint environment message
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  // Method to fetch SharePoint lists for the property pane dropdown
  private _getLists(): Promise<void> {
    const listUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 100 and Hidden eq false`;
    return this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: any) => {
        this._lists = data.value.map((list: { Id: any; Title: any; }) => ({
          key: list.Id,
          text: list.Title
        }));
        this.context.propertyPane.refresh(); // Refresh the property pane to show the lists
      });
  }

  // Method to handle theme changes (e.g., light to dark mode)
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted; // Update theme status
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  // Cleanup method when the web part is disposed
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // Define the data version for this web part
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Configuration for the property pane (UI to configure web part properties)
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription // Header description
          },
          groups: [
            {
              groupName: strings.BasicGroupName, // Group name for property fields
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel // Description field
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select a list', // Dropdown for selecting a SharePoint list
                  options: this._lists
                }),
                PropertyPaneToggle('randomize', {
                  label: 'Randomize Order', // Toggle for randomizing order
                  onText: 'Randomize',
                  offText: 'In Order'
                }),
                PropertyPaneToggle('showFieldLabels', {
                  label: 'Show Field Labels', // Toggle for showing field labels
                  onText: 'Show',
                  offText: 'Hide'
                }),
                
                PropertyPaneTextField('customMessage', {
                  label: 'Custom Message' // Text field for custom message
                }),
                PropertyPaneToggle('showCustomMessage', {
                  label: 'Show Custom Message', // Toggle for showing custom message
                  onText: 'Show',
                  offText: 'Hide'
                }),
                
                PropertyPaneToggle('customMessageBold', { // Toggle for bolding custom message
                  label: 'Bold Custom Message',
                  onText: 'Bold',
                  offText: 'Normal'
                }),
                PropertyPaneTextField('headerColor', {
                  label: 'Header Color (HEX)', // Text field for header color
                  description: 'Enter the HEX value for the header color, e.g., #FF5733',
                  value: this.properties.headerColor
                }),
                PropertyPaneTextField('customMessageColor', { // Text field for custom message color
                  label: 'Custom Message Color (HEX)',
                  description: 'Enter the HEX value for the custom message color, e.g., #FF5733',
                  value: this.properties.customMessageColor
                }),
                PropertyPaneTextField('imageUrl', { // Text field for image URL
                  label: 'Image URL',
                  description: 'Enter the URL of the image to display above the marquee',
                  value: this.properties.imageUrl
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
