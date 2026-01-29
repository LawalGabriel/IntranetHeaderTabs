import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
 
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'IntranetHeaderTabsWebPartStrings';
import IntranetHeaderTabs from './components/IntranetHeaderTabs';
import { IIntranetHeaderTabsProps } from './components/IIntranetHeaderTabsProps';

export interface IIntranetHeaderTabsWebPartProps {
  description: string;
  headerTitle: string;
  headerBackgroundColor: string;
  headerTextColor: string;
  welcomeBackgroundColor: string;
  welcomeTextColor: string;
  welcomeMessage: string;
  maxTabsToShow: number;
  listTitle: string;
}

export default class IntranetHeaderTabsWebPart extends BaseClientSideWebPart<IIntranetHeaderTabsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // Default property values
  protected getDefaultProperties(): Partial<IIntranetHeaderTabsWebPartProps> {
    return {
      headerTitle: 'THE HUB',
      headerBackgroundColor: '#1a1a1a',
      headerTextColor: '#ffffff',
      welcomeBackgroundColor: '#f3f2f1',
      welcomeTextColor: '#323130',
      welcomeMessage: 'Welcome, {user}',
      maxTabsToShow: 0, // 0 means show all
      listTitle: 'HeaderTabs'
    };
  }

  public render(): void {
    const element: React.ReactElement<IIntranetHeaderTabsProps> = React.createElement(
      IntranetHeaderTabs,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        listTitle: this.properties.listTitle,
        context: this.context,
        headerTitle: this.properties.headerTitle,
        headerBackgroundColor: this.properties.headerBackgroundColor,
        headerTextColor: this.properties.headerTextColor,
        welcomeBackgroundColor: this.properties.welcomeBackgroundColor,
        welcomeTextColor: this.properties.welcomeTextColor,
        welcomeMessage: this.properties.welcomeMessage,
        maxTabsToShow: this.properties.maxTabsToShow
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
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
              groupName: 'Header Configuration',
              groupFields: [
                PropertyPaneTextField('headerTitle', {
                  label: 'Header Title',
                  value: this.properties.headerTitle
                }),
                PropertyPaneTextField('headerBackgroundColor', {
                  label: 'Header Background Color',
                  value: this.properties.headerBackgroundColor,
                  description: 'Enter hex color code (e.g., #1a1a1a)'
                }),
                PropertyPaneTextField('headerTextColor', {
                  label: 'Header Text Color',
                  value: this.properties.headerTextColor,
                  description: 'Enter hex color code (e.g., #ffffff)'
                })
              ]
            },
            {
              groupName: 'Welcome Section',
              groupFields: [
                PropertyPaneTextField('welcomeMessage', {
                  label: 'Welcome Message',
                  multiline: true,
                  rows: 3,
                  value: this.properties.welcomeMessage,
                  description: 'Use {user} to insert the user\'s name'
                }),
                PropertyPaneTextField('welcomeBackgroundColor', {
                  label: 'Background Color',
                  value: this.properties.welcomeBackgroundColor,
                  description: 'Enter hex color code (e.g., #f3f2f1)'
                }),
                PropertyPaneTextField('welcomeTextColor', {
                  label: 'Text Color',
                  value: this.properties.welcomeTextColor,
                  description: 'Enter hex color code (e.g., #323130)'
                })
              ]
            },
            {
              groupName: 'Tabs Configuration',
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: 'SharePoint List Title',
                  value: this.properties.listTitle,
                  description: 'Name of the list containing header tabs'
                }),
                PropertyPaneSlider('maxTabsToShow', {
                  label: 'Maximum Tabs to Show',
                  min: 0,
                  max: 20,
                  step: 1,
                  value: this.properties.maxTabsToShow,
                  showValue: true,
                  //description: '0 = Show all tabs'
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}