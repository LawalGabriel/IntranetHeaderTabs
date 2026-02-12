import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'IntranetHeaderTabsWebPartStrings';
import IntranetHeaderTabs from './components/IntranetHeaderTabs';
import { IIntranetHeaderTabsProps } from './components/IIntranetHeaderTabsProps';

export interface IIntranetHeaderTabsWebPartProps {
  tabsFontSize: string | undefined;
  headerTitleFontSize: string | undefined;
  headerHeight: number | undefined;
  logoListTitle: string;
  description: string;
  headerTitle: string;
  headerTitleColor: string;
  headerBackgroundColor: string;
  headerTextColor: string;
  logoUrl: string;
  welcomeBackgroundColor: string;
  welcomeTextColor: string;
  welcomeMessage: string;
  maxTabsToShow: number;
  listTitle: string;
   showWelcomeSection: boolean;
}

export default class IntranetHeaderTabsWebPart extends BaseClientSideWebPart<IIntranetHeaderTabsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // Default property values
  protected getDefaultProperties(): Partial<IIntranetHeaderTabsWebPartProps> {
    return {
      headerTitle: 'THE HUB',
      headerTitleColor: '#ffffff',
      headerBackgroundColor: '#1a1a1a',
      headerTextColor: '#ffffff',
      logoUrl: '',
      welcomeBackgroundColor: '#f3f2f1',
      welcomeTextColor: '#323130',
      welcomeMessage: 'Welcome, {user}',
      maxTabsToShow: 0,
      listTitle: 'HeaderTabs',
      headerHeight: 60,
    headerTitleFontSize: '24px',
    tabsFontSize: '16px',
     showWelcomeSection: true,
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
        headerTitleColor: this.properties.headerTitleColor,
        headerBackgroundColor: this.properties.headerBackgroundColor,
        headerTextColor: this.properties.headerTextColor,
        logoUrl: this.properties.logoUrl,
        welcomeBackgroundColor: this.properties.welcomeBackgroundColor,
        welcomeTextColor: this.properties.welcomeTextColor,
        welcomeMessage: this.properties.welcomeMessage,
        maxTabsToShow: this.properties.maxTabsToShow,
        logoListTitle: this.properties.logoListTitle || "LogoList",
         headerHeight: this.properties.headerHeight,
      headerTitleFontSize: this.properties.headerTitleFontSize,
      tabsFontSize: this.properties.tabsFontSize,
        showWelcomeSection: this.properties.showWelcomeSection !== false 
        
        
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
                PropertyPaneTextField('headerTitleColor', {
                  label: 'Header Title Color',
                  value: this.properties.headerTitleColor,
                  description: 'Enter hex color code (e.g., #ffffff)'
                }),
                PropertyPaneTextField('headerBackgroundColor', {
                  label: 'Header Background Color',
                  value: this.properties.headerBackgroundColor,
                  description: 'Enter hex color code (e.g., #1a1a1a)'
                }),
                PropertyPaneTextField('headerTextColor', {
                  label: 'Header Tabs Text Color',
                  value: this.properties.headerTextColor,
                  description: 'Enter hex color code (e.g., #ffffff)'
                }),
                PropertyPaneTextField('logoUrl', {
                  label: 'Logo Image URL',
                  value: this.properties.logoUrl,
                  description: 'Enter the full URL of your logo image'
                }),
                PropertyPaneTextField('headerHeight', {
                label: 'Header Height (px)',
                description: 'Set the height of the header in pixels (e.g., 60)',
              }),
              PropertyPaneTextField('headerTitleFontSize', {
                label: 'Header Title Font Size',
                description: 'Set the font size for header title (e.g., 24px, 1.5rem)',
              }),
              PropertyPaneTextField('tabsFontSize', {
                label: 'Tabs Font Size',
                description: 'Set the font size for navigation tabs (e.g., 16px, 1rem)',
              }),
              
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
                }),
                PropertyPaneToggle('showWelcomeSection', {
                label: 'Show Welcome Section',
                onText: 'Visible',
                offText: 'Hidden',
                checked: true   // default state
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