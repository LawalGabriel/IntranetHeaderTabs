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
  welcomeSectionHeight: number | undefined;
  tabsFontSize: string | undefined;
  headerTitleFontSize: string | undefined;
  headerHeight: number | undefined;
  logoListTitle: string;
  description: string;
  ourHomesListName: string;
  enableOurHomesDropdown: boolean;
  dropdownBackgroundColor: string;
  dropdownTextColor: string;
  dropdownHoverBackgroundColor: string;
  dropdownFontSize: string;
  dropdownFontWeight: string;
  dropdownOpenInNewTab: boolean;
  dropdownIconColor: string;
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
    headerHeight: 60,
    headerTitleFontSize: '24px',
    tabsFontSize: '16px',

    logoListTitle: 'HeaderImages',

    welcomeBackgroundColor: '#f3f2f1',
    welcomeTextColor: '#323130',
    welcomeMessage: 'Welcome, {user}',
    showWelcomeSection: true,
    welcomeSectionHeight: 200,           // ✅ default height

    listTitle: 'HeaderTabs',
    maxTabsToShow: 0,

    // Our Homes defaults
    ourHomesListName: 'OurHomes',
    enableOurHomesDropdown: true,
    dropdownBackgroundColor: '#ffffff',
    dropdownTextColor: '#323130',
    dropdownHoverBackgroundColor: '#f3f2f1',
    dropdownFontSize: '14px',
    dropdownFontWeight: '400',
    dropdownOpenInNewTab: false,
    dropdownIconColor: '#ffffff',
  };
}

public render(): void {
  const element: React.ReactElement<IIntranetHeaderTabsProps> = React.createElement(
    IntranetHeaderTabs,
    {
      // ----- Header -----
      headerTitle: this.properties.headerTitle,
      headerTitleColor: this.properties.headerTitleColor,
      headerBackgroundColor: this.properties.headerBackgroundColor,
      headerTextColor: this.properties.headerTextColor,
      headerHeight: this.properties.headerHeight,
      headerTitleFontSize: this.properties.headerTitleFontSize,
      tabsFontSize: this.properties.tabsFontSize,

      // ----- Logo -----
      logoListTitle: this.properties.logoListTitle,

      // ----- Welcome Section -----
      welcomeBackgroundColor: this.properties.welcomeBackgroundColor,
      welcomeTextColor: this.properties.welcomeTextColor,
      welcomeMessage: this.properties.welcomeMessage,
      showWelcomeSection: this.properties.showWelcomeSection,
      welcomeSectionHeight: this.properties.welcomeSectionHeight,   // ✅ already there

      // ----- Tabs -----
      listTitle: this.properties.listTitle,
      maxTabsToShow: this.properties.maxTabsToShow,

      // ----- Our Homes Dropdown (NEW – add these lines!) -----
      ourHomesListName: this.properties.ourHomesListName,
      enableOurHomesDropdown: this.properties.enableOurHomesDropdown,
      dropdownBackgroundColor: this.properties.dropdownBackgroundColor,
      dropdownTextColor: this.properties.dropdownTextColor,
      dropdownHoverBackgroundColor: this.properties.dropdownHoverBackgroundColor,
      dropdownFontSize: this.properties.dropdownFontSize,
      dropdownFontWeight: this.properties.dropdownFontWeight,
      dropdownOpenInNewTab: this.properties.dropdownOpenInNewTab,
      dropdownIconColor: this.properties.dropdownIconColor,

      // ----- SPFx required -----
      context: this.context,
      userDisplayName: this.context.pageContext.user.displayName,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
    }
  );

  ReactDom.render(element, this.domElement);
}


  protected onInit(): Promise<void> {
    // Ensure Default Chrome Is Disabled
    this._ensureDefaultChromeIsDisabled();
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


  private _ensureDefaultChromeIsDisabled(): void {
    //'.SuiteNavWrapper','#spSiteHeader','.sp-appBar','#sp-appBar','#workbenchPageContent', '.SPCanvas-canvas', '.CanvasZone', '.ms-CommandBar', '#spSiteHeader', '.commandBarWrapper'
    const displayElements = ['#SuiteNavWrapper', '#spSiteHeader', '.sp-appBar', '.ms-CommandBar', '.commandBarWrapper', '#spCommandBar', '.ms-SPLegacyFabric', '.ms-footer', '.sp-pageLayout-footer', '.ms-workbenchFooter'];
    const widthElements = ["#workbenchPageContent", '.CanvasZone', '.SPCanvas-canvas'];
    displayElements.forEach(selector => {
      document.querySelectorAll(selector).forEach(element => {
        (element as HTMLElement).style.display = 'none';
      });
    });

    widthElements.forEach(selector => {
      document.querySelectorAll(selector).forEach(element => {
        (element as HTMLElement).style.maxWidth = 'none';
      });
    });

    //document.querySelector<HTMLElement>('#spCommandBar')?.style.setProperty('min-height', '0', 'important');




    // Check if the device is mobile and apply mobile-specific styles
    if (this._isMobileDevice()) {
      const mobileElements = [
        '.spMobileHeader',
        '.ms-FocusZone',
        '.ms-CommandBar',
        '.spMobileNav',
        '#O365_MainLink_NavContainer', // Waffle (App Launcher)
        '.ms-Nav', // Additional possible mobile navigation elements
        '.ms-Nav-item' // Possible item within the navigation
      ];
      mobileElements.forEach(selector => {
        document.querySelectorAll(selector).forEach(element => {
          (element as HTMLElement).style.display = 'none';
        });
      });
    }
  }

  private _isMobileDevice(): boolean {
    // Check if the user is on a mobile device based on user agent or screen width
    const userAgent = navigator.userAgent.toLowerCase();
    const isMobile = /iphone|ipod|ipad|android|blackberry|windows phone/i.test(userAgent);
    return isMobile || window.innerWidth <= 768; // Custom breakpoint for mobile
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
        header: { description: strings.PropertyPaneDescription },
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
                placeholder: '#ffffff'
              }),
              PropertyPaneTextField('headerBackgroundColor', {
                label: 'Header Background Color',
                placeholder: '#1a1a1a'
              }),
              PropertyPaneTextField('headerTextColor', {
                label: 'Header Tabs Text Color',
                placeholder: '#ffffff'
              }),
              PropertyPaneTextField('headerHeight', {
                label: 'Header Height (px)',
                placeholder: '60'
              }),
              PropertyPaneTextField('headerTitleFontSize', {
                label: 'Header Title Font Size',
                placeholder: '24px'
              }),
              PropertyPaneTextField('tabsFontSize', {
                label: 'Tabs Font Size',
                placeholder: '16px'
              }),
            ]
          },
          {
            groupName: 'Logo Configuration',
            groupFields: [
              PropertyPaneTextField('logoListTitle', {
                label: 'Logo SharePoint List Name',
                value: this.properties.logoListTitle,
                description: 'List with attachments containing the header logo'
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
                description: 'Use {user} to insert the user\'s name'
              }),
              PropertyPaneTextField('welcomeBackgroundColor', {
                label: 'Background Color',
                placeholder: '#f3f2f1'
              }),
              PropertyPaneTextField('welcomeTextColor', {
                label: 'Text Color',
                placeholder: '#323130'
              }),
              PropertyPaneToggle('showWelcomeSection', {
                label: 'Show Welcome Section',
                onText: 'Visible',
                offText: 'Hidden'
              }),
              PropertyPaneSlider('welcomeSectionHeight', {
                label: 'Welcome Section Height (px)',
                min: 100,
                max: 500,
                step: 10,
                value: 200
              }),
            ]
          },
          {
            groupName: 'Tabs Configuration',
            groupFields: [
              PropertyPaneTextField('listTitle', {
                label: 'Tabs SharePoint List Name',
                value: this.properties.listTitle
              }),
              PropertyPaneSlider('maxTabsToShow', {
                label: 'Maximum Tabs to Show',
                min: 0,
                max: 20,
                step: 1,
                showValue: true,
                //description: '0 = Show all tabs'
              })
            ]
          },
          {
            groupName: 'Our Homes Dropdown',   // ✅ new dedicated group
            groupFields: [
              PropertyPaneTextField('ourHomesListName', {
                label: 'OurHomes List Name',
                value: 'OurHomes'
              }),
              PropertyPaneToggle('enableOurHomesDropdown', {
                label: 'Enable Our Homes Dropdown',
                checked: true
              }),
              PropertyPaneTextField('dropdownBackgroundColor', {
                label: 'Dropdown Background Color',
                placeholder: '#ffffff'
              }),
              PropertyPaneTextField('dropdownTextColor', {
                label: 'Dropdown Text Color',
                placeholder: '#323130'
              }),
              PropertyPaneTextField('dropdownHoverBackgroundColor', {
                label: 'Dropdown Hover Background',
                placeholder: '#f3f2f1'
              }),
              PropertyPaneTextField('dropdownFontSize', {
                label: 'Dropdown Font Size',
                placeholder: '14px'
              }),
              PropertyPaneTextField('dropdownFontWeight', {
                label: 'Dropdown Font Weight',
                placeholder: '400'
              }),
              PropertyPaneToggle('dropdownOpenInNewTab', {
                label: 'Open Subsite Links in New Tab'
              }),
              PropertyPaneTextField('dropdownIconColor', {
                label: 'Dropdown Icon Color',
                placeholder: '#ffffff'
              })
            ]
          }
        ]
      }
    ]
  };
}
}