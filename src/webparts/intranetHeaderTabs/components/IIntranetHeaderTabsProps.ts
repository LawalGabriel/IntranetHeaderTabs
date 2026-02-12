/* eslint-disable @typescript-eslint/no-explicit-any */
export interface IIntranetHeaderTabsProps {
  // Header
  headerTitle: string;
  headerTitleColor: string;
  headerBackgroundColor: string;
  headerTextColor: string;
  headerHeight?: number;
  headerTitleFontSize?: string;
  tabsFontSize?: string;

  // Logo
  logoListTitle?: string;

  // Welcome Section
  welcomeBackgroundColor: string;
  welcomeTextColor: string;
  welcomeMessage: string;
  showWelcomeSection?: boolean;
  welcomeSectionHeight?: number;          // ✅ ensure this is present

  // Tabs
  listTitle: string;
  maxTabsToShow: number;

  // Our Homes Dropdown (✅ must be here!)
  ourHomesListName?: string;
  enableOurHomesDropdown?: boolean;
  dropdownBackgroundColor?: string;
  dropdownTextColor?: string;
  dropdownHoverBackgroundColor?: string;
  dropdownFontSize?: string;
  dropdownFontWeight?: string;
  dropdownOpenInNewTab?: boolean;
  dropdownIconColor?: string;

  // SPFx context & user info
  context: any;
  userDisplayName: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
}


export interface IHeaderTab {
  Id: number;
  Title: string;
  Link: string | { Url: string };
  OpenInNewTab: boolean;
  Created: string;
  Status: number;
  Order?: number;
}

export interface IAttachmentFile {
  Title: string;
  AttachmentFiles: { ServerRelativeUrl: string }[];
  Created: string;
  LogoImageUrl?: string;
}

// ===== NEW: Subsite info for OurHomes dropdown =====
export interface ISubsiteInfo {
  Title: string;
  SubsiteLink: string;           // we enforce string; if stored as hyperlink field, we'll extract URL
}