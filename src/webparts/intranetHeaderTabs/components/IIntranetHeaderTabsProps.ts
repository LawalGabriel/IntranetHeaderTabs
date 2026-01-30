/* eslint-disable @typescript-eslint/no-explicit-any */
export interface IIntranetHeaderTabsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listTitle: string;
  logoListTitle?: string;
  context: any;
  
  // Header Configuration
  headerTitle: string;
  headerTitleColor: string;
  headerBackgroundColor: string;
  headerTextColor: string;
  logoUrl: string;
  
  // Welcome Section Configuration
  welcomeBackgroundColor: string;
  welcomeTextColor: string;
  welcomeMessage: string;
  
  // Tabs Configuration
  maxTabsToShow: number;
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