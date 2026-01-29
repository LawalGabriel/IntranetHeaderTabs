/* eslint-disable @typescript-eslint/no-explicit-any */
export interface IIntranetHeaderTabsProps {
 
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listTitle: string;
  context: any;
  errorContainer?: string;
  
  // New properties
  headerTitle: string;
  headerBackgroundColor: string;
  headerTextColor: string;
  welcomeBackgroundColor: string;
  welcomeTextColor: string;
  welcomeMessage: string;
  maxTabsToShow: number;
  loadingcontainer?: string;
  description?: string;
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