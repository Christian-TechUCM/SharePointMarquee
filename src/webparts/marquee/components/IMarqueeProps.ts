import { SPHttpClient } from '@microsoft/sp-http';

export interface IMarqueeProps {
  description: string;
  randomize: boolean;
  customMessage: string;
  selectedList: string;
  showFieldLabels: boolean;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  showCustomMessage: boolean;
  headerColor: string;
  customMessageColor: string;
  customMessageBold: boolean;
  imageUrl: string; // Add this line
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string;
}
