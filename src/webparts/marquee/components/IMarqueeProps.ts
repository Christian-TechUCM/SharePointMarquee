import { SPHttpClient } from '@microsoft/sp-http';

export interface IMarqueeProps {
  description: string;
  randomize: boolean;
  customMessage: string;
  selectedList: string;
  showFieldLabels: boolean;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  isDarkTheme?: boolean;
  environmentMessage?: string;
  hasTeamsContext?: boolean;
  userDisplayName?: string; // Add this line if needed
}
