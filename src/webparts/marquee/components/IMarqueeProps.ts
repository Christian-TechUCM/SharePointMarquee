import { SPHttpClient } from '@microsoft/sp-http';
//The IMarqueeProps.ts file defines an interface that ensures the necessary data and 
//configuration settings are passed to the React component in a structured, 
//type-safe manner, improving code clarity, maintainability, and reusability.
// It also allows the component to interact with SharePoint-specific resources like 
//lists and APIs, enabling dynamic and customizable web part behavior.

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
  imageWidth?: string; // Optional image width
  imageHeight?: string; // Optional image height
}
