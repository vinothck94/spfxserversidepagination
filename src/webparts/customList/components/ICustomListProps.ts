import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICustomListProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteUrl: string;
  context: WebPartContext;
}
