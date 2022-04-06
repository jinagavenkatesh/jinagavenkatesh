import { DisplayMode } from "@microsoft/sp-core-library";
import { HttpClient } from "@microsoft/sp-http";  
export interface IHelpDeskProps {
  title: string;
  description: string;
  jiraServiceAccount: string;
  jiraAPIToken: string;
  jiraUrl: string;
  jiraJqlQuery: string;
  jiraCloudId: string;
  jiraDateFilter: string;
  jiraUserFilter: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  httpclient: HttpClient; 
  userEmail: string;
  /* title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void; */
}
