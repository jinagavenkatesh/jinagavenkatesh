declare interface IHelpDeskWebPartStrings {
  PropertyPaneDescription: string;
  HelpDeskGroupName: string;
  TitleFieldLabel: string;
  DescriptionFieldLabel: string;
  JiraGroupName: string;
  JiraServiceAccountFieldLabel: string;
  JiraAPITokenFieldLabel: string;
  JiraUrlFieldLabel: string;
  JiraCloudIdFieldLabel: string;
  JiraJqlQueryFieldLabel: string;
  JiraDateFilterFieldLabel: string;
  JiraUserFilterFieldLabel: string;
  OtherUserEmailFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'HelpDeskWebPartStrings' {
  const strings: IHelpDeskWebPartStrings;
  export = strings;
}
