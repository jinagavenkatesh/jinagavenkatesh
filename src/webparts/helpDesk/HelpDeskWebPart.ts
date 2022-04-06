import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelpDeskWebPartStrings';
import HelpDesk from './components/HelpDesk';
import { IHelpDeskProps } from './components/IHelpDeskProps';

import { IJSMService } from "../helpDesk/services/IJSMService";
import { JSMService } from "../helpDesk/services/JSMService";


export interface IHelpDeskWebPartProps {
  title: string;
  description: string;
  jiraServiceAccount: string;
  jiraAPIToken: string;
  jiraUrl: string;
  jiraCloudId: string;
  jiraJqlQuery: string;
  jiraDateFilter: string;
  jiraUserFilter: string;
  otherUserEmail: string;
  /* title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void; */
}

export default class HelpDeskWebPart extends BaseClientSideWebPart<IHelpDeskWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    return super.onInit();
    
  }

  public render(): void {
    const element: React.ReactElement<IHelpDeskProps> = React.createElement(
      HelpDesk,
      {
        title: this.properties.title,
        description: this.properties.description,
        jiraServiceAccount: this.properties.jiraServiceAccount,
        jiraAPIToken: this.properties.jiraAPIToken,
        jiraUrl: this.properties.jiraUrl,
        jiraCloudId: this.properties.jiraCloudId,
        jiraJqlQuery: this.properties.jiraJqlQuery,
        jiraDateFilter:  "last6months",//this.properties.jiraDateFilter,
        jiraUserFilter: this.properties.jiraUserFilter,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        httpclient: this.context.httpClient,
        userEmail: this.properties.jiraUserFilter=="otheruser" ? this.properties.otherUserEmail : this.context.pageContext.user.email
        /* title: "My Tickets",
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        } */
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.HelpDeskGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                })
              ]
            },
            {
              groupName: strings.JiraGroupName,
              groupFields: [
                PropertyPaneTextField('jiraServiceAccount', {
                  label: strings.JiraServiceAccountFieldLabel
                }),
                PropertyPaneTextField('jiraAPIToken', {
                  label: strings.JiraAPITokenFieldLabel
                }),
                PropertyPaneTextField('jiraUrl', {
                  label: strings.JiraUrlFieldLabel
                }),
                PropertyPaneTextField('jiraCloudId', {
                  label: strings.JiraCloudIdFieldLabel
                }),
                PropertyPaneTextField('jiraJqlQuery', {
                  label: strings.JiraJqlQueryFieldLabel
                }),
               /*  PropertyPaneDropdown('jiraDateFilter',{
                  label: strings.JiraDateFilterFieldLabel,
                  options: [{
                    key: 'last6months',
                    text: 'Last 6 Months'
                  },
                  {
                    key: 'last1year',
                    text: 'Last 1 Year'
                  }]
                }), */
                PropertyPaneDropdown('jiraUserFilter',{
                  label: strings.JiraUserFilterFieldLabel,
                  options: [{
                    key: 'currentuser',
                    text: 'Current User'
                  },
                  {
                    key: 'otheruser',
                    text: 'Other User'
                  }]
                }),
                PropertyPaneTextField('otherUserEmail', {
                  label: strings.OtherUserEmailFieldLabel,
                  disabled: this.properties.jiraUserFilter == "currentuser"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
