import { IJSMService } from "./IJSMService";
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { IIssueType, IItem, IUser } from "../components/IHelpDeskState";
import * as _ from "lodash";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export class JSMService implements IJSMService{

    private _httpClient: HttpClient;
    private _context: WebPartContext;
    private _jiraServiceAccount: string;
    private _jiraAPIToken: string;
    private _jiraUrl: string;
    private _jiraCloudId: string;
    private _items: any;
    
    constructor(
        httpClient: HttpClient, 
        context: WebPartContext, 
        jiraServiceAccount: string,
        jiraAPIToken: string,
        jiraUrl: string,
        jiraCloudId: string,
        ) {
            this._httpClient = httpClient;
            this._context = context;
            this._jiraServiceAccount = jiraServiceAccount;
            this._jiraAPIToken = jiraAPIToken;
            this._jiraUrl = jiraUrl;
            this._jiraCloudId = jiraCloudId;
            this._items = [];
    }

    public getTickets(jiraJqlQuery: string, jiraDateFilter: string, userEmail: string, nextPageToken: number): Promise<any> {
        const requestHeaders = new Headers();
        requestHeaders.append("Authorization", `Basic ${btoa(this._jiraServiceAccount + ':' + this._jiraAPIToken)}`);
        requestHeaders.append("Content-Type", "application/json");

        const requestGetOptions: IHttpClientOptions = {
            method: "GET",
            headers: requestHeaders,
        };
        let uri = `https://api.atlassian.com/ex/jira/${this._jiraCloudId}/rest/api/3/search?startAt=${nextPageToken}&maxResults=100&fields=issuetype,key,summary,creator,assignee,reporter,status,priority,description,created&expand=renderedFields`;
        let jiraDateFilterString = this.getDateFilterString(jiraDateFilter);

        uri = jiraJqlQuery != null || undefined || "" ? `${uri}&jql=(reporter = '${userEmail}' OR creator = '${userEmail}') AND created >= '${jiraDateFilterString}' AND ${jiraJqlQuery}` : `${uri}&reporter=${userEmail} AND ${jiraDateFilterString}`;
        return this._httpClient  
        .fetch(  
            uri,  
            HttpClient.configurations.v1,
            requestGetOptions
        )  
        .then(response => response.json())
            .then(json => {
                nextPageToken  = nextPageToken < json.total ? nextPageToken + 100 : json.total;
                this._items = this._getRowsFromData(json, nextPageToken); 
                return {'total': json.total ,'issues':this._items} ;
            });
    }

    private getDateFilterString = (dateFilter:string): string => {
        var currentdate = new Date();
        var pastDate;
        switch (dateFilter) {
          case "last6months":
            pastDate = new Date(
              currentdate.getFullYear(),
              currentdate.getMonth() - 6, 
              currentdate.getDate()
            );    
            break;
          case "last1year":
            pastDate = new Date(
              currentdate.getFullYear()-1,
              currentdate.getMonth(), 
              currentdate.getDate()
            );
            break;
          default:
            pastDate = new Date(
              currentdate.getFullYear(),
              currentdate.getMonth() - 6, 
              currentdate.getDate()
            );
            break;
        }
        return pastDate.getFullYear() + "-" + (pastDate.getMonth() + 1) + "-" + pastDate.getDate();  
    }

    private _getRowsFromData(response, nextPageToken) {
        let issues = response.issues.map((issue, i) => {
        
            let tempDescription = issue.renderedFields["description"];
        
            var m,
            urls = [], 
            rexImgs = /<img[^>]+src="?([^"\s]+)"?\s*/gi,
            rexLinks = /<a[^>]+href="?([^"\s]+)"?\s*/gi;
            
            while ( m = rexLinks.exec( tempDescription ) ) {
                //if(m[1].startsWith("/")){
                urls.push( m[1] );
                //}
            }
            while ( m = rexImgs.exec( tempDescription ) ) {
                if(m[1].startsWith("/")){
                urls.push( m[1] );
                }
            }
            if(urls.length != 0){
                urls.forEach(url => {
                if(url.startsWith("/")){
                    tempDescription = tempDescription.replace(url, this._jiraUrl + url);
                }
                if(url.startsWith("http")){
                    tempDescription = tempDescription.replace(url, url + '" target="_blank"');
                }
                });
            }
            let item : IItem =  {
                index: i.toString(),
                key: issue.key,
                summary: issue.fields["summary"],
                assignee: this.getUser(issue, "assignee"),
                reporter: this.getUser(issue, "reporter"),
                creator: this.getUser(issue, "creator"),
                issueType: this.getIssueType(issue),
                status: issue.fields["status"].name,
                priority: issue.fields["priority"].name,
                created: new Date(issue.fields["created"]),
                description: tempDescription
            };
            return item;
        });
        return issues;
    }

    private getUser = (issue:any, userType:string): IUser =>  {
        var user : IUser = {name: "", iconUrl: ""};
        
        if(issue.fields[userType] != null){
            user.name = issue.fields[userType].displayName;
            user.email = issue.fields[userType].emailAddress;
            user.iconUrl = issue.fields[userType]["avatarUrls"]["16x16"];
        }
        else{
            user.name = "Unassigned";
            user.email = "";
            user.iconUrl = "https://cdn-icons-png.flaticon.com/512/2948/2948035.png";
        }
        return user;
    }
    
    private getIssueType = (issue:any): IIssueType =>  {
        var issuetype : IIssueType = {name: "", iconUrl: ""};
        if(issue.fields["issuetype"] != null){
            issuetype.name = issue.fields["issuetype"].name;
            issuetype.iconUrl = issue.fields["issuetype"].iconUrl;
        
            issuetype.iconUrl = `${this._jiraUrl}/` + issuetype.iconUrl.split(`/${this._jiraCloudId}/`)[1];
        
            switch (issuetype.name) {
            case "[System] Service request":
                issuetype.iconUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAHXRFWHRqaXJhLXN5c3RlbS1pbWFnZS10eXBlAGF2YXRhcuQCGmEAAAGfSURBVHjaY2AYBUggeZetRvIum/VA/AmI/1MZg8xcl7zTVh2f5e9pYDE6fp+8G4sjoD7/Tye8DpsDPtHRAR+xOeA/PfHQdUDzieT/ex+t+f/sy/3/v/7+/P/335//b7+/+H/82c7/k89X/E/ZZUsbB+Tu8/h/+sW+/4TAg083/9cei6WuAwoP+P5/8fXRf2LBjz/f/redyqCOA0BBeuvdhf+kgs+/PvwvOuhHuQNmXmr4Ty7Y/3g95Q64+vY02Q749vvL/7TdDpQ54OuvT/8pAbgSJNEO+PPvN0UOaDuZTpkD3gDzOCWg7HAIZQ448nQr2Za/+vaU8kTYeDzp/z8gJAcsvzGJOgXR/sfrSLb89vtLOHMAyQ5YfmMiSZbf+3j1f95+L+oVxdvvLyU632++O/9/+h5H6tYFx57tAFvwBVgmrLo19f+UC5X/t95f/P/k893/z748CK4dZ1ys+58DrLBoUh2vuDn5/5U3J4Flu/9og4SmDhjwRuk6OjpgLaYDgD0WenVMknbaqGHvHe0GOwIUEh9pEewgn+O0fMQCAAwHWUP3k/r8AAAAAElFTkSuQmCC';
                break;
            case "[System] Incident":
                issuetype.iconUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAHXRFWHRqaXJhLXN5c3RlbS1pbWFnZS10eXBlAGF2YXRhcuQCGmEAAAEBSURBVHjaY2AYBUjgkYeNxmMPq/VA/AmI/1MZg8xc98TdWh2f5e9pYDE6fv/EA4sjoD7/Tye8DpsDPtHRAR+xOeA/PfGoA4aOA17Xlfz/++7tf1IBSM/r2mLKHUCO5QhHvKHcATDwqigdp2FPw73+/7p/F2Lpl8//30/qguujmgPwWn7vDtzyl/kpROmjigNwWU4XB+CznOYOIGQ5TR1AjOU0cwCxltPEAaBCBTmrvcxLxmn4q+IM6jmAkoLoD1Av5UUx0Od/3r4h3XKgntf1paO14agDSHLAgDdK19HRAWsxHADqsdCrY/LI00INa+8I1GOBhsRHWgQ7yOc4LR+xAACYCvrMKv1qCAAAAABJRU5ErkJggg==';
                break;
            default:
                break;
            }
        }
        else{
            issuetype.name = "";
            issuetype.iconUrl = "";
        }
        return issuetype;
    }
}