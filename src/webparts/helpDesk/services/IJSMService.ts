export interface IJSMService {
  getTickets: (
    jiraJqlQuery: string,
    jiraDateFilter: string,
    userEmail: string,
    nextPageToken: number
  ) => Promise<any>;

  // TODO: Implement the below methods
  //createTicket: (ticketData: any) => Promise<any>;
  //getTicket: (key: string) => Promise<any>;
  //editTicket: (editTicketData: any) => Promise<any>;
}
