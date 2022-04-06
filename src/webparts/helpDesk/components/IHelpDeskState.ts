import { IColumn } from '@fluentui/react';

export interface IHelpDeskState {
  columns?: IColumn[];
  items?: IItem[];
  isLoading?: boolean;
  nextPageToken?: number;
  selectedItem?: IItem;
  isModalOpen?: boolean;
  total?: number;
  loadingMessage?: string;
  userFilter?: string;
  requestTypeFilter?: string;
  statusFilter?: string;
  dateRangeFilterKey?: number | string;
  dateRangeFilterDate?: Date;
  showClearAllFilter?: boolean;
  searchText?: string;
}
export interface IItem {
  index?:string;
  key?: string;
  summary?: string;
  assignee?: IUser;
  reporter?: IUser;
  creator?: IUser;
  issueType?: IIssueType;
  status?: string;
  priority?: string;
  created?: Date;
  description?: string;
}
export interface IUser {
  name?: string;
  email?: string;
  iconUrl?: string;
}
export interface IIssueType {
  name: string;
  iconUrl: string;
}