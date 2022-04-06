import * as React from 'react';
import styles from './HelpDesk.module.scss';
import { IHelpDeskProps } from './IHelpDeskProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IHelpDeskState, IIssueType, IItem, IUser } from './IHelpDeskState';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { getTheme, ThemeProvider, Stack, StackItem, TextField, IColumn, mergeStyleSets, TooltipHost, IIconProps, DetailsList, SelectionMode, DetailsListLayoutMode, IconButton, Spinner, Selection, MarqueeSelection, IDetailsRowProps, DetailsRow, IDetailsHeaderProps, IRenderFunction, Sticky, ConstrainMode, IDetailsColumnRenderTooltipProps, IDetailsListStyles, ImageFit, Modal, IButtonStyles, ActionButton, FontWeights, Label, Dropdown, IDropdownOption, PrimaryButton, ScrollablePane, StickyPositionType, ScrollbarVisibility, Button, DefaultButton, SearchBox, ISearchBoxStyles, IDropdownStyles  } from '@fluentui/react';
import { findLastIndex } from 'lodash';
//import { WebPartTitle } from '@pnp/spfx-controls-react';
import 'office-ui-fabric-react/dist/css/fabric.css';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { JSMService } from '../services/JSMService';
import { IJSMService } from '../services/IJSMService';
const theme = getTheme();

const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '800px'
  },
  header: [
    theme.fonts.large,
    {
      flex: '1 1 auto',
      color: theme.palette.neutralDark,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      padding: '12px 12px 14px 24px',
      backgroundColor: theme.palette.neutralLighter
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'auto',
    maxHeight: '400px',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
  footer: {

  }
});

const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};  
const gridStyles: Partial<IDetailsListStyles> = {
  root: {
    //overflowX: 'scroll',
    selectors: {
      '& [role=grid]': {
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'start',
        height:  '63vh',
        marginLeft: '20px'
      },
    },
  },
  headerWrapper: {
    flex: '0 0 auto',
  },
  contentWrapper: {
    flex: '1 1 auto',
    overflowY: 'auto',
    overflowX: 'hidden',
  },
};


const ddStyles: Partial<IDropdownStyles> = {
  root: {
    selectors: {
      '& [role=combobox]': {
        border:  '1px solid black',
      },
      '&::after': {
        border: '1px solid black',
      },
    },
  },
};


const classNames = mergeStyleSets({
  header: {
    margin: 0,
  },
  row: {
    flex: '0 0 auto',
  },
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px',
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden',
      },
    },
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '24px',
    maxWidth: '24px',
    marginTop: '-3px',
    marginRight: '5px'
  },
  columnTextColor:{
    color:'#323130'
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px',
  },
  selectionDetails: {
    marginBottom: '20px',
  },
  descriptionContainer:{
    display: 'flex',
    alignItems: 'center',
    alignContent: 'center',
    justifyContent: 'center',
    backgroundColor: theme.palette.neutralLighter,// '#f5f5f5',
    //marginBottom: '25px',
    borderRadius: '6px',
    paddingTop: '5px',
    paddingBottom: '5px'
  },
  descriptionThumbnail: {
    marginRight: '15px',
    marginLeft: '15px'
  },
  descriptionHeader: {
    marginBotton: '5px',
    fontWeight: '600',
    fontSize: '1.3em'
  },
  descriptionContent: {
    marginRight: '15px',
  },
  myTicketsHeader:{
    fontWeight: '600',
    fontSize: '1.3em',
    margin: '20px 0 0 0'
  },
  modalContainer:{
  },
  itemKeyButton:{
    height: '20px',
    color: theme.palette.themePrimary,
    margin: '-10px'
  },
  itemSummaryButton:{
    height: '20px',
    color: theme.palette.themePrimary,
    margin: '-10px'
  },
  noItems:{
    textAlign: 'center',
    paddingTop: '100px'
  },
  ddInput:{
    marginRight: '5px',
    marginLeft: '5px',
    selectors: {
      '& [role=combobox]:focus::after': {
        border: '1px solid black'
      },
      '& [role=combobox]:active': {
        '.ms-Dropdown-title':{
          borderColor: 'black'
        }
      }
    }
  },
  clearFiltersButton:{
    marginLeft: '15px',
    width: '16px',
    height: '16px'
  },
  /* viewTicketModalContainer:{
    width:'800px'
  }, */
  mt20:{
    marginTop: '20px !important'
  },
  mb20:{
    marginBottom: '20px !important'
  },
  viewTicketLabels:{
    color: `${theme.palette.black}`
  },
  viewTicketModal:{
    maxHeight: 'inherit'
  }

});



const searchControlStyles: Partial<ISearchBoxStyles> = {
  root: {
    margin: '0px 15px 18px 25px',
    width: '200px',
    selectors: {
      '&::after': {
        border: '1px solid black',
      }
    }
  },
};

const refreshButtonStyles = {
  root: {
    margin: '20px 0px 0px 10px',
    verticalAlign: 'middle',
    color: 'black'
  }
};

const clearFiltersButtonStyles = {
  root: {
    //margin: '5px 0px 0px 10px',
    verticalAlign: 'middle',
    selectors:{
    img: {
      height: '16px',
      width: '16px'
      }
    }
  }
};

const addIcon: IIconProps = { iconName: 'Add' };
const clearFiltersIcon: IIconProps = { imageProps:{src:`${require('../assets/clear-filter-icon.png')}`, height:'16', width: '16'} };
const searchIcon: IIconProps = { iconName: 'Search' };
const cancelIcon: IIconProps = { iconName: 'Cancel', style:{ color: 'black'} };

export default class HelpDesk extends React.Component<IHelpDeskProps, IHelpDeskState> {
  private _allItems: IItem[];
  private _selection;
  private jsmService: IJSMService;

  constructor(props: IHelpDeskProps, state: IHelpDeskState) {  
    super(props);
    this.jsmService = new JSMService(
      this.props.httpclient,
      this.context, 
      this.props.jiraServiceAccount,
      this.props.jiraAPIToken,
      this.props.jiraUrl,
      this.props.jiraCloudId);

    const columns: IColumn[] = [
      {
        key: 'issueType',
        name: 'Type',
        ariaLabel: 'Press to sort on Issue Type',
        fieldName: 'issueType',
        minWidth: 200,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: IItem) => (
          <TooltipHost content={`${item.issueType.name}`}>
            <img src={item.issueType.iconUrl} className={classNames.fileIconImg} alt={`${item.issueType.name}`} /><span className={classNames.columnTextColor}>{item.issueType.name}</span>
          </TooltipHost>
        ),
      },
      {
        key: 'key',
        name: 'Key',
        fieldName: 'key',
        ariaLabel: 'Press to sort on Key',
        onColumnClick: this._onColumnClick,
        minWidth: 50,
        isResizable: true,
        isRowHeader: true,
        data: 'string',
        isPadded: true,
        onRender: (item: IItem) => (
          // <TooltipHost content={`${item.key}`}>
            <ActionButton onClick={ () => this._onItemClick(item)} className={classNames.itemKeyButton} >
              {item.key}
            </ActionButton>
          // </TooltipHost>
        ),
      },
      {
        key: 'summary',
        name: 'Summary',
        fieldName: 'summary',   
        /* commented the sorting functionality for summary field */
        
        //ariaLabel: 'Press to sort on Summary',
        //onColumnClick: this._onColumnClick,
        onRender: (item: IItem) => (
          <TooltipHost content={`${item.summary}`}>
            <ActionButton onClick={ () => this._onItemClick(item)} className={classNames.itemSummaryButton} >
              {item.summary}
            </ActionButton>
          </TooltipHost>
        ),
        minWidth: 325,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'creator',
        name: 'Creator',
        fieldName: 'creator',
        ariaLabel: 'Press to sort on Creator',
        onColumnClick: this._onColumnClick,
        minWidth: 175,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        onRender: (item: IItem) =>  (
          <TooltipHost content={`${ item.creator && item.creator.name}`}>
            <img src={item.creator && item.creator.iconUrl} className={classNames.fileIconImg} alt={`${item.creator && item.creator.name}`} /><span className={classNames.columnTextColor}>{item.creator.name}</span>
          </TooltipHost>
        ),
        isPadded: true,
      },
      {
        key: 'reporter',
        name: 'Reporter',
        fieldName: 'reporter',
        ariaLabel: 'Press to sort on Reporter',
        onColumnClick: this._onColumnClick,
        minWidth: 175,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        onRender: (item: IItem) => (
          <TooltipHost content={`${item.reporter.name}`}>
            <img src={item.reporter.iconUrl} className={classNames.fileIconImg} alt={`${item.reporter.name}`} /><span className={classNames.columnTextColor}>{item.reporter.name}</span>
          </TooltipHost>
        ),
        isPadded: true,
      },
      {
        key: 'status',
        name: 'Status',
        fieldName: 'status',
        ariaLabel: 'Press to sort on Status',
        onColumnClick: this._onColumnClick,
        onRender: (item: IItem) => (
          <TooltipHost content={`${item.status}`}>
            <span className={classNames.columnTextColor}>{item.status}</span>
          </TooltipHost>
        ),
        minWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'created',
        name: 'Created',
        fieldName: 'created',
        ariaLabel: 'Press to sort on Created',
        onColumnClick: this._onColumnClick,
        onRender: (item: IItem) => (
          <TooltipHost content={`${item.created.toLocaleDateString()}`}>
            <span className={classNames.columnTextColor}>{item.created.toLocaleDateString()}</span>
          </TooltipHost>
        ),
        minWidth: 75,
        isRowHeader: true,
        isResizable: true,
        data: 'date',
        isPadded: true,
      },
    ];

    this.state = {  
      columns: columns,
      items: null,
      isLoading: false,
      nextPageToken: 0,
      selectedItem: {},
      showClearAllFilter: false,
      searchText:''
    };  
    this._selection = new Selection();

    this._onColumnClick = this._onColumnClick.bind(this);
    this._onItemClick = this._onItemClick.bind(this);
    this._onReloadClick = this._onReloadClick.bind(this);
    this._onRenderMissingItem = this._onRenderMissingItem.bind(this);
    this._onHideModal = this._onHideModal.bind(this);
    this._onResetFilters = this._onResetFilters.bind(this);
    this._onCreatororReporterClear = this._onCreatororReporterClear.bind(this);
    this._onRequestTypeClear = this._onRequestTypeClear.bind(this);
    this._onStatusClear = this._onStatusClear.bind(this);
    this._onDateRangeClear = this._onDateRangeClear.bind(this);
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems,
    });
  }

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    //if(key == 'reporter' || 'creator' || 'issueType'  ){
      if((key == 'reporter') || (key == 'creator') || (key == 'issueType')){
      return items.filter(i => i).slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key]['name'] < b[key]['name'] : a[key]['name'] > b[key]['name']) ? 1 : -1));
    }
    else{
      return items.filter(i => i).slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    }
  }

  private _onReloadClick() {
    this.setState({ items: null, isLoading:true, nextPageToken: 0, userFilter: null, requestTypeFilter:null, statusFilter:null, dateRangeFilterDate:null, dateRangeFilterKey:null, searchText:'' });
    setTimeout(() => {
      this._onLoadNextPage();
    }, 1); 
    //this._onLoadNextPage();
  }

  private _onRenderMissingItem() :(index?: number, rowProps?: IDetailsRowProps) => React.ReactNode {
    console.log('onRenderMissingItem getting called');
    let { isLoading } = this.state;
    if (!isLoading) {
      this.setState({ isLoading: true });
    }
    setTimeout(() => {
      this._onLoadNextPage();
    }, 1);
    //this._onLoadNextPage();  
    return;
  }

  private _onLoadNextPage() {
    let {jiraServiceAccount, jiraAPIToken, jiraCloudId, jiraJqlQuery, jiraDateFilter, userEmail, jiraUserFilter } = this.props;
    let { items, nextPageToken, total, isLoading, loadingMessage } = this.state;
    this.setState({ isLoading});
    this.jsmService.getTickets(jiraJqlQuery,jiraDateFilter, userEmail, nextPageToken)
    .then(response => {
      total = response.total;
      nextPageToken  = nextPageToken < total ? nextPageToken + 100 : total;
      let issues = [];
      issues = response.issues; 
      
      // adding the issues to the collection if there are issues in the next page
      if (items && nextPageToken) {
        issues = items.slice(0, items.length - 1).concat(issues);
      }
      // adding null to the end of the issues array, so that  DetailsList - onRenderMissingItem  method will be triggered
      if (response.total > nextPageToken ) {
          issues.push(null);
      }
      if(issues[issues.length -1] != null){
        this._allItems = issues;
        this._selection.setItems(issues);
      }
      this.setState({ 
        items: issues,
        nextPageToken,
        isLoading: false,
        total
      });
    });
  }

  public componentDidMount(){  
    this._onReloadClick();
  }
  
  private _getKey(item: any, index?: number): string {
    return item.index;
  }

  private _onItemClick(item: any): void {
    let {isModalOpen, selectedItem} = this.state;
    if(!isModalOpen){
      isModalOpen = true;
      selectedItem = item;
    }
    this.setState({isModalOpen, selectedItem});
  }

  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter} = this.state;
    userFilter = requestTypeFilter = statusFilter = dateRangeFilterKey = dateRangeFilterDate  = null;
    showClearAllFilter = false;
    this.setState({
      userFilter,requestTypeFilter,statusFilter,dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter,
      searchText: text,
      items: text ? this._allItems
      .filter(j => j).filter(i =>
        (i.issueType.name.toLowerCase().indexOf(text.toLowerCase()) > -1 ||
        i.key.toLowerCase().indexOf(text.toLowerCase()) > -1 ||
        i.summary.toLowerCase().indexOf(text.toLowerCase()) > -1 ||
        i.creator.name.toLowerCase().indexOf(text.toLowerCase()) > -1 ||
        i.reporter.name.toLowerCase().indexOf(text.toLowerCase()) > -1 ||
        i.status.toLowerCase().indexOf(text.toLowerCase()) > -1
        )
      ) 
      : this._allItems,
    });
  }
  private _onResetFilters(){
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter} = this.state;
    userFilter = requestTypeFilter = statusFilter = dateRangeFilterKey = dateRangeFilterDate  = null;
    items = this._allItems;
    showClearAllFilter = false;
    this.setState({items,userFilter,requestTypeFilter,statusFilter,dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter});
  }
  private _onCreatororReporterChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void =>{
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter, searchText} = this.state;
    let {userEmail} = this.props;
    userFilter = option.text;
    searchText = '';
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          (userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true) : true)
      )
      );
      this.setState({searchText, userFilter, showClearAllFilter, items});
  }

  private _onCreatororReporterClear()  {
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter} = this.state;
    let {userEmail} = this.props;
    userFilter = null;
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true)  &&
        (userFilter ? 
          (userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true) : true) 
      )
      );
      this.setState({userFilter, showClearAllFilter, items});
  }

  private _onRequestTypeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void =>{
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate,showClearAllFilter, searchText} = this.state;
    let {userEmail} = this.props;
    requestTypeFilter = option.text;
    searchText = '';
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          (userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true) : true)
      )
      );
      this.setState({searchText, requestTypeFilter, showClearAllFilter, items});
   }

   private _onRequestTypeClear(){
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate,showClearAllFilter} = this.state;
    let {userEmail} = this.props;
    requestTypeFilter = null;
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          (userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true) : true)
      )
      );
      this.setState({requestTypeFilter, showClearAllFilter, items});
   }
  private _onStatusChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void =>{
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter, searchText} = this.state;
    let {userEmail} = this.props;
    statusFilter = option.text;
    searchText = '';
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true : true)
      )
    );
    this.setState({searchText, statusFilter, showClearAllFilter, items});
  }

  private _onStatusClear(){
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter} = this.state;
    let {userEmail} = this.props;
    statusFilter = null;
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true : true)
      )
    );
    this.setState({statusFilter, showClearAllFilter, items});
  }
  private _onDateRangeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void =>{
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate,showClearAllFilter, searchText} = this.state;
    let {userEmail} = this.props;
    searchText = '';
    var pastDate;
    var currentDate = new Date();
    switch (option.key) {
      case "last6months":
        pastDate = new Date(
          currentDate.getFullYear(),
          currentDate.getMonth() - 6, 
          currentDate.getDate()
        );    
        break;
      case "last3months":
        pastDate = new Date(
          currentDate.getFullYear(),
          currentDate.getMonth() - 3 , 
          currentDate.getDate()
        );
        break;
      case "last1month":
        pastDate = new Date(
          currentDate.getFullYear(),
          currentDate.getMonth() - 1, 
          currentDate.getDate()
        );
        break;  
      default:
        break;
    }
    dateRangeFilterDate = pastDate;
    dateRangeFilterKey = option.key;
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          (userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true) : true)
      )
    );
    this.setState({searchText, dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter, items});
  }

  private _onDateRangeClear() {
    let {items, userFilter, requestTypeFilter,statusFilter, dateRangeFilterKey, dateRangeFilterDate,showClearAllFilter} = this.state;
    let {userEmail} = this.props;
    dateRangeFilterDate = null;
    dateRangeFilterKey = null;
    let filtersArr = [userFilter, requestTypeFilter, statusFilter, dateRangeFilterKey];
    if(filtersArr.filter(x => x !== null).length > 0){
      showClearAllFilter = true;
    }
    else{
      showClearAllFilter = false;
    }
    items = this._allItems.filter(j => j).filter(i =>
      (
        (requestTypeFilter ? i.issueType.name.toLowerCase().indexOf(requestTypeFilter.toLowerCase()) > -1 : true) &&
        (statusFilter ? 
          (statusFilter == "Open" ? i.status.toLowerCase() != "done": 
          statusFilter == "Closed" ? i.status.toLowerCase() == "done": true) : true) &&
        (dateRangeFilterDate ? i.created >= dateRangeFilterDate: true) &&
        (userFilter ? 
          (userFilter == "Created by me" ? i.creator.email.toLowerCase() == userEmail.toLowerCase() :
          userFilter == "Reported by me" ? i.reporter.email.toLowerCase() == userEmail.toLowerCase() : true) : true)
      )
    );
    this.setState({dateRangeFilterKey, dateRangeFilterDate, showClearAllFilter, items});
  }

  private _onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props) {
      return null;
    }
    const onRenderColumnHeaderTooltip: IRenderFunction<IDetailsColumnRenderTooltipProps> = 
       tooltipHostProps => (
          <TooltipHost {...tooltipHostProps} />
        );
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender!({
           ...props,
           onRenderColumnHeaderTooltip,
        })}
      </Sticky>
    );
  }

  private _onHideModal(){
    let {isModalOpen}= this.state;
    if(isModalOpen){
      isModalOpen = false;
    } 
    this.setState({isModalOpen});
  }

  public render(): React.ReactElement<IHelpDeskProps> {
    {SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css');}
    const {
      jiraCloudId,
      title,
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;
    let { items, isLoading, isModalOpen, selectedItem, nextPageToken, total, loadingMessage, showClearAllFilter, searchText, userFilter, requestTypeFilter, statusFilter, dateRangeFilterDate, dateRangeFilterKey } = this.state;

    return (
      <ThemeProvider>
          {/* <Stack horizontal horizontalAlign='start' className={classNames.descriptionContainer} >
            <StackItem>
              <img src={require('../assets/help-desk-ticket.png')} height={64} width={64} className={classNames.descriptionThumbnail} ></img>
            </StackItem>
            <StackItem grow>
              <h3 className={classNames.descriptionHeader}>{title}</h3>
              <p className={classNames.descriptionContent}>{description}</p>
            </StackItem>
          </Stack> */}
          <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
          <div>
          <Stack horizontal horizontalAlign='start' className={classNames.descriptionContainer}>
            <StackItem>
              <img src={require('../assets/help-desk-ticket.png')} height={64} width={64} className={classNames.descriptionThumbnail} ></img>
            </StackItem>
              <StackItem align='start'>
                <h3 className={classNames.myTicketsHeader}>My Tickets {items && (<span style={{color: '#FA4C06'}}>({items.length})</span>)} </h3>
              </StackItem>
              <StackItem grow align='start'>
                <TooltipHost
                  content="Reload tickets"
                >
                <IconButton 
                    styles={ refreshButtonStyles }
                    iconProps={ { iconName:'Refresh' }}
                    onClick={this._onReloadClick} 
                  />
                </TooltipHost>
              </StackItem>
              {/* <IconButton 
                styles={ refreshButtonStyles }
                iconProps={ { iconName:'Filter' }}
              /> */}
              <Dropdown
                id="ddUser"
                placeholder='Creator/Reporter'
                selectedKey={userFilter}
                //disabled={this.props.jiraUserFilter == "otheruser"}
                options={[
                  { key: 'Created by me', text: 'Created by me',title: 'Created by me' },
                  { key: 'Reported by me', text: 'Reported by me', title: 'Reported by me' },
                ]}
                onChange={this._onCreatororReporterChange}
                //onRenderTitle = {options => <span><b>User:</b>{options[0].text}</span>}
                className={classNames.ddInput}
                //styles={ddStyles}
              />
              {userFilter &&
              <TooltipHost
                content="Clear Creator/Reporter"
              >
                <IconButton iconProps={cancelIcon} onClick={this._onCreatororReporterClear}/>
              </TooltipHost>
              }
              <Dropdown
                id="ddRequestType"
                placeholder='Request Type'
                selectedKey={requestTypeFilter}
                options={[
                  { key: '[System] Service request', text: '[System] Service request', title: '[System] Service request' },
                  { key: '[System] Incident', text: '[System] Incident', title: '[System] Incident' },
                ]}
                onChange={this._onRequestTypeChange}
                onRenderTitle = {options => <span><b>Type: </b>{options[0].text}</span>}
                className={classNames.ddInput}
              />
              {requestTypeFilter &&
              <TooltipHost
                content="Clear Request Type"
              >
                <IconButton iconProps={cancelIcon} onClick={this._onRequestTypeClear}/>
              </TooltipHost>
              }
              <Dropdown
                id="ddStatus"
                placeholder="Status"
                selectedKey={statusFilter}
                options={[
                  { key: 'Open', text: 'Open', title: 'Open' },
                  { key: 'Closed', text: 'Closed', title: 'Closed' },
                ]}
                onChange={this._onStatusChange}
                onRenderTitle = {options => <span><b>Status: </b>{options[0].text}</span>}
                className={classNames.ddInput}
              />
              {statusFilter &&
              <TooltipHost
                content="Clear Status"
              >
                <IconButton iconProps={cancelIcon} onClick={this._onStatusClear}/>
              </TooltipHost>
              }
              <Dropdown
                id="ddDateRange"
                placeholder="Date Range"
                selectedKey={dateRangeFilterKey}
                options={[
                  { key: 'last6months', text: 'Last 6 Months', title: 'Last 6 Months' },
                  { key: 'last3months', text: 'Last 3 Months', title: 'Last 3 Months' },
                  { key: 'last1month', text: 'Last 1 Month', title: 'Last 1 Month' },
                ]}
                onChange={this._onDateRangeChange}
                onRenderTitle = {options => <span><b>Date Range: </b>{options[0].text}</span>}
                className={classNames.ddInput}
              />
              {dateRangeFilterKey &&
              <TooltipHost
                content="Clear Date Range"
              >
                <IconButton iconProps={cancelIcon} onClick={this._onDateRangeClear}/>
              </TooltipHost>
              }
              {
                showClearAllFilter &&
              <TooltipHost
                content="Clear all filters"
              >
                <IconButton styles={clearFiltersButtonStyles} iconProps={clearFiltersIcon} onClick={this._onResetFilters}/>
              </TooltipHost>
              }
              <StackItem align='end'>
                <SearchBox
                  styles={searchControlStyles}
                  placeholder="Search"
                  defaultValue={searchText}
                  value={searchText}
                  onEscape={ev => {
                    console.log('Custom onEscape Called');
                  }}
                  onClear={ev => {
                    console.log('Custom onClear Called');
                  }}
                  //onChange={(_, newValue) => console.log('SearchBox onChange fired: ' + newValue)}
                  onChange={this._onChangeText}
                  onSearch={newValue => console.log('SearchBox onSearch fired: ' + newValue)}
                  disableAnimation={true}
                />
                {/* <TextField iconProps={searchIcon} placeholder='Search' onChange={this._onChangeText} styles={searchControlStyles} />   */}
              </StackItem>
            </Stack>
            </div>
          </Sticky>
          {/* { items != null && (
            <div>
              <span style={{float:'right', marginRight: '20px'}}>Tickets count: {items.length}</span>
            </div>
          )} */}
          { isLoading && (
            <Spinner style={{'paddingTop':'20px'}} label={'Loading...' /* loadingMessage */} labelPosition='right' />
          )}
          {(items && items.length > 0) && (
                <MarqueeSelection selection={ this._selection }>
                    <DetailsList
                    items={this.state.items}
                    columns={this.state.columns}
                    selectionMode={SelectionMode.none}
                    selection={ this._selection }
                    //getKey={this._getKey}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    //constrainMode={ConstrainMode.unconstrained}
                    //onRenderDetailsHeader={this._onRenderDetailsHeader}
                    selectionPreservedOnEmptyClick={false}
                    isHeaderVisible={true}
                    styles={gridStyles}
                    onRenderMissingItem={ this._onRenderMissingItem }
                    onRenderRow={ (props, defaultRender) => <div onClick={ () => console.log('clicking: ' + this.props.jiraUrl)}>{defaultRender(props)}</div> }
                    onShouldVirtualize={() => false}
                  />
                </MarqueeSelection>
          )}
          { (items != null && items.length == 0 ) && (
            <div className={classNames.noItems}>
              <img style={{'height':'90px', 'width':'auto'}} src={require('../assets/no-tickets.png')} ></img>
              <br/><span>There are no tickets to display.</span>
            </div>
          )}
        <Modal
          isOpen={isModalOpen}
          onDismiss={()=>this._onHideModal}
          isBlocking={false}
          className={classNames.viewTicketModal}
        >
          <div className={contentStyles.container}>
            <div className={contentStyles.header}>
              <span>{selectedItem && selectedItem.key + " : " + selectedItem.summary} </span>
              <IconButton
                styles={iconButtonStyles}
                iconProps={cancelIcon}
                ariaLabel="Close popup modal"
                onClick={this._onHideModal}
              />
            </div>
            <div className={contentStyles.body}>
              <div className='ms-Grid' dir='ltr'>
                <div className={`ms-Grid-row ${classNames.mb20} ${classNames.mt20}`}>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Request Type</Label>  
                    {selectedItem && selectedItem.issueType && selectedItem.issueType.name}
                  </div>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Priority</Label>  
                    {selectedItem && selectedItem.priority}
                  </div>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Created Date</Label>  
                    {selectedItem && selectedItem.created && selectedItem.created.toLocaleDateString()}
                  </div>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Status</Label>  
                    {selectedItem && selectedItem.status}
                  </div>
                </div>
                <div className={`ms-Grid-row ${classNames.mb20}`}>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Reporter</Label>
                    {selectedItem && selectedItem.reporter && selectedItem.reporter.name}
                  </div>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Creator</Label>  
                    {selectedItem && selectedItem.creator && selectedItem.creator.name}
                  </div>
                  <div className='ms-Grid-col ms-md3'>
                    <Label className={classNames.viewTicketLabels}>Assignee</Label>  
                    {selectedItem && selectedItem.assignee && selectedItem.assignee.name}
                  </div>
                </div>
                <div className={`ms-Grid-row ${classNames.mb20}`}>
                  <div className='ms-Grid-col ms-md12'>
                    <Label className={classNames.viewTicketLabels}>Summary</Label>  
                    {selectedItem && selectedItem.summary}
                  </div>
                </div>
                <div className={`ms-Grid-row ${classNames.mb20}`}>
                  <div className='ms-Grid-col ms-md12'>
                    <Label className={classNames.viewTicketLabels}>Description</Label>  
                    {selectedItem && (<span dangerouslySetInnerHTML={{__html: selectedItem.description}}></span>)}
                  </div>
                </div>
              </div>
            </div>
            <div className={contentStyles.footer}>
              <div className='ms-Grid'>
                <div className={`ms-Grid-row ${classNames.mb20} ${classNames.mt20}`}>
                  <div className='ms-Grid-col' style={{float:'right', marginRight: '20px'}}>
                    <DefaultButton text='Cancel' onClick={this._onHideModal}></DefaultButton>              
                  </div>
                </div>
              </div>
            </div>
          </div>
        </Modal>
      </ThemeProvider>    
    );
  }
}
