import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import SPListViewService from './listService/listService';
import FileExport from './FileExport/FileExport';
import UrlQueryParameters from './UrlQueryParameters/UrlQueryParameters';

export interface IExcelExportCommandSetProperties {
  enablelistview:any[];
  exportType: string;
}

const LOG_SOURCE: string = 'ExcelExportCommandSet';

let showInView: boolean = false;

export default class ExcelExportCommandSet extends BaseListViewCommandSet<IExcelExportCommandSetProperties> {

  private getUrlViewId = () => {
    const urlParams = location.search;
    let result = '';
    urlParams.split('&').forEach(param => {
      if(param.indexOf("viewid") > -1) {
        result = decodeURI(param.split('=')[1]);
      }
    });
    return result;
  }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ExcelExportCommandSet');
    const view = this.getUrlViewId();
    this.properties.enablelistview.forEach(async listview => {
      const isMember:boolean = await SPListViewService.isGroupMember(listview.spGroupsTitle);
      if(isMember) {
        if(listview.listId === view){
          showInView = true;
          //console.info(showInView);
        }
      }
    });
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    //console.info(showInView);
    if (compareOneCommand && showInView) {
      compareOneCommand.visible = event.selectedRows.length > 0;
    } else {
      compareOneCommand.visible = false;
    }
  }

  private getRows = (event: IListViewCommandSetExecuteEventParameters) => {
    let result:any[] = [];
    event.selectedRows.forEach(row => {
      let rowObj = {};
      row.fields.forEach(field => {
        let value = row.getValue(field);
        if(typeof value === 'object') {
          value = value[0].lookupValue;
        }
        if(field.fieldType === "Number") {
          value = Number(row.getValueByName(field.internalName+'.'));
        }
        rowObj[field.displayName] = value;
      });
      result.push(rowObj);
    });
    return result;
  }

  private export = (event: IListViewCommandSetExecuteEventParameters) => {
    const fields:string[] = event.selectedRows[0].fields.map(field => {
      return field.displayName;
    });
    FileExport[this.properties.exportType](fields, this.getRows(event), this.context.pageContext.list.title);
    console.info(this.updateItems(event));  
  }

  private updateItems = (event: IListViewCommandSetExecuteEventParameters) => {
    SPListViewService.updateListItems(
      String(this.context.pageContext.list.id),
      event.selectedRows.map(row => {
        return Number(row.getValueByName('ID'));
      })
    ).then(result => {
      SPListViewService.createEmailListItems(
        event.selectedRows.map(row => {
          return {
            Title: 'Completed', 
            ListName: this.context.pageContext.list.title, 
            Affiliate: row.getValueByName('Affiliate') ? row.getValueByName('Affiliate') : row.getValueByName('RepresentativeOffice'),
            RelatedItemID: Number(row.getValueByName('ID'))
          };
        })
      ).then(resultEmail => {
        console.info(result, resultEmail);
        location.reload();
      });
    });
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        this.export(event);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
