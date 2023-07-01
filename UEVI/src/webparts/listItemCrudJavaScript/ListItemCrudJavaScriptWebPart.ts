//gavdcodebegin 005
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import { IListItem } from './IListItem'; 

import styles from './ListItemCrudJavaScriptWebPart.module.scss';
import * as strings from 'ListItemCrudJavaScriptWebPartStrings';

export interface IListItemCrudJavaScriptWebPartProps {
  description: string;
  listName: string;
}
//gavdcodeend 005

export default class ListItemCrudJavaScriptWebPart extends BaseClientSideWebPart<IListItemCrudJavaScriptWebPartProps> {

//gavdcodebegin 009
  private CreateItem(): void {
    var myDate = new Date();
    const myBody: string = JSON.stringify({
      'Title': `Item-${myDate.getHours()}${myDate.getMinutes()}${myDate.getSeconds()}`
    });  
    
    var myAbsUrl = this.context.pageContext.web.absoluteUrl;
    var myQuery = '/_api/web/lists/getbytitle(';
    var myListName = this.properties.listName;
    var myOdata = ')/items';
    this.context.spHttpClient.post(
            `${myAbsUrl}${myQuery}'${myListName}'${myOdata}`,  
            SPHttpClient.configurations.v1,  
    {  
      headers: {  
        'Accept': 'application/json;odata=nometadata',  
        'Content-type': 'application/json;odata=nometadata',  
        'odata-version': ''  
      },  
      body: myBody  
    })  
    .then((myResponse: SPHttpClientResponse): Promise<IListItem> => {  
      return myResponse.json();  
    })  
    .then((myItem: IListItem): void => {  
      this.ResponseMessage(`Item '${myItem.Title}' with ID '${myItem.Id}' created`);  
    }, (myError: any): void => {  
      this.ResponseMessage('Error creating Item: ' + myError);  
    });  
  }  
//gavdcodeend 009
    
//gavdcodebegin 011
private ReadItem(): void {  
  this.GetLatestItemId()  
    .then((myItemId: number): Promise<SPHttpClientResponse> => {  
      if (myItemId === -1) {  
        throw new Error('List has no Items');  
      }
        
      var myAbsUrl = this.context.pageContext.web.absoluteUrl;
      var myQuery = '/_api/web/lists/getbytitle(';
      var myListName = this.properties.listName;
      var myOdata1 = ')/items(';
      var myOdata2 = ')?$select=Title,Id';
      return this.context.spHttpClient.get(
              `${myAbsUrl}${myQuery}'${myListName}'${myOdata1}${myItemId}${myOdata2}`,
              SPHttpClient.configurations.v1,
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        });  
    })  
    .then((myResponse: SPHttpClientResponse): Promise<IListItem> => {  
      return myResponse.json();  
    })  
    .then((myItem: IListItem): void => {  
      this.ResponseMessage(`Last Item Title: '${myItem.Title}' - Item ID: '${myItem.Id}'`);  
    }, (myError: any): void => {  
      this.ResponseMessage('Error finding Item: ' + myError);  
    });  
}  
//gavdcodeend 011

//gavdcodebegin 013
private UpdateItem(): void {
  this.GetLatestItemId()  
    .then((myItemId: number): Promise<SPHttpClientResponse> => {  
      if (myItemId === -1) {  
        throw new Error('List has no Items');  
      }

      var myAbsUrl = this.context.pageContext.web.absoluteUrl;
      var myQuery = '/_api/web/lists/getbytitle(';
      var myListName = this.properties.listName;
      var myOdata1 = ')/items(';
      var myOdata2 = ')?$select=Title,Id';
      return this.context.spHttpClient.get(
              `${myAbsUrl}${myQuery}'${myListName}'${myOdata1}${myItemId}${myOdata2}`,
              SPHttpClient.configurations.v1,
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        });  
    })  
    .then((myResponse: SPHttpClientResponse): Promise<IListItem> => {  
      return myResponse.json();  
    })  
    .then((myItem: IListItem): void => {
      const myBody: string = JSON.stringify({  
        'Title': `${myItem.Title}_Updated`  
      });

      var myAbsUrl = this.context.pageContext.web.absoluteUrl;
      var myQuery = '/_api/web/lists/getbytitle(';
      var myListName = this.properties.listName;
      var myOdata1 = ')/items(';
      var myOdata2 = ')';
      this.context.spHttpClient.post(
              `${myAbsUrl}${myQuery}'${myListName}'${myOdata1}${myItem.Id}${myOdata2}`,
              SPHttpClient.configurations.v1,
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'MERGE'  
          },  
          body: myBody  
        })  
        .then((myResponse: SPHttpClientResponse): void => {  
          this.ResponseMessage(`Item ID '${myItem.Id}' updated`);  
        }, (myError: any): void => {  
          this.ResponseMessage(`Error updating Item: ${myError}`);  
        });  
    });  
}  
//gavdcodeend 013

//gavdcodebegin 014
private DeleteItem(): void {
  let etag: string = undefined;  
  this.GetLatestItemId()  
    .then((myItemId: number): Promise<SPHttpClientResponse> => {  
      if (myItemId === -1) {  
        throw new Error('List has no Items');  
      }  
  
      var myAbsUrl = this.context.pageContext.web.absoluteUrl;
      var myQuery = '/_api/web/lists/getbytitle(';
      var myListName = this.properties.listName;
      var myOdata1 = ')/items(';
      var myOdata2 = ')?$select=Id';
      return this.context.spHttpClient.get(
              `${myAbsUrl}${myQuery}'${myListName}'${myOdata1}${myItemId}${myOdata2}`,
              SPHttpClient.configurations.v1,
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        });  
    })  
    .then((myResponse: SPHttpClientResponse): Promise<IListItem> => {  
      etag = myResponse.headers.get('ETag');  
      return myResponse.json();  
    })  
    .then((myItem: IListItem): Promise<SPHttpClientResponse> => {
      var myAbsUrl = this.context.pageContext.web.absoluteUrl;
      var myQuery = '/_api/web/lists/getbytitle(';
      var myListName = this.properties.listName;
      var myOdata1 = ')/items(';
      var myOdata2 = ')';
      return this.context.spHttpClient.post(
              `${myAbsUrl}${myQuery}'${myListName}'${myOdata1}${myItem.Id}${myOdata2}`,
              SPHttpClient.configurations.v1,
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=verbose',  
            'odata-version': '',  
            'IF-MATCH': etag,  
            'X-HTTP-Method': 'DELETE'  
          }  
        });  
    })  
    .then((myResponse: SPHttpClientResponse): void => {  
      this.ResponseMessage(`Last Item deleted`);  
    }, (myError: any): void => {  
      this.ResponseMessage(`Error deleting Item: ${myError}`);  
    });  
} 
//gavdcodeend 014

//gavdcodebegin 012
private GetLatestItemId(): Promise<number> {  
  return new Promise<number>(
      (resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
    
        var myAbsUrl = this.context.pageContext.web.absoluteUrl;
        var myQuery = '/_api/web/lists/getbytitle(';
        var myListName = this.properties.listName;
        var myOdata = ')/items?$orderby=Id desc&$top=1&$select=id';
        this.context.spHttpClient.get(`${myAbsUrl}${myQuery}'${myListName}'${myOdata}`,  
                                        SPHttpClient.configurations.v1,
      {  
        headers: {  
          'Accept': 'application/json;odata=nometadata',  
          'odata-version': ''  
        }  
      })  
      .then((myResponse: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {  
        return myResponse.json();  
      }, (myError: any): void => {  
        reject(myError);  
      })  
      .then((myResponse: { value: { Id: number }[] }): void => {  
        if (myResponse.value.length === 0) {  
          resolve(-1);  
        }  
        else {  
          resolve(myResponse.value[0].Id);  
        }  
      });  
  });  
}  
//gavdcodeend 012

//gavdcodebegin 010
  private ResponseMessage(myResponse: string): void {  
    this.domElement.querySelector('.lblMessage').innerHTML = myResponse;  
  }
//gavdcodeend 010

//gavdcodebegin 008
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listItemCrudJavaScript }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
                <p><button id="btnCreate" class="${ styles.button }">
                <span class="${styles.label}">Create Item</span></button></p>
                <p><button id="btnRead" class="${ styles.button }">
                <span class="${styles.label}">Find Last Item</span></button></p>
                <p><button id="btnUpdate" class="${ styles.button }">
                <span class="${styles.label}">Update Last Item</span></button></p>
                <p><button id="btnDelete" class="${ styles.button }">
                <span class="${styles.label}">Delete Last Item</span></button></p>
                <div class="lblMessage"></div>  
            </div>
          </div>
        </div>
      </div>`;

      document.getElementById("btnCreate").onclick = 
                                          this.CreateItem.bind(this);
      document.getElementById("btnRead").onclick = 
                                          this.ReadItem.bind(this);
      document.getElementById("btnUpdate").onclick = 
                                          this.UpdateItem.bind(this);
      document.getElementById("btnDelete").onclick = 
                                          this.DeleteItem.bind(this);
  }
//gavdcodeend 008

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

//gavdcodebegin 007
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
//gavdcodeend 007
}
