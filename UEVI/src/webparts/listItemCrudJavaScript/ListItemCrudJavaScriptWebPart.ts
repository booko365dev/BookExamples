//gavdcodebegin 005
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';  
import { IListItem } from './IListItem';

import * as strings from 'ListItemCrudJavaScriptWebPartStrings';

export interface IListItemCrudJavaScriptWebPartProps {
  description: string;
  listName: string;
}
//gavdcodeend 005

export default class ListItemCrudJavaScriptWebPart extends 
            BaseClientSideWebPart<IListItemCrudJavaScriptWebPartProps> {

  //gavdcodebegin 009
  private CreateItem(): void {
    const myDate = new Date();
    const myBody: string = JSON.stringify({
      'Title': 
        `Item-${myDate.getHours()}${myDate.getMinutes()}
                                      ${myDate.getSeconds()}`
    });  
    
    const myAbsUrl = this.context.pageContext.web.absoluteUrl;
    const myQuery = '/_api/web/lists/getbytitle(';
    const myListName = this.properties.listName;
    const myOdata = ')/items';
    
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
      alert(`Item '${myItem.Title}' with ID '${myItem.Id}' created`); 
    }, (myError: Error): void => {  
      alert('Error creating Item: ' + myError);  
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
          
        const myAbsUrl = this.context.pageContext.web.absoluteUrl;
        const myQuery = '/_api/web/lists/getbytitle(';
        const myListName = this.properties.listName;
        const myOdata1 = ')/items(';
        const myOdata2 = ')?$select=Title,Id';

        return this.context.spHttpClient.get(
              `${myAbsUrl}${myQuery}'${myListName}'
               ${myOdata1}${myItemId}${myOdata2}`,
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
        alert(`Last Item Title: 
                  '${myItem.Title}' - Item ID: '${myItem.Id}'`);  
      }, (myError: Error): void => {  
        alert('Error finding Item: ' + myError);  
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

        const myAbsUrl = this.context.pageContext.web.absoluteUrl;
        const myQuery = '/_api/web/lists/getbytitle(';
        const myListName = this.properties.listName;
        const myOdata1 = ')/items(';
        const myOdata2 = ')?$select=Title,Id';

        return this.context.spHttpClient.get(
              `${myAbsUrl}${myQuery}'${myListName}'
               ${myOdata1}${myItemId}${myOdata2}`,
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

        const myAbsUrl = this.context.pageContext.web.absoluteUrl;
        const myQuery = '/_api/web/lists/getbytitle(';
        const myListName = this.properties.listName;
        const myOdata1 = ')/items(';
        const myOdata2 = ')';

        this.context.spHttpClient.post(
              `${myAbsUrl}${myQuery}'${myListName}'
               ${myOdata1}${myItem.Id}${myOdata2}`,
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
            alert(`Item ID '${myItem.Id}' updated`);
          }, (myError: Error): void => {
            alert(`Error updating Item: ${myError}`);
          });
      })
      .catch((myError: Error): void => {
        alert(`Error updating Item: ${myError}`);
      });
  }
  //gavdcodeend 013

  //gavdcodebegin 014
  private DeleteItem(): void {
    let myEtag: string;  
    this.GetLatestItemId()  
      .then((myItemId: number): Promise<SPHttpClientResponse> => {  
        if (myItemId === -1) {  
          throw new Error('List has no Items');  
        }  
    
        const myAbsUrl = this.context.pageContext.web.absoluteUrl;
        const myQuery = '/_api/web/lists/getbytitle(';
        const myListName = this.properties.listName;
        const myOdata1 = ')/items(';
        const myOdata2 = ')?$select=Id';

        return this.context.spHttpClient.get(
                `${myAbsUrl}${myQuery}'${myListName}'
                 ${myOdata1}${myItemId}${myOdata2}`,
            SPHttpClient.configurations.v1,
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'odata-version': ''  
            }  
          });  
      })  
      .then((myResponse: SPHttpClientResponse): Promise<IListItem> => {  
        myEtag = myResponse.headers.get('ETag')!;  
        return myResponse.json();  
      })  
      .then((myItem: IListItem): Promise<SPHttpClientResponse> => {
        const myAbsUrl = this.context.pageContext.web.absoluteUrl;
        const myQuery = '/_api/web/lists/getbytitle(';
        const myListName = this.properties.listName;
        const myOdata1 = ')/items(';
        const myOdata2 = ')';
        return this.context.spHttpClient.post(
                  `${myAbsUrl}${myQuery}'${myListName}'
                   ${myOdata1}${myItem.Id}${myOdata2}`,
            SPHttpClient.configurations.v1,
          {  
            headers: {  
              'Accept': 'application/json;odata=nometadata',  
              'Content-type': 'application/json;odata=verbose',  
              'odata-version': '',  
              'IF-MATCH': myEtag,  
              'X-HTTP-Method': 'DELETE'  
            }  
          });  
      })  
      .then((myResponse: SPHttpClientResponse): void => {  
        alert(`Last Item deleted`);  
      }, (myError: Error): void => {  
        alert(`Error deleting Item: ${myError}`);  
      });  
  } 
  //gavdcodeend 014

  //gavdcodebegin 012
  private GetLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve, reject) => { 

      const myAbsUrl = this.context.pageContext.web.absoluteUrl;
      const myQuery = '/_api/web/lists/getbytitle(';
      const myListName = this.properties.listName;
      const myOdata = ')/items?$orderby=Id desc&$top=1&$select=Id';

      this.context.spHttpClient.get(
        `${myAbsUrl}${myQuery}'${myListName}'${myOdata}`,  
        SPHttpClient.configurations.v1,
        {  
          headers: {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
          }  
        })  
        .then((myResponse: SPHttpClientResponse): 
                      Promise<{ value: { Id: number }[] }> => {  
          return myResponse.json();  
        })  
        .then((myResponse: { value: { Id: number }[] }): void => {  
          if (myResponse.value.length === 0) {  
            resolve(-1);  
          } else {  
            resolve(myResponse.value[0].Id);  
          }  
        }, (myError: Error): void => {  
          reject(myError);  
        });  
    });  
  }
  //gavdcodeend 012
  
  //gavdcodebegin 008
  public render(): void {
    this.domElement.innerHTML = `
      <div>
        <p><button id="btnCreate">
        <span>Create Item</span></button></p>
        <p><button id="btnRead">
        <span>Find Last Item</span></button></p>
        <p><button id="btnUpdate">
        <span>Update Last Item</span></button></p>
        <p><button id="btnDelete">
        <span>Delete Last Item</span></button></p>
        <div></div>  
      </div>`;

    const btnCreate: HTMLElement | null = 
                    this.domElement.querySelector('#btnCreate'); 
    if (btnCreate) { 
      btnCreate.addEventListener('click', () => 
                    this.handleBtnCreate());
    }

    const btnRead: HTMLElement | null = 
                    this.domElement.querySelector('#btnRead'); 
    if (btnRead) { 
      btnRead.addEventListener('click', () => 
                    this.handleBtnRead());
    }

    const btnUpdate: HTMLElement | null = 
                    this.domElement.querySelector('#btnUpdate'); 
    if (btnUpdate) { 
      btnUpdate.addEventListener('click', () => 
                    this.handleBtnUpdate());
    }

    const btnDelete: HTMLElement | null = 
                    this.domElement.querySelector('#btnDelete'); 
    if (btnDelete) { 
      btnDelete.addEventListener('click', () => 
                    this.handleBtnDelete());
    }
  }
  private handleBtnCreate(): void { 
    this.CreateItem();
  }
  private handleBtnRead(): void { 
    this.ReadItem();
  }
  private handleBtnUpdate(): void { 
    this.UpdateItem();
  }
  private handleBtnDelete(): void { 
    this.DeleteItem();
  }
  //gavdcodeend 008

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //gavdcodebegin 007
  protected getPropertyPaneConfiguration(): 
                  IPropertyPaneConfiguration {
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
