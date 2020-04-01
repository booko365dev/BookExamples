//gavdcodebegin 22
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { sp, ItemAddResult, ItemUpdateResult } from "@pnp/sp";
import { IListItem } from './IListItem'; 

import styles from './ListItemCrudPnPWebPart.module.scss';
import * as strings from 'ListItemCrudPnPWebPartStrings';

export interface IListItemCrudPnPWebPartProps {
  description: string;
  listName: string;
}
//gavdcodeend 22

export default class ListItemCrudPnPWebPart extends BaseClientSideWebPart<IListItemCrudPnPWebPartProps> {

//gavdcodebegin 26
  private CreateItem(): void {
    var myDate = new Date();
    sp.web.lists.getByTitle(this.properties.listName).items.add({
      Title: "Item-" + myDate.getHours() + myDate.getMinutes() + myDate.getSeconds()})
      .then((myResult: ItemAddResult): void => {
      const myItem: IListItem = myResult.data as IListItem;
      this.ResponseMessage(`Item '${myItem.Title}' with ID '${myItem.Id}' created`);
    }, (myError: any): void => {  
      this.ResponseMessage('Error creating Item: ' + myError);
     });
  }
//gavdcodeend 26

//gavdcodebegin 28
  private ReadItem(): void {
    this.GetLatestItemId()  
      .then((myItemId: number): Promise<IListItem> => {  
        if (myItemId === -1) {  
          throw new Error('List has no Items');  
        }  
  
        return sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(myItemId).select('Title', 'Id').get();  
      })  
      .then((myItem: IListItem): void => {  
        this.ResponseMessage(`Last Item Title: '${myItem.Title}' - Item ID: '${myItem.Id}`);  
      }, (myError: any): void => {
        this.ResponseMessage('Error finding Item: ' + myError);  
      });  
  }
//gavdcodeend 28

//gavdcodebegin 30
  private UpdateItem(): void {
    let etag: string = undefined;  
  
    this.GetLatestItemId()  
      .then((myItemId: number): Promise<IListItem> => {  
        if (myItemId === -1) {  
          throw new Error('List has no Items');  
        }  
  
        return sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(myItemId).get(undefined, {  
            headers: {  
              'Accept': 'application/json;odata=minimalmetadata'  
            }  
          });  
      })  
      .then((myItem: IListItem): Promise<IListItem> => {  
        etag = myItem["odata.etag"];  
        return Promise.resolve((myItem as any) as IListItem);  
      })  
      .then((myItem: IListItem): Promise<ItemUpdateResult> => {  
        return sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(myItem.Id).update({  
            'Title': `${myItem.Title}_Updated`  
          }, etag);  
      })  
      .then((myResult: ItemUpdateResult): void => {
        this.ResponseMessage(`Item updated`);  
      }, (myError: any): void => {  
        this.ResponseMessage('Error updating Item: ' + myError);  
      });
  }
//gavdcodeend 30

//gavdcodebegin 31
  private DeleteItem(): void {
    let etag: string = undefined;

    this.GetLatestItemId()  
      .then((myItemId: number): Promise<IListItem> => {  
        if (myItemId === -1) {  
          throw new Error('List has no Items');  
        }  
    
        return sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(myItemId).select('Id').get(undefined, {  
            headers: {  
              'Accept': 'application/json;odata=minimalmetadata'  
            }  
          });  
      })  
      .then((myItem: IListItem): Promise<IListItem> => {  
        etag = myItem["odata.etag"];  
        return Promise.resolve((myItem as any) as IListItem);  
      })  
      .then((myItem: IListItem): Promise<void> => {  
        return sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(myItem.Id).delete(etag);  
      })  
      .then((): void => {  
        this.ResponseMessage(`Last Item deleted`);  
      }, (myError: any): void => {  
        this.ResponseMessage(`Error deleting item: ${myError}`);  
      });
  }
//gavdcodeend 31

//gavdcodebegin 29
  private GetLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve: (itemId: number) => 
                        void, reject: (error: any) => void): void => {  
      sp.web.lists.getByTitle(this.properties.listName)  
        .items.orderBy('Id', false).top(1).select('Id').get()  
        .then((items: { Id: number }[]): void => {  
          if (items.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(items[0].Id);  
          }  
        }, (error: any): void => {  
          reject(error);  
        });  
    });  
  }  
//gavdcodeend 29

//gavdcodebegin 27
  private ResponseMessage(myResponse: string): void {  
    this.domElement.querySelector('.lblMessage').innerHTML = myResponse;  
  }
//gavdcodeend 27

//gavdcodebegin 25
public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listItemCrudPnP }">
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
//gavdcodeend 25

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

//gavdcodebegin 24
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
//gavdcodeend 24
}
