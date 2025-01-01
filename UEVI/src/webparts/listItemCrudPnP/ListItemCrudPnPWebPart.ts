//gavdcodebegin 022
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items";
import { IListItem } from './IListItem'; 

import * as strings from 'ListItemCrudPnPWebPartStrings';

export interface IListItemCrudPnPWebPartProps {
  description: string;
  listName: string;
}
//gavdcodeend 022

export default class ListItemCrudPnPWebPart extends 
            BaseClientSideWebPart<IListItemCrudPnPWebPartProps> {

private GetAllItems(): void {
  const sp = spfi().using(SPFx(this.context));
  
  sp.web.lists.getByTitle(this.properties.listName).items()
    .then((allItems: IListItem[]) => {
      let itemsString = '';
      allItems.forEach(oneItem => {
        itemsString += `Title: ${oneItem.Title}, ID: ${oneItem.Id}\n`;
      });
      alert(itemsString);
    })
    .catch((myError: Error) => {
      alert(myError);
    });
}

//gavdcodebegin 026
private CreateItem(): void {
  const sp = spfi().using(SPFx(this.context));
  const myDate = new Date();
  const myTitle = "Item-" + myDate.getHours() + 
                            myDate.getMinutes() + myDate.getSeconds();

  sp.web.lists.getByTitle(this.properties.listName).items.add({
    Title: myTitle
  })
  .then((myResult) => {
    alert(`New item created`);
  })
  .catch((myError) => {
    alert(`Error creating Item: ${myError}`);
  });
}
//gavdcodeend 026

//gavdcodebegin 028
  private ReadLastItem(): void {
    this.GetLatestItemId()
      .then((myItemId: number): void => {
        if (myItemId === -1) {
          throw new Error('List has no Items');
        }
  
          alert(`Last Item ID: ${myItemId}`);
      }
    )
    .catch((myError: Error) => {
      alert(`Error finding Item: ${myError}`);
    });
  }
//gavdcodeend 028

//gavdcodebegin 030
  private UpdateItem(): void {
    const sp = spfi().using(SPFx(this.context));

    this.GetLatestItemId()
    .then((myItemId: number): void => {
      if (myItemId === -1) {
        throw new Error('List has no Items');
      }

      const myDate = new Date();
      const myTitle = "Item-" + myDate.getHours() + 
                                myDate.getMinutes() + myDate.getSeconds();
      const myList = sp.web.lists.getByTitle(this.properties.listName);

      myList.items.getById(myItemId).update({
        Title: `${myTitle}_Updated`
      })
      .then(() => {
        alert("Item updated");
      })
      .catch((myError: Error) => {
        alert(`Error updating Item: ${myError}`);
      });
    })
    .catch((myError: Error) => {
      alert(`Error updating Item: ${myError}`);
    });
  }
//gavdcodeend 030

//gavdcodebegin 031
  private DeleteItem(): void {
    const sp = spfi().using(SPFx(this.context));

    this.GetLatestItemId()
    .then((myItemId: number): void => {
      if (myItemId === -1) {
        throw new Error('List has no Items');
      }

      const myList = sp.web.lists.getByTitle(this.properties.listName);

      myList.items.getById(myItemId).delete()
      .then(() => {
        alert("Item deleted");
      })
      .catch((myError: Error) => {
        alert(`Error deleting Item: ${myError}`);
      });
    })
    .catch((myError: Error) => {
      alert(`Error deleting Item: ${myError}`);
    });
  }
//gavdcodeend 031

//gavdcodebegin 029
  private GetLatestItemId(): Promise<number> {  
    const sp = spfi().using(SPFx(this.context));

    return sp.web.lists.getByTitle(this.properties.listName).items()
      .then((allItems: IListItem[]) => {
        const lastItemId = 
            allItems.length > 0 ? allItems[allItems.length - 1].Id : -1;
        return lastItemId;
      })
      .catch((myError: Error) => {
        console.error(myError);
        return -1;
      });
  }  
//gavdcodeend 029

//gavdcodebegin 027
  // private ResponseMessage(myResponse: string): void {  
  //   alert(myResponse);  
  // }
//gavdcodeend 027

//gavdcodebegin 025
  public render(): void {
    this.domElement.innerHTML = `
      <div>
        <p><button id="btnCreate">
        <span>Create Item</span></button></p>
        <p><button id="btnRead">
        <span>Find All Item</span></button></p>
        <p><button id="btnReadLast">
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

    const btnReadLast: HTMLElement | null = 
              this.domElement.querySelector('#btnReadLast'); 
    if (btnReadLast) { 
      btnReadLast.addEventListener('click', () => 
              this.handleBtnReadLast());
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
      this.GetAllItems();
    }
  private handleBtnReadLast(): void { 
      this.ReadLastItem();
    }
  private handleBtnUpdate(): void { 
    this.UpdateItem();
  }
  private handleBtnDelete(): void { 
    this.DeleteItem();
  }
//gavdcodeend 025

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

//gavdcodebegin 024
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
//gavdcodeend 024
}
