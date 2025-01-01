//gavdcodebegin 040
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { spfi, SPFx } from "@pnp/sp";
import { SearchResults } from "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/profiles";
import "@pnp/sp/search";

import * as strings from 'OtherPnPjsWebPartStrings';
import { IListInfo } from '@pnp/sp/lists';
import { IFileInfo } from '@pnp/sp/files';

export interface IOtherPnPjsWebPartProps {
  description: string;
}
//gavdcodeend 040

export default class OtherPnPjsWebPart extends 
            BaseClientSideWebPart<IOtherPnPjsWebPartProps> {

//gavdcodebegin 032
  private GetUrlFromContext(): void {
    alert(this.context.pageContext.web.absoluteUrl);
  }
//gavdcodeend 032

//gavdcodebegin 034
  private CreateList(): void {
    const sp = spfi().using(SPFx(this.context));
    const spListTitle = "SPFxPnPjsList";  
    const spListDescription = "New List created with PnPjs";  
    const spListTemplateId = 100;  
    const spEnableCT = false;  

    sp.web.lists.add(spListTitle, spListDescription, 
                     spListTemplateId, spEnableCT)
    .then((myResult: IListInfo): void => {
        alert(`List created`);
      }, (myError: Error): void => {  
        alert('Error creating List: ' + myError);
    });
  }   
//gavdcodeend 034

//gavdcodebegin 035
  private UploadFile(): void {
    const sp = spfi().using(SPFx(this.context));
    const myFiles = (<HTMLInputElement>document
                          .getElementById('inpFile')).files;
    const myFile = myFiles![0];

    if (myFile !== undefined || myFile !== null) {
      if (myFile.size <= 10485760) {
        sp.web.getFolderByServerRelativePath("Shared Documents")
        .files.addUsingPath(myFile.name, myFile)
        .then((myData: IFileInfo): void => {
          alert(`Small file uploaded`);
        })
        .catch((myError: Error): void => {
          alert(`Error uploading file` + myError);
        });
      }
      else {
        sp.web.getFolderByServerRelativePath("Shared Documents")
          .files.addChunked(myFile.name, myFile,
            {
              progress: data => { console.log(`progress`); }, 
              Overwrite: true
            })
          .then((myData: IFileInfo): void =>{
            alert(`Big file uploaded`);
          })
          .catch((myError: Error): void => {
            alert(`Error uploading file` + myError);
        });
      }
    }
  }
//gavdcodeend 035

//gavdcodebegin 036
private GetUserProperties(): void {
  const sp = spfi().using(SPFx(this.context));
  let userPropertyValues = ""; 

  sp.profiles.myProperties()
  .then(function(propResult) {  
    const userProperties = propResult.UserProfileProperties;  

    userProperties.forEach((oneProperty: { 
                          Key: string; Value: string }) => {
      userPropertyValues += `${oneProperty.Key} - 
                             ${oneProperty.Value}<br/>`;
    });
  })
   .then((myResult: void): void => {
    alert(userPropertyValues);
    }, (myError: Error): void => {  
      alert('Error geting user: ' + myError);
  });
}
//gavdcodeend 036

//gavdcodebegin 037
private GetOtherUserProperties(): void {  
  const sp = spfi().using(SPFx(this.context));
  let userPropertyValues = "";
  const loginName = "i:0#.f|membership|user@domain.onmicrosoft.com";  

  sp.profiles.getPropertiesFor(loginName)
  .then(function(propResult) {  
    const userProperties = propResult.UserProfileProperties;  

    userProperties.forEach((oneProperty: { 
                          Key: string; Value: string }) => {
      userPropertyValues += `${oneProperty.Key} - 
                             ${oneProperty.Value}<br/>`;
    });
  })
  .then((myResult: void): void => {
    alert(userPropertyValues);
    }, (myError: Error): void => {  
      alert('Error getting user: ' + myError);
  });
}
//gavdcodeend 037

//gavdcodebegin 038
private UpdateUserProperties(): void {
  const sp = spfi().using(SPFx(this.context));
  let userAccName = ""; 

  sp.profiles.myProperties()
  .then(function(propResult) {
    const userProperties = propResult.UserProfileProperties;  

    userProperties.forEach((oneProperty: { 
                        Key: string; Value: string }) => {
      if (oneProperty.Key === "AccountName") {
      userAccName = oneProperty.Value;
      }
    });
  })
  .then((myResult: void): void => {
    sp.profiles.setSingleValueProfileProperty(
                    userAccName, 'AboutMe', 'Books writter')
      .then(() => {
        // handle success
      })
      .catch((error: Error) => {
        alert('Error updating AboutMe property: ' + error);
      });
  })
  .then((myResult: void): void => {
    const mySkills = ["SharePoint", "Microsoft365"]; 
    sp.profiles.setMultiValuedProfileProperty(
                    userAccName, 'SPS-Skills', mySkills)
      .catch((error: Error) => {
        alert('Error updating SPS-Skills property: ' + error);
      });
  })
  .then((myResult: void): void => {
    alert("User properties updated");
  }, (myError: Error): void => {  
    alert('Error updating user properties: ' + myError);
  });
}
//gavdcodeend 038

//gavdcodebegin 039
private GetSearchResults(): void {
  const sp = spfi().using(SPFx(this.context));
  let allSearchRes = "";

  sp.search("SharePoint")  // Search for the word "SharePoint"
  .then((myResult : SearchResults) => {
    const mySearchRes = myResult.PrimarySearchResults;
  
    let counter = 1;
    mySearchRes.forEach(function(object) {
      allSearchRes += counter++ + 
              ". Title - " + object.Title + "<br/>" + 
              "Rank - " + object.Rank + "<br/>" + 
              "File Type - " + object.FileType + "<br/>" + 
              "Original Path - " + object.OriginalPath + "<br/>" + 
              "Summary - " + object.HitHighlightedSummary + "<br/>" + 
              "<br/>";
    });
   })
   .then((myResult: void): void => {
      alert(allSearchRes);
    }, (myError: Error): void => {
      alert('Error getting search: ' + myError);
    });
  }
//gavdcodeend 039

//gavdcodebegin 033
  // private ResponseMessage(myResponse: string): void {  
  //   this.domElement.querySelector('.lblMessage').innerHTML = myResponse;  
  // }
//gavdcodeend 033

//gavdcodebegin 041
public render(): void {
    this.domElement.innerHTML = `
      <div>
        <p><button id="btnGetUrlFromContext">
        <span>Get Url From Context</span></button></p>
        <p><button id="btnCreateList">
        <span>Create a Custom List</span></button></p>
        <input type="file" id="inpFile"></input>
        <p><button id="btnUploadFile">
        <span>Upload a file</span></button></p>
        <p><button id="btnGetUserProperties">
        <span>Get User Properties</span></button></p>
        <p><button id="btnGetOtherUserProperties">
        <span>Get other User Properties</span></button></p>
        <p><button id="btnUpdateUserProperties">
        <span>Update User Properties</span></button></p>
        <p><button id="btnGetSearchResults">
        <span>Get search results</span></button></p>
        <div class="lblMessage"></div>  
      </div>`;

      const btnGetUrlFromContext: HTMLElement | null = 
                this.domElement.querySelector('#btnGetUrlFromContext'); 
      if (btnGetUrlFromContext) { 
        btnGetUrlFromContext.addEventListener('click', () => 
                this.handleBtnGetUrlFromContext());
      }

      const btnCreateList: HTMLElement | null = 
                this.domElement.querySelector('#btnCreateList'); 
      if (btnCreateList) { 
        btnCreateList.addEventListener('click', () => 
                this.handleBtnCreateList());
      }

      const btnUploadFile: HTMLElement | null = 
                this.domElement.querySelector('#btnUploadFile'); 
      if (btnUploadFile) { 
        btnUploadFile.addEventListener('click', () => 
                this.handleBtnUploadFile());
      }

      const btnGetUserProperties: HTMLElement | null = 
                this.domElement.querySelector('#btnGetUserProperties'); 
      if (btnGetUserProperties) { 
        btnGetUserProperties.addEventListener('click', () => 
                this.handleBtnGetUserProperties());
      }

      const btnGetOtherUserProperties: HTMLElement | null = 
                this.domElement.querySelector('#btnGetOtherUserProperties'); 
      if (btnGetOtherUserProperties) { 
        btnGetOtherUserProperties.addEventListener('click', () => 
                this.handleBtnGetOtherUserProperties());
      }

      const btnUpdateUserProperties: HTMLElement | null = 
                this.domElement.querySelector('#btnUpdateUserProperties'); 
      if (btnUpdateUserProperties) { 
        btnUpdateUserProperties.addEventListener('click', () => 
                this.handleBtnUpdateUserProperties());
      }

      const btnGetSearchResults: HTMLElement | null = 
                this.domElement.querySelector('#btnGetSearchResults'); 
      if (btnGetSearchResults) { 
        btnGetSearchResults.addEventListener('click', () => 
                this.handleBbtnGetSearchResults());
      }
  }
  private handleBtnGetUrlFromContext(): void { 
    this.GetUrlFromContext();
  }
  private handleBtnCreateList(): void { 
    this.CreateList();
  }
  private handleBtnUploadFile(): void { 
    this.UploadFile();
  }
  private handleBtnGetUserProperties(): void { 
    this.GetUserProperties();
  }
  private handleBtnGetOtherUserProperties(): void { 
    this.GetOtherUserProperties();
  }
  private handleBtnUpdateUserProperties(): void { 
    this.UpdateUserProperties();
  }
  private handleBbtnGetSearchResults(): void { 
    this.GetSearchResults();
  }
//gavdcodeend 041

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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
