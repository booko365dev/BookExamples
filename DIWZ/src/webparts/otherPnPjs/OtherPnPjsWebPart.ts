//gavdcodebegin 40
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { sp, Web, List, ListAddResult, SearchResults } from "@pnp/sp";

import styles from './OtherPnPjsWebPart.module.scss';
import * as strings from 'OtherPnPjsWebPartStrings';
import { any } from 'prop-types';

export interface IOtherPnPjsWebPartProps {
  description: string;
}
//gavdcodeend 40

export default class OtherPnPjsWebPart extends BaseClientSideWebPart<IOtherPnPjsWebPartProps> {

//gavdcodebegin 32
  private GetUrlFromContext(): void {
    this.ResponseMessage(this.context.pageContext.web.absoluteUrl);
  }
//gavdcodeend 32

//gavdcodebegin 34
  private CreateList(): void {
    let spListTitle = "SPFxPnPjsList";  
    let spListDescription = "New List created with PnPjs";  
    let spListTemplateId = 100;  
    let spEnableCT = false;  

    sp.web.lists.add(spListTitle, spListDescription, spListTemplateId, spEnableCT)
    .then((myResult: ListAddResult): void => {
        this.ResponseMessage(`List created`);
      }, (myError: any): void => {  
        this.ResponseMessage('Error creating List: ' + myError);
    });
  }   
//gavdcodeend 34

//gavdcodebegin 35
  private UploadFile(): void {
    var myFiles = (<HTMLInputElement>document.getElementById('inpFile')).files;
    var myFile = myFiles[0];

    if (myFile!=undefined || myFile!=null) {
      if (myFile.size <= 10485760) {
        sp.web.getFolderByServerRelativeUrl(
        this.context.pageContext.web.serverRelativeUrl + "/Shared Documents")
        .files.add(myFile.name, myFile, true)
        .then((myUploadedFile: any): void => {
          myUploadedFile.file.getItem()
          .then((myItem: any): void => {
            myItem.update({
              Title:'TestFile'
            })
          })
        })
        .then((myData: any): void => {
          this.ResponseMessage(`Small file uploaded`);
        })
        .catch((myError: any): void => {
          this.ResponseMessage(`Error uploading file` + myError);
        });
      }
      else {
        sp.web.getFolderByServerRelativeUrl(
          this.context.pageContext.web.serverRelativeUrl + "/Shared Documents")
          .files.addChunked(myFile.name, myFile, myData =>
            {
               console.log({ data: myData, message: "progress" });
            }, true)
          .then((myData: any): void =>{
            this.ResponseMessage(`Big file uploaded`);
          })
          .catch((myError: any): void => {
            this.ResponseMessage(`Error uploading file` + myError);
        });
      }
    }
  }
//gavdcodeend 35

//gavdcodebegin 36
private GetUserProperties(): void {
  var userPropertyValues = ""; 

  sp.profiles.myProperties.get()
  .then(function(propResult) {  
    var userProperties = propResult.UserProfileProperties;  

    userProperties.forEach(function(oneProperty) {  
        userPropertyValues += oneProperty.Key + " - " + oneProperty.Value + "<br/>";  
    })
  })
  .then((myResult: any): void => {
    this.ResponseMessage(userPropertyValues);
  }, (myError: any): void => {  
    this.ResponseMessage('Error geting user: ' + myError);
  });
}
//gavdcodeend 36

//gavdcodebegin 37
private GetOtherUserProperties(): void {  
  var userPropertyValues = "";

  let loginName = "i:0#.f|membership|oneuser@onedomain.onmicrosoft.com";  
  sp.profiles.getPropertiesFor(loginName)
  .then(function(propResult) {  
    var userProperties = propResult.UserProfileProperties;  

    userProperties.forEach(function(oneProperty) {  
        userPropertyValues += oneProperty.Key + " - " + oneProperty.Value + "<br/>";  
    })
  })
  .then((myResult: any): void => {
    this.ResponseMessage(userPropertyValues);
  }, (myError: any): void => {  
    this.ResponseMessage('Error getting user: ' + myError);
  });
}
//gavdcodeend 37

//gavdcodebegin 38
private UpdateUserProperties(): void {
  var userAccName = ""; 

  sp.profiles.myProperties.get()
  .then(function(propResult) {
    var userProperties = propResult.UserProfileProperties;  

    userProperties.forEach(function(oneProperty) { 
      if(oneProperty.Key == "AccountName") {
        userAccName = oneProperty.Value;
      } 
    })
  })
  .then((myResult: any): void => {
    sp.profiles.setSingleValueProfileProperty(userAccName, 'AboutMe', 'Books writter');
  })
  .then((myResult: any): void => {
    let mySkills = ["SharePoint", "Office365"]; 
    sp.profiles.setMultiValuedProfileProperty(userAccName, 'SPS-Skills', mySkills);
  })
  .then((myResult: any): void => {
    this.ResponseMessage("User properties updated");
  }, (myError: any): void => {  
    this.ResponseMessage('Error updating user properties: ' + myError);
  });
}
//gavdcodeend 38

//gavdcodebegin 39
private GetSearchResults(): void {
  var allSearchRes = "";

  sp.search("SharePoint")  // Search for the word "SharePoint"
  .then((myResult : SearchResults) => {
    var mySearchRes = myResult.PrimarySearchResults;
  
    var counter = 1;
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
   .then((myResult: any): void => {
      this.ResponseMessage(allSearchRes);
    }, (myError: any): void => {
      this.ResponseMessage('Error getting search: ' + myError);
    });
  }
//gavdcodeend 39

//gavdcodebegin 33
  private ResponseMessage(myResponse: string): void {  
    this.domElement.querySelector('.lblMessage').innerHTML = myResponse;  
  }
//gavdcodeend 33

//gavdcodebegin 41
public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.otherPnPjs }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <p><button id="btnGetUrlFromContext" class="${ styles.button }">
              <span class="${styles.label}">Get Url From Context</span></button></p>
              <p><button id="btnCreateList" class="${ styles.button }">
              <span class="${styles.label}">Create a Custom List</span></button></p>
              <input type="file" id="inpFile" class="${ styles.button }"></input>
              <p><button id="btnUploadFile" class="${ styles.button }">
              <span class="${styles.label}">Upload a file</span></button></p>
              <p><button id="btnGetUserProperties" class="${ styles.button }">
              <span class="${styles.label}">Get User Properties</span></button></p>
              <p><button id="btnGetOtherUserProperties" class="${ styles.button }">
              <span class="${styles.label}">Get other User Properties</span></button></p>
              <p><button id="btnUpdateUserProperties" class="${ styles.button }">
              <span class="${styles.label}">Update User Properties</span></button></p>
              <p><button id="btnGetSearchResults" class="${ styles.button }">
              <span class="${styles.label}">Get search results</span></button></p>
              <div class="lblMessage"></div>  
            </div>
          </div>
        </div>
      </div>`;

      document.getElementById("btnGetUrlFromContext").onclick = 
                                          this.GetUrlFromContext.bind(this);
      document.getElementById("btnCreateList").onclick = 
                                          this.CreateList.bind(this);
      document.getElementById("btnUploadFile").onclick = 
                                          this.UploadFile.bind(this);
      document.getElementById("btnGetUserProperties").onclick = 
                                          this.GetUserProperties.bind(this);
      document.getElementById("btnGetOtherUserProperties").onclick = 
                                          this.GetOtherUserProperties.bind(this);
      document.getElementById("btnUpdateUserProperties").onclick = 
                                          this.UpdateUserProperties.bind(this);
      document.getElementById("btnGetSearchResults").onclick = 
                                          this.GetSearchResults.bind(this);
  }
//gavdcodeend 41

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
