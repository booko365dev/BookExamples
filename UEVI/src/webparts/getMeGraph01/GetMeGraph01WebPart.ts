//gavdcodebegin 401
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ConnectToMicrosoftGraphApiWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGetMeGraph01WebPartProps {
  description: string;
}
//gavdcodeend 401

export default class GetMeGraph01WebPart extends 
            BaseClientSideWebPart<IGetMeGraph01WebPartProps> {

//gavdcodebegin 402
  public render(): void {
    
    this.context.msGraphClientFactory
    .getClient('3')
      .then((graphClient: MSGraphClientV3): void => {

        interface IUserResponse {
          displayName: string;
          mail: string;
          businessPhones: string[];
        }
  
        graphClient
        .api('/me')
        .version('v1.0')
        .get()
        .then((myResponse: IUserResponse) => {
          
          // Only for debugging purpose
          //console.log("My Response - " + JSON.stringify(myResponse));
          
          this.domElement.innerHTML = `
            <h2>Information about Me 01</h2>
            <p>DisplayName - ${myResponse.displayName}</p>
            <p>Email - ${myResponse.mail}</p>
            <p>Phone Number - ${myResponse.businessPhones[0]}</p>
            </div>
          `;
        })
        .catch((myError: Error) => {
          console.error("Error in Get - " + myError);
        });
    }).catch((myError: Error) => {
      console.error("Error in Client - " + myError);
    });
  }
//gavdcodeend 402

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
