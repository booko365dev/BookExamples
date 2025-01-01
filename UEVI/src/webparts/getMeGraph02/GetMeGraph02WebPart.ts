//gavdcodebegin 403
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ConnectToMicrosoftGraphApiWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IGetMeGraph02WebPartProps {
  description: string;
}
//gavdcodeend 403

//gavdcodebegin 404
export default class GetMeGraph02WebPart extends 
            BaseClientSideWebPart<IGetMeGraph02WebPartProps> {

  private graphClient: MSGraphClientV3;

  protected onInit(): Promise<void> {
    return this.context.msGraphClientFactory.getClient('3')
      .then(client => {
        this.graphClient = client;
      });
  }
//gavdcodeend 404

//gavdcodebegin 405
public render(): void {
    this.domElement.innerHTML = `
      <div>
        <h2>Information about Me 02</h2>
        <div id="userInfo"></div>
      </div>`;

    this.getUserInfo().catch(error => {
      console.error('Error in getUserInfo:', error);
    });
  }
//gavdcodeend 405

//gavdcodebegin 406
private getUserInfo(): Promise<void> {
    return this.graphClient
    .api('/me')
    .version('v1.0')
    .get()
      .then(myResponse => {
        const myMe = this.domElement.querySelector('#userInfo');
        if (myMe) {
          myMe.innerHTML = `
            <p>Name - ${myResponse.displayName}</p>
            <p>Email - ${myResponse.mail}</p>
            <p>Phone Number - ${myResponse.businessPhones[0]}</p>
          `;
        }
      })
      .catch((myError: Error) => {
        console.error('Error getting user info:', myError);
      });
  }
//gavdcodeend 406

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
