//gavdcodebegin 004
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PersonWp01WebPartStrings';

import { Providers, SharePointProvider, ProviderState } from '@microsoft/mgt-spfx';

export interface IPersonWp01WebPartProps {
  description: string;
}

export default class PersonWp01WebPart extends BaseClientSideWebPart<IPersonWp01WebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <mgt-login></mgt-login>
    <mgt-person person-query="me" view="twolines" show-name show-email person-card="click">
      <template data-type="person-card">
        <mgt-person-card person-details="{{person}}" 
            person-image="{{personImage}}">
          <template data-type="additional-details">
            <h3>My cars:</h3>
            <ol>
              <li>Ferrari</li>
              <li>Lamborghini</li>
              <li>Rolls Royce</li>
            </ol>
          </template>
        </mgt-person-card>
      </template>
    </mgt-person>
    `;

    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }

    // Check if the user is signed in 
    Providers.globalProvider.onStateChanged(() => { 
      if (Providers.globalProvider.state === ProviderState.SignedIn) { 
        console.log("-- User is signed in"); 
      } 
      else { 
        console.log("-- User is not signed in"); 
      }
    });
  }

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
//gavdcodeend 004
