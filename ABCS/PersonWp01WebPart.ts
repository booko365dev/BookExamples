//gavdcodebegin 04
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PersonWp01WebPart.module.scss';
import * as strings from 'PersonWp01WebPartStrings';

import { Providers, SharePointProvider } from '@microsoft/mgt';

export interface IPersonWp01WebPartProps {
  description: string;
}

export default class PersonWp01WebPart extends
    BaseClientSideWebPart<IPersonWp01WebPartProps> {

  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context)
  }

  public render(): void {
    this.domElement.innerHTML = `
    <h1>Info about Me</h1>
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
//gavdcodeend 04
