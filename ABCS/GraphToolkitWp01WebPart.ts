//gavdcodebegin 003
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GraphToolkitWp01WebPartStrings';

import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

export interface IGraphToolkitWp01WebPartProps {
  description: string;
}

export default class GraphToolkitWp01WebPart extends BaseClientSideWebPart<IGraphToolkitWp01WebPartProps> {

  protected async onInit(): Promise<void> {
    Providers.globalProvider = new SharePointProvider(this.context)
  }

  public render(): void {
    this.domElement.innerHTML = `
      <h1>People in one Group<h1>  
      <mgt-people group-id="b624b3f3-ec54-4e2b-a5d5-2f53197d21ed"></mgt-people>
      <hr>
      <h1>My Agenda</h1>  
      <mgt-agenda group-by-day></mgt-agenda>  
      <hr>  
      <h1>My info</h1>
      <mgt-person person-query="me" view="twolines" show-name show-email><mgt-person>
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
//gavdcodeend 003