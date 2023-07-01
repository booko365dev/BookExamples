//gavdcodebegin 003
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphToolkitWp01WebPart.module.scss';
import * as strings from 'GraphToolkitWp01WebPartStrings';

import { Providers, SharePointProvider } from '@microsoft/mgt';

export interface IGraphToolkitWp01WebPartProps {
  description: string;
}

export default class GraphToolkitWp01WebPart extends
    BaseClientSideWebPart<IGraphToolkitWp01WebPartProps> {

  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context)
  }

  public render(): void {
    this.domElement.innerHTML = `
      <h1>People in one Group<h1>  
      <mgt-people group-id="194b9866-e05c-489e-aaf6-51b31ce22a91"></mgt-people>
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
