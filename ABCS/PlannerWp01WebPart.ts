//gavdcodebegin 001
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GraphToolkitWp01WebPartStrings';

import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

export interface IPlannerWp01WebPartProps {
  description: string;
}

export default class PlannerWp01WebPart extends BaseClientSideWebPart<IPlannerWp01WebPartProps> {

  protected async onInit(): Promise<void> {
    Providers.globalProvider = new SharePointProvider(this.context)
  }

  public render(): void {
    this.domElement.innerHTML = `
      <mgt-login></mgt-login>
      <mgt-tasks></mgt-tasks>
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
//gavdcodeend 001
