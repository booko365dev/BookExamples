//gavdcodebegin 001
import { Version } from '@microsoft/sp-core-library';
import * as microsoftTeams from '@microsoft/teams-js';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
//gavdcodeend 001

import styles from './SpFxAsTeamsTabWebPart.module.scss';
import * as strings from 'SpFxAsTeamsTabWebPartStrings';

export interface ISpFxAsTeamsTabWebPartProps {
  description: string;
}

//gavdcodebegin 002
export default class SpFxAsTeamsTabWebPart extends BaseClientSideWebPart<ISpFxAsTeamsTabWebPartProps> {
  private myTeamsContext: microsoftTeams.Context;

protected onInit(): Promise<any> {
  let retVal: Promise<any> = Promise.resolve();
  if (this.context.microsoftTeams) {
    retVal = new Promise((resolve, reject) => {
      this.context.microsoftTeams.getContext(context => {
        this.myTeamsContext = context;
        resolve();
      });
    });
  }
  return retVal;
}
//gavdcodeend 002

//gavdcodebegin 003
  public render(): void {

    let webpartContext: string = '';

    if (this.myTeamsContext) {
      webpartContext = "Team: " + this.myTeamsContext.teamName;
    }
    else {
      webpartContext = "SharePoint site: " + this.context.pageContext.web.title;
    }

    this.domElement.innerHTML = `
      <div class="${ styles.spFxAsTeamsTab }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${webpartContext}</span>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
            </div>
          </div>
        </div>
      </div>`;
  }
//gavdcodeend 003

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
