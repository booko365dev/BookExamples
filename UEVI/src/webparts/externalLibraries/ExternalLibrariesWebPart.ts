import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ExternalLibrariesWebPart.module.scss';
import * as strings from 'ExternalLibrariesWebPartStrings';

//gavdcodebegin 15
import * as validator from 'validator';
//gavdcodeend 15

//gavdcodebegin 18
import * as marked from 'marked'; 
//gavdcodeend 18

//gavdcodebegin 20
import * as $ from 'jquery'; 
//gavdcodeend 20

export interface IExternalLibrariesWebPartProps {
  description: string;
}

export default class ExternalLibrariesWebPart extends BaseClientSideWebPart<IExternalLibrariesWebPartProps> {

//gavdcodebegin 17
  private ValidateInteger(): void {
    document.getElementById("divMessage").innerHTML = 
                      validator.isInt('somestring');  // Should get 'false'
  }
//gavdcodeend 17

//gavdcodebegin 19
private ShowMarkdown(): void {
  document.getElementById("divMessage").innerHTML = 
                  marked('This string is __bold__');  // Should show html text
}
//gavdcodeend 19

//gavdcodebegin 21
private ShowWithJquery(): void {
	$("#divMessage").text("Some text in the label");   // Should show text in the label
}
//gavdcodeend 21

//gavdcodebegin 16
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.externalLibraries }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <p><button id="btnValidate" class="${ styles.button }">
              <span class="${styles.label}">Validate if it is Integer</span></button></p>
              <p><button id="btnMarked" class="${ styles.button }">
              <span class="${styles.label}">Show some markdown</span></button></p>
              <p><button id="btnJquery" class="${ styles.button }">
              <span class="${styles.label}">Using jQuery</span></button></p>
              <div id="divMessage" class="${ styles.label }" /></div></p>
              </div>
          </div>
        </div>
      </div>`;

      document.getElementById("btnValidate").onclick = 
                                          this.ValidateInteger.bind(this);
      document.getElementById("btnMarked").onclick = 
                                          this.ShowMarkdown.bind(this);
      document.getElementById("btnJquery").onclick = 
                                          this.ShowWithJquery.bind(this);
  }
//gavdcodeend 16

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
