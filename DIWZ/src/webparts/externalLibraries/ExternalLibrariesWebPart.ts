import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ExternalLibrariesWebPart.module.scss';
import * as strings from 'ExternalLibrariesWebPartStrings';

//gavdcodebegin 015
import * as validator from 'validator';
//gavdcodeend 015

//gavdcodebegin 018
import * as marked from 'marked'; 
//gavdcodeend 018

//gavdcodebegin 020
import * as $ from 'jquery'; 
//gavdcodeend 020

export interface IExternalLibrariesWebPartProps {
  description: string;
}

export default class ExternalLibrariesWebPart extends BaseClientSideWebPart<IExternalLibrariesWebPartProps> {

//gavdcodebegin 017
  private ValidateInteger(): void {
    document.getElementById("divMessage").innerHTML = 
                      validator.isInt('somestring');  // Should get 'false'
  }
//gavdcodeend 017

//gavdcodebegin 019
private ShowMarkdown(): void {
  document.getElementById("divMessage").innerHTML = 
                  marked('This string is __bold__');  // Should show html text
}
//gavdcodeend 019

//gavdcodebegin 021
private ShowWithJquery(): void {
	$("#divMessage").text("Some text in the label");   // Should show text in the label
}
//gavdcodeend 021

//gavdcodebegin 016
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
//gavdcodeend 016

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
