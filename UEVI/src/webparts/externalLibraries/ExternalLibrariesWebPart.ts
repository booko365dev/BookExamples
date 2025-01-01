import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

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
    const divMessage = document.getElementById("divMessage");
    if (divMessage) {
      divMessage.innerHTML = validator.isInt('somestring').toString();  // Should get 'false'
    }
  }
//gavdcodeend 017

//gavdcodebegin 019
private ShowMarkdown(): void {
  const divMessage = document.getElementById("divMessage");
  if (divMessage) {
    const result = marked.parse('This string is __bold__');
    if (result instanceof Promise) {
      result.then(res => { divMessage.innerHTML = res; })
      .catch(err => console.error(err));
    } else {
      divMessage.innerHTML = result;
    }
  }
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
      <div>
        <p><button id="btnValidate">
        <span>Validate if it is Integer</span></button></p>
        <p><button id="btnMarked">
        <span>Show some markdown</span></button></p>
        <p><button id="btnJquery">
        <span>Using jQuery</span></button></p>
        <p><label id="divMessage">This is a message</label></p>
        </div>
      </div>`;

      const btnValidate: HTMLElement | null = this.domElement.querySelector('#btnValidate'); 
      if (btnValidate) { 
        btnValidate.addEventListener('click', () => this.handleBtnValidate());
      }

      const btnMarked: HTMLElement | null = this.domElement.querySelector('#btnMarked'); 
      if (btnMarked) { 
        btnMarked.addEventListener('click', () => this.handleBtnMarked());
      }

      const btnJquery: HTMLElement | null = this.domElement.querySelector('#btnJquery'); 
      if (btnJquery) { 
        btnJquery.addEventListener('click', () => this.handleBtnJquery());
      }
  }
  private handleBtnValidate(): void { 
    this.ValidateInteger();
  }
  private handleBtnMarked(): void { 
    this.ShowMarkdown();
  }
  private handleBtnJquery(): void { 
    this.ShowWithJquery();
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