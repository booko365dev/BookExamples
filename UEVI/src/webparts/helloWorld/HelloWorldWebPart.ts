
//gavdcodebegin 01
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
  maxRandom: string;
}
//gavdcodeend 01

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

//gavdcodebegin 04
  private GetSomeRandom(): void {
    console.log("GetSomeRandom started");
    var myMax = Number(this.properties.maxRandom);
    document.getElementById("divRandomString").innerHTML = 
                      String(Math.floor(Math.random() * myMax));
  }
//gavdcodeend 04

//gavdcodebegin 02
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">
                        Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">
                        ${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
                    
              <p><button id="btnRandom" class="${ styles.button }">
                          Get Random Number</button>
              <div id="divRandomString" class="${ styles.label }" /></div></p>

            </div>
          </div>
        </div>
      </div>`;

      document.getElementById("btnRandom").onclick = 
                                          this.GetSomeRandom.bind(this);

  }
//gavdcodeend 02

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

//gavdcodebegin 03
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
                }),
                PropertyPaneSlider('maxRandom', {
                  label: 'Max Random',
                  min:1,
                  max:10
                })
              ]
            }
          ]
        }
      ]
    };
  }
//gavdcodebegin 03
}
