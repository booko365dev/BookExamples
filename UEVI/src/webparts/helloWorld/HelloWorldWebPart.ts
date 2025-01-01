//gavdcodebegin 001
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
  maxRandom: number
}
//gavdcodeend 001

export default class HelloWorldWebPart extends 
            BaseClientSideWebPart<IHelloWorldWebPartProps> {

//gavdcodebegin 004
private GetSomeRandom(): string {
  console.log("GetSomeRandom started");
  const myMax = Number(this.properties.maxRandom);
  return String(Math.floor(Math.random() * myMax));
}
//gavdcodeend 004

//gavdcodebegin 002
public render(): void {
    this.domElement.innerHTML = `
    <p><button id="btnRandom">
          Get Random Number</button>
    </p>`;

    const btnRandom: HTMLElement | null = this.domElement.querySelector('#btnRandom'); 
    if (btnRandom) { 
      btnRandom.addEventListener('click', () => this.handleMyButtonClick());    
    }

  }

  private handleMyButtonClick(): void { 
    alert(this.GetSomeRandom());
  }
  //gavdcodeend 002

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //gavdcodebegin 003
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
                  min: 0, 
                  max: 100, 
                  step: 1, 
                  value: 50, 
                  showValue: true                })
              ]
            }
          ]
        }
      ]
    };
  }
//gavdcodeend 003
}
