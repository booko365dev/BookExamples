//gavdcodebegin 02
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AppCustExtensionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AppCustExtensionApplicationCustomizer';
//gavdcodeend 02

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppCustExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppCustExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IAppCustExtensionApplicationCustomizerProperties> {

//gavdcodebegin 01
@override
  public onInit(): Promise<void> {
    let topPholder: PlaceholderContent = 
          this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    if (topPholder) {
      topPholder.domElement.innerHTML = `
              <p style="height:34px;text-align:center;font-weight:bold;color:white;
              background-color:#02767a;line-height:2.5;">Guitaca Editors</p>`;
    }

    let bottomPholder: PlaceholderContent = 
          this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
    if (bottomPholder) {
      bottomPholder.domElement.innerHTML = `<div class="">
              <p style="height:34px;text-align:center;font-weight:bold;color:white;
              background-color:#02767a;line-height:2.5;">https://www.guitaca.com</p>`;
    }

    return Promise.resolve();
  }
}
//gavdcodeend 01
