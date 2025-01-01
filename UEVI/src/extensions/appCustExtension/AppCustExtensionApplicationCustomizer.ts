//gavdcodebegin 102
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
//gavdcodeend 102

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

//gavdcodebegin 101
public onInit(): Promise<void> {
  const topPholder: PlaceholderContent | undefined = 
  this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    if (topPholder) {
    topPholder.domElement.innerHTML = `
          <p style="height:34px;text-align:center;font-weight:bold;color:white;
          background-color:#02767a;line-height:2.5;">Guitaca Editors</p>`;
    }

    const bottomPholder: PlaceholderContent | undefined = 
      this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
    if (bottomPholder) {
    bottomPholder.domElement.innerHTML = `<div class="">
          <p style="height:34px;text-align:center;font-weight:bold;color:white;
          background-color:#02767a;line-height:2.5;">https://www.guitaca.com</p>`;
    }

    return Promise.resolve();
  }
}
//gavdcodeend 101
