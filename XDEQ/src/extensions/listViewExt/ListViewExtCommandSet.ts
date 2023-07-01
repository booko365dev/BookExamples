import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ListViewExtCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
//gavdcodebegin 001
export interface IListViewExtCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
  sampleTextThree: string;
}
//gavdcodeend 001

const LOG_SOURCE: string = 'ListViewExtCommandSet';

export default class ListViewExtCommandSet extends BaseListViewCommandSet<IListViewExtCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ListViewExtCommandSet');
    return Promise.resolve();
  }

//gavdcodebegin 002
@override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_3');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }
//gavdcodeend 002

//gavdcodebegin 003
@override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne}`);
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      case 'COMMAND_3':
        Dialog.prompt(`${this.properties.sampleTextThree}`)
        .then(retString => {
          Dialog.alert(`Input string = ${retString}`);
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
//gavdcodeend 003
}
