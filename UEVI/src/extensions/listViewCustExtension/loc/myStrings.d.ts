declare interface IListViewCustExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ListViewCustExtensionCommandSetStrings' {
  const strings: IListViewCustExtensionCommandSetStrings;
  export = strings;
}
