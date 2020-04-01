declare interface IListViewExtCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ListViewExtCommandSetStrings' {
  const strings: IListViewExtCommandSetStrings;
  export = strings;
}
