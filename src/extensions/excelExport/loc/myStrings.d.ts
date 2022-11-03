declare interface IExcelExportCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ExcelExportCommandSetStrings' {
  const strings: IExcelExportCommandSetStrings;
  export = strings;
}
