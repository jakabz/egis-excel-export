import * as XLSX from 'xlsx';
import { CSVLink, CSVDownload } from "react-csv";
import {saveAs}  from 'file-saver';

export class fileExport {

    public Excel = (head:string[], rows:any[], fileName:string) => {
      const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';  
      const fileExtension = '.xlsx';  
      const Heading = [ head ];  
      const ws = XLSX.utils.book_new();   
      XLSX.utils.sheet_add_aoa(ws, Heading);  
      XLSX.utils.sheet_add_json(ws, rows, { origin: 'A2', skipHeader: true });         
      const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };  
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });  
      const data = new Blob([excelBuffer], {type: fileType});  
      saveAs(data, fileName + fileExtension);  
    }

}

const FileExport = new fileExport();  
export default FileExport;