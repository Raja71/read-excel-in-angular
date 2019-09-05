import { Component } from '@angular/core'  
import * as XLSX from 'xlsx'  
import * as FileSaver from 'file-saver'  

@Component({  
  selector: 'app-root',  
  templateUrl: './app.component.html',  
  styleUrls: ['./app.component.css']  
})  
export class AppComponent {  

  title = 'read-excel-in-angular'  
  storeData: any  
  csvData: any  
  jsonData: any  
  textData: any  
  htmlData: any  
  fileUploaded: File  
  worksheet: any 

  uploadedFile(event: any) {  
    this.fileUploaded = event.target.files[0]  
    this.readExcel()  
  }

  //function to read excel file
  readExcel() {  
    let readFile = new FileReader()  
    readFile.onload = (e) => {  
      this.storeData = readFile.result  
      let data = new Uint8Array(this.storeData)  
      let arr = new Array()  
      for (let i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i])  
      let bstr = arr.join("")  
      let workbook = XLSX.read(bstr, { type: "binary" })  
      let first_sheet_name = workbook.SheetNames[0]  
      this.worksheet = workbook.Sheets[first_sheet_name]  
    }  
    readFile.readAsArrayBuffer(this.fileUploaded)  
  }  

  //function to download as csv file
  readAsCSV() {  
    this.csvData = XLSX.utils.sheet_to_csv(this.worksheet)  
    const data: Blob = new Blob([this.csvData], { type: 'text/csvcharset=utf-8' })  
    FileSaver.saveAs(data, "CSVFile" + new Date().getTime() + '.csv')  
  }  

  //function to download as json file
  readAsJson() {  
    this.jsonData = XLSX.utils.sheet_to_json(this.worksheet, { raw: false })  
    this.jsonData = JSON.stringify(this.jsonData)  
    const data: Blob = new Blob([this.jsonData], { type: "application/json" })  
    FileSaver.saveAs(data, "JsonFile" + new Date().getTime() + '.json')  
  }  

  //function to download as html file
  readAsHTML() {  
    this.htmlData = XLSX.utils.sheet_to_html(this.worksheet)  
    const data: Blob = new Blob([this.htmlData], { type: "text/htmlcharset=utf-8" })  
    FileSaver.saveAs(data, "HtmlFile" + new Date().getTime() + '.html')  
  }  

  //function to download as text file
  readAsText() {  
    this.textData = XLSX.utils.sheet_to_txt(this.worksheet)  
    const data: Blob = new Blob([this.textData], { type: 'text/plaincharset=utf-8' })  
    FileSaver.saveAs(data, "TextFile" + new Date().getTime() + '.txt')  
  }  

}  