import { Injectable } from '@angular/core';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import { TranslateService } from '@ngx-translate/core';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const EXCEL_EXTENSION = '.xlsx';
const EXCEL_BOOKTYPE = 'xlsx';
const CSV_EXTENSION = '.csv';
const CSV_BOOKTYPE = 'csv';

@Injectable({
  providedIn: 'root'
})
export class ExportToExcelService {
  translatedHeaderNames: any[];
  headers: any[];

  constructor(
    private translate: TranslateService
  ) { }

  public exportAsExcelFile(recordObj: any[], headerNames: any[], fileName: string, is_export_to_excel: boolean = true): void {

    this.removeJsonIdProperty(recordObj);

    const workSheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(recordObj);
    this.translatedHeaderNames = this.translateHeaderNames(headerNames);
    // workSheet['!cols'] = [{ wch: 200 }, { wch: 200 }, { wch: 200 }, { wch: 200 },
    // { wch: 200 }, { wch: 200 }, { wch: 200 }, { wch: 200 },
    // { wch: 200 }, { wch: 200 }, { wch: 200 }, { wch: 200 },
    // { wch: 200 }, { wch: 200 }, { wch: 200 }, { wch: 200 }]

    workSheet['!cols'] = this.autoColumnWidth(recordObj);

    this.setHeaderRow(workSheet, this.translateHeaderNames(headerNames));

    const workBook: XLSX.WorkBook = { Sheets: { 'data': workSheet }, SheetNames: ['data'] };

    const excelBuffer: any = XLSX.write(workBook, { bookType: (is_export_to_excel ? EXCEL_BOOKTYPE : CSV_BOOKTYPE), type: 'array' });

    this.saveExcelFile(excelBuffer, fileName, is_export_to_excel);
  }

  private autoColumnWidth(json: any[]) {
    let cols = [];
    let wsCols = [];
    let header = '';
    this.headers = Object.keys(json[0]);

    for (let i = 0; i < this.headers.length;) {
      header = this.headers[i];
      json.forEach(element => {
        cols.push(element[header]);
      });
      wsCols.push({ wch: Math.max(...(cols.map(el => this.getWidth(el, this.translatedHeaderNames[i])))) });
      cols = [];
      i++;
    }
    return wsCols;
  }

  private getWidth(el: any, header: string) {
    el = el != null ? el : '';
    return header.length > el.toString().length ? header.length : el.toString().length;
  }

  private saveExcelFile(bufferObj: any, fileName: string, is_export_to_excel: boolean): void {
    const data: Blob = new Blob([bufferObj], { type: EXCEL_TYPE });
    FileSaver.saveAs(data, fileName + (is_export_to_excel ? EXCEL_EXTENSION : CSV_EXTENSION));
  }

  translateHeaderNames(headerNames: string[]): string[] {
    const translatedHeaders = [];
    this.translate.get(headerNames).subscribe((res: string) => {
      Object.keys(res).map(key => translatedHeaders.push(res[key]));
    });
    return translatedHeaders;
  }

  removeJsonIdProperty(recordObj: any[]) {
    recordObj.forEach(function (record) {
      if (record.id) {
        delete record.id;
      }
    });
  }
  setHeaderRow(sheet: any, headerNames: any): void {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    let C = range.s.r; /* start in the first row */
    const R = range.s.r; /* start in the first row */
    /* walk every column in the range */
    let i = 0;
    for (C = range.s.c; C <= range.e.c; ++C) {
      const cell = sheet[XLSX.utils.encode_cell({ c: C, r: R })]; /* find the cell in the first row */

      if (cell && cell.t) {
        cell.v = headerNames[i];
      }
      i++;
    }
  }
}
