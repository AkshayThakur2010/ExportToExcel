import { Component, OnInit } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { LabelConstants } from './constants/label.const';
import { ExportToExcelService } from './services/export-to-excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'ExportToExcel';
  isEnglishLanguage: boolean;
  tableColumn = LabelConstants.TableColumn;
  employeeList: any[] = [
    {
      empId: 'E_101',
      name: 'SMITH',
      dept: 'SALESMAN',
      manager: 'ADAMS'
    },
    {
      empId: 'E_201',
      name: 'ALLEN',
      dept: 'CLERK',
      manager: 'ADAMS'
    }
  ];
  constructor(
    private translate: TranslateService,
    private exportExcel: ExportToExcelService
  ) {
    this.translate.addLangs(['en', 'pt']);
    this.translate.use('en');
    this.isEnglishLanguage = true;
  }

  ngOnInit() {
  }


  changeLanguage() {
    if (this.isEnglishLanguage) {
      this.translate.use('pt');
      this.isEnglishLanguage = false;
    } else {
      this.translate.use('en');
      this.isEnglishLanguage = true;
    }
  }

  exportToExcel(is_csv: boolean) {
    const tableHeaders = [];
    this.tableColumn.forEach(function (value) {
      tableHeaders.push(value.header);
    });

    this.exportExcel.exportAsExcelFile(this.employeeList, tableHeaders,
      LabelConstants.FileName, is_csv);
  }
}
