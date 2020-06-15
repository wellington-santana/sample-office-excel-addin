//eslint-disable-next-line
import { Component, ViewChild, Inject } from "@angular/core";
import * as CPF from "@fnando/cpf";
// import * as OfficeHelpers from "@microsoft/office-js-helpers";
//eslint-disable-next-line
import ImportResult from "../../models/import-result.model";
import ExcelAddinHelperService from "../excel-addin-helper/excel-addin-helper.service";
// eslint-disable-next-line no-unused-vars
import ExcelSelectionMap from "../excel-addin-helper/models/excel-selection-map.model";
const template = require("./main.component.html");
//eslint-disable-next-line
/* global require, Excel, console, ngModel, OfficeHelpers, Office */

@Component({
  selector: "main-component",
  template: template
})
export default class MainComponent {
  targetColumn: String = "";
  feedbackColumn: String = "";
  targetLinhaInicial: String = "";
  targetLinhaFinal: String = "";
  hasHeaders: boolean = true;
  log: String;
  importResults: ImportResult[] = [];
  excelAddinService: ExcelAddinHelperService;
  isImporting: boolean = false;

  @ViewChild("modal") modal: any;

  constructor(@Inject(ExcelAddinHelperService) excelAddinService: ExcelAddinHelperService) {
    this.excelAddinService = excelAddinService;
  }

  setColumnCpf(text: String) {
    this.targetColumn = text;
  }

  setColumnFeedback(text: String) {
    this.feedbackColumn = text;
  }

  setInicial(text: String) {
    this.targetLinhaInicial = text;
  }

  setFinal(text: String) {
    this.targetLinhaFinal = text;
  }

  async insertFeedback(columnName: string, columnIndex: number, rowIndex: number, isSuccess: boolean, message: string) {
    this.importResults.push(new ImportResult(columnName, columnIndex, rowIndex, isSuccess, message));
  }

  async goToCell(result: ImportResult) {
    Excel.run(async context => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getCell(result.rowIndex, result.columnIndex).select();
    });
  }

  private countImportErrors(): number {
    return this.importResults.filter(r => !r.isSuccess).length;
  }

  private countImportSuccess(): number {
    return this.importResults.filter(r => r.isSuccess).length;
  }

  private importRangeSpecified(selecionList: ExcelSelectionMap, sheet: Excel.Worksheet, range: Excel.Range) {
    let contRow = selecionList.firtsCell.rowIndex;

    range.values.forEach(lineValues => {
      //A ordenação que o excel devolve é diferente da ordenação do mapeamento
      //o filtro abaixo ajusta conforme a itereção ocorre
      let contColumn = 0;
      let selecionListFiltered = selecionList.selectedCells.filter(v => v.rowIndex == contRow);
      lineValues.forEach(columnValue => {
        let cpf: string = columnValue;
        let selectedCell = selecionListFiltered[contColumn];
        contColumn++;

        if (!cpf) {
          this.insertFeedback(
            selectedCell.columnName,
            selectedCell.columnIndex,
            selectedCell.rowIndex,
            false,
            "Valor vazio"
          );
          return;
        }

        if (!CPF.isValid(cpf)) {
          this.insertFeedback(
            selectedCell.columnName,
            selectedCell.columnIndex,
            selectedCell.rowIndex,
            false,
            "CPF inválido"
          );
          return;
        }

        this.insertFeedback(
          selectedCell.columnName,
          selectedCell.columnIndex,
          selectedCell.rowIndex,
          true,
          "Importado com sucesso"
        );
      });
      contRow++;
    });
  }

  async send() {
    this.importResults = [];
    try {
      Excel.run(async context => {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = context.workbook.getSelectedRange().getUsedRange();

        range.load(["values", "columnIndex", "rowIndex", "address"]);

        context.sync().then(() => {
          let selecionList = this.excelAddinService.mapSelection(range.address);
          this.importRangeSpecified(selecionList, sheet, range);
        });
      });
    } catch (error) {
      console.error(error);
      this.modal.open();
    }
  }
}
