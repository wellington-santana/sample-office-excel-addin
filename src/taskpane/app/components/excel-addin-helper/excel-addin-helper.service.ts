import { Injectable } from "@angular/core";
import ExcelSelectionMap from "./models/excel-selection-map.model";
import ExcelSelectedCell from "./models/excel-selection-cell.model";

@Injectable()
export default class ExcelAddinHelperService {
  private _patterValidExcelSelection = new RegExp("\\w+\\![A-Z]+");
  private _patterSheetName = new RegExp("^(\\w+)\\!");
  private _patternEntireSelection = new RegExp("([A-Z]+):([A-Z]+)|(\\d+):(\\d+)$");
  private _patternSpecificCellSelection = new RegExp("([A-Z]+\\d+:[A-Z]+\\d+)|([A-Z]+\\d+)");

  public mapSelection(excelSelecion: string): ExcelSelectionMap {
    if (!this._patterValidExcelSelection.test(excelSelecion)) {
      throw Error("Valor e enviado para mapeamento não é válido");
    }

    if (this._patternEntireSelection.test(excelSelecion)) {
      return this.mapEntireSelection(excelSelecion);
    } else if (this._patternSpecificCellSelection.test(excelSelecion)) {
      return this.mapSpecificSelection(excelSelecion);
    }

    throw Error("Valor e enviado de seleção para mapeamento não é válido");
  }

  private mapEntireSelection(excelSelecion: string): ExcelSelectionMap {
    let matches = this._patternEntireSelection.exec(excelSelecion);
    let selectionMap = new ExcelSelectionMap();
    let firstCell = new ExcelSelectedCell();
    let lastCell = new ExcelSelectedCell();
    let param1 = matches[1].split(":")[0];
    let param2 = matches[1].split(":")[1];

    if (Number(param1)) {
      selectionMap.isEntireRow = true;
      firstCell.rowName = param1;
      firstCell.rowIndex = Number(param1) - 1;
      lastCell.rowName = param2;
      lastCell.rowIndex = Number(param2) - 1;
    } else {
      selectionMap.isEntireColumn = true;
      firstCell.columnName = param1;
      firstCell.columnIndex = this.transformColumnLetterToNumber(param1);
      lastCell.columnName = param2;
      lastCell.columnIndex = this.transformColumnLetterToNumber(param2);
    }

    selectionMap.firtsCell = firstCell;
    selectionMap.lastCell = lastCell;
    selectionMap.sheetName = this.getWorksheetName(excelSelecion);

    return selectionMap;
  }

  private mapSpecificSelection(excelSelecion: string): ExcelSelectionMap {
    let matches = this._patternSpecificCellSelection.exec(excelSelecion);
    let selectionMap = new ExcelSelectionMap();
    let selectedCells: ExcelSelectedCell[] = [];

    matches.splice(0, 1).forEach(selection => {
      if (!selection) {
        return;
      }
      let isRangeSelection = selection.indexOf(":") > -1;

      if (isRangeSelection) {
        selectedCells = selectedCells.concat(this.transformRangeSelectedCell(selection));
      } else {
        selectedCells.push(this.transformSelectedCell(selection));
      }
    });

    selectionMap.selectedCells = selectedCells;
    selectionMap.firtsCell = selectedCells[0];
    selectionMap.lastCell = selectedCells[selectedCells.length - 1];
    selectionMap.sheetName = this.getWorksheetName(excelSelecion);

    return selectionMap;
  }

  private transformRangeSelectedCell(rangeSelectedCellString: string): ExcelSelectedCell[] {
    let selectionList: ExcelSelectedCell[] = [];
    let fistSelection = this.transformSelectedCell(rangeSelectedCellString.split(":")[0]);
    let lastSelection = this.transformSelectedCell(rangeSelectedCellString.split(":")[1]);

    for (let columnIndex = fistSelection.columnIndex; columnIndex <= lastSelection.columnIndex; columnIndex++) {
      for (let rowIndex = fistSelection.rowIndex; rowIndex <= lastSelection.rowIndex; rowIndex++) {
        selectionList.push(this.transformSelectedCell(`${this.transformColumnNumberToLetter(columnIndex + 1)}${rowIndex + 1}`));
      }
    }

    return selectionList;
  }

  private transformSelectedCell(selectedCellString: string): ExcelSelectedCell {
    let selectedCell = new ExcelSelectedCell();
    let matches = new RegExp("([A-Z]+)(\\d+)").exec(selectedCellString);

    selectedCell.columnName = matches[1];
    selectedCell.columnIndex = this.transformColumnLetterToNumber(matches[1]) - 1;
    selectedCell.rowName = matches[2];
    selectedCell.rowIndex = Number(matches[2]) - 1;

    return selectedCell;
  }

  //https://stackoverflow.com/questions/9905533/convert-excel-column-alphabet-e-g-aa-to-number-e-g-25
  private transformColumnLetterToNumber(columnLetter: string): number {
    var base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
      i,
      j,
      result = 0;

    for (i = 0, j = columnLetter.length - 1; i < columnLetter.length; i += 1, j -= 1) {
      result += Math.pow(base.length, j) * (base.indexOf(columnLetter[i]) + 1);
    }

    return result;
  }

  //https://stackoverflow.com/questions/9905533/convert-excel-column-alphabet-e-g-aa-to-number-e-g-25
  private transformColumnNumberToLetter(columnIndex: number): string {
    for (var ret = "", a = 1, b = 26; (columnIndex -= a) >= 0; a = b, b *= 26) {
      ret = String.fromCharCode((columnIndex % b) / a + 65) + ret;
    }
    return ret;
  }

  private getWorksheetName(excelSelecion: string): string {
    return this._patterSheetName.exec(excelSelecion)[1];
  }
}
