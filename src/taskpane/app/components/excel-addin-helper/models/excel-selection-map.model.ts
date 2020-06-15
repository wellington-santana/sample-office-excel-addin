// eslint-disable-next-line no-unused-vars
import ExcelSelectedCell from "./excel-selection-cell.model";

export default class ExcelSelectionMap {
  isEntireColumn: boolean;
  isEntireRow: boolean;
  selectedCells: ExcelSelectedCell[] = [];
  sheetName: string;
  firtsCell: ExcelSelectedCell;
  lastCell: ExcelSelectedCell;
}
