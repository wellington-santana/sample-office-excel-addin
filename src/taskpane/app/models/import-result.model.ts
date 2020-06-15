export default class ImportResult{
    rowIndex: number;
    rowName: string;
    columnIndex: number;
    columnName: string;
    isSuccess: boolean;
    message: string;

    constructor(columnName: string, columnIndex: number, rowIndex: number, isSuccess: boolean, message: string){
        this.columnName = columnName;
        this.columnIndex = columnIndex;
        this.rowIndex = rowIndex;
        this.isSuccess = isSuccess;
        this.message = message;
        this.rowName = (++rowIndex).toString();
    }
}