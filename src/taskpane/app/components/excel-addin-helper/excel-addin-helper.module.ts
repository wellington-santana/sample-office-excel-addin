import { NgModule } from "@angular/core";
import ExcelAddinHelperService from "./excel-addin-helper.service";
import ExcelSelectedCell from "./models/excel-selection-cell.model";
import ExcelSelectionMap from "./models/excel-selection-map.model";

@NgModule({
  providers: [ExcelAddinHelperService],
  exports: [ExcelAddinHelperService, ExcelSelectedCell, ExcelSelectionMap],
  declarations: [ExcelAddinHelperService]
})
export default class ExcelAddinHelperModule {}
