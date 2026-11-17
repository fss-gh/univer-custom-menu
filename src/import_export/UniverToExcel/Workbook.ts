import exceljs from "@zwight/exceljs";
import { jsonParse } from "../common/method";
import { ExcelWorkSheet } from "./WorkSheet";

const Workbook = exceljs.Workbook;
export class WorkBook extends Workbook {
    constructor(snapshot: any) {
        super()
        this.init(snapshot);
    }

    private init(snapshot: any) {
        // this.properties.date1904 = true;
        this.calcProperties.fullCalcOnLoad = true;
        this.setDefineNames(snapshot.resources);
        ExcelWorkSheet(this, snapshot)
    }

    private setDefineNames(resources: any[]) {
        const definedNames = jsonParse(resources.find(d => d.name === 'SHEET_DEFINED_NAME_PLUGIN')?.data);
        for (const key in definedNames) {
            const element = definedNames[key]
            this.definedNames.add(element.formulaOrRefString, element.name);
        }
    }
}
