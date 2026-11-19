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
            // 打印区域由 SHEET_PAGE_LAYOUT_PLUGIN 资源通过 pageSetup.printArea 写回，
            // 这里跳过 _xlnm.Print_Area，避免 ExcelJS 重复生成定义名称。
            if (element.name === '_xlnm.Print_Area') continue;
            this.definedNames.add(element.formulaOrRefString, element.name);
        }
    }
}
