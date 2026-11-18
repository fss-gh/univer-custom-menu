import { ChartTypeBits } from "../common/constant";
import { IExcelXmlFile, IExcelXmlFileInfo, IExcelXmlSheet, IExcelXmlSheetCelldata, IExcelXmlSheetConfig, IExcelXmlSheetCelldataValue, IExcelXmlSheetCelldataValueMerge, IExcelXmlSheetCellFormat, IExcelXmlSheetConfigMerges, IExcelXmlSheetConfigMerge, IExcelXmlSheetborderInfoCellValue, IExcelXmlSheetborderInfoCellValueStyle, IExcelXmlSheetborderInfoCellForImp, IExcelXmlSheetRowAndColumnLen, IExcelXmlSheetRowAndColumnHidden, IExcelXmlSheetSelection, IExcelXmlSheetFrozen, IExcelXmlSheetChart, IExcelXmlSheetPivotTable, IExcelXmlSheetConditionFormat, IExcelXmlSheetCalcChain, IExcelXmlInlineString, IExcelXmlImage, IExcelXmlImageBorder, IExcelXmlImageCrop, IExcelXmlImageDefault, IExcelXmlImages, IExcelXmlSheetHyperlink, IExcelXmlSheetDataVerification, IWorkBookInfo, IExcelXmlCharts, IExcelXmlChart, IExcelXmlChartContext, IExcelXmlChartStyle, IExcelXmlChartImage, IExcelXmlChartImageData } from "./types";
import type { ExcelXmlConditionFormat } from './ExcelXmlCondition'
import { IDefinedNames } from "./ExcelXmlDefineName";
import { ExcelXmlFilterFormat } from "./ExcelXmlFilter";
import { ExcelXmlVerificationFormat } from "./ExcelXmlVerification";



export class ExcelXmlFileBase implements IExcelXmlFile {
    workbook!: IWorkBookInfo;
    info!: IExcelXmlFileInfo
    sheets!: IExcelXmlSheet[]
}
export class WorkBookInfo implements IWorkBookInfo {
    defineNames?: IDefinedNames;
}

export class ExcelXmlSheetBase implements IExcelXmlSheet {
    id!: string
    name!: string
    color!: string
    config!: IExcelXmlSheetConfig
    index!: string
    status!: string
    order!: string
    row!: number
    column!: number
    excelXmlSelectSave!: IExcelXmlSheetSelection[]
    scrollLeft!: number
    scrollTop!: number
    zoomRatio!: number
    showGridLines!: string
    defaultColWidth!: number
    defaultRowHeight!: number

    celldata!: IExcelXmlSheetCelldata[]
    chart!: IExcelXmlSheetChart[]

    isPivotTable!: boolean
    pivotTable!: IExcelXmlSheetPivotTable

    excelXmlConditionFormatSave!: IExcelXmlSheetConditionFormat[]
    freezen!: IExcelXmlSheetFrozen

    calcChain!: IExcelXmlSheetCalcChain[]

    images!: IExcelXmlImages
    charts!: IExcelXmlChart[]

    dataVerification!: IExcelXmlSheetDataVerification;
    hyperlink!: IExcelXmlSheetHyperlink
    hide!: number;

    conditionalFormatting!: ExcelXmlConditionFormat[];
    dataVerificationList!: ExcelXmlVerificationFormat[];
    filter!: ExcelXmlFilterFormat;

}

export class ExcelXmlFileInfo implements IExcelXmlFileInfo {
    name!: string
    creator!: string
    lastmodifiedby!: string
    createdTime!: string
    modifiedTime!: string
    company!: string
    appversion!: string
}

export class ExcelXmlSheetCelldataBase implements IExcelXmlSheetCelldata {
    r!: number
    c!: number
    v!: IExcelXmlSheetCelldataValue | string | null
}

export class ExcelXmlSheetCelldataValue implements IExcelXmlSheetCelldataValue {
    ct: ExcelXmlSheetCellFormat | undefined //celltype,Cell value format: text, time, etc. numfmts
    bg: string | undefined//background,#fff000,	fill
    ff: string | undefined//fontfamily, fonts
    fc: string | undefined//fontcolor fonts
    bl: number | undefined//Bold, fonts
    it: number | undefined//italic, fonts
    fs: number | undefined//font size, fonts
    cl: number | undefined//strike, 0 Regular, 1 strikes, fonts
    un: number | undefined//underline, 0 Regular, 1 underlines, fonts
    vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
    ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
    mc: IExcelXmlSheetCelldataValueMerge | undefined //Merge Cells, mergeCells
    tr: number | undefined //Text rotation,0: 0、3 Vertical text alignment
    tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment
    v: string | undefined //Original value, v
    m: string | undefined //Display value, v
    f: string | undefined //formula, f
    rt: number | undefined //text rotation angle 0-180 alignment
    qp: number | undefined //quotePrefix, show number as string
    va: number | undefined //1sub and 2super and 0none
    ti: number | undefined //text indent
}


export class ExcelXmlSheetCellFormat implements IExcelXmlSheetCellFormat {
    s!: Array<IExcelXmlSheetCelldataValue | IExcelXmlInlineString>;
    ci: any
    fa!: string
    t!: string
}

export class ExcelXmlInlineString implements IExcelXmlInlineString {
    ff: string | undefined //font family
    fc: string | undefined//font color
    fs: number | undefined//font size
    cl: number | undefined//strike
    un: number | undefined//underline
    bl: number | undefined//blod
    it: number | undefined//italic
    va: number | undefined//1sub and 2super and 0none
    v: string | undefined
}

export class ExcelXmlConfig implements IExcelXmlSheetConfig {
    merge!: IExcelXmlSheetConfigMerges
    borderInfo!: IExcelXmlSheetborderInfoCellForImp[]
    // _borderInfo: IMapExcelXmlSheetborderInfoCellForImp
    rowlen!: IExcelXmlSheetRowAndColumnLen
    columnlen!: IExcelXmlSheetRowAndColumnLen
    rowhidden!: IExcelXmlSheetRowAndColumnHidden
    colhidden!: IExcelXmlSheetRowAndColumnHidden

    customHeight!: IExcelXmlSheetRowAndColumnHidden
    customWidth!: IExcelXmlSheetRowAndColumnHidden
}

export class ExcelXmlSheetborderInfoCellForImp implements IExcelXmlSheetborderInfoCellForImp {
    rangeType!: string
    // cells:string[]
    value!: IExcelXmlSheetborderInfoCellValue
}

export class ExcelXmlSheetborderInfoCellValue implements IExcelXmlSheetborderInfoCellValue {
    row_index!: number
    col_index!: number
    l!: IExcelXmlSheetborderInfoCellValueStyle
    r!: IExcelXmlSheetborderInfoCellValueStyle
    t!: IExcelXmlSheetborderInfoCellValueStyle
    b!: IExcelXmlSheetborderInfoCellValueStyle
    tl_br!: IExcelXmlSheetborderInfoCellValueStyle
    bl_tr!: IExcelXmlSheetborderInfoCellValueStyle
}

export class ExcelXmlSheetborderInfoCellValueStyle implements IExcelXmlSheetborderInfoCellValueStyle {
    "style": number
    "color": string
}

export class ExcelXmlSheetConfigMerge implements IExcelXmlSheetConfigMerge {
    r!: number
    c!: number
    rs!: number
    cs!: number
}

export class ExcelXmlSheetCalcChain implements IExcelXmlSheetCalcChain {
    r!: number
    c!: number
    index!: string | undefined
}


export class ExcelXmlImageBase implements IExcelXmlImage {
    border!: IExcelXmlImageBorder
    crop!: IExcelXmlImageCrop
    default!: IExcelXmlImageDefault

    fixedLeft!: number
    fixedTop!: number
    isFixedPos!: Boolean
    originHeight!: number
    originWidth!: number
    src!: string
    type!: string
}

export class ExcelXmlChartImageBase implements IExcelXmlChartImage {
    border!: IExcelXmlImageBorder
    crop!: IExcelXmlImageCrop
    default!: IExcelXmlImageDefault
    transform!: IExcelXmlImageDefault

    fixedLeft!: number
    fixedTop!: number
    isFixedPos!: Boolean
    originHeight!: number
    originWidth!: number
    data!: IExcelXmlChartImageData
    type!: string
}

export class ExcelXmlChart implements IExcelXmlChart {
    id!: string;
    range!: string;
    chartType!: ChartTypeBits;
    context!: IExcelXmlChartContext;
    style!: IExcelXmlChartStyle;
    isRowDirection!: boolean;
}