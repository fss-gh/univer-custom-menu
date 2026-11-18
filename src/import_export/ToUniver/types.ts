import { ChartTypeBits } from '../common/constant';

/**
 * 单元格样式中间值（仅在解析过程中使用，最终会转成 Univer IStyleData）
 */
export interface CellFormat {
    fa: string;
    t: string;
    s?: Array<CellStyleValue | InlineString>;
    ci?: any;
}

export interface InlineString {
    ff?: string;
    fc?: string;
    fs?: number;
    cl?: number;
    un?: number;
    bl?: number;
    it?: number;
    va?: number;
    v?: string;
}

export interface CellStyleValue {
    ct?: CellFormat;
    bg?: string;
    ff?: string;
    fc?: string;
    bl?: number;
    it?: number;
    fs?: number;
    cl?: number;
    un?: number;
    vt?: number;
    ht?: number;
    mc?: any;
    tr?: number;
    tb?: number;
    v?: string;
    m?: string;
    rt?: number;
    f?: string;
    va?: number;
    qp?: number;
    ti?: number;
}

export interface BorderStyleValue {
    style: number;
    color: string;
}

export interface BorderCellValue {
    row_index?: number;
    col_index?: number;
    l?: BorderStyleValue;
    r?: BorderStyleValue;
    t?: BorderStyleValue;
    b?: BorderStyleValue;
    tl_br?: BorderStyleValue;
    bl_tr?: BorderStyleValue;
}

export interface BorderCellForImp {
    rangeType: string;
    value?: BorderCellValue;
}

export interface RowAndColumnLen {
    [index: string]: number;
}

export interface RowAndColumnHidden {
    [index: string]: number;
}

export interface CellRange {
    r: number;
    c: number;
    v?: CellStyleValue | string | null;
}

export interface FormulaRefValue {
    t: string;
    ref: string | null;
    si: string;
    fv?: string;
    cellValue: CellRange;
}

export interface FormulaRefList {
    [index: string]: FormulaRefListValue;
}

export interface FormulaRefListValue {
    mainRef?: FormulaRefValue;
    [key: string]: FormulaRefValue | FormulaRefListValue | string | null | undefined;
}

export interface FormulaListItem {
    r: number;
    c: number;
}

export interface FrozenData {
    horizen: number | undefined;
    vertical: number | undefined;
}

/** 图片对象（导入过程中使用，最终转成 Univer SHEET_DRAWING_PLUGIN 资源） */
export interface ImageBorder {
    color: string;
    radius: number;
    style: string;
    width: number;
}

export interface ImageCrop {
    height: number;
    offsetLeft: number;
    offsetTop: number;
    width: number;
}

export interface ImageDefault {
    height: number;
    left: number;
    top: number;
    width: number;
}

export interface ImageObject {
    id?: string;
    src?: string;
    type?: string;

    fromCol?: number;
    fromColOff?: number;
    fromRow?: number;
    fromRowOff?: number;

    toCol?: number;
    toColOff?: number;
    toRow?: number;
    toRowOff?: number;

    originWidth?: number;
    originHeight?: number;

    isFixedPos?: boolean;
    fixedLeft?: number;
    fixedTop?: number;

    border?: ImageBorder;
    crop?: ImageCrop;
    default?: ImageDefault;
    transform?: any;
    data?: any;

    [key: string]: any;
}

export interface ImagesMap {
    [index: string]: ImageObject;
}

export interface ChartImageData {
    border: string;
    background: string;
    range: string;
    chartType: ChartTypeBits;
    isRowDirection: boolean;
}

export interface ChartImageObject extends ImageObject {
    data: ChartImageData;
}

/** 图表对象 */
export interface ChartContext {
    categoryIndex: number;
    seriesIndexes: number[];
}

export interface ChartStyleBase {
    color?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
}

export interface ChartStyleTitle extends ChartStyleBase {
    content: string;
}

export interface ChartStyleTitles {
    titlePosition?: 'top' | 'bottom' | 'left' | 'right' | 'hide';
    [key: string]: ChartStyleTitle | string | undefined | null | any;
}

export interface ChartStyleAxisLabel extends ChartStyleBase {
    axisTitleAlign: 'start' | 'center' | 'end';
    rotate: number;
}

export interface ChartStyleAxis {
    label: ChartStyleAxisLabel;
    reverse: boolean;
    lineVisible: boolean;
    gridLine: {
        visible: boolean;
        width: number | undefined;
        color: string | undefined;
    };
    tick: {
        visible: boolean;
        position: 'inside' | 'outside';
        length?: number;
        lineWidth?: number;
    };
    min: number | undefined;
    max: number | undefined;
}

export interface ChartStyleBorder {
    dashType: 'solid' | 'dotted' | 'dashed';
    width: number;
    opacity: number;
    color: string;
}

export interface ChartStyleSeries {
    border?: ChartStyleBorder;
    label?: {
        visible?: boolean;
    };
    fillOpacity?: number;
}

export interface ChartStyleSerieStyle {
    border?: ChartStyleBorder;
    rightYAxis?: boolean;
    label: {
        visible: boolean;
        position: 'inside' | 'outside';
        contentType: number;
    } & ChartStyleBase;
}

export interface ChartStyleSerieStyles {
    [key: string]: ChartStyleSerieStyle;
}

export interface ChartStyleLegend {
    position: 'top' | 'bottom' | 'left' | 'right' | 'hide';
    selectMode?: 'single' | 'multiple' | 'close';
    label?: ChartStyleBase;
}

export interface ChartStyle {
    titles: ChartStyleTitles;
    runtime: any;
    width: number;
    height: number;
    backgroundColor?: string;
    borderColor?: string;
    xAxis?: ChartStyleAxis;
    yAxis?: ChartStyleAxis;
    allSeriesStyle?: ChartStyleSeries;
    seriesStyleMap?: ChartStyleSerieStyles;
    legend?: ChartStyleLegend;
    gradientFill?: boolean;
}

export interface ChartModel {
    id: string;
    range: string;
    chartType: ChartTypeBits;
    context: ChartContext;
    style: ChartStyle;
    isRowDirection: boolean;
}

/**
 * 以下为 Excel 解析过程中的中间数据结构（原 ILuck 精简后保留在 ToUniver 内部），
 * 仅用于连接 XML 解析层与 Univer 数据组装层，不对外暴露。
 */
import { ExcelXmlConditionFormat } from "./ExcelXmlCondition";
import { IDefinedNames } from "./ExcelXmlDefineName";
import { ExcelXmlFilterFormat } from "./ExcelXmlFilter";
import { ExcelXmlVerificationFormat } from "./ExcelXmlVerification";

export interface IExcelXmlFile {
    info: IExcelXmlFileInfo,//File information, name,password,date,createor etc. 
    sheets: IExcelXmlSheet[],//Sheets, include all sheet data
    workbook: IWorkBookInfo,
}

export interface IWorkBookInfo {
    defineNames?: IDefinedNames
}

export interface IExcelXmlFileInfo {
    name: string,// File name
    creator: string,//Create user
    lastmodifiedby: string,//Edit user
    createdTime: string,
    modifiedTime: string,
    company: string,
    appversion: string,//Excel version, Is it necessary?
}

export interface IExcelXmlSheet {
    name: string,// Sheet name, it will show on sheet bar, must be unique
    color: string,// Sheet color, it will show on sheet bar
    config?: IExcelXmlSheetConfig, // Row height, column width, hidden, and so on
    index: string, //A sheet uniquely identifies, 
    status: string, //If 1 , it means current shown sheet, else means hidden
    order: string, //Order of sheet
    row: number,// Sheet the number of rows, contain blank cell
    column: number, // Sheet the number of columns, contain blank cell
    // visibledatarow:number[],
    // visibledatacolumn:number[],
    excelXmlSelectSave?: IExcelXmlSheetSelection[],//selection defines
    scrollLeft: number,//horizen scroll offset
    scrollTop: number,//verticel scroll offset

    celldata?: IExcelXmlSheetCelldata[],// cells
    dynamicArray?: IExcelXmlSheetDynamicArray[]

    chart?: IExcelXmlSheetChart[],

    isPivotTable: boolean,
    pivotTable?: IExcelXmlSheetPivotTable,

    excelXmlConditionFormatSave?: IExcelXmlSheetConditionFormat[],
    freezen?: IExcelXmlSheetFrozen,

    calcChain?: IExcelXmlSheetCalcChain[],

    zoomRatio: number, //sheet zoom ratio 10%-400%

    showGridLines: string, // show grid lines

    defaultColWidth: number, //cloumn width pixel
    defaultRowHeight: number, //row height pixel

    images: IExcelXmlImages,//image list
    charts: IExcelXmlChart[],//chart list

    dataVerification: IExcelXmlSheetDataVerification;
    hyperlink: IExcelXmlSheetHyperlink, // hyperlinks
    hide: number; // sheet hide

    conditionalFormatting: ExcelXmlConditionFormat[];
    dataVerificationList: ExcelXmlVerificationFormat[];
    filter: ExcelXmlFilterFormat

    comments?: IExcelXmlSheetComment[]; // Excel 传统批注
}

//excelXmlSheet general selection
export interface IExcelXmlSheetSelection {
    row: number[], //selection start row and end row
    column: number[], //selection start column and end column
}

export interface IExcelXmlSheetDynamicArray {
    r: number,//cell row number
    c: number,//cell column number
    f: string | undefined, //formula
    index: string,
    data: number[][],
    refcells: string[][] | undefined
}

export interface IExcelXmlSheetChart {

}

//pivot table interface
export interface IExcelXmlSheetPivotTable {
    pivot_select_save: IExcelXmlSheetSelection,//Pivot table data source range
    pivotDataSheetIndex: string | undefined, //data source sheet index, index is unique id
    column: IExcelXmlSheetPivotTableField[],// column area, include filed 
    row: IExcelXmlSheetPivotTableField[], // row area, include filed 
    filter: IExcelXmlSheetPivotTableField[], // filter area, include filed 
    filterparm: IExcelXmlSheetPivotTablefilterParam,// save param after apply filter
    values: IExcelXmlSheetPivotTableField[],
    showType: string,
    pivotDatas: any[][],
    drawPivotTable: boolean,
    pivotTableBoundary: number[],
}

export interface IExcelXmlSheetPivotTableField {
    index: number,
    name: string,
    fullname: string,
    sumtype: string,
    nameindex: number
}

export interface IExcelXmlSheetPivotTablefilterParam {
    [index: string]: IExcelXmlSheetPivotTablefilterParamItem
}

export interface IExcelXmlSheetPivotTablefilterParamItem {
    caljs: IExcelXmlSheetPivotTablefilterParamItemCaljs,
    rowhidden: IExcelXmlSheetPivotTablefilterParamItemRowhidden,
    selected: IExcelXmlSheetPivotTablefilterParamItemSelected,
}

export interface IExcelXmlSheetPivotTablefilterParamItemCaljs {
    text: string,
    type: string,
    value: string,
    value1: string,
}

export interface IExcelXmlSheetPivotTablefilterParamItemRowhidden {
    [index: number]: number
}

export interface IExcelXmlSheetPivotTablefilterParamItemSelected {
    [index: number]: number
}


export interface IExcelXmlSheetFrozen {
    horizen: number | undefined, //freeze horizen row number
    vertical: number | undefined, //freeze horizen column number
}

export interface IExcelXmlSheetConditionFormat {
    type: string, //Option:defualt,databar,colorGradation,icons,
    cellrange: IExcelXmlSheetSelection[],//Valid range
    format: string[] | IExcelXmlSheetCFDefaultFormat | IExcelXmlSheetCFIconsFormat,//style
    conditionName: string | undefined,//Detailed settings,comparison parameters
    conditionRange: IExcelXmlSheetSelection[],//Detailed settings,comparison range
    conditionValue: any[],//Detailed settings,comparison value
}

export interface IExcelXmlSheetCFDefaultFormat {
    textColor: string | undefined | null,
    cellColor: string | undefined | null
}

export interface IExcelXmlSheetCFIconsFormat {
    len: string | number,
    leftMin: string | number,
    top: string | number,
}

export interface IExcelXmlSheetCalcChain {
    r: number,
    c: number,
    index: string | undefined,
}

export interface IExcelXmlSheetCelldata {
    r: number,//cell row number
    c: number,//cell column number
    v: IExcelXmlSheetCelldataValue | string | null, //cell value
}

export interface IExcelXmlSheetCelldataValue {
    ct: IExcelXmlSheetCellFormat | undefined, //celltype,Cell value format: text, time, etc.
    bg: string | undefined,//background,#fff000	
    ff: string | undefined,//fontfamily,
    fc: string | undefined,//fontcolor
    bl: number | undefined,//Bold
    it: number | undefined,//italic
    fs: number | undefined,//font size	
    cl: number | undefined,//Cancelline, 0 Regular, 1 Cancelline
    un: number | undefined//underline, 0 Regular, 1 underlines, fonts
    vt: number | undefined,//Vertical alignment, 0 middle, 1 up, 2 down
    ht: number | undefined,//Horizontal alignment,0 center, 1 left, 2 right
    mc: IExcelXmlSheetCelldataValueMerge | undefined, //Merge Cells
    tr: number | undefined, //Text rotation, 3 表示竖排文字
    tb: number | undefined, //Text wrap,0 truncation, 1 overflow, 2 word wrap
    v: string | undefined, //Original value	
    m: string | undefined, //Display value	
    rt: number | undefined, //原始 textRotation 角度 0-180（255 已用 tr=3 表示）
    f: string | undefined, //formula
    va: number | undefined,//(Subscript 下标 /Superscript上标 Text)
    qp: number | undefined //quotePrefix, show number as string
    ti: number | undefined //text indent
}


export interface IExcelXmlSheetCellFormat {
    fa: string //Format definition string
    t: string // Cell Type
    s: Array<IExcelXmlSheetCelldataValue | IExcelXmlInlineString> | undefined;
    ci: any // Cell Image
}

export interface IExcelXmlSheetCelldataValueMerge {
    rs?: number, //row of merge cell  length, only main merge cell, every merge cell has only one main mergeCell
    cs?: number, //column of merge cell  length, only main merge cell, every merge cell has only one main mergeCell
    r: number, // main merge cell row Number, other cell link to main cell
    c: number, // main merge cell column Number, other cell link to main cell
}

//ExcelXml sheet config attribute
export interface IExcelXmlSheetConfig {
    merge?: IExcelXmlSheetConfigMerges, //merge handdler
    // _borderInfo?: IMapExcelXmlSheetborderInfoCellForImp, //range border
    borderInfo: IExcelXmlSheetborderInfoCellForImp[],//range border
    rowlen?: IExcelXmlSheetRowAndColumnLen, // every row's height
    columnlen?: IExcelXmlSheetRowAndColumnLen,// every column's width
    rowhidden?: IExcelXmlSheetRowAndColumnHidden,//setting be hidden rows
    colhidden?: IExcelXmlSheetRowAndColumnHidden,//setting be hidden columns

    customHeight: IExcelXmlSheetRowAndColumnHidden,//user operate row height
    customWidth: IExcelXmlSheetRowAndColumnHidden//user operate column width
}

//Merge cells interface
export interface IExcelXmlSheetConfigMerges {
    [firstRange: string]: IExcelXmlSheetConfigMerge // "r_s":{ r,c,rs,cs } format, define a main merge cell 
}
//Merge cell interface
export interface IExcelXmlSheetConfigMerge {
    r: number,
    c: number,
    rs: number,
    cs: number
}

//Border cell interface
export interface IExcelXmlSheetborderInfoCell {
    rangeType: string,
    value: IExcelXmlSheetborderInfoCellValue,
}
export interface IExcelXmlSheetborderInfoCellValue {
    row_index: number,
    col_index: number,
    l: IExcelXmlSheetborderInfoCellValueStyle,
    r: IExcelXmlSheetborderInfoCellValueStyle,
    t: IExcelXmlSheetborderInfoCellValueStyle,
    b: IExcelXmlSheetborderInfoCellValueStyle,
    tl_br: IExcelXmlSheetborderInfoCellValueStyle
    bl_tr: IExcelXmlSheetborderInfoCellValueStyle
}
export interface IExcelXmlSheetborderInfoCellValueStyle {
    "style": number,
    "color": string
}

//border range interface
export interface IExcelXmlSheetborderInfoRange {
    rangeType: string,
    borderType: string,
    style: string,
    color: string,
    range: IExcelXmlSheetSelection[],
}

//Border cell interface for import
export interface IExcelXmlSheetborderInfoCellForImp {
    rangeType: string,
    cells?: string[],
    value: IExcelXmlSheetborderInfoCellValue,
}

//Border cell interface for import
export interface IMapExcelXmlSheetborderInfoCellForImp {
    [index: number]: IExcelXmlSheetborderInfoCellForImp
}

//row length and column length interface
export interface IExcelXmlSheetRowAndColumnLen {
    [index: string]: number
}

//row and column hidden interface
export interface IExcelXmlSheetRowAndColumnHidden {
    [index: string]: number
}

export interface IFormulaSI {
    [index: string]: IFormulaCell
}

export interface IFormulaCell {
    [index: string]: IFormulaCellValue
}

export interface IFormulaCellValue {
    t: string
    ref: string
    si: string
    fv?: string
    cellValue: IExcelXmlSheetCelldata
}

export interface IExcelXmlInlineString {
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



//Image
export interface IExcelXmlImage {
    border: IExcelXmlImageBorder
    crop: IExcelXmlImageCrop
    default: IExcelXmlImageDefault

    fixedLeft: number
    fixedTop: number
    isFixedPos: Boolean
    originHeight: number
    originWidth: number
    src: string
    type: string
    [key: string]: any;
}

export interface IExcelXmlImageBorder {
    color: string
    radius: number
    style: string
    width: number
}

export interface IExcelXmlImageCrop {
    height: number
    offsetLeft: number
    offsetTop: number
    width: number
}

export interface IExcelXmlImageDefault {
    height: number
    left: number
    top: number
    width: number
}

export interface IExcelXmlImages {
    [index: string]: IExcelXmlImage
}


export interface IcellOtherInfo {
    [index: string]: IformulaList
}

export interface IformulaList {
    [index: string]: IformulaListItem | null
}

export interface IformulaListItem {
    r: number,
    c: number
}
// chart
export interface IExcelXmlChartImage {
    border: IExcelXmlImageBorder
    crop: IExcelXmlImageCrop
    default: IExcelXmlImageDefault

    fixedLeft: number
    fixedTop: number
    isFixedPos: Boolean
    originHeight: number
    originWidth: number
    data: IExcelXmlChartImageData
    type: string
    [key: string]: any;
}
export interface IExcelXmlChartImageData {
    border: string,
    background: string,
    range: string,
    chartType: ChartTypeBits,
    isRowDirection: boolean
}
export interface IExcelXmlCharts {
    [index: string]: IExcelXmlChart
}
export interface IExcelXmlChart {
    id: string,
    range: string,
    chartType: ChartTypeBits,
    context: IExcelXmlChartContext,
    style: IExcelXmlChartStyle,
    isRowDirection: boolean,
}
export interface IExcelXmlChartContext {
    categoryIndex: number;
    seriesIndexes: number[];
}
export interface IExcelXmlChartStyle {
    titles: IExcelXmlChartStyleTitles;
    runtime: any;
    width: number;
    height: number;
    backgroundColor?: string;
    borderColor?: string;
    xAxis?: IExcelXmlChartStyleAxis;
    yAxis?: IExcelXmlChartStyleAxis;
    allSeriesStyle?: IExcelXmlChartStyleSeries;
    seriesStyleMap?: IExcelXmlChartStyleSerieStyles;
    legend?: IExcelXmlChartStyleLegend;
    gradientFill?: boolean
}

export interface IExcelXmlChartStyleTitles {
    titlePosition?: 'top' | 'bottom' | 'left' | 'right' | 'hide';
    [key: string]: IExcelXmlChartStyleTitle | string | undefined | null | any,
}

export interface IExcelXmlChartStyleTitle extends IExcelXmlChartStyleBase {
    content: string;
}
export interface IExcelXmlChartStyleBase {
    color?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
}

export interface IExcelXmlChartStyleAxis {
    label: IExcelXmlChartStyleAxisLabel;
    reverse: boolean;
    lineVisible: boolean;
    gridLine: {
        visible: boolean;
        width: number | undefined;
        color: string | undefined;
    },
    tick: {
        visible: boolean;
        position: 'inside' | 'outside',
        length?: number;
        lineWidth?: number;
    },
    min: number | undefined;
    max: number | undefined;
}

export interface IExcelXmlChartStyleAxisLabel extends IExcelXmlChartStyleBase {
    axisTitleAlign: 'start' | 'center' | 'end';
    rotate: number;
}

export interface IExcelXmlChartStyleSeries {
    border?: IExcelXmlChartStyleBorder;
    label?: {
        visible?: boolean;
    }
    fillOpacity?: number;
}
export interface IExcelXmlChartStyleBorder {
    dashType: 'solid' | 'dotted' | 'dashed';
    width: number;
    opacity: number;
    color: string;
}
export interface IExcelXmlChartStyleSerieStyles {
    [key: string]: IExcelXmlChartStyleSerieStyle
}
export interface IExcelXmlChartStyleSerieStyle {
    border?: IExcelXmlChartStyleBorder;
    rightYAxis?: boolean;
    label: {
        visible: boolean;
        position: 'inside' | 'outside',
        contentType: number
    } & IExcelXmlChartStyleBase
}
export interface IExcelXmlChartStyleLegend {
    position: 'top' | 'bottom' | 'left' | 'right' | 'hide';
    selectMode?: 'single' | 'multiple' | 'close';
    label?: IExcelXmlChartStyleBase;
}

// DataVerification
export interface IExcelXmlSheetDataVerification {
    [key: string]: IExcelXmlSheetDataVerificationValue;
}

export interface IExcelXmlSheetDataVerificationValue {
    type: IExcelXmlSheetDataVerificationType;
    type2: string | null;
    value1: string | number | null;
    value2: string | number | null;
    checked: boolean;
    remote: boolean;
    prohibitInput: boolean;
    hintShow: boolean;
    hintText: string | null;
}

export type IExcelXmlSheetDataVerificationType =
    | "dropdown"
    | "checkbox"
    | "number"
    | "number_integer"
    | "number_decimal"
    | "text_content"
    | "text_length"
    | "date"
    | "validity";

export interface IExcelXmlSheetHyperlink {
    [key: string]: IExcelXmlSheetHyperlinkValue;
}

export interface IExcelXmlSheetHyperlinkValue {
    linkAddress: string | null;
    linkTooltip: string;
    linkType: IExcelXmlSheetHyperlinkType;
    display: string;
}

export type IExcelXmlSheetHyperlinkType = "internal" | "external";

/**
 * Excel 传统批注（comments/comment）
 */
export interface IExcelXmlSheetComment {
    /** 批注所在单元格，如 "B2" */
    ref: string;
    /** 0 基行号 */
    row: number;
    /** 0 基列号 */
    column: number;
    /** 批注作者 */
    author?: string;
    /** 批注文本（纯文本，保留换行） */
    text: string;
    /** 批注框宽度（像素，若 XML 可解析） */
    width?: number;
    /** 批注框高度（像素，若 XML 可解析） */
    height?: number;
}
