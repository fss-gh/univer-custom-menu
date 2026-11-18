import {
    DrawingTypeEnum,
    IResources,
    IRange,
    IStyleData,
    IWorkbookData,
    IWorksheetData,
    LocaleType,
    Nullable,
} from '@univerjs/core';
import { HyperLink, UniverSheet } from './UniverSheet';
// import { ISheetDrawing, SheetDrawingAnchorType } from '@univerjs/sheets-drawing';
import { handleStyle } from './utils';
import { PRINT_AREA_PLUGIN_NAME } from '../../plugin/print-area/types';
import { generateRandomId } from '../common/method';
import { IExcelXmlSheet, IExcelXmlFile, IWorkBookInfo } from './types';
import { ImageSourceType } from './ExcelXmlInterface';
import { handleRanges } from './style';

interface Sheets {
    [sheetId: string]: Partial<IWorksheetData & { hyperLink: HyperLink[] }>;
}
interface ExcelXmlSheetObj {
    [sheetId: string]: Partial<IExcelXmlSheet>;
}
export class UniverWorkBook implements IWorkbookData {
    id!: string;
    rev?: number | undefined;
    name!: string;
    appVersion!: string;
    locale!: LocaleType;
    styles!: Record<string, Nullable<IStyleData>>;
    sheetOrder!: string[];
    sheets!: Sheets;
    resources?: IResources | undefined = [];
    constructor(file: IExcelXmlFile) {
        const { info, sheets, workbook } = file;
        this.id = generateRandomId(6);
        this.name = info.name;
        this.appVersion = info.appversion;
        this.locale = LocaleType.ZH_CN;

        const workSheets: Sheets = {},
            order: string[] = [],
            sheetsObj: ExcelXmlSheetObj = {};
        sheets
            .sort((a, b) => Number(a.order) - Number(b.order))
            .forEach((d) => {
                const sheet = new UniverSheet(d);
                workSheets[sheet.id] = sheet.mode;
                sheetsObj[sheet.id] = d;
                order.push(sheet.id);
            });

        // console.log(workSheets,sheets)
        this.handleHyperLinks(workSheets);
        this.handleImage(workSheets, sheets);
        this.handleChart(workSheets, sheets);
        this.handleNames(workbook);
        this.handleCondition(sheetsObj);
        this.handleVerification(sheetsObj);
        this.handleFilter(sheetsObj);
        this.handleNotes(sheetsObj);
        this.sheetOrder = order;
        this.handlePrintArea(workbook);

        this.sheets = workSheets;
    }

    get mode(): IWorkbookData {
        return {
            id: this.id,
            rev: this.rev,
            name: this.name,
            appVersion: this.appVersion,
            locale: this.locale,
            styles: this.styles,
            sheetOrder: this.sheetOrder,
            sheets: this.sheets,
            resources: this.resources,
        };
    }
    private handleHyperLinks = (workSheets: Sheets) => {
        const hyperLinks: { [key: string]: HyperLink[] } = {};
        for (const key in workSheets) {
            const link = workSheets[key].hyperLink;
            if (!link?.length) continue;
            hyperLinks[key] = link.map((d: any) => {
                let payload = d.payload;
                if (typeof d.payload !== 'string') {
                    payload = '#';
                    const gid = d.payload.gid.replace(/'|"/g, '');
                    const sheetId = Object.values(workSheets).find(
                        (sheet) => sheet.name === gid
                    )?.id;

                    if (gid && sheetId) {
                        payload += `gid=${sheetId}`;
                    }
                    if (gid && sheetId && d.payload.range) payload += '&';
                    if (d.payload.range) payload += `range=${d.payload.range}`;
                }
                return {
                    ...d,
                    payload,
                };
            });
        }
        // console.log(workSheets, hyperLinks)
        this.resources?.push({
            name: 'SHEET_HYPER_LINK_PLUGIN',
            data: JSON.stringify(hyperLinks),
        });
    };
    private handleImage = (workSheets: Sheets, sheets: IExcelXmlSheet[]) => {
        const drawerList: {
            [key: string]: { order: string[]; data: { [key: string]: any } };
        } = {};

        Object.values(workSheets).forEach((sheet) => {
            const images = sheets.find((d) => d.name === sheet.name)?.images;
            if (!images) return;
            const order = Object.keys(images);
            const data: { [key: string]: any } = {};
            order.forEach((key) => {
                const image = images[key];
                if ((sheet.columnCount||0) < image.toCol) {
                    sheet.columnCount = image.toCol + 1;
                }
                if ((sheet.rowCount||0) < image.toRow) {
                    sheet.rowCount = image.toRow + 1;
                }
                const defaultImage = image.default || {};
                const transform = {
                    width: defaultImage.width || image.originWidth || 0,
                    height: defaultImage.height || image.originHeight || 0,
                    scaleX: 1,
                    scaleY: 1,
                    left: defaultImage.left || 0,
                    top: defaultImage.top || 0,
                    angle: 0,
                    skewX: 0,
                    skewY: 0,
                    flipX: false,
                    flipY: false,
                    ...(image.transform || {}),
                };
                // Excel 的 editAs 与 Univer 锚点类型对应：
                // twoCell → Both("1")，oneCell → Position("0")，absolute → None("2")
                const anchorType = image.type === '2' ? '0' : image.type === '3' ? '2' : '1';
                const sheetTransform = {
                    angle: 0,
                    skewX: 0,
                    skewY: 0,
                    flipX: false,
                    flipY: false,
                    from: {
                        column: image.fromCol,
                        columnOffset: image.fromColOff,
                        row: image.fromRow,
                        rowOffset: image.fromRowOff,
                    },
                    to: {
                        column: image.toCol,
                        columnOffset: image.toColOff,
                        row: image.toRow,
                        rowOffset: image.toRowOff,
                    },
                };
                let imageObj: any = {
                    unitId: this.id,
                    subUnitId: sheet.id || '',
                    drawingId: key,
                    transform,
                    sheetTransform,
                    axisAlignSheetTransform: sheetTransform,
                }
                if (image.type === 'chart') {
                    imageObj = {
                        ...imageObj,
                        drawingType: DrawingTypeEnum.DRAWING_CHART,
                        componentKey: 'Chart',
                        data: {
                            ...(image.data || {}),
                            range: `${sheet.name}!${image.data.range}`
                        },
                        allowTransform: true
                    }
                } else {
                    imageObj = {
                        ...imageObj,
                        drawingType: DrawingTypeEnum.DRAWING_IMAGE,
                        imageSourceType: ImageSourceType.BASE64,
                        source: image.src,
                        prstGeom: 'rect',
                        anchorType,
                    }
                }

                data[key] = imageObj;
            });
            drawerList[sheet.id!] = {
                data,
                order,
            };
        });
        this.resources?.push({
            name: 'SHEET_DRAWING_PLUGIN',
            data: JSON.stringify(drawerList),
        });
    };

    private handleChart = (workSheets: Sheets, sheets: IExcelXmlSheet[]) => {
        const chartList: {
            [key: string]: any
        } = {};
        Object.values(workSheets).forEach((sheet) => {
            const charts = sheets.find((d) => d.name === sheet.name)?.charts;
            if (!charts) return;
            charts.forEach((chart) => {
                if (!chartList[sheet.id!]) {
                    chartList[sheet.id!] = []
                }
                chartList[sheet.id!].push({
                    rangeInfo: {
                        isRowDirection: chart.isRowDirection,
                        rangeInfo: {
                            unitId: this.id,
                            subUnitId: sheet.id || '',
                            range: handleRanges(chart.range)[0]
                        }
                    },
                    id: chart.id,
                    chartType: chart.chartType,
                    context: chart.context,
                    style: chart.style,
                    dataAggregation: {}
                })
            });
        })
        // console.log('chartList', chartList)
        this.resources?.push({
            name: 'SHEET_CHART_PLUGIN',
            data: JSON.stringify(chartList),
        });
    }
    private handleNames = (workbook: IWorkBookInfo) => {
        this.resources?.push({
            name: 'SHEET_DEFINED_NAME_PLUGIN',
            data: JSON.stringify(workbook.defineNames),
        });
    };
    /**
     * 把 Excel 的 _xlnm.Print_Area 定义名称转换为打印区域插件资源。
     * Excel 的 localSheetId 是工作表索引，需要映射到 Univer 随机生成的 sheet id。
     */
    private handlePrintArea = (workbook: IWorkBookInfo) => {
        const printAreas: Record<string, { ranges: IRange[] }> = {};
        const definedNames = workbook.defineNames;

        if (definedNames) {
            for (const key in definedNames) {
                const definedName = definedNames[key];
                if (definedName.name !== '_xlnm.Print_Area') continue;

                const localSheetId = Number(definedName.localSheetId);
                const sheetId = this.sheetOrder[localSheetId];
                if (!sheetId) continue;

                const ranges = this.parsePrintAreaRef(definedName.formulaOrRefString);
                if (!ranges.length) continue;

                if (!printAreas[sheetId]) {
                    printAreas[sheetId] = { ranges: [] };
                }
                printAreas[sheetId].ranges.push(...ranges);
            }
        }

        this.resources?.push({
            name: PRINT_AREA_PLUGIN_NAME,
            data: JSON.stringify(printAreas),
        });
    };

    /** 解析 "Sheet1'!$A$1:$D$10" 或 "A1:B2&&C3:D4" / "A1:B2,C3:D4" 这样的 A1 引用，返回一组 IRange */
    private parsePrintAreaRef = (ref: string): IRange[] => {
        // Excel 多个打印区域通常用逗号分隔，ExcelJS 的 pageSetup.printArea 用 && 分隔
        // 每个区域可能都带 sheet 前缀，例如 'Sheet1'!$A$1:$B$2,'Sheet1'!$C$3:$D$4
        const refs = ref.split(/&&|,/);
        const result: IRange[] = [];

        refs.forEach((singleRef) => {
            const rangePart = singleRef.split('!').pop()?.trim();
            if (!rangePart) return;

            const cells = rangePart.split(':');
            const first = this.parseCellAddress(cells[0]);
            if (!first) return;
            const second = cells.length > 1 ? this.parseCellAddress(cells[1]) : first;
            if (!second) return;

            result.push({
                startRow: Math.min(first.row, second.row),
                startColumn: Math.min(first.column, second.column),
                endRow: Math.max(first.row, second.row),
                endColumn: Math.max(first.column, second.column),
            });
        });

        return result;
    };

    private parseCellAddress = (address: string): { row: number; column: number } | null => {
        const match = address.replace(/\$/g, '').trim().match(/^([A-Z]+)(\d+)$/i);
        if (!match) return null;

        const colStr = match[1].toUpperCase();
        let column = 0;
        for (let i = 0; i < colStr.length; i++) {
            column = column * 26 + (colStr.charCodeAt(i) - 64);
        }

        return {
            row: Number(match[2]) - 1,
            column: column - 1,
        };
    };
    private handleCondition = (sheets: ExcelXmlSheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            const condition = sheets[d].conditionalFormatting?.map((d: any) => {
                if (d.rule?.style) {
                    d.rule.style = handleStyle(
                        { v: d.rule.style, r: 0, c: 0 },
                        { value: d.rule?.style?.border, rangeType: '' }
                    );
                }
                return d;
            });
            obj[d] = condition;
        });
        this.resources?.push({
            name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
            data: JSON.stringify(obj),
        });
    };

    private handleVerification = (sheets: ExcelXmlSheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            obj[d] = sheets[d].dataVerificationList;
        });
        this.resources?.push({
            name: 'SHEET_DATA_VALIDATION_PLUGIN',
            data: JSON.stringify(obj),
        });
    };
    private handleFilter = (sheets: ExcelXmlSheetObj) => {
        const obj: any = {};
        Object.keys(sheets).forEach((d) => {
            obj[d] = sheets[d].filter;
        });
        this.resources?.push({
            name: 'SHEET_FILTER_PLUGIN',
            data: JSON.stringify(obj),
        });
    };

    /**
     * Excel 传统批注 → Univer SHEET_NOTE_PLUGIN 资源
     * 资源结构：{ [sheetId]: { [row]: { [col]: ISheetNote } } }
     */
    private handleNotes = (sheets: ExcelXmlSheetObj) => {
        const notes: { [sheetId: string]: { [row: string]: { [col: string]: any } } } = {};
        Object.keys(sheets).forEach((sheetId) => {
            const comments = sheets[sheetId].comments;
            if (!comments || !comments.length) {
                return;
            }

            const sheetNotes: { [row: string]: { [col: string]: any } } = {};
            comments.forEach((comment) => {
                const rowKey = String(comment.row);
                const colKey = String(comment.column);
                if (!sheetNotes[rowKey]) {
                    sheetNotes[rowKey] = {};
                }
                sheetNotes[rowKey][colKey] = {
                    id: `${sheetId}_${comment.row}_${comment.column}`,
                    row: comment.row,
                    col: comment.column,
                    width: comment.width || 200,
                    height: comment.height || 80,
                    note: comment.text,
                    show: true,
                };
            });

            notes[sheetId] = sheetNotes;
        });

        if (Object.keys(notes).length) {
            this.resources?.push({
                name: 'SHEET_NOTE_PLUGIN',
                data: JSON.stringify(notes),
            });
        }
    };
}
