import { Workbook, Worksheet, WorksheetViewCommon, WorksheetViewFrozen } from "@zwight/exceljs";
import { convertSheetIdToName, heightConvert, hex2argb, wdithConvert } from "./util";
import { PAGE_LAYOUT_PLUGIN_NAME } from "../../plugin/page-layout/types";
import { cellStyle, fontConvert } from "./CellStyle";
import { getImageExtension, jsonParse, removeEmptyAttr } from "../common/method";
import { Resource } from "./Resource";

export class ViewCommon implements WorksheetViewCommon {
    rightToLeft!: boolean;
    activeCell!: string;
    showRuler!: boolean;
    showRowColHeaders!: boolean;
    showGridLines!: boolean;
    zoomScale!: number;
    zoomScaleNormal!: number;
}
export class FrozenView implements WorksheetViewFrozen {
    state!: 'frozen';
    style?: "pageBreakPreview";
    xSplit?: number;
    ySplit?: number;
    topLeftCell?: string;
}
export function ExcelWorkSheet(workbook: Workbook, snapshot: any) {
    const { sheetOrder, sheets, styles, resources } = snapshot;
    sheetOrder.forEach((sheetId: string) => {
        const sheet = sheets[sheetId];
        const {
            id,
            name,
            tabColor,
            defaultRowHeight,
            defaultColumnWidth,
            hidden,
            rightToLeft,
            showGridlines,
            freeze,
            mergeData
        } = sheet;
        const commonView = new ViewCommon();
        commonView.rightToLeft = rightToLeft === 1;
        commonView.showGridLines = showGridlines === 1;
        const frozenView = new FrozenView();
        if (freeze.xSplit > 0 || freeze.ySplit > 0) {
            frozenView.state = 'frozen';
            frozenView.xSplit = freeze.xSplit;
            frozenView.ySplit = freeze.ySplit;
        }
        const views = Object.assign(commonView, frozenView)

        const defaultColWidth = wdithConvert(defaultColumnWidth);
        const defaultRowHeightR = heightConvert(defaultRowHeight);
        const worksheet = workbook.addWorksheet(name, {
            views: [views],
            state: hidden === 1 ? 'hidden' : 'visible',
            properties: {
                tabColor: tabColor ? { argb: hex2argb(tabColor) } : undefined,
                defaultColWidth: defaultColWidth,
                defaultRowHeight: defaultRowHeightR,
                dyDescent: 0
            }
        })
        setColumns(worksheet, sheet.columnData, defaultColWidth);
        setRows(worksheet, sheet.rowData, defaultRowHeightR)
        setCell(worksheet, sheet, styles, snapshot, workbook);
        setMerges(worksheet, mergeData);
        setPageLayout(worksheet, id, resources);

        new Resource(id, workbook, worksheet, resources);
    });
}


function setMerges(worksheet: Worksheet, mergeData: any[]) {
    mergeData.forEach(d => {
        worksheet.mergeCells(d.startRow + 1, d.startColumn + 1, d.endRow + 1, d.endColumn + 1)
    })
}

/**
 * 从 SHEET_PAGE_LAYOUT_PLUGIN 资源写回 Excel 的页面布局：
 * - 打印区域 -> pageSetup.printArea（ExcelJS 会转成 _xlnm.Print_Area 定义名称）
 * - 页边距 -> pageMargins
 * - 纸张方向/大小 -> pageSetup
 * - 打印缩放 -> pageSetUpPr.fitToPage + pageSetup.fitToWidth/fitToHeight/scale
 * - 水平分页符 -> rowBreaks（ExcelJS 不支持 colBreaks，列分页符在 sanitizeExportedXml 中注入）
 */
function setPageLayout(worksheet: Worksheet, sheetId: string, resources: any[]) {
    const pageLayouts = jsonParse(resources.find((d: any) => d.name === PAGE_LAYOUT_PLUGIN_NAME)?.data);
    const layout = pageLayouts?.[sheetId];
    if (!layout) return;

    if (layout.printArea?.length) {
        worksheet.pageSetup.printArea = layout.printArea
            .map((range: any) => `${toExcelAddress(range.startRow, range.startColumn)}:${toExcelAddress(range.endRow, range.endColumn)}`)
            .join('&&');
    }

    if (layout.margins) {
        worksheet.pageSetup.margins = { ...layout.margins };
    }

    if (layout.pageSetup) {
        worksheet.pageSetup.orientation = layout.pageSetup.orientation;
        worksheet.pageSetup.paperSize = layout.pageSetup.paperSize as any;

        const scaleMode = layout.pageSetup.scaleMode ?? 'none';
        if (scaleMode === 'none') {
            worksheet.pageSetup.fitToPage = false;
            worksheet.pageSetup.scale = 100;
        } else {
            worksheet.pageSetup.fitToPage = true;
            // Excel 用 0 表示“自动”：只限宽时高度为 0，只限高时宽度为 0
            worksheet.pageSetup.fitToWidth = scaleMode === 'fitRows' ? 0 : 1;
            worksheet.pageSetup.fitToHeight = scaleMode === 'fitColumns' ? 0 : 1;
        }
    }

    if (layout.rowBreaks?.length) {
        (worksheet as any).rowBreaks = layout.rowBreaks.map((id: number) => ({
            // Univer 存 0-based 起始行，Excel brk id 从 1 开始
            id: id + 1,
            max: 16383,
            min: 0,
            man: 1,
        }));
    }
}

function setCell(worksheet: Worksheet, sheet: any, styles: any, snapshot: any, workbook: Workbook) {
    const { resources, sheets } = snapshot;
    const { cellData, id } = sheet;
    const formulaMap = getFormulaMap(cellData);
    for (const rowid in cellData) {
        const row = cellData[rowid];
        for (const columnid in row) {
            const cell = row[columnid];
            if (!cell) continue;
            // console.log(rowid + 1, columnid + 1)
            const target = worksheet.getCell(Number(rowid) + 1, Number(columnid) + 1)

            target.value = handleValue(cell, {
                resources,
                sheetId: id,
                rowId: rowid,
                columnId: columnid,
                sheets,
                formulaMap
            }, workbook);

            let originStyle = cell.s;
            if (typeof cell.s === 'string') {
                originStyle = styles[cell.s]
            }
            const style = removeEmptyAttr(cellStyle(originStyle, originStyle?.n?.pattern))
            Object.assign(target, style)
            // console.log(target)
        }
    }
}

function getHyperLink(cellSource: any) {
    const { resources, sheetId, rowId, columnId } = cellSource;
    const hyperlinks = jsonParse(resources.find((d: any) => d.name === 'SHEET_HYPER_LINK_PLUGIN')?.data);
    const list = hyperlinks?.[sheetId] || [];
    const hyperlink = list.find((d: any) => d.row === Number(rowId) && d.column === Number(columnId));
    return hyperlink
}

function handleHyperLink(hyperlink: any, sheets: any) {
    let hyperlinks;
    if (hyperlink) {
        const { payload } = hyperlink;
        let link = '';
        let model = '';
        if (payload.includes('#gid=') || payload.includes('range=')) {
            const str = payload.replace('#', '');
            const arr = str.split('&');
            link += '';
            if (arr.length === 1 && arr[0].includes('range=')) {
                link += arr[0].replace('range=')
            }
            if (arr.length === 2) {
                link += `\'${convertSheetIdToName(sheets, arr[0].replace('gid=', ''))}\'`;
                link += `!${arr[1].replace('range=', '')}`
            }
        } else {
            link = payload
            model = 'External'
        }

        if (link) hyperlinks = {
            hyperlink: link,
            hyperlinkModel: model
        }
    }
    return hyperlinks
}

function getFormulaMap(cellData: any) {
    const map: Record<string, { row: number; column: number; formula: string }> = {};
    for (const rowid in cellData) {
        const row = cellData[rowid];
        for (const columnid in row) {
            const cell = row[columnid];
            if (cell?.si && cell?.f) {
                map[cell.si] = {
                    row: Number(rowid),
                    column: Number(columnid),
                    formula: String(cell.f).replace(/^=/, ''),
                };
            }
        }
    }
    return map;
}

function toExcelAddress(row: number, column: number) {
    let col = '';
    let n = column;
    while (n >= 0) {
        col = String.fromCharCode((n % 26) + 65) + col;
        n = Math.floor(n / 26) - 1;
    }
    return `${col}${row + 1}`;
}

function handleValue(cell: any, cellSource: any, workbook: Workbook) {
    const { sheets } = cellSource
    const hyperlink = getHyperLink(cellSource)
    const hyperlinks = handleHyperLink(hyperlink, sheets)
    let value;
    if (cell.p) {
        const body = cell.p?.body;
        if (cell.p.drawingsOrder?.length) {
            const image = cell.p.drawings[cell.p.drawingsOrder[0]];
            const { id, value: imgId } = workbook.addCellImage({
                base64: image.source,
                extension: getImageExtension(image.source),
                descr: image.description,
                ext: {
                    width: image.transform.width,
                    height: image.transform.height,
                }
            })
            value = { id, cellImageId: imgId, ...(hyperlinks || {}) }
            return value
        } else {
            value = {
                richText: body?.textRuns.map((d: any) => {
                    return {
                        text: body.dataStream.substring(d.st, d.ed),
                        font: fontConvert(d.ts)
                    }
                })
            }
        }
    } else if (cell.f) {
        // Univer 中公式存放在 f 字段（例如 =SUM(A1:A2)），ExcelJS 需要去掉开头的 =
        const formula = String(cell.f).replace(/^=/, '');
        value = { formula, result: cell.v }
    } else if (cell.si) {
        // 共享公式：Univer 子单元格可能只有 si，没有完整 f，需要指向主公式单元格
        const { formulaMap } = cellSource;
        const master = formulaMap?.[cell.si];
        if (master) {
            value = {
                sharedFormula: toExcelAddress(master.row, master.column),
                result: cell.v,
            };
        } else {
            // 兼容旧数据：si 作为共享公式的兜底
            value = { formula: String(cell.si).replace(/^=/, ''), result: cell.v }
        }
    } else {
        value = cell.v
    }
    if (hyperlinks) {
        const text = value?.richText?.map?.((d: any) => d.text)?.join('') || value?.result || value;
        value = {
            text: text,
            ...hyperlinks
        }
    }
    return value
}

function setColumns(worksheet: Worksheet, columnData: any = {}, defaultColumnWidth: number) {
    for (const key in columnData) {
        if (Object.prototype.hasOwnProperty.call(columnData, key)) {
            const element = columnData[key];
            const column = worksheet.getColumn(Number(key) + 1)
            column.width = element.w ? wdithConvert(element.w) : defaultColumnWidth;
            column.hidden = element.hd === 1;
        }
    }
}

function setRows(worksheet: Worksheet, rowData: any = {}, defaultRowHeight: number) {
    for (const key in rowData) {
        if (Object.prototype.hasOwnProperty.call(rowData, key)) {
            const element = rowData[key];
            const row = worksheet.getRow(Number(key) + 1)
            row.height = element.h ? heightConvert(element.h) : defaultRowHeight;
            row.hidden = element.hd === 1;
        }
    }
}