import JSZip from "@progress/jszip-esm";
import { jsonParse } from "../common/method";

interface IFilters {
    blank?: boolean;
    filters?: string[];
}

interface ICustomFilter {
    val: string | number;
    operator?: string;
}

interface ICustomFilters {
    and?: boolean | number;
    customFilters: ICustomFilter[];
}

interface IFilterColumn {
    colId: number;
    filters?: IFilters;
    customFilters?: ICustomFilters;
}

interface IAutoFilter {
    ref?: {
        startRow: number;
        startColumn: number;
        endRow: number;
        endColumn: number;
    };
    filterColumns?: IFilterColumn[];
}

function escapeXml(value: string | number | boolean): string {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function buildFiltersXml(filters: IFilters): string {
    const blank = filters.blank ? ' blank="1"' : "";
    const values = filters.filters || [];

    if (!blank && !values.length) {
        return "";
    }

    if (!values.length) {
        return `<filters${blank}/>`;
    }

    const filterNodes = values
        .map((value) => `<filter val="${escapeXml(value)}"/>`)
        .join("");
    return `<filters${blank}>${filterNodes}</filters>`;
}

function buildCustomFiltersXml(customFilters: ICustomFilters): string {
    const and = customFilters.and ? ' and="1"' : "";
    const customFilterNodes = (customFilters.customFilters || [])
        .map((item) => {
            const operator = item.operator ? ` operator="${escapeXml(item.operator)}"` : "";
            return `<customFilter${operator} val="${escapeXml(item.val)}"/>`;
        })
        .join("");
    return `<customFilters${and}>${customFilterNodes}</customFilters>`;
}

function buildFilterColumnsXml(autoFilter: IAutoFilter): string {
    const ref = autoFilter.ref;
    const filterColumns = autoFilter.filterColumns || [];

    return filterColumns
        .map((column) => {
            const filterXml = column.filters ? buildFiltersXml(column.filters) : "";
            const customFilterXml = column.customFilters ? buildCustomFiltersXml(column.customFilters) : "";
            const children = filterXml + customFilterXml;
            if (!children) {
                return "";
            }

            // Univer 的 colId 是工作表绝对列号，Excel 的 filterColumn.colId 是相对 autoFilter 起始列的偏移。
            const colId = ref ? column.colId - ref.startColumn : column.colId;
            return `<filterColumn colId="${colId}">${children}</filterColumn>`;
        })
        .join("");
}

/**
 * 有筛选条件时，行隐藏状态应该由 autoFilter 重新计算，而不是沿用旧的行 hidden 标记。
 * 否则在 Univer 中修改筛选条件后，旧的行 hidden 会残留，导致导出后清除筛选仍看不到部分行。
 */
function clearHiddenRowsInFilterRange(
    xml: string,
    ref: IAutoFilter["ref"]
): string {
    if (!ref) {
        return xml;
    }

    const startRow = ref.startRow + 1;
    const endRow = ref.endRow + 1;

    return xml.replace(/<row\s+([^>]*?)>/g, (full, attrs: string) => {
        const rowMatch = attrs.match(/\br="(\d+)"/);
        if (!rowMatch) {
            return full;
        }

        const row = Number(rowMatch[1]);
        if (row < startRow || row > endRow) {
            return full;
        }

        const nextAttrs = attrs.replace(/\s*hidden="1"/, "");
        return nextAttrs === attrs ? full : `<row ${nextAttrs}>`;
    });
}

/**
 * ExcelJS 只支持写入 `<autoFilter ref="..."/>`，不支持写入 filterColumn 筛选条件。
 * 这里在 ExcelJS 生成 xlsx 后，把 Univer SHEET_FILTER_PLUGIN 中的筛选条件注入到
 * `xl/worksheets/sheetN.xml` 的 autoFilter 节点。
 */
export async function injectAutoFilterColumns(
    buffer: ArrayBuffer,
    snapshot: any
): Promise<ArrayBuffer> {
    const filterResource = (snapshot?.resources || []).find(
        (item: any) => item.name === "SHEET_FILTER_PLUGIN"
    );
    const filtersBySheetId = filterResource ? jsonParse(filterResource.data) : undefined;
    if (!filtersBySheetId) {
        return buffer;
    }

    const sheetOrder: string[] = snapshot?.sheetOrder || [];
    const targets: Array<{ path: string; autoFilter: IAutoFilter }> = [];

    sheetOrder.forEach((sheetId, index) => {
        const autoFilter = filtersBySheetId[sheetId] as IAutoFilter | undefined;
        if (autoFilter?.filterColumns?.length) {
            targets.push({
                path: `xl/worksheets/sheet${index + 1}.xml`,
                autoFilter,
            });
        }
    });

    if (!targets.length) {
        return buffer;
    }

    const zip = new JSZip();
    await zip.loadAsync(buffer);
    let changed = false;

    for (const target of targets) {
        const file = zip.file(target.path);
        if (!file) {
            continue;
        }

        const xml = await file.async("string");
        const filterColumnsXml = buildFilterColumnsXml(target.autoFilter);
        if (!filterColumnsXml) {
            continue;
        }

        const updatedXml = clearHiddenRowsInFilterRange(
            injectFilterColumnsIntoXml(xml, filterColumnsXml),
            target.autoFilter.ref
        );
        if (updatedXml !== xml) {
            zip.file(target.path, updatedXml);
            changed = true;
        }
    }

    if (!changed) {
        return buffer;
    }

    return zip.generateAsync({
        type: "arraybuffer",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
    });
}

function injectFilterColumnsIntoXml(xml: string, filterColumnsXml: string): string {
    // ExcelJS 默认输出自闭合标签：<autoFilter ref="A1:A4"/>
    const selfClosingMatch = xml.match(/<autoFilter\s+ref="([^"]+)"\s*\/>/);
    if (selfClosingMatch) {
        return xml.replace(
            selfClosingMatch[0],
            `<autoFilter ref="${selfClosingMatch[1]}">${filterColumnsXml}</autoFilter>`
        );
    }

    // 兼容已经带子节点的 autoFilter，插到 </autoFilter> 前。
    const openCloseMatch = xml.match(
        /<autoFilter\s+ref="([^"]+)"[^>]*>([\s\S]*?)<\/autoFilter>/
    );
    if (openCloseMatch) {
        const existingChildren = openCloseMatch[2] || "";
        if (existingChildren.includes("<filterColumn")) {
            return xml;
        }
        return xml.replace(
            openCloseMatch[0],
            `<autoFilter ref="${openCloseMatch[1]}">${existingChildren}${filterColumnsXml}</autoFilter>`
        );
    }

    return xml;
}
