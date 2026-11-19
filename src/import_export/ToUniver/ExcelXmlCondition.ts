/**:
* @description: 返回的condition格式为univer使用的格式，但是规则配置中的style是ExcelXml的格式，
* 后面考虑新增配置用来输出univer所需的数据结构
* @author: Created by zwight on 2024-09-20 16:13:34
*/
import { escapeCharacter, generateRandomIndex, getXmlAttibute, str2num } from "../common/method";
import { IattributeList } from "../ICommon";
import { CellFormat } from "./types";
import { Element, getColor, IStyleCollections, ReadXml } from "./ReadXml";
import { getBackgroundByFill, getFontStyle, handleBorder, handleRanges, IRange } from "./style";

/** Excel 默认图标集；源文件省略 iconSet 属性时使用 */
export const DEFAULT_ICON_SET = '3TrafficLights1';


export interface ConditionFormationRule {
    type: string,
    subType?: string,
    style?: {},
    value?: string,
    isShowValue?: boolean
}
export interface ExcelXmlConditionFormat<C = any> {
    ranges: IRange[],
    cfId: string,
    stopIfTrue: boolean,
    order: number,
    rule: C
}

export enum CFTextOperator {
    beginsWith = 'beginsWith',
    endsWith = 'endsWith',
    containsText = 'containsText',
    notContainsText = 'notContainsText',
    equal = 'equal',
    notEqual = 'notEqual',
    containsBlanks = 'containsBlanks',
    notContainsBlanks = 'notContainsBlanks',
    containsErrors = 'containsErrors',
    notContainsErrors = 'notContainsErrors',
}
export enum CFTimePeriodOperator {
    today = 'today',
    yesterday = 'yesterday',
    tomorrow = 'tomorrow',
    last7Days = 'last7Days',
    thisMonth = 'thisMonth',
    lastMonth = 'lastMonth',
    nextMonth = 'nextMonth',
    thisWeek = 'thisWeek',
    lastWeek = 'lastWeek',
    nextWeek = 'nextWeek',
}
export enum CFNumberOperator {
    greaterThan = 'greaterThan',
    greaterThanOrEqual = 'greaterThanOrEqual',
    lessThan = 'lessThan',
    lessThanOrEqual = 'lessThanOrEqual',
    notBetween = 'notBetween',
    between = 'between',
    equal = 'equal',
    notEqual = 'notEqual',
}

export enum CFRuleType {
    highlightCell = 'highlightCell',
    dataBar = 'dataBar',
    colorScale = 'colorScale',
    iconSet = 'iconSet',
}
export enum CFSubRuleType {
    uniqueValues = 'uniqueValues',
    duplicateValues = 'duplicateValues',
    rank = 'rank',
    text = 'text',
    timePeriod = 'timePeriod',
    number = 'number',
    average = 'average',
    formula = 'formula',
}

interface ConditionElement extends Element {
    parentAttribute: IattributeList
    extLst: Element,
    isExtLst: boolean
}

function cellRefToIndex(ref: string): { row: number; col: number } | null {
    const match = ref.match(/([A-Z]+)(\d+)/);
    if (!match) return null;

    let col = 0;
    for (const ch of match[1]) {
        col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    return {
        row: Number(match[2]) - 1,
        col: col - 1,
    };
}

function getDataBarActualMinMax(
    ranges: IRange[],
    readXml: ReadXml,
    sheetFile?: string
): { min: number | null; max: number | null } {
    if (!sheetFile) {
        return { min: null, max: null };
    }

    const cells = readXml.getElementsByTagName("sheetData/c", sheetFile, true);
    let min: number | null = null;
    let max: number | null = null;

    cells.forEach((cell) => {
        const cellType = cell.attributeList?.t;
        // 只统计数值单元格
        if (cellType && cellType !== "n") return;

        const ref = cellRefToIndex(cell.attributeList?.r || "");
        if (!ref) return;

        const inRange = ranges.some((range) =>
            ref.row >= range.startRow &&
            ref.row <= range.endRow &&
            ref.col >= range.startColumn &&
            ref.col <= range.endColumn
        );
        if (!inRange) return;

        const vElement = readXml.getElementsByTagName("v", cell.value, false)?.[0];
        const value = str2num(vElement?.value);
        if (typeof value !== "number" || Number.isNaN(value)) return;

        if (min === null || value < min) min = value;
        if (max === null || value > max) max = value;
    });

    return { min, max };
}

/**
 * Univer 的数据条渲染器在“自动 min/max”边界下，最小值单元格会被计算成 0 长度而返回 undefined，
 * 导致该单元格不显示任何数据条（例如 Excel 中 5 和 6 的数据条，5 会显示短条，而 Univer 中 5 完全消失）。
 *
 * 这里把 Excel 的“自动 min/max”边界统一转成 Univer 的数值类型（num）：
 * - 全正数：min 转成 0，max 保持实际最大值，使数据条从 0 开始显示（与 Excel 数值类型效果一致）；
 * - 全负数：max 向上外扩 10%，让最接近 0 的单元格也有非零长度；
 * - 两侧都改成 num，同时避免导出时 num + autoMin/autoMax 混用导致 Excel 报错。
 */
function adjustDataBarAutoToNum(
    config: any,
    ranges: IRange[],
    readXml: ReadXml,
    sheetFile?: string
) {
    const hasAuto = config.min?.type === "min" || config.max?.type === "max";
    if (!hasAuto) return;

    const { min: actualMin, max: actualMax } = getDataBarActualMinMax(ranges, readXml, sheetFile);

    // 正数范围自动最小值统一转成 0，这样与 Excel 中数值类型数据条从 0 开始的效果一致；
    // 负数范围保持实际最小值，避免破坏负数数据条。
    const hasActual = actualMin !== null && actualMax !== null && actualMax > actualMin;
    const rangeLength = hasActual ? actualMax - actualMin : 100;
    const padding = rangeLength * 0.1;

    if (config.min?.type === "min") {
        const value = hasActual && actualMin >= 0
            ? 0
            : (hasActual ? actualMin : 0);
        config.min = { type: "num", value };
    }

    if (config.max?.type === "max") {
        const value = hasActual
            ? (actualMax <= 0 ? actualMax + padding : actualMax)
            : 100;
        config.max = { type: "num", value };
    }
}

export class ExcelXmlCondition implements ExcelXmlConditionFormat {
    ranges!: IRange[];
    cfId: string;
    stopIfTrue: boolean = false;
    order: number;
    rule: any;
    constructor(ele: ConditionElement, readXml: ReadXml, styles: IStyleCollections, sheetFile?: string) {
        const { attributeList, parentAttribute } = ele;
        if (parentAttribute?.sqref) this.ranges = handleRanges(parentAttribute.sqref);
        this.order = Number(getXmlAttibute(attributeList, 'priority', '1'));
        this.cfId = generateRandomIndex('condition');
        this.stopIfTrue = getXmlAttibute(attributeList, 'stopIfTrue', '0') === '1' ? true : false;
        this.handleRules(ele, readXml, styles, sheetFile)
    }

    private handleRules = (ele: ConditionElement, readXml: ReadXml, styles: IStyleCollections, sheetFile?: string) => {
        const { attributeList, value, extLst, isExtLst } = ele
        const type = getXmlAttibute(attributeList, 'type', 'expression')
        const rule: any = {
            type: CFRuleType.highlightCell
        }
        const operator = getXmlAttibute(attributeList, 'operator', '')
        const rank = getXmlAttibute(attributeList, 'rank', '');

        const formula = readXml.getElementsByTagName("formula", value, false);
        const aboveAverage = getXmlAttibute(attributeList, 'aboveAverage', '');

        if (operator) rule.operator = operator;
        if (formula[0]?.value || formula[0]?.value == '0') rule.value = str2num(formula[0]?.value);
        switch (type) {
            case 'expression':
                rule.subType = CFSubRuleType.formula;
                break;
            case 'cellIs':
                rule.subType = CFSubRuleType.number;
                break;
            case 'top10':
                rule.subType = CFSubRuleType.rank;
                const percent = getXmlAttibute(attributeList, 'percent', '0');
                const bottom = getXmlAttibute(attributeList, 'bottom', '0');
                if (rank) rule.value = str2num(rank);
                rule.isBottom = bottom === '1' ? true : false;
                rule.isPercent = percent === '1' ? true : false;
                break;
            case 'aboveAverage':
                rule.subType = CFSubRuleType.average;
                rule.operator = rule.operator || CFNumberOperator.lessThan
                break;
            case 'timePeriod':
                rule.subType = CFSubRuleType.timePeriod;
                rule.operator = getXmlAttibute(attributeList, 'timePeriod', null);
                break;
            case 'duplicateValues':
                rule.subType = CFSubRuleType.duplicateValues;
                break;
            case 'containsText':
                rule.subType = CFSubRuleType.text;
                rule.operator = 'containsText'
                rule.value = getXmlAttibute(attributeList, 'text', '');
                break;
            case 'colorScale':
                const cfvo = readXml.getElementsByTagName("colorScale/cfvo", value, false)
                const color = readXml.getElementsByTagName("colorScale/color", value, false)

                rule.type = CFRuleType.colorScale;
                rule.config = cfvo.map((d, index) => {
                    const type = getXmlAttibute(d.attributeList, 'type', '');
                    const value = getXmlAttibute(d.attributeList, 'val', null);
                    const colorVal = color[index] ? getColor(color[index], styles) : undefined
                    return {
                        index: 0,
                        color: colorVal,
                        value: {
                            type,
                            value: str2num(value)
                        }
                    }
                })
                break;
            case 'dataBar':
                rule.type = CFRuleType.dataBar;

                const dataBar = readXml.getElementsByTagName("dataBar", value, false)?.[0];
                const barCfvo = readXml.getElementsByTagName("cfvo", dataBar.value, false);
                const barColor = readXml.getElementsByTagName("color", dataBar.value, false);
                const isShowValue = getXmlAttibute(dataBar.attributeList, 'showValue', '1');
                rule.isShowValue = isShowValue === '1';

                let positiveColor = barColor[0] ? getColor(barColor[0], styles) : undefined;
                let isGradient = true;
                let nativeColor: string | null | undefined = '';
                if (extLst) {
                    const x14DataBar = readXml.getElementsByTagName("x14:dataBar", extLst.value, false)[0];
                    const gradient = getXmlAttibute(x14DataBar.attributeList, 'gradient', null);
                    const negativeFillColor = readXml.getElementsByTagName("x14:dataBar/x14:negativeFillColor", extLst.value, false)?.[0];
                    nativeColor = negativeFillColor ? getColor(negativeFillColor, styles) : undefined;
                    isGradient = gradient !== '0';
                }
                rule.config = {
                    min: {
                        type: getXmlAttibute(barCfvo[0]?.attributeList, 'type', 'min'),
                        value: str2num(getXmlAttibute(barCfvo[0]?.attributeList, 'val', null))
                    },
                    max: {
                        type: getXmlAttibute(barCfvo[1]?.attributeList, 'type', 'max'),
                        value: str2num(getXmlAttibute(barCfvo[1]?.attributeList, 'val', null))
                    },
                    isGradient,
                    positiveColor,
                    nativeColor,
                }
                // Univer 在数值等于最小/最大边界时会不绘制数据条，这里把自动 min/max 统一转成
                // 数值类型并略微外扩，保证 Excel 中最短的那根数据条在 Univer 中也能显示出来。
                adjustDataBarAutoToNum(rule.config, this.ranges, readXml, sheetFile);

                break;
            case 'iconSet':
                rule.type = CFRuleType.iconSet;
                if (!isExtLst) {
                    const iconSet = readXml.getElementsByTagName("iconSet", value, false);
                    const iconCfvo = readXml.getElementsByTagName("iconSet/cfvo", value, false);

                    rule.isShowValue = iconSet[0]?.attributeList?.showValue === '0' ? false : true;
                    const iconType = iconSet[0]?.attributeList?.iconSet || DEFAULT_ICON_SET;
                    rule.config = iconCfvo.map((d, index) => {
                        return {
                            operator: rule.operator || CFNumberOperator.greaterThanOrEqual,
                            value: {
                                type: d.attributeList.type,
                                value: str2num(d.attributeList.val)
                            },
                            iconType,
                            // Excel 的 iconSet/cfvo 从低阈值到高阈值排列，Univer 的图标集配置从高到低排列，
                            // 因此这里映射 iconId 并 reverse，保证导入后在 Univer 中显示正确的图标。
                            iconId: (str2num(iconType.charAt(0)) as number) - index - 1
                        }
                    }).reverse();
                } else {
                    const iconSet = readXml.getElementsByTagName("x14:iconSet", value, false)[0];
                    const iconCfvo = readXml.getElementsByTagName("x14:iconSet/x14:cfvo", value, false);
                    const cfIcon = readXml.getElementsByTagName("x14:iconSet/x14:cfIcon", value, false);
                    const isCustom = getXmlAttibute(iconSet?.attributeList, 'custom', '0') === '1'

                    rule.isShowValue = iconSet?.attributeList?.showValue === '0' ? false : true;
                    rule.config = iconCfvo.map((d, index) => {
                        const value = readXml.getElementsByTagName("xm:f", d.value, false)[0];

                        let iconInfo = cfIcon[index]?.attributeList
                        const iconType = (isCustom ? iconInfo?.iconSet : iconSet?.attributeList.iconSet) || DEFAULT_ICON_SET;
                        // 普通图标集（非 custom）没有 x14:cfIcon，无法直接取得 iconId。
                        // Excel 的 iconSet/cfvo 从低阈值到高阈值排列，默认图标 id 依次是 N-1、N-2、...、0，
                        // 这里先按同样规则映射，后面统一 reverse 成 Univer 的高阈值到低阈值顺序。
                        const iconCount = str2num(iconType.charAt(0)) as number;
                        const rawIconId = isCustom ? str2num(iconInfo?.iconId) : index;
                        const iconId = typeof rawIconId !== 'number'
                            ? iconCount - index - 1
                            : iconCount - rawIconId - 1;
                        return {
                            operator: CFNumberOperator.greaterThanOrEqual,
                            value: {
                                type: d.attributeList.type,
                                value: str2num(value?.value)
                            },
                            iconType,
                            iconId,
                            // sort: str2num(iconInfo?.iconId)
                        }
                    }).reverse();
                }

                break;
        }
        const dxfId = getXmlAttibute(attributeList, 'dxfId', null);
        if (dxfId) {
            let dxfs = styles["dxfs"] as Element[];
            const dxf = dxfId !== null ? dxfs[Number(dxfId)] : undefined;
            // let numFmt,font,fill,border;
            if (dxf) {
                const font = dxf.getInnerElements('font')?.[0];
                const numFmt = dxf.getInnerElements('numFmt')?.[0];
                const fill = dxf.getInnerElements('fill')?.[0];
                const border = dxf.getInnerElements('border')?.[0];

                let numfmts = styles["numfmts"] as IattributeList;

                let style: any = {}
                if (border) {
                    const { borderCellValue } = handleBorder(border, styles);
                    style.border = borderCellValue
                }
                if (fill) {
                    const bg = getBackgroundByFill(fill, styles);
                    style.bg = bg;
                }

                if (numFmt) {
                    let numf = numfmts[parseInt(numFmt?.attributeList?.numFmtId || '0')];
                    const cellFormat: CellFormat = {
                        fa: (numf ? escapeCharacter(numf) : numFmt.attributeList.formatCode) || '',
                        t: '',
                    };
                    style.ct = cellFormat;
                }

                if (font) {
                    const { cellValue } = getFontStyle(font, style)
                    style = { ...style, ...cellValue };
                }

                rule.style = style;
            }
        }
        this.rule = rule;
    }

}