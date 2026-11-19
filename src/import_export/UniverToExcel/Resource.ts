import { Anchor, Workbook, Worksheet } from "@zwight/exceljs";
import { escapeCharacter, getEmusByPx, getImageExtension, getRangetxt, isEmpty, jsonParse, removeEmptyAttr, str2num } from "../common/method";
import { CFRuleType, CFSubRuleType, DEFAULT_ICON_SET } from "../ToUniver/ExcelXmlCondition";
import { hex2argb } from "./util";
import { cellStyle } from "./CellStyle";

import '@zwight/exceljs';

declare module '@zwight/exceljs' {
  interface DataValidations {
    add(address: string, validation: DataValidation): void;

    find(address: string): DataValidation | undefined;

    remove(address: string): void;
  }

  interface Worksheet {
    dataValidations: DataValidations;
  }
}

export class Resource {
    sheetId: string;
    workbook: Workbook;
    worksheet: Worksheet;
    resources: any;
    constructor(
        sheetId: string,
        workbook: Workbook,
        worksheet: Worksheet,
        resources: any
    ) {
        this.sheetId = sheetId;
        this.workbook = workbook;
        this.worksheet = worksheet;
        this.resources = resources;

        this.setImages();
        this.setNotes();
        this.setConditional();
        this.setDataValidation();
        this.setFilter();
    }

    private handleRang(range: any) {
        const { startRow, startColumn, endRow, endColumn } = range;
        return getRangetxt({
            row: [startRow, endRow],
            column: [startColumn, endColumn],
        }, '')
    }
    // private setRangeProtection() {
    //     const rangeProtection = this.getSheetResource('SHEET_RANGE_PROTECTION_PLUGIN');
    // }
    private setFilter() {
        const filters = this.getSheetResource('SHEET_FILTER_PLUGIN');
        if (!filters) return;
        this.worksheet.autoFilter = this.handleRang(filters.ref);
    }
    private setConditional() {
        const conditionals = this.getSheetResource('SHEET_CONDITIONAL_FORMATTING_PLUGIN');
        const ruleList: any[] = [];
        if (!conditionals) return;
        conditionals.forEach((conditional: any) => {
            const { ranges, rule, stopIfTrue } = conditional;
            ranges.forEach((range: any) => {
                const ref = this.handleRang(range);
                const index = ruleList.findIndex(d => d.ref === ref);
                const ruleValue = this.handleRule(conditional);
                if (index > -1) {
                    ruleList[index].rules.push(ruleValue)
                    return;
                }
                ruleList.push({
                    ref,
                    rules: [ruleValue]
                })
            });
        });
        // console.log(this.worksheet.name, ruleList)
        ruleList.forEach(d => {
            this.worksheet.addConditionalFormatting(d);
        });
    }

    private setDataValidation() {
        const datavalidations = this.getSheetResource('SHEET_DATA_VALIDATION_PLUGIN');
        const excelToDate = (v: number, date1904 = false) => {
            // eslint-disable-next-line no-mixed-operators
            const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
            return new Date(millisecondSinceEpoch);
        }
        datavalidations?.forEach((validate: any) => {
            const {
                ranges = [],
                type,
                allowBlank,
                operator,
                formula1,
                formula2,
                showErrorMessage,
                showInputMessage,
                prompt,
                promptTitle,
                error,
                errorTitle,
                errorStyle
            } = validate || {};


            const styleMap = ['information', 'stop', 'warning']
            
            let formulae = [formula1];
            if (!isEmpty(formula2)) formulae.push(formula2);

            if (type === 'list') {
                formulae = [`"${formula1}"`];
                if (!isEmpty(formula2)) formulae.push(`"${formula2}"`);
            }
            if (type === 'date') {
                formulae = [excelToDate(formula1)];
                if (!isEmpty(formula2)) formulae.push(excelToDate(formula2));
            }

            const list = ranges.map((range: any) => this.handleRang(range));

            
            const valid = {
                type,
                allowBlank,
                operator,
                formulae,
                showErrorMessage,
                showInputMessage,
                prompt,
                promptTitle,
                error,
                errorTitle,
                errorStyle: styleMap[errorStyle]
            }
            // this.worksheet.dataValidations.add(list.join(' '), removeEmptyAttr(valid))
            // console.log(this.sheetId, this.worksheet.name, list, valid);
            list.forEach((address: string) => {
                this.worksheet.dataValidations.add(address, removeEmptyAttr(valid))
            });
        });
    }

    private setImages() {
        const images = this.getSheetResource('SHEET_DRAWING_PLUGIN');
        const sheetIamges = images?.data;
        if (!sheetIamges) return;
        for (const key in sheetIamges) {
            const element = sheetIamges[key];
            if (!element?.source) {
                continue;
            }

            const images = this.workbook.getImages();
            const extension = getImageExtension(element.source);
            let imageId = images.findIndex(d => d.extension === extension && d.base64 === element.source);
            if (imageId === -1) {
                imageId = this.workbook.addImage({
                    base64: element.source,
                    extension,
                })
            }
            
            const handlePosition = (position: any) => {
                return {
                    nativeCol: position.column,
                    // xdr:colOff / xdr:rowOff 必须是整数 EMU，否则 Excel 会提示修复 drawing1.xml
                    nativeColOff: Math.round(getEmusByPx(position.columnOffset)),
                    nativeRow: position.row,
                    nativeRowOff: Math.round(getEmusByPx(position.rowOffset)),
                }
            }
            // console.log(handlePosition(element.sheetTransform.from), handlePosition(element.sheetTransform.to))

            const imageOptions: any = {
                tl: handlePosition(element.sheetTransform.from) as Anchor,
                br: handlePosition(element.sheetTransform.to) as Anchor,
            };
            if (element.anchorType) {
                // Univer SheetDrawingAnchorType：0=Position(oneCell)、1=Both(twoCell)、2=None(absolute)
                const anchorType = String(element.anchorType);
                imageOptions.editAs = anchorType === '0' ? 'oneCell' : anchorType === '2' ? 'absolute' : 'twoCell';
            }

            this.worksheet.addImage(imageId, imageOptions)
        }
    }

    /**
     * Univer SHEET_NOTE_PLUGIN 资源 → Excel 传统批注（cell.note）
     */
    private setNotes() {
        const notes = this.getSheetResource('SHEET_NOTE_PLUGIN');
        if (!notes) return;
        for (const rowKey in notes) {
            const rowNotes = notes[rowKey];
            if (!rowNotes) continue;
            for (const colKey in rowNotes) {
                const note = rowNotes[colKey];
                if (!note?.note) continue;
                const cell = this.worksheet.getCell(Number(rowKey) + 1, Number(colKey) + 1);
                cell.note = { texts: [{ text: note.note || '' }] };
            }
        }
    }

    private getSheetResource(name: string) {
        const resources = jsonParse(this.resources.find((d: any) => d.name === name)?.data);
        const sheetResources = resources[this.sheetId];
        return sheetResources;
    }

    private handleRule(conditional: any) {
        const { rule, stopIfTrue, order } = conditional;
        const ruleValue: any = {};
        if (stopIfTrue) ruleValue.stopIfTrue = 1;
        ruleValue.priority = order;
        if (rule.style) {
            // const formatCode = rule.style.n?.pattern ? { formatCode: rule.style.n?.pattern } : undefined;
            ruleValue.style = cellStyle(rule.style, rule.style.n?.pattern, true);
        }
        if (rule.operator) ruleValue.operator = rule.operator;
        switch (rule.type) {
            case CFRuleType.colorScale:
                ruleValue.type = CFRuleType.colorScale;
                ruleValue.cfvo = rule.config?.map((d: any) => {
                    return {
                        type: d.value.type,
                        // min/max/percent 等类型没有 val 属性；null 会让 ExcelJS 调用 toString 报错
                        value: d.value.value ?? undefined
                    }
                });
                ruleValue.color = rule.config?.map((d: any) => {
                    return {
                        argb: hex2argb(d.color)
                    }
                });
                break;
            case CFRuleType.dataBar:
                ruleValue.type = CFRuleType.dataBar;
                ruleValue.showValue = rule.isShowValue;
                // ExcelJS 对 gradient=true 的数据条会在 cfRule 的 extLst 中生成空的 x14:id，
                // 这会导致 Excel 报告 /xl/worksheets/sheet1.xml 有 XML 错误。
                // 只有 gradient=false 时才显式写 gradient=0；否则不设置 gradient，
                // 让 ExcelJS 按默认渐变数据条写出完整的 x14:dataBar 扩展。
                if (rule.config.isGradient === false) {
                    ruleValue.gradient = false;
                }
                ruleValue.cfvo = [{
                    type: rule.config.min.type,
                    value: rule.config.min.value ?? undefined
                }, {
                    type: rule.config.max.type,
                    value: rule.config.max.value ?? undefined
                }];
                if (rule.config.nativeColor) {
                    ruleValue.negativeFillColor = { argb: hex2argb(rule.config.nativeColor) };
                    ruleValue.negativeBorderColor = { argb: hex2argb(rule.config.nativeColor) };
                    ruleValue.border = true;
                    ruleValue.borderColor = { argb: hex2argb(rule.config.positiveColor) };
                    ruleValue.axisColor = { argb: "FF000000" };
                    // 与 Excel 原文件保持一致：不显式写 negativeBarColorSameAsPositive，
                    // 由 negativeFillColor 表达负值填充；只声明负值边框不同色。
                    ruleValue.negativeBarBorderColorSameAsPositive = false;
                }
                ruleValue.color = { argb: hex2argb(rule.config.positiveColor) }
                ruleValue.axisPosition = 'auto';
                ruleValue.direction = 'leftToRight';
                ruleValue.minLength = 0;
                ruleValue.maxLength = 100;

                break;
            case CFRuleType.iconSet:
                ruleValue.type = CFRuleType.iconSet;
                ruleValue.reverse = false;
                ruleValue.showValue = rule.isShowValue;
                // x14:iconSet 的 iconSet 属性是必填项，缺失会导致 Excel 报 XML 错误。
                ruleValue.iconSet = rule.config?.[0]?.iconType || DEFAULT_ICON_SET;
                // Univer 的默认图标顺序是高阈值→低阈值 iconId 为 0、1、2...，
                // 与 Excel 默认图标顺序一致时不需要写 custom/cfIcon，保持 Excel 原生写法；
                // 只有自定义过图标顺序才写 custom="1" + x14:cfIcon。
                const isDefaultIconOrder = rule.config?.every((d: any, index: number) => {
                    const iconId = str2num(d?.iconId);
                    return typeof iconId === "number" && iconId === index;
                });
                if (!isDefaultIconOrder) {
                    ruleValue.icons = rule.config?.map((d: any) => {
                        const iconId = (str2num(d.iconType.charAt(0)) as number) - (str2num(d?.iconId) as number) -1
                        return {
                            iconId,
                            iconSet: d.iconType
                        }
                    }).reverse();
                }
                ruleValue.cfvo = rule.config?.map((d: any) => {
                    return {
                        type: d.value.type,
                        value: d.value.value ?? undefined
                    }
                }).reverse();
                break;
            case CFRuleType.highlightCell:
                switch (rule.subType) {
                    case CFSubRuleType.average:
                        ruleValue.type = 'aboveAverage';
                        ruleValue.aboveAverage = false;
                        break;
                    case CFSubRuleType.duplicateValues:
                        ruleValue.type = 'duplicateValues';
                        break;
                    case CFSubRuleType.formula:
                        ruleValue.type = 'expression';
                        ruleValue.formulae = [escapeCharacter(rule.value)];
                        break;
                    case CFSubRuleType.number:
                        ruleValue.type = 'cellIs';
                        // Univer 中字符串条件值可能保留 XML/HTML 实体（如 &quot;F&quot;）。
                        // 直接写入会被 ExcelJS 二次转义成 &amp;quot;，Excel 解析公式失败后会删除条件格式。
                        ruleValue.formulae = [
                            typeof rule.value === 'string' ? escapeCharacter(rule.value) : rule.value
                        ];
                        break;
                    case CFSubRuleType.rank:
                        ruleValue.type = 'top10';
                        ruleValue.rank = rule.value;
                        ruleValue.percent = rule.isPercent;
                        ruleValue.bottom = rule.isBottom;
                        break;
                    case CFSubRuleType.text:
                        ruleValue.type = 'containsText';
                        ruleValue.text = rule.value
                        break;
                    case CFSubRuleType.timePeriod:
                        ruleValue.type = 'timePeriod';
                        ruleValue.timePeriod = rule.operator
                        break;
                    // case CFSubRuleType.uniqueValues:
                    //     ruleValue.type = 'uniqueValues';
                    //     break;
                }
                break;
        }
        return ruleValue;
    }
}