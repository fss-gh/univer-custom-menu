import {
    BooleanNumber,
    VerticalAlign,
    Nullable,
    IBorderData,
    IBorderStyleData,
    WrapStrategy,
    TextDecoration,
    ThemeColorType,
    HorizontalAlign,
    IStyleBase,
} from '@univerjs/core';
import { isObject } from '../common/method';
import { IExcelXmlSheetborderInfoCellForImp, IExcelXmlSheetborderInfoCellValueStyle, IExcelXmlSheetCelldata } from './types';

/**
 * 删除对象中含undefined的值
 * @param object
 * @returns
 */
export function removeEmptyAttr(object: any) {
    for (const key in object) {
        if (Object.prototype.hasOwnProperty.call(object, key)) {
            if (object[key] === undefined) {
                delete object[key]; // 删除值为 undefined 的属性
            } else if (isObject(object[key]) && object[key] !== null) {
                removeEmptyAttr(object[key]); // 对子对象递归
            }
        }
    }
    return object;
}
export const handleStyle = (
    row: IExcelXmlSheetCelldata,
    borderInfo?: IExcelXmlSheetborderInfoCellForImp,
    domContent: boolean = false
) => {
    const { v } = row;
    if (typeof v === 'string' || v === null || v === undefined) {
        return undefined;
    }
    // 0 middle, 1 up, 2 down
    const VerticalAlignMap: any = {
        0: VerticalAlign.MIDDLE,
        1: VerticalAlign.TOP,
        2: VerticalAlign.BOTTOM,
    };
    let border: Nullable<IBorderData> = undefined;
    if (borderInfo?.value && !domContent) {
        const handleBorder: (
            con: IExcelXmlSheetborderInfoCellValueStyle
        ) => IBorderStyleData | null = (con: IExcelXmlSheetborderInfoCellValueStyle) => {
            if (!con) return null;
            return {
                s: con.style,
                cl: { rgb: con.color, th: ThemeColorType.DARK1 },
            };
        };
        border = {
            t: handleBorder(borderInfo.value?.t),
            r: handleBorder(borderInfo.value?.r),
            b: handleBorder(borderInfo.value?.b),
            l: handleBorder(borderInfo.value?.l),
            bl_tr: handleBorder(borderInfo.value?.bl_tr),
            tl_br: handleBorder(borderInfo.value?.tl_br),
        };
    }
    const TextWrap: any = {
        0: WrapStrategy.CLIP,
        1: WrapStrategy.OVERFLOW,
        2: WrapStrategy.WRAP,
    };

    let angle: number | undefined;
    let isVertical = BooleanNumber.FALSE;
    if (v.tr === 3) {
        // 竖排文字
        angle = 0;
        isVertical = BooleanNumber.TRUE;
    } else if (v.rt !== undefined) {
        // Excel textRotation：0-90 为逆时针（上倾），91-180 为顺时针（下倾）
        // Univer 角度：负数表示逆时针/上倾，正数表示顺时针/下倾
        angle = v.rt <= 90 ? -v.rt : v.rt - 90;
    }
    const unMap: any = {
        0: TextDecoration.DASH,
        1: TextDecoration.SINGLE,
        2: TextDecoration.DOUBLE,
        3: TextDecoration.SINGLE,
        4: TextDecoration.DOUBLE,
    };
    // 0 center, 1 left, 2 right
    const htMap: any = {
        0: HorizontalAlign.CENTER,
        1: HorizontalAlign.LEFT,
        2: HorizontalAlign.RIGHT,
    };
    return {
        // bbl: , // bottomBorerLine
        bd: border, // border
        bg: v.bg !== undefined ? { rgb: v.bg, th: ThemeColorType.DARK1 } : undefined, // background
        bl: v.bl, // bold 0: false 1: true
        cl: v.fc !== undefined ? { rgb: v.fc, th: ThemeColorType.DARK1 } : undefined, // foreground
        ff: v.ff, // fontFamily
        fs: v.fs, // fontSize
        ht: v.ht !== undefined ? htMap[v.ht] : undefined, // horizontalAlignment
        it: v.it, // italic 0: false 1: true
        n: v.ct?.fa !== undefined ? { pattern: v.ct.fa } : undefined, //Numfmt pattern
        // ol: { s: v.cl === 0 ? BooleanNumber.TRUE : BooleanNumber.FALSE}, // overline
        // pd: , // padding
        st:
            v.cl !== undefined
                ? { s: v.cl === 1 ? BooleanNumber.TRUE : BooleanNumber.FALSE }
                : undefined, // strikethrough
        tb: v.tb !== undefined ? TextWrap[v.tb] : undefined, // wrapStrategy
        // td: , // textDirection
        tr:
            angle !== undefined
                ? {
                      a: angle,
                      ...(isVertical === BooleanNumber.TRUE ? { v: BooleanNumber.TRUE } : {}),
                  }
                : undefined, // textRotation
        ul:
            v.un !== undefined
                ? {
                      s: v.un === undefined ? BooleanNumber.FALSE : BooleanNumber.TRUE,
                      t: v.un ? unMap[v.un] : TextDecoration.DASH,
                  }
                : undefined, // underline
        va: v.va?v.va*1+1:undefined, // (Subscript 2 下标 /Superscript上标 3 Text)
        vt: v.vt !== undefined ? VerticalAlignMap[v.vt] : undefined, // verticalAlignment
    } as IStyleBase;
};