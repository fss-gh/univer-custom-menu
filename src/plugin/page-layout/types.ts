import type { IRange } from '@univerjs/core';

/** 页面布局插件在 Univer 快照 resources 中使用的资源名 */
export const PAGE_LAYOUT_PLUGIN_NAME = 'SHEET_PAGE_LAYOUT_PLUGIN';

/** 插件配置 key（通过 configService 读取） */
export const PAGE_LAYOUT_PLUGIN_CONFIG = 'PAGE_LAYOUT_PLUGIN_CONFIG';

/** 打开页面布局面板 */
export const OPEN_PAGE_LAYOUT_PANEL_COMMAND_ID = 'custom-menu.operation.open-page-layout-panel';

/** 设置打印区域 */
export const SET_PRINT_AREA_COMMAND_ID = 'custom-menu.operation.set-print-area';

/** 清除打印区域 */
export const CLEAR_PRINT_AREA_COMMAND_ID = 'custom-menu.operation.clear-print-area';

/** 添加分页符，参数 { type: 'row' | 'col', index?: number } */
export const ADD_PAGE_BREAK_COMMAND_ID = 'custom-menu.operation.add-page-break';

/** 删除分页符，参数 { type: 'row' | 'col', index: number } */
export const REMOVE_PAGE_BREAK_COMMAND_ID = 'custom-menu.operation.remove-page-break';

/** 清除当前工作表所有分页符 */
export const CLEAR_PAGE_BREAKS_COMMAND_ID = 'custom-menu.operation.clear-page-breaks';

/** 侧边栏面板 id */
export const PAGE_LAYOUT_PANEL_ID = 'custom-menu.page-layout-panel';

/** 侧边栏面板组件名 */
export const PAGE_LAYOUT_PANEL_COMPONENT = 'custom-menu.page-layout-panel-component';

/** 页面方向 */
export type PageOrientation = 'portrait' | 'landscape';

/** 页边距，单位：英寸（与 Excel pageMargins 一致） */
export interface IPageMargins {
    top: number;
    bottom: number;
    left: number;
    right: number;
    header: number;
    footer: number;
}

/** 纸张设置 */
export interface IPageSetup {
    orientation: PageOrientation;
    /** Excel paperSize，例：9 = A4，8 = A3，1 = Letter */
    paperSize: number;
}

/** 单个工作表保存的页面布局数据 */
export interface ISheetPageLayoutData {
    /** 打印区域，支持 Excel 的多个打印区域 */
    printArea?: IRange[];
    /** 页边距（英寸） */
    margins?: IPageMargins;
    /** 纸张方向/大小 */
    pageSetup?: IPageSetup;
    /** 水平分页符：新页起始行的 0-based 行索引 */
    rowBreaks?: number[];
    /** 垂直分页符：新页起始列的 0-based 列索引 */
    colBreaks?: number[];
}

/** 整个工作簿保存的页面布局数据：subUnitId -> 页面布局 */
export type IPageLayoutResourceData = Record<string, ISheetPageLayoutData>;

/** 页面布局插件配置 */
export interface IPageLayoutPluginConfig {
    /** 页面布局按钮 tooltip，默认使用语言包 customMenu.pageLayout */
    tooltip?: string;
    /** 打印区域边界颜色，默认蓝色 */
    stroke?: string;
    /** 打印区域边界线宽，默认 2 */
    strokeWidth?: number;
    /** 是否显示打印区域边界，默认 true */
    showBorder?: boolean;
    /** 分页符颜色，默认蓝色 */
    breakStroke?: string;
    /** 是否显示分页符，默认 true */
    showPageBreaks?: boolean;
}

/** 默认页边距（英寸），与 Excel 默认值一致 */
export const DEFAULT_PAGE_MARGINS: IPageMargins = {
    top: 0.75,
    bottom: 0.75,
    left: 0.7,
    right: 0.7,
    header: 0.3,
    footer: 0.3,
};

/** 默认纸张设置：A4 纵向 */
export const DEFAULT_PAGE_SETUP: IPageSetup = {
    orientation: 'portrait',
    paperSize: 9,
};

/** 英寸转厘米 */
export const inchToCm = (value: number): number => Math.round(value * 2.54 * 100) / 100;

/** 厘米转英寸 */
export const cmToInch = (value: number): number => Math.round((value / 2.54) * 10000) / 10000;
