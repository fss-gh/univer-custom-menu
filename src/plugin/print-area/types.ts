import type { IRange } from '@univerjs/core';

/** 打印区域插件在 Univer 快照 resources 中使用的资源名 */
export const PRINT_AREA_PLUGIN_NAME = 'SHEET_PRINT_AREA_PLUGIN';

/** 插件配置 key（通过 configService 读取） */
export const PRINT_AREA_PLUGIN_CONFIG = 'PRINT_AREA_PLUGIN_CONFIG';

/** 设置打印区域命令 */
export const SET_PRINT_AREA_COMMAND_ID = 'custom-menu.operation.set-print-area';

/** 清除打印区域命令 */
export const CLEAR_PRINT_AREA_COMMAND_ID = 'custom-menu.operation.clear-print-area';

/** 单个工作表保存的打印区域数据 */
export interface ISheetPrintAreaData {
    /** 支持 Excel 的多个打印区域，例如 A1:B2 和 C3:D4 */
    ranges: IRange[];
}

/** 整个工作簿保存的打印区域数据：subUnitId -> 打印区域 */
export type IPrintAreaResourceData = Record<string, ISheetPrintAreaData>;

/** 打印区域插件配置 */
export interface IPrintAreaPluginConfig {
    /** 设置按钮 tooltip，默认使用语言包 customMenu.printAreaSet */
    tooltip?: string;
    /** 清除按钮 tooltip，默认使用语言包 customMenu.printAreaClear */
    clearTooltip?: string;
    /** 边界颜色，默认蓝色 */
    stroke?: string;
    /** 边界线宽，默认 2 */
    strokeWidth?: number;
    /** 是否显示边界，默认 true */
    showBorder?: boolean;
}
