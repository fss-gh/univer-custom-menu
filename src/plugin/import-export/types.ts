import type { IWorkbookData, Workbook } from '@univerjs/core';

export const IMPORT_EXPORT_PLUGIN_CONFIG = 'SHEET_IMPORT_EXPORT_MENU_PLUGIN_CONFIG';
export const IMPORT_EXPORT_PLUGIN_NAME = 'SHEET_IMPORT_EXPORT_MENU_PLUGIN';

/** 自定义导入菜单命令 ID */
export const CUSTOM_IMPORT_OPERATION_ID = 'custom-menu.operation.import';
/** 自定义导出菜单命令 ID */
export const CUSTOM_EXPORT_OPERATION_ID = 'custom-menu.operation.export';

/** 导入菜单配置 */
export interface IImportMenuConfig {
    /** 点击导入前调用；返回 false 可中止操作 */
    before?: () => void | Promise<any>;
    /** 导入完成后调用：params.workbook 为新创建的工作簿，params.error 为错误信息 */
    after?: (params?: { workbook?: Workbook; error?: any; message?: string }) => void;
}

/** 导出菜单配置 */
export interface IExportMenuConfig {
    /** 点击导出前调用；返回值会合并到导出参数中，返回 false 可中止操作 */
    before?: () => void | Promise<any>;
    /** 导出完成后调用 */
    after?: (params?: { error?: any; message?: string }) => void;
    /** 自定义工作簿快照；不传时插件自动取当前 Sheet 的 snapshot */
    snapshot?: () => IWorkbookData | Promise<IWorkbookData>;
    /** 导出的文件名 */
    fileName?: string;
}

/** UniverSheetsImportExportPlugin 配置 */
export interface IImportExportMenuPluginConfig {
    import?: IImportMenuConfig;
    export?: IExportMenuConfig;
}
