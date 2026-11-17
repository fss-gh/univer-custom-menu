/**
 * Univer 导入/导出菜单插件 —— 库入口
 *
 * 供 Vue/React 项目本地打包引用：
 * ```ts
 * import { UniverSheetsImportExportPlugin } from 'univer-custom-menu';
 *
 * univer.registerPlugin(UniverSheetsImportExportPlugin, {
 *   import: { before, after },
 *   export: { snapshot, fileName, before, after },
 * });
 * ```
 */
export { UniverSheetsImportExportPlugin } from './plugin/import-export/plugin';
export {
    CUSTOM_IMPORT_OPERATION_ID,
    CUSTOM_EXPORT_OPERATION_ID,
    IMPORT_EXPORT_PLUGIN_CONFIG,
    IMPORT_EXPORT_PLUGIN_NAME,
} from './plugin/import-export/types';
export type {
    IImportMenuConfig,
    IExportMenuConfig,
    IImportExportMenuPluginConfig,
} from './plugin/import-export/types';

/** 上下标插件（上标 + 下标） */
export { UniverSheetsSuperSubscriptPlugin } from './plugin/super-subscript/plugin';
export {
    SUPER_SUBSCRIPT_PLUGIN_CONFIG,
    SUPER_SUBSCRIPT_PLUGIN_NAME,
    SUPERSCRIPT_OPERATION_ID,
    SUBSCRIPT_OPERATION_ID,
} from './plugin/super-subscript/types';
export type {
    ISubScriptMenuConfig,
    ISuperSubscriptPluginConfig,
} from './plugin/super-subscript/types';

/** 打印区域插件（导入显示蓝色虚线边界 + 设置/清除 + JSON 保存） */
export { UniverSheetsPrintAreaPlugin } from './plugin/print-area/plugin';
export {
    CLEAR_PRINT_AREA_COMMAND_ID,
    PRINT_AREA_PLUGIN_CONFIG,
    PRINT_AREA_PLUGIN_NAME,
    SET_PRINT_AREA_COMMAND_ID,
} from './plugin/print-area/types';
export type {
    IPrintAreaPluginConfig,
    IPrintAreaResourceData,
    ISheetPrintAreaData,
} from './plugin/print-area/types';

/** Excel 转换工具（需要时可直接使用） */
export { LuckyExcel } from './import_export';
/** 文件选择工具 */
export { waitUserSelectExcelFile } from './utils';
