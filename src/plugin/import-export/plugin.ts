import { Dependency, IConfigService, Inject, Injector, LocaleService, Plugin, UniverInstanceType } from '@univerjs/core';
import zhCN from '../locale/zh-CN';
import enUS from '../locale/en-US';
import { ImportExportMenuController } from './import-export-menu.controller';
import { IMPORT_EXPORT_PLUGIN_CONFIG, IMPORT_EXPORT_PLUGIN_NAME, IImportExportMenuPluginConfig } from './types';

const AllImportExportController = [ImportExportMenuController];

/**
 * 专用「导入 / 导出」菜单插件。
 *
 * 在 Vue/React 项目中按以下方式注册到 Univer：
 * ```ts
 * univer.registerPlugin(UniverSheetsImportExportPlugin, {
 *   import: { before, after },
 *   export: { snapshot, fileName, before, after },
 * });
 * ```
 */
export class UniverSheetsImportExportPlugin extends Plugin {
    static override type = UniverInstanceType.UNIVER_SHEET;
    static override pluginName = IMPORT_EXPORT_PLUGIN_NAME;

    constructor(
        readonly config: IImportExportMenuPluginConfig,
        @Inject(Injector) protected readonly _injector: Injector,
        @Inject(LocaleService) private readonly _localeService: LocaleService,
        @IConfigService private readonly _configService: IConfigService
    ) {
        super();

        this._localeService.load({ zhCN, enUS });
        this._configService.setConfig(IMPORT_EXPORT_PLUGIN_CONFIG, config || {});
    }

    override onStarting(): void {
        ([AllImportExportController] as Dependency[]).forEach((d) =>
            this._injector.add(d)
        );
    }

    override onReady(): void {
        this._injector.get(ImportExportMenuController);
    }
}
