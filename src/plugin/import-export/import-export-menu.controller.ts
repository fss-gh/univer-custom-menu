import { Disposable, ICommandService, IConfigService, Inject } from '@univerjs/core';
import { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import type { UniverMenuConfig } from '..';
import CustomImportMenu from '../controllers/menu/import.menu';
import CustomExportMenu from '../controllers/menu/export.menu';
import { registerMenuConfigs } from '../controllers/menu/register-menu';
import { IMPORT_EXPORT_PLUGIN_CONFIG, IImportExportMenuPluginConfig } from './types';

/**
 * 内置「导入 / 导出」菜单的控制器：
 * - 把命令、图标、菜单项注册到 Univer
 * - 配置通过 UniverSheetsImportExportPlugin 传入
 */
export class ImportExportMenuController extends Disposable {
    constructor(
        @ICommandService private readonly _commandService: ICommandService,
        @IMenuManagerService private readonly _menuMangerService: IMenuManagerService,
        @Inject(ComponentManager) private readonly _componentManager: ComponentManager,
        @Inject(IconManager) private readonly _iconManager: IconManager,
        @IShortcutService private readonly _shortcutService: IShortcutService,
        @IConfigService private readonly _configService: IConfigService
    ) {
        super();

        this._init();
    }

    private _init(): void {
        const config: IImportExportMenuPluginConfig =
            this._configService.getConfig(IMPORT_EXPORT_PLUGIN_CONFIG) ?? {};

        const menus: UniverMenuConfig[] = [];
        // 只注册配置项里启用的菜单
        if (config.import) {
            menus.push(CustomImportMenu(config.import) as UniverMenuConfig);
        }
        if (config.export) {
            menus.push(CustomExportMenu(config.export) as UniverMenuConfig);
        }

        registerMenuConfigs(
            {
                disposable: this,
                commandService: this._commandService,
                shortcutService: this._shortcutService,
                menuManagerService: this._menuMangerService,
                componentManager: this._componentManager,
                iconManager: this._iconManager,
            },
            menus
        );
    }
}
