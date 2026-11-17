import { Disposable, ICommandService, IConfigService, Inject, IResourceManagerService, UniverInstanceType } from '@univerjs/core';
import { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import { registerMenuConfigs } from '../controllers/menu/register-menu';
import { CustomClearPrintAreaMenu, CustomPrintAreaMenu } from './menu';
import { PrintAreaModel } from './model';
import type { IPrintAreaPluginConfig, IPrintAreaResourceData } from './types';
import { PRINT_AREA_PLUGIN_CONFIG, PRINT_AREA_PLUGIN_NAME } from './types';

/**
 * 打印区域插件控制器：
 * - 注册 resources，让打印区域数据随 Univer 快照保存/加载
 * - 注册“设置打印区域”和“清除打印区域”菜单
 */
export class PrintAreaMenuController extends Disposable {
    constructor(
        @ICommandService private readonly _commandService: ICommandService,
        @IMenuManagerService private readonly _menuManagerService: IMenuManagerService,
        @Inject(ComponentManager) private readonly _componentManager: ComponentManager,
        @Inject(IconManager) private readonly _iconManager: IconManager,
        @IShortcutService private readonly _shortcutService: IShortcutService,
        @IConfigService private readonly _configService: IConfigService,
        @IResourceManagerService private readonly _resourceManagerService: IResourceManagerService,
        @Inject(PrintAreaModel) private readonly _printAreaModel: PrintAreaModel
    ) {
        super();

        this._init();
    }

    private _init(): void {
        this._registerResource();
        this._registerMenus();
    }

    /** 把打印区域数据注册为 Univer 快照中的 SHEET_PRINT_AREA_PLUGIN 资源 */
    private _registerResource(): void {
        this.disposeWithMe(
            this._resourceManagerService.registerPluginResource<IPrintAreaResourceData>({
                pluginName: PRINT_AREA_PLUGIN_NAME,
                businesses: [UniverInstanceType.UNIVER_SHEET],
                onLoad: (unitId, resource) => {
                    this._printAreaModel.load(unitId, resource || {});
                },
                onUnLoad: (unitId) => {
                    this._printAreaModel.unload(unitId);
                },
                toJson: (unitId) => JSON.stringify(this._printAreaModel.getUnitData(unitId)),
                parseJson: (json) => JSON.parse(json),
            })
        );
    }

    private _registerMenus(): void {
        const config: IPrintAreaPluginConfig =
            this._configService.getConfig(PRINT_AREA_PLUGIN_CONFIG) ?? {};

        registerMenuConfigs(
            {
                disposable: this,
                commandService: this._commandService,
                shortcutService: this._shortcutService,
                menuManagerService: this._menuManagerService,
                componentManager: this._componentManager,
                iconManager: this._iconManager,
            },
            [CustomPrintAreaMenu(config), CustomClearPrintAreaMenu(config)]
        );
    }
}
