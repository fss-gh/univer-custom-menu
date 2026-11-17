import { Disposable, ICommandService, IConfigService, Inject } from '@univerjs/core';
import { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import type { UniverMenuConfig } from '..';
import { registerMenuConfigs } from '../controllers/menu/register-menu';
import { CustomSubscriptMenu, CustomSuperscriptMenu } from './menu';
import { ISuperSubscriptPluginConfig, SUPER_SUBSCRIPT_PLUGIN_CONFIG } from './types';

/**
 * 上下标插件菜单控制器：
 * - 注册上标 / 下标两个命令、图标和工具栏按钮
 * - 配置通过 UniverSheetsSuperSubscriptPlugin 传入
 */
export class SuperSubscriptMenuController extends Disposable {
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
        const config: ISuperSubscriptPluginConfig =
            this._configService.getConfig<ISuperSubscriptPluginConfig>(SUPER_SUBSCRIPT_PLUGIN_CONFIG) ?? {};

        const menus: UniverMenuConfig[] = [];

        if (config.superscript !== false && config.superscript?.show !== false) {
            menus.push(CustomSuperscriptMenu(config.superscript || undefined));
        }
        if (config.subscript !== false && config.subscript?.show !== false) {
            menus.push(CustomSubscriptMenu(config.subscript || undefined));
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
