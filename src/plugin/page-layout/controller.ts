import { Disposable, ICommandService, IConfigService, Inject, IResourceManagerService, UniverInstanceType } from '@univerjs/core';
import { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import { registerMenuConfigs } from '../controllers/menu/register-menu';
import {
    AddPageBreakCommand,
    ClearPageBreaksCommand,
    ClearPrintAreaCommand,
    RemovePageBreakCommand,
    SetPrintAreaCommand,
} from './commands';
import { CustomPageLayoutMenu } from './menu';
import { PageLayoutModel } from './model';
import { PageLayoutPanel } from './panel';
import type { IPageLayoutPluginConfig, IPageLayoutResourceData } from './types';
import {
    PAGE_LAYOUT_PANEL_COMPONENT,
    PAGE_LAYOUT_PLUGIN_CONFIG,
    PAGE_LAYOUT_PLUGIN_NAME,
} from './types';

/**
 * 页面布局插件控制器：
 * - 注册 resources，让页面布局数据随 Univer 快照保存/加载
 * - 注册页面布局面板和工具栏入口
 * - 注册打印区域、分页符相关命令
 */
export class PageLayoutMenuController extends Disposable {
    constructor(
        @ICommandService private readonly _commandService: ICommandService,
        @IMenuManagerService private readonly _menuManagerService: IMenuManagerService,
        @Inject(ComponentManager) private readonly _componentManager: ComponentManager,
        @Inject(IconManager) private readonly _iconManager: IconManager,
        @IShortcutService private readonly _shortcutService: IShortcutService,
        @IConfigService private readonly _configService: IConfigService,
        @IResourceManagerService private readonly _resourceManagerService: IResourceManagerService,
        @Inject(PageLayoutModel) private readonly _pageLayoutModel: PageLayoutModel
    ) {
        super();

        this._init();
    }

    private _init(): void {
        this._registerResource();
        this._registerCommands();
        this._registerComponent();
        this._registerMenus();
    }

    private _registerResource(): void {
        this.disposeWithMe(
            this._resourceManagerService.registerPluginResource<IPageLayoutResourceData>({
                pluginName: PAGE_LAYOUT_PLUGIN_NAME,
                businesses: [UniverInstanceType.UNIVER_SHEET],
                onLoad: (unitId, resource) => {
                    this._pageLayoutModel.load(unitId, resource || {});
                },
                onUnLoad: (unitId) => {
                    this._pageLayoutModel.unload(unitId);
                },
                toJson: (unitId) => JSON.stringify(this._pageLayoutModel.getUnitData(unitId)),
                parseJson: (json) => JSON.parse(json),
            })
        );
    }

    private _registerCommands(): void {
        [
            SetPrintAreaCommand,
            ClearPrintAreaCommand,
            AddPageBreakCommand,
            RemovePageBreakCommand,
            ClearPageBreaksCommand,
        ].forEach((command) => {
            this.disposeWithMe(this._commandService.registerCommand(command));
        });
    }

    private _registerComponent(): void {
        this.disposeWithMe(
            this._componentManager.register(PAGE_LAYOUT_PANEL_COMPONENT, PageLayoutPanel)
        );
    }

    private _registerMenus(): void {
        const config: IPageLayoutPluginConfig =
            this._configService.getConfig(PAGE_LAYOUT_PLUGIN_CONFIG) ?? {};

        registerMenuConfigs(
            {
                disposable: this,
                commandService: this._commandService,
                shortcutService: this._shortcutService,
                menuManagerService: this._menuManagerService,
                componentManager: this._componentManager,
                iconManager: this._iconManager,
            },
            [CustomPageLayoutMenu(config)]
        );
    }
}
