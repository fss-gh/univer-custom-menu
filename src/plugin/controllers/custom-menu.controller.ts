import { Disposable, ICommandService, IConfigService, Inject } from '@univerjs/core';
import { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import { CUSTOM_PLUGIN_CONFIG, ICustomMenuPluginConfig } from '..';
import { registerMenuConfigs } from './menu/register-menu';

export class CustomMenuController extends Disposable {
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
        const config: ICustomMenuPluginConfig =
            this._configService.getConfig(CUSTOM_PLUGIN_CONFIG)!;
        const { menu } = config || { menu: [] };

        registerMenuConfigs(
            {
                disposable: this,
                commandService: this._commandService,
                shortcutService: this._shortcutService,
                menuManagerService: this._menuMangerService,
                componentManager: this._componentManager,
                iconManager: this._iconManager,
            },
            menu
        );
    }

    // /**
    //  * register commands
    // */
    // private _initCommands(): void {
    //     [
    //         SingleButtonOperation,
    //         DropdownListFirstItemOperation,
    //         DropdownListSecondItemOperation,
    //     ].forEach((c) => {
    //         this.disposeWithMe(this._commandService.registerCommand(c));
    //     });
    // }

    // /**
    //  * register icon components
    // */
    // private _registerComponents(): void {
    //     const componentManager = this._componentManager;
    //     this.disposeWithMe(componentManager.register("ExportBtnIcon", ExportBtnIcon));
    //     this.disposeWithMe(componentManager.register("ImportBtnIcon", ImportBtnIcon));
    //     this.disposeWithMe(componentManager.register("SubBtnIcon", SubBtnIcon));
    //     this.disposeWithMe(componentManager.register("SuperBtnIcon", SuperBtnIcon));
    //     this.disposeWithMe(componentManager.register("SaveBtnIcon", SaveBtnIcon));
    //     this.disposeWithMe(componentManager.register("ButtonIcon", ButtonIcon));
    //     this.disposeWithMe(componentManager.register("MainButtonIcon", MainButtonIcon));
    //     this.disposeWithMe(componentManager.register("ItemIcon", ItemIcon));
    // }

    // /**
    //  * register menu items
    // */
    // private _initMenus(): void {
    //     this._menuMangerService.mergeMenu({
    //         [RibbonStartGroup.OTHERS]: {
    //             [SingleButtonOperation.id]: {
    //                 order: 10,
    //                 menuItemFactory: CustomMenuItemSingleButtonFactory
    //             },
    //             [CUSTOM_MENU_DROPDOWN_LIST_OPERATION_ID]: {
    //                 order: 11,
    //                 menuItemFactory: CustomMenuItemDropdownListMainButtonFactory,
    //                 [DropdownListFirstItemOperation.id]: {
    //                     order: 0,
    //                     menuItemFactory: CustomMenuItemDropdownListFirstItemFactory
    //                 },
    //                 [DropdownListSecondItemOperation.id]: {
    //                     order: 1,
    //                     menuItemFactory: CustomMenuItemDropdownListSecondItemFactory
    //                 }
    //             }
    //         },
    //         [ContextMenuPosition.MAIN_AREA]: {
    //             [ContextMenuGroup.OTHERS]: {
    //                 [SingleButtonOperation.id]: {
    //                     order: 12,
    //                     menuItemFactory: CustomMenuItemSingleButtonFactory
    //                 },
    //                 [CUSTOM_MENU_DROPDOWN_LIST_OPERATION_ID]: {
    //                     order: 9,
    //                     menuItemFactory: CustomMenuItemDropdownListMainButtonFactory,
    //                     [DropdownListFirstItemOperation.id]: {
    //                         order: 0,
    //                         menuItemFactory: CustomMenuItemDropdownListFirstItemFactory
    //                     },
    //                     [DropdownListSecondItemOperation.id]: {
    //                         order: 1,
    //                         menuItemFactory: CustomMenuItemDropdownListSecondItemFactory
    //                     }
    //                 }
    //             }
    //         }
    //     })
    // }
}
