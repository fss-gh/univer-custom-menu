import { Disposable, ICommandService, IConfigService, Inject } from '@univerjs/core';
import { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import { CUSTOM_PLUGIN_CONFIG, ICustomMenuPluginConfig, UniverMenuConfig } from '..';

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
        const { menu } = config || {};
        // 递归处理菜单项（含子菜单）
        const buildMenuItem = (item: Omit<UniverMenuConfig,'menuGroup'|'menuPosition'>, menu: Omit<UniverMenuConfig,'menuGroup'|'menuPosition'>[], index: number): Record<string, any> => {

            if (item.operation)
                this.disposeWithMe(
                    this._commandService.registerCommand(item.operation)
                ); // register command
            if (item.shortcut)
                this.disposeWithMe(
                    this._shortcutService.registerShortcut(item.shortcut)
                ); // register shortcut
            if (item.icon) {
                // Univer 1.5.x 的工具栏/菜单图标通过 IconManager 渲染，
                // ComponentManager 用于 label 等自定义组件，二者需要分别注册。
                this.disposeWithMe(
                    this._iconManager.register(item.icon.name, item.icon.component)
                );
                this.disposeWithMe(
                    this._componentManager.register(item.icon.name, item.icon.component)
                );
            }

            // 基础配置：当前菜单项的order和factory
            const baseConfig = {
                order: item.order ?? -((menu.length || 0) - index) - 2,
                menuItemFactory: item.menu,
            };

            // 若存在子菜单，递归处理并合并到当前配置中
            if (item.subMenus && item.subMenus.length > 0) {
                // 递归处理每个子菜单，key为子菜单的id，值为子菜单配置
                const subMenuConfigs = item.subMenus.reduce((subObj, subItem, subIndex) => {
                    const subKey = subItem.id;
                    // 递归调用，传入子菜单列表和当前子项索引
                    subObj[subKey] = buildMenuItem(subItem, item.subMenus!, subIndex);
                    return subObj;
                }, {} as Record<string, any>);

                // 合并基础配置和子菜单配置
                return { ...baseConfig, ...subMenuConfigs };
            }

            // 无subMenus时直接返回基础配置
            return baseConfig;
        };
        menu.forEach((item, index) => {
            if (item.menuPosition) {
                this._menuMangerService.mergeMenu({
                    [item.menuPosition]: {
                        [item.menuGroup]: {
                            [item.id]: buildMenuItem(item, menu, index), // 调用递归函数
                        },
                    }
                });
            } else {
                this._menuMangerService.mergeMenu({
                    [item.menuGroup]: {
                        [item.id]: buildMenuItem(item, menu, index), // 调用递归函数
                    },
                });
            }
        });
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
