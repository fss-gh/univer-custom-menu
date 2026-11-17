import { Disposable } from '@univerjs/core';
import type { ICommandService } from '@univerjs/core';
import type { ComponentManager, IconManager, IMenuManagerService, IShortcutService } from '@univerjs/ui';
import type { UniverMenuConfig } from '../..';

export interface IMenuRegisterServices {
    /** 用于把注册的资源绑定到插件生命周期 */
    disposable: Disposable;
    commandService: ICommandService;
    shortcutService: IShortcutService;
    menuManagerService: IMenuManagerService;
    componentManager: ComponentManager;
    iconManager: IconManager;
}

/**
 * 将菜单配置递归注册到 Univer：
 * - 注册 command / shortcut
 * - 把图标同时注册到 IconManager 和 ComponentManager
 * - 通过 MenuManagerService.mergeMenu 合并到工具栏/右键菜单
 */
export function registerMenuConfigs(
    services: IMenuRegisterServices,
    menu: UniverMenuConfig[]
): void {
    const { disposable, commandService, shortcutService, menuManagerService, componentManager, iconManager } = services;

    // 递归处理菜单项（含子菜单）
    const buildMenuItem = (item: Omit<UniverMenuConfig, 'menuGroup' | 'menuPosition'>, items: Omit<UniverMenuConfig, 'menuGroup' | 'menuPosition'>[], index: number): Record<string, any> => {
        if (item.operation) {
            disposable.disposeWithMe(commandService.registerCommand(item.operation));
        }

        if (item.shortcut) {
            disposable.disposeWithMe(shortcutService.registerShortcut(item.shortcut));
        }

        if (item.icon) {
            // Univer 1.5.x 的工具栏/菜单图标通过 IconManager 渲染，
            // ComponentManager 用于 label 等自定义组件，二者需要分别注册。
            disposable.disposeWithMe(iconManager.register(item.icon.name, item.icon.component));
            disposable.disposeWithMe(componentManager.register(item.icon.name, item.icon.component));
        }

        // 基础配置：当前菜单项的 order 和 factory
        const baseConfig = {
            order: item.order ?? -((items.length || 0) - index) - 2,
            menuItemFactory: item.menu,
        };

        // 若存在子菜单，递归处理并合并到当前配置中
        if (item.subMenus && item.subMenus.length > 0) {
            const subMenuConfigs = item.subMenus.reduce((subObj, subItem, subIndex) => {
                const subKey = subItem.id;
                subObj[subKey] = buildMenuItem(subItem, item.subMenus!, subIndex);
                return subObj;
            }, {} as Record<string, any>);

            return { ...baseConfig, ...subMenuConfigs };
        }

        return baseConfig;
    };

    menu.forEach((item, index) => {
        if (item.menuPosition) {
            menuManagerService.mergeMenu({
                [item.menuPosition]: {
                    [item.menuGroup]: {
                        [item.id]: buildMenuItem(item, menu, index),
                    },
                },
            });
        } else {
            menuManagerService.mergeMenu({
                [item.menuGroup]: {
                    [item.id]: buildMenuItem(item, menu, index),
                },
            });
        }
    });
}
