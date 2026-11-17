import { IAccessor, ICommand, Univer } from '@univerjs/core';
import { ContextMenuGroup, ContextMenuPosition, IMenuButtonItem, IMenuSelectorItem, IShortcutItem, RibbonDataGroup, RibbonFormulasGroup, RibbonInsertGroup, RibbonOthersGroup, RibbonPosition, RibbonStartGroup, RibbonViewGroup } from '@univerjs/ui';

export const CUSTOM_PLUGIN_CONFIG = 'CUSTOM_PLUGIN_CONFIG';
export interface ICustomMenuPluginConfig {
    instance: Univer;
    menu: Array<UniverMenuConfig>;
}
export interface UniverMenuConfig {
    id: string;
    operation?: ICommand<object, boolean>;
    shortcut?: IShortcutItem<object>;
    menu: (accessor:IAccessor) => IMenuButtonItem<string> | IMenuSelectorItem<string>;
    order?: number;
    menuPosition?: ContextMenuPosition | RibbonPosition;
    menuGroup: RibbonStartGroup|RibbonInsertGroup|RibbonFormulasGroup|RibbonDataGroup|RibbonViewGroup|RibbonOthersGroup|ContextMenuGroup;
    subMenus?: Array<Omit<UniverMenuConfig, 'menuGroup'|'menuPosition'>>;
    icon?: { name: string; component: any };
    onlyOperation?: boolean;
}
export interface ICustomMenuPulginParams {
    [key: string]: any;
    before?: () => void | Promise<any>;
    after?: (param?: any) => void;
}
export { UniverSheetsCustomMenuPlugin } from './plugin';