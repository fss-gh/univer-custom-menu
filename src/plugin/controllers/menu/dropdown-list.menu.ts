import type { IMenuButtonItem, IMenuSelectorItem } from '@univerjs/ui';
import { MenuItemType, RibbonStartGroup } from '@univerjs/ui';

import { CommandType, IAccessor, ICommand } from '@univerjs/core';
import { ICustomMenuPulginParams } from '../..';
import { MainButtonIcon } from '../../components/main-button-icon/MainButtonIcon';
import { ItemIcon } from '../../components/item-icon/ItemIcon';

export const CUSTOM_MENU_DROPDOWN_LIST_OPERATION_ID = 'custom-menu.operation.dropdown-list';

const DropdownListFirstItemOperation: ICommand = {
    id: 'custom-menu.operation.dropdown-list-first-item',
    type: CommandType.OPERATION,
    handler: async (accessor: IAccessor) => {
        alert('Dropdown list first item operation');
        return true;
    },
};

const DropdownListSecondItemOperation: ICommand = {
    id: 'custom-menu.operation.dropdown-list-second-item',
    type: CommandType.OPERATION,
    handler: async (accessor: IAccessor) => {
        alert('Dropdown list second item operation');
        return true;
    },
};

export function CustomMenuItemDropdownListMainButtonFactory(): IMenuSelectorItem<string> {
    return {
        // When type is MenuItemType.SUBITEMS, this factory serves as a container for the drop-down list, and you can set any unique id
        id: CUSTOM_MENU_DROPDOWN_LIST_OPERATION_ID,
        // The type of the menu item, in this case, it is a subitems
        type: MenuItemType.SUBITEMS,
        icon: 'MainButtonIcon',
        tooltip: 'customMenu.dropdownList',
        title: 'customMenu.dropdown',
    };
}

export function CustomMenuItemDropdownListFirstItemFactory(): IMenuButtonItem<string> {
    return {
        id: DropdownListFirstItemOperation.id,
        type: MenuItemType.BUTTON,
        title: 'customMenu.itemOne',
        icon: 'ItemIcon'
    };
}

export function CustomMenuItemDropdownListSecondItemFactory(): IMenuButtonItem<string> {
    return {
        id: DropdownListSecondItemOperation.id,
        type: MenuItemType.BUTTON,
        title: 'customMenu.itemTwo',
        icon: 'ItemIcon'
    };
}

export const CustomDropdownListMenu = (config?: ICustomMenuPulginParams) => ({
    id: CUSTOM_MENU_DROPDOWN_LIST_OPERATION_ID,
    menu: CustomMenuItemDropdownListMainButtonFactory,
    menuGroup: RibbonStartGroup.OTHERS,
    subMenus: [
        {
            id: DropdownListFirstItemOperation.id,
            operation: DropdownListFirstItemOperation,
            menu: CustomMenuItemDropdownListFirstItemFactory,
            icon: { name: 'ItemIcon', component: ItemIcon },
        },
        {
            id: DropdownListSecondItemOperation.id,
            operation: DropdownListSecondItemOperation,
            menu: CustomMenuItemDropdownListSecondItemFactory,
            icon: { name: 'ItemIcon', component: ItemIcon },
        }
    ],
    icon: { name: 'MainButtonIcon', component: MainButtonIcon },
});

