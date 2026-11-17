import {
  MenuItemType,
  IShortcutItem,
  KeyCode,
  MetaKeys,
  IMenuButtonItem,
  RibbonPosition,
  RibbonStartGroup,
} from '@univerjs/ui';
import {
  CommandType,
  ICommand,
} from '@univerjs/core';
import { ICustomMenuPulginParams, UniverMenuConfig } from '../..';
import { SaveBtnIcon } from '../../components/button-icon/SaveBtnIcon';

const OperationId = 'custom-menu.operation.save';
const SaveButtonOperation: (config?: ICustomMenuPulginParams) => ICommand = (
  config
) => ({
  id: OperationId,
  type: CommandType.OPERATION,
  handler: async () => {
    config?.after?.();
    return true;
  },
});

const SaveShortcutItem: IShortcutItem = {
  id: OperationId,
  description: 'shortcut.save',
  group: '1_common-edit',
  binding: MetaKeys.CTRL_COMMAND | KeyCode.S,
};

function CustomMenuItemSaveButtonFactory(): IMenuButtonItem<string> {
  return {
    // Bind the command id, clicking the button will trigger this command
    id: OperationId,
    // The type of the menu item, in this case, it is a button
    type: MenuItemType.BUTTON,
    // The icon of the button, which needs to be registered in ComponentManager
    icon: 'SaveBtnIcon',
    // The tooltip of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    tooltip: 'customMenu.save',
    // The title of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
    title: 'customMenu.save',
  };
}

const CustomSaveMenu: (config?: ICustomMenuPulginParams) => UniverMenuConfig = (
  config?: ICustomMenuPulginParams
) => ({
  id: OperationId,
  operation: SaveButtonOperation(config),
  // shortcut: SaveShortcutItem,
  menuGroup: RibbonStartGroup.OTHERS,
  menu: CustomMenuItemSaveButtonFactory,
  icon: { name: 'SaveBtnIcon', component: SaveBtnIcon },
});

export default CustomSaveMenu;
