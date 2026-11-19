import type { IAccessor } from '@univerjs/core';
import { UniverInstanceType } from '@univerjs/core';
import type { IMenuButtonItem } from '@univerjs/ui';
import { getMenuHiddenObservable, MenuItemType, RibbonStartGroup } from '@univerjs/ui';
import type { UniverMenuConfig } from '..';
import { PageLayoutBtnIcon } from '../components/button-icon/PageLayoutBtnIcon';
import { OpenPageLayoutPanelCommand } from './commands';
import type { IPageLayoutPluginConfig } from './types';
import { OPEN_PAGE_LAYOUT_PANEL_COMMAND_ID } from './types';

const createPageLayoutMenuFactory = (
    tooltip: string
): ((accessor: IAccessor) => IMenuButtonItem<string>) => {
    return (accessor) => {
        return {
            id: OPEN_PAGE_LAYOUT_PANEL_COMMAND_ID,
            type: MenuItemType.BUTTON,
            icon: 'PageLayoutBtnIcon',
            tooltip,
            title: tooltip,
            hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET),
        };
    };
};

/** 页面布局菜单：打开打印区域/页边距/分页符/纸张/打印缩放设置面板 */
export function CustomPageLayoutMenu(config?: IPageLayoutPluginConfig): UniverMenuConfig {
    const tooltip = config?.tooltip || 'customMenu.pageLayout';
    return {
        id: OPEN_PAGE_LAYOUT_PANEL_COMMAND_ID,
        operation: OpenPageLayoutPanelCommand,
        menuGroup: RibbonStartGroup.LAYOUT,
        order: 9,
        menu: createPageLayoutMenuFactory(tooltip),
        icon: { name: 'PageLayoutBtnIcon', component: PageLayoutBtnIcon },
    };
}
