import type { IMenuButtonItem } from '@univerjs/ui';
import { getMenuHiddenObservable, MenuItemType, RibbonStartGroup } from '@univerjs/ui';
import type { IAccessor } from '@univerjs/core';
import { UniverInstanceType } from '@univerjs/core';
import { SheetsSelectionsService } from '@univerjs/sheets';
import { Observable } from 'rxjs';
import type { UniverMenuConfig } from '..';
import { ClearPrintAreaBtnIcon } from '../components/button-icon/ClearPrintAreaBtnIcon';
import { PrintAreaBtnIcon } from '../components/button-icon/PrintAreaBtnIcon';
import { ClearPrintAreaCommand, SetPrintAreaCommand } from './commands';
import type { IPrintAreaPluginConfig } from './types';
import { CLEAR_PRINT_AREA_COMMAND_ID, SET_PRINT_AREA_COMMAND_ID } from './types';

/**
 * 设置打印区域按钮：有选区时可用（包含单个单元格）。
 * 点击后把当前选区设为/替换为该工作表的打印区域。
 */
export function createSetPrintAreaDisabled$(accessor: IAccessor): Observable<boolean> {
    const selectionsService = accessor.get(SheetsSelectionsService);

    return new Observable<boolean>((subscriber) => {
        const emit = () => {
            const selections = selectionsService.getCurrentSelections();
            subscriber.next(!selections?.[0]?.range);
        };

        emit();
        const subscription = selectionsService.selectionChanged$.subscribe(() => emit());
        return () => subscription.unsubscribe();
    });
}

const createPrintAreaMenuFactory = (
    operationId: string,
    iconName: string,
    tooltip: string,
    disabledFactory?: (accessor: IAccessor) => Observable<boolean>
): ((accessor: IAccessor) => IMenuButtonItem<string>) => {
    return (accessor) => {
        return {
            id: operationId,
            type: MenuItemType.BUTTON,
            icon: iconName,
            tooltip,
            title: tooltip,
            disabled$: disabledFactory?.(accessor),
            hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET),
        };
    };
};

/** 设置打印区域菜单 */
export function CustomPrintAreaMenu(config?: IPrintAreaPluginConfig): UniverMenuConfig {
    return {
        id: SET_PRINT_AREA_COMMAND_ID,
        operation: SetPrintAreaCommand,
        menuGroup: RibbonStartGroup.LAYOUT,
        order: 9,
        menu: createPrintAreaMenuFactory(
            SET_PRINT_AREA_COMMAND_ID,
            'PrintAreaBtnIcon',
            config?.tooltip || 'customMenu.printAreaSet',
            createSetPrintAreaDisabled$
        ),
        icon: { name: 'PrintAreaBtnIcon', component: PrintAreaBtnIcon },
    };
}

/** 清除打印区域菜单 */
export function CustomClearPrintAreaMenu(config?: IPrintAreaPluginConfig): UniverMenuConfig {
    return {
        id: CLEAR_PRINT_AREA_COMMAND_ID,
        operation: ClearPrintAreaCommand,
        menuGroup: RibbonStartGroup.LAYOUT,
        order: 10,
        menu: createPrintAreaMenuFactory(
            CLEAR_PRINT_AREA_COMMAND_ID,
            'ClearPrintAreaBtnIcon',
            config?.clearTooltip || 'customMenu.printAreaClear'
        ),
        icon: { name: 'ClearPrintAreaBtnIcon', component: ClearPrintAreaBtnIcon },
    };
}
