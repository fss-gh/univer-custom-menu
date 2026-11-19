import {
    CommandType,
    type IAccessor,
    type ICommand,
    IUniverInstanceService,
    LocaleService,
    UniverInstanceType,
    type Workbook,
} from '@univerjs/core';
import { SheetsSelectionsService } from '@univerjs/sheets';
import { ISidebarService } from '@univerjs/ui';
import { PageLayoutModel } from './model';
import {
    ADD_PAGE_BREAK_COMMAND_ID,
    CLEAR_PAGE_BREAKS_COMMAND_ID,
    CLEAR_PRINT_AREA_COMMAND_ID,
    OPEN_PAGE_LAYOUT_PANEL_COMMAND_ID,
    PAGE_LAYOUT_PANEL_COMPONENT,
    PAGE_LAYOUT_PANEL_ID,
    REMOVE_PAGE_BREAK_COMMAND_ID,
    SET_PRINT_AREA_COMMAND_ID,
} from './types';

interface IPageLayoutTarget {
    unitId: string;
    subUnitId: string;
}

function getTarget(accessor: IAccessor): IPageLayoutTarget | null {
    const workbook = accessor
        .get(IUniverInstanceService)
        .getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET);
    const sheet = workbook?.getActiveSheet();
    if (!workbook || !sheet) {
        return null;
    }
    return { unitId: workbook.getUnitId(), subUnitId: sheet.getSheetId() };
}

/** 打开页面布局侧边栏 */
export const OpenPageLayoutPanelCommand: ICommand = {
    id: OPEN_PAGE_LAYOUT_PANEL_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor: IAccessor) => {
        if (!getTarget(accessor)) {
            return false;
        }

        accessor.get(ISidebarService).open({
            id: PAGE_LAYOUT_PANEL_ID,
            header: { title: accessor.get(LocaleService).t('customMenu.pageLayout') },
            children: { label: PAGE_LAYOUT_PANEL_COMPONENT },
            width: 320,
        });

        return true;
    },
};

/**
 * 设置打印区域：将当前选区作为该工作表的打印区域。
 * 多次设置会用新选区替换旧打印区域。
 */
export const SetPrintAreaCommand: ICommand = {
    id: SET_PRINT_AREA_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor: IAccessor) => {
        const target = getTarget(accessor);
        if (!target) {
            return false;
        }

        const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection();
        if (!selection?.range) {
            return false;
        }

        accessor.get(PageLayoutModel).setPrintArea(target.unitId, target.subUnitId, [selection.range]);
        return true;
    },
};

/** 清除当前工作表的打印区域 */
export const ClearPrintAreaCommand: ICommand = {
    id: CLEAR_PRINT_AREA_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor: IAccessor) => {
        const target = getTarget(accessor);
        if (!target) {
            return false;
        }

        accessor.get(PageLayoutModel).clearPrintArea(target.unitId, target.subUnitId);
        return true;
    },
};

/**
 * 添加分页符：
 * - params.index 省略时使用当前选区起始行/列；
 * - type=row 时在 index 行上方分页，type=col 时在 index 列左侧分页。
 */
export const AddPageBreakCommand: ICommand = {
    id: ADD_PAGE_BREAK_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor: IAccessor, params?: { type?: 'row' | 'col'; index?: number }) => {
        const target = getTarget(accessor);
        if (!target || !params?.type) {
            return false;
        }

        let index = params.index;
        if (index === undefined) {
            const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection();
            if (!selection?.range) {
                return false;
            }
            index = params.type === 'row' ? selection.range.startRow : selection.range.startColumn;
        }

        accessor.get(PageLayoutModel).addPageBreak(target.unitId, target.subUnitId, params.type, index);
        return true;
    },
};

/** 删除指定行/列分页符 */
export const RemovePageBreakCommand: ICommand = {
    id: REMOVE_PAGE_BREAK_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor: IAccessor, params?: { type?: 'row' | 'col'; index?: number }) => {
        const target = getTarget(accessor);
        if (!target || !params?.type || params.index === undefined) {
            return false;
        }

        accessor
            .get(PageLayoutModel)
            .removePageBreak(target.unitId, target.subUnitId, params.type, params.index);
        return true;
    },
};

/** 清除当前工作表所有分页符 */
export const ClearPageBreaksCommand: ICommand = {
    id: CLEAR_PAGE_BREAKS_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor: IAccessor) => {
        const target = getTarget(accessor);
        if (!target) {
            return false;
        }

        accessor.get(PageLayoutModel).clearPageBreaks(target.unitId, target.subUnitId);
        return true;
    },
};
