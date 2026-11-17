import { CommandType, type IAccessor, type ICommand, IUniverInstanceService, UniverInstanceType, type Workbook } from '@univerjs/core';
import { SheetsSelectionsService } from '@univerjs/sheets';
import { PrintAreaModel } from './model';
import { CLEAR_PRINT_AREA_COMMAND_ID, SET_PRINT_AREA_COMMAND_ID } from './types';

/**
 * 设置打印区域：将当前选区作为该工作表的打印区域。
 * 多次点击会用新选区替换旧打印区域，从而支持“修改”。
 */
export const SetPrintAreaCommand: ICommand = {
    id: SET_PRINT_AREA_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: async (accessor: IAccessor) => {
        const instanceService = accessor.get(IUniverInstanceService);
        const workbook = instanceService.getCurrentUnitOfType(UniverInstanceType.UNIVER_SHEET) as Workbook | null;
        if (!workbook) {
            return false;
        }

        const sheet = workbook.getActiveSheet();
        if (!sheet) {
            return false;
        }

        const selectionsService = accessor.get(SheetsSelectionsService);
        const selection = selectionsService.getCurrentLastSelection();
        if (!selection?.range) {
            return false;
        }

        accessor.get(PrintAreaModel).setPrintArea(
            workbook.getUnitId(),
            sheet.getSheetId(),
            [selection.range]
        );

        return true;
    },
};

/** 清除当前工作表的打印区域 */
export const ClearPrintAreaCommand: ICommand = {
    id: CLEAR_PRINT_AREA_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: async (accessor: IAccessor) => {
        const instanceService = accessor.get(IUniverInstanceService);
        const workbook = instanceService.getCurrentUnitOfType(UniverInstanceType.UNIVER_SHEET) as Workbook | null;
        if (!workbook) {
            return false;
        }

        const sheet = workbook.getActiveSheet();
        if (!sheet) {
            return false;
        }

        accessor.get(PrintAreaModel).clearPrintArea(workbook.getUnitId(), sheet.getSheetId());

        return true;
    },
};
