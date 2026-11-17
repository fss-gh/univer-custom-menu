import type { IMenuButtonItem } from '@univerjs/ui';
import { getMenuHiddenObservable, MenuItemType, RibbonStartGroup } from '@univerjs/ui';
import { ButtonIcon } from '../../components/button-icon/ButtonIcon';
import { BaselineOffset, CommandType, DOCS_NORMAL_EDITOR_UNIT_ID_KEY, DocumentDataModel, EDITOR_ACTIVATED, FOCUSING_SHEET, IAccessor, ICommand, ICommandService, IContextService, IUniverInstanceService, UniverInstanceType, VerticalAlign } from '@univerjs/core';
import { ICustomMenuPulginParams } from '../..';
import { SuperBtnIcon } from '../../components/button-icon/SuperBtnIcon';
import { SetInlineFormatCommand, SetInlineFormatSuperscriptCommand } from '@univerjs/docs-ui';
import { deriveStateFromActiveSheet$, getCurrentRangeDisable$ } from '@univerjs/sheets-ui';
import { RangeProtectionPermissionEditPoint, SetRangeValuesMutation, SetSelectionsOperation, SetWorksheetActiveOperation, SheetsSelectionsService, WorkbookEditablePermission, WorksheetEditPermission, WorksheetSetCellStylePermission } from '@univerjs/sheets';
import { Observable } from 'rxjs';
import { DocSelectionManagerService, SetTextSelectionsOperation } from '@univerjs/docs';

const OperationId = 'custom-menu.operation.single-button';
const SingleButtonOperation: (config?: ICustomMenuPulginParams) => ICommand = (
    config
) => ({
    id: OperationId,
    type: CommandType.COMMAND,
    handler: async (accessor) => {
        const commandService = accessor.get(ICommandService);
        const contextService = accessor.get(IContextService);
        const isCellEditorFocus = contextService.getContextValue(EDITOR_ACTIVATED);

        if (isCellEditorFocus) {
            return commandService.executeCommand(SetInlineFormatSuperscriptCommand.id);
        }

        return false;
    },
});
export function CustomMenuItemSingleButtonFactory(accessor: IAccessor): IMenuButtonItem<string> {
    const commandService = accessor.get(ICommandService);
    const univerInstanceService = accessor.get(IUniverInstanceService);
    const contextService = accessor.get(IContextService);
    return {
        // Bind the command id, clicking the button will trigger this command
        id: OperationId,
        // The type of the menu item, in this case, it is a button
        type: MenuItemType.BUTTON,
        // The icon of the button, which needs to be registered in ComponentManager
        icon: 'SuperBtnIcon',
        // The tooltip of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
        tooltip: 'customMenu.singleButton',
        // The title of the button. Prioritize matching internationalization. If no match is found, the original string will be displayed
        title: 'customMenu.button',
        disabled$: disableMenuWhenSelectRange(accessor),
        activated$: deriveStateFromActiveSheet$(univerInstanceService, false, () => new Observable<boolean>((subscriber) => {
            const disposable = commandService.onCommandExecuted((c) => {
                const id = c.id;
                if (
                    (id === SetTextSelectionsOperation.id || id === SetInlineFormatCommand.id) &&
                    contextService.getContextValue(EDITOR_ACTIVATED) &&
                    contextService.getContextValue(FOCUSING_SHEET)
                ) {
                    const textRun = getFontStyleAtCursor(accessor);
                    if (textRun == null) {
                        return;
                    }

                    const va = textRun.ts?.va;
                    subscriber.next(va === BaselineOffset.SUPERSCRIPT);
                }
            });
            return disposable.dispose;
        })),
        hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET),
    };
}
function getFontStyleAtCursor(accessor: IAccessor) {
    const univerInstanceService = accessor.get(IUniverInstanceService);
    const textSelectionService = accessor.get(DocSelectionManagerService);

    const editorDataModel = univerInstanceService.getUnit<DocumentDataModel>(DOCS_NORMAL_EDITOR_UNIT_ID_KEY);
    const activeTextRange = textSelectionService.getActiveTextRange();

    if (editorDataModel == null || activeTextRange == null) return null;

    const textRuns = editorDataModel.getBody()?.textRuns;
    if (textRuns == null) return;

    const { startOffset } = activeTextRange;
    const textRun = textRuns.find(({ st, ed }) => startOffset >= st && startOffset <= ed);
    return textRun;
}
export function disableMenuWhenSelectRange(accessor: IAccessor): Observable<boolean> {
    const docSelectionManagerService = accessor.get(SheetsSelectionsService);

    return new Observable((subscriber) => {
        const subscription = docSelectionManagerService.selectionChanged$.subscribe((selection) => {
            if (selection && selection[0]) {
                const { startRow, endRow, startColumn, endColumn } = selection[0].range
                if (startRow === endRow && startColumn === endColumn) {
                    subscriber.next(false);
                    return;
                } else {
                    subscriber.next(true);
                }
            }
            subscriber.next(true);
        });

        return () => subscription.unsubscribe();
    });
}

const CustomSingleButtonMenu = (config?: ICustomMenuPulginParams) => ({
    id: OperationId,
    operation: SingleButtonOperation(config),
    menuGroup: RibbonStartGroup.FORMAT,
    order: 8,
    menu: CustomMenuItemSingleButtonFactory,
    icon: { name: 'SuperBtnIcon', component: SuperBtnIcon },
});

export default CustomSingleButtonMenu;