import type { IMenuButtonItem } from '@univerjs/ui';
import { getMenuHiddenObservable, MenuItemType, RibbonStartGroup } from '@univerjs/ui';
import {
    BaselineOffset,
    CommandType,
    EDITOR_ACTIVATED,
    FOCUSING_SHEET,
    IAccessor,
    ICommand,
    ICommandService,
    IContextService,
    IUniverInstanceService,
    UniverInstanceType,
} from '@univerjs/core';
import { SetInlineFormatCommand, SetInlineFormatSubscriptCommand, SetInlineFormatSuperscriptCommand } from '@univerjs/docs-ui';
import { deriveStateFromActiveSheet$ } from '@univerjs/sheets-ui';
import { SetTextSelectionsOperation } from '@univerjs/docs';
import { Observable } from 'rxjs';
import { SuperBtnIcon } from '../components/button-icon/SuperBtnIcon';
import { SubBtnIcon } from '../components/button-icon/SubBtnIcon';
import { disableMenuWhenSelectRange, getFontStyleAtCursor } from '../controllers/menu/utils';
import type { UniverMenuConfig } from '..';
import { ISubScriptMenuConfig, SUBSCRIPT_OPERATION_ID, SUPERSCRIPT_OPERATION_ID } from './types';

const createScriptOperation = (
    inlineCommandId: string,
    operationId: string,
    config?: ISubScriptMenuConfig
): ICommand => ({
    id: operationId,
    type: CommandType.COMMAND,
    handler: async (accessor) => {
        const contextService = accessor.get(IContextService);

        // 上下标只作用于单元格编辑器内的文本，编辑器未激活时不处理
        if (!contextService.getContextValue(EDITOR_ACTIVATED)) {
            return false;
        }

        const beforeResult = await config?.before?.();
        if (beforeResult === false) {
            return false;
        }

        const commandService = accessor.get(ICommandService);
        const result = await commandService.executeCommand(inlineCommandId);
        config?.after?.(result);
        return result;
    },
});

const createScriptMenuFactory = (
    operationId: string,
    iconName: string,
    tooltip: string,
    baselineOffset: BaselineOffset
): ((accessor: IAccessor) => IMenuButtonItem<string>) => {
    return (accessor) => {
        const commandService = accessor.get(ICommandService);
        const contextService = accessor.get(IContextService);
        const univerInstanceService = accessor.get(IUniverInstanceService);

        return {
            id: operationId,
            type: MenuItemType.BUTTON,
            icon: iconName,
            tooltip,
            title: tooltip,
            disabled$: disableMenuWhenSelectRange(accessor),
            activated$: deriveStateFromActiveSheet$(univerInstanceService, false, () => {
                return new Observable<boolean>((subscriber) => {
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
                            subscriber.next(va === baselineOffset);
                        }
                    });
                    return disposable.dispose;
                });
            }),
            hidden$: getMenuHiddenObservable(accessor, UniverInstanceType.UNIVER_SHEET),
        };
    };
};

/** 上标菜单配置 */
export function CustomSuperscriptMenu(config?: ISubScriptMenuConfig): UniverMenuConfig {
    return {
        id: SUPERSCRIPT_OPERATION_ID,
        operation: createScriptOperation(SetInlineFormatSuperscriptCommand.id, SUPERSCRIPT_OPERATION_ID, config),
        menuGroup: RibbonStartGroup.FORMAT,
        order: 8,
        menu: createScriptMenuFactory(
            SUPERSCRIPT_OPERATION_ID,
            'SuperBtnIcon',
            config?.tooltip || 'customMenu.superscript',
            BaselineOffset.SUPERSCRIPT
        ),
        icon: { name: 'SuperBtnIcon', component: SuperBtnIcon },
    };
}

/** 下标菜单配置 */
export function CustomSubscriptMenu(config?: ISubScriptMenuConfig): UniverMenuConfig {
    return {
        id: SUBSCRIPT_OPERATION_ID,
        operation: createScriptOperation(SetInlineFormatSubscriptCommand.id, SUBSCRIPT_OPERATION_ID, config),
        menuGroup: RibbonStartGroup.FORMAT,
        order: 8,
        menu: createScriptMenuFactory(
            SUBSCRIPT_OPERATION_ID,
            'SubBtnIcon',
            config?.tooltip || 'customMenu.subscript',
            BaselineOffset.SUBSCRIPT
        ),
        icon: { name: 'SubBtnIcon', component: SubBtnIcon },
    };
}
