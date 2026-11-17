import { DOCS_NORMAL_EDITOR_UNIT_ID_KEY, DocumentDataModel, IAccessor, IUniverInstanceService } from "@univerjs/core";
import { DocSelectionManagerService } from "@univerjs/docs";
import { SheetsSelectionsService } from "@univerjs/sheets";
import { Observable } from "rxjs";

export function getFontStyleAtCursor(accessor: IAccessor) {
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
           const { startRow,endRow,startColumn,endColumn } = selection![0].range
            if (startRow === endRow && startColumn===endColumn) {
                subscriber.next(false);
                return;
            }
            subscriber.next(true);
        });

        return () => subscription.unsubscribe();
    });
}