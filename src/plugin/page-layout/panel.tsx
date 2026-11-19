import React from 'react';
import { of } from 'rxjs';
import type { Workbook } from '@univerjs/core';
import { ICommandService, IUniverInstanceService, LocaleService, UniverInstanceType } from '@univerjs/core';
import { Button, FormLayout, InputNumber, Select } from '@univerjs/design';
import { SheetsSelectionsService } from '@univerjs/sheets';
import { useDependency, useObservable } from '@univerjs/ui';
import { PageLayoutModel } from './model';
import type { IPageMargins, PageOrientation } from './types';
import {
    DEFAULT_PAGE_MARGINS,
    DEFAULT_PAGE_SETUP,
    ADD_PAGE_BREAK_COMMAND_ID,
    CLEAR_PAGE_BREAKS_COMMAND_ID,
    CLEAR_PRINT_AREA_COMMAND_ID,
    REMOVE_PAGE_BREAK_COMMAND_ID,
    SET_PRINT_AREA_COMMAND_ID,
    inchToCm,
    cmToInch,
} from './types';

const PAPER_SIZES = [
    { value: '9', label: 'A4' },
    { value: '8', label: 'A3' },
    { value: '11', label: 'A5' },
    { value: '1', label: 'Letter' },
    { value: '5', label: 'Legal' },
    { value: '7', label: 'Executive' },
    { value: '13', label: 'B5 (JIS)' },
    { value: '119', label: '16K' },
];

const MARGIN_FIELDS: Array<{ key: keyof IPageMargins; label: string }> = [
    { key: 'top', label: 'customMenu.pageLayoutMarginTop' },
    { key: 'bottom', label: 'customMenu.pageLayoutMarginBottom' },
    { key: 'left', label: 'customMenu.pageLayoutMarginLeft' },
    { key: 'right', label: 'customMenu.pageLayoutMarginRight' },
    { key: 'header', label: 'customMenu.pageLayoutMarginHeader' },
    { key: 'footer', label: 'customMenu.pageLayoutMarginFooter' },
];

function columnToLetter(column: number): string {
    let result = '';
    let value = column;
    while (value >= 0) {
        result = String.fromCharCode((value % 26) + 65) + result;
        value = Math.floor(value / 26) - 1;
    }
    return result;
}

function rangeToA1(range: { startRow: number; startColumn: number; endRow: number; endColumn: number }): string {
    const start = `${columnToLetter(range.startColumn)}${range.startRow + 1}`;
    const end = `${columnToLetter(range.endColumn)}${range.endRow + 1}`;
    return start === end ? start : `${start}:${end}`;
}

const sectionStyle: React.CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    gap: 8,
};

const sectionTitleStyle: React.CSSProperties = {
    fontSize: 13,
    fontWeight: 600,
};

const rowStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    flexWrap: 'wrap',
};

const hintStyle: React.CSSProperties = {
    fontSize: 12,
    opacity: 0.65,
};

export function PageLayoutPanel() {
    const localeService = useDependency(LocaleService);
    const commandService = useDependency(ICommandService);
    const instanceService = useDependency(IUniverInstanceService);
    const selectionService = useDependency(SheetsSelectionsService);
    const pageLayoutModel = useDependency(PageLayoutModel);

    const workbook = useObservable(
        () => instanceService.getCurrentTypeOfUnit$(UniverInstanceType.UNIVER_SHEET),
        undefined,
        undefined,
        []
    ) as Workbook | undefined;
    const worksheet = useObservable(
        () => workbook?.activeSheet$ ?? of(null),
        undefined,
        undefined,
        [workbook]
    );
    useObservable(() => pageLayoutModel.update$, undefined, undefined, []);
    useObservable(() => selectionService.selectionChanged$, undefined, undefined, []);

    const unitId = workbook?.getUnitId();
    const subUnitId = worksheet?.getSheetId();
    const data = unitId && subUnitId ? pageLayoutModel.getSheetData(unitId, subUnitId) : undefined;

    const margins = data?.margins ?? DEFAULT_PAGE_MARGINS;
    const pageSetup = data?.pageSetup ?? DEFAULT_PAGE_SETUP;
    const printArea = data?.printArea ?? [];
    const rowBreaks = data?.rowBreaks ?? [];
    const colBreaks = data?.colBreaks ?? [];
    const selection = selectionService.getCurrentLastSelection();

    const t = (key: string, ...args: string[]) => localeService.t(key, ...args);

    const updateMargins = (next: IPageMargins) => {
        if (!unitId || !subUnitId) return;
        pageLayoutModel.setMargins(unitId, subUnitId, next);
    };

    const updateOrientation = (orientation: PageOrientation) => {
        if (!unitId || !subUnitId) return;
        pageLayoutModel.setOrientation(unitId, subUnitId, orientation);
    };

    const updatePaperSize = (paperSize: number) => {
        if (!unitId || !subUnitId) return;
        pageLayoutModel.setPaperSize(unitId, subUnitId, paperSize);
    };

    const addBreak = (type: 'row' | 'col') => {
        commandService.executeCommand(ADD_PAGE_BREAK_COMMAND_ID, { type });
    };

    const removeBreak = (type: 'row' | 'col', index: number) => {
        commandService.executeCommand(REMOVE_PAGE_BREAK_COMMAND_ID, { type, index });
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 20, padding: '12px 16px 24px' }}>
            <section style={sectionStyle}>
                <div style={sectionTitleStyle}>{t('customMenu.pageLayoutPrintArea')}</div>
                <div style={hintStyle}>
                    {printArea.length
                        ? printArea.map((range) => rangeToA1(range)).join(', ')
                        : t('customMenu.pageLayoutNotSet')}
                </div>
                <div style={rowStyle}>
                    <Button
                        size="small"
                        variant="primary"
                        disabled={!selection?.range}
                        onClick={() => commandService.executeCommand(SET_PRINT_AREA_COMMAND_ID)}
                    >
                        {t('customMenu.printAreaSet')}
                    </Button>
                    <Button
                        size="small"
                        disabled={!printArea.length}
                        onClick={() => commandService.executeCommand(CLEAR_PRINT_AREA_COMMAND_ID)}
                    >
                        {t('customMenu.printAreaClear')}
                    </Button>
                </div>
            </section>

            <section style={sectionStyle}>
                <div style={sectionTitleStyle}>{t('customMenu.pageLayoutMargins')}</div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                    {MARGIN_FIELDS.map((field) => (
                        <label
                            key={field.key}
                            style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 12 }}
                        >
                            <span>{t(field.label)}</span>
                            <InputNumber
                                size="small"
                                min={0}
                                step={0.1}
                                precision={2}
                                controls={false}
                                value={inchToCm(margins[field.key])}
                                onChange={(value) => {
                                    if (value === null || value < 0) return;
                                    updateMargins({ ...margins, [field.key]: cmToInch(value) });
                                }}
                            />
                        </label>
                    ))}
                </div>
                <div style={hintStyle}>{t('customMenu.pageLayoutMarginUnit')}</div>
            </section>

            <section style={sectionStyle}>
                <div style={sectionTitleStyle}>{t('customMenu.pageLayoutPaper')}</div>
                <FormLayout label={t('customMenu.pageLayoutOrientation')}>
                    <Select
                        value={pageSetup.orientation}
                        options={[
                            { value: 'portrait', label: t('customMenu.pageLayoutPortrait') },
                            { value: 'landscape', label: t('customMenu.pageLayoutLandscape') },
                        ]}
                        onChange={(value) => updateOrientation(value as PageOrientation)}
                    />
                </FormLayout>
                <FormLayout label={t('customMenu.pageLayoutPaperSize')}>
                    <Select
                        value={String(pageSetup.paperSize)}
                        options={PAPER_SIZES}
                        onChange={(value) => updatePaperSize(Number(value))}
                    />
                </FormLayout>
            </section>

            <section style={sectionStyle}>
                <div style={sectionTitleStyle}>{t('customMenu.pageLayoutPageBreaks')}</div>
                <div style={rowStyle}>
                    <Button size="small" disabled={!selection?.range} onClick={() => addBreak('row')}>
                        {t('customMenu.pageLayoutAddRowBreak')}
                    </Button>
                    <Button size="small" disabled={!selection?.range} onClick={() => addBreak('col')}>
                        {t('customMenu.pageLayoutAddColBreak')}
                    </Button>
                </div>

                {rowBreaks.length === 0 && colBreaks.length === 0 ? (
                    <div style={hintStyle}>{t('customMenu.pageLayoutNoBreaks')}</div>
                ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                        {rowBreaks.map((index) => (
                            <div key={`row-${index}`} style={rowStyle}>
                                <span style={{ flex: 1, fontSize: 12 }}>
                                    {t('customMenu.pageLayoutBeforeRow', String(index + 1))}
                                </span>
                                <Button size="small" variant="text" onClick={() => removeBreak('row', index)}>
                                    {t('customMenu.delete')}
                                </Button>
                            </div>
                        ))}
                        {colBreaks.map((index) => (
                            <div key={`col-${index}`} style={rowStyle}>
                                <span style={{ flex: 1, fontSize: 12 }}>
                                    {t('customMenu.pageLayoutBeforeColumn', columnToLetter(index))}
                                </span>
                                <Button size="small" variant="text" onClick={() => removeBreak('col', index)}>
                                    {t('customMenu.delete')}
                                </Button>
                            </div>
                        ))}
                    </div>
                )}

                {rowBreaks.length > 0 || colBreaks.length > 0 ? (
                    <Button
                        size="small"
                        onClick={() => commandService.executeCommand(CLEAR_PAGE_BREAKS_COMMAND_ID)}
                    >
                        {t('customMenu.pageLayoutClearBreaks')}
                    </Button>
                ) : null}
            </section>
        </div>
    );
}
