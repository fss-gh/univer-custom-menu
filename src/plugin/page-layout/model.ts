import type { IRange } from '@univerjs/core';
import { Subject } from 'rxjs';
import type {
    IPageLayoutResourceData,
    IPageMargins,
    IPageSetup,
    ISheetPageLayoutData,
    PageOrientation,
    PageScaleMode,
} from './types';

/**
 * 页面布局数据模型：
 * - 以工作簿 unitId 为一级 key，工作表 subUnitId 为二级 key
 * - 数据直接写入 Univer 快照 resources 中的 SHEET_PAGE_LAYOUT_PLUGIN
 */
export class PageLayoutModel {
    private readonly _data = new Map<string, IPageLayoutResourceData>();

    private readonly _update$ = new Subject<{ unitId: string; subUnitId: string }>();

    /** 页面布局变更通知（设置/清除/加载/卸载） */
    readonly update$ = this._update$.asObservable();

    private _getSheetData(unitId: string, subUnitId: string): ISheetPageLayoutData {
        if (!this._data.has(unitId)) {
            this._data.set(unitId, {});
        }
        const unitData = this._data.get(unitId)!;
        if (!unitData[subUnitId]) {
            unitData[subUnitId] = {};
        }
        return unitData[subUnitId];
    }

    private _notify(unitId: string, subUnitId: string): void {
        this._update$.next({ unitId, subUnitId });
    }

    setPrintArea(unitId: string, subUnitId: string, ranges: IRange[]): void {
        if (!ranges.length) {
            this.clearPrintArea(unitId, subUnitId);
            return;
        }

        const data = this._getSheetData(unitId, subUnitId);
        data.printArea = ranges.map((r) => ({ ...r }));
        this._notify(unitId, subUnitId);
    }

    clearPrintArea(unitId: string, subUnitId: string): void {
        const data = this._data.get(unitId)?.[subUnitId];
        if (!data?.printArea) {
            return;
        }
        delete data.printArea;
        this._notify(unitId, subUnitId);
    }

    setMargins(unitId: string, subUnitId: string, margins: IPageMargins): void {
        const data = this._getSheetData(unitId, subUnitId);
        data.margins = { ...margins };
        this._notify(unitId, subUnitId);
    }

    setPageSetup(
        unitId: string,
        subUnitId: string,
        pageSetup: IPageSetup
    ): void {
        const data = this._getSheetData(unitId, subUnitId);
        data.pageSetup = { ...pageSetup };
        this._notify(unitId, subUnitId);
    }

    setOrientation(unitId: string, subUnitId: string, orientation: PageOrientation): void {
        const data = this._getSheetData(unitId, subUnitId);
        data.pageSetup = {
            paperSize: data.pageSetup?.paperSize ?? 9,
            ...data.pageSetup,
            orientation,
        };
        this._notify(unitId, subUnitId);
    }

    setPaperSize(unitId: string, subUnitId: string, paperSize: number): void {
        const data = this._getSheetData(unitId, subUnitId);
        data.pageSetup = {
            orientation: data.pageSetup?.orientation ?? 'portrait',
            ...data.pageSetup,
            paperSize,
        };
        this._notify(unitId, subUnitId);
    }

    setScaleMode(unitId: string, subUnitId: string, scaleMode: PageScaleMode): void {
        const data = this._getSheetData(unitId, subUnitId);
        data.pageSetup = {
            orientation: data.pageSetup?.orientation ?? 'portrait',
            paperSize: data.pageSetup?.paperSize ?? 9,
            ...data.pageSetup,
            scaleMode,
        };
        this._notify(unitId, subUnitId);
    }

    addPageBreak(unitId: string, subUnitId: string, type: 'row' | 'col', index: number): void {
        const data = this._getSheetData(unitId, subUnitId);
        const key = type === 'row' ? 'rowBreaks' : 'colBreaks';
        const breaks = data[key] ?? [];
        if (!breaks.includes(index)) {
            breaks.push(index);
            breaks.sort((a, b) => a - b);
            data[key] = breaks;
            this._notify(unitId, subUnitId);
        }
    }

    removePageBreak(unitId: string, subUnitId: string, type: 'row' | 'col', index: number): void {
        const data = this._data.get(unitId)?.[subUnitId];
        const key = type === 'row' ? 'rowBreaks' : 'colBreaks';
        const breaks = data?.[key];
        if (!breaks) {
            return;
        }
        const next = breaks.filter((item) => item !== index);
        if (next.length === 0) {
            delete data![key];
        } else {
            data![key] = next;
        }
        this._notify(unitId, subUnitId);
    }

    clearPageBreaks(unitId: string, subUnitId: string): void {
        const data = this._data.get(unitId)?.[subUnitId];
        if (!data?.rowBreaks && !data?.colBreaks) {
            return;
        }
        delete data.rowBreaks;
        delete data.colBreaks;
        this._notify(unitId, subUnitId);
    }

    getSheetData(unitId: string, subUnitId: string): ISheetPageLayoutData | undefined {
        return this._data.get(unitId)?.[subUnitId];
    }

    getUnitData(unitId: string): IPageLayoutResourceData {
        return this._data.get(unitId) ?? {};
    }

    load(unitId: string, data: IPageLayoutResourceData): void {
        this._data.set(unitId, data || {});
        this._notify(unitId, '');
    }

    unload(unitId: string): void {
        this._data.delete(unitId);
        this._notify(unitId, '');
    }
}
