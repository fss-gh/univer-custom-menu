import type { IRange } from '@univerjs/core';
import { Subject } from 'rxjs';
import type { IPrintAreaResourceData, ISheetPrintAreaData } from './types';

/**
 * 打印区域数据模型：
 * - 以工作簿 unitId 为一级 key，工作表 subUnitId 为二级 key
 * - 数据直接写入 Univer 快照 resources 中的 SHEET_PRINT_AREA_PLUGIN
 */
export class PrintAreaModel {
    private readonly _data = new Map<string, IPrintAreaResourceData>();

    private readonly _update$ = new Subject<{ unitId: string; subUnitId: string }>();

    /** 打印区域变更通知（设置/清除/加载/卸载） */
    readonly update$ = this._update$.asObservable();

    setPrintArea(unitId: string, subUnitId: string, ranges: IRange[]): void {
        if (!ranges.length) {
            this.clearPrintArea(unitId, subUnitId);
            return;
        }

        if (!this._data.has(unitId)) {
            this._data.set(unitId, {});
        }

        this._data.get(unitId)![subUnitId] = {
            ranges: ranges.map((r) => ({ ...r })),
        };

        this._update$.next({ unitId, subUnitId });
    }

    clearPrintArea(unitId: string, subUnitId: string): void {
        const unitData = this._data.get(unitId);
        if (!unitData) {
            return;
        }

        if (delete unitData[subUnitId]) {
            this._update$.next({ unitId, subUnitId });
        }
    }

    getSheetPrintArea(unitId: string, subUnitId: string): ISheetPrintAreaData | undefined {
        return this._data.get(unitId)?.[subUnitId];
    }

    getUnitData(unitId: string): IPrintAreaResourceData {
        return this._data.get(unitId) ?? {};
    }

    load(unitId: string, data: IPrintAreaResourceData): void {
        this._data.set(unitId, data || {});
        this._update$.next({ unitId, subUnitId: '' });
    }

    unload(unitId: string): void {
        this._data.delete(unitId);
        this._update$.next({ unitId, subUnitId: '' });
    }
}
