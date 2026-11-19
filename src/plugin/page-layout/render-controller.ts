import { Disposable, IConfigService, Inject, type Nullable, Workbook } from '@univerjs/core';
import type { IRange } from '@univerjs/core';
import type { IRenderContext, IRenderModule } from '@univerjs/engine-render';
import { DashedRect } from '@univerjs/engine-render';
import { attachSelectionWithCoord, ISheetSkeletonManagerParam, ISelectionWithStyle } from '@univerjs/sheets';
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui';
import { PageLayoutModel } from './model';
import type { IPageLayoutPluginConfig } from './types';
import { PAGE_LAYOUT_PLUGIN_CONFIG } from './types';

/**
 * 页面布局渲染控制器：
 * - 监听当前工作表 skeleton 和页面布局数据变化
 * - 打印区域绘制为蓝色虚线矩形
 * - 分页符绘制为贯穿当前视图的蓝色虚线
 */
export class PageLayoutRenderController extends Disposable implements IRenderModule {
    private readonly _rects: DashedRect[] = [];

    private _currentSkeletonParam: Nullable<ISheetSkeletonManagerParam> = null;

    constructor(
        private readonly _context: IRenderContext<Workbook>,
        @Inject(SheetSkeletonManagerService)
        private readonly _sheetSkeletonManagerService: SheetSkeletonManagerService,
        @Inject(PageLayoutModel)
        private readonly _pageLayoutModel: PageLayoutModel,
        @IConfigService
        private readonly _configService: IConfigService
    ) {
        super();

        this._init();
    }

    private _init(): void {
        this.disposeWithMe(
            this._sheetSkeletonManagerService.currentSkeleton$.subscribe((param) => {
                this._currentSkeletonParam = param || null;
                this._update();
            })
        );

        this.disposeWithMe(
            this._pageLayoutModel.update$.subscribe(() => {
                this._update();
            })
        );
    }

    override dispose(): void {
        this._clearRects();
        super.dispose();
    }

    private _update(): void {
        const param = this._currentSkeletonParam;
        if (!param) {
            return;
        }

        const data = this._pageLayoutModel.getSheetData(param.unitId, param.sheetId);
        this._clearRects();
        if (!data) {
            return;
        }

        const config = this._configService.getConfig<IPageLayoutPluginConfig>(PAGE_LAYOUT_PLUGIN_CONFIG) ?? {};
        const stroke = config.stroke || '#2F6BFF';
        const strokeWidth = config.strokeWidth ?? 2;

        if (config.showBorder !== false && data.printArea?.length) {
            data.printArea.forEach((range, index) => {
                const selection: ISelectionWithStyle = { range, primary: null, style: null };
                const rangeWithCoord = attachSelectionWithCoord(selection, param.skeleton).rangeWithCoord;

                this._addRect(`PAGE_LAYOUT_PRINT_AREA_${index}`, {
                    left: rangeWithCoord.startX,
                    top: rangeWithCoord.startY,
                    width: rangeWithCoord.endX - rangeWithCoord.startX,
                    height: rangeWithCoord.endY - rangeWithCoord.startY,
                    stroke,
                    strokeWidth,
                });
            });
        }

        if (config.showPageBreaks !== false) {
            const breakStroke = config.breakStroke || '#2F6BFF';
            const scene = this._context.scene;

            (data.rowBreaks ?? []).forEach((row, index) => {
                const y = this._getBoundaryCoordinate(param, row, 0).startY;
                this._addRect(`PAGE_LAYOUT_ROW_BREAK_${index}`, {
                    left: 0,
                    top: y,
                    width: scene.width,
                    height: 1,
                    stroke: breakStroke,
                    strokeWidth: 1,
                    strokeDashArray: [6, 4],
                });
            });

            (data.colBreaks ?? []).forEach((col, index) => {
                const x = this._getBoundaryCoordinate(param, 0, col).startX;
                this._addRect(`PAGE_LAYOUT_COL_BREAK_${index}`, {
                    left: x,
                    top: 0,
                    width: 1,
                    height: scene.height,
                    stroke: breakStroke,
                    strokeWidth: 1,
                    strokeDashArray: [6, 4],
                });
            });
        }
    }

    private _getBoundaryCoordinate(
        param: ISheetSkeletonManagerParam,
        row: number,
        col: number
    ): { startX: number; startY: number } {
        const range: IRange = { startRow: row, endRow: row, startColumn: col, endColumn: col };
        const selection: ISelectionWithStyle = { range, primary: null, style: null };
        return attachSelectionWithCoord(selection, param.skeleton).rangeWithCoord;
    }

    private _addRect(
        key: string,
        options: {
            left: number;
            top: number;
            width: number;
            height: number;
            stroke: string;
            strokeWidth: number;
            strokeDashArray?: number[];
        }
    ): void {
        const rect = new DashedRect(key, {
            ...options,
            strokeDashArray: options.strokeDashArray ?? [4, 4],
            evented: false,
            zIndex: 3,
            forceRender: true,
        });

        this._context.scene.addObject(rect, 3);
        this._rects.push(rect);
    }

    private _clearRects(): void {
        this._rects.forEach((rect) => {
            this._context.scene.removeObject(rect);
            rect.dispose();
        });
        this._rects.length = 0;
    }
}
