import { Disposable, IConfigService, Inject, type Nullable, Workbook } from '@univerjs/core';
import type { IRenderContext, IRenderModule } from '@univerjs/engine-render';
import { DashedRect } from '@univerjs/engine-render';
import { attachSelectionWithCoord, ISheetSkeletonManagerParam, ISelectionWithStyle } from '@univerjs/sheets';
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui';
import { PrintAreaModel } from './model';
import type { IPrintAreaPluginConfig } from './types';
import { PRINT_AREA_PLUGIN_CONFIG } from './types';

/**
 * 打印区域渲染控制器：
 * - 监听当前工作表 skeleton 变化和打印区域数据变化
 * - 把每个打印区域画成一个蓝色虚线矩形
 */
export class PrintAreaRenderController extends Disposable implements IRenderModule {
    private readonly _rects: DashedRect[] = [];

    private _currentSkeletonParam: Nullable<ISheetSkeletonManagerParam> = null;

    constructor(
        private readonly _context: IRenderContext<Workbook>,
        @Inject(SheetSkeletonManagerService)
        private readonly _sheetSkeletonManagerService: SheetSkeletonManagerService,
        @Inject(PrintAreaModel)
        private readonly _printAreaModel: PrintAreaModel,
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
            this._printAreaModel.update$.subscribe(() => {
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

        const printArea = this._printAreaModel.getSheetPrintArea(param.unitId, param.sheetId);
        this._clearRects();
        if (!printArea?.ranges?.length) {
            return;
        }

        const config = this._configService.getConfig<IPrintAreaPluginConfig>(PRINT_AREA_PLUGIN_CONFIG) ?? {};
        if (config.showBorder === false) {
            return;
        }

        const stroke = config.stroke || '#2F6BFF';
        const strokeWidth = config.strokeWidth ?? 2;

        printArea.ranges.forEach((range, index) => {
            const selection: ISelectionWithStyle = { range, primary: null, style: null };
            const rangeWithCoord = attachSelectionWithCoord(selection, param.skeleton).rangeWithCoord;

            const rect = new DashedRect(`PRINT_AREA_${index}`, {
                left: rangeWithCoord.startX,
                top: rangeWithCoord.startY,
                width: rangeWithCoord.endX - rangeWithCoord.startX,
                height: rangeWithCoord.endY - rangeWithCoord.startY,
                stroke,
                strokeWidth,
                strokeDashArray: [4, 4],
                evented: false,
                zIndex: 3,
                forceRender: true,
            });

            this._context.scene.addObject(rect, 3);
            this._rects.push(rect);
        });
    }

    private _clearRects(): void {
        this._rects.forEach((rect) => {
            this._context.scene.removeObject(rect);
            rect.dispose();
        });
        this._rects.length = 0;
    }
}
