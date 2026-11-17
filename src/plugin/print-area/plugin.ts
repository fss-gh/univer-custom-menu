import { Dependency, IConfigService, Inject, Injector, LocaleService, Plugin, UniverInstanceType } from '@univerjs/core';
import { IRenderManagerService } from '@univerjs/engine-render';
import zhCN from '../locale/zh-CN';
import enUS from '../locale/en-US';
import { PrintAreaMenuController } from './controller';
import { PrintAreaModel } from './model';
import { PrintAreaRenderController } from './render-controller';
import type { IPrintAreaPluginConfig } from './types';
import { PRINT_AREA_PLUGIN_CONFIG, PRINT_AREA_PLUGIN_NAME } from './types';

const AllPrintAreaController = [PrintAreaMenuController];

/**
 * 打印区域自定义插件：
 * - 导入 Excel 时解析 `_xlnm.Print_Area`，在界面上以蓝色虚线显示打印区域边界
 * - 工具栏提供「设置打印区域 / 清除打印区域」按钮
 * - 打印区域数据保存在 Univer 快照 resources 的 SHEET_PRINT_AREA_PLUGIN 中
 *
 * 在 Vue/React 项目中注册到 Univer：
 * ```ts
 * univer.registerPlugin(UniverSheetsPrintAreaPlugin, {
 *   stroke: '#2F6BFF',
 *   strokeWidth: 2,
 * });
 * ```
 */
export class UniverSheetsPrintAreaPlugin extends Plugin {
    static override type = UniverInstanceType.UNIVER_SHEET;
    static override pluginName = PRINT_AREA_PLUGIN_NAME;

    constructor(
        readonly config: IPrintAreaPluginConfig,
        @Inject(Injector) protected readonly _injector: Injector,
        @Inject(LocaleService) private readonly _localeService: LocaleService,
        @IConfigService private readonly _configService: IConfigService,
        @IRenderManagerService private readonly _renderManagerService: IRenderManagerService
    ) {
        super();

        this._localeService.load({ zhCN, enUS });
        this._configService.setConfig(PRINT_AREA_PLUGIN_CONFIG, config || {});
    }

    override onStarting(): void {
        ([AllPrintAreaController, [PrintAreaModel]] as Dependency[]).forEach((d) =>
            this._injector.add(d)
        );
    }

    override onRendered(): void {
        // 注册渲染模块：负责在表格画布上绘制打印区域蓝色虚线边界
        this.disposeWithMe(
            this._renderManagerService.registerRenderModule(UniverInstanceType.UNIVER_SHEET, [
                PrintAreaRenderController,
            ])
        );
    }

    override onReady(): void {
        this._injector.get(PrintAreaMenuController);
    }
}
