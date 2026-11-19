import { Dependency, IConfigService, Inject, Injector, LocaleService, Plugin, UniverInstanceType } from '@univerjs/core';
import { IRenderManagerService } from '@univerjs/engine-render';
import zhCN from '../locale/zh-CN';
import enUS from '../locale/en-US';
import { PageLayoutMenuController } from './controller';
import { PageLayoutModel } from './model';
import { PageLayoutRenderController } from './render-controller';
import type { IPageLayoutPluginConfig } from './types';
import { PAGE_LAYOUT_PLUGIN_CONFIG, PAGE_LAYOUT_PLUGIN_NAME } from './types';

const AllPageLayoutController = [PageLayoutMenuController];

/**
 * 页面布局自定义插件：
 * - 导入 Excel 时解析打印区域、页边距、纸张方向/大小、打印缩放和分页符
 * - 工具栏提供「页面布局」入口，侧边栏可设置打印区域、页边距、分页符、纸张方向/大小和打印缩放
 * - 画布上显示打印区域边框和分页符虚线
 * - 页面布局数据保存在 Univer 快照 resources 的 SHEET_PAGE_LAYOUT_PLUGIN 中
 *
 * 在 Vue/React 项目中注册到 Univer：
 * ```ts
 * univer.registerPlugin(UniverSheetsPageLayoutPlugin, {
 *   stroke: '#2F6BFF',
 *   strokeWidth: 2,
 * });
 * ```
 */
export class UniverSheetsPageLayoutPlugin extends Plugin {
    static override type = UniverInstanceType.UNIVER_SHEET;
    static override pluginName = PAGE_LAYOUT_PLUGIN_NAME;

    constructor(
        readonly config: IPageLayoutPluginConfig,
        @Inject(Injector) protected readonly _injector: Injector,
        @Inject(LocaleService) private readonly _localeService: LocaleService,
        @IConfigService private readonly _configService: IConfigService,
        @IRenderManagerService private readonly _renderManagerService: IRenderManagerService
    ) {
        super();

        this._localeService.load({ zhCN, enUS });
        this._configService.setConfig(PAGE_LAYOUT_PLUGIN_CONFIG, config || {});
    }

    override onStarting(): void {
        ([AllPageLayoutController, [PageLayoutModel]] as Dependency[]).forEach((d) =>
            this._injector.add(d)
        );
    }

    override onRendered(): void {
        this.disposeWithMe(
            this._renderManagerService.registerRenderModule(UniverInstanceType.UNIVER_SHEET, [
                PageLayoutRenderController,
            ])
        );
    }

    override onReady(): void {
        this._injector.get(PageLayoutMenuController);
    }
}
