import { Dependency, IConfigService, Inject, Injector, LocaleService, Plugin, UniverInstanceType } from '@univerjs/core';
import zhCN from '../locale/zh-CN';
import enUS from '../locale/en-US';
import { SuperSubscriptMenuController } from './controller';
import { ISuperSubscriptPluginConfig, SUPER_SUBSCRIPT_PLUGIN_CONFIG, SUPER_SUBSCRIPT_PLUGIN_NAME } from './types';

const AllSuperSubscriptController = [SuperSubscriptMenuController];

/**
 * 上下标插件：在工具栏中同时注册「上标」和「下标」两个按钮。
 *
 * 在 Vue/React 项目中注册到 Univer：
 * ```ts
 * univer.registerPlugin(UniverSheetsSuperSubscriptPlugin, {
 *   superscript: { tooltip: '上标' },
 *   subscript: { tooltip: '下标' },
 * });
 * ```
 */
export class UniverSheetsSuperSubscriptPlugin extends Plugin {
    static override type = UniverInstanceType.UNIVER_SHEET;
    static override pluginName = SUPER_SUBSCRIPT_PLUGIN_NAME;

    constructor(
        readonly config: ISuperSubscriptPluginConfig,
        @Inject(Injector) protected readonly _injector: Injector,
        @Inject(LocaleService) private readonly _localeService: LocaleService,
        @IConfigService private readonly _configService: IConfigService
    ) {
        super();

        this._localeService.load({ zhCN, enUS });
        this._configService.setConfig(SUPER_SUBSCRIPT_PLUGIN_CONFIG, config || {});
    }

    override onStarting(): void {
        ([AllSuperSubscriptController] as Dependency[]).forEach((d) =>
            this._injector.add(d)
        );
    }

    override onReady(): void {
        this._injector.get(SuperSubscriptMenuController);
    }
}
