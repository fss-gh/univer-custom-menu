/** 上下标插件配置 key */
export const SUPER_SUBSCRIPT_PLUGIN_CONFIG = 'SUPER_SUBSCRIPT_PLUGIN_CONFIG';
/** 上下标插件名称 */
export const SUPER_SUBSCRIPT_PLUGIN_NAME = 'SUPER_SUBSCRIPT_PLUGIN';

/** 上标命令 / 菜单 ID */
export const SUPERSCRIPT_OPERATION_ID = 'custom-menu.operation.superscript';
/** 下标命令 / 菜单 ID */
export const SUBSCRIPT_OPERATION_ID = 'custom-menu.operation.subscript';

/** 单个上下标按钮的配置 */
export interface ISubScriptMenuConfig {
    /** 点击前回调；返回 false 可中止执行 */
    before?: () => void | Promise<boolean | void>;
    /** 执行完成后回调 */
    after?: (result?: boolean) => void;
    /** 自定义 tooltip / 标题文案 */
    tooltip?: string;
    /** 是否显示该按钮，默认显示 */
    show?: boolean;
}

/** 上下标插件配置 */
export interface ISuperSubscriptPluginConfig {
    /** 上标按钮配置；传 false 可隐藏上标按钮 */
    superscript?: ISubScriptMenuConfig | false;
    /** 下标按钮配置；传 false 可隐藏下标按钮 */
    subscript?: ISubScriptMenuConfig | false;
}
