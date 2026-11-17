import "./style.css";
import "@univerjs/design/lib/index.css";
import "@univerjs/ui/lib/index.css";
import "@univerjs/docs-ui/lib/index.css";
import "@univerjs/sheets-ui/lib/index.css";
import "@univerjs/drawing-ui/lib/index.css";
import "@univerjs/find-replace/lib/index.css";
import "@univerjs/sheets-conditional-formatting-ui/lib/index.css";
import "@univerjs/sheets-crosshair-highlight/lib/index.css";
import "@univerjs/sheets-data-validation-ui/lib/index.css";
import "@univerjs/sheets-drawing-ui/lib/index.css";
import "@univerjs/sheets-filter-ui/lib/index.css";
import "@univerjs/sheets-formula-ui/lib/index.css";
import "@univerjs/sheets-hyper-link-ui/lib/index.css";
import "@univerjs/sheets-note-ui/lib/index.css";
import "@univerjs/sheets-numfmt-ui/lib/index.css";
import "@univerjs/sheets-sort-ui/lib/index.css";
import "@univerjs/sheets-table-ui/lib/index.css";
import "@univerjs/sheets-thread-comment-ui/lib/index.css";
import "@univerjs/thread-comment-ui/lib/index.css";
import "@univerjs-pro/sheets-print/lib/index.css";

import { LocaleType, Univer, UniverInstanceType } from "@univerjs/core";
import { defaultTheme } from "@univerjs/themes";
import { UniverDocsPlugin } from "@univerjs/docs";
import { UniverDocsUIPlugin } from "@univerjs/docs-ui";
import { UniverRenderEnginePlugin } from "@univerjs/engine-render";
import { UniverSheetsPlugin } from "@univerjs/sheets";
//数值格式
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsNumfmtUIPlugin } from '@univerjs/sheets-numfmt-ui';
//公式
import { UniverFormulaEnginePlugin } from "@univerjs/engine-formula";
import { UniverSheetsFormulaPlugin } from "@univerjs/sheets-formula";
import { UniverSheetsFormulaUIPlugin } from "@univerjs/sheets-formula-ui";
// 筛选
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter';
import { UniverSheetsFilterUIPlugin } from '@univerjs/sheets-filter-ui';
import '@univerjs/sheets-filter/facade'
//排序
import { UniverSheetsSortPlugin } from '@univerjs/sheets-sort'
import { UniverSheetsSortUIPlugin } from '@univerjs/sheets-sort-ui'
import '@univerjs/sheets-sort/facade'
//数据验证
import { UniverDataValidationPlugin } from '@univerjs/data-validation'
import { UniverSheetsDataValidationPlugin } from '@univerjs/sheets-data-validation'
import { UniverSheetsDataValidationUIPlugin } from '@univerjs/sheets-data-validation-ui'
import '@univerjs/sheets-data-validation/facade'
//条件格式
import { UniverSheetsConditionalFormattingPlugin } from '@univerjs/sheets-conditional-formatting'
import { UniverSheetsConditionalFormattingUIPlugin } from '@univerjs/sheets-conditional-formatting-ui'
import '@univerjs/sheets-conditional-formatting/facade'
//超链接
import { UniverSheetsHyperLinkPlugin } from '@univerjs/sheets-hyper-link'
import { UniverSheetsHyperLinkUIPlugin } from '@univerjs/sheets-hyper-link-ui'
import '@univerjs/sheets-hyper-link/facade'
import '@univerjs/sheets-hyper-link-ui/facade'
//查找替换
import { UniverFindReplacePlugin } from "@univerjs/find-replace";
import { UniverSheetsFindReplacePlugin } from '@univerjs/sheets-find-replace'
import '@univerjs/sheets-find-replace/facade'
//图片
import { UniverDrawingPlugin } from '@univerjs/drawing'
import { UniverDrawingUIPlugin } from '@univerjs/drawing-ui'
import { UniverSheetsDrawingPlugin } from '@univerjs/sheets-drawing'
import { UniverSheetsDrawingUIPlugin } from '@univerjs/sheets-drawing-ui'
import '@univerjs/sheets-drawing-ui/facade'
//评论
import { UniverSheetsThreadCommentPlugin } from '@univerjs/sheets-thread-comment'
import { UniverSheetsThreadCommentUIPlugin } from '@univerjs/sheets-thread-comment-ui'
import { UniverThreadCommentPlugin } from '@univerjs/thread-comment'
import { UniverThreadCommentUIPlugin } from '@univerjs/thread-comment-ui'
import '@univerjs/sheets-thread-comment/facade'
//批注
import { UniverSheetsNotePlugin } from '@univerjs/sheets-note'
import { UniverSheetsNoteUIPlugin } from '@univerjs/sheets-note-ui'
import '@univerjs/sheets-note/facade'
//表格
import { UniverSheetsTablePlugin } from '@univerjs/sheets-table'
import { UniverSheetsTableUIPlugin } from '@univerjs/sheets-table-ui'
import '@univerjs/sheets-table/facade'
//水印（当前未启用，注释在下方，启用时再恢复 import）
//十字高亮
import { UniverSheetsCrosshairHighlightPlugin } from '@univerjs/sheets-crosshair-highlight'
import '@univerjs/sheets-crosshair-highlight/facade'
//禅模式/Uniscript：0.25.2 与 1.5.2 不兼容（暂无 1.x 版本），已移除；后续官方发布适配版本后再恢复
//打印
import { UniverLicensePlugin } from '@univerjs-pro/license';
import { UniverSheetsPrintPlugin } from '@univerjs-pro/sheets-print'
//UI
import { UniverSheetsUIPlugin } from "@univerjs/sheets-ui";
import { UniverUIPlugin } from "@univerjs/ui";
// 自定义插件
import { UniverSheetsCustomMenuPlugin } from "./plugin";
import { UniverSheetsImportExportPlugin, UniverSheetsSuperSubscriptPlugin } from "./index";
import { FUniver } from '@univerjs/core/facade';
//引入Facade API 实现
import '@univerjs/ui/facade'
import '@univerjs/docs-ui/facade'
import '@univerjs/sheets/facade'
/**
 * 语言包：当前未使用 @univerjs/vite-plugin，改为从各插件包手动导入并合并（见 ./locales）。
 * 样式也改为在文件顶部手动引入各包的 CSS。
 */
import { zhCN, enUS } from './locales'
import CustomSaveMenu from "./plugin/controllers/menu/save.menu";
import { CustomDropdownListMenu } from "./plugin/controllers/menu/dropdown-list.menu";


const univer = new Univer({
  theme: defaultTheme,
  locale: LocaleType.ZH_CN,
  locales: {
    [LocaleType.ZH_CN]: zhCN,
    [LocaleType.EN_US]: enUS,
  },
});

univer.registerPlugin(UniverRenderEnginePlugin);

univer.registerPlugin(UniverUIPlugin, {
  container: 'app',
});

univer.registerPlugin(UniverDocsPlugin, {
  hasScroll: false,
});
univer.registerPlugin(UniverDocsUIPlugin);

univer.registerPlugin(UniverSheetsPlugin);
univer.registerPlugin(UniverSheetsUIPlugin);

univer.registerPlugin(UniverFormulaEnginePlugin);
univer.registerPlugin(UniverSheetsFormulaPlugin);
univer.registerPlugin(UniverSheetsFormulaUIPlugin);

univer.registerPlugin(UniverSheetsFilterPlugin);
univer.registerPlugin(UniverSheetsFilterUIPlugin);

univer.registerPlugin(UniverSheetsSortPlugin);
univer.registerPlugin(UniverSheetsSortUIPlugin);

univer.registerPlugin(UniverDataValidationPlugin)
univer.registerPlugin(UniverSheetsDataValidationPlugin)
univer.registerPlugin(UniverSheetsDataValidationUIPlugin)

univer.registerPlugin(UniverSheetsConditionalFormattingPlugin)
univer.registerPlugin(UniverSheetsConditionalFormattingUIPlugin)


univer.registerPlugin(UniverSheetsHyperLinkPlugin)
univer.registerPlugin(UniverSheetsHyperLinkUIPlugin)

univer.registerPlugin(UniverFindReplacePlugin)
univer.registerPlugin(UniverSheetsFindReplacePlugin)

univer.registerPlugin(UniverSheetsNumfmtPlugin);
univer.registerPlugin(UniverSheetsNumfmtUIPlugin);

univer.registerPlugin(UniverDrawingPlugin);
univer.registerPlugin(UniverDrawingUIPlugin);
univer.registerPlugin(UniverSheetsDrawingPlugin);
univer.registerPlugin(UniverSheetsDrawingUIPlugin);

univer.registerPlugin(UniverThreadCommentPlugin)
univer.registerPlugin(UniverThreadCommentUIPlugin)
univer.registerPlugin(UniverSheetsThreadCommentPlugin)
univer.registerPlugin(UniverSheetsThreadCommentUIPlugin)

univer.registerPlugin(UniverSheetsNotePlugin)
univer.registerPlugin(UniverSheetsNoteUIPlugin)

univer.registerPlugin(UniverSheetsTablePlugin)
univer.registerPlugin(UniverSheetsTableUIPlugin)

univer.registerPlugin(UniverSheetsCrosshairHighlightPlugin)

// univer.registerPlugin(UniverWatermarkPlugin, {
//   textWatermarkSettings: {
//     content: 'Hello, Univer!',
//     fontSize: 36,
//   },
// })

const univerAPI = FUniver.newAPI(univer);
const getUniverSnapshot = () => {
  const activeWorkbook = univerAPI.getActiveWorkbook();
  if (!activeWorkbook) {
    throw new Error('Workbook is not initialized');
  }
  // 返回 IWorkbookData 快照，供导出/保存逻辑使用
  return activeWorkbook.save();
};
// 导入/导出专用插件
univer.registerPlugin(UniverSheetsImportExportPlugin, {
  import: {
    before: () => {
      // loading.value = true;
    },
    after: () => {
      // loading.value = false;
      // if (error) {
      //   message.error(error.message || '导入失败');
      //   return;
      // }
      // message.success('导入成功');
    },
  },
  export: {
    snapshot: getUniverSnapshot,
    after: () => {
      // loading.value = false;
      // if (res) {
      //   message.error(res.message || '导出失败');
      // } else {
      //   message.success('导出成功');
      // }
    },
  },
});

// 上下标插件（上标 + 下标）
univer.registerPlugin(UniverSheetsSuperSubscriptPlugin, {});

// 其余自定义菜单（保存/下拉列表）
univer.registerPlugin(UniverSheetsCustomMenuPlugin, {
  instance: univer,
  menu: [
    CustomSaveMenu({
      after: () => {
        const saveData = getUniverSnapshot();
        console.log(saveData);
      },
    }),
    CustomDropdownListMenu(),
  ],
});
// create univer sheet instance
univer.createUnit(UniverInstanceType.UNIVER_SHEET, {});
