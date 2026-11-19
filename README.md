# univer-custom-menu

Univer 自定义菜单插件，目前提供 **Excel 导入 / 导出**、**上下标（Superscript / Subscript）** 和 **页面布局（Page Layout）** 功能，并可以直接在 Vue / React 项目中使用。

## 功能

- 导入菜单：选择 Excel 文件后，解析并直接在 Univer 中创建工作表
- 导出菜单：将当前 Univer 工作簿导出为 Excel / CSV
- 支持条件格式、数据验证、筛选、冻结、超链接、传统批注、单元格图片/浮动图片等常见 Excel 特性
- 图表/透视表：暂不支持导入与导出
- 上下标插件：在工具栏中同时提供上标、下标两个按钮
- 页面布局插件：导入/导出打印区域、页边距、分页符、纸张方向/大小与打印缩放，支持在工具栏打开侧边栏进行设置，画布上以虚线显示打印区域和分页符，并随 Univer 快照以 JSON 形式保存
- 不依赖 `@univerjs/vite-plugin`，插件通过 `Plugin` 基类手动注册

## 目录结构

```text
src/
├── import_export/
│   ├── ToUniver/          # Excel -> IExcelXmlFile -> UniverWorkBook -> IWorkbookData
│   ├── UniverToExcel/     # Univer -> Excel
│   ├── UniverToCsv/       # Univer -> CSV
│   └── common/            # 公共工具（XML、EMF、常量等）
└── plugin/
    ├── import-export/     # 导入/导出菜单插件
    ├── super-subscript/   # 上下标插件
    └── page-layout/       # 页面布局插件（打印区域/页边距/分页符/纸张方向大小/打印缩放）
```

> 说明：代码现在直接从 Excel XML 解析为 `IExcelXmlFile`，再由 `UniverWorkBook` 生成 `IWorkbookData`，不再经过额外的中间层。

## 本地构建与打包

项目当前保持 `private: true`，不会发布到 npm，只做本地打包使用。

```bash
# 安装依赖
pnpm install

# 构建库产物（ESM + CJS + 类型声明）
pnpm build:lib

# 生成本地 tgz 包
pnpm pack --pack-destination /tmp
# 产物：/tmp/univer-custom-menu-0.0.1.tgz
```

构建输出：

```text
dist/lib/
├── index.js          # ESM
├── index.cjs         # CJS
└── types/            # 类型声明
```

## 在 Vue 项目中安装

```bash
pnpm add /tmp/univer-custom-menu-0.0.1.tgz
```

使用方项目需要自行安装以下 peer 依赖：

```bash
pnpm add @univerjs/core @univerjs/sheets @univerjs/ui react react-dom
```

同时请确保 Univer 的基础样式已正常引入：

```ts
import '@univerjs/design/lib/index.css';
import '@univerjs/ui/lib/index.css';
import '@univerjs/sheets-ui/lib/index.css';
```

## 注册插件

```ts
import { Univer } from '@univerjs/core';
import { UniverSheetsImportExportPlugin } from 'univer-custom-menu';

const univer = new Univer();

univer.registerPlugin(UniverSheetsImportExportPlugin, {
  import: {
    before: () => {
      console.log('准备导入');
    },
    after: ({ workbook, error, message }) => {
      if (error) {
        console.error(error);
      } else {
        console.log('导入完成', workbook, message);
      }
    },
  },
  export: {
    fileName: '导出文件.xlsx',
    before: () => {
      console.log('准备导出');
    },
    after: ({ error, message }) => {
      console.log('导出结束', error, message);
    },
    snapshot: () => {
      // 可选，自定义导出快照；不传时插件自动获取当前工作簿
      return currentWorkbookData;
    },
  },
});
```

## 注册上下标插件

```ts
import { UniverSheetsSuperSubscriptPlugin } from 'univer-custom-menu';

univer.registerPlugin(UniverSheetsSuperSubscriptPlugin, {
  superscript: {
    // 可选，默认使用内置语言包文案
    tooltip: '上标',
  },
  subscript: {
    tooltip: '下标',
  },
});
```

如果只需要其中一个按钮，把另一个传 `false` 即可隐藏：

```ts
univer.registerPlugin(UniverSheetsSuperSubscriptPlugin, {
  superscript: {},
  subscript: false, // 只显示上标
});
```

## 注册页面布局插件

```ts
import { UniverSheetsPageLayoutPlugin } from 'univer-custom-menu';

univer.registerPlugin(UniverSheetsPageLayoutPlugin, {
  // 可选：打印区域边界颜色与线宽，默认蓝色虚线
  stroke: '#2F6BFF',
  strokeWidth: 2,
  // 可选：分页符颜色
  breakStroke: '#2F6BFF',
  // showBorder: false,       // 只保存数据，不显示打印区域边界
  // showPageBreaks: false,   // 不显示分页符
});
```

使用方式：

1. 通过“导入”菜单打开 Excel 文件时，插件会自动解析打印区域、`pageMargins`、`pageSetup`（纸张方向/大小、打印缩放）以及 `rowBreaks`/`colBreaks`，并保存为 `SHEET_PAGE_LAYOUT_PLUGIN` 资源。
2. 点击工具栏「页面布局」打开侧边栏，可设置/清除打印区域、调整上/下/左/右/页眉/页脚页边距（界面单位为厘米）、切换纸张方向与大小、选择 4 种打印缩放模式，并在选中行/列处添加、删除或清空分页符。
3. 画布上会显示打印区域边框和分页符虚线。
4. 页面布局数据会随 `activeWorkbook.save()` 返回的 JSON 快照保存到 `resources` 中；导出 Excel 时会写回 `pageSetup.printArea`、`pageMargins`、`pageSetup.orientation/paperSize/scaleMode`、`rowBreaks` 和 `colBreaks`。打印缩放模式会写入 Excel 的 `sheetPr/pageSetUpPr@fitToPage` 与 `pageSetup@fitToWidth/fitToHeight/scale`。

数据 JSON 示例：

```json
{
  "resources": [
    {
      "name": "SHEET_PAGE_LAYOUT_PLUGIN",
      "data": "{\"<sheetId>\":{\"printArea\":[{\"startRow\":0,\"startColumn\":0,\"endRow\":9,\"endColumn\":3}],\"margins\":{\"top\":0.75,\"bottom\":0.75,\"left\":0.7,\"right\":0.7,\"header\":0.3,\"footer\":0.3},\"pageSetup\":{\"orientation\":\"portrait\",\"paperSize\":9,\"scaleMode\":\"none\"},\"rowBreaks\":[19],\"colBreaks\":[3]}}"
    }
  ]
}
```

> 页边距在资源中按 Excel 原生单位（英寸）保存，侧边栏展示与输入时自动换算为厘米；`rowBreaks`/`colBreaks` 保存 0 基的起始行/列索引。`scaleMode` 支持 `none`（无缩放）、`fitSheet`（将工作表调整为一页）、`fitColumns`（将所有列调整为一页）、`fitRows`（将所有行调整为一页）。

## 直接使用转换工具

插件内部使用 `ExcelXml` 完成转换，也可以直接调用：

```ts
import { ExcelXml } from 'univer-custom-menu';

// Excel -> Univer 工作簿数据
ExcelXml.transformExcelToUniver(file, (workbookData) => {
  univer.createUniver(workbookData);
});

// Univer 工作簿数据 -> Excel
await ExcelXml.transformUniverToExcel({
  snapshot: workbookData,
  fileName: '导出.xlsx',
});

// Univer 工作簿数据 -> CSV
await ExcelXml.transformUniverToCsv({
  snapshot: workbookData,
  fileName: '导出.csv',
});

// 如果需要中间 ExcelXml JSON 格式，仍可使用：
ExcelXml.transformExcelToExcelXml(file, (excelXmlFile, jsonString) => {
  // excelXmlFile: IExcelXmlFile 对象
  // jsonString: JSON 字符串
});
```

## 注意事项

- 包名为 `univer-custom-menu`，当前版本为 `0.0.1`
- `package.json` 中保持 `private: true`，仅用于本地打包
- `@univerjs/*`、`react`、`react-dom`、`exceljs` 等作为外部依赖，不会打进产物
- 如果要正式分发，需要把 `private` 改为 `false`，并更新包名与版本
