# univer-custom-menu

Univer 自定义菜单插件，目前提供 **Excel 导入 / 导出** 和 **上下标（Superscript / Subscript）** 功能，并可以直接在 Vue / React 项目中使用。

## 功能

- 导入菜单：选择 Excel 文件后，解析并直接在 Univer 中创建工作表
- 导出菜单：将当前 Univer 工作簿导出为 Excel / CSV
- 支持条件格式、数据验证、筛选、图表、冻结、超链接等常见 Excel 特性
- 上下标插件：在工具栏中同时提供上标、下标两个按钮
- 不依赖 `@univerjs/vite-plugin`，插件通过 `Plugin` 基类手动注册

## 目录结构

```text
src/
├── import_export/
│   ├── ToUniver/          # Excel -> ILuckyFile -> UniverWorkBook -> IWorkbookData
│   ├── UniverToExcel/     # Univer -> Excel
│   ├── UniverToCsv/       # Univer -> CSV
│   └── common/            # 公共工具（XML、EMF、常量等）
└── plugin/
    ├── import-export/     # 导入/导出菜单插件
    └── super-subscript/   # 上下标插件
```

> 说明：已移除原来的 `ToLuckySheet` / `LuckyToUniver` 中间层，`LuckyFile.Parse()` 现在直接返回 `ILuckyFile` 对象，再由 `UniverWorkBook` 生成 `IWorkbookData`，逻辑更少一层。

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

## 直接使用转换工具

插件内部使用 `LuckyExcel` 完成转换，也可以直接调用：

```ts
import { LuckyExcel } from 'univer-custom-menu';

// Excel -> Univer 工作簿数据
LuckyExcel.transformExcelToUniver(file, (workbookData) => {
  univer.createUniver(workbookData);
});

// Univer 工作簿数据 -> Excel
await LuckyExcel.transformUniverToExcel({
  snapshot: workbookData,
  fileName: '导出.xlsx',
});

// Univer 工作簿数据 -> CSV
await LuckyExcel.transformUniverToCsv({
  snapshot: workbookData,
  fileName: '导出.csv',
});

// 如果需要旧的 Luckysheet JSON 格式，仍可使用：
LuckyExcel.transformExcelToLucky(file, (luckysheetFile, jsonString) => {
  // luckysheetFile: ILuckyFile 对象
  // jsonString: JSON 字符串
});
```

## 注意事项

- 包名为 `univer-custom-menu`，当前版本为 `0.0.1`
- `package.json` 中保持 `private: true`，仅用于本地打包
- `@univerjs/*`、`react`、`react-dom`、`exceljs` 等作为外部依赖，不会打进产物
- 如果要正式分发，需要把 `private` 改为 `false`，并更新包名与版本
