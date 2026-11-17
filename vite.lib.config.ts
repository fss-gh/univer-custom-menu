import { fileURLToPath, URL } from 'node:url';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

/**
 * 本地库构建配置：
 * 把「导入 / 导出菜单插件」打成 ESM + CJS 两个产物，
 * Univer、React、ExcelJS 等作为外部依赖，由使用方项目提供。
 */
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist/lib',
    lib: {
      entry: fileURLToPath(new URL('./src/index.ts', import.meta.url)),
      name: 'UniverImportExportMenu',
      formats: ['es', 'cjs'],
      fileName: (format) => (format === 'es' ? 'index.js' : 'index.cjs'),
    },
    rollupOptions: {
      external: [
        /^@univerjs\//,
        'react',
        'react-dom',
        'react/jsx-runtime',
        'rxjs',
        '@zwight/exceljs',
        'papaparse',
        '@progress/jszip-esm',
        'dayjs',
        'nanoid',
      ],
    },
    emptyOutDir: true,
    copyPublicDir: false,
    sourcemap: true,
  },
});
