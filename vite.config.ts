import { defineConfig } from "vite";
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [
    react()
  ],
  build: {
    // Demo 应用与库产物分开：应用输出到 dist/app，库输出到 dist/lib
    outDir: 'dist/app'
  }
});
