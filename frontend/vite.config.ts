import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // 本番ビルドから console.* / debugger を除去し、ソースマップも出力しない
  // (ブラウザの開発者コンソールに内部情報やソースを露出させないため)
  esbuild: {
    drop: ['console', 'debugger'],
  },
  build: {
    sourcemap: false,
  },
  server: {
    port: 5173,
    proxy: {
      '/api': { target: 'http://localhost:8000', changeOrigin: true },
      '/ws': { target: 'ws://localhost:8000', ws: true },
    },
  },
})
