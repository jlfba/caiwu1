import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
  base: './',
  server: {
    host: '0.0.0.0',
    port: 59323,
    proxy: {
      '/api': 'http://127.0.0.1:15618'
    }
  },
  build: {
    outDir: 'dist'
  }
})
