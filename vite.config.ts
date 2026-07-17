import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    // Delegated Azure import uses a same-window MSAL redirect, so retain strict
    // opener isolation in local development as well as production.
    headers: {
      'Cross-Origin-Opener-Policy': 'same-origin',
    },
    proxy: {
      // Forward /api/ to the local speech token server (mirrors nginx in production)
      '/api': {
        target: 'http://127.0.0.1:3001',
        changeOrigin: false,
      },
    },
  },
  build: {
    outDir: 'dist',
  },
})
