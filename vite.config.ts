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
    // Deliberately above the ELK layout engine's lazy chunk (~1.4 MB) so the
    // known, code-split heavy modules do not spam warnings, while still
    // flagging any unexpected growth in the eagerly-loaded chunks.
    chunkSizeWarningLimit: 1500,
    rollupOptions: {
      output: {
        // Split stable vendor code into long-cached chunks so a deploy that
        // only touches app code doesn't bust the React/ReactFlow caches.
        // React + react-dom + scheduler MUST stay together to avoid
        // "Cannot access before initialization" from a bad init order.
        manualChunks(id) {
          if (!id.includes('node_modules')) {
            return undefined;
          }
          if (/[\\/]node_modules[\\/](react|react-dom|scheduler|use-sync-external-store)[\\/]/.test(id)) {
            return 'react-vendor';
          }
          if (/[\\/]node_modules[\\/]@?reactflow[\\/]/.test(id)) {
            return 'reactflow-vendor';
          }
          // Heavy libraries reached ONLY through dynamic import() must keep
          // their own lazy chunks. Folding them into the eager vendor chunk
          // would block first paint and undo the lazy auth/telemetry split
          // (msal + Application Insights) and the code-split exporters/layout.
          if (/[\\/]node_modules[\\/](?:@azure[\\/]msal-[^\\/]+|@microsoft[\\/]applicationinsights-[^\\/]+|@microsoft[\\/]dynamicproto-js|@nevware21[\\/][^\\/]+|pptxgenjs|jszip|html-to-image|elkjs|dagre|graphlib|lodash|image-size)[\\/]/.test(id)) {
            return undefined;
          }
          // Remaining eagerly-loaded dependencies (e.g. lucide-react) form a
          // stable, long-cached vendor chunk.
          return 'vendor';
        },
      },
    },
  },
})
