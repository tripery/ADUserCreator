import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    host: '0.0.0.0',
    port: 5173,
    strictPort: true,

    watch: {
      usePolling: true,
      interval: 1000
    },

    hmr: {
      protocol: 'ws',
      host: 'localhost',
      port: 5173,
      clientPort: 5173
    },

    proxy: {
      '/api': {
        target: 'http://host.docker.internal:8787',
        changeOrigin: true,
        secure: false
      }
    }
  }
})
