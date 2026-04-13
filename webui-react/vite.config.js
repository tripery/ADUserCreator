import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

const repositoryName = process.env.GITHUB_REPOSITORY?.split('/')[1]
const defaultBase =
  process.env.VITE_BASE_PATH ||
  (process.env.GITHUB_ACTIONS && repositoryName ? `/${repositoryName}/` : '/')

export default defineConfig({
  base: defaultBase,
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
