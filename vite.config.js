import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5174,
    proxy: {
      '/anthropic': {
        target: 'https://api.anthropic.com',
        changeOrigin: true,
        rewrite: path => path.replace(/^\/anthropic/, ''),
        headers: { 'anthropic-version': '2023-06-01' },
      },
      '/devops': {
        target: 'https://dev.azure.com',
        changeOrigin: true,
        rewrite: path => path.replace(/^\/devops/, ''),
      },
    },
  },
})
