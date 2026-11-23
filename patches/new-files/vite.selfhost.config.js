import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath, URL } from 'node:url'

import tailwindcss from '@tailwindcss/vite'
import vue from '@vitejs/plugin-vue'
import { defineConfig } from 'vite'

const certDir = path.join(os.homedir(), '.office-addin-dev-certs')
const https = {
  key: fs.readFileSync(path.join(certDir, 'localhost.key')),
  cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
}

export default defineConfig({
  plugins: [vue(), tailwindcss()],
  resolve: {
    alias: [
      { find: '@', replacement: fileURLToPath(new URL('./src', import.meta.url)) },
      { find: /^(?:node:)?async_hooks$/, replacement: fileURLToPath(new URL('./async_hook.js', import.meta.url)) },
    ],
  },
  server: { port: 3000, strictPort: true, https },
  preview: { port: 3000, strictPort: true, https },
  build: { target: 'es2022' },
})
