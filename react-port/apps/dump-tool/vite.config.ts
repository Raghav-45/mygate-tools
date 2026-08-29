import { crx } from '@crxjs/vite-plugin'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export const manifest = {
  manifest_version: 3,
  name: 'MyGate Dump Tool',
  version: '1.0.0',
  description:
    "Autonomous Multi-Year Ticket Dump Exporter. Bypasses MyGate's 1-year export restriction and automatically merges data into a master Excel file.",
  permissions: ['storage', 'activeTab', 'scripting', 'downloads'],
  host_permissions: ['https://*.mygate.com/*', 'https://*.cloudfront.net/*'],
  action: {
    default_popup: 'index.html',
    default_icon: {
      16: 'icons/icon16.png',
      48: 'icons/icon48.png',
      128: 'icons/icon128.png',
    },
  },
  background: {
    service_worker: 'src/background/index.ts',
    type: 'module',
  },
  icons: {
    16: 'icons/icon16.png',
    48: 'icons/icon48.png',
    128: 'icons/icon128.png',
  },
}

export default defineConfig({
  plugins: [react(), crx({ manifest })],
  build: {
    sourcemap: false,
  },
})
