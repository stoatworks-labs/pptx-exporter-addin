import { defineConfig } from 'vitest/config'
import { resolve } from 'node:path'

// The deck parser uses DOMParser, so tests run under jsdom. Anything that needs
// a real canvas (PNG/JPEG encoding, resampling) is exercised in a browser, not
// here — see README §Testing.
export default defineConfig({
  define: { __APP_VERSION__: JSON.stringify('test') },
  // As in vite.config.ts: the font manager's scanner, vendored.
  resolve: { alias: { '@fm': resolve(import.meta.dirname, 'vendor/pptx-font-manager/src') } },
  test: { environment: 'jsdom', include: ['test/**/*.test.ts'] },
})
