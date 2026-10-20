import { defineConfig } from 'vitest/config'

// The deck parser uses DOMParser, so tests run under jsdom. Anything that needs
// a real canvas (PNG/JPEG encoding, resampling) is exercised in a browser, not
// here — see README §Testing.
export default defineConfig({
  define: { __APP_VERSION__: JSON.stringify('test') },
  test: { environment: 'jsdom', include: ['test/**/*.test.ts'] },
})
