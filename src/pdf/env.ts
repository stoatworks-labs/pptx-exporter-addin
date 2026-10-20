// Where the PDF formats get their fonts and their fallback text drawing —
// swappable for tests, which have neither fetch-able assets nor a canvas.

import type { FontBytes, LineRasterizer } from './kit'
import { canvasRasterizer } from './raster-text'

export const pdfEnv: { fonts: () => Promise<FontBytes>; rasterizer?: LineRasterizer } = {
  fonts: async () => (await import('./kit')).loadBundledFonts(),
  rasterizer: canvasRasterizer,
}
