// The web view's own text engine, for lines the bundled PDF font can't set.

import { ctx2d, encodeNative, makeCanvas, type Ctx2D } from '../render/canvas'
import type { LineRasterizer } from './kit'

const SCALE = 4 // pixels per point: about 288 dpi, sharp in print
const FAMILY = '-apple-system, "Segoe UI", "Hiragino Sans", "Yu Gothic", "Noto Sans CJK JP", "Noto Sans", "Apple Color Emoji", "Segoe UI Emoji", sans-serif'

let measurer: Ctx2D | undefined

function font(sizePt: number, bold: boolean, scale = 1): string {
  return `${bold ? '700 ' : ''}${sizePt * scale}px ${FAMILY}`
}

export const canvasRasterizer: LineRasterizer = {
  measure(text, sizePt, bold) {
    measurer ??= ctx2d(makeCanvas(1, 1))
    measurer.font = font(sizePt, bold)
    return measurer.measureText(text).width
  },
  async render(text, sizePt, bold, color) {
    const width = this.measure(text, sizePt, bold)
    const heightPt = sizePt * 1.35
    const baselinePt = sizePt * 1.05
    const c = makeCanvas(Math.ceil(width * SCALE) + 2, Math.ceil(heightPt * SCALE))
    const g = ctx2d(c)
    g.font = font(sizePt, bold, SCALE)
    g.fillStyle = color
    g.textBaseline = 'alphabetic'
    g.fillText(text, 1, baselinePt * SCALE)
    return { png: await encodeNative(c, 'image/png'), widthPt: c.width / SCALE, heightPt, baselinePt }
  },
}
