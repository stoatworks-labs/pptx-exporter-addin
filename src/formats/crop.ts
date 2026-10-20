// Pictures of part of a slide — for the things a text format can't rebuild
// (charts, SmartArt, drawn shapes, EMF pictures): PowerPoint draws the slide,
// and the shape's box is cut out of it.

import type { Rect, Shape, Slide } from '../deck/model'
import { ctx2d, encodeNative, makeCanvas } from '../render/canvas'
import type { Frame } from '../render/sizing'
import type { ExportContext } from './types'

/** PNG of a region of a slide; `slideWidthPx` is the width of the whole slide render. */
export async function cropRegion(ctx: ExportContext, slide: Slide, r: Rect, slideWidthPx = 1920): Promise<Uint8Array> {
  const w = Math.round(slideWidthPx)
  const h = Math.round(w * (ctx.deck.height / ctx.deck.width))
  const f: Frame = { width: w, height: h, dx: 0, dy: 0, dw: w, dh: h, renderW: w, renderH: h, dpi: 96 }
  const full = await ctx.images.frame(slide.index, f, '#ffffff')
  const sx = Math.max(0, Math.floor((r.x / ctx.deck.width) * w))
  const sy = Math.max(0, Math.floor((r.y / ctx.deck.height) * h))
  const sw = Math.max(1, Math.min(w - sx, Math.ceil((r.w / ctx.deck.width) * w)))
  const sh = Math.max(1, Math.min(h - sy, Math.ceil((r.h / ctx.deck.height) * h)))
  const c = makeCanvas(sw, sh)
  ctx2d(c).drawImage(full as CanvasImageSource, sx, sy, sw, sh, 0, 0, sw, sh)
  return encodeNative(c, 'image/png')
}

/** Picture formats a document format can carry as they are. */
export function passThroughImage(contentType: string): 'png' | 'jpg' | 'gif' | null {
  if (contentType === 'image/png') return 'png'
  if (contentType === 'image/jpeg') return 'jpg'
  if (contentType === 'image/gif') return 'gif'
  return null
}

/** Shapes a structured export has to show as a picture rather than rebuild. */
export function needsPicture(s: Shape): boolean {
  if (s.kind === 'chart' || s.kind === 'diagram' || s.kind === 'other' || s.kind === 'media') return true
  if (s.kind === 'picture') return !s.image || !passThroughImage(s.image.contentType)
  return false
}
