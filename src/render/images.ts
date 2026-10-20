// Slide pictures: asked for once, at the size a format needs, and shared by
// every format in the same export.
//
// Two engines:
//   powerpoint — Slide.getImageAsBase64. PowerPoint itself draws the slide at
//                exactly the pixel size asked for. Needs PowerPointApi 1.8.
//   pdf        — PowerPoint's own PDF of the deck, rasterised by pdf.js. For
//                hosts without 1.8 (Office LTSC on Windows).
//
// PowerPoint won't draw a slide past a certain size: on PowerPoint for Mac
// 16.112 a 16:9 slide 4032 px wide works and 4040 px answers InvalidArgument
// (measured 2026-10-20). Rather than hard-code that, the first refusal sets
// off a search for the largest size it will draw; anything bigger comes from
// PowerPoint's vector PDF through pdf.js at full resolution, or — for a hidden
// slide, which the PDF hasn't got — from PowerPoint's largest picture scaled
// up, with a note saying so.

import type { Deck } from '../deck/model'
import type { Host } from '../host/types'
import { ctx2d, decodeImage, encodeNative, imageSize, makeCanvas, type AnyCanvas } from './canvas'
import { mapPdf, pageForSlide, type PdfMapping } from './pdfmap'
import type { Frame } from './sizing'

export type Engine = 'powerpoint' | 'pdf'

interface PdfDoc {
  numPages: number
  getPage(n: number): Promise<{
    view: number[]
    getViewport(o: { scale: number }): { width: number; height: number }
    render(o: { canvasContext: unknown; viewport: unknown; canvas?: unknown; background?: string }): { promise: Promise<void> }
    cleanup(): void
  }>
}

const CACHE_LIMIT = 384 * 1024 * 1024

export class SlideImages {
  private cache = new Map<string, Promise<Uint8Array>>()
  private sizes = new Map<string, number>()
  private cached = 0
  private pdfDoc?: Promise<{ doc: PdfDoc; map: PdfMapping }>
  /** Narrowest width PowerPoint has refused, and widest it has drawn. */
  private refusedW = Infinity
  private drawnW = 0
  private ceiling?: Promise<number>
  private noted = new Set<string>()

  constructor(
    private host: Host,
    private deck: Deck,
    readonly engine: Engine,
    private pdfBytes: () => Promise<Uint8Array>,
    /** Things the user should hear about (once each), e.g. an upscaled slide. */
    private note: (message: string) => void = () => {},
  ) {}

  /** The slide alone, as a PNG of exactly w × h (or as near as PowerPoint gets). */
  slidePng(index: number, w: number, h: number): Promise<Uint8Array> {
    const key = `${index}@${w}x${h}`
    let p = this.cache.get(key)
    if (p) {
      // Refresh its place in the eviction order.
      this.cache.delete(key)
      this.cache.set(key, p)
      return p
    }
    p = this.engine === 'powerpoint' ? this.powerpointPng(index, w, h) : this.pdfPng(index, w, h)
    this.cache.set(key, p)
    p.then((bytes) => {
      this.sizes.set(key, bytes.length)
      this.cached += bytes.length
      this.evict()
    }, () => this.cache.delete(key))
    return p
  }

  private async powerpointPng(index: number, w: number, h: number): Promise<Uint8Array> {
    if (w < this.refusedW) {
      try {
        const png = await this.host.renderSlide(index, w, h)
        this.drawnW = Math.max(this.drawnW, w)
        return png
      } catch (e) {
        if (!tooBig(e)) throw e
        this.refusedW = Math.min(this.refusedW, w)
      }
    }
    return this.oversize(index, w, h)
  }

  /** Bigger than PowerPoint will draw: its vector PDF at full size, or its
   * largest picture scaled up when the slide isn't in the PDF. */
  private async oversize(index: number, w: number, h: number): Promise<Uint8Array> {
    const s = this.deck.slides[index]
    let viaPdf = false
    try {
      const { map } = await this.loadPdf()
      viaPdf = pageForSlide(map, this.deck.slides.map((x) => x.hidden), index) != null
    } catch {
      viaPdf = false
    }
    if (viaPdf) {
      this.once('pdf', `PowerPoint won't draw slides wider than ${await this.largestWidth(index)} px, so pictures this size come from its vector PDF instead.`)
      return this.pdfPng(index, w, h)
    }
    const maxW = await this.largestWidth(index)
    this.once(`up${index}`, `Slide ${s.index + 1}${s.hidden ? ' is hidden, so it isn’t in PowerPoint’s PDF, and' : ''} PowerPoint won't draw it wider than ${maxW} px; it has been scaled up to ${w} px and will look softer.`)
    const png = await this.host.renderSlide(index, maxW, Math.round((maxW * h) / w))
    const img = await decodeImage(png)
    const c = makeCanvas(w, h)
    ctx2d(c).drawImage(img as CanvasImageSource, 0, 0, w, h)
    if ('close' in img) img.close()
    return encodeNative(c, 'image/png')
  }

  /** The widest PowerPoint will draw this deck's slides, found once by bisection. */
  private largestWidth(index: number): Promise<number> {
    this.ceiling ??= (async () => {
      const ar = this.deck.width / this.deck.height
      let lo = Math.max(this.drawnW, 0)
      let hi = Math.min(this.refusedW, 16384) - 1
      if (!lo) {
        lo = 1280
        try {
          await this.host.renderSlide(index, lo, Math.round(lo / ar))
        } catch {
          return 640
        }
      }
      while (hi - lo > 8) {
        const mid = Math.floor((lo + hi) / 2)
        try {
          await this.host.renderSlide(index, mid, Math.round(mid / ar))
          lo = mid
        } catch (e) {
          if (!tooBig(e)) throw e
          hi = mid - 1
        }
      }
      return lo
    })()
    return this.ceiling
  }

  private once(key: string, message: string): void {
    if (this.noted.has(key)) return
    this.noted.add(key)
    this.note(message)
  }

  private evict(): void {
    for (const key of this.cache.keys()) {
      if (this.cached <= CACHE_LIMIT) break
      const n = this.sizes.get(key)
      if (n == null) continue
      this.cache.delete(key)
      this.sizes.delete(key)
      this.cached -= n
    }
  }

  /** One output frame: the slide drawn into its place, background behind. */
  async frame(index: number, f: Frame, background: string): Promise<AnyCanvas> {
    const png = await this.slidePng(index, f.renderW, f.renderH)
    const img = await decodeImage(png)
    const c = makeCanvas(f.width, f.height)
    const ctx = ctx2d(c)
    ctx.fillStyle = background
    ctx.fillRect(0, 0, c.width, c.height)
    ctx.drawImage(img as CanvasImageSource, f.dx, f.dy, f.dw, f.dh)
    if ('close' in img) img.close()
    return c
  }

  /** The frame as PNG bytes — PowerPoint's own file untouched when the frame
   * is just the slide at the size it drew it. */
  async framePng(index: number, f: Frame, background: string): Promise<Uint8Array> {
    if (f.dx === 0 && f.dy === 0 && f.width === f.dw && f.height === f.dh && f.renderW === f.dw && f.renderH === f.dh) {
      const png = await this.slidePng(index, f.renderW, f.renderH)
      const dims = pngSize(png)
      if (dims && dims.w === f.width && dims.h === f.height) return png
    }
    return encodeNative(await this.frame(index, f, background), 'image/png')
  }

  private loadPdf(): Promise<{ doc: PdfDoc; map: PdfMapping }> {
    this.pdfDoc ??= (async () => {
      const [pdfjs, worker] = await Promise.all([
        import('pdfjs-dist'),
        import('pdfjs-dist/build/pdf.worker.min.mjs?url'),
      ])
      pdfjs.GlobalWorkerOptions.workerSrc = worker.default
      const bytes = await this.pdfBytes()
      const doc = (await pdfjs.getDocument({ data: bytes.slice(), useSystemFonts: true }).promise) as unknown as PdfDoc
      return { doc, map: mapPdf(doc.numPages, this.deck.slides.map((s) => s.hidden)) }
    })()
    return this.pdfDoc
  }

  async pdfMapping(): Promise<PdfMapping> {
    return (await this.loadPdf()).map
  }

  private async pdfPng(index: number, w: number, h: number): Promise<Uint8Array> {
    const { doc, map } = await this.loadPdf()
    const pageIndex = pageForSlide(map, this.deck.slides.map((s) => s.hidden), index)
    if (pageIndex == null) {
      throw new Error(map.kind === 'mismatch'
        ? `PowerPoint's PDF has ${map.pages} pages for ${map.slides} slides, so its pages can't be matched to slides. Use the PowerPoint renderer.`
        : `Slide ${index + 1} is hidden, and PowerPoint's PDF leaves hidden slides out.`)
    }
    const page = await doc.getPage(pageIndex + 1)
    const base = page.getViewport({ scale: 1 })
    const scale = Math.min(w / base.width, h / base.height)
    const vp = page.getViewport({ scale })
    const c = makeCanvas(Math.round(vp.width), Math.round(vp.height))
    const ctx = ctx2d(c)
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, c.width, c.height)
    await page.render({ canvasContext: ctx, viewport: vp, canvas: c }).promise
    page.cleanup()
    return encodeNative(c, 'image/png')
  }
}

/** PowerPoint's answer to a picture bigger than it will draw. */
function tooBig(e: unknown): boolean {
  const err = e as { code?: string; message?: string }
  return err?.code === 'InvalidArgument' || err?.message === 'InvalidArgument'
}

/** Width and height from a PNG's IHDR, without decoding it. */
export function pngSize(b: Uint8Array): { w: number; h: number } | null {
  if (b.length < 24 || b[0] !== 0x89 || b[1] !== 0x50) return null
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength)
  return { w: dv.getUint32(16), h: dv.getUint32(20) }
}

/** Decode any PNG into an RGBA canvas at its own size. */
export async function pngToCanvas(png: Uint8Array): Promise<AnyCanvas> {
  const img = await decodeImage(png)
  const { w, h } = imageSize(img)
  const c = makeCanvas(w, h)
  ctx2d(c).drawImage(img as CanvasImageSource, 0, 0)
  if ('close' in img) img.close()
  return c
}
