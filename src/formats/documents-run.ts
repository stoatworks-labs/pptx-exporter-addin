// The PDF formats' working parts — pdf-lib and fontkit are most of the pane's
// weight, so this module is loaded only when a PDF format runs (see
// documents.ts for the formats and their options).

import { PDFDocument, rgb, type PDFEmbeddedPage, type PDFImage, type PDFPage } from 'pdf-lib'
import type { Slide } from '../deck/model'
import { EMU_PER_POINT } from '../deck/model'
import { addLink, addOutline, PAGE_SIZES, setInfo, Typesetter } from '../pdf/kit'
import { pdfEnv } from '../pdf/env'
import { encodeNative } from '../render/canvas'
import { mapPdf, pageForSlide, type PdfMapping } from '../render/pdfmap'
import type { Frame } from '../render/sizing'
import { eachSlide, expandName } from './common'
import { HEAD, handoutLayout, MARGIN } from './documents'
import { hotspots, safeUrl } from './overlay'
import { checkCancel, type ExportContext, type Values } from './types'

interface Source {
  doc: PDFDocument
  map: PdfMapping
  hidden: boolean[]
}

async function loadSource(ctx: ExportContext): Promise<Source> {
  const doc = await PDFDocument.load(await ctx.pdf(), { updateMetadata: false, ignoreEncryption: true })
  const hidden = ctx.deck.slides.map((s) => s.hidden)
  return { doc, map: mapPdf(doc.getPageCount(), hidden), hidden }
}

function slidePoints(ctx: ExportContext): [number, number] {
  return [ctx.deck.width / EMU_PER_POINT, ctx.deck.height / EMU_PER_POINT]
}

/** "{deck}" and "{date}" in header and footer text. */
function expandHeader(t: string, ctx: ExportContext): string {
  return t.replace(/\{deck\}/g, ctx.deck.title || ctx.base).replace(/\{date\}/g, new Date().toLocaleDateString())
}

// --- PowerPoint's PDF --------------------------------------------------------

export async function runVector(ctx: ExportContext, v: Values): Promise<void> {
    ctx.progress(0, 'Asking PowerPoint for its PDF')
    const src = await loadSource(ctx)
    const all = ctx.slides.length === ctx.deck.slides.length
    let pages: { slide: Slide; page: number }[] = []
    if (src.map.kind === 'mismatch') {
      if (!all) {
        throw new Error(`PowerPoint's PDF has ${src.map.pages} pages for ${src.map.slides} slides (${src.map.visible} visible), so its pages can't be matched to slides to pick a range. Export every slide, or use "PDF — pictures".`)
      }
      ctx.warn(`PowerPoint's PDF has ${src.map.pages} pages for ${src.map.slides} slides; it has been passed through whole.`)
      if (!v.split) {
        ctx.add(`${ctx.base}.pdf`, await ctx.pdf())
        return
      }
      pages = Array.from({ length: src.map.pages }, (_, i) => ({ slide: ctx.deck.slides[Math.min(i, ctx.deck.slides.length - 1)], page: i }))
    } else {
      const missing: number[] = []
      for (const s of ctx.slides) {
        const p = pageForSlide(src.map, src.hidden, s.index)
        if (p == null) missing.push(s.index + 1)
        else pages.push({ slide: s, page: p })
      }
      if (missing.length) {
        ctx.warn(`PowerPoint's PDF leaves hidden slides out, so slide${missing.length > 1 ? 's' : ''} ${missing.join(', ')} ${missing.length > 1 ? 'are' : 'is'} missing. "PDF — pictures" includes them.`)
      }
    }
    if (!pages.length) throw new Error('None of the chosen slides are in PowerPoint’s PDF.')

    if (v.split) {
      for (let i = 0; i < pages.length; i++) {
        checkCancel(ctx.signal)
        ctx.progress(i / pages.length, `Slide ${pages[i].slide.index + 1}`)
        const out = await PDFDocument.create()
        const [p] = await out.copyPages(src.doc, [pages[i].page])
        out.addPage(p)
        if (v.info) setInfo(out, pages[i].slide.title || ctx.deck.title, ctx.deck.author)
        ctx.add(`${expandName(String(v.name), ctx, pages[i].slide, i)}.pdf`, await out.save())
      }
      return
    }
    const out = await PDFDocument.create()
    const copied = await out.copyPages(src.doc, pages.map((p) => p.page))
    copied.forEach((p) => out.addPage(p))
    if (v.info) setInfo(out, ctx.deck.title, ctx.deck.author)
    if (v.bookmarks) addOutline(out, pages.map((p, i) => ({ title: p.slide.title || `Slide ${p.slide.index + 1}`, page: copied[i] })))
    ctx.add(`${ctx.base}.pdf`, await out.save())
}

// --- Pictures -----------------------------------------------------------------

/** Embeds slide pictures (JPEG or PNG) into a document. */
async function embedSlidePicture(ctx: ExportContext, doc: PDFDocument, slide: Slide, widthPx: number, format: string, quality: number): Promise<PDFImage> {
  const ar = ctx.deck.width / ctx.deck.height
  const w = Math.max(16, Math.round(widthPx))
  const h = Math.max(16, Math.round(w / ar))
  const f: Frame = { width: w, height: h, dx: 0, dy: 0, dw: w, dh: h, renderW: w, renderH: h, dpi: 96 }
  if (format === 'png') return doc.embedPng(await ctx.images.framePng(slide.index, f, '#ffffff'))
  const c = await ctx.images.frame(slide.index, f, '#ffffff')
  return doc.embedJpg(await encodeNative(c, 'image/jpeg', quality / 100))
}

export async function runPictures(ctx: ExportContext, v: Values): Promise<void> {
    const doc = await PDFDocument.create()
    setInfo(doc, ctx.deck.title, ctx.deck.author)
    const [pw, ph] = slidePoints(ctx)
    const pageOf = new Map<number, PDFPage>()
    await eachSlide(ctx, async (s) => {
      const img = await embedSlidePicture(ctx, doc, s, (pw / 72) * Number(v.dpi), String(v.image), Number(v.quality))
      const page = doc.addPage([pw, ph])
      page.drawImage(img, { x: 0, y: 0, width: pw, height: ph })
      pageOf.set(s.index, page)
    })
    if (v.links) {
      for (const s of ctx.slides) {
        const page = pageOf.get(s.index)!
        for (const h of hotspots(ctx.deck, s)) {
          const rect: [number, number, number, number] = [h.x * pw, ph - (h.y + h.h) * ph, (h.x + h.w) * pw, ph - h.y * ph]
          const url = h.link.url ? safeUrl(h.link.url) : null
          const target = h.link.slide != null ? pageOf.get(h.link.slide) : undefined
          if (url) addLink(page, rect, { url })
          else if (target) addLink(page, rect, { page: target })
        }
      }
    }
    if (v.bookmarks) addOutline(doc, ctx.slides.map((s) => ({ title: s.title || `Slide ${s.index + 1}`, page: pageOf.get(s.index)! })))
    ctx.add(`${ctx.base}.pdf`, await doc.save())
}

// --- Handouts and notes pages ----------------------------------------------------

/** Draws a slide into a box: PowerPoint's vector page when it has one, otherwise a picture. */
class SlidePainter {
  private embedded = new Map<number, PDFEmbeddedPage>()
  private warnedRaster = false

  constructor(
    private ctx: ExportContext,
    private doc: PDFDocument,
    private src: Source | null,
    private dpi: number,
    private image: string,
    private quality: number,
  ) {}

  async draw(page: PDFPage, slide: Slide, x: number, y: number, w: number, h: number, frame: boolean): Promise<void> {
    const pageIndex = this.src ? pageForSlide(this.src.map, this.src.hidden, slide.index) : null
    if (this.src && pageIndex != null) {
      let ep = this.embedded.get(pageIndex)
      if (!ep) {
        ep = await this.doc.embedPage(this.src.doc.getPage(pageIndex))
        this.embedded.set(pageIndex, ep)
      }
      page.drawPage(ep, { x, y, width: w, height: h })
    } else {
      try {
        const img = await embedSlidePicture(this.ctx, this.doc, slide, (w / 72) * this.dpi, this.image, this.quality)
        page.drawImage(img, { x, y, width: w, height: h })
        if (this.src && !this.warnedRaster) {
          this.warnedRaster = true
          this.ctx.warn("Slides missing from PowerPoint's PDF (hidden ones) are drawn as pictures.")
        }
      } catch (e) {
        // On the PDF engine a hidden slide has no picture either: leave a
        // labelled gap rather than lose the whole document.
        if (this.ctx.signal.aborted) throw e
        page.drawRectangle({ x, y, width: w, height: h, color: rgb(0.93, 0.93, 0.93) })
        this.ctx.warn(`Slide ${slide.index + 1} couldn't be drawn (${(e as Error).message}); its place is left blank.`)
      }
    }
    if (frame) page.drawRectangle({ x, y, width: w, height: h, borderColor: rgb(0.6, 0.6, 0.6), borderWidth: 0.5 })
  }
}

async function decorate(ts: Typesetter, page: PDFPage, ctx: ExportContext, v: Values, n: number, total: number): Promise<void> {
  const { width, height } = page.getSize()
  const header = expandHeader(String(v.header ?? ''), ctx)
  const footer = expandHeader(String(v.footer ?? ''), ctx)
  const grey = '#555555'
  if (header) await ts.draw(page, header, MARGIN, height - MARGIN + 4, 9, { color: grey })
  if (footer) await ts.draw(page, footer, MARGIN, MARGIN - 14, 9, { color: grey })
  if (v.pageNumbers) {
    const t = `${n} / ${total}`
    await ts.draw(page, t, width - MARGIN - ts.width(t, 9), MARGIN - 14, 9, { color: grey })
  }
}

function pageSize(v: Values, landscape: boolean): [number, number] {
  const [w, h] = PAGE_SIZES[String(v.page)] ?? PAGE_SIZES.a4
  return landscape ? [h, w] : [w, h]
}

export async function runLayout(ctx: ExportContext, v: Values, kind: 'handouts' | 'notes'): Promise<void> {
  const doc = await PDFDocument.create()
  setInfo(doc, ctx.deck.title, ctx.deck.author)
  const ts = await Typesetter.create(doc, await pdfEnv.fonts(), pdfEnv.rasterizer, (m) => ctx.warn(m))
  let src: Source | null = null
  if (v.source !== 'pictures') {
    ctx.progress(0, 'Asking PowerPoint for its PDF')
    src = await loadSource(ctx)
    if (src.map.kind === 'mismatch') {
      ctx.warn(`PowerPoint's PDF has ${src.map.pages} pages for ${src.map.slides} slides, so slides are drawn as pictures instead.`)
      src = null
    }
  }
  const painter = new SlidePainter(ctx, doc, src, Number(v.dpi) || 200, String(v.image ?? 'jpeg'), Number(v.quality) || 85)
  const aspect = ctx.deck.width / ctx.deck.height
  const header = !!expandHeader(String(v.header ?? ''), ctx)
  const footer = !!expandHeader(String(v.footer ?? ''), ctx) || !!v.pageNumbers

  if (kind === 'handouts') {
    const per = Number(v.perPage) || 3
    const [pw, ph] = pageSize(v, v.orientation === 'landscape')
    const layout = handoutLayout(per, pw, ph, aspect, { header, footer, numbers: !!v.numbers })
    const total = Math.ceil(ctx.slides.length / per)
    let page: PDFPage | null = null
    let pageNo = 0
    await eachSlide(ctx, async (s, i) => {
      const slot = i % per
      if (slot === 0) {
        page = doc.addPage([pw, ph])
        pageNo++
        await decorate(ts, page, ctx, v, pageNo, total)
      }
      const b = layout.slides[slot]
      await painter.draw(page!, s, b.x, b.y, b.w, b.h, !!v.frame)
      if (v.numbers) await ts.draw(page!, `${s.index + 1}`, b.x, b.y - 10, 8, { color: '#555555' })
      const l = layout.lines[slot]
      if (l) {
        for (let y = l.y + l.h - 18; y >= l.y; y -= 18) {
          page!.drawLine({ start: { x: l.x, y }, end: { x: l.x + l.w, y }, thickness: 0.4, color: rgb(0.7, 0.7, 0.7) })
        }
      }
    })
    ctx.add(`${ctx.base}-handouts.pdf`, await doc.save())
    return
  }

  // Notes pages.
  const [pw, ph] = pageSize(v, false)
  const size = Number(v.fontSize) || 12
  const leading = size * 1.35
  const top = ph - MARGIN - (header ? HEAD : 0)
  const bottom = MARGIN + (footer ? HEAD : 0)
  const width = pw - 2 * MARGIN
  const slides = v.skipEmpty ? ctx.slides.filter((s) => s.notesText.trim()) : ctx.slides
  if (!slides.length) throw new Error('None of the chosen slides have speaker notes.')
  const pages: PDFPage[] = []
  const sub = { ...ctx, slides }
  await eachSlide(sub, async (s) => {
    let page = doc.addPage([pw, ph])
    pages.push(page)
    // The slide across most of the width, but never more than half the page.
    const sh = Math.min((width * 0.86) / aspect, (top - bottom) * 0.5)
    const sw = sh * aspect
    const sy = top - sh
    await painter.draw(page, s, MARGIN + (width - sw) / 2, sy, sw, sh, !!v.frame)
    await ts.draw(page, `Slide ${s.index + 1}${s.title ? ` — ${s.title}` : ''}`, MARGIN, sy - 24, size, { bold: true })
    let y = sy - 24 - leading * 1.4
    const paras = s.notesText ? s.notesText.split('\n') : []
    for (const para of paras) {
      for (const line of ts.wrap(para, size, width)) {
        if (y < bottom) {
          page = doc.addPage([pw, ph])
          pages.push(page)
          await ts.draw(page, `Slide ${s.index + 1} (continued)`, MARGIN, top - size, size, { bold: true, color: '#555555' })
          y = top - size - leading * 1.4
        }
        await ts.draw(page, line, MARGIN, y, size)
        y -= leading
      }
    }
  })
  for (let i = 0; i < pages.length; i++) await decorate(ts, pages[i], ctx, v, i + 1, pages.length)
  ctx.add(`${ctx.base}-notes.pdf`, await doc.save())
}

