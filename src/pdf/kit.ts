// pdf-lib helpers: Unicode text that never comes out as boxes, bookmarks, and
// link annotations (pdf-lib has no high-level API for the last two).
//
// Text goes through Noto Sans (bundled, OFL; Latin, Greek and Cyrillic) when
// the font has every glyph a line needs. A line it can't cover — Japanese notes, emoji, ✓ — is drawn by
// the web view's own text engine into a small high-resolution picture
// instead, so it still reads correctly; it just can't be selected.

import fontkit from '@pdf-lib/fontkit'
import { PDFArray, PDFDocument, PDFHexString, PDFName, PDFString, rgb, type PDFFont, type PDFPage } from 'pdf-lib'

export interface FontBytes {
  regular: Uint8Array
  bold: Uint8Array
}

/** Measures and draws a line with the web view's fonts, for text Noto can't cover. */
export interface LineRasterizer {
  measure(text: string, sizePt: number, bold: boolean): number
  render(text: string, sizePt: number, bold: boolean, color: string): Promise<{ png: Uint8Array; widthPt: number; heightPt: number; baselinePt: number }>
}

export async function loadBundledFonts(): Promise<FontBytes> {
  const [r, b] = await Promise.all([
    import('../assets/fonts/NotoSans-Regular.ttf?url'),
    import('../assets/fonts/NotoSans-Bold.ttf?url'),
  ])
  const get = async (u: string) => new Uint8Array(await (await fetch(u)).arrayBuffer())
  const [regular, bold] = await Promise.all([get(r.default), get(b.default)])
  return { regular, bold }
}

export class Typesetter {
  private warnedMissing = false

  private constructor(
    private regular: PDFFont,
    private bold: PDFFont,
    private covers: (cp: number) => boolean,
    private raster: LineRasterizer | undefined,
    private onMissing: (msg: string) => void,
  ) {}

  static async create(doc: PDFDocument, fonts: FontBytes, raster?: LineRasterizer, onMissing: (msg: string) => void = () => {}): Promise<Typesetter> {
    doc.registerFontkit(fontkit)
    // Embedded whole: pdf-lib's own subsetting blanks most of Noto Sans's
    // glyphs, so the font is pre-subset at build time instead
    // (scripts/make-pdf-fonts.sh) — about 90 KB a PDF.
    const [regular, bold] = await Promise.all([
      doc.embedFont(fonts.regular, { subset: false }),
      doc.embedFont(fonts.bold, { subset: false }),
    ])
    const fk = fontkit.create(fonts.regular)
    const cache = new Map<number, boolean>()
    const covers = (cp: number) => {
      let v = cache.get(cp)
      if (v == null) cache.set(cp, (v = fk.hasGlyphForCodePoint(cp)))
      return v
    }
    return new Typesetter(regular, bold, covers, raster, onMissing)
  }

  covered(text: string): boolean {
    for (const ch of text) {
      const cp = ch.codePointAt(0)!
      if (cp < 0x20) continue
      if (!this.covers(cp)) return false
    }
    return true
  }

  private font(bold: boolean): PDFFont {
    return bold ? this.bold : this.regular
  }

  width(text: string, size: number, bold = false): number {
    if (this.covered(text) || !this.raster) return this.font(bold).widthOfTextAtSize(this.coverable(text), size)
    return this.raster.measure(text, size, bold)
  }

  /** Without a rasteriser (tests), replace what the font lacks with '?'. */
  private coverable(text: string): string {
    if (this.covered(text)) return text
    if (!this.warnedMissing) {
      this.warnedMissing = true
      this.onMissing('Some characters in the notes are outside the PDF font and were replaced.')
    }
    let out = ''
    for (const ch of text) out += this.covers(ch.codePointAt(0)!) ? ch : '?'
    return out
  }

  async draw(page: PDFPage, text: string, x: number, baseline: number, size: number, o: { bold?: boolean; color?: string } = {}): Promise<void> {
    if (!text) return
    const color = o.color ?? '#000000'
    if (this.covered(text) || !this.raster) {
      const [r, g, b] = hexRgb(color)
      page.drawText(this.coverable(text), { x, y: baseline, size, font: this.font(!!o.bold), color: rgb(r, g, b) })
      return
    }
    const img = await this.raster.render(text, size, !!o.bold, color)
    const embedded = await page.doc.embedPng(img.png)
    page.drawImage(embedded, { x, y: baseline - (img.heightPt - img.baselinePt), width: img.widthPt, height: img.heightPt })
  }

  /** Word-wrap one paragraph to a width. Text without spaces (CJK) breaks anywhere. */
  wrap(text: string, size: number, maxWidth: number, bold = false): string[] {
    if (!text) return ['']
    const lines: string[] = []
    const tokens = text.match(/\S+\s*|\s+/g) ?? [text]
    let line = ''
    const push = () => {
      lines.push(line.trimEnd())
      line = ''
    }
    for (const tok of tokens) {
      const trial = line + tok
      if (this.width(trial.trimEnd(), size, bold) <= maxWidth) {
        line = trial
        continue
      }
      if (line) push()
      if (this.width(tok.trimEnd(), size, bold) <= maxWidth) {
        line = tok
        continue
      }
      // A single token wider than the line: break it by character.
      for (const ch of tok) {
        if (line && this.width(line + ch, size, bold) > maxWidth) push()
        line += ch
      }
    }
    if (line.trim() || !lines.length) push()
    return lines
  }
}

export function hexRgb(color: string): [number, number, number] {
  let h = color.replace('#', '')
  if (h.length === 3) h = h.split('').map((c) => c + c).join('')
  const n = parseInt(h, 16)
  return [((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255]
}

/** Bookmarks, one per entry, in order. */
export function addOutline(doc: PDFDocument, entries: { title: string; page: PDFPage }[]): void {
  if (!entries.length) return
  const ctx = doc.context
  const outlinesRef = ctx.nextRef()
  const refs = entries.map(() => ctx.nextRef())
  entries.forEach((e, i) => {
    const item = ctx.obj({
      Title: PDFHexString.fromText(e.title),
      Parent: outlinesRef,
      Dest: [e.page.ref, 'Fit'],
    })
    if (i > 0) item.set(PDFName.of('Prev'), refs[i - 1])
    if (i < refs.length - 1) item.set(PDFName.of('Next'), refs[i + 1])
    ctx.assign(refs[i], item)
  })
  ctx.assign(outlinesRef, ctx.obj({ Type: 'Outlines', First: refs[0], Last: refs[refs.length - 1], Count: refs.length }))
  doc.catalog.set(PDFName.of('Outlines'), outlinesRef)
}

/** A clickable rectangle (PDF coordinates, origin bottom-left). */
export function addLink(page: PDFPage, rect: [number, number, number, number], target: { url: string } | { page: PDFPage }): void {
  const ctx = page.doc.context
  const annot = ctx.obj({
    Type: 'Annot',
    Subtype: 'Link',
    Rect: rect,
    Border: [0, 0, 0],
  })
  if ('url' in target) annot.set(PDFName.of('A'), ctx.obj({ Type: 'Action', S: 'URI', URI: PDFString.of(target.url) }))
  else annot.set(PDFName.of('Dest'), ctx.obj([target.page.ref, 'Fit']))
  const ref = ctx.register(annot)
  let annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) {
    annots = ctx.obj([])
    page.node.set(PDFName.of('Annots'), annots)
  }
  annots.push(ref)
}

export function setInfo(doc: PDFDocument, title: string, author?: string): void {
  doc.setTitle(title, { showInWindowTitleBar: true })
  if (author) doc.setAuthor(author)
  doc.setCreator('Advanced Exporter for PowerPoint')
  doc.setProducer('Advanced Exporter (pdf-lib)')
}

export const PAGE_SIZES: Record<string, [number, number]> = {
  a4: [595.28, 841.89],
  letter: [612, 792],
  legal: [612, 1008],
  a3: [841.89, 1190.55],
  tabloid: [792, 1224],
}
