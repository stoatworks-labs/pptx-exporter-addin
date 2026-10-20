// A stand-in for PowerPoint, for running the pane in an ordinary browser.
//
// It reads a .pptx (from ?mock=<url> or a file the user picks) and draws each
// slide itself: the words where the shapes are, the pictures, a box for
// anything else. It is nothing like PowerPoint's rendering and isn't meant to
// be — it exists so every format's plumbing can be driven and tested without
// PowerPoint. Its "PowerPoint PDF" leaves hidden slides out, as the real one
// is expected to.

import type { Deck, Shape } from '../deck/model'
import { EMU_PER_POINT } from '../deck/model'
import { parseDeck } from '../deck/parse'
import { ctx2d, decodeImage, encodeNative, makeCanvas, type Ctx2D } from '../render/canvas'
import type { FileProgress, Host, HostCapabilities } from './types'

export class MockHost implements Host {
  readonly kind = 'mock' as const
  readonly label = 'Browser preview (no PowerPoint)'
  readonly caps: HostCapabilities = { slideImages: true, pdf: true, exportSlide: false, selection: true, apiVersion: 'mock' }
  readonly deck: Deck
  private pdfBytes?: Promise<Uint8Array>
  selection: number[] = [0]

  constructor(private bytes: Uint8Array, private name: string) {
    this.deck = parseDeck(bytes, name)
  }

  static async fromUrl(url: string): Promise<MockHost> {
    const res = await fetch(url)
    if (!res.ok) throw new Error(`Could not fetch ${url}: ${res.status}`)
    const name = decodeURIComponent(url.split('/').pop() ?? 'deck').replace(/\.pptx$/i, '')
    return new MockHost(new Uint8Array(await res.arrayBuffer()), name)
  }

  documentName(): string {
    return this.name
  }

  async pptx(onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
    onProgress?.({ done: 1, total: 1 })
    return this.bytes
  }

  pdf(onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
    this.pdfBytes ??= (async () => {
      const { PDFDocument } = await import('pdf-lib')
      const doc = await PDFDocument.create()
      const w = this.deck.width / EMU_PER_POINT
      const h = this.deck.height / EMU_PER_POINT
      for (const s of this.deck.slides) {
        if (s.hidden) continue
        const png = await this.renderSlide(s.index, Math.round(w * 2), Math.round(h * 2))
        const img = await doc.embedPng(png)
        doc.addPage([w, h]).drawImage(img, { x: 0, y: 0, width: w, height: h })
      }
      onProgress?.({ done: 1, total: 1 })
      return doc.save()
    })()
    return this.pdfBytes
  }

  async slideCount(): Promise<number> {
    return this.deck.slides.length
  }

  async selectedSlides(): Promise<number[]> {
    return this.selection
  }

  async renderSlide(index: number, width: number, height: number): Promise<Uint8Array> {
    const s = this.deck.slides[index]
    const scale = Math.min(width / this.deck.width, height / this.deck.height)
    const c = makeCanvas(Math.round(this.deck.width * scale), Math.round(this.deck.height * scale))
    const g = ctx2d(c)
    g.fillStyle = '#ffffff'
    g.fillRect(0, 0, c.width, c.height)
    for (const sh of s.shapes) await this.drawShape(g, sh, scale)
    g.fillStyle = 'rgba(0,0,0,0.35)'
    g.font = `${Math.max(10, c.height * 0.025)}px sans-serif`
    g.textAlign = 'right'
    g.fillText(`${index + 1}${s.hidden ? ' (hidden)' : ''} · mock render`, c.width - c.height * 0.02, c.height * 0.97)
    return encodeNative(c, 'image/png')
  }

  async exportSlide(): Promise<Uint8Array> {
    throw new Error('The browser preview cannot split a deck; that needs PowerPoint.')
  }

  private async drawShape(g: Ctx2D, sh: Shape, k: number): Promise<void> {
    if (!sh.rect) return
    const x = sh.rect.x * k, y = sh.rect.y * k, w = sh.rect.w * k, h = sh.rect.h * k
    if (sh.kind === 'picture' && sh.image) {
      const bytes = this.deck.readPart(sh.image.part)
      if (bytes && /png|jpe?g|gif/.test(sh.image.contentType)) {
        try {
          const img = await decodeImage(bytes, sh.image.contentType)
          g.drawImage(img as CanvasImageSource, x, y, w, h)
          return
        } catch {
          /* draw the placeholder box */
        }
      }
    }
    if (sh.kind === 'media' || sh.kind === 'chart' || sh.kind === 'diagram' || sh.kind === 'picture' || sh.kind === 'other') {
      g.fillStyle = sh.kind === 'media' ? '#111' : '#dde3ea'
      g.fillRect(x, y, w, h)
      g.fillStyle = sh.kind === 'media' ? '#fff' : '#456'
      g.font = `${Math.max(10, h * 0.12)}px sans-serif`
      g.textAlign = 'center'
      g.fillText(sh.kind === 'media' ? '▶' : sh.kind, x + w / 2, y + h / 2)
      g.textAlign = 'left'
      return
    }
    if (sh.kind === 'table' && sh.table) {
      const rows = sh.table.rows.length
      const cols = sh.table.colWidths.length || 1
      g.strokeStyle = '#678'
      g.lineWidth = Math.max(1, k * 12700)
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cx = x + (w / cols) * c, cy = y + (h / rows) * r
          g.strokeRect(cx, cy, w / cols, h / rows)
          const cell = sh.table.rows[r][c]
          if (cell && !cell.merged) {
            g.fillStyle = '#222'
            g.font = `${Math.max(8, (h / rows) * 0.35)}px sans-serif`
            g.fillText(cell.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join(' '), cx + 6 * k * 12700, cy + (h / rows) * 0.62, w / cols - 12 * k * 12700)
          }
        }
      }
      return
    }
    if (sh.paragraphs) this.drawText(g, sh, x, y, w, k)
  }

  private drawText(g: Ctx2D, sh: Shape, x: number, y: number, w: number, k: number): void {
    const isTitle = sh.placeholder?.type === 'title' || sh.placeholder?.type === 'ctrTitle'
    let cy = y
    for (const p of sh.paragraphs ?? []) {
      const pt = p.runs.find((r) => r.size)?.size ?? (isTitle ? 40 : 24 - p.level * 3)
      const px = pt * EMU_PER_POINT * k
      cy += px * 1.2
      let cx = x + p.level * px * 1.2
      if (p.bullet !== 'none') {
        g.fillStyle = '#333'
        g.font = `${px}px sans-serif`
        g.fillText('•', cx, cy)
        cx += px * 0.9
      }
      for (const r of p.runs) {
        if (r.br) continue
        g.font = `${r.italic ? 'italic ' : ''}${r.bold || isTitle ? '700 ' : ''}${px}px sans-serif`
        g.fillStyle = r.link ? '#1a5fb4' : r.color ?? '#222'
        g.fillText(r.text, cx, cy, Math.max(1, x + w - cx))
        cx += g.measureText(r.text).width
      }
    }
  }
}
