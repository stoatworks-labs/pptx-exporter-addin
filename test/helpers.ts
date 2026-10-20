// A stand-in export context for running formats under vitest (no PowerPoint,
// no canvas). Formats that need slide pictures can't run here; everything
// else can, against the real fixture decks.
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { parseDeck } from '../src/deck/parse'
import type { Deck } from '../src/deck/model'
import type { ExportContext, FormatDef, Values } from '../src/formats/types'
import { normalise } from '../src/formats/types'

export const fixture = (name: string) => new Uint8Array(readFileSync(resolve(__dirname, 'fixtures', name)))

/** A PDF shaped like PowerPoint's: one page per visible slide, with the title on it. */
export async function fakePowerPointPdf(deck: Deck): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const w = deck.width / 12700, h = deck.height / 12700
  for (const s of deck.slides) {
    if (s.hidden) continue
    const p = doc.addPage([w, h])
    p.drawText(`Slide ${s.index + 1}`, { x: 40, y: h - 60, size: 28, font })
  }
  return doc.save()
}

export interface Run {
  files: Map<string, Uint8Array | string>
  warnings: string[]
}

export async function runFormat(def: FormatDef, values: Partial<Values>, opts: { deck?: string; slides?: number[] } = {}): Promise<Run> {
  const bytes = fixture(opts.deck ?? 'sample-16x9.pptx')
  const deck = parseDeck(bytes)
  const files = new Map<string, Uint8Array | string>()
  const warnings: string[] = []
  const pdf = await fakePowerPointPdf(deck)
  const ctx: ExportContext = {
    deck,
    slides: opts.slides ? opts.slides.map((i) => deck.slides[i]) : deck.slides,
    base: 'sample',
    host: {} as ExportContext['host'],
    images: new Proxy({}, { get: () => () => { throw new Error('no slide pictures under vitest') } }) as ExportContext['images'],
    add: (p, d) => void files.set(p, d),
    pdf: async () => pdf,
    pptx: bytes,
    progress: () => {},
    warn: (m) => void warnings.push(m),
    signal: new AbortController().signal,
  }
  await def.run(ctx, normalise(def, values as Values))
  return { files, warnings }
}

export const text = (d: Uint8Array | string | undefined) => (typeof d === 'string' ? d : new TextDecoder().decode(d))
