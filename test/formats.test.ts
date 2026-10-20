import { readFileSync, writeFileSync, mkdirSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { describe, expect, it, beforeAll } from 'vitest'
import { PDFDocument, PDFName } from 'pdf-lib'
import { fixture, runFormat, text } from './helpers'
import { parseDeck } from '../src/deck/parse'
import { latex, texEscape, paragraphsTex } from '../src/formats/latex'
import { markdown, outline, speakerNotes, csv, json, csvField, outlineOf } from '../src/formats/text'
import { pdfVector, handouts, notesPages, handoutLayout } from '../src/formats/documents'
import { pdfEnv } from '../src/pdf/env'
import { structure } from '../src/formats/structure'
import { expandName, slug } from '../src/formats/common'
import { frameFor } from '../src/render/sizing'
import { mapPdf, pageForSlide } from '../src/render/pdfmap'

const deck = parseDeck(fixture('sample-16x9.pptx'))
const dump = (name: string, data: Uint8Array | string) => {
  if (!process.env.SAMPLES_DIR) return
  const p = resolve(process.env.SAMPLES_DIR, name)
  mkdirSync(dirname(p), { recursive: true })
  writeFileSync(p, data)
}

beforeAll(() => {
  const font = (f: string) => new Uint8Array(readFileSync(resolve(__dirname, '../src/assets/fonts', f)))
  pdfEnv.fonts = async () => ({ regular: font('NotoSans-Regular.ttf'), bold: font('NotoSans-Bold.ttf') })
  pdfEnv.rasterizer = undefined
})

describe('sizing', () => {
  const W = deck.width, H = deck.height
  it('fits a 16:9 slide to 1080p exactly', () => {
    expect(frameFor(W, H, { size: '1080p' })).toMatchObject({ width: 1920, height: 1080, dx: 0, dy: 0 })
  })
  it("uses the slide's own 96 dpi size for native", () => {
    expect(frameFor(W, H, { size: 'native' })).toMatchObject({ width: 1280, height: 720 })
  })
  it('letterboxes a 4:3 deck into 1080p when asked', () => {
    const f = frameFor(9144000, 6858000, { size: '1080p', fit: 'pad' })
    expect(f).toMatchObject({ width: 1920, height: 1080, dw: 1440, dh: 1080, dx: 240, dy: 0, renderW: 1440 })
  })
  it('keeps the shape of a 4:3 deck by default', () => {
    expect(frameFor(9144000, 6858000, { size: '1080p' })).toMatchObject({ width: 1440, height: 1080 })
  })
  it('crops to cover', () => {
    const f = frameFor(9144000, 6858000, { size: '1080p', fit: 'cover' })
    expect(f).toMatchObject({ width: 1920, height: 1080, dw: 1920, dh: 1440, dy: -180 })
  })
  it('works out print sizes from dpi', () => {
    const f = frameFor(W, H, { size: 'dpi', dpi: 300 })
    expect(f.width).toBe(4000)
    expect(Math.round(f.dpi)).toBe(300)
  })
})

describe('PDF page mapping', () => {
  const hidden = [false, false, true, false]
  it('maps when the PDF has every slide', () => {
    const m = mapPdf(4, hidden)
    expect(m.kind).toBe('all')
    expect(pageForSlide(m, hidden, 3)).toBe(3)
  })
  it('maps around hidden slides', () => {
    const m = mapPdf(3, hidden)
    expect(m.kind).toBe('visible')
    expect(pageForSlide(m, hidden, 2)).toBeNull()
    expect(pageForSlide(m, hidden, 3)).toBe(2)
  })
  it('refuses to guess on a mismatch', () => {
    const m = mapPdf(5, hidden)
    expect(m.kind).toBe('mismatch')
    expect(pageForSlide(m, hidden, 0)).toBeNull()
  })
})

describe('names', () => {
  it('expands patterns', () => {
    const ctx = { base: 'Deck', slides: deck.slides }
    expect(expandName('{deck}-{nn}', ctx, deck.slides[2], 2)).toBe('Deck-03')
    expect(expandName('{nnn}_{title}', ctx, deck.slides[7], 0)).toBe('008_specials-50-5-1')
    expect(slug('Fixture — généré & 日本語')).toBe('fixture-genere-日本語')
  })
})

describe('structure', () => {
  it('finds two columns on the two-content slide', () => {
    const st = structure(deck.slides[2])
    expect(st.columns.length).toBe(2)
  })
  it('marks the title slide', () => {
    expect(structure(deck.slides[0]).isTitleSlide).toBe(true)
    expect(structure(deck.slides[1]).isTitleSlide).toBe(false)
  })
  it('shows video and pictures it can carry', () => {
    expect(structure(deck.slides[8]).columns.flat().map((b) => b.kind)).toEqual(['crop'])
    expect(structure(deck.slides[4]).columns.flat().map((b) => b.kind)).toEqual(['image'])
  })
})

describe('LaTeX', () => {
  it('escapes the specials', () => {
    expect(texEscape('50% & $5 #1 {x} ~ ^ _ \\', 'xelatex')).toBe('50\\% \\& \\$5 \\#1 \\{x\\} \\textasciitilde{} \\textasciicircum{} \\_ \\textbackslash{}')
  })
  it('maps or flags what pdfLaTeX cannot set', () => {
    const bad = new Set<string>()
    expect(texEscape('ok ✓ é 日', 'pdflatex', bad)).toBe('ok \\checkmark{} é ?')
    expect([...bad]).toEqual(['日'])
  })
  it('nests lists by level', () => {
    const body = deck.slides[1].shapes.find((s) => s.placeholder?.type === 'obj')!
    const t = paragraphsTex({ engine: 'xelatex', unsupported: new Set(), v: { links: true }, exported: new Set([0, 1]) }, body.paragraphs!)
    expect(t.match(/\\begin\{itemize\}/g)?.length).toBe(3)
    expect(t.match(/\\end\{itemize\}/g)?.length).toBe(3)
    expect(t).toContain('\\textbf{bold}')
    expect(t).toContain('\\emph{italic}')
    expect(t).toContain('\\href{https://example.com/}{A link to example.com}')
  })
  it('builds a structured Beamer deck', async () => {
    const { files, warnings } = await runFormat(latex, { mode: 'structured', crops: false })
    const tex = text(files.get('main.tex'))
    dump('latex/main.tex', tex)
    for (const [p, d] of files) dump(`latex/${p}`, d)
    expect(tex).toContain('\\documentclass[aspectratio=169]{beamer}')
    expect(tex).toContain('\\titlepage')
    expect(tex).toContain('\\begin{columns}[T]')
    expect(tex).toContain('\\multicolumn{2}{l}{Merged cell}')
    expect(tex).toContain('\\hyperlink{slide-2}{Back to the agenda}')
    expect(tex).toContain('\\section{Intro}')
    expect(tex).toContain('\\usepackage{xeCJK}')
    expect(tex).toContain('\\note{Welcome notes')
    expect(tex).toContain('Specials: 50\\% \\& \\$5 \\#1')
    expect(files.has('media/image1.png')).toBe(true)
    expect(files.get('latexmkrc')).toContain('$pdf_mode = 5;')
    expect((tex.match(/\\begin\{frame\}/g) ?? []).length).toBe(9)
    expect(warnings).toEqual([])
  })
})

describe('Markdown', () => {
  it('writes GFM with nested lists, links, a table and notes', async () => {
    const { files } = await runFormat(markdown, { flavour: 'gfm', content: 'text', crops: false })
    const md = text(files.get('sample.md'))
    dump('sample.md', md)
    expect(md).toContain('## 2. Agenda')
    expect(md).toContain('- First point\n  - **bold** and *italic*\n    - Third level\n- [A link to example.com](https://example.com/)')
    expect(md).toContain('| Kit | Qty | Notes |')
    expect(md).toContain('> Agenda notes.\n>\n> A paragraph after a blank line.')
    expect(md).toContain('[Back to the agenda](#slide-2)')
    expect(md).toContain('Specials: 50% & $5 #1')
  })
  it('writes Marp with notes as comments', async () => {
    const { files } = await runFormat(markdown, { flavour: 'marp', content: 'text', crops: false })
    const md = text(files.get('sample.md'))
    expect(md.startsWith('---\nmarp: true')).toBe(true)
    expect(md).toContain('<!--\nWelcome notes')
    expect(md.split('\n---\n').length).toBeGreaterThan(8)
  })
})

describe('Outline, notes, CSV, JSON', () => {
  it('outlines slides with levels', () => {
    expect(outlineOf(deck.slides[1]).map((i) => i.level)).toEqual([0, 1, 2, 3, 1])
  })
  it('writes RTF with outline levels and escaped Unicode', async () => {
    const { files } = await runFormat(outline, { as: 'rtf' })
    const rtf = text(files.get('sample-outline.rtf'))
    dump('sample-outline.rtf', rtf)
    expect(rtf.startsWith('{\\rtf1')).toBe(true)
    expect(rtf).toContain('\\outlinelevel0')
    expect(rtf).toContain('\\u26085?') // 日
    expect(rtf.split('{').length).toBe(rtf.split('}').length)
  })
  it('writes OPML that parses', async () => {
    const { files } = await runFormat(outline, { as: 'opml' })
    const doc = new DOMParser().parseFromString(text(files.get('sample-outline.opml')), 'application/xml')
    expect(doc.getElementsByTagName('parsererror').length).toBe(0)
    expect(doc.querySelectorAll('body > outline').length).toBe(9)
  })
  it('writes speaker notes', async () => {
    const { files } = await runFormat(speakerNotes, { skipEmpty: true })
    const t = text(files.get('sample-notes.txt'))
    expect(t).toContain('6. Hidden Slide (hidden)')
    expect(t).not.toContain('Two Columns')
  })
  it('quotes CSV properly', async () => {
    expect(csvField('a,b', ',')).toBe('"a,b"')
    expect(csvField('say "hi"', ',')).toBe('"say ""hi"""')
    const { files } = await runFormat(csv, {})
    const c = text(files.get('sample.csv'))
    expect(c.startsWith('\ufeffSlide,Title')).toBe(true)
    expect(c).toContain('7,Fade and Advance,Finale,,')
    expect(c).toContain(',fade,3,')
  })
  it('writes JSON that round-trips', async () => {
    const { files } = await runFormat(json, {}, { slides: [0, 1] })
    const j = JSON.parse(text(files.get('sample.json')))
    expect(j.slides.length).toBe(2)
    expect(j.slides[1].shapes.some((s: { text?: unknown }) => s.text)).toBe(true)
  })
})

describe('PDF', () => {
  it('lays handouts out inside the page, slide shape kept', () => {
    for (const n of [1, 2, 3, 4, 6, 9]) {
      for (const [w, h] of [[595.28, 841.89], [841.89, 595.28]]) {
        const l = handoutLayout(n, w, h, 16 / 9, { header: true, footer: true, numbers: true })
        expect(l.slides.length).toBe(n)
        for (const b of l.slides) {
          expect(b.x).toBeGreaterThanOrEqual(39)
          expect(b.y).toBeGreaterThanOrEqual(39)
          expect(b.x + b.w).toBeLessThanOrEqual(w - 39)
          expect(b.y + b.h).toBeLessThanOrEqual(h - 39)
          expect(b.w / b.h).toBeCloseTo(16 / 9, 1)
        }
        if (n === 3) expect(l.lines.length).toBe(3)
      }
    }
  })
  it("trims PowerPoint's PDF to a range around the hidden slide", async () => {
    const { files, warnings } = await runFormat(pdfVector, {}, { slides: [4, 5, 6] })
    const doc = await PDFDocument.load(files.get('sample.pdf') as Uint8Array)
    expect(doc.getPageCount()).toBe(2)
    expect(warnings[0]).toMatch(/hidden slides out/)
    expect(doc.catalog.get(PDFName.of('Outlines'))).toBeDefined()
    dump('pdf/sample-range.pdf', files.get('sample.pdf') as Uint8Array)
  })
  it('splits into one PDF per slide', async () => {
    const { files } = await runFormat(pdfVector, { split: true }, { slides: [0, 1] })
    expect([...files.keys()]).toEqual(['sample-01.pdf', 'sample-02.pdf'])
  })
  it('builds vector handouts', async () => {
    const { files } = await runFormat(handouts, { perPage: '6' }, { slides: [0, 1, 2, 3, 4, 6, 7, 8] })
    const pdf = files.get('sample-handouts.pdf') as Uint8Array
    dump('pdf/sample-handouts.pdf', pdf)
    expect((await PDFDocument.load(pdf)).getPageCount()).toBe(2)
  })
  it('builds notes pages, overflowing long notes', async () => {
    const { files, warnings } = await runFormat(notesPages, { fontSize: 12 }, { slides: [0, 1, 3, 7] })
    const pdf = files.get('sample-notes.pdf') as Uint8Array
    dump('pdf/sample-notes.pdf', pdf)
    expect((await PDFDocument.load(pdf)).getPageCount()).toBe(4)
    // ✓ isn't in Noto Sans and there's no canvas here to draw it.
    expect(warnings.join(' ')).toMatch(/replaced/)
  })
})
