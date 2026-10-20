import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { encode as encodePng } from 'fast-png'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import type { Deck, Paragraph, Shape, Slide } from '../src/deck/model'
import { parseDeck, resolvePart } from '../src/deck/parse'
import { pptxCopy, splitSlides } from '../src/formats/office'
import { rewritePptx } from '../src/formats/office-copy'
import { buildNotesDocx, type HandoutLayout } from '../src/formats/office-docx'
import { buildFlatPptx, transitionXml, type PictureSlide } from '../src/formats/office-flat'
import { buildOdp, ODP_MIME } from '../src/formats/office-odp'
import { crc32, CT, readZip } from '../src/formats/ooxml-write'
import { defaults, type ExportContext, type Values } from '../src/formats/types'
import { all, first, kids, NS, parseXml } from '../src/lib/xml'

const fixture = () => new Uint8Array(readFileSync(resolve(__dirname, 'fixtures', 'sample-16x9.pptx')))
const SAMPLES = process.env.SAMPLES_DIR

function sample(name: string, bytes: Uint8Array): void {
  if (!SAMPLES) return
  mkdirSync(SAMPLES, { recursive: true })
  writeFileSync(join(SAMPLES, name), bytes)
}

type Files = Record<string, Uint8Array>
const unzip = (b: Uint8Array): Files => unzipSync(b)
const xml = (f: Files, name: string) => parseXml(f[name])
const text = (f: Files, name: string) => strFromU8(f[name])

/** Everything PowerPoint and Word check on open that a test can: every
 * Override and relationship target exists, every part has a content type,
 * every XML part parses, and every r:id / r:embed in a part is one of its
 * relationships. */
function packageProblems(files: Files): string[] {
  const out: string[] = []
  const lower = new Map(Object.keys(files).map((n) => [n.toLowerCase(), n]))
  const real = (p: string) => lower.get(decodeURIComponent(p).toLowerCase())
  const ct = xml(files, '[Content_Types].xml')
  const defaults = new Set(all(ct, 'Default').map((d) => (d.getAttribute('Extension') ?? '').toLowerCase()))
  const overrides = new Set<string>()
  for (const ov of all(ct, 'Override')) {
    const p = (ov.getAttribute('PartName') ?? '').replace(/^\//, '')
    overrides.add(p.toLowerCase())
    if (!real(p)) out.push(`Override for a missing part: ${p}`)
  }
  const relIds = new Map<string, Set<string>>()
  for (const name of Object.keys(files)) {
    if (name.endsWith('/') || name === '[Content_Types].xml') continue
    const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase()
    if (!overrides.has(name.toLowerCase()) && !defaults.has(ext)) out.push(`No content type: ${name}`)
    let doc: Document | null = null
    if (/\.(xml|rels)$/i.test(name)) {
      try {
        doc = parseXml(files[name])
      } catch (e) {
        out.push(`${name}: ${(e as Error).message}`)
        continue
      }
    }
    const m = /^(.*)_rels\/([^/]*)\.rels$/.exec(name)
    if (!m || !doc) continue
    const source = `${m[1]}${m[2]}`
    if (source && !real(source)) out.push(`Relationships for a missing part: ${source}`)
    const ids = new Set<string>()
    for (const r of all(doc, 'Relationship')) {
      ids.add(r.getAttribute('Id') ?? '')
      if (r.getAttribute('TargetMode') === 'External') continue
      const target = resolvePart(source, r.getAttribute('Target') ?? '')
      if (!real(target)) out.push(`${name} ${r.getAttribute('Id')} → missing ${target}`)
    }
    relIds.set(source, ids)
  }
  for (const name of Object.keys(files)) {
    if (!/\.xml$/i.test(name) || name === '[Content_Types].xml') continue
    const list = parseXml(files[name]).getElementsByTagName('*')
    for (let i = 0; i < list.length; i++) {
      for (const a of Array.from(list[i].attributes)) {
        if (a.namespaceURI !== NS.r || !a.value) continue
        if (!relIds.get(name)?.has(a.value)) out.push(`${name}: ${list[i].localName} r:${a.localName}="${a.value}" has no relationship`)
      }
    }
  }
  return out
}

function mainType(files: Files): string {
  const ov = all(xml(files, '[Content_Types].xml'), 'Override').find((o) => o.getAttribute('PartName') === '/ppt/presentation.xml')
  return ov?.getAttribute('ContentType') ?? ''
}

function sldIds(files: Files): number[] {
  return kids(first(xml(files, 'ppt/presentation.xml'), 'sldIdLst'), 'sldId').map((s) => Number(s.getAttribute('id')))
}

// ---------------------------------------------------------------------------
// test decks with things the fixture lacks

/** The fixture plus comments of both kinds: legacy on slide 1, modern on slide 2. */
function withComments(bytes: Uint8Array): Uint8Array {
  const f = unzip(bytes)
  const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
  f['ppt/commentAuthors.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmAuthorLst ${P}><p:cmAuthor id="1" name="Reviewer" initials="R" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>`)
  f['ppt/comments/comment1.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst ${P}><p:cm authorId="1" dt="2026-10-20T09:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>Legacy comment</p:text></p:cm></p:cmLst>`)
  f['ppt/authors.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{11111111-1111-1111-1111-111111111111}" name="Reviewer" initials="R" userId="Reviewer" providerId="None"/></p188:authorLst>')
  f['ppt/comments/modernComment_101_1.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"/>')
  const addRel = (rels: string, id: string, type: string, target: string) => {
    f[rels] = strToU8(text(f, rels).replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`))
  }
  addRel('ppt/_rels/presentation.xml.rels', 'rId90', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors', 'commentAuthors.xml')
  addRel('ppt/_rels/presentation.xml.rels', 'rId91', 'http://schemas.microsoft.com/office/2018/10/relationships/authors', 'authors.xml')
  addRel('ppt/slides/_rels/slide1.xml.rels', 'rId90', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', '../comments/comment1.xml')
  addRel('ppt/slides/_rels/slide2.xml.rels', 'rId90', 'http://schemas.microsoft.com/office/2018/10/relationships/comments', '../comments/modernComment_101_1.xml')
  f['ppt/slides/slide2.xml'] = strToU8(text(f, 'ppt/slides/slide2.xml').replace('</p:sld>',
    '<p:extLst><p:ext uri="{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"><p188:commentRel xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" r:id="rId90"/></p:ext></p:extLst></p:sld>'))
  const pml = 'application/vnd.openxmlformats-officedocument.presentationml.'
  f['[Content_Types].xml'] = strToU8(text(f, '[Content_Types].xml').replace('</Types>',
    `<Override PartName="/ppt/commentAuthors.xml" ContentType="${pml}commentAuthors+xml"/>`
    + `<Override PartName="/ppt/comments/comment1.xml" ContentType="${pml}comments+xml"/>`
    + '<Override PartName="/ppt/authors.xml" ContentType="application/vnd.ms-powerpoint.authors+xml"/>'
    + '<Override PartName="/ppt/comments/modernComment_101_1.xml" ContentType="application/vnd.ms-powerpoint.comments+xml"/></Types>'))
  return zipSync(f)
}

/** The fixture dressed up as a macro-enabled deck. */
function withMacros(bytes: Uint8Array): Uint8Array {
  const f = unzip(bytes)
  f['ppt/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4])
  f['ppt/_rels/presentation.xml.rels'] = strToU8(text(f, 'ppt/_rels/presentation.xml.rels').replace('</Relationships>',
    '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'))
  f['[Content_Types].xml'] = strToU8(text(f, '[Content_Types].xml')
    .replace(CT.pptx, 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml')
    .replace('</Types>', '<Override PartName="/ppt/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'))
  return zipSync(f)
}

/** The same zip with ZIP64 end records, as System.IO.Packaging writes even
 * for small files: the classic end record's counts become 0xFFFF…. */
function withZip64End(bytes: Uint8Array): Uint8Array {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = bytes.length - 22
  const count = dv.getUint16(end + 10, true)
  const cdSize = dv.getUint32(end + 12, true)
  const cdOffset = dv.getUint32(end + 16, true)
  const out = new Uint8Array(bytes.length + 56 + 20)
  out.set(bytes.subarray(0, end))
  const o = new DataView(out.buffer)
  const rec = end
  o.setUint32(rec, 0x06064b50, true)
  o.setBigUint64(rec + 4, 44n, true)
  o.setUint16(rec + 12, 45, true)
  o.setUint16(rec + 14, 45, true)
  o.setBigUint64(rec + 24, BigInt(count), true)
  o.setBigUint64(rec + 32, BigInt(count), true)
  o.setBigUint64(rec + 40, BigInt(cdSize), true)
  o.setBigUint64(rec + 48, BigInt(cdOffset), true)
  o.setUint32(rec + 56, 0x07064b50, true)
  o.setBigUint64(rec + 64, BigInt(rec), true)
  o.setUint32(rec + 72, 1, true)
  const eocd = rec + 76
  o.setUint32(eocd, 0x06054b50, true)
  o.setUint16(eocd + 8, 0xffff, true)
  o.setUint16(eocd + 10, 0xffff, true)
  o.setUint32(eocd + 12, 0xffffffff, true)
  o.setUint32(eocd + 16, 0xffffffff, true)
  return out
}

/** A slide picture: a colour per slide and its number in big block digits,
 * so the order is visible when the samples are opened. */
function slidePng(n: number, w = 640, h = 360): Uint8Array {
  const DIGITS = ['111101101101111', '010110010010111', '111001111100111', '111001111001111', '101101111001001',
    '111100111001111', '111100111101111', '111001001001001', '111101111101111', '111101111001111']
  const data = new Uint8Array(w * h * 4)
  const hue = (n * 47) % 360
  const [r, g, b] = [0, 120, 240].map((o) => Math.round(110 + 90 * Math.cos(((hue - o) * Math.PI) / 180)))
  for (let i = 0; i < w * h; i++) data.set([r, g, b, 255], i * 4)
  const s = Math.floor(h / 9)
  const digits = String(n).split('').map(Number)
  const x0 = Math.floor((w - (digits.length * 4 - 1) * s) / 2)
  const y0 = Math.floor((h - 5 * s) / 2)
  digits.forEach((d, k) => {
    for (let cy = 0; cy < 5; cy++) for (let cx = 0; cx < 3; cx++) {
      if (DIGITS[d][cy * 3 + cx] !== '1') continue
      for (let y = 0; y < s; y++) for (let x = 0; x < s; x++) {
        const px = x0 + (k * 4 + cx) * s + x
        const py = y0 + cy * s + y
        data.set([255, 255, 255, 255], (py * w + px) * 4)
      }
    }
  })
  return encodePng({ width: w, height: h, data, channels: 4, depth: 8 })
}

/** The fixture deck, with formatted notes on slide 1 and link shapes on
 * slides 1 and 3 (the fixture's own links are inside text, too little of
 * it to make a click area). */
function richDeck(): Deck {
  const deck = parseDeck(fixture())
  const notes: Paragraph[] = [
    { level: 0, bullet: 'none', runs: [{ text: 'Plain, ' }, { text: 'bold', bold: true }, { text: ' and ' }, { text: 'italic', italic: true }, { text: '.' }] },
    { level: 0, bullet: 'none', runs: [] },
    { level: 0, bullet: 'none', runs: [{ text: 'See ' }, { text: 'example.com', link: { url: 'https://example.com/' } }, { text: '\tafter a tab' }] },
  ]
  const link = (id: number, name: string, x: number, l: Shape['link']): Shape =>
    ({ id, name, kind: 'shape', rect: { x, y: deck.height * 0.75, w: deck.width * 0.2, h: deck.height * 0.15 }, link: l, description: name })
  const slides: Slide[] = deck.slides.map((s) => {
    if (s.index === 0) return { ...s, notes, notesText: 'Plain, bold and italic.\n\nSee example.com\tafter a tab', shapes: [...s.shapes, link(90, 'To slide 3', deck.width * 0.1, { slide: 2 }), link(91, 'Next', deck.width * 0.7, { jump: 'next' })] }
    if (s.index === 2) return { ...s, shapes: [...s.shapes, link(90, 'Example', deck.width * 0.1, { url: 'https://example.com/', tooltip: 'Example site' }), link(91, 'Back to the agenda', deck.width * 0.7, { slide: 1 })] }
    return s
  })
  return { ...deck, slides }
}

function pictures(slides: Slide[]): PictureSlide[] {
  return slides.map((s) => ({ slide: s, bytes: slidePng(s.index + 1), type: 'png' as const }))
}

// ---------------------------------------------------------------------------

describe('zip plumbing', () => {
  it('computes the standard CRC-32', () => {
    expect(crc32(strToU8('123456789'))).toBe(0xcbf43926)
  })
})

describe('rewritePptx', () => {
  const original = fixture()
  const before = unzip(original)
  const deck = parseDeck(original)

  it('starts from a clean fixture', () => {
    expect(packageProblems(before)).toEqual([])
  })

  it('hands back the very same bytes when nothing changes', () => {
    expect(rewritePptx(original, { kind: 'pptx' }).bytes).toBe(original)
  })

  it('leaves out a hidden slide and a range, and everything only they used', () => {
    const out = rewritePptx(original, { kind: 'pptx', removeSlides: [5, 7, 8] })
    sample('sample-trimmed.pptx', out.bytes)
    const f = unzip(out.bytes)
    expect(packageProblems(f)).toEqual([])
    expect(sldIds(f)).toEqual([256, 257, 258, 259, 260, 262])
    for (const gone of ['ppt/slides/slide6.xml', 'ppt/slides/slide8.xml', 'ppt/slides/slide9.xml',
      'ppt/notesSlides/notesSlide4.xml', 'ppt/notesSlides/notesSlide5.xml', 'ppt/notesSlides/notesSlide6.xml',
      'ppt/media/media1.mp4', 'ppt/media/image2.png', 'ppt/slides/_rels/slide9.xml.rels']) {
      expect(f[gone], gone).toBeUndefined()
    }
    expect(f['ppt/media/image1.png']).toBeDefined()
    expect(f['ppt/notesMasters/notesMaster1.xml']).toBeDefined()
    const back = parseDeck(out.bytes)
    expect(back.slides.map((s) => s.title)).toEqual(['Advanced Exporter Test Deck', 'Agenda', 'Two Columns', 'A Table', 'A Picture', 'Fade and Advance'])
    expect(back.sections.map((s) => [s.name, s.slideIndexes])).toEqual([['Intro', [0, 1]], ['Body', [2, 3, 4]], ['Finale', [5]]])
    expect(back.slides[1].notesText).toBe(deck.slides[1].notesText)
    expect(back.slides[5].transition).toMatchObject({ type: 'fade', advanceAfterMs: 3000 })
    const app = xml(f, 'docProps/app.xml')
    expect(first(app, 'Slides')?.textContent).toBe('6')
    expect(first(app, 'Notes')?.textContent).toBe('3')
    expect(first(app, 'HiddenSlides')?.textContent).toBe('0')
  })

  it('cuts links that pointed at a removed slide, and keeps the rest of that slide', () => {
    const out = rewritePptx(original, { kind: 'pptx', removeSlides: [1] })
    const f = unzip(out.bytes)
    expect(packageProblems(f)).toEqual([])
    expect(text(f, 'ppt/slides/_rels/slide8.xml.rels')).not.toMatch(/relationships\/slide"/)
    expect(text(f, 'ppt/slides/slide8.xml')).not.toContain('hlinkClick')
    expect(text(f, 'ppt/slides/slide8.xml')).toContain('Back to the agenda')
    const back = parseDeck(out.bytes)
    expect(back.slides.map((s) => s.title)).not.toContain('Agenda')
    expect(back.sections[0]).toEqual({ name: 'Intro', slideIndexes: [0] })
  })

  it('drops a section it empties, and the thumbnail of a first slide that went', () => {
    const out = rewritePptx(original, { kind: 'pptx', removeSlides: [0, 1] })
    const f = unzip(out.bytes)
    expect(packageProblems(f)).toEqual([])
    expect(parseDeck(out.bytes).sections.map((s) => s.name)).toEqual(['Body', 'Finale'])
    expect(f['docProps/thumbnail.jpeg']).toBeUndefined()
    expect(text(f, '_rels/.rels')).not.toContain('thumbnail')

    const one = unzip(rewritePptx(original, { kind: 'pptx', removeSlides: [2, 3, 4, 5, 6, 7, 8] }).bytes)
    expect(packageProblems(one)).toEqual([])
    expect(all(xml(one, 'ppt/presentation.xml'), 'section').map((s) => s.getAttribute('name'))).toEqual(['Intro'])
    expect(one['docProps/thumbnail.jpeg']).toBeDefined()
  })

  it('switches the main content type for a show or a template, and changes nothing else', () => {
    for (const kind of ['ppsx', 'potx'] as const) {
      const out = rewritePptx(original, { kind })
      sample(`sample-${kind === 'ppsx' ? 'show' : 'template'}.${kind}`, out.bytes)
      const f = unzip(out.bytes)
      expect(mainType(f)).toBe(CT[kind])
      expect(packageProblems(f)).toEqual([])
      expect(Object.keys(f).sort()).toEqual(Object.keys(before).sort())
      for (const name of Object.keys(before)) {
        if (name !== '[Content_Types].xml') expect(f[name], name).toEqual(before[name])
      }
    }
  })

  it('copies untouched entries as their compressed bytes', () => {
    const out = rewritePptx(original, { kind: 'pptx', removeSlides: [8] })
    const was = new Map(readZip(original).map((e) => [e.name, e]))
    const now = new Map(readZip(out.bytes).map((e) => [e.name, e]))
    for (const name of ['ppt/slideLayouts/slideLayout1.xml', 'ppt/media/image1.png', 'ppt/slides/slide8.xml', 'ppt/theme/theme1.xml']) {
      const a = was.get(name)!
      const b = now.get(name)!
      expect(b.raw, name).toEqual(a.raw)
      expect([b.crc, b.method, b.time, b.date]).toEqual([a.crc, a.method, a.time, a.date])
    }
  })

  it('removes speaker notes and every link to them, keeping the notes master', () => {
    const out = rewritePptx(original, { kind: 'pptx', removeNotes: true })
    sample('sample-no-notes.pptx', out.bytes)
    const f = unzip(out.bytes)
    expect(packageProblems(f)).toEqual([])
    expect(Object.keys(f).filter((n) => n.includes('notesSlide'))).toEqual([])
    expect(text(f, '[Content_Types].xml')).not.toContain('notesSlide+xml')
    expect(f['ppt/notesMasters/notesMaster1.xml']).toBeDefined()
    expect(parseDeck(out.bytes).slides.every((s) => s.notesText === '')).toBe(true)
    expect(first(xml(f, 'docProps/app.xml'), 'Notes')?.textContent).toBe('0')
  })

  it('removes legacy and modern comments, their authors and their anchors', () => {
    const commented = withComments(original)
    expect(packageProblems(unzip(commented))).toEqual([])
    const out = rewritePptx(commented, { kind: 'pptx', removeComments: true })
    sample('sample-no-comments.pptx', out.bytes)
    const f = unzip(out.bytes)
    expect(packageProblems(f)).toEqual([])
    expect(Object.keys(f).filter((n) => /comment|authors/i.test(n))).toEqual([])
    expect(text(f, 'ppt/slides/slide2.xml')).not.toMatch(/commentRel|extLst/)
    expect(text(f, '[Content_Types].xml')).not.toMatch(/comment|authors/i)
    // Comments kept when not asked to remove them.
    const kept = unzip(rewritePptx(commented, { kind: 'pptx', removeSlides: [8] }).bytes)
    expect(kept['ppt/comments/modernComment_101_1.xml']).toBeDefined()
    expect(packageProblems(kept)).toEqual([])
  })

  it('clears personal names but keeps the title', () => {
    const f = unzip(rewritePptx(original, { kind: 'pptx', stripPersonal: true }).bytes)
    const core = xml(f, 'docProps/core.xml')
    expect(first(core, 'creator')?.textContent).toBe('')
    expect(first(core, 'lastModifiedBy')?.textContent).toBe('')
    expect(first(core, 'title')?.textContent).toBe('Advanced Exporter Test Deck')
  })

  it('moves a slide-show range with the slides, or widens it to all when it empties', () => {
    const f = unzip(original)
    f['ppt/presProps.xml'] = strToU8(text(f, 'ppt/presProps.xml').replace('<p:extLst>', '<p:showPr><p:sldRg st="3" end="8"/></p:showPr><p:extLst>'))
    const ranged = zipSync(f)
    const range = (b: Uint8Array) => {
      const showPr = first(xml(unzip(b), 'ppt/presProps.xml'), 'showPr')!
      const rg = first(showPr, 'sldRg')
      return rg ? [rg.getAttribute('st'), rg.getAttribute('end')] : showPr.firstElementChild?.localName
    }
    expect(range(rewritePptx(ranged, { kind: 'pptx', removeSlides: [0, 5] }).bytes)).toEqual(['2', '6'])
    expect(range(rewritePptx(ranged, { kind: 'pptx', removeSlides: [2, 3, 4, 5, 6, 7] }).bytes)).toBe('sldAll')
    expect(range(rewritePptx(ranged, { kind: 'pptx', removeSlides: [8] }).bytes)).toEqual(['3', '8'])
  })

  it('reads a package with ZIP64 end records', () => {
    const z64 = withZip64End(original)
    expect(() => readZip(z64)).toThrow(/ZIP64/)
    const out = rewritePptx(z64, { kind: 'pptx', removeSlides: [8] })
    expect(packageProblems(unzip(out.bytes))).toEqual([])
    expect(parseDeck(out.bytes).slides.length).toBe(8)
  })

  it('takes the macros out of a macro-enabled deck', () => {
    const out = rewritePptx(withMacros(original), { kind: 'pptx' })
    const f = unzip(out.bytes)
    expect(out.warnings.join(' ')).toMatch(/Macros/)
    expect(mainType(f)).toBe(CT.pptx)
    expect(f['ppt/vbaProject.bin']).toBeUndefined()
    expect(packageProblems(f)).toEqual([])
  })
})

function fakeContext(deck: Deck, pptx: Uint8Array, slides: Slide[], extra: Partial<ExportContext> = {}) {
  const files = new Map<string, Uint8Array | string>()
  const warnings: string[] = []
  const ctx = {
    deck, slides, base: 'sample', pptx,
    host: {}, images: {},
    add: (p: string, d: Uint8Array | string) => files.set(p, d),
    pdf: async () => new Uint8Array(),
    progress: () => {},
    warn: (m: string) => warnings.push(m),
    signal: new AbortController().signal,
    ...extra,
  } as unknown as ExportContext
  return { ctx, files, warnings }
}

describe('pptxCopy and splitSlides formats', () => {
  const bytes = fixture()
  const deck = parseDeck(bytes)

  it('keeps only the exported slides, less the hidden ones when asked', async () => {
    const { ctx, files } = fakeContext(deck, bytes, deck.slides.slice(3, 8))
    const v: Values = { ...defaults(pptxCopy), type: 'ppsx', removeHidden: true }
    await pptxCopy.run(ctx, v)
    const out = files.get('sample.ppsx') as Uint8Array
    expect(out).toBeInstanceOf(Uint8Array)
    expect(parseDeck(out).slides.map((s) => s.title)).toEqual(['A Table', 'A Picture', 'Fade and Advance', 'Specials: 50% & $5 #1'])
    expect(mainType(unzip(out))).toBe(CT.ppsx)
  })

  it('asks PowerPoint for each slide and names the files', async () => {
    const asked: number[] = []
    const host = { exportSlide: async (i: number) => (asked.push(i), new Uint8Array([i])) }
    const { ctx, files } = fakeContext(deck, bytes, [deck.slides[0], deck.slides[8]], { host } as never)
    await splitSlides.run(ctx, defaults(splitSlides))
    expect(asked).toEqual([0, 8])
    expect([...files.keys()]).toEqual(['sample-01.pptx', 'sample-09.pptx'])
  })
})

describe('transitionXml', () => {
  const t = (type: string, durationMs: number, extra = {}) => transitionXml({ type, durationMs, advanceOnClick: true, ...extra })

  it('writes the plain form for the three standard speeds', () => {
    expect(t('fade', 750, { advanceAfterMs: 3000 })).toBe('<p:transition spd="med" advTm="3000"><p:fade/></p:transition>')
    expect(t('wipe', 500, { dir: 'd', advanceOnClick: false })).toBe('<p:transition spd="fast" advClick="0"><p:wipe dir="d"/></p:transition>')
  })

  it('keeps an exact duration in p14:dur with a plain fallback', () => {
    const x = t('push', 1600, { dir: 'u' })
    expect(x).toContain('<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14"><p:transition spd="slow" p14:dur="1600"><p:push dir="u"/>')
    expect(x).toContain('<mc:Fallback><p:transition spd="slow"><p:push dir="u"/></p:transition></mc:Fallback>')
  })

  it('turns newer effects into a fade and drops directions an effect does not take', () => {
    expect(t('vortex', 1000, { dir: 'r' })).toBe('<p:transition spd="slow"><p:fade/></p:transition>')
    expect(t('push', 500, { dir: 'lu' })).toBe('<p:transition spd="fast"><p:push/></p:transition>')
    expect(t('none', 500, { advanceAfterMs: 2000 })).toBe('<p:transition advTm="2000"/>')
  })
})

describe('buildFlatPptx', () => {
  const deck = richDeck()

  it('builds a valid deck that reads back with the same slides, size, hidden flags and notes', () => {
    const { bytes, droppedLinks } = buildFlatPptx(deck, pictures(deck.slides), { notes: true, transitions: true, links: true })
    sample('sample-flattened.pptx', bytes)
    expect(droppedLinks).toBe(0)
    const f = unzip(bytes)
    expect(packageProblems(f)).toEqual([])
    expect(Object.keys(f)).toContain('ppt/notesMasters/notesMaster1.xml')
    const back = parseDeck(bytes)
    expect([back.width, back.height]).toEqual([deck.width, deck.height])
    expect(back.slides.length).toBe(9)
    expect(back.slides.map((s) => s.hidden)).toEqual(deck.slides.map((s) => s.hidden))
    expect(back.slides.map((s) => s.notesText)).toEqual(deck.slides.map((s) => s.notesText))
    expect(back.slides[0].notes[0].runs.find((r) => r.text === 'bold')?.bold).toBe(true)
    expect(back.slides[6].transition).toMatchObject({ type: 'fade', durationMs: 750, advanceAfterMs: 3000 })
    expect(back.slides[7].transition).toMatchObject({ type: 'push', durationMs: 1600, dir: 'u' })
    expect(back.slides[0].transition).toBeUndefined()
    expect(back.title).toBe(deck.title)
  })

  it('puts each picture full-bleed with the title as alt text and a hidden title', () => {
    const f = unzip(buildFlatPptx(deck, pictures(deck.slides), { notes: true, transitions: true, links: true }).bytes)
    const slide = xml(f, 'ppt/slides/slide4.xml')
    const pic = first(slide, 'pic')!
    expect(first(pic, 'cNvPr')?.getAttribute('descr')).toBe('A Table')
    expect(first(first(pic, 'xfrm'), 'ext')?.getAttribute('cx')).toBe(String(deck.width))
    const title = all(slide, 'sp').find((sp) => first(sp, 'ph')?.getAttribute('type') === 'title')!
    expect(first(title, 'cNvPr')?.getAttribute('hidden')).toBe('1')
    expect(title.textContent).toBe('A Table')
    expect(f['ppt/media/slide4.png']).toEqual(slidePng(4))
  })

  it('re-creates links as clear shapes: web, slide jump and show jump', () => {
    const bytes = buildFlatPptx(deck, pictures(deck.slides), { notes: true, transitions: true, links: true }).bytes
    const back = parseDeck(bytes)
    expect(back.slides[0].shapes.filter((s) => s.link).map((s) => s.link)).toEqual([{ slide: 2, tooltip: undefined }, { jump: 'next', tooltip: undefined }])
    expect(back.slides[2].shapes.filter((s) => s.link).map((s) => s.link)).toEqual([
      { url: 'https://example.com/', tooltip: 'Example site' }, { slide: 1, tooltip: undefined }])
    const fill = first(xml(unzip(bytes), 'ppt/slides/slide1.xml'), 'alpha')
    expect(fill?.getAttribute('val')).toBe('0')
  })

  it('numbers a subset afresh and drops links to slides that are not in it', () => {
    const subset = [deck.slides[0], deck.slides[2], deck.slides[5]]
    const { bytes, droppedLinks } = buildFlatPptx(deck, pictures(subset), { notes: true, transitions: true, links: true })
    expect(droppedLinks).toBe(1) // slide 3 → the agenda, which isn't in the subset
    const back = parseDeck(bytes)
    expect(back.slides.map((s) => s.hidden)).toEqual([false, false, true])
    expect(back.slides[0].shapes.find((s) => s.link?.slide != null)?.link?.slide).toBe(1)
    expect(packageProblems(unzip(bytes))).toEqual([])
  })

  it('leaves out the notes master when there are no notes to keep', () => {
    const { bytes } = buildFlatPptx(deck, pictures(deck.slides), { notes: false, transitions: false, links: false })
    const f = unzip(bytes)
    expect(packageProblems(f)).toEqual([])
    expect(Object.keys(f).filter((n) => /notes/i.test(n))).toEqual([])
    expect(parseDeck(bytes).slides.every((s) => !s.transition && !s.notesText)).toBe(true)
  })
})

describe('buildNotesDocx', () => {
  const deck = richDeck()
  const pics = pictures(deck.slides)

  for (const layout of ['below', 'beside', 'lines'] as HandoutLayout[]) {
    it(`builds a valid ${layout} handout`, () => {
      const bytes = buildNotesDocx(deck, pics, { layout, imageWidthCm: layout === 'beside' ? 8 : 15, page: 'a4', outline: layout === 'beside' })
      sample(`sample-notes-${layout}.docx`, bytes)
      const f = unzip(bytes)
      expect(packageProblems(f)).toEqual([])
      expect(Object.keys(f).sort()).toEqual([
        '[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml',
        ...deck.slides.map((s) => `word/media/slide-${String(s.index + 1).padStart(3, '0')}.png`),
        'word/_rels/document.xml.rels', 'word/document.xml', 'word/settings.xml', 'word/styles.xml',
      ].sort())
      const doc = xml(f, 'word/document.xml')
      expect(all(doc, 'drawing').length).toBe(9)
      const headings = all(doc, 'p').filter((p) => first(p, 'pStyle')?.getAttribute('w:val') === 'Heading2')
      expect(headings.map((h) => h.textContent)[1]).toBe('Slide 2 — Agenda')
      expect(headings[5].textContent).toBe('Slide 6 — Hidden Slide (hidden)')
      const extent = first(doc, 'extent')!
      const cx = Number(extent.getAttribute('cx'))
      expect(cx).toBe(Math.round((layout === 'beside' ? 8 : 15) * 1440 / 2.54) * 635)
      expect(Number(extent.getAttribute('cy'))).toBe(Math.round(cx * deck.height / deck.width))
      expect(first(doc, 'docPr')?.getAttribute('descr')).toBe('Advanced Exporter Test Deck')
      const body = doc.documentElement.textContent ?? ''
      if (layout === 'lines') {
        expect(body).not.toContain('Agenda notes.')
        expect(all(doc, 'between').length).toBeGreaterThan(9 * 3)
      } else {
        expect(body).toContain('Agenda notes.')
        expect(body).toContain('日本語のノート。')
        const bold = all(doc, 'r').find((r) => r.textContent === 'bold')!
        expect(first(bold, 'b')).not.toBeNull()
        expect(all(doc, 'hyperlink').length).toBeGreaterThan(0)
      }
      if (layout === 'beside') {
        expect(all(doc, 'tbl').length).toBeGreaterThanOrEqual(9)
        expect(body).toContain('Projector') // the table's words, from the outline
      }
    })
  }

  it('keeps blank lines in the notes', () => {
    const f = unzip(buildNotesDocx(deck, [pics[1]], { layout: 'below', imageWidthCm: 15, page: 'letter', outline: false }))
    const doc = xml(f, 'word/document.xml')
    const ps = all(doc, 'p').map((p) => p.textContent)
    const i = ps.indexOf('Agenda notes.')
    expect(ps.slice(i, i + 3)).toEqual(['Agenda notes.', '', 'A paragraph after a blank line.'])
    expect(first(doc, 'pgSz')?.getAttribute('w:w')).toBe('12240')
  })
})

describe('buildOdp', () => {
  const deck = richDeck()
  const bytes = buildOdp(deck, pictures(deck.slides), { notes: true })
  sample('sample.odp', bytes)

  it('starts with an uncompressed mimetype entry and no extra field', () => {
    const entries = readZip(bytes)
    expect(entries[0].name).toBe('mimetype')
    expect(entries[0].method).toBe(0)
    expect(strFromU8(bytes.subarray(30, 38))).toBe('mimetype')
    expect(strFromU8(bytes.subarray(38, 38 + ODP_MIME.length))).toBe(ODP_MIME)
  })

  it('lists every file in the manifest', () => {
    const f = unzip(bytes)
    const listed = all(xml(f, 'META-INF/manifest.xml'), 'file-entry').map((e) => e.getAttribute('manifest:full-path'))
    const files = Object.keys(f).filter((n) => n !== 'mimetype' && n !== 'META-INF/manifest.xml')
    expect(listed.filter((p) => p !== '/').sort()).toEqual(files.sort())
    for (const n of ['content.xml', 'styles.xml', 'meta.xml']) expect(() => xml(f, n)).not.toThrow()
  })

  it('has a page per slide, full-size, with titles, hidden flags and notes', () => {
    const f = unzip(bytes)
    const content = xml(f, 'content.xml')
    const pages = all(content, 'page')
    expect(pages.length).toBe(9)
    expect(pages.map((p) => p.getAttribute('draw:style-name'))).toEqual(deck.slides.map((s) => (s.hidden ? 'dp2' : 'dp1')))
    const frame = kids(pages[3], 'frame')[0]
    expect(frame.getAttribute('svg:width')).toBe('33.867cm')
    expect(first(frame, 'title')?.textContent).toBe('A Table')
    expect(first(frame, 'image')?.getAttribute('xlink:href')).toBe('Pictures/slide-004.png')
    const notes = first(pages[1], 'notes')!
    expect(all(notes, 'p').map((p) => p.textContent)).toEqual(['Agenda notes.', '', 'A paragraph after a blank line.'])
    expect(first(pages[2], 'notes')).toBeNull() // slide 3 has no notes
    expect(all(first(pages[0], 'notes'), 'span').map((s) => s.textContent)).toEqual(['bold', 'italic'])
    expect(first(first(pages[0], 'notes'), 'tab')).not.toBeNull()
    const layout = first(xml(f, 'styles.xml'), 'page-layout-properties')!
    expect([layout.getAttribute('fo:page-width'), layout.getAttribute('fo:page-height')]).toEqual(['33.867cm', '19.050cm'])
  })
})
