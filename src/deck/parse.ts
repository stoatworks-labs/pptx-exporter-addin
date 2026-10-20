// Read a .pptx package into a Deck.
//
// Only the XML parts are inflated up front. Media (a deck can carry a gigabyte
// of video) is left compressed until a format actually asks for it.

import { unzipSync } from 'fflate'
import {
  all, attrLocal, boolAttr, first, intAttr, kid, kids, parseXml, path, rAttr,
} from '../lib/xml'
import type {
  Deck, ImageRef, Link, MediaRef, Paragraph, PlaceholderType, Rect, Run, Section, Shape,
  Slide, TableCell, Transition,
} from './model'

const REL = {
  slide: '/slide',
  layout: '/slideLayout',
  master: '/slideMaster',
  notes: '/notesSlide',
  image: '/image',
  hyperlink: '/hyperlink',
  video: '/video',
  audio: '/audio',
  media: '/media',
  officeDocument: '/officeDocument',
  coreProps: '/core-properties',
  presProps: '/presProps',
}

interface Rel {
  id: string
  type: string
  target: string
  external: boolean
}

const MONO_FACES = /^(courier|consolas|menlo|monaco|andale mono|lucida console|source code|fira (code|mono)|jetbrains mono|sf mono|cascadia|dejavu sans mono|liberation mono|roboto mono|ubuntu mono|inconsolata|ibm plex mono)/i

const SPEED_MS: Record<string, number> = { fast: 500, med: 750, slow: 1000 }

class Package {
  private xml = new Map<string, Uint8Array>()
  private docs = new Map<string, Document>()
  private relsCache = new Map<string, Map<string, Rel>>()
  private types = { defaults: new Map<string, string>(), overrides: new Map<string, string>() }

  constructor(private bytes: Uint8Array) {
    const files = unzipSync(bytes, { filter: (f) => /\.(xml|rels)$/i.test(f.name) })
    for (const [name, data] of Object.entries(files)) this.xml.set(name, data)
    const ct = this.doc('[Content_Types].xml')
    if (ct) {
      for (const d of all(ct, 'Default')) {
        this.types.defaults.set((d.getAttribute('Extension') ?? '').toLowerCase(), d.getAttribute('ContentType') ?? '')
      }
      for (const o of all(ct, 'Override')) {
        this.types.overrides.set((o.getAttribute('PartName') ?? '').replace(/^\//, ''), o.getAttribute('ContentType') ?? '')
      }
    }
  }

  has(part: string): boolean {
    return this.xml.has(part)
  }

  doc(part: string): Document | null {
    let d = this.docs.get(part)
    if (d) return d
    const raw = this.xml.get(part)
    if (!raw) return null
    d = parseXml(raw)
    this.docs.set(part, d)
    return d
  }

  rels(part: string): Map<string, Rel> {
    const cached = this.relsCache.get(part)
    if (cached) return cached
    const slash = part.lastIndexOf('/')
    const relsPart = `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`
    const map = new Map<string, Rel>()
    const d = this.doc(relsPart)
    if (d) {
      for (const r of all(d, 'Relationship')) {
        const external = r.getAttribute('TargetMode') === 'External'
        const target = r.getAttribute('Target') ?? ''
        map.set(r.getAttribute('Id') ?? '', {
          id: r.getAttribute('Id') ?? '',
          type: r.getAttribute('Type') ?? '',
          target: external ? target : resolvePart(part, target),
          external,
        })
      }
    }
    this.relsCache.set(part, map)
    return map
  }

  relOfType(part: string, suffix: string): Rel | undefined {
    for (const r of this.rels(part).values()) if (r.type.endsWith(suffix)) return r
    return undefined
  }

  read(part: string): Uint8Array | undefined {
    if (this.xml.has(part)) return this.xml.get(part)
    const files = unzipSync(this.bytes, { filter: (f) => f.name === part })
    return files[part]
  }

  contentType(part: string): string | undefined {
    const o = this.types.overrides.get(part)
    if (o) return o
    const ext = part.slice(part.lastIndexOf('.') + 1).toLowerCase()
    return this.types.defaults.get(ext) ?? guessType(ext)
  }
}

function guessType(ext: string): string | undefined {
  return ({
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
    tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf',
    mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', mp3: 'audio/mpeg',
    m4a: 'audio/mp4', wav: 'audio/wav', wmv: 'video/x-ms-wmv', avi: 'video/x-msvideo',
  } as Record<string, string>)[ext]
}

/** Resolve a relationship target against the part that owns it. */
export function resolvePart(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = fromPart.split('/').slice(0, -1)
  for (const seg of target.split('/')) {
    if (seg === '..') base.pop()
    else if (seg !== '.' && seg !== '') base.push(seg)
  }
  return base.join('/')
}

export function parseDeck(bytes: Uint8Array, fallbackTitle = 'Presentation'): Deck {
  const pkg = new Package(bytes)
  const presPart = pkg.relOfType('', REL.officeDocument)?.target ?? 'ppt/presentation.xml'
  const pres = pkg.doc(presPart)
  if (!pres) throw new Error('This does not look like a PowerPoint file (no presentation part).')

  const sz = first(pres, 'sldSz')
  const width = intAttr(sz, 'cx') ?? 12192000
  const height = intAttr(sz, 'cy') ?? 6858000

  const presRels = pkg.rels(presPart)
  const sldIds = kids(first(pres, 'sldIdLst'), 'sldId')
  const slideParts: { sldId: number; part: string }[] = []
  for (const s of sldIds) {
    const rel = presRels.get(rAttr(s, 'id') ?? '')
    if (rel) slideParts.push({ sldId: intAttr(s, 'id') ?? 0, part: rel.target })
  }
  const indexOfPart = new Map(slideParts.map((s, i) => [s.part, i]))

  const ctx: Ctx = { pkg, indexOfPart, slideCount: slideParts.length }
  const slides = slideParts.map((s, i) => parseSlide(ctx, s.part, s.sldId, i))

  // Sections (PowerPoint 2010+, an extension list on the presentation).
  const sections: Section[] = []
  const byId = new Map(slides.map((s) => [s.sldId, s.index]))
  for (const sec of all(pres, 'section')) {
    const name = sec.getAttribute('name') ?? ''
    const idxs: number[] = []
    for (const id of all(sec, 'sldId')) {
      const i = byId.get(intAttr(id, 'id') ?? -1)
      if (i != null) idxs.push(i)
    }
    sections.push({ name, slideIndexes: idxs })
    for (const i of idxs) slides[i].section = name
  }

  // Title and author.
  let title = ''
  let author: string | undefined
  const coreRel = pkg.relOfType('', REL.coreProps)
  const core = coreRel ? pkg.doc(coreRel.target) : pkg.doc('docProps/core.xml')
  if (core) {
    title = first(core, 'title')?.textContent?.trim() ?? ''
    author = first(core, 'creator')?.textContent?.trim() || undefined
  }

  // Show settings: loop and use-timings.
  const show = { loop: false, useTimings: true }
  const ppRel = pkg.relOfType(presPart, REL.presProps)
  const pp = ppRel ? pkg.doc(ppRel.target) : null
  const showPr = first(pp, 'showPr')
  if (showPr) {
    show.loop = boolAttr(showPr, 'loop') ?? false
    show.useTimings = boolAttr(showPr, 'useTimings') ?? true
  }

  return {
    title: title || fallbackTitle,
    author,
    lang: deckLanguage(pkg, slideParts.map((s) => s.part)),
    width,
    height,
    slides,
    sections,
    show,
    readPart: (p) => pkg.read(p),
    contentTypeOf: (p) => pkg.contentType(p),
  }
}

/** The commonest run language on the slides — what an exported page's
 * <html lang> should say, for screen readers and hyphenation. */
function deckLanguage(pkg: Package, parts: string[]): string | undefined {
  const counts = new Map<string, number>()
  for (const part of parts) {
    for (const el of all(pkg.doc(part), 'rPr').concat(all(pkg.doc(part), 'endParaRPr'))) {
      const lang = el.getAttribute('lang')
      if (lang && /^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$/.test(lang)) counts.set(lang, (counts.get(lang) ?? 0) + 1)
    }
  }
  let best: string | undefined
  let n = 0
  for (const [l, c] of counts) if (c > n) [best, n] = [l, c]
  return best
}

interface Ctx {
  pkg: Package
  indexOfPart: Map<string, number>
  slideCount: number
}

interface PlaceholderGeom {
  type: string
  idx?: number
  rect?: Rect
}

/** Placeholders on a layout or master, for slides that inherit their position. */
function placeholderGeometry(pkg: Package, part: string | undefined): PlaceholderGeom[] {
  if (!part) return []
  const d = pkg.doc(part)
  if (!d) return []
  const out: PlaceholderGeom[] = []
  for (const sp of all(d, 'sp')) {
    const ph = path(sp, 'nvSpPr', 'nvPr', 'ph')
    if (!ph) continue
    out.push({
      type: ph.getAttribute('type') ?? 'obj',
      idx: intAttr(ph, 'idx'),
      rect: xfrmRect(path(sp, 'spPr', 'xfrm')),
    })
  }
  return out
}

function phFamily(t: string): string {
  if (t === 'title' || t === 'ctrTitle') return 'title'
  if (t === 'body' || t === 'obj' || t === 'subTitle') return 'body'
  return t
}

function inheritRect(ph: { type: string; idx?: number }, layout: PlaceholderGeom[], master: PlaceholderGeom[]): Rect | undefined {
  const byIdx = ph.idx != null ? layout.find((l) => l.idx === ph.idx) : undefined
  const byType = layout.find((l) => l.type === ph.type) ?? layout.find((l) => phFamily(l.type) === phFamily(ph.type))
  const cand = byIdx ?? byType
  if (cand?.rect) return cand.rect
  const fam = phFamily(cand?.type ?? ph.type)
  return master.find((m) => phFamily(m.type) === fam && m.rect)?.rect
}

function xfrmRect(x: Element | null): Rect | undefined {
  if (!x) return undefined
  const off = kid(x, 'off')
  const ext = kid(x, 'ext')
  if (!off || !ext) return undefined
  return {
    x: intAttr(off, 'x') ?? 0,
    y: intAttr(off, 'y') ?? 0,
    w: intAttr(ext, 'cx') ?? 0,
    h: intAttr(ext, 'cy') ?? 0,
  }
}

interface Xform {
  // child → parent: parent = off + (child - chOff) * scale
  ox: number
  oy: number
  sx: number
  sy: number
  cx: number
  cy: number
}

const IDENTITY: Xform = { ox: 0, oy: 0, sx: 1, sy: 1, cx: 0, cy: 0 }

function applyXform(r: Rect, t: Xform): Rect {
  return {
    x: t.ox + (r.x - t.cx) * t.sx,
    y: t.oy + (r.y - t.cy) * t.sy,
    w: r.w * t.sx,
    h: r.h * t.sy,
  }
}

function composeGroup(grpXfrm: Element | null, parent: Xform): Xform {
  if (!grpXfrm) return parent
  const off = kid(grpXfrm, 'off')
  const ext = kid(grpXfrm, 'ext')
  const chOff = kid(grpXfrm, 'chOff')
  const chExt = kid(grpXfrm, 'chExt')
  const local: Xform = {
    ox: intAttr(off, 'x') ?? 0,
    oy: intAttr(off, 'y') ?? 0,
    cx: intAttr(chOff, 'x') ?? 0,
    cy: intAttr(chOff, 'y') ?? 0,
    sx: (intAttr(chExt, 'cx') ?? 0) ? (intAttr(ext, 'cx') ?? 0) / (intAttr(chExt, 'cx') as number) : 1,
    sy: (intAttr(chExt, 'cy') ?? 0) ? (intAttr(ext, 'cy') ?? 0) / (intAttr(chExt, 'cy') as number) : 1,
  }
  // parent(local(p)) = parent.o + (local.o + (p - local.c)*local.s - parent.c) * parent.s
  return {
    ox: parent.ox + (local.ox - local.cx * local.sx - parent.cx) * parent.sx,
    oy: parent.oy + (local.oy - local.cy * local.sy - parent.cy) * parent.sy,
    sx: local.sx * parent.sx,
    sy: local.sy * parent.sy,
    cx: 0,
    cy: 0,
  }
}

function parseSlide(ctx: Ctx, part: string, sldId: number, index: number): Slide {
  const { pkg } = ctx
  const d = pkg.doc(part)
  const root = d?.documentElement ?? null
  const layoutRel = pkg.relOfType(part, REL.layout)
  const layoutPart = layoutRel?.target
  const masterPart = layoutPart ? pkg.relOfType(layoutPart, REL.master)?.target : undefined
  const layoutGeom = placeholderGeometry(pkg, layoutPart)
  const masterGeom = placeholderGeometry(pkg, masterPart)
  const layoutName = layoutPart ? first(pkg.doc(layoutPart), 'cSld')?.getAttribute('name') || undefined : undefined

  const rels = pkg.rels(part)
  const shapes: Shape[] = []
  const spTree = path(root, 'cSld', 'spTree')
  if (spTree) walkTree(ctx, part, rels, spTree, IDENTITY, layoutGeom, masterGeom, shapes)

  const titleShape = shapes.find((s) => s.placeholder && phFamily(s.placeholder.type) === 'title')
  const title = titleShape?.paragraphs
    ? titleShape.paragraphs.map((p) => p.runs.map((r) => (r.br ? ' ' : r.text)).join('')).join(' ').replace(/\s+/g, ' ').trim()
    : ''

  // Speaker notes: the body placeholder on the notes slide.
  let notes: Paragraph[] = []
  const notesRel = pkg.relOfType(part, REL.notes)
  if (notesRel) {
    const nd = pkg.doc(notesRel.target)
    const nrels = pkg.rels(notesRel.target)
    for (const sp of all(nd, 'sp')) {
      const ph = path(sp, 'nvSpPr', 'nvPr', 'ph')
      if (ph?.getAttribute('type') === 'body') {
        notes = parseTextBody(ctx, nrels, kid(sp, 'txBody'), false)
        break
      }
    }
  }
  // Trim empty paragraphs at either end; PowerPoint leaves a trailing one.
  while (notes.length && isEmptyPara(notes[notes.length - 1])) notes.pop()
  while (notes.length && isEmptyPara(notes[0])) notes.shift()
  const notesText = notes.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join('')).join('\n')

  return {
    index,
    sldId,
    part,
    hidden: root?.getAttribute('show') === '0',
    layoutName,
    title,
    shapes,
    notes,
    notesText,
    transition: parseTransition(root),
    hasAnimations: !!first(kid(root, 'timing'), 'par'),
  }
}

function isEmptyPara(p: Paragraph): boolean {
  return p.runs.every((r) => !r.text.trim())
}

function parseTransition(root: Element | null): Transition | undefined {
  if (!root) return undefined
  // Prefer the richer mc:Choice form (p14 durations, newer effects); fall back
  // to a plain p:transition.
  let tr: Element | null = null
  for (const ac of kids(root, 'AlternateContent')) {
    const choice = kid(ac, 'Choice')
    tr = kid(choice, 'transition') ?? kid(kid(ac, 'Fallback'), 'transition')
    if (tr) break
  }
  tr = tr ?? kid(root, 'transition')
  if (!tr) return undefined
  const effect = tr.firstElementChild?.localName === 'extLst' ? null : tr.firstElementChild
  const dur = attrLocal(tr, 'dur')
  const spd = tr.getAttribute('spd') ?? 'fast'
  const advTm = intAttr(tr, 'advTm')
  return {
    type: effect?.localName ?? 'none',
    durationMs: dur ? parseInt(dur, 10) : (SPEED_MS[spd] ?? 500),
    dir: effect?.getAttribute('dir') ?? undefined,
    advanceOnClick: boolAttr(tr, 'advClick') ?? true,
    advanceAfterMs: advTm,
  }
}

function walkTree(
  ctx: Ctx, part: string, rels: Map<string, Rel>, tree: Element, xf: Xform,
  layout: PlaceholderGeom[], master: PlaceholderGeom[], out: Shape[],
): void {
  for (let el = tree.firstElementChild; el; el = el.nextElementSibling) {
    // Content-control alternates: take the Choice when it holds shapes.
    if (el.localName === 'AlternateContent') {
      const holder = kid(el, 'Choice') ?? kid(el, 'Fallback')
      if (holder) walkTree(ctx, part, rels, holder, xf, layout, master, out)
      continue
    }
    if (el.localName === 'grpSp') {
      if (boolAttr(path(el, 'nvGrpSpPr', 'cNvPr'), 'hidden')) continue
      walkTree(ctx, part, rels, el, composeGroup(path(el, 'grpSpPr', 'xfrm'), xf), layout, master, out)
      continue
    }
    const s = parseShape(ctx, rels, el, layout, master)
    if (!s) continue
    if (s.rect) s.rect = applyXform(s.rect, xf)
    out.push(s)
  }
}

function parseShape(ctx: Ctx, rels: Map<string, Rel>, el: Element, layout: PlaceholderGeom[], master: PlaceholderGeom[]): Shape | null {
  const ln = el.localName
  if (ln !== 'sp' && ln !== 'pic' && ln !== 'graphicFrame' && ln !== 'cxnSp') return null
  const nv = kid(el, `nv${ln[0].toUpperCase()}${ln.slice(1)}Pr`)
  const cNvPr = kid(nv, 'cNvPr')
  if (boolAttr(cNvPr, 'hidden')) return null
  const phEl = path(nv, 'nvPr', 'ph')
  const placeholder = phEl ? { type: (phEl.getAttribute('type') ?? 'obj') as PlaceholderType, idx: intAttr(phEl, 'idx') } : undefined

  const xfrm = ln === 'graphicFrame' ? kid(el, 'xfrm') : path(el, 'spPr', 'xfrm')
  let rect = xfrmRect(xfrm)
  if (!rect && placeholder) rect = inheritRect(placeholder, layout, master)
  const rot = intAttr(xfrm, 'rot')

  const shape: Shape = {
    id: intAttr(cNvPr, 'id') ?? 0,
    name: cNvPr?.getAttribute('name') ?? '',
    kind: 'shape',
    rect,
    rotation: rot ? rot / 60000 : undefined,
    placeholder,
    description: cNvPr?.getAttribute('descr') || undefined,
    link: parseLink(ctx, rels, kid(cNvPr, 'hlinkClick')),
  }

  if (ln === 'sp') {
    const body = kid(el, 'txBody')
    // Body and content placeholders are bulleted unless a paragraph says
    // otherwise; a subtitle never is.
    const isBodyPh = !!placeholder && (placeholder.type === 'body' || placeholder.type === 'obj')
    const paragraphs = parseTextBody(ctx, rels, body, isBodyPh)
    if (paragraphs.some((p) => p.runs.some((r) => r.text.trim()))) {
      shape.kind = 'text'
      shape.paragraphs = paragraphs
    }
    return shape
  }

  if (ln === 'cxnSp') return shape

  if (ln === 'pic') {
    const blip = path(el, 'blipFill', 'blip')
    const embed = rAttr(blip, 'embed')
    const imgRel = embed ? rels.get(embed) : undefined
    if (imgRel && !imgRel.external) {
      shape.image = { part: imgRel.target, contentType: ctx.pkg.contentType(imgRel.target) ?? 'application/octet-stream' } satisfies ImageRef
    }
    const nvPr = kid(nv, 'nvPr')
    const vf = kid(nvPr, 'videoFile')
    const af = kid(nvPr, 'audioFile')
    if (vf || af) {
      const media: MediaRef = { kind: vf ? 'video' : 'audio', poster: shape.image?.part }
      // PowerPoint 2010+ embeds via p14:media r:embed; older decks link only.
      const p14 = all(nvPr, 'media').find((m) => rAttr(m, 'embed'))
      const embedId = p14 ? rAttr(p14, 'embed') : null
      const linkId = rAttr(vf ?? af, 'link')
      const er = embedId ? rels.get(embedId) : undefined
      const lr = linkId ? rels.get(linkId) : undefined
      const r = er && !er.external ? er : lr
      if (r) {
        if (r.external) media.external = r.target
        else {
          media.part = r.target
          media.contentType = ctx.pkg.contentType(r.target)
        }
      }
      shape.kind = 'media'
      shape.media = media
    } else {
      shape.kind = 'picture'
    }
    return shape
  }

  // graphicFrame: table, chart, SmartArt, OLE.
  const gd = path(el, 'graphic', 'graphicData')
  const uri = gd?.getAttribute('uri') ?? ''
  const tbl = kid(gd, 'tbl')
  if (tbl) {
    shape.kind = 'table'
    shape.table = parseTable(ctx, rels, tbl)
  } else if (uri.endsWith('/chart')) shape.kind = 'chart'
  else if (uri.endsWith('/diagram')) shape.kind = 'diagram'
  else shape.kind = 'other'
  return shape
}

function parseTable(ctx: Ctx, rels: Map<string, Rel>, tbl: Element): { rows: TableCell[][]; colWidths: number[] } {
  const colWidths = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => intAttr(g, 'w') ?? 0)
  const rows = kids(tbl, 'tr').map((tr) =>
    kids(tr, 'tc').map((tc) => ({
      paragraphs: parseTextBody(ctx, rels, kid(tc, 'txBody'), false),
      colSpan: intAttr(tc, 'gridSpan') ?? 1,
      rowSpan: intAttr(tc, 'rowSpan') ?? 1,
      merged: !!(boolAttr(tc, 'hMerge') || boolAttr(tc, 'vMerge')),
    })),
  )
  return { rows, colWidths }
}

function parseLink(ctx: Ctx, rels: Map<string, Rel>, h: Element | null): Link | undefined {
  if (!h) return undefined
  const action = h.getAttribute('action') ?? ''
  const tooltip = h.getAttribute('tooltip') ?? undefined
  const rid = rAttr(h, 'id')
  const rel = rid ? rels.get(rid) : undefined
  const jump = /jump=(\w+)/.exec(action)?.[1]
  if (jump) {
    const map: Record<string, Link['jump']> = {
      firstslide: 'first', lastslide: 'last', nextslide: 'next', previousslide: 'previous', endshow: 'end',
      lastslideviewed: 'previous',
    }
    if (map[jump]) return { jump: map[jump], tooltip }
  }
  if (rel) {
    if (rel.external) return { url: rel.target, tooltip }
    const i = ctx.indexOfPart.get(rel.target)
    if (i != null) return { slide: i, tooltip }
  }
  return undefined
}

function parseTextBody(ctx: Ctx, rels: Map<string, Rel>, body: Element | null, bulletsByDefault: boolean): Paragraph[] {
  if (!body) return []
  return kids(body, 'p').map((p) => {
    const pPr = kid(p, 'pPr')
    const level = intAttr(pPr, 'lvl') ?? 0
    let bullet: Paragraph['bullet'] = bulletsByDefault ? 'char' : 'none'
    let numberScheme: string | undefined
    if (kid(pPr, 'buNone')) bullet = 'none'
    else if (kid(pPr, 'buAutoNum')) {
      bullet = 'number'
      numberScheme = kid(pPr, 'buAutoNum')?.getAttribute('type') ?? 'arabicPeriod'
    } else if (kid(pPr, 'buChar') || kid(pPr, 'buBlip')) bullet = 'char'
    const algn = pPr?.getAttribute('algn') as Paragraph['align'] | null
    const runs: Run[] = []
    for (let c = p.firstElementChild; c; c = c.nextElementSibling) {
      if (c.localName === 'r' || c.localName === 'fld') {
        const text = kid(c, 't')?.textContent ?? ''
        if (!text) continue
        runs.push(parseRun(ctx, rels, kid(c, 'rPr'), text))
      } else if (c.localName === 'br') {
        runs.push({ text: '', br: true })
      }
    }
    // An empty paragraph never shows a bullet.
    if (!runs.some((r) => r.text.trim())) bullet = 'none'
    return { level, bullet, numberScheme, align: algn ?? undefined, runs }
  })
}

function parseRun(ctx: Ctx, rels: Map<string, Rel>, rPr: Element | null, text: string): Run {
  const run: Run = { text }
  if (!rPr) return run
  if (boolAttr(rPr, 'b')) run.bold = true
  if (boolAttr(rPr, 'i')) run.italic = true
  const u = rPr.getAttribute('u')
  if (u && u !== 'none') run.underline = true
  const strike = rPr.getAttribute('strike')
  if (strike && strike !== 'noStrike') run.strike = true
  const bl = intAttr(rPr, 'baseline')
  if (bl) run.baseline = bl
  const sz = intAttr(rPr, 'sz')
  if (sz) run.size = sz / 100
  const srgb = path(rPr, 'solidFill', 'srgbClr')?.getAttribute('val')
  if (srgb && /^[0-9a-f]{6}$/i.test(srgb)) run.color = `#${srgb.toLowerCase()}`
  const face = kid(rPr, 'latin')?.getAttribute('typeface') ?? ''
  if (MONO_FACES.test(face)) run.mono = true
  const link = parseLink(ctx, rels, kid(rPr, 'hlinkClick'))
  if (link) run.link = link
  return run
}
