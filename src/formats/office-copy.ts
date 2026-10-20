// A copy of the deck's own package: some slides left out, notes or comments
// taken out, saved as a show or a template.
//
// Deleting a slide by hand in the XML is how decks end up with PowerPoint's
// repair dialog, so this works the way the package is wired: decide which
// parts are dead, cut every relationship that points at one (and the markup
// that used it — a link to a removed slide, a modern comment's anchor), then
// keep only what can still be reached from the package root. Anything not
// cut is copied byte for byte, compressed bytes included.

import { unzipSync } from 'fflate'
import { resolvePart } from '../deck/parse'
import { all, first, kids, NS, parseXml, rAttr } from '../lib/xml'
import { CT, entryData, makeEntry, readZip, REL, relsPathFor, writeZip, XML_DECL, type ZipEntry } from './ooxml-write'

export type PptxKind = 'pptx' | 'ppsx' | 'potx'

export interface CopyOptions {
  kind: PptxKind
  /** Zero-based positions in the deck (sldIdLst order) to leave out. */
  removeSlides?: number[]
  removeNotes?: boolean
  removeComments?: boolean
  /** Clear author, last-modified-by, manager and company. */
  stripPersonal?: boolean
}

export interface CopyResult {
  bytes: Uint8Array
  warnings: string[]
}

const STRICT_R = 'http://purl.oclc.org/ooxml/officeDocument/relationships'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const P14 = NS.p14
const VBA = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
const MODERN_COMMENTS = 'http://schemas.microsoft.com/office/2018/10/relationships/comments'
const MODERN_AUTHORS = 'http://schemas.microsoft.com/office/2018/10/relationships/authors'

interface Rel {
  el: Element
  id: string
  type: string
  /** Resolved part name for internal targets; the raw address for external ones. */
  target: string
  external: boolean
}

/** The package being edited: zip entries plus parsed XML on demand. */
class Pkg {
  private byName = new Map<string, ZipEntry>()
  private lower = new Map<string, string>()
  private docs = new Map<string, Document>()
  readonly dirty = new Set<string>()

  constructor(readonly entries: ZipEntry[]) {
    for (const e of entries) {
      this.byName.set(e.name, e)
      this.lower.set(e.name.toLowerCase(), e.name)
    }
  }

  /** The zip entry a part name refers to: OPC names are case-insensitive and
   * relationship targets may be percent-encoded. */
  find(part: string): string | undefined {
    if (this.byName.has(part)) return part
    let decoded = part
    try {
      decoded = decodeURIComponent(part)
    } catch { /* keep as is */ }
    return this.lower.get(decoded.toLowerCase()) ?? this.lower.get(part.toLowerCase())
  }

  doc(name: string): Document | null {
    const real = this.find(name)
    if (!real) return null
    let d = this.docs.get(real)
    if (!d) {
      d = parseXml(entryData(this.byName.get(real)!))
      this.docs.set(real, d)
    }
    return d
  }

  touch(name: string): void {
    const real = this.find(name)
    if (real) this.dirty.add(real)
  }

  rels(part: string): Rel[] {
    const d = this.doc(relsPathFor(part))
    if (!d) return []
    return all(d, 'Relationship').map((el) => {
      const external = el.getAttribute('TargetMode') === 'External'
      const raw = el.getAttribute('Target') ?? ''
      return {
        el,
        id: el.getAttribute('Id') ?? '',
        type: el.getAttribute('Type') ?? '',
        target: external ? raw : resolvePart(part, raw.split('#')[0]),
        external,
      }
    })
  }

  serialize(name: string): Uint8Array {
    const xml = new XMLSerializer().serializeToString(this.docs.get(name)!).replace(/^<\?xml[^>]*\?>\s*/, '')
    return new TextEncoder().encode(XML_DECL + xml)
  }
}

/** Raw entries, so untouched parts are copied as they are. A package with
 * ZIP64 records (System.IO.Packaging writes them even for small files) is
 * read by fflate instead and re-compressed. */
function entriesOf(input: Uint8Array): ZipEntry[] {
  try {
    return readZip(input)
  } catch (e) {
    if (!/ZIP64/.test((e as Error).message)) throw e
    return Object.entries(unzipSync(input)).map(([name, data]) => makeEntry(name, data))
  }
}

export function rewritePptx(input: Uint8Array, o: CopyOptions): CopyResult {
  const warnings: string[] = []
  const pkg = new Pkg(entriesOf(input))
  const types = pkg.doc('[Content_Types].xml')
  if (!types) throw new Error('This package has no [Content_Types].xml.')
  const main = pkg.rels('').find((r) => r.type.endsWith('/officeDocument'))
  const presName = main && pkg.find(main.target)
  const pres = presName ? pkg.doc(presName) : null
  if (!presName || !pres) throw new Error('This does not look like a PowerPoint file (no presentation part).')

  const mainOverride = all(types, 'Override').find((ov) => pkg.find(partName(ov)) === presName)
  const mainType = mainOverride?.getAttribute('ContentType') ?? ''
  const macro = /macroEnabled/i.test(mainType)

  // Which slides go: their parts, and their p:sldId numbers for the sections.
  const presRels = new Map(pkg.rels(presName).map((r) => [r.id, r]))
  const sldIds = kids(first(pres, 'sldIdLst'), 'sldId')
  const dropAt = new Set((o.removeSlides ?? []).filter((i) => i >= 0 && i < sldIds.length))
  const deadParts = new Set<string>()
  const deadSldIds = new Set<string>()
  sldIds.forEach((el, i) => {
    if (!dropAt.has(i)) return
    const r = presRels.get(rAttr(el, 'id') ?? '')
    const part = r && !r.external ? pkg.find(r.target) : undefined
    if (part) deadParts.add(part)
    deadSldIds.add(el.getAttribute('id') ?? '')
  })

  const changesContent = dropAt.size > 0 || !!o.removeNotes || !!o.removeComments || !!o.stripPersonal || macro
  if (!changesContent && mainType === CT[o.kind]) return { bytes: input, warnings }
  if (macro) warnings.push(`Macros were left out: a .${o.kind} can't carry them.`)

  const deadType = (type: string): boolean =>
    (!!o.removeNotes && type.endsWith('/notesSlide'))
    || (!!o.removeComments && isCommentRel(type))
    || (macro && type === VBA)
    // The thumbnail is a picture of slide 1; don't keep one of a slide that's gone.
    || (dropAt.has(0) && type === REL.thumbnail)
    // Any signature is broken by the first changed byte.
    || type === REL.signatureOrigin

  // Walk from the root, cutting dead relationships as they're met.
  const reachable = new Set<string>()
  const cut = new Map<string, Set<string>>()
  const queue = ['']
  while (queue.length) {
    const src = queue.pop()!
    for (const r of pkg.rels(src)) {
      const target = r.external ? undefined : pkg.find(r.target)
      if (deadType(r.type) || (target && deadParts.has(target))) {
        if (!cut.has(src)) cut.set(src, new Set())
        cut.get(src)!.add(r.id)
        continue
      }
      if (target && !reachable.has(target)) {
        reachable.add(target)
        queue.push(target)
      }
    }
  }

  for (const [src, ids] of cut) {
    const relsName = pkg.find(relsPathFor(src))!
    for (const r of pkg.rels(src)) if (ids.has(r.id)) r.el.remove()
    pkg.touch(relsName)
    const d = src && /\.xml$/i.test(src) ? pkg.doc(src) : null
    if (d && stripRefs(d, ids)) pkg.touch(src)
  }

  // The presentation part: empty slide list, sections.
  const lst = first(pres, 'sldIdLst')
  if (lst && !lst.firstElementChild) lst.remove()
  if (deadSldIds.size && pruneSections(pres, deadSldIds)) pkg.touch(presName)

  const kept = (name: string): boolean => {
    if (name === '[Content_Types].xml' || name === '_rels/.rels') return true
    if (reachable.has(name)) return true
    const m = /^(.*)_rels\/([^/]+)\.rels$/.exec(name)
    return !!m && reachable.has(`${m[1]}${m[2]}`)
  }
  const survivors = pkg.entries.filter((e) => kept(e.name))
  const names = new Set(survivors.map((e) => e.name))

  // Content types: the main part's type, and no Override for a part that's gone.
  let typesChanged = false
  for (const ov of all(types, 'Override')) {
    const real = pkg.find(partName(ov))
    if (!real || !names.has(real)) {
      ov.remove()
      typesChanged = true
    }
  }
  if (mainOverride && mainType !== CT[o.kind]) {
    mainOverride.setAttribute('ContentType', CT[o.kind])
    typesChanged = true
  }
  if (typesChanged) pkg.touch('[Content_Types].xml')

  updateProps(pkg, presName, o, [...dropAt].sort((a, b) => a - b), sldIds.length)
  if (dropAt.size) renumberShowRange(pkg, presName, dropAt)

  const out = survivors.map((e) => (pkg.dirty.has(e.name) ? makeEntry(e.name, pkg.serialize(e.name)) : e))
  return { bytes: writeZip(out), warnings }
}

function partName(ov: Element): string {
  return (ov.getAttribute('PartName') ?? '').replace(/^\//, '')
}

function isCommentRel(type: string): boolean {
  return type === MODERN_COMMENTS || type === MODERN_AUTHORS || /\/(comments|commentAuthors)$/.test(type)
}

/** Take out every element that used one of these relationship ids — a link
 * to a removed slide, a custom show's entry, a modern comment's anchor. */
function stripRefs(doc: Document, ids: Set<string>): boolean {
  const victims: Element[] = []
  const list = doc.getElementsByTagName('*')
  for (let i = 0; i < list.length; i++) {
    const el = list[i]
    for (const a of Array.from(el.attributes)) {
      if ((a.namespaceURI === NS.r || a.namespaceURI === STRICT_R) && ids.has(a.value)) {
        victims.push(el)
        break
      }
    }
  }
  for (const el of victims) removeTidily(el)
  return victims.length > 0
}

/** Remove an element, and the p:ext / p:extLst wrapper it leaves empty. */
function removeTidily(el: Element): void {
  let parent: Element | null = el.parentElement
  el.remove()
  while (parent && (parent.localName === 'ext' || parent.localName === 'extLst') && !parent.firstElementChild) {
    const up: Element | null = parent.parentElement
    parent.remove()
    parent = up
  }
}

/** Drop removed slides from p14 sections, and the sections that were emptied. */
function pruneSections(pres: Document, dead: Set<string>): boolean {
  let changed = false
  for (const sec of all(pres, 'section').filter((s) => s.namespaceURI === P14)) {
    const ids = all(sec, 'sldId')
    if (!ids.length) continue
    for (const id of ids) {
      if (dead.has(id.getAttribute('id') ?? '')) {
        id.remove()
        changed = true
      }
    }
    if (!all(sec, 'sldId').length) sec.remove()
  }
  for (const lst of all(pres, 'sectionLst').filter((s) => s.namespaceURI === P14)) {
    if (!lst.firstElementChild) removeTidily(lst)
  }
  return changed
}

/** docProps: the counts PowerPoint keeps in app.xml, the titles of slides
 * that are gone (they'd leak what was taken out), and the personal fields. */
function updateProps(pkg: Pkg, presName: string, o: CopyOptions, dropped: number[], originalCount: number): void {
  const root = pkg.rels('')
  const appRel = root.find((r) => r.type === REL.extendedProps)
  const appName = appRel && pkg.find(appRel.target)
  const app = appName ? pkg.doc(appName) : null
  if (app && appName && (dropped.length || o.removeNotes || o.stripPersonal)) {
    const slides = remainingSlides(pkg, presName)
    setText(app, 'Slides', slides.length)
    setText(app, 'Notes', slides.filter((s) => pkg.rels(s).some((r) => r.type.endsWith('/notesSlide') && pkg.find(r.target))).length)
    setText(app, 'HiddenSlides', slides.filter((s) => /^(0|false)$/.test(pkg.doc(s)?.documentElement.getAttribute('show') ?? '')).length)
    if (dropped.length) pruneTitles(app, dropped, originalCount)
    if (o.stripPersonal) {
      setText(app, 'Manager', '')
      setText(app, 'Company', '')
    }
    pkg.touch(appName)
  }
  if (o.stripPersonal) {
    const coreRel = root.find((r) => r.type === REL.coreProps)
    const coreName = coreRel && pkg.find(coreRel.target)
    const core = coreName ? pkg.doc(coreName) : null
    if (core && coreName) {
      for (const el of [...all(core, 'creator'), ...all(core, 'lastModifiedBy')]) el.textContent = ''
      pkg.touch(coreName)
    }
  }
}

/** Set Up Slide Show's "from slide … to …" is by position, so it moves with
 * the slides that went before it; a range with nothing left becomes "all". */
function renumberShowRange(pkg: Pkg, presName: string, dropped: Set<number>): void {
  const rel = pkg.rels(presName).find((r) => r.type === REL.presProps)
  const name = rel && pkg.find(rel.target)
  const rg = name ? first(pkg.doc(name), 'sldRg') : null
  if (!name || !rg) return
  const st = parseInt(rg.getAttribute('st') ?? '1', 10)
  const end = parseInt(rg.getAttribute('end') ?? '1', 10)
  const keptBefore = (n: number) => { // kept slides among the first n positions
    let k = 0
    for (let i = 0; i < n; i++) if (!dropped.has(i)) k++
    return k
  }
  const newSt = keptBefore(st - 1) + 1
  const newEnd = keptBefore(end)
  if (newEnd >= newSt) {
    rg.setAttribute('st', String(newSt))
    rg.setAttribute('end', String(newEnd))
  } else {
    rg.parentNode!.replaceChild(rg.ownerDocument.createElementNS(NS_P, 'p:sldAll'), rg)
  }
  pkg.touch(name)
}

function remainingSlides(pkg: Pkg, presName: string): string[] {
  const pres = pkg.doc(presName)!
  const rels = new Map(pkg.rels(presName).map((r) => [r.id, r]))
  return kids(first(pres, 'sldIdLst'), 'sldId')
    .map((el) => rels.get(rAttr(el, 'id') ?? ''))
    .map((r) => (r ? pkg.find(r.target) : undefined))
    .filter((p): p is string => !!p)
}

function setText(doc: Document, local: string, value: string | number): void {
  const el = kids(doc.documentElement, local)[0]
  if (el) el.textContent = String(value)
}

/** app.xml lists every slide's title under a "Slide Titles" heading pair (the
 * name is localised, so failing that: the last group, if it has one entry per
 * slide). Take out the removed slides' entries. */
function pruneTitles(app: Document, dropped: number[], originalCount: number): void {
  const pairs = kids(first(first(app, 'HeadingPairs'), 'vector'), 'variant')
  const titles = first(first(app, 'TitlesOfParts'), 'vector')
  if (!titles) return
  const entries = Array.from(titles.children)
  let offset = 0
  for (let i = 0; i + 1 < pairs.length; i += 2) {
    const name = first(pairs[i], 'lpstr')?.textContent ?? ''
    const countEl = first(pairs[i + 1], 'i4')
    const count = parseInt(countEl?.textContent ?? '0', 10) || 0
    const last = i + 2 >= pairs.length
    if (countEl && count === originalCount && count > 0 && (name === 'Slide Titles' || last)) {
      if (offset + count > entries.length) return
      for (const at of dropped) entries[offset + at]?.remove()
      countEl.textContent = String(count - dropped.length)
      titles.setAttribute('size', String(titles.children.length))
      return
    }
    offset += count
  }
}
