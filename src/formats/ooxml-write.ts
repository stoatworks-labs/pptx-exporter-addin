// Writing Office (OOXML) and OpenDocument packages.
//
// A small zip reader and writer rather than fflate's zipSync, because the
// .pptx copy has to keep every part it doesn't touch exactly as it was: an
// untouched entry is copied as its compressed bytes, CRC and date included,
// so a deck with half a gigabyte of video costs one copy in memory rather
// than an inflate and a re-deflate of every clip. fflate still does the
// actual (de)compression. No ZIP64: nothing a task pane can hold needs it.

import { deflateSync, inflateSync } from 'fflate'
import { xmlEscape, xmlSafe } from '../lib/xml'

export interface ZipEntry {
  name: string
  /** 0 stored, 8 deflated. */
  method: number
  crc: number
  /** Uncompressed size. */
  size: number
  /** The bytes as they sit in the archive (compressed when method is 8). */
  raw: Uint8Array
  /** MS-DOS time and date fields. */
  time: number
  date: number
}

const SIG_LOCAL = 0x04034b50
const SIG_CENTRAL = 0x02014b50
const SIG_END = 0x06054b50
const UTF8_FLAG = 0x0800

const decoder = new TextDecoder('utf-8')
const encoder = new TextEncoder()

export function readZip(bytes: Uint8Array): ZipEntry[] {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let end = -1
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65535); i--) {
    if (dv.getUint32(i, true) === SIG_END) {
      end = i
      break
    }
  }
  if (end < 0) throw new Error('Not a zip package (no end-of-directory record).')
  const count = dv.getUint16(end + 10, true)
  let p = dv.getUint32(end + 16, true)
  if (count === 0xffff || p === 0xffffffff) throw new Error('ZIP64 packages are not supported.')
  const out: ZipEntry[] = []
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== SIG_CENTRAL) throw new Error('Damaged zip central directory.')
    const flags = dv.getUint16(p + 8, true)
    const method = dv.getUint16(p + 10, true)
    const csize = dv.getUint32(p + 20, true)
    const size = dv.getUint32(p + 24, true)
    const nlen = dv.getUint16(p + 28, true)
    const xlen = dv.getUint16(p + 30, true)
    const clen = dv.getUint16(p + 32, true)
    const local = dv.getUint32(p + 42, true)
    if (flags & 1) throw new Error('Encrypted packages are not supported.')
    if (csize === 0xffffffff || size === 0xffffffff || local === 0xffffffff) throw new Error('ZIP64 packages are not supported.')
    if (dv.getUint32(local, true) !== SIG_LOCAL) throw new Error('Damaged zip entry.')
    const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true)
    out.push({
      name: decoder.decode(bytes.subarray(p + 46, p + 46 + nlen)),
      method,
      crc: dv.getUint32(p + 16, true),
      size,
      raw: bytes.subarray(start, start + csize),
      time: dv.getUint16(p + 12, true),
      date: dv.getUint16(p + 14, true),
    })
    p += 46 + nlen + xlen + clen
  }
  return out
}

export function entryData(e: ZipEntry): Uint8Array {
  if (e.method === 0) return e.raw
  if (e.method === 8) return inflateSync(e.raw, { out: new Uint8Array(e.size) })
  throw new Error(`${e.name} uses zip compression method ${e.method}, which this can't read.`)
}

/** Already-compressed media: deflating it again only costs time. */
const PRECOMPRESSED = /\.(png|jpe?g|gif|webp|avif|mp4|m4v|mov|webm|mp3|m4a|wma|wmv|zip|docx|xlsx|pptx|odp)$/i

export function makeEntry(name: string, data: Uint8Array | string, when = new Date()): ZipEntry {
  const bytes = typeof data === 'string' ? encoder.encode(data) : data
  const { time, date } = dosTime(when)
  const base = { name, crc: crc32(bytes), size: bytes.length, time, date }
  if (name === 'mimetype' || PRECOMPRESSED.test(name) || bytes.length < 64) return { ...base, method: 0, raw: bytes }
  const deflated = deflateSync(bytes, { level: 6 })
  return deflated.length < bytes.length ? { ...base, method: 8, raw: deflated } : { ...base, method: 0, raw: bytes }
}

export function writeZip(entries: ZipEntry[]): Uint8Array {
  if (entries.length > 0xfffe) throw new Error('Too many files for one package.')
  const names = entries.map((e) => encoder.encode(e.name))
  let total = 22
  entries.forEach((e, i) => (total += 30 + 46 + 2 * names[i].length + e.raw.length))
  if (total > 0xffffffff) throw new Error('The package would be past the 4 GB zip limit.')
  const out = new Uint8Array(total)
  const dv = new DataView(out.buffer)
  // The fields local and central headers share, from "version needed" on.
  const common = (e: ZipEntry, name: Uint8Array, q: number): void => {
    dv.setUint16(q, 20, true)
    dv.setUint16(q + 2, /[^\x20-\x7e]/.test(e.name) ? UTF8_FLAG : 0, true)
    dv.setUint16(q + 4, e.method, true)
    dv.setUint16(q + 6, e.time, true)
    dv.setUint16(q + 8, e.date, true)
    dv.setUint32(q + 10, e.crc >>> 0, true)
    dv.setUint32(q + 14, e.raw.length, true)
    dv.setUint32(q + 18, e.size, true)
    dv.setUint16(q + 22, name.length, true)
    dv.setUint16(q + 24, 0, true) // no extra field
  }
  const offsets: number[] = []
  let p = 0
  entries.forEach((e, i) => {
    offsets.push(p)
    dv.setUint32(p, SIG_LOCAL, true)
    common(e, names[i], p + 4)
    out.set(names[i], p + 30)
    p += 30 + names[i].length
    out.set(e.raw, p)
    p += e.raw.length
  })
  const cdStart = p
  entries.forEach((e, i) => {
    dv.setUint32(p, SIG_CENTRAL, true)
    dv.setUint16(p + 4, 20, true) // made by: MS-DOS, zip 2.0
    common(e, names[i], p + 6)
    // comment length, disk number, internal and external attributes stay 0
    dv.setUint32(p + 42, offsets[i], true)
    out.set(names[i], p + 46)
    p += 46 + names[i].length
  })
  dv.setUint32(p, SIG_END, true)
  dv.setUint16(p + 8, entries.length, true)
  dv.setUint16(p + 10, entries.length, true)
  dv.setUint32(p + 12, p - cdStart, true)
  dv.setUint32(p + 16, cdStart, true)
  return out
}

function dosTime(d: Date): { time: number; date: number } {
  const y = Math.max(1980, d.getFullYear())
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
    date: ((y - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  }
}

let CRC_TABLE: Uint32Array | undefined

export function crc32(data: Uint8Array): number {
  if (!CRC_TABLE) {
    CRC_TABLE = new Uint32Array(256)
    for (let n = 0; n < 256; n++) {
      let c = n
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      CRC_TABLE[n] = c >>> 0
    }
  }
  let crc = 0xffffffff
  for (let i = 0; i < data.length; i++) crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

// ---------------------------------------------------------------------------
// XML and OPC

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

/** Text from a deck, made safe for XML content or an attribute value. */
export function esc(s: string): string {
  return xmlEscape(xmlSafe(s))
}

export const NSX = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rels: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const

const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'

export const REL = {
  officeDocument: `${OD}officeDocument`,
  extendedProps: `${OD}extended-properties`,
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  thumbnail: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
  signatureOrigin: 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin',
  slide: `${OD}slide`,
  slideLayout: `${OD}slideLayout`,
  slideMaster: `${OD}slideMaster`,
  notesSlide: `${OD}notesSlide`,
  notesMaster: `${OD}notesMaster`,
  theme: `${OD}theme`,
  presProps: `${OD}presProps`,
  viewProps: `${OD}viewProps`,
  tableStyles: `${OD}tableStyles`,
  image: `${OD}image`,
  hyperlink: `${OD}hyperlink`,
  styles: `${OD}styles`,
  settings: `${OD}settings`,
} as const

const PML = 'application/vnd.openxmlformats-officedocument.presentationml.'
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml.'

export const CT = {
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  pptx: `${PML}presentation.main+xml`,
  ppsx: `${PML}slideshow.main+xml`,
  potx: `${PML}template.main+xml`,
  slide: `${PML}slide+xml`,
  slideLayout: `${PML}slideLayout+xml`,
  slideMaster: `${PML}slideMaster+xml`,
  notesSlide: `${PML}notesSlide+xml`,
  notesMaster: `${PML}notesMaster+xml`,
  presProps: `${PML}presProps+xml`,
  viewProps: `${PML}viewProps+xml`,
  tableStyles: `${PML}tableStyles+xml`,
  docx: `${WML}document.main+xml`,
  styles: `${WML}styles+xml`,
  settings: `${WML}settings+xml`,
} as const

/** The _rels/*.rels part for a part ('' is the package itself). */
export function relsPathFor(part: string): string {
  const slash = part.lastIndexOf('/')
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`
}

/** One part's relationships, numbered rId1… in the order they're added. */
export class Rels {
  private list: { id: string; type: string; target: string; external: boolean }[] = []

  add(type: string, target: string, external = false): string {
    const same = this.list.find((r) => r.type === type && r.target === target && r.external === external)
    if (same) return same.id
    const id = `rId${this.list.length + 1}`
    this.list.push({ id, type, target, external })
    return id
  }

  get size(): number {
    return this.list.length
  }

  xml(): string {
    const items = this.list.map((r) =>
      `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`)
    return `${XML_DECL}<Relationships xmlns="${NSX.rels}">${items.join('')}</Relationships>`
  }
}

/** Collects the parts of a package built from scratch, then zips it with
 * [Content_Types].xml first, as Office writes it. */
export class PackageWriter {
  private entries: ZipEntry[] = []
  private defaults = new Map<string, string>([['rels', CT.rels], ['xml', 'application/xml']])
  private overrides = new Map<string, string>()
  private when = new Date()

  /** Add a part; an XML part with its own content type gets an Override. */
  add(path: string, data: Uint8Array | string, contentType?: string): void {
    if (contentType) this.overrides.set(`/${path}`, contentType)
    this.entries.push(makeEntry(path, data, this.when))
  }

  /** Add a part's relationships, when it has any. */
  rels(part: string, rels: Rels): void {
    if (rels.size) this.add(relsPathFor(part), rels.xml())
  }

  defaultType(ext: string, contentType: string): void {
    this.defaults.set(ext, contentType)
  }

  zip(): Uint8Array {
    const defs = [...this.defaults].map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`)
    const ovs = [...this.overrides].map(([p, t]) => `<Override PartName="${esc(p)}" ContentType="${t}"/>`)
    const ct = `${XML_DECL}<Types xmlns="${NSX.ct}">${defs.join('')}${ovs.join('')}</Types>`
    return writeZip([makeEntry('[Content_Types].xml', ct, this.when), ...this.entries])
  }
}

/** W3CDTF, which core.xml insists on (no milliseconds). */
export function w3cdtf(d = new Date()): string {
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

export function corePropsXml(p: { title: string; creator?: string }): string {
  const now = w3cdtf()
  return `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" `
    + 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    + 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    + `<dc:title>${esc(p.title)}</dc:title>`
    + (p.creator ? `<dc:creator>${esc(p.creator)}</dc:creator>` : '')
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`
    + `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
    + '</cp:coreProperties>'
}

/** docProps/app.xml. `fields` are extra elements in the order given (the
 * schema doesn't mind the order). */
export function appPropsXml(fields: Record<string, string | number> = {}): string {
  const body = Object.entries(fields).map(([k, v]) => `<${k}>${esc(String(v))}</${k}>`).join('')
  return `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" `
    + 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    + `<Application>Advanced Exporter</Application>${body}</Properties>`
}

/** EMU → centimetres, for ODF. */
export function cm(emu: number): string {
  return `${(emu / 360000).toFixed(3)}cm`
}
