// An OpenDocument presentation of slide pictures, for LibreOffice, Keynote's
// importer and the public-sector systems that insist on ODF.
//
// ODF's one packaging rule that bites: the first zip entry must be
// "mimetype", stored uncompressed with no extra field, so the file can be
// recognised from its first bytes. ODF 1.2, which every current reader takes.

import { slideText, type Paragraph, type Run } from '../deck/model'
import type { Deck } from '../deck/model'
import type { PictureSlide } from './office-flat'
import { cm, esc, makeEntry, w3cdtf, writeZip, XML_DECL } from './ooxml-write'

export interface OdpOptions {
  notes: boolean
}

export const ODP_MIME = 'application/vnd.oasis.opendocument.presentation'

const NS = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
  dc: 'http://purl.org/dc/elements/1.1/',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
}
const DOC_NS = ['office', 'style', 'text', 'draw', 'fo', 'xlink', 'svg', 'presentation']
  .map((k) => `xmlns:${k}="${NS[k as keyof typeof NS]}"`).join(' ')

/** Notes pages are A4 portrait; the slide sits in the top half. */
const NOTES = { w: 7560000, h: 10692000, margin: 720000 }

export function buildOdp(deck: Deck, pics: PictureSlide[], o: OdpOptions): Uint8Array {
  const when = new Date()
  const pictures = pics.map((p) => `Pictures/slide-${String(p.slide.index + 1).padStart(3, '0')}.${p.type === 'png' ? 'png' : 'jpg'}`)
  const entries = [
    makeEntry('mimetype', ODP_MIME, when),
    makeEntry('content.xml', contentXml(deck, pics, pictures, o), when),
    makeEntry('styles.xml', stylesXml(deck), when),
    makeEntry('meta.xml', metaXml(deck, pics.length, when), when),
    ...pics.map((p, i) => makeEntry(pictures[i], p.bytes, when)),
    makeEntry('META-INF/manifest.xml', manifestXml(pics, pictures), when),
  ]
  return writeZip(entries)
}

function manifestXml(pics: PictureSlide[], pictures: string[]): string {
  const file = (path: string, type: string) => `<manifest:file-entry manifest:full-path="${esc(path)}" manifest:media-type="${type}"/>`
  return `${XML_DECL}<manifest:manifest xmlns:manifest="${NS.manifest}" manifest:version="1.2">`
    + `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODP_MIME}"/>`
    + file('content.xml', 'text/xml') + file('styles.xml', 'text/xml') + file('meta.xml', 'text/xml')
    + pics.map((p, i) => file(pictures[i], p.type === 'png' ? 'image/png' : 'image/jpeg')).join('')
    + '</manifest:manifest>'
}

function metaXml(deck: Deck, pages: number, when: Date): string {
  return `${XML_DECL}<office:document-meta xmlns:office="${NS.office}" xmlns:meta="${NS.meta}" xmlns:dc="${NS.dc}" office:version="1.2">`
    + '<office:meta><meta:generator>Advanced Exporter</meta:generator>'
    + `<dc:title>${esc(deck.title)}</dc:title>`
    + (deck.author ? `<meta:initial-creator>${esc(deck.author)}</meta:initial-creator>` : '')
    + `<meta:creation-date>${w3cdtf(when)}</meta:creation-date><dc:date>${w3cdtf(when)}</dc:date>`
    + `<meta:document-statistic meta:page-count="${pages}"/>`
    + '</office:meta></office:document-meta>'
}

/** The slide's page size, a notes page, and the one master every slide uses. */
function stylesXml(deck: Deck): string {
  const thumb = notesThumb(deck)
  const body = notesBody(thumb)
  const layout = (name: string, w: number, h: number) => `<style:page-layout style:name="${name}"><style:page-layout-properties `
    + `fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${cm(w)}" fo:page-height="${cm(h)}" `
    + `style:print-orientation="${w > h ? 'landscape' : 'portrait'}"/></style:page-layout>`
  return `${XML_DECL}<office:document-styles ${DOC_NS} office:version="1.2">`
    + '<office:styles/>'
    + '<office:automatic-styles>'
    + layout('PM1', deck.width, deck.height)
    + layout('PM2', NOTES.w, NOTES.h)
    + '<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#000000"/></style:style>'
    + '</office:automatic-styles>'
    + '<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1">'
    + '<presentation:notes style:page-layout-name="PM2">'
    + `<draw:page-thumbnail presentation:class="page" ${geometry(thumb)}/>`
    + `<draw:frame presentation:class="notes" presentation:placeholder="true" ${geometry(body)}><draw:text-box/></draw:frame>`
    + '</presentation:notes></style:master-page></office:master-styles>'
    + '</office:document-styles>'
}

interface Box {
  x: number
  y: number
  w: number
  h: number
}

const geometry = (b: Box) => `svg:x="${cm(b.x)}" svg:y="${cm(b.y)}" svg:width="${cm(b.w)}" svg:height="${cm(b.h)}"`

function notesThumb(deck: Deck): Box {
  const maxW = NOTES.w - 2 * NOTES.margin
  const maxH = NOTES.h * 0.4
  const s = Math.min(maxW / deck.width, maxH / deck.height)
  const w = deck.width * s
  return { x: (NOTES.w - w) / 2, y: NOTES.margin, w, h: deck.height * s }
}

function notesBody(thumb: Box): Box {
  const y = thumb.y + thumb.h + NOTES.margin / 2
  return { x: NOTES.margin, y, w: NOTES.w - 2 * NOTES.margin, h: NOTES.h - y - NOTES.margin }
}

function contentXml(deck: Deck, pics: PictureSlide[], pictures: string[], o: OdpOptions): string {
  const spans = new SpanStyles()
  const full: Box = { x: 0, y: 0, w: deck.width, h: deck.height }
  const thumb = notesThumb(deck)
  const body = notesBody(thumb)
  const pages = pics.map((p, i) => {
    const s = p.slide
    const title = s.title || `Slide ${s.index + 1}`
    const desc = slideText(s) || title
    const notes = o.notes && s.notes.length
      ? '<presentation:notes>'
        + `<draw:page-thumbnail presentation:class="page" draw:page-number="${i + 1}" ${geometry(thumb)}/>`
        + `<draw:frame presentation:class="notes" ${geometry(body)}><draw:text-box>${odfParagraphs(s.notes, spans)}</draw:text-box></draw:frame>`
        + '</presentation:notes>'
      : ''
    // draw:frame: its content first, then svg:title and svg:desc.
    return `<draw:page draw:name="Slide ${s.index + 1}" draw:style-name="${s.hidden ? 'dp2' : 'dp1'}" draw:master-page-name="Default">`
      + `<draw:frame draw:name="${esc(title)}" ${geometry(full)}>`
      + `<draw:image xlink:href="${esc(pictures[i])}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`
      + `<svg:title>${esc(title)}</svg:title><svg:desc>${esc(desc)}</svg:desc></draw:frame>`
      + notes
      + '</draw:page>'
  })
  return `${XML_DECL}<office:document-content ${DOC_NS} office:version="1.2">`
    + '<office:automatic-styles>'
    + '<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:visibility="visible"/></style:style>'
    + '<style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:visibility="hidden"/></style:style>'
    + spans.xml()
    + '</office:automatic-styles>'
    + `<office:body><office:presentation>${pages.join('')}</office:presentation></office:body>`
    + '</office:document-content>'
}

/** Automatic text styles, one per combination of bold/italic/underline/strike in use. */
class SpanStyles {
  private names = new Map<string, string>()

  name(r: Run): string | null {
    const key = `${r.bold ? 'b' : ''}${r.italic ? 'i' : ''}${r.underline ? 'u' : ''}${r.strike ? 's' : ''}`
    if (!key) return null
    let n = this.names.get(key)
    if (!n) {
      n = `T${this.names.size + 1}`
      this.names.set(key, n)
    }
    return n
  }

  xml(): string {
    return [...this.names].map(([key, name]) => {
      const p: string[] = []
      if (key.includes('b')) p.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"')
      if (key.includes('i')) p.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"')
      if (key.includes('u')) p.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"')
      if (key.includes('s')) p.push('style:text-line-through-style="solid"')
      return `<style:style style:name="${name}" style:family="text"><style:text-properties ${p.join(' ')}/></style:style>`
    }).join('')
  }
}

function odfParagraphs(ps: Paragraph[], spans: SpanStyles): string {
  return ps.map((p) => {
    const content = p.runs.map((r) => {
      if (r.br) return '<text:line-break/>'
      const t = odfText(r.text)
      const style = spans.name(r)
      return style ? `<text:span text:style-name="${style}">${t}</text:span>` : t
    }).join('')
    return `<text:p>${content}</text:p>`
  }).join('')
}

/** ODF collapses white space in text, so tabs and runs of spaces are spelt out. */
function odfText(s: string): string {
  return s.split('\t').map((seg) =>
    esc(seg).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`).replace(/^ /, '<text:s/>'),
  ).join('<text:tab/>')
}
