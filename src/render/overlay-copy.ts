// The slide copy an overlay is drawn from.
//
// PowerPoint always paints a slide's background when it draws a picture of
// it, so a transparent picture is pulled out of two drawings of the same
// slide, one on solid black and one on solid white (render/matte.ts). The
// slide itself is never changed for this: PowerPoint hands over a copy of it
// as a one-slide .pptx (Slide.exportAsBase64), this edits the copy, and the
// host inserts it, draws it and deletes it again.
//
// The edit: a solid background on the slide itself (which overrides whatever
// its layout and master say), the layout's and master's own graphics hidden
// unless asked for, and optionally a picture that fills the slide from the
// very back taken out — the way many decks fake a background.

import { all, kid, kids, parseXml } from '../lib/xml'
import { entryData, makeEntry, readZip, writeZip, XML_DECL } from '../formats/ooxml-write'

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

export interface OverlayCopyOptions {
  /** Keep logos, bars and other graphics that come from the layout or master. */
  masterGraphics: boolean
  /** Take out pictures that fill the slide from the very back. */
  dropBackPictures: boolean
}

export interface OverlayCopy {
  bytes: Uint8Array
  /** Pictures taken out from the back, over every slide in the package. */
  dropped: number
}

/** Every slide in the package on a solid `color` (#rrggbb), edited as asked. */
export function overlayCopy(pptx: Uint8Array, color: string, o: OverlayCopyOptions): OverlayCopy {
  const hex = color.replace(/^#/, '').toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(hex)) throw new Error(`Not a colour: ${color}`)
  const entries = readZip(pptx)
  const size = slideSize(entries.find((e) => e.name === 'ppt/presentation.xml'))
  let dropped = 0
  let slides = 0
  const out = entries.map((e) => {
    if (!/^ppt\/slides\/slide\d+\.xml$/i.test(e.name)) return e
    slides++
    const doc = parseXml(entryData(e))
    const sld = doc.documentElement
    const cSld = kid(sld, 'cSld')
    if (sld.localName !== 'sld' || !cSld) throw new Error(`${e.name} is not a slide.`)
    if (!o.masterGraphics) sld.setAttribute('showMasterSp', '0')
    for (const bg of kids(cSld, 'bg')) cSld.removeChild(bg)
    cSld.insertBefore(solidBackground(doc, sld, hex), cSld.firstChild)
    if (o.dropBackPictures && size) dropped += dropBack(kid(cSld, 'spTree'), size)
    const xml = XML_DECL + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '')
    return makeEntry(e.name, xml)
  })
  if (!slides) throw new Error('PowerPoint handed over a copy of the slide with no slide in it.')
  return { bytes: writeZip(out), dropped }
}

/** <p:bg><p:bgPr><a:solidFill><a:srgbClr val=…/></a:solidFill><a:effectLst/></p:bgPr></p:bg>,
 * in whatever prefixes this part already uses. */
function solidBackground(doc: Document, sld: Element, hex: string): Element {
  const p = sld.namespaceURI
  const pq = (local: string) => (sld.prefix ? `${sld.prefix}:${local}` : local)
  const ap = sld.lookupPrefix(A) ?? 'a'
  const aq = (local: string) => `${ap}:${local}`
  const bg = doc.createElementNS(p, pq('bg'))
  const bgPr = doc.createElementNS(p, pq('bgPr'))
  const fill = doc.createElementNS(A, aq('solidFill'))
  const clr = doc.createElementNS(A, aq('srgbClr'))
  clr.setAttribute('val', hex)
  fill.appendChild(clr)
  bgPr.appendChild(fill)
  bgPr.appendChild(doc.createElementNS(A, aq('effectLst')))
  bg.appendChild(bgPr)
  return bg
}

/** Pictures at the very back of the shape tree that cover the slide (bleed
 * allowed). Stops at the first thing that isn't one. */
function dropBack(tree: Element | null, size: { w: number; h: number }): number {
  if (!tree) return 0
  let n = 0
  for (let c = tree.firstElementChild; c; ) {
    const next = c.nextElementSibling
    if (c.localName === 'nvGrpSpPr' || c.localName === 'grpSpPr' || c.localName === 'extLst') {
      c = next
      continue
    }
    if (c.localName !== 'pic' || !coversSlide(c, size)) break
    tree.removeChild(c)
    n++
    c = next
  }
  return n
}

function coversSlide(pic: Element, size: { w: number; h: number }): boolean {
  const xfrm = kid(kid(pic, 'spPr'), 'xfrm')
  const off = kid(xfrm, 'off')
  const ext = kid(xfrm, 'ext')
  if (!off || !ext) return false
  const x = Number(off.getAttribute('x')), y = Number(off.getAttribute('y'))
  const w = Number(ext.getAttribute('cx')), h = Number(ext.getAttribute('cy'))
  if (![x, y, w, h].every(Number.isFinite)) return false
  // Within 1% of each edge, or past it.
  const tx = size.w * 0.01, ty = size.h * 0.01
  return x <= tx && y <= ty && x + w >= size.w - tx && y + h >= size.h - ty
}

function slideSize(e: ReturnType<typeof readZip>[number] | undefined): { w: number; h: number } | null {
  if (!e) return null
  const sz = all(parseXml(entryData(e)), 'sldSz')[0]
  const w = Number(sz?.getAttribute('cx')), h = Number(sz?.getAttribute('cy'))
  return w > 0 && h > 0 ? { w, h } : null
}
