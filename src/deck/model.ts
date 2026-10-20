// What the exporter knows about a deck, read out of the .pptx package.
//
// PowerPoint draws the pixels (Slide.getImageAsBase64); this model supplies
// everything a picture cannot: the words, the speaker notes, which slides are
// hidden, transitions and timings, where the links and videos sit. All
// geometry is in EMU (914,400 per inch), relative to the slide's top-left.

export const EMU_PER_INCH = 914400
export const EMU_PER_POINT = 12700
/** PowerPoint's own "true size" for a slide image is the slide at 96 dpi. */
export const EMU_PER_PX96 = 9525

export interface Rect {
  x: number
  y: number
  w: number
  h: number
}

export interface Link {
  /** External address (http, mailto, file…). */
  url?: string
  /** Zero-based index of a slide in this deck. */
  slide?: number
  /** first | last | next | previous | end — PowerPoint's show-jump actions. */
  jump?: 'first' | 'last' | 'next' | 'previous' | 'end'
  tooltip?: string
}

export interface Run {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  /** Positive = superscript, negative = subscript. */
  baseline?: number
  /** A monospaced typeface was asked for — rendered as code where that means something. */
  mono?: boolean
  /** Point size, where the run says (otherwise it comes from the layout). */
  size?: number
  /** Hex colour where the run sets one directly (#rrggbb). */
  color?: string
  link?: Link
  /** A soft line break (<a:br/>), not text. */
  br?: boolean
}

export interface Paragraph {
  level: number
  /** none, a bullet character, or an automatic number scheme. */
  bullet: 'none' | 'char' | 'number'
  /** For numbered paragraphs, the scheme (arabicPeriod, romanUcPeriod…). */
  numberScheme?: string
  align?: 'l' | 'ctr' | 'r' | 'just'
  runs: Run[]
}

export type PlaceholderType =
  | 'title' | 'ctrTitle' | 'subTitle' | 'body' | 'obj' | 'dt' | 'ftr' | 'sldNum'
  | 'pic' | 'tbl' | 'chart' | 'dgm' | 'media' | 'clipArt' | 'hdr' | 'sldImg' | string

export interface TableCell {
  paragraphs: Paragraph[]
  colSpan: number
  rowSpan: number
  /** Covered by a neighbour's span — emit nothing for it. */
  merged: boolean
}

export interface MediaRef {
  kind: 'video' | 'audio'
  /** Package part name, e.g. ppt/media/media1.mp4 — absent for linked media. */
  part?: string
  /** A linked (not embedded) file's address. */
  external?: string
  contentType?: string
  /** The poster frame PowerPoint shows when it isn't playing. */
  poster?: string
}

export interface ImageRef {
  part: string
  contentType: string
}

export type ShapeKind = 'text' | 'picture' | 'table' | 'media' | 'chart' | 'diagram' | 'shape' | 'other'

export interface Shape {
  id: number
  name: string
  kind: ShapeKind
  rect?: Rect
  /** Clockwise, in degrees. */
  rotation?: number
  placeholder?: { type: PlaceholderType; idx?: number }
  paragraphs?: Paragraph[]
  table?: { rows: TableCell[][]; colWidths: number[] }
  image?: ImageRef
  media?: MediaRef
  /** Click action on the whole shape. */
  link?: Link
  /** Alt text. */
  description?: string
}

export interface Transition {
  /** The DrawingML element name: fade, push, wipe, cut, morph, prism… */
  type: string
  durationMs: number
  /** Direction attribute where the effect has one (l, r, u, d…). */
  dir?: string
  /** Whether a click advances (advClick, default true). */
  advanceOnClick: boolean
  /** Rehearsed / set timing: advance after this long. */
  advanceAfterMs?: number
}

export interface Slide {
  /** Zero-based position in the deck, hidden slides included. */
  index: number
  /** p:sldId@id — stable for the life of the slide. */
  sldId: number
  part: string
  hidden: boolean
  layoutName?: string
  title: string
  shapes: Shape[]
  notes: Paragraph[]
  notesText: string
  transition?: Transition
  hasAnimations: boolean
  section?: string
}

export interface Section {
  name: string
  slideIndexes: number[]
}

export interface Deck {
  title: string
  author?: string
  /** The language most of the text is tagged with (BCP 47, e.g. en-GB), if any. */
  lang?: string
  /** Slide size in EMU. */
  width: number
  height: number
  slides: Slide[]
  sections: Section[]
  /** presProps: loop until Esc, use timings. */
  show: { loop: boolean; useTimings: boolean }
  /** Read a part's bytes (media, images) out of the package on demand. */
  readPart(part: string): Uint8Array | undefined
  contentTypeOf(part: string): string | undefined
}

export function aspect(deck: Pick<Deck, 'width' | 'height'>): number {
  return deck.width / deck.height
}

/** Plain text of some paragraphs, one line per paragraph, soft breaks kept. */
export function paragraphsText(ps: Paragraph[]): string {
  return ps.map((p) => p.runs.map((r) => (r.br ? '\n' : r.text)).join('')).join('\n')
}

/** Everything readable on a slide, title first, in reading order. */
export function slideText(s: Slide): string {
  const parts: string[] = []
  for (const sh of s.shapes) {
    if (sh.paragraphs) {
      const t = paragraphsText(sh.paragraphs).trim()
      if (t) parts.push(t)
    } else if (sh.table) {
      for (const row of sh.table.rows) {
        const cells = row.filter((c) => !c.merged).map((c) => paragraphsText(c.paragraphs).trim())
        if (cells.some(Boolean)) parts.push(cells.join('\t'))
      }
    } else if (sh.description) {
      parts.push(sh.description)
    }
  }
  return parts.join('\n')
}

/** Ratio as a label: 16:9, 4:3, 16:10 — or the decimal when it's none of those. */
export function aspectLabel(w: number, h: number): string {
  const r = w / h
  const known: [number, string][] = [
    [16 / 9, '16:9'], [4 / 3, '4:3'], [16 / 10, '16:10'], [3 / 2, '3:2'], [5 / 4, '5:4'],
    [1, '1:1'], [21 / 9, '21:9'], [9 / 16, '9:16'], [210 / 297, 'A4 portrait'], [297 / 210, 'A4'],
  ]
  for (const [v, label] of known) if (Math.abs(r - v) < 0.01) return label
  return `${r.toFixed(2)}:1`
}
