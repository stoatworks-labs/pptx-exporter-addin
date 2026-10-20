// A slide's content in reading order, for the formats that rebuild slides as
// text (LaTeX, Markdown). Pure — unit tested.

import type { Shape, Slide } from '../deck/model'
import { needsPicture } from './crop'

export type Block =
  | { kind: 'text'; shape: Shape }
  | { kind: 'table'; shape: Shape }
  | { kind: 'image'; shape: Shape }
  /** Shown as a picture of its region of the slide. */
  | { kind: 'crop'; shape: Shape }

export interface SlideStructure {
  title?: Shape
  subtitle?: Shape
  /** A title slide (ctrTitle placeholder). */
  isTitleSlide: boolean
  /** One list per column, left to right; usually just one. */
  columns: Block[][]
}

const FURNITURE = new Set(['dt', 'ftr', 'sldNum', 'hdr', 'sldImg'])

export function structure(slide: Slide): SlideStructure {
  const title = slide.shapes.find((s) => s.placeholder && (s.placeholder.type === 'title' || s.placeholder.type === 'ctrTitle'))
  const isTitleSlide = title?.placeholder?.type === 'ctrTitle'
  const subtitle = slide.shapes.find((s) => s.placeholder?.type === 'subTitle')
  const blocks: Block[] = []
  for (const s of slide.shapes) {
    if (s === title || s === subtitle) continue
    if (s.placeholder && FURNITURE.has(s.placeholder.type)) continue
    if (s.kind === 'text' && s.paragraphs) blocks.push({ kind: 'text', shape: s })
    else if (s.kind === 'table' && s.table) blocks.push({ kind: 'table', shape: s })
    else if (needsPicture(s)) blocks.push({ kind: 'crop', shape: s })
    else if (s.kind === 'picture' && s.image) blocks.push({ kind: 'image', shape: s })
    // Plain drawn shapes without words are decoration; leave them out.
  }
  return { title, subtitle, isTitleSlide, columns: toColumns(blocks) }
}

/** Side-by-side content becomes columns; everything else reads top to bottom. */
export function toColumns(blocks: Block[]): Block[][] {
  const placed = blocks.filter((b) => b.shape.rect)
  const loose = blocks.filter((b) => !b.shape.rect)
  if (placed.length < 2) return [sortByY(blocks)]
  const cols: { x0: number; x1: number; items: Block[] }[] = []
  for (const b of [...placed].sort((a, c) => a.shape.rect!.x - c.shape.rect!.x)) {
    const r = b.shape.rect!
    const col = cols.find((c) => overlap(c.x0, c.x1, r.x, r.x + r.w) > 0.3 * Math.min(c.x1 - c.x0, r.w))
    if (col) {
      col.items.push(b)
      col.x0 = Math.min(col.x0, r.x)
      col.x1 = Math.max(col.x1, r.x + r.w)
    } else cols.push({ x0: r.x, x1: r.x + r.w, items: [b] })
  }
  // Only call it columns if they really sit beside each other.
  if (cols.length < 2 || cols.length > 4 || !sideBySide(cols.map((c) => c.items))) return [sortByY(blocks)]
  const out = cols.map((c) => sortByY(c.items))
  if (loose.length) out[0].push(...loose)
  return out
}

function overlap(a0: number, a1: number, b0: number, b1: number): number {
  return Math.max(0, Math.min(a1, b1) - Math.max(a0, b0))
}

function sideBySide(cols: Block[][]): boolean {
  const span = (items: Block[]) => [
    Math.min(...items.map((b) => b.shape.rect!.y)),
    Math.max(...items.map((b) => b.shape.rect!.y + b.shape.rect!.h)),
  ]
  const spans = cols.map(span)
  for (let i = 1; i < spans.length; i++) {
    const [a0, a1] = spans[i - 1]
    const [b0, b1] = spans[i]
    if (overlap(a0, a1, b0, b1) < 0.25 * Math.min(a1 - a0, b1 - b0)) return false
  }
  return true
}

function sortByY(blocks: Block[]): Block[] {
  return [...blocks].sort((a, b) => (a.shape.rect?.y ?? 0) - (b.shape.rect?.y ?? 0) || (a.shape.rect?.x ?? 0) - (b.shape.rect?.x ?? 0))
}

/** Fraction of the slide's width a shape takes, for sizing pictures. */
export function widthFraction(s: Shape, slideWidth: number): number {
  if (!s.rect) return 0.8
  return Math.min(1, Math.max(0.1, s.rect.w / slideWidth))
}
