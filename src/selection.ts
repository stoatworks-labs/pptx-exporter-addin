// Which slides an export covers. Pure — unit tested.

import type { Deck, Slide } from './deck/model'

export type Scope = 'all' | 'selected' | 'range' | 'section'

export interface SelectionInput {
  scope: Scope
  /** "1-3, 7, 9-" — one-based, as the user sees slide numbers. */
  range: string
  section: string
  /** Zero-based indexes selected in PowerPoint's thumbnail pane. */
  selected: number[]
  includeHidden: boolean
}

export interface RangeResult {
  indexes: number[]
  error?: string
}

/** Parse "1-3, 7, 9-" against a deck of `count` slides. Open ends run to the
 * first or last slide; order is the deck's, duplicates dropped. */
export function parseRange(text: string, count: number): RangeResult {
  const picked = new Set<number>()
  const t = text.trim()
  if (!t) return { indexes: [], error: 'Type some slide numbers, e.g. 1-3, 7.' }
  for (const raw of t.split(/[,;\s]+/).filter(Boolean)) {
    const m = /^(\d*)\s*[-–—]\s*(\d*)$/.exec(raw) ?? /^(\d+)$/.exec(raw)
    if (!m) return { indexes: [], error: `“${raw}” isn't a slide number or a range like 3-5.` }
    let a: number, b: number
    if (m.length === 2) a = b = parseInt(m[1], 10)
    else {
      a = m[1] ? parseInt(m[1], 10) : 1
      b = m[2] ? parseInt(m[2], 10) : count
    }
    if (a > b) [a, b] = [b, a]
    if (a < 1 || b > count) return { indexes: [], error: `This deck has slides 1 to ${count}; “${raw}” is outside that.` }
    for (let i = a; i <= b; i++) picked.add(i - 1)
  }
  return { indexes: [...picked].sort((x, y) => x - y) }
}

export interface Selection {
  slides: Slide[]
  /** Why the selection is empty or odd, for the pane to say. */
  note?: string
  error?: string
}

export function selectSlides(deck: Deck, s: SelectionInput): Selection {
  let picked: Slide[]
  let error: string | undefined
  switch (s.scope) {
    case 'selected':
      picked = s.selected.map((i) => deck.slides[i]).filter(Boolean)
      if (!picked.length) error = 'No slides are selected in PowerPoint. Click one in the thumbnails.'
      break
    case 'range': {
      const r = parseRange(s.range, deck.slides.length)
      error = r.error
      picked = r.indexes.map((i) => deck.slides[i])
      break
    }
    case 'section': {
      const sec = deck.sections.find((x) => x.name === s.section)
      picked = sec ? sec.slideIndexes.map((i) => deck.slides[i]) : []
      if (!sec) error = 'Pick a section.'
      break
    }
    default:
      picked = deck.slides
  }
  // Hidden slides go only when asked for — except slides the user picked one
  // by one in PowerPoint, which they presumably meant.
  const hiddenDropped = s.includeHidden || s.scope === 'selected' ? 0 : picked.filter((x) => x.hidden).length
  if (!s.includeHidden && s.scope !== 'selected') picked = picked.filter((x) => !x.hidden)
  const note = hiddenDropped ? `${hiddenDropped} hidden slide${hiddenDropped > 1 ? 's' : ''} left out.` : undefined
  if (!error && !picked.length) error = 'That leaves no slides to export.'
  return { slides: picked, note, error }
}
