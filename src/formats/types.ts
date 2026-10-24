// The shape every export format takes.
//
// A format is a list of options (rendered generically by the pane, saved in
// presets) and a run() that writes files into the bundle. Formats never talk
// to Office.js directly — they ask the context for slide pictures, the .pptx
// or the PDF, so the same code runs against the mock host in a browser.

import type { Deck, Slide } from '../deck/model'
import type { Host } from '../host/types'
import type { SlideImages } from '../render/images'

export type Group = 'images' | 'documents' | 'web' | 'video' | 'text' | 'office'

export const GROUPS: { id: Group; label: string }[] = [
  { id: 'images', label: 'Images' },
  { id: 'documents', label: 'Documents' },
  { id: 'web', label: 'Web' },
  { id: 'video', label: 'Video & animation' },
  { id: 'text', label: 'LaTeX, Markdown & data' },
  { id: 'office', label: 'PowerPoint & Office' },
]

export type Values = Record<string, string | number | boolean>

interface Base {
  key: string
  label: string
  help?: string
  /** Only shown (and only meaningful) when this returns true. */
  when?: (v: Values) => boolean
}

export interface Choice {
  value: string
  label: string
}

export type OptionDef =
  | (Base & { type: 'select'; choices: Choice[]; default: string })
  | (Base & { type: 'number'; min: number; max: number; step?: number; unit?: string; default: number })
  | (Base & { type: 'slider'; min: number; max: number; step?: number; unit?: string; default: number })
  | (Base & { type: 'toggle'; default: boolean })
  | (Base & { type: 'color'; default: string })
  | (Base & { type: 'text'; default: string; placeholder?: string })

export interface Needs {
  /** Slide pictures (PowerPoint's renderer or the PDF fallback). */
  images?: boolean
  /** PowerPoint's PDF. */
  pdf?: boolean
  /** Slide.exportAsBase64. */
  exportSlide?: boolean
}

/**
 * What a format does with the deck's video and sound:
 * - `still`: pictures of slides (or none), so nothing plays;
 * - `kept`: the deck's own media goes with it (PowerPoint copies);
 * - `plays`: it plays embedded media itself and says what it couldn't (HTML5);
 * - `files`: it is the media files (Video bundle);
 * - `none`: text only, where media has nothing to lose.
 */
export type MediaHandling = 'still' | 'kept' | 'plays' | 'files' | 'none'

export interface FormatDef {
  id: string
  label: string
  group: Group
  /** One line for the format list. */
  blurb: string
  /** File extension(s) produced, for the list. */
  ext: string
  options: OptionDef[]
  needs: Needs
  /** When what's needed depends on the options (handouts drawn from the PDF or from pictures). */
  needsFor?: (v: Values) => Needs
  /** A reason this format can't run in this web view, or null. Checked once. */
  unavailable?: () => Promise<string | null>
  /** What becomes of the deck's video and sound. The runner turns it into one note (formats/media.ts). */
  media: MediaHandling
  run(ctx: ExportContext, v: Values): Promise<void>
}

export interface ExportContext {
  deck: Deck
  /** The slides to export, in deck order. */
  slides: Slide[]
  /** Base for file names: the document's name, tidied. */
  base: string
  host: Host
  images: SlideImages
  /** Write a file. Paths are relative to this format's folder. */
  add(path: string, data: Uint8Array | string): void
  /** PowerPoint's PDF (fetched once per export, shared). */
  pdf(): Promise<Uint8Array>
  /** The .pptx the deck was read from. */
  pptx: Uint8Array
  /** 0..1 within this format. */
  progress(fraction: number, label?: string): void
  warn(message: string): void
  signal: AbortSignal
}

export function defaults(def: FormatDef): Values {
  const v: Values = {}
  for (const o of def.options) v[o.key] = o.default
  return v
}

/** Fill gaps from defaults and drop values that no longer fit (a preset from an older version). */
export function normalise(def: FormatDef, given: Values | undefined): Values {
  const v = defaults(def)
  if (!given) return v
  for (const o of def.options) {
    const g = given[o.key]
    if (g == null) continue
    switch (o.type) {
      case 'select':
        if (typeof g === 'string' && o.choices.some((c) => c.value === g)) v[o.key] = g
        break
      case 'number':
      case 'slider':
        if (typeof g === 'number' && Number.isFinite(g)) v[o.key] = Math.min(o.max, Math.max(o.min, g))
        break
      case 'toggle':
        if (typeof g === 'boolean') v[o.key] = g
        break
      case 'color':
        if (typeof g === 'string' && /^#[0-9a-f]{6}$/i.test(g)) v[o.key] = g
        break
      case 'text':
        if (typeof g === 'string') v[o.key] = g
        break
    }
  }
  return v
}

export class Cancelled extends Error {
  constructor() {
    super('Export cancelled.')
  }
}

export function checkCancel(signal: AbortSignal): void {
  if (signal.aborted) throw new Cancelled()
}
