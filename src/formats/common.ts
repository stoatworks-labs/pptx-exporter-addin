// Options and helpers several formats share.

import type { Slide } from '../deck/model'
import type { OverlayCopyOptions } from '../render/overlay-copy'
import { frameFor, frameProblem, SIZE_CHOICES, type Fit, type Frame, type SizePreset } from '../render/sizing'
import { checkCancel, type ExportContext, type OptionDef, type Values } from './types'

export interface SizeDefaults {
  size?: SizePreset
  fit?: Fit
  width?: number
  height?: number
  dpi?: number
  /** Leave out the fit control (formats whose output is always the slide's shape). */
  noFit?: boolean
  /** Leave out the print-resolution choice. */
  noDpi?: boolean
}

export function sizeOptions(d: SizeDefaults = {}): OptionDef[] {
  const choices = SIZE_CHOICES.filter((c) => !(d.noDpi && c.value === 'dpi'))
  const opts: OptionDef[] = [
    { key: 'size', label: 'Size', type: 'select', choices, default: d.size ?? '1080p' },
    {
      key: 'width', label: 'Width', type: 'number', min: 16, max: 16384, step: 1, unit: 'px', default: d.width ?? 1920,
      when: (v) => v.size === 'width' || v.size === 'box',
    },
    {
      key: 'height', label: 'Height', type: 'number', min: 16, max: 16384, step: 1, unit: 'px', default: d.height ?? 1080,
      when: (v) => v.size === 'height' || v.size === 'box',
    },
    {
      key: 'dpi', label: 'Resolution', type: 'number', min: 36, max: 1200, step: 1, unit: 'dpi', default: d.dpi ?? 300,
      when: (v) => v.size === 'dpi',
    },
    {
      key: 'scale', label: 'Scale', type: 'number', min: 10, max: 800, step: 5, unit: '%', default: 200,
      when: (v) => v.size === 'scale',
    },
  ]
  if (!d.noFit) {
    opts.push(
      {
        key: 'fit', label: 'Fit', type: 'select', default: d.fit ?? 'contain',
        choices: [
          { value: 'contain', label: "Keep the slide's shape" },
          { value: 'pad', label: 'Exact size, letterboxed' },
          { value: 'cover', label: 'Exact size, cropped to fill' },
          { value: 'stretch', label: 'Exact size, stretched' },
        ],
        help: "Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say.",
        when: (v) => v.size !== 'native' && v.size !== 'width' && v.size !== 'height' && v.size !== 'dpi' && v.size !== 'scale',
      },
      {
        key: 'background', label: 'Bars', type: 'color', default: '#000000',
        when: (v) => v.fit === 'pad' && v.size !== 'native' && v.size !== 'width' && v.size !== 'height' && v.size !== 'dpi' && v.size !== 'scale',
      },
    )
  }
  return opts
}

export function frameFromValues(ctx: ExportContext, v: Values): Frame {
  const f = frameFor(ctx.deck.width, ctx.deck.height, {
    size: v.size as SizePreset,
    width: v.width as number,
    height: v.height as number,
    dpi: v.dpi as number,
    scale: v.scale as number,
    fit: (v.fit as Fit) ?? 'contain',
  })
  const problem = frameProblem(f)
  if (problem) throw new Error(problem)
  return f
}

export function background(v: Values): string {
  return typeof v.background === 'string' ? v.background : '#000000'
}

// --- Transparent backgrounds (overlays) ---------------------------------------

const on = (v: Values) => !!v.transparent

/** The options a format that can carry alpha offers for leaving the slide's
 * background out. `fillKey` adds the separate fill-and-key choice. */
export function overlayOptions(fillKey: boolean): OptionDef[] {
  const opts: OptionDef[] = [
    {
      key: 'transparent', label: 'Transparent background', type: 'toggle', default: false,
      help: "Only what's on the slide, with an alpha channel: for overlays on video and in media servers. PowerPoint draws a copy of each slide on black and on white at the end of the deck, then deletes it; your slides aren't changed, but the copies are in PowerPoint's Undo history.",
    },
    {
      key: 'masterGraphics', label: 'Keep layout and master graphics', type: 'toggle', default: false, when: on,
      help: 'Logos, bars and other pictures that come from the slide layout or master rather than the slide itself.',
    },
    {
      key: 'backPicture', label: 'Drop a full-slide picture at the back', type: 'toggle', default: false, when: on,
      help: 'For decks whose background is a picture placed on the slide rather than set as the background.',
    },
  ]
  if (fillKey) {
    opts.push({
      key: 'alphaAs', label: 'Alpha', type: 'select', default: 'file', when: on,
      choices: [
        { value: 'file', label: 'In the file' },
        { value: 'pair', label: 'Separate fill and key files' },
        { value: 'both', label: 'Both' },
      ],
      help: 'Fill and key for vision mixers and CG kit that take them as two signals: the fill is the slide on black (pre-multiplied), the key is white where it is solid.',
    })
  }
  return opts
}

/** What to leave out, or null when the background stays. */
export function overlayOf(v: Values): OverlayCopyOptions | null {
  return v.transparent ? { masterGraphics: !!v.masterGraphics, dropBackPictures: !!v.backPicture } : null
}

/** Which files a transparent export writes: the picture with alpha, the fill
 * and key pair, or both. */
export function alphaFiles(v: Values): { alpha: boolean; pair: boolean } {
  const as = v.transparent ? String(v.alphaAs ?? 'file') : 'file'
  return { alpha: as !== 'pair', pair: !!v.transparent && as !== 'file' }
}

/** Size options for a format offering transparency: the letterbox colour
 * means nothing when the bars are transparent too. */
export function withOverlay(opts: OptionDef[], fillKey: boolean): OptionDef[] {
  const size = opts.map((o) => (o.key === 'background' && o.when ? { ...o, when: (v: Values) => !v.transparent && o.when!(v) } : o))
  return [...size, ...overlayOptions(fillKey)]
}

export const NAME_HELP = '{deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export'

export function nameOption(def: string): OptionDef {
  return { key: 'name', label: 'File names', type: 'text', default: def, help: NAME_HELP }
}

/** Expand a file-name pattern for one slide. */
export function expandName(pattern: string, ctx: Pick<ExportContext, 'base' | 'slides'>, slide: Slide, position: number): string {
  const n = slide.index + 1
  const width = String(ctx.slides.length ? Math.max(...ctx.slides.map((s) => s.index + 1)) : n).length
  const out = (pattern || '{deck}-{nn}')
    .replace(/\{deck\}/g, ctx.base)
    .replace(/\{nnn\}/g, String(n).padStart(3, '0'))
    .replace(/\{nn\}/g, String(n).padStart(Math.max(2, width), '0'))
    .replace(/\{n\}/g, String(n))
    .replace(/\{i\}/g, String(position + 1))
    .replace(/\{title\}/g, slug(slide.title) || `slide-${n}`)
  return safeName(out)
}

/** Lower-case, dash-separated, ASCII-ish — for file names from titles. */
export function slug(s: string): string {
  return s
    .normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^\p{L}\p{N}]+/gu, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase()
    .slice(0, 60)
}

/** Strip characters Windows, macOS or zip tools object to. */
export function safeName(s: string): string {
  return s.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-').replace(/^\.+/, '').trim() || 'untitled'
}

/** Run over the export's slides with progress, a cancel check, and a little
 * breathing room for the pane to repaint. */
export async function eachSlide(
  ctx: ExportContext,
  fn: (slide: Slide, position: number) => Promise<void>,
  label = 'Slide',
): Promise<void> {
  const total = ctx.slides.length
  for (let i = 0; i < total; i++) {
    checkCancel(ctx.signal)
    const s = ctx.slides[i]
    ctx.progress(i / total, `${label} ${s.index + 1}`)
    await fn(s, i)
    await new Promise((r) => setTimeout(r, 0))
  }
  ctx.progress(1)
}

/** Seconds each slide stays up when turning a deck into time: its rehearsed
 * timing if it has one and the user wants those, otherwise the default. */
export function slideSeconds(slide: Slide, v: Values): number {
  const fallback = Number(v.seconds) || 5
  if (v.timings && slide.transition?.advanceAfterMs) return Math.max(0.5, slide.transition.advanceAfterMs / 1000)
  return fallback
}
