// Turning "1080p, letterboxed" into pixels. Pure — unit tested.

import { EMU_PER_INCH, EMU_PER_PX96 } from '../deck/model'

export type SizePreset =
  | 'native' | '480p' | '720p' | '1080p' | '1440p' | '4k' | '5k' | '8k'
  | 'width' | 'height' | 'box' | 'dpi' | 'scale'

/** contain: the picture is the slide's shape, as big as fits the box.
 *  pad: exactly the box, slide centred, bars in the background colour.
 *  cover: exactly the box, slide fills it, edges cropped.
 *  stretch: exactly the box, slide distorted to fit. */
export type Fit = 'contain' | 'pad' | 'cover' | 'stretch'

export interface SizeOptions {
  size: SizePreset
  width?: number
  height?: number
  dpi?: number
  scale?: number
  fit?: Fit
}

export const PRESET_BOXES: Record<string, [number, number]> = {
  '480p': [854, 480],
  '720p': [1280, 720],
  '1080p': [1920, 1080],
  '1440p': [2560, 1440],
  '4k': [3840, 2160],
  '5k': [5120, 2880],
  '8k': [7680, 4320],
}

export const SIZE_CHOICES: { value: SizePreset; label: string }[] = [
  { value: 'native', label: "Slide's own size (96 dpi)" },
  { value: '720p', label: 'HD 1280 × 720' },
  { value: '1080p', label: 'Full HD 1920 × 1080' },
  { value: '1440p', label: 'QHD 2560 × 1440' },
  { value: '4k', label: '4K UHD 3840 × 2160' },
  { value: '5k', label: '5K 5120 × 2880' },
  { value: '8k', label: '8K 7680 × 4320' },
  { value: 'width', label: 'Width…' },
  { value: 'height', label: 'Height…' },
  { value: 'box', label: 'Width × height…' },
  { value: 'dpi', label: 'Print resolution (dpi)…' },
  { value: 'scale', label: 'Scale (% of slide size)…' },
]

/** Everything the renderer needs to build one output frame. */
export interface Frame {
  /** The output picture. */
  width: number
  height: number
  /** Where the slide lands inside it (may overhang for cover). */
  dx: number
  dy: number
  dw: number
  dh: number
  /** What to ask PowerPoint for: the slide's shape, at the drawn size. */
  renderW: number
  renderH: number
  /** Pixels per inch of slide, for DPI metadata. */
  dpi: number
}

/** The biggest picture the pane will build. PowerPoint for Mac's WKWebView
 * made, filled, read back and encoded a 16384 × 9216 canvas (measured
 * 2026-10-20); Chromium (WebView2) allows 16384 × 16384. */
export const MAX_PIXELS = 16384 * 9216
export const MAX_SIDE = 16384

export function frameFor(slideW: number, slideH: number, o: SizeOptions): Frame {
  const ar = slideW / slideH
  const nativeW = slideW / EMU_PER_PX96
  const nativeH = slideH / EMU_PER_PX96
  let boxW: number
  let boxH: number
  let fit: Fit = o.fit ?? 'contain'

  switch (o.size) {
    case 'native':
      boxW = nativeW
      boxH = nativeH
      fit = 'contain'
      break
    case 'width':
      boxW = o.width || 1920
      boxH = boxW / ar
      fit = 'contain'
      break
    case 'height':
      boxH = o.height || 1080
      boxW = boxH * ar
      fit = 'contain'
      break
    case 'box':
      boxW = o.width || 1920
      boxH = o.height || 1080
      break
    case 'dpi': {
      const dpi = o.dpi || 300
      boxW = (slideW / EMU_PER_INCH) * dpi
      boxH = (slideH / EMU_PER_INCH) * dpi
      fit = 'contain'
      break
    }
    case 'scale': {
      const s = (o.scale || 100) / 100
      boxW = nativeW * s
      boxH = nativeH * s
      fit = 'contain'
      break
    }
    default: {
      const box = PRESET_BOXES[o.size] ?? PRESET_BOXES['1080p']
      ;[boxW, boxH] = box
    }
  }

  boxW = Math.max(1, Math.round(boxW))
  boxH = Math.max(1, Math.round(boxH))

  let width: number, height: number, dw: number, dh: number
  if (fit === 'stretch') {
    width = dw = boxW
    height = dh = boxH
  } else if (fit === 'cover') {
    width = boxW
    height = boxH
    const s = Math.max(boxW / slideW, boxH / slideH)
    dw = Math.round(slideW * s)
    dh = Math.round(slideH * s)
  } else {
    const s = Math.min(boxW / slideW, boxH / slideH)
    dw = Math.max(1, Math.round(slideW * s))
    dh = Math.max(1, Math.round(slideH * s))
    if (fit === 'pad') {
      width = boxW
      height = boxH
    } else {
      width = dw
      height = dh
    }
  }
  const dx = Math.round((width - dw) / 2)
  const dy = Math.round((height - dh) / 2)
  // PowerPoint is asked for the slide at its drawn size, so nothing is
  // resampled after it has drawn it — except for stretch, which can't be
  // drawn any other way.
  const renderW = fit === 'stretch' ? Math.round(boxH * ar) <= boxW ? Math.round(boxH * ar) : boxW : dw
  const renderH = fit === 'stretch' ? Math.round(renderW / ar) : dh
  const dpi = (dw / slideW) * EMU_PER_INCH
  return { width, height, dx, dy, dw, dh, renderW, renderH, dpi }
}

export function frameProblem(f: Frame): string | null {
  if (f.width > MAX_SIDE || f.height > MAX_SIDE) return `${f.width} × ${f.height} is past the ${MAX_SIDE}-pixel side limit web views keep to.`
  if (f.width * f.height > MAX_PIXELS) return `${f.width} × ${f.height} is more pixels than the pane can build in one picture (16384 × 9216 is the most).`
  return null
}

export function describeFrame(f: Frame): string {
  return `${f.width} × ${f.height}`
}
