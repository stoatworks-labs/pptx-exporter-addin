// GIF frames: quantise to a palette, optionally with Floyd–Steinberg error
// diffusion (gifenc maps each pixel to its nearest colour, which bands on the
// gradients slides love).

import { applyPalette, GIFEncoder, nearestColorIndex, quantize, type Palette } from 'gifenc'

export interface GifFrame {
  width: number
  height: number
  rgba: Uint8Array | Uint8ClampedArray
  delayMs?: number
}

export interface GifOptions {
  colors: number
  dither: boolean
  /** 0 = forever, -1 = play once, n = repeat n times. */
  repeat?: number
  /** One palette for the whole animation (smaller, steadier) or one per frame. */
  globalPalette?: boolean
}

export function paletteFor(frames: GifFrame[], colors: number): Palette {
  // Sample every frame into one buffer so the palette serves them all.
  const budget = 1_000_000
  const total = frames.reduce((n, f) => n + f.width * f.height, 0)
  const step = Math.max(1, Math.ceil(total / budget))
  const sample: number[] = []
  for (const f of frames) {
    for (let i = 0; i < f.width * f.height; i += step) {
      sample.push(f.rgba[i * 4], f.rgba[i * 4 + 1], f.rgba[i * 4 + 2], 255)
    }
  }
  return quantize(Uint8Array.from(sample), colors)
}

export function indexFrame(f: GifFrame, palette: Palette, dither: boolean): Uint8Array {
  if (!dither) return applyPalette(f.rgba, palette)
  const { width: w, height: h } = f
  const err = new Float32Array(w * h * 3)
  for (let i = 0; i < w * h; i++) {
    err[i * 3] = f.rgba[i * 4]
    err[i * 3 + 1] = f.rgba[i * 4 + 1]
    err[i * 3 + 2] = f.rgba[i * 4 + 2]
  }
  const out = new Uint8Array(w * h)
  const cache = new Map<number, number>()
  const clamp = (v: number) => (v < 0 ? 0 : v > 255 ? 255 : v)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x
      const r = clamp(err[i * 3]), g = clamp(err[i * 3 + 1]), b = clamp(err[i * 3 + 2])
      const key = ((r >> 2) << 12) | ((g >> 2) << 6) | (b >> 2)
      let idx = cache.get(key)
      if (idx == null) {
        idx = nearestColorIndex(palette, [r, g, b])
        cache.set(key, idx)
      }
      out[i] = idx
      const p = palette[idx]
      const er = r - p[0], eg = g - p[1], eb = b - p[2]
      const spread = (dx: number, dy: number, k: number) => {
        const xx = x + dx, yy = y + dy
        if (xx < 0 || xx >= w || yy >= h) return
        const j = (yy * w + xx) * 3
        err[j] += er * k
        err[j + 1] += eg * k
        err[j + 2] += eb * k
      }
      spread(1, 0, 7 / 16)
      spread(-1, 1, 3 / 16)
      spread(0, 1, 5 / 16)
      spread(1, 1, 1 / 16)
    }
  }
  return out
}

export function encodeGif(frames: GifFrame[], o: GifOptions): Uint8Array {
  const gif = GIFEncoder()
  const global = o.globalPalette !== false ? paletteFor(frames, o.colors) : undefined
  frames.forEach((f, i) => {
    const palette = global ?? quantize(f.rgba, o.colors)
    gif.writeFrame(indexFrame(f, palette, o.dither), f.width, f.height, {
      palette: i === 0 || !global ? palette : undefined,
      delay: f.delayMs ?? 0,
      repeat: i === 0 ? (o.repeat ?? 0) : undefined,
    })
  })
  gif.finish()
  return gif.bytes()
}
