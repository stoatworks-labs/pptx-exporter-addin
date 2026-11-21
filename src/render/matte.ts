// Transparency from two opaque pictures. Pure — unit tested.
//
// PowerPoint only draws opaque slide pictures, so an overlay is pulled from
// two drawings of the same slide, one on solid black and one on solid white.
// Each pixel PowerPoint drew is the content composited over the background:
//
//   onBlack = α·F            onWhite = α·F + (1 − α)·255
//
// so onWhite − onBlack = (1 − α)·255, which gives α, and onBlack / α gives the
// content's own colour F. Where the content is solid the two agree (α = 1);
// where there is none they are 0 and 255 (α = 0); soft edges, shadows, glows
// and see-through fills land in between, at the alpha they really have.
// Unlike keying out a colour, white or black content survives.

/**
 * Straight (un-premultiplied) RGBA from the same picture on black and on
 * white, both RGBA of the same size. The three channels' estimates of alpha
 * are averaged: they agree for greyscale anti-aliasing and differ only where
 * a renderer smooths text per sub-pixel.
 */
export function pullMatte(
  onBlack: ArrayLike<number>,
  onWhite: ArrayLike<number>,
  out: Uint8ClampedArray = new Uint8ClampedArray(onBlack.length),
): Uint8ClampedArray {
  if (onBlack.length !== onWhite.length) throw new Error('The two pictures are not the same size.')
  for (let i = 0; i < onBlack.length; i += 4) {
    const br = onBlack[i], bg = onBlack[i + 1], bb = onBlack[i + 2]
    const d = onWhite[i] - br + (onWhite[i + 1] - bg) + (onWhite[i + 2] - bb)
    const a = Math.min(255, Math.max(0, Math.round(255 - d / 3)))
    if (a === 0) {
      out[i] = out[i + 1] = out[i + 2] = out[i + 3] = 0
      continue
    }
    const k = 255 / a
    out[i] = br * k
    out[i + 1] = bg * k
    out[i + 2] = bb * k
    out[i + 3] = a
  }
  return out
}

/**
 * The two pictures vision mixers and CG kit take as separate signals, both
 * opaque RGBA: the fill (the content composited on black, i.e. pre-multiplied)
 * and the key (white where the content is solid, black where there is none).
 */
export function fillAndKey(rgba: ArrayLike<number>): { fill: Uint8ClampedArray; key: Uint8ClampedArray } {
  const fill = new Uint8ClampedArray(rgba.length)
  const key = new Uint8ClampedArray(rgba.length)
  for (let i = 0; i < rgba.length; i += 4) {
    const a = rgba[i + 3]
    fill[i] = (rgba[i] * a) / 255
    fill[i + 1] = (rgba[i + 1] * a) / 255
    fill[i + 2] = (rgba[i + 2] * a) / 255
    fill[i + 3] = 255
    key[i] = key[i + 1] = key[i + 2] = a
    key[i + 3] = 255
  }
  return { fill, key }
}

/** Whether any pixel is less than fully opaque. */
export function hasTransparency(rgba: ArrayLike<number>): boolean {
  for (let i = 3; i < rgba.length; i += 4) if (rgba[i] < 255) return true
  return false
}
