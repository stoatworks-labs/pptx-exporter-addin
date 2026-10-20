// BMP and TGA writers, and DPI metadata for PNG and JPEG. Pure — unit tested.

/** 24-bit Windows bitmap, bottom-up, rows padded to four bytes. */
export function encodeBmp(width: number, height: number, rgba: Uint8Array | Uint8ClampedArray, dpi = 96): Uint8Array {
  const rowSize = (width * 3 + 3) & ~3
  const pixels = rowSize * height
  const out = new Uint8Array(54 + pixels)
  const dv = new DataView(out.buffer)
  out[0] = 0x42 // B
  out[1] = 0x4d // M
  dv.setUint32(2, out.length, true)
  dv.setUint32(10, 54, true)
  dv.setUint32(14, 40, true) // BITMAPINFOHEADER
  dv.setInt32(18, width, true)
  dv.setInt32(22, height, true) // positive: bottom-up
  dv.setUint16(26, 1, true)
  dv.setUint16(28, 24, true)
  dv.setUint32(30, 0, true) // BI_RGB
  dv.setUint32(34, pixels, true)
  const ppm = Math.round(dpi / 0.0254)
  dv.setInt32(38, ppm, true)
  dv.setInt32(42, ppm, true)
  for (let y = 0; y < height; y++) {
    let si = (height - 1 - y) * width * 4
    let di = 54 + y * rowSize
    for (let x = 0; x < width; x++, si += 4, di += 3) {
      out[di] = rgba[si + 2]
      out[di + 1] = rgba[si + 1]
      out[di + 2] = rgba[si]
    }
  }
  return out
}

/** Truevision TGA, 24-bit, top-left origin; optionally run-length encoded. */
export function encodeTga(width: number, height: number, rgba: Uint8Array | Uint8ClampedArray, rle = true): Uint8Array {
  const header = new Uint8Array(18)
  header[2] = rle ? 10 : 2
  header[12] = width & 0xff
  header[13] = width >> 8
  header[14] = height & 0xff
  header[15] = height >> 8
  header[16] = 24
  header[17] = 0x20 // top-left origin
  const footer = new TextEncoder().encode('\0\0\0\0\0\0\0\0TRUEVISION-XFILE.\0')
  const body: number[] | Uint8Array = rle ? tgaRle(width, height, rgba) : tgaRaw(width, height, rgba)
  const out = new Uint8Array(18 + body.length + footer.length)
  out.set(header)
  out.set(body as ArrayLike<number>, 18)
  out.set(footer, 18 + body.length)
  return out
}

function tgaRaw(width: number, height: number, rgba: Uint8Array | Uint8ClampedArray): Uint8Array {
  const out = new Uint8Array(width * height * 3)
  for (let i = 0, j = 0; i < width * height * 4; i += 4, j += 3) {
    out[j] = rgba[i + 2]
    out[j + 1] = rgba[i + 1]
    out[j + 2] = rgba[i]
  }
  return out
}

function tgaRle(width: number, height: number, rgba: Uint8Array | Uint8ClampedArray): Uint8Array {
  let out = new Uint8Array(width * height * 3 + Math.ceil((width * height) / 128) * 2 + 16)
  let o = 0
  const px = (i: number) => (rgba[i * 4] << 16) | (rgba[i * 4 + 1] << 8) | rgba[i * 4 + 2]
  const put = (v: number) => {
    out[o++] = v & 0xff
    out[o++] = (v >> 8) & 0xff
    out[o++] = (v >> 16) & 0xff
  }
  // Packets never cross a scanline (the spec's recommendation; some readers insist).
  for (let y = 0; y < height; y++) {
    const rowStart = y * width
    const rowEnd = rowStart + width
    let i = rowStart
    while (i < rowEnd) {
      let run = 1
      const v = px(i)
      while (i + run < rowEnd && run < 128 && px(i + run) === v) run++
      if (run > 1) {
        out[o++] = 0x80 | (run - 1)
        put(v)
        i += run
        continue
      }
      let lit = 1
      while (i + lit < rowEnd && lit < 128 && px(i + lit) !== px(i + lit - 1)) lit++
      // Leave the start of a run for the next packet.
      if (i + lit < rowEnd && lit > 1 && px(i + lit) === px(i + lit - 1)) lit--
      out[o++] = lit - 1
      for (let k = 0; k < lit; k++) put(px(i + k))
      i += lit
    }
    if (o > out.length - width * 4 - 8) {
      const grown = new Uint8Array(out.length * 2)
      grown.set(out)
      out = grown
    }
  }
  return out.slice(0, o)
}

// --- DPI metadata -----------------------------------------------------------

let crcTable: Uint32Array | undefined

export function crc32(bytes: Uint8Array, crc = 0xffffffff): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256)
    for (let n = 0; n < 256; n++) {
      let c = n
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      crcTable[n] = c >>> 0
    }
  }
  for (let i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8)
  return crc >>> 0
}

/** Set a PNG's physical size (pHYs), replacing any it has. */
export function pngWithDpi(png: Uint8Array, dpi: number): Uint8Array {
  const ppm = Math.round(dpi / 0.0254)
  const chunk = new Uint8Array(4 + 4 + 9 + 4)
  const dv = new DataView(chunk.buffer)
  dv.setUint32(0, 9)
  chunk.set([0x70, 0x48, 0x59, 0x73], 4) // pHYs
  dv.setUint32(8, ppm)
  dv.setUint32(12, ppm)
  chunk[16] = 1 // metre
  dv.setUint32(17, (crc32(chunk.subarray(4, 17)) ^ 0xffffffff) >>> 0)
  // Walk the chunks: drop an existing pHYs, insert ours after IHDR.
  const parts: Uint8Array[] = [png.subarray(0, 8)]
  const src = new DataView(png.buffer, png.byteOffset, png.byteLength)
  let p = 8
  while (p + 8 <= png.length) {
    const len = src.getUint32(p)
    const type = String.fromCharCode(png[p + 4], png[p + 5], png[p + 6], png[p + 7])
    const end = p + 12 + len
    if (type !== 'pHYs') parts.push(png.subarray(p, end))
    if (type === 'IHDR') parts.push(chunk)
    p = end
  }
  const out = new Uint8Array(parts.reduce((n, a) => n + a.length, 0))
  let o = 0
  for (const a of parts) {
    out.set(a, o)
    o += a.length
  }
  return out
}

/** Set a JPEG's JFIF density, adding a JFIF header if the encoder left it out. */
export function jpegWithDpi(jpg: Uint8Array, dpi: number): Uint8Array {
  const d = Math.max(1, Math.min(65535, Math.round(dpi)))
  if (jpg[0] !== 0xff || jpg[1] !== 0xd8) return jpg
  // An APP0 JFIF straight after SOI: patch it in place.
  if (jpg[2] === 0xff && jpg[3] === 0xe0 && String.fromCharCode(...jpg.subarray(6, 10)) === 'JFIF') {
    const out = jpg.slice()
    out[13] = 1 // dots per inch
    out[14] = d >> 8
    out[15] = d & 0xff
    out[16] = d >> 8
    out[17] = d & 0xff
    return out
  }
  const app0 = Uint8Array.from([
    0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x01,
    d >> 8, d & 0xff, d >> 8, d & 0xff, 0x00, 0x00,
  ])
  const out = new Uint8Array(jpg.length + app0.length)
  out.set(jpg.subarray(0, 2))
  out.set(app0, 2)
  out.set(jpg.subarray(2), 2 + app0.length)
  return out
}
