import { writeFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import * as UTIF from 'utif2'
import { decode as decodePng, encode as encodePng } from 'fast-png'
import { encodeTiff, lzwEncode, packBitsRows, type TiffCompression } from '../src/encode/tiff'
import { encodeBmp, encodeTga, jpegWithDpi, pngWithDpi, crc32 } from '../src/encode/simple'
import { encodeGif } from '../src/encode/gif'

/** A test card: gradients, flat areas (runs) and noise (literals). */
function card(w: number, h: number, seed = 1): Uint8Array {
  const px = new Uint8Array(w * h * 4)
  let s = seed
  const rnd = () => ((s = (s * 1103515245 + 12345) & 0x7fffffff) & 0xff)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4
      if (y < h / 3) { px[i] = (x * 255 / w) | 0; px[i + 1] = (y * 255 / h) | 0; px[i + 2] = 128 }
      else if (y < (2 * h) / 3) { px[i] = 200; px[i + 1] = 30; px[i + 2] = 60 }
      else { px[i] = rnd(); px[i + 1] = rnd(); px[i + 2] = rnd() }
      px[i + 3] = 255
    }
  }
  return px
}

function rgbOf(rgba: ArrayLike<number>): number[] {
  const out: number[] = []
  for (let i = 0; i < rgba.length; i += 4) out.push(rgba[i], rgba[i + 1], rgba[i + 2])
  return out
}

describe('TIFF', () => {
  for (const c of ['none', 'lzw', 'deflate', 'packbits'] as TiffCompression[]) {
    it(`round-trips pixels exactly (${c}), two pages`, () => {
      const a = card(317, 211, 1)
      const b = card(64, 48, 7)
      const tif = encodeTiff([{ width: 317, height: 211, rgba: a, dpi: 300 }, { width: 64, height: 48, rgba: b }], c)
      const ifds = UTIF.decode((tif.buffer as ArrayBuffer).slice(tif.byteOffset, tif.byteOffset + tif.byteLength) as ArrayBuffer)
      expect(ifds.length).toBe(2)
      UTIF.decodeImage((tif.buffer as ArrayBuffer), ifds[0])
      UTIF.decodeImage((tif.buffer as ArrayBuffer), ifds[1])
      expect(rgbOf(UTIF.toRGBA8(ifds[0]))).toEqual(rgbOf(a))
      expect(rgbOf(UTIF.toRGBA8(ifds[1]))).toEqual(rgbOf(b))
      expect(ifds[0].t282).toEqual([[300, 1]])
      if (process.env.SAMPLES_DIR) writeFileSync(`${process.env.SAMPLES_DIR}/card-${c}.tif`, tif)
    })
  }

  it('LZW survives a table reset (long varied input)', () => {
    const big = card(1024, 700, 3)
    const tif = encodeTiff([{ width: 1024, height: 700, rgba: big }], 'lzw')
    const ifds = UTIF.decode((tif.buffer as ArrayBuffer))
    UTIF.decodeImage((tif.buffer as ArrayBuffer), ifds[0])
    expect(rgbOf(UTIF.toRGBA8(ifds[0]))).toEqual(rgbOf(big))
  })

  it('LZW of empty input is Clear + EOI', () => {
    expect(lzwEncode(new Uint8Array()).length).toBe(3)
  })

  it('PackBits handles runs and literals', () => {
    const row = Uint8Array.from([1, 1, 1, 1, 2, 3, 4, 4, 5])
    expect([...packBitsRows(row, row.length)]).toEqual([253, 1, 4, 2, 3, 4, 4, 5])
  })
})

describe('BMP', () => {
  it('writes a valid 24-bit bottom-up bitmap', () => {
    const w = 5, h = 3
    const px = card(w, h)
    const bmp = encodeBmp(w, h, px, 300)
    if (process.env.SAMPLES_DIR) writeFileSync(`${process.env.SAMPLES_DIR}/card.bmp`, encodeBmp(301, 97, card(301, 97), 300))
    const dv = new DataView(bmp.buffer)
    expect(String.fromCharCode(bmp[0], bmp[1])).toBe('BM')
    expect(dv.getUint32(2, true)).toBe(bmp.length)
    expect(dv.getUint16(28, true)).toBe(24)
    const rowSize = 16 // 5*3 padded to 16
    expect(bmp.length).toBe(54 + rowSize * h)
    // Top-left pixel lives in the last row of the file, BGR.
    const last = 54 + rowSize * (h - 1)
    expect([bmp[last + 2], bmp[last + 1], bmp[last]]).toEqual([px[0], px[1], px[2]])
    expect(Math.round(dv.getInt32(38, true) * 0.0254)).toBe(300)
  })
})

describe('TGA', () => {
  function decodeTga(t: Uint8Array): number[] {
    const w = t[12] | (t[13] << 8), h = t[14] | (t[15] << 8)
    const rle = t[2] === 10
    const out: number[] = []
    let p = 18
    while (out.length < w * h * 3) {
      if (!rle) { out.push(t[p + 2], t[p + 1], t[p]); p += 3; continue }
      const hdr = t[p++]
      const n = (hdr & 0x7f) + 1
      if (hdr & 0x80) { for (let k = 0; k < n; k++) out.push(t[p + 2], t[p + 1], t[p]); p += 3 }
      else for (let k = 0; k < n; k++) { out.push(t[p + 2], t[p + 1], t[p]); p += 3 }
    }
    return out
  }
  for (const rle of [false, true]) {
    it(`round-trips (${rle ? 'RLE' : 'raw'})`, () => {
      const px = card(300, 120, 5)
      const t = encodeTga(300, 120, px, rle)
      if (process.env.SAMPLES_DIR) writeFileSync(`${process.env.SAMPLES_DIR}/card-${rle ? 'rle' : 'raw'}.tga`, t)
      expect(t[17]).toBe(0x20)
      expect(decodeTga(t)).toEqual(rgbOf(px))
      expect(new TextDecoder().decode(t.subarray(t.length - 18, t.length - 2))).toBe('TRUEVISION-XFILE')
    })
  }
})

describe('DPI metadata', () => {
  it('adds pHYs to a PNG and keeps it decodable', () => {
    const png = encodePng({ width: 4, height: 2, data: card(4, 2), channels: 4 })
    const out = pngWithDpi(png, 300)
    const dec = decodePng(out)
    expect(dec.width).toBe(4)
    expect(Math.round((dec.resolution?.x ?? 0) * 0.0254)).toBe(300)
    // and replacing it again doesn't duplicate the chunk
    const twice = pngWithDpi(out, 150)
    expect(twice.length).toBe(out.length)
  })

  it('crc32 matches the PNG spec check value', () => {
    expect(crc32(new TextEncoder().encode('123456789')) ^ 0xffffffff).toBe(0xcbf43926 | 0)
  })

  it('patches a JFIF density, or inserts one', () => {
    const jfif = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 0xff, 0xd9])
    const a = jpegWithDpi(jfif, 300)
    expect([a[13], (a[14] << 8) | a[15], (a[16] << 8) | a[17]]).toEqual([1, 300, 300])
    const bare = Uint8Array.from([0xff, 0xd8, 0xff, 0xdb, 0, 2, 0xff, 0xd9])
    const b = jpegWithDpi(bare, 72)
    expect([b[2], b[3]]).toEqual([0xff, 0xe0])
    expect(b.length).toBe(bare.length + 18)
  })
})

describe('GIF', () => {
  it('writes an animated GIF with a loop and per-frame delays', () => {
    const f1 = { width: 40, height: 30, rgba: card(40, 30, 1), delayMs: 500 }
    const f2 = { width: 40, height: 30, rgba: card(40, 30, 2), delayMs: 1500 }
    const gif = encodeGif([f1, f2], { colors: 64, dither: true, repeat: 0 })
    expect(new TextDecoder().decode(gif.subarray(0, 6))).toBe('GIF89a')
    expect(gif[gif.length - 1]).toBe(0x3b)
    // NETSCAPE2.0 loop extension present
    expect(new TextDecoder('latin1').decode(gif)).toContain('NETSCAPE2.0')
  })
})
