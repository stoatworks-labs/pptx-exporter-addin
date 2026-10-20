// A baseline TIFF writer: 8-bit RGB, one or many pages, uncompressed, LZW,
// Deflate or PackBits, with the horizontal predictor for the two that gain
// from it. Pure — unit tested by decoding the result with UTIF.

import { zlibSync } from 'fflate'

export type TiffCompression = 'none' | 'lzw' | 'deflate' | 'packbits'

export interface TiffPage {
  width: number
  height: number
  /** RGBA, 4 bytes a pixel (alpha is dropped — slides are opaque). */
  rgba: Uint8Array | Uint8ClampedArray
  dpi?: number
}

const COMPRESSION_CODE: Record<TiffCompression, number> = { none: 1, lzw: 5, deflate: 8, packbits: 32773 }

// Field types
const SHORT = 3
const LONG = 4
const RATIONAL = 5
const ASCII = 2

interface Entry {
  tag: number
  type: number
  values: number[] | string
}

export function encodeTiff(pages: TiffPage[], compression: TiffCompression = 'lzw', software = 'Advanced Exporter'): Uint8Array {
  if (!pages.length) throw new Error('A TIFF needs at least one page.')
  const w = new TiffWriter(compression, pages.length, software)
  for (const p of pages) w.addPage(p)
  return w.finish()
}

/** Page-at-a-time writer, so a 100-page TIFF never holds 100 pages of RGBA. */
export class TiffWriter {
  private chunks: Uint8Array[] = []
  private offset = 8
  private pageIndex = 0
  // Patched once each IFD's position is known.
  private prevNextPtr: { chunk: Uint8Array; at: number }
  private predictor: boolean

  constructor(private compression: TiffCompression, private pageCount: number, private software = 'Advanced Exporter') {
    const header = new Uint8Array(8)
    header[0] = 0x49 // 'I'
    header[1] = 0x49
    new DataView(header.buffer).setUint16(2, 42, true)
    this.chunks.push(header)
    this.prevNextPtr = { chunk: header, at: 4 }
    this.predictor = compression === 'lzw' || compression === 'deflate'
  }

  addPage(page: TiffPage): void {
    const { chunks, compression, predictor, software } = this
    const pages = { length: this.pageCount }
    const pageIndex = this.pageIndex++
    let offset = this.offset
    const { width, height } = page
    const rowBytes = width * 3
    // Strips of about 256 KB uncompressed: small enough for every reader.
    const rowsPerStrip = Math.max(1, Math.min(height, Math.floor((256 * 1024) / rowBytes)))
    const stripOffsets: number[] = []
    const stripCounts: number[] = []
    for (let y0 = 0; y0 < height; y0 += rowsPerStrip) {
      const rows = Math.min(rowsPerStrip, height - y0)
      const raw = new Uint8Array(rows * rowBytes)
      for (let y = 0; y < rows; y++) {
        let si = (y0 + y) * width * 4
        let di = y * rowBytes
        for (let x = 0; x < width; x++, si += 4, di += 3) {
          raw[di] = page.rgba[si]
          raw[di + 1] = page.rgba[si + 1]
          raw[di + 2] = page.rgba[si + 2]
        }
        if (predictor) {
          const rs = y * rowBytes
          for (let i = rs + rowBytes - 1; i >= rs + 3; i--) raw[i] = (raw[i] - raw[i - 3]) & 0xff
        }
      }
      let data: Uint8Array
      switch (compression) {
        case 'lzw': data = lzwEncode(raw); break
        case 'deflate': data = zlibSync(raw, { level: 6 }); break
        case 'packbits': data = packBitsRows(raw, rowBytes); break
        default: data = raw
      }
      stripOffsets.push(offset)
      stripCounts.push(data.length)
      chunks.push(data)
      offset += data.length
    }
    if (offset % 2) {
      chunks.push(new Uint8Array(1))
      offset++
    }

    const dpi = Math.max(1, Math.round(page.dpi ?? 72))
    const entries: Entry[] = [
      { tag: 254, type: LONG, values: [pages.length > 1 ? 2 : 0] },
      { tag: 256, type: LONG, values: [width] },
      { tag: 257, type: LONG, values: [height] },
      { tag: 258, type: SHORT, values: [8, 8, 8] },
      { tag: 259, type: SHORT, values: [COMPRESSION_CODE[compression]] },
      { tag: 262, type: SHORT, values: [2] },
      { tag: 273, type: LONG, values: stripOffsets },
      { tag: 277, type: SHORT, values: [3] },
      { tag: 278, type: LONG, values: [rowsPerStrip] },
      { tag: 279, type: LONG, values: stripCounts },
      { tag: 282, type: RATIONAL, values: [dpi, 1] },
      { tag: 283, type: RATIONAL, values: [dpi, 1] },
      { tag: 284, type: SHORT, values: [1] },
      { tag: 296, type: SHORT, values: [2] },
      { tag: 297, type: SHORT, values: [pageIndex, pages.length] },
      { tag: 305, type: ASCII, values: software },
    ]
    if (predictor) entries.push({ tag: 317, type: SHORT, values: [2] })
    entries.sort((a, b) => a.tag - b.tag)

    // IFD: count, 12-byte entries, next-IFD pointer; then out-of-line values.
    const ifdSize = 2 + entries.length * 12 + 4
    const ifdStart = offset
    let extra = ifdStart + ifdSize
    const extras: Uint8Array[] = []
    const ifd = new Uint8Array(ifdSize)
    const dv = new DataView(ifd.buffer)
    dv.setUint16(0, entries.length, true)
    entries.forEach((e, i) => {
      const p = 2 + i * 12
      const bytes = valueBytes(e)
      const count = typeof e.values === 'string' ? e.values.length + 1 : e.type === RATIONAL ? e.values.length / 2 : e.values.length
      dv.setUint16(p, e.tag, true)
      dv.setUint16(p + 2, e.type, true)
      dv.setUint32(p + 4, count, true)
      if (bytes.length <= 4) {
        ifd.set(bytes, p + 8)
      } else {
        dv.setUint32(p + 8, extra, true)
        extras.push(bytes)
        extra += bytes.length
        if (bytes.length % 2) {
          extras.push(new Uint8Array(1))
          extra++
        }
      }
    })
    // Point the previous IFD (or the header) at this one.
    const prev = this.prevNextPtr
    new DataView(prev.chunk.buffer, prev.chunk.byteOffset).setUint32(prev.at, ifdStart, true)
    this.prevNextPtr = { chunk: ifd, at: ifdSize - 4 }
    chunks.push(ifd, ...extras)
    this.offset = extra
  }

  finish(): Uint8Array {
    const out = new Uint8Array(this.offset)
    let p = 0
    for (const c of this.chunks) {
      out.set(c, p)
      p += c.length
    }
    return out
  }
}

function valueBytes(e: Entry): Uint8Array {
  if (typeof e.values === 'string') {
    const s = new TextEncoder().encode(e.values.replace(/[^\x20-\x7e]/g, '?'))
    const b = new Uint8Array(s.length + 1)
    b.set(s)
    return b
  }
  const size = e.type === SHORT ? 2 : e.type === LONG ? 4 : 4
  const b = new Uint8Array(e.values.length * size)
  const dv = new DataView(b.buffer)
  e.values.forEach((v, i) => {
    if (e.type === SHORT) dv.setUint16(i * 2, v, true)
    else dv.setUint32(i * 4, v, true)
  })
  return b
}

/** TIFF-flavoured LZW: MSB-first codes, 9–12 bits, written the way libtiff writes it. */
export function lzwEncode(data: Uint8Array): Uint8Array {
  const CLEAR = 256
  const EOI = 257
  // Dictionary as a flat table indexed by (prefix code << 8 | byte), with a
  // generation stamp so a Clear costs nothing.
  const codeOf = new Int16Array(4096 * 256)
  const gen = new Uint16Array(4096 * 256)
  let curGen = 1
  let out = new Uint8Array(Math.max(1024, data.length >> 1))
  let outLen = 0
  let bitBuf = 0
  let bitCount = 0
  let width = 9
  let next = 258

  const emit = (code: number) => {
    bitBuf = (bitBuf << width) | code
    bitCount += width
    while (bitCount >= 8) {
      if (outLen >= out.length) {
        const grown = new Uint8Array(out.length * 2)
        grown.set(out)
        out = grown
      }
      out[outLen++] = (bitBuf >>> (bitCount - 8)) & 0xff
      bitCount -= 8
    }
    bitBuf &= (1 << bitCount) - 1
  }
  // libtiff's rule: widen once the next free code passes the current width's
  // largest; the decoder, a code behind, widens one entry earlier ("early
  // change"). Clear when the table reaches 4094.
  const advance = () => {
    next++
    if (next === 4094) {
      emit(CLEAR)
      reset()
    } else if (next === 512) width = 10
    else if (next === 1024) width = 11
    else if (next === 2048) width = 12
  }
  const reset = () => {
    curGen++
    if (curGen === 0xffff) {
      gen.fill(0)
      curGen = 1
    }
    next = 258
    width = 9
  }

  emit(CLEAR)
  if (data.length === 0) {
    emit(EOI)
  } else {
    let prefix = data[0]
    for (let i = 1; i < data.length; i++) {
      const c = data[i]
      const key = (prefix << 8) | c
      if (gen[key] === curGen) {
        prefix = codeOf[key]
        continue
      }
      emit(prefix)
      codeOf[key] = next
      gen[key] = curGen
      advance()
      prefix = c
    }
    // The decoder adds one more entry on reading the last code, so the width
    // may step up (or the table clear) before EOI — as libtiff's encoder does.
    emit(prefix)
    advance()
    emit(EOI)
  }
  if (bitCount > 0) {
    if (outLen >= out.length) {
      const grown = new Uint8Array(out.length + 1)
      grown.set(out)
      out = grown
    }
    out[outLen++] = (bitBuf << (8 - bitCount)) & 0xff
  }
  return out.slice(0, outLen)
}

/** PackBits, row by row as the spec asks (runs never cross a row). */
export function packBitsRows(data: Uint8Array, rowBytes: number): Uint8Array {
  const out: number[] = []
  for (let r = 0; r < data.length; r += rowBytes) {
    const end = Math.min(data.length, r + rowBytes)
    let i = r
    while (i < end) {
      // A run of 3+ identical bytes?
      let run = 1
      while (i + run < end && run < 128 && data[i + run] === data[i]) run++
      if (run >= 3) {
        out.push((257 - run) & 0xff, data[i])
        i += run
        continue
      }
      // Literal stretch until the next run of 3.
      let lit = 0
      const start = i
      while (i < end && lit < 128) {
        if (i + 2 < end && data[i] === data[i + 1] && data[i] === data[i + 2]) break
        i++
        lit++
      }
      out.push(lit - 1)
      for (let k = start; k < start + lit; k++) out.push(data[k])
    }
  }
  return Uint8Array.from(out)
}
