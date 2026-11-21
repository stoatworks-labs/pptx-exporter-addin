// Transparent backgrounds: the matte maths, the slide copy PowerPoint draws
// them from, and the encoders that carry the alpha.
import { describe, expect, it } from 'vitest'
import * as UTIF from 'utif2'
import { fillAndKey, hasTransparency, pullMatte } from '../src/render/matte'
import { overlayCopy } from '../src/render/overlay-copy'
import { entryData, makeEntry, readZip, writeZip } from '../src/formats/ooxml-write'
import { encodeTiff } from '../src/encode/tiff'
import { encodeTga } from '../src/encode/simple'
import { parseDeck } from '../src/deck/parse'
import { all, kid, parseXml } from '../src/lib/xml'
import { fixture } from './helpers'

/** Content of colour `f` at alpha `a` (0..255) composited over a background, the way a renderer does. */
function over(f: number[], a: number, bg: number): number[] {
  return f.map((c) => Math.round((c * a + bg * (255 - a)) / 255))
}

function px(...pixels: number[][]): Uint8ClampedArray {
  return Uint8ClampedArray.from(pixels.flatMap((p) => [...p, 255]))
}

describe('matte', () => {
  const cases: { f: number[]; a: number }[] = [
    { f: [255, 255, 255], a: 255 }, // white text: a colour key would lose it
    { f: [0, 0, 0], a: 255 }, // black text
    { f: [200, 30, 60], a: 255 },
    { f: [200, 30, 60], a: 128 }, // a see-through fill
    { f: [0, 0, 0], a: 64 }, // a soft shadow
    { f: [20, 120, 240], a: 12 }, // the faint end of a glow
    { f: [0, 0, 0], a: 0 }, // nothing there
  ]
  const onBlack = px(...cases.map((c) => over(c.f, c.a, 0)))
  const onWhite = px(...cases.map((c) => over(c.f, c.a, 255)))
  const out = pullMatte(onBlack, onWhite)

  cases.forEach((c, i) => {
    it(`recovers ${c.f.join(',')} at alpha ${c.a}`, () => {
      const a = out[i * 4 + 3]
      expect(Math.abs(a - c.a)).toBeLessThanOrEqual(1)
      if (c.a === 0) {
        expect([...out.subarray(i * 4, i * 4 + 3)]).toEqual([0, 0, 0])
        return
      }
      // Colour error, as it shows once composited: small at every alpha.
      for (let k = 0; k < 3; k++) expect(Math.abs(out[i * 4 + k] - c.f[k]) * (c.a / 255)).toBeLessThanOrEqual(1)
    })
  })

  it('composites back to what PowerPoint drew, on either background', () => {
    for (let i = 0; i < cases.length; i++) {
      const f = [...out.subarray(i * 4, i * 4 + 3)], a = out[i * 4 + 3]
      for (const [bg, ref] of [[0, onBlack], [255, onWhite]] as const) {
        over(f, a, bg).forEach((v, k) => expect(Math.abs(v - ref[i * 4 + k])).toBeLessThanOrEqual(1))
      }
    }
  })

  it('treats a white render darker than the black one as solid', () => {
    const m = pullMatte(px([100, 100, 100]), px([90, 90, 90]))
    expect(m[3]).toBe(255)
  })

  it('refuses pictures of different sizes', () => {
    expect(() => pullMatte(px([0, 0, 0]), px([0, 0, 0], [0, 0, 0]))).toThrow(/same size/)
  })

  it('splits into a pre-multiplied fill and a key', () => {
    const { fill, key } = fillAndKey(Uint8ClampedArray.from([200, 100, 50, 128, 10, 20, 30, 255, 99, 99, 99, 0]))
    expect([...fill]).toEqual([100, 50, 25, 255, 10, 20, 30, 255, 0, 0, 0, 255])
    expect([...key]).toEqual([128, 128, 128, 255, 255, 255, 255, 255, 0, 0, 0, 255])
    expect(hasTransparency(fill)).toBe(false)
    expect(hasTransparency(out)).toBe(true)
  })
})

// A one-slide package like Slide.exportAsBase64's: a picture filling the
// slide at the back, a background of its own, a title and a small picture.
const PIC = (x: number, y: number, w: number, h: number, name: string, p = 'p:', a = 'a:') =>
  `<${p}pic><${p}nvPicPr><${p}cNvPr id="${name.length + 10}" name="${name}"/><${p}cNvPicPr/><${p}nvPr/></${p}nvPicPr>` +
  `<${p}blipFill><${a}blip r:embed="rId2"/></${p}blipFill>` +
  `<${p}spPr><${a}xfrm><${a}off x="${x}" y="${y}"/><${a}ext cx="${w}" cy="${h}"/></${a}xfrm><${a}prstGeom prst="rect"/></${p}spPr></${p}pic>`

function onePackage(slideXml: string): Uint8Array {
  return writeZip([
    makeEntry('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="12192000" cy="6858000"/></p:presentation>'),
    makeEntry('ppt/slides/slide1.xml', slideXml),
  ])
}

const PREFIXED = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
  + '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>'
  + '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
  + PIC(-20000, -10000, 12250000, 6900000, 'Backdrop')
  + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>'
  + PIC(1000000, 1000000, 3000000, 2000000, 'Logo')
  + '</p:spTree></p:cSld></p:sld>'

function slideOf(bytes: Uint8Array): Document {
  const e = readZip(bytes).find((x) => x.name === 'ppt/slides/slide1.xml')!
  return parseXml(entryData(e))
}

const names = (doc: Document) => all(doc, 'cNvPr').map((e) => e.getAttribute('name'))

describe('overlay copy', () => {
  it('puts the slide on a solid colour and hides the master graphics', () => {
    const { bytes, dropped } = overlayCopy(onePackage(PREFIXED), '#000000', { masterGraphics: false, dropBackPictures: false })
    const doc = slideOf(bytes)
    expect(dropped).toBe(0)
    expect(doc.documentElement.getAttribute('showMasterSp')).toBe('0')
    const cSld = kid(doc.documentElement, 'cSld')!
    // bg must be cSld's first child, and there must be one.
    expect(cSld.firstElementChild?.localName).toBe('bg')
    expect(all(doc, 'bg')).toHaveLength(1)
    expect(all(doc, 'bgRef')).toHaveLength(0)
    expect(all(cSld.firstElementChild, 'srgbClr')[0].getAttribute('val')).toBe('000000')
    expect(cSld.firstElementChild?.prefix).toBe('p')
    expect(all(cSld.firstElementChild, 'solidFill')[0].prefix).toBe('a')
    expect(names(doc)).toEqual(['', 'Backdrop', 'Title', 'Logo'])
  })

  it('keeps the master graphics when asked', () => {
    const doc = slideOf(overlayCopy(onePackage(PREFIXED), '#FFFFFF', { masterGraphics: true, dropBackPictures: false }).bytes)
    expect(doc.documentElement.hasAttribute('showMasterSp')).toBe(false)
    expect(all(doc, 'srgbClr')[0].getAttribute('val')).toBe('FFFFFF')
  })

  it('drops only the picture that fills the slide from the back', () => {
    const { bytes, dropped } = overlayCopy(onePackage(PREFIXED), '#ffffff', { masterGraphics: false, dropBackPictures: true })
    expect(dropped).toBe(1)
    expect(names(slideOf(bytes))).toEqual(['', 'Title', 'Logo'])
  })

  it('leaves a full-slide picture alone when something is behind it', () => {
    const xml = PREFIXED.replace(PIC(-20000, -10000, 12250000, 6900000, 'Backdrop'), PIC(1000000, 1000000, 3000000, 2000000, 'Small') + PIC(0, 0, 12192000, 6858000, 'Backdrop'))
    const { dropped } = overlayCopy(onePackage(xml), '#ffffff', { masterGraphics: false, dropBackPictures: true })
    expect(dropped).toBe(0)
  })

  it('matches on local names, whatever the prefixes', () => {
    const xml = '<sld xmlns="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<cSld><spTree><nvGrpSpPr/><grpSpPr/>' + PIC(0, 0, 12192000, 6858000, 'Backdrop', '', 'd:') + '</spTree></cSld></sld>'
    const { bytes, dropped } = overlayCopy(onePackage(xml), '#000000', { masterGraphics: false, dropBackPictures: true })
    const doc = slideOf(bytes)
    expect(dropped).toBe(1)
    const bg = kid(kid(doc.documentElement, 'cSld'), 'bg')!
    expect(bg.namespaceURI).toBe('http://schemas.openxmlformats.org/presentationml/2006/main')
    expect(bg.prefix).toBeNull()
    expect(all(bg, 'solidFill')[0].prefix).toBe('d')
  })

  it('leaves a real deck readable, every slide on the colour', () => {
    const src = fixture('sample-16x9.pptx')
    const before = parseDeck(src)
    const { bytes } = overlayCopy(src, '#123456', { masterGraphics: false, dropBackPictures: false })
    const after = parseDeck(bytes)
    expect(after.slides.map((s) => s.title)).toEqual(before.slides.map((s) => s.title))
    const slides = readZip(bytes).filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.name))
    expect(slides.length).toBe(before.slides.length)
    for (const e of slides) {
      const doc = parseXml(entryData(e))
      expect(all(doc, 'srgbClr').some((c) => c.getAttribute('val') === '123456')).toBe(true)
    }
    // Parts it didn't touch are copied as they were.
    const media = readZip(src).filter((e) => e.name.startsWith('ppt/media/'))
    for (const m of media) expect(readZip(bytes).find((e) => e.name === m.name)?.raw).toEqual(m.raw)
  })

  it('refuses a colour that isn’t one', () => {
    expect(() => overlayCopy(onePackage(PREFIXED), 'black', { masterGraphics: false, dropBackPictures: false })).toThrow(/Not a colour/)
  })
})

/** A card with every alpha level across it. */
function alphaCard(w: number, h: number): Uint8Array {
  const out = new Uint8Array(w * h * 4)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4
      out[i] = (x * 7) & 0xff
      out[i + 1] = (y * 13) & 0xff
      out[i + 2] = 200
      out[i + 3] = (x * 255 / (w - 1)) | 0
    }
  }
  return out
}

describe('alpha in the encoders', () => {
  for (const c of ['none', 'lzw', 'deflate', 'packbits'] as const) {
    it(`TIFF keeps unassociated alpha (${c})`, () => {
      const card = alphaCard(257, 9)
      const tif = encodeTiff([{ width: 257, height: 9, rgba: card, dpi: 72, alpha: true }], c)
      const buf = (tif.buffer as ArrayBuffer).slice(tif.byteOffset, tif.byteOffset + tif.byteLength) as ArrayBuffer
      const [ifd] = UTIF.decode(buf)
      expect(ifd.t277).toEqual([4])
      expect(ifd.t338).toEqual([2])
      expect(ifd.t258).toEqual([8, 8, 8, 8])
      UTIF.decodeImage(buf, ifd)
      expect([...UTIF.toRGBA8(ifd)]).toEqual([...card])
    })
  }

  function decodeTga32(t: Uint8Array): number[] {
    const w = t[12] | (t[13] << 8), h = t[14] | (t[15] << 8)
    const rle = t[2] === 10
    const out: number[] = []
    let p = 18
    const one = () => { out.push(t[p + 2], t[p + 1], t[p], t[p + 3]) }
    while (out.length < w * h * 4) {
      if (!rle) { one(); p += 4; continue }
      const hdr = t[p++]
      const n = (hdr & 0x7f) + 1
      if (hdr & 0x80) { for (let k = 0; k < n; k++) one(); p += 4 }
      else for (let k = 0; k < n; k++) { one(); p += 4 }
    }
    return out
  }

  for (const rle of [false, true]) {
    it(`TGA writes 32-bit with straight alpha (${rle ? 'RLE' : 'raw'})`, () => {
      const card = alphaCard(300, 7)
      // A run of transparent pixels for the RLE packets.
      card.fill(0, 0, 4 * 150)
      const t = encodeTga(300, 7, card, rle, true)
      expect(t[16]).toBe(32)
      expect(t[17]).toBe(0x28)
      expect(decodeTga32(t)).toEqual([...card])
      const dv = new DataView(t.buffer, t.byteOffset)
      const ext = dv.getUint32(t.length - 26, true)
      expect(dv.getUint16(ext, true)).toBe(495)
      expect(t[ext + 494]).toBe(3)
      expect(ext + 495).toBe(t.length - 26)
      expect(new TextDecoder().decode(t.subarray(t.length - 18, t.length - 2))).toBe('TRUEVISION-XFILE')
    })
  }

  it('TGA without alpha is unchanged: 24-bit, no extension area', () => {
    const t = encodeTga(4, 1, new Uint8Array(16).fill(9), false)
    expect(t[16]).toBe(24)
    expect(t[17]).toBe(0x20)
    expect(t.length).toBe(18 + 12 + 26)
    expect(new DataView(t.buffer).getUint32(t.length - 26, true)).toBe(0)
  })
})
