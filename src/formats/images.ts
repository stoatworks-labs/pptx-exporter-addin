// Still images, one per slide (or one multi-page TIFF).

import { slideText } from '../deck/model'
import { encodeGif } from '../encode/gif'
import { encodeBmp, encodeTga, jpegWithDpi, pngWithDpi } from '../encode/simple'
import { TiffWriter, type TiffCompression } from '../encode/tiff'
import { encodeAvif, encodeWebp } from '../encode/wasm'
import { canEncodeNatively, ctx2d, encodeNative, makeCanvas, rgba, type AnyCanvas } from '../render/canvas'
import { MAX_PIXELS, MAX_SIDE, type Frame } from '../render/sizing'
import { bytesToBase64 } from '../host/types'
import { xmlEscape, xmlSafe } from '../lib/xml'
import { background, eachSlide, expandName, frameFromValues, nameOption, sizeOptions } from './common'
import { hotspots, linkHref, safeUrl } from './overlay'
import { checkCancel, type FormatDef, type OptionDef, type Values } from './types'

type Encoder = (c: AnyCanvas, f: Frame, v: Values) => Promise<Uint8Array>

function imageDef(id: string, label: string, ext: string, blurb: string, extra: OptionDef[]): Omit<FormatDef, 'run'> {
  return {
    id,
    label,
    group: 'images',
    blurb,
    ext: `.${ext}`,
    needs: { images: true },
    options: [...sizeOptions(), ...extra, nameOption('{deck}-{nn}')],
  }
}

/** One file per slide: build the frame, encode it. */
function perSlide(id: string, label: string, ext: string, blurb: string, extra: OptionDef[], encode: Encoder): FormatDef {
  return {
    ...imageDef(id, label, ext, blurb, extra),
    async run(ctx, v) {
      const f = frameFromValues(ctx, v)
      await eachSlide(ctx, async (s, i) => {
        const c = await ctx.images.frame(s.index, f, background(v))
        ctx.add(`${expandName(String(v.name), ctx, s, i)}.${ext}`, await encode(c, f, v))
      })
    },
  }
}

const quality = (def: number): OptionDef => ({ key: 'quality', label: 'Quality', type: 'slider', min: 1, max: 100, step: 1, default: def })

export const png: FormatDef = {
  ...imageDef('png', 'PNG', 'png', 'Lossless. What PowerPoint drew, untouched when the size is the slide’s own shape.', []),
  async run(ctx, v) {
    const f = frameFromValues(ctx, v)
    await eachSlide(ctx, async (s, i) => {
      // framePng hands PowerPoint's own PNG through when nothing needs compositing.
      const bytes = await ctx.images.framePng(s.index, f, background(v))
      ctx.add(`${expandName(String(v.name), ctx, s, i)}.png`, pngWithDpi(bytes, f.dpi))
    })
  },
}

export const jpeg = perSlide(
  'jpeg', 'JPEG', 'jpg', 'Small files for email and the web. Quality is yours to pick.',
  [quality(90)],
  async (c, f, v) => jpegWithDpi(await encodeNative(c, 'image/jpeg', Number(v.quality) / 100), f.dpi),
)

export const webp = perSlide(
  'webp', 'WebP', 'webp', 'Smaller than JPEG at the same quality, or lossless and smaller than PNG.',
  [quality(85), { key: 'lossless', label: 'Lossless', type: 'toggle', default: false }],
  async (c, _f, v) => {
    if (!v.lossless && (await canEncodeNatively('image/webp'))) return encodeNative(c, 'image/webp', Number(v.quality) / 100)
    return encodeWebp(rgba(c), Number(v.quality), !!v.lossless)
  },
)

export const avif = perSlide(
  'avif', 'AVIF', 'avif', 'The smallest of the lot. Slow to encode — this runs in the pane, not on a GPU.',
  [
    quality(60),
    { key: 'speed', label: 'Encoder speed', type: 'slider', min: 0, max: 10, step: 1, default: 7, help: '0 is smallest and slowest; 10 is quick and a bit bigger.' },
  ],
  async (c, _f, v) => {
    if (await canEncodeNatively('image/avif')) return encodeNative(c, 'image/avif', Number(v.quality) / 100)
    return encodeAvif(rgba(c), Number(v.quality), Number(v.speed))
  },
)

export const bmp = perSlide('bmp', 'BMP', 'bmp', '24-bit Windows bitmap, uncompressed — for the kit that still insists.', [],
  async (c, f) => encodeBmp(c.width, c.height, rgba(c).data, f.dpi))

export const tga = perSlide('tga', 'TGA', 'tga', 'Targa, 24-bit — the still format older media servers and vision mixers load.',
  [{ key: 'rle', label: 'Run-length compressed', type: 'toggle', default: true, help: 'Smaller. Turn it off for kit that only reads raw Targa.' }],
  async (c, _f, v) => encodeTga(c.width, c.height, rgba(c).data, !!v.rle))

export const gif = perSlide('gif', 'GIF (still)', 'gif', '256 colours at most. Fine for diagrams, unkind to photos.',
  [
    { key: 'colors', label: 'Colours', type: 'slider', min: 2, max: 256, step: 1, default: 256 },
    { key: 'dither', label: 'Dither', type: 'toggle', default: true, help: 'Smooths gradients into dots instead of bands.' },
  ],
  async (c, _f, v) => encodeGif([{ width: c.width, height: c.height, rgba: rgba(c).data }], { colors: Number(v.colors), dither: !!v.dither, repeat: -1 }))

export const tiff: FormatDef = {
  ...imageDef('tiff', 'TIFF', 'tif', 'Lossless, with real DPI — for print shops and DTP. One file per slide, or one multi-page file.',
    [
      {
        key: 'compression', label: 'Compression', type: 'select', default: 'lzw',
        choices: [
          { value: 'lzw', label: 'LZW (reads everywhere)' },
          { value: 'deflate', label: 'Deflate / ZIP (smaller)' },
          { value: 'packbits', label: 'PackBits' },
          { value: 'none', label: 'None' },
        ],
      },
      { key: 'multipage', label: 'One multi-page file', type: 'toggle', default: false },
    ]),
  async run(ctx, v) {
    const f = frameFromValues(ctx, v)
    const compression = v.compression as TiffCompression
    if (v.multipage) {
      const w = new TiffWriter(compression, ctx.slides.length)
      await eachSlide(ctx, async (s) => {
        const c = await ctx.images.frame(s.index, f, background(v))
        w.addPage({ width: c.width, height: c.height, rgba: rgba(c).data, dpi: f.dpi })
      })
      ctx.add(`${ctx.base}.tif`, w.finish())
      return
    }
    await eachSlide(ctx, async (s, i) => {
      const c = await ctx.images.frame(s.index, f, background(v))
      const w = new TiffWriter(compression, 1)
      w.addPage({ width: c.width, height: c.height, rgba: rgba(c).data, dpi: f.dpi })
      ctx.add(`${expandName(String(v.name), ctx, s, i)}.tif`, w.finish())
    })
  },
}

export const svg: FormatDef = {
  id: 'svg',
  label: 'SVG',
  group: 'images',
  blurb: 'The slide as a picture inside an SVG, with its links still clickable and its text still findable.',
  ext: '.svg',
  needs: { images: true },
  options: [
    ...sizeOptions({ noFit: true }),
    {
      key: 'embed', label: 'Picture inside', type: 'select', default: 'png',
      choices: [{ value: 'png', label: 'PNG (sharp)' }, { value: 'jpeg', label: 'JPEG (small)' }],
    },
    { ...quality(88), when: (v) => v.embed === 'jpeg' },
    { key: 'links', label: 'Clickable links', type: 'toggle', default: true },
    { key: 'text', label: 'Invisible text layer', type: 'toggle', default: true, help: 'Lets search, screen readers and copy-paste find the words.' },
    nameOption('{deck}-{nn}'),
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, { ...v, fit: 'contain' })
    const names = new Map(ctx.slides.map((s, i) => [s.index, `${expandName(String(v.name), ctx, s, i)}.svg`]))
    await eachSlide(ctx, async (s) => {
      const c = await ctx.images.frame(s.index, f, '#ffffff')
      const img = v.embed === 'jpeg' ? await encodeNative(c, 'image/jpeg', Number(v.quality) / 100) : await ctx.images.framePng(s.index, f, '#ffffff')
      const mime = v.embed === 'jpeg' ? 'image/jpeg' : 'image/png'
      const W = f.width, H = f.height
      const parts: string[] = [
        `<?xml version="1.0" encoding="UTF-8"?>`,
        `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">`,
        `<title>${xmlEscape(xmlSafe(s.title || `Slide ${s.index + 1}`))}</title>`,
        // xlink:href alone: every browser reads it, and so do Illustrator and
        // older Inkscape, which ignore SVG 2's plain href.
        `<image width="${W}" height="${H}" xlink:href="data:${mime};base64,${bytesToBase64(img)}"/>`,
      ]
      if (v.text) {
        const text = slideText(s)
        if (text) parts.push(`<desc>${xmlEscape(xmlSafe(text))}</desc>`)
        for (const sh of s.shapes) {
          if (!sh.paragraphs || !sh.rect) continue
          const x = (sh.rect.x / ctx.deck.width) * W
          let y = (sh.rect.y / ctx.deck.height) * H
          for (const p of sh.paragraphs) {
            const line = p.runs.map((r) => (r.br ? ' ' : r.text)).join('').trim()
            if (!line) continue
            const pt = p.runs.find((r) => r.size)?.size ?? 18
            const px = (pt * 12700 / ctx.deck.height) * H
            y += px
            parts.push(`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" font-size="${px.toFixed(1)}" font-family="sans-serif" fill="#000" fill-opacity="0">${xmlEscape(xmlSafe(line))}</text>`)
            y += px * 0.2
          }
        }
      }
      if (v.links) {
        for (const h of hotspots(ctx.deck, s)) {
          const raw = linkHref(h.link, (i) => names.get(i) ?? '', s.index, ctx.deck.slides.length)
          const href = h.link.url ? safeUrl(h.link.url) : raw
          if (!href) continue
          parts.push(`<a href="${xmlEscape(href)}" xlink:href="${xmlEscape(href)}"><title>${xmlEscape(xmlSafe(h.link.tooltip ?? h.label))}</title><rect x="${(h.x * W).toFixed(1)}" y="${(h.y * H).toFixed(1)}" width="${(h.w * W).toFixed(1)}" height="${(h.h * H).toFixed(1)}" fill="#fff" fill-opacity="0"/></a>`)
        }
      }
      parts.push('</svg>')
      ctx.add(names.get(s.index)!, parts.join('\n'))
    })
  },
}

export const contactSheet: FormatDef = {
  id: 'contact',
  label: 'Contact sheet',
  group: 'images',
  blurb: 'Every slide as a thumbnail on one big picture — for running orders and sign-off.',
  ext: '.png / .jpg',
  needs: { images: true },
  options: [
    { key: 'columns', label: 'Columns', type: 'number', min: 1, max: 12, step: 1, default: 4 },
    { key: 'thumb', label: 'Thumbnail width', type: 'number', min: 80, max: 1920, step: 10, unit: 'px', default: 480 },
    { key: 'gap', label: 'Gap', type: 'number', min: 0, max: 200, step: 2, unit: 'px', default: 24 },
    {
      key: 'labels', label: 'Labels', type: 'select', default: 'title',
      choices: [{ value: 'none', label: 'None' }, { value: 'number', label: 'Slide number' }, { value: 'title', label: 'Number and title' }],
    },
    { key: 'background', label: 'Background', type: 'color', default: '#1f2328' },
    { key: 'ink', label: 'Label colour', type: 'color', default: '#f0f0f0' },
    { key: 'hiddenMark', label: 'Mark hidden slides', type: 'toggle', default: true },
    {
      key: 'format', label: 'Save as', type: 'select', default: 'png',
      choices: [{ value: 'png', label: 'PNG' }, { value: 'jpeg', label: 'JPEG' }],
    },
  ],
  async run(ctx, v) {
    const cols = Math.max(1, Number(v.columns))
    const tw = Number(v.thumb)
    const th = Math.round(tw * (ctx.deck.height / ctx.deck.width))
    const gap = Number(v.gap)
    const labelH = v.labels === 'none' ? 0 : Math.max(16, Math.round(tw * 0.06))
    const W = cols * tw + (cols + 1) * gap
    const rowH = th + labelH + gap
    // A web view won't draw past ~16k pixels a side or ~35M pixels, so a long
    // deck goes onto several sheets.
    const maxRows = Math.max(1, Math.min(Math.floor((MAX_SIDE - gap) / rowH), Math.floor(MAX_PIXELS / W / rowH)))
    const perSheet = maxRows * cols
    const sheets = Math.ceil(ctx.slides.length / perSheet)
    const font = `${Math.round(labelH * 0.7)}px -apple-system, "Segoe UI", system-ui, sans-serif`
    const thumbFrame: Frame = { width: tw, height: th, dx: 0, dy: 0, dw: tw, dh: th, renderW: tw, renderH: th, dpi: 96 }
    for (let sheet = 0; sheet < sheets; sheet++) {
      const slides = ctx.slides.slice(sheet * perSheet, (sheet + 1) * perSheet)
      const rows = Math.ceil(slides.length / cols)
      const H = rows * rowH + gap
      const c = makeCanvas(W, H)
      const g = ctx2d(c)
      g.fillStyle = String(v.background)
      g.fillRect(0, 0, W, H)
      for (let i = 0; i < slides.length; i++) {
        checkCancel(ctx.signal)
        const s = slides[i]
        ctx.progress((sheet * perSheet + i) / ctx.slides.length, `Thumbnail ${s.index + 1}`)
        const x = gap + (i % cols) * (tw + gap)
        const y = gap + Math.floor(i / cols) * rowH
        const thumb = await ctx.images.frame(s.index, thumbFrame, '#000')
        g.drawImage(thumb as CanvasImageSource, x, y)
        if (s.hidden && v.hiddenMark) {
          g.fillStyle = 'rgba(0,0,0,0.55)'
          g.fillRect(x, y, tw, th)
          g.fillStyle = '#fff'
          g.font = font
          g.textAlign = 'center'
          g.textBaseline = 'middle'
          g.fillText('HIDDEN', x + tw / 2, y + th / 2)
          g.textAlign = 'left'
        }
        if (labelH) {
          g.fillStyle = String(v.ink)
          g.font = font
          g.textBaseline = 'middle'
          const label = v.labels === 'title' && s.title ? `${s.index + 1}  ${s.title}` : String(s.index + 1)
          g.fillText(clip(g, label, tw), x, y + th + labelH / 2 + 2)
        }
      }
      const bytes = v.format === 'jpeg' ? await encodeNative(c, 'image/jpeg', 0.9) : await encodeNative(c, 'image/png')
      const suffix = sheets > 1 ? `-${sheet + 1}` : ''
      ctx.add(`${ctx.base}-contact-sheet${suffix}.${v.format === 'jpeg' ? 'jpg' : 'png'}`, bytes)
    }
    ctx.progress(1)
  },
}

function clip(g: ReturnType<typeof ctx2d>, text: string, width: number): string {
  if (g.measureText(text).width <= width) return text
  let t = text
  while (t.length > 1 && g.measureText(`${t}…`).width > width) t = t.slice(0, -1)
  return `${t}…`
}

export const IMAGE_FORMATS: FormatDef[] = [png, jpeg, webp, avif, tiff, bmp, tga, gif, svg, contactSheet]
