// The deck as something a browser plays or reads.
//
// Every slide goes out as the picture PowerPoint drew, so it looks right
// everywhere. What a picture can't carry — links, video, the words, the notes
// — goes alongside as data for the page to put back. Three flavours: our own
// player (no dependencies, works opened from disk), reveal.js for people who
// already live there, and one long page for reading on a phone and printing.

import PLAYER_CSS from '../../player/player.css?raw'
import PLAYER_JS from '../../player/player.js?raw'
import type { Deck, Link, Paragraph, Run, Slide, Transition } from '../deck/model'
import { slideText } from '../deck/model'
import { encodeWebp } from '../encode/wasm'
import { bytesToBase64 } from '../host/types'
import { xmlSafe } from '../lib/xml'
import { canEncodeNatively, encodeNative, rgba } from '../render/canvas'
import type { Frame } from '../render/sizing'
import { eachSlide, frameFromValues, sizeOptions } from './common'
import { allLinks, hotspots, mediaSpots, safeUrl, type Box, type MediaSpot } from './overlay'
import { checkCancel, type ExportContext, type FormatDef, type OptionDef, type Values } from './types'

// ---------------------------------------------------------------- the player's data

/** The deck as player/player.js reads it, from
 * <script type="application/json" id="deck-data">. Boxes (x, y, w, h) are
 * fractions of the slide, 0–1, from its top-left corner. */
export interface PlayerDeck {
  version: 1
  title: string
  /** BCP 47 language of the deck's text, for <html lang>. */
  lang?: string
  /** Width ÷ height. */
  aspect: number
  /** Pixel size of the slide pictures. */
  width: number
  height: number
  /** Fills the window around a slide that isn't the window's shape. */
  background: string
  slides: PlayerSlide[]
  options: PlayerOptions
}

export interface PlayerSlide {
  /** The picture: a relative path, a data: URI, or '#id' — the src of that
   * element in the page (a single file keeps each picture once, in its
   * no-script fallback). */
  src: string
  /** Alt text for the picture. The full text is in `text`. */
  alt: string
  title: string
  /** Number in the deck (1-based); differs from its place when slides are left out. */
  number: number
  /** Everything readable on the slide, a line per paragraph — for screen readers and Find. */
  text: string
  /** Speaker notes as HTML built here, everything from the deck escaped.
   * Absent when notes are off. */
  notes?: string
  /** Hidden in PowerPoint: stepped over, but reachable by link or number. */
  hidden?: boolean
  transition?: PlayerTransition
  /** The slide's own timing (PowerPoint's "After"), in ms. */
  advanceMs?: number
  links?: PlayerLink[]
  media?: PlayerMedia[]
}

export interface PlayerTransition {
  type: 'none' | 'fade' | 'slide'
  ms: number
  /** Which way the slides travel, for 'slide'. */
  dir?: 'left' | 'right' | 'up' | 'down'
}

/** With a box it is a click area over the slide; without one, it is only in
 * the text layer (a link inside a paragraph, which a picture can't place). */
export interface PlayerLink extends Partial<Box> {
  /** http(s), mailto or tel only. */
  href?: string
  /** Another slide, by its place in the export (0-based). */
  slide?: number
  label: string
  tooltip?: string
}

export interface PlayerMedia extends Box {
  kind: 'video' | 'audio'
  src: string
  poster?: string
  type?: string
  label: string
}

export interface PlayerOptions {
  controls: boolean
  progress: boolean
  numbers: boolean
  loop: boolean
  auto: 'off' | 'deck' | 'every'
  autoSeconds: number
  transitions: 'deck' | 'fade' | 'slide' | 'none'
  /** 0 = each slide's own duration. */
  transitionMs: number
  theme: 'dark' | 'light'
  /** The presenter view (S). Only offered when the notes were written. */
  presenter: boolean
  /** Step over hidden slides, as PowerPoint's slide show does. */
  skipHidden: boolean
}

export interface DeckDataInput {
  deck: Deck
  slides: Slide[]
  title: string
  background: string
  width: number
  height: number
  /** The picture for the slide at this place in the export. */
  src: (slide: Slide, position: number) => string
  /** Where an embedded video or audio went, or null to leave it out. */
  media?: (spot: MediaSpot, slide: Slide) => { src: string; poster?: string; type?: string } | null
  notes: boolean
  links: boolean
  options: Omit<PlayerOptions, 'presenter'>
}

export function buildDeckData(o: DeckDataInput): PlayerDeck {
  // Links and jumps name slides by their deck index; the player knows only
  // its own list, which may have slides left out.
  const positions = new Map(o.slides.map((s, i) => [s.index, i]))
  const slides = o.slides.map((s, i): PlayerSlide => {
    const d: PlayerSlide = {
      src: o.src(s, i),
      alt: clean(s.title) || `Slide ${s.index + 1}`,
      title: clean(s.title),
      number: s.index + 1,
      text: clean(slideText(s)),
    }
    if (o.notes) {
      const notes = notesHtml(s.notes)
      if (notes) d.notes = notes
    }
    if (s.hidden) d.hidden = true
    const t = mapTransition(s.transition)
    if (t.type !== 'none') d.transition = t
    if (s.transition?.advanceAfterMs != null) d.advanceMs = s.transition.advanceAfterMs
    if (o.links) {
      const links = slideLinks(o.deck, s, i, positions, o.slides.length)
      if (links.length) d.links = links
    }
    if (o.media) {
      const media: PlayerMedia[] = []
      for (const spot of mediaSpots(o.deck, s)) {
        const m = o.media(spot, s)
        if (!m) continue
        media.push({
          ...roundBox(spot),
          kind: spot.media.kind,
          src: m.src,
          ...(m.poster ? { poster: m.poster } : {}),
          ...(m.type ? { type: m.type } : {}),
          label: clean(spot.label),
        })
      }
      if (media.length) d.media = media
    }
    return d
  })
  return {
    version: 1,
    lang: o.deck.lang,
    title: clean(o.title),
    aspect: Math.round((o.deck.width / o.deck.height) * 1e6) / 1e6,
    width: o.width,
    height: o.height,
    background: cssColor(o.background),
    slides,
    options: { ...o.options, presenter: o.notes },
  }
}

const FADES = new Set(['fade', 'dissolve'])
const SLIDES = new Set(['push', 'pull', 'cover', 'uncover', 'wipe', 'split'])
const DIRS: Record<string, PlayerTransition['dir']> = { l: 'left', r: 'right', u: 'up', d: 'down' }

/** PowerPoint's dozens of transitions, as the three a picture player can do
 * honestly. `dir` in the file is the way the slide travels (push dir="u"
 * comes up from the bottom); diagonals keep their sideways part. */
export function mapTransition(t: Transition | undefined): PlayerTransition {
  if (!t || t.type === 'none' || t.type === 'cut') return { type: 'none', ms: 0 }
  const ms = Math.max(0, Math.min(10000, Math.round(t.durationMs || 0)))
  if (SLIDES.has(t.type)) return { type: 'slide', ms, dir: DIRS[(t.dir ?? 'l').charAt(0)] ?? 'left' }
  if (FADES.has(t.type)) return { type: 'fade', ms }
  return { type: 'fade', ms }
}

function roundBox(b: Box): Box {
  const r = (n: number) => Math.round(n * 1e4) / 1e4
  return { x: r(b.x), y: r(b.y), w: r(b.w), h: r(b.h) }
}

function slideLinks(deck: Deck, slide: Slide, position: number, positions: Map<number, number>, count: number): PlayerLink[] {
  const target = (l: Link): { href: string } | { slide: number } | null => {
    if (l.url) {
      const href = safeUrl(l.url)
      return href ? { href } : null
    }
    if (l.slide != null) {
      // A slide that isn't in the export has nowhere to go.
      const p = positions.get(l.slide)
      return p == null ? null : { slide: p }
    }
    switch (l.jump) {
      case 'first': return { slide: 0 }
      case 'last': return { slide: count - 1 }
      case 'next': return position + 1 < count ? { slide: position + 1 } : null
      case 'previous': return position > 0 ? { slide: position - 1 } : null
      default: return null
    }
  }
  const key = (t: { href: string } | { slide: number }) => ('href' in t ? `u:${t.href}` : `s:${t.slide}`)
  const out: PlayerLink[] = []
  const seen = new Set<string>()
  for (const h of hotspots(deck, slide)) {
    const t = target(h.link)
    if (!t) continue
    seen.add(key(t))
    out.push({ ...roundBox(h), ...t, label: clean(h.label), ...(h.link.tooltip ? { tooltip: clean(h.link.tooltip) } : {}) })
  }
  // Links inside text a picture can't place: listed, not drawn.
  for (const { text, link } of allLinks(slide)) {
    const t = target(link)
    if (!t || seen.has(key(t))) continue
    seen.add(key(t))
    const label = clean(text).trim() || ('href' in t ? t.href : `Slide ${t.slide + 1}`)
    out.push({ ...t, label, ...(link.tooltip ? { tooltip: clean(link.tooltip) } : {}) })
  }
  return out
}

/** Notes as a little HTML: paragraphs, emphasis and safe links, nothing else. */
export function notesHtml(ps: Paragraph[]): string {
  let start = 0
  let end = ps.length
  const blank = (p: Paragraph) => p.runs.every((r) => r.br || !r.text.trim())
  while (start < end && blank(ps[start])) start++
  while (end > start && blank(ps[end - 1])) end--
  const out: string[] = []
  for (const p of ps.slice(start, end)) {
    if (blank(p)) {
      out.push('<p><br></p>')
      continue
    }
    const indent = p.level > 0 ? ` style="margin-left:${p.level * 1.25}em"` : ''
    const bullet = p.bullet === 'char' ? '• ' : ''
    out.push(`<p${indent}>${bullet}${p.runs.map(runHtml).join('')}</p>`)
  }
  return out.join('')
}

function runHtml(r: Run): string {
  if (r.br) return '<br>'
  let s = esc(r.text)
  if (!s) return ''
  if (r.bold) s = `<strong>${s}</strong>`
  if (r.italic) s = `<em>${s}</em>`
  if (r.underline) s = `<u>${s}</u>`
  if (r.strike) s = `<s>${s}</s>`
  if (r.baseline) s = r.baseline > 0 ? `<sup>${s}</sup>` : `<sub>${s}</sub>`
  const href = r.link?.url ? safeUrl(r.link.url) : null
  if (href) s = `<a href="${esc(href)}" target="_blank" rel="noopener noreferrer">${s}</a>`
  return s
}

/** JSON that can sit inside a <script> element: nothing in it can close the
 * element or open a comment, and the two line separators JavaScript once
 * choked on are escaped. All of these only ever occur inside strings. */
export function jsonForScript(value: unknown): string {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029')
}

/** Escape for HTML text and attribute values (and drop what HTML can't carry). */
export function esc(s: string): string {
  return xmlSafe(s).replace(/[<>&"']/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&#39;' })[c]!)
}

function clean(s: string | undefined): string {
  return xmlSafe(s ?? '')
}

function cssColor(c: unknown): string {
  return typeof c === 'string' && /^#[0-9a-f]{3,8}$/i.test(c) ? c : '#000000'
}

/** Our own scripts and styles, safe to put inline. */
function inlineJs(js: string): string {
  return js.replace(/<\/(script)/gi, '<\\/$1').replace(/<!--/g, '<\\!--')
}

function inlineCss(css: string): string {
  return css.replace(/<\/(style)/gi, '<\\/$1')
}

function isDark(hex: string): boolean {
  let h = cssColor(hex).slice(1)
  if (h.length < 6) h = h.slice(0, 3).split('').map((c) => c + c).join('')
  const n = parseInt(h.slice(0, 6), 16)
  const lum = 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)
  return lum < 140
}

// ---------------------------------------------------------------- shared options and pictures

const PICTURE_CHOICES = [
  { value: 'webp', label: 'WebP (smallest)' },
  { value: 'jpeg', label: 'JPEG' },
  { value: 'png', label: 'PNG (lossless, biggest)' },
]

function pictureOptions(): OptionDef[] {
  return [
    {
      key: 'image', label: 'Slide pictures', type: 'select', default: 'webp', choices: PICTURE_CHOICES,
      help: 'WebP is roughly a third the size of JPEG for slides, and every current browser shows it.',
    },
    { key: 'quality', label: 'Quality', type: 'slider', min: 1, max: 100, step: 1, default: 85, when: (v) => v.image !== 'png' },
    // The pictures keep the slide's shape; the page letterboxes them.
    ...sizeOptions({ size: '1080p', noFit: true, noDpi: true }),
  ]
}

function packagingOption(help: string): OptionDef {
  return {
    key: 'packaging', label: 'Packaging', type: 'select', default: 'folder', help,
    choices: [
      { value: 'folder', label: 'A folder: index.html and its files' },
      { value: 'single', label: 'One self-contained .html file' },
    ],
  }
}

function titleOption(): OptionDef {
  return { key: 'title', label: 'Page title', type: 'text', default: '', placeholder: 'The deck’s title' }
}

const NOTES_PRIVATE = 'Off leaves the notes out of the files altogether, not merely out of sight — worth it when they were for your eyes only and the export is going to anyone else.'

interface Picture {
  bytes: Uint8Array
  ext: string
  mime: string
}

async function picture(ctx: ExportContext, s: Slide, f: Frame, v: Values): Promise<Picture> {
  const q = Number(v.quality)
  // Slides are opaque, so the colour behind only matters to a transparent one.
  if (v.image === 'png') return { bytes: await ctx.images.framePng(s.index, f, '#ffffff'), ext: 'png', mime: 'image/png' }
  const c = await ctx.images.frame(s.index, f, '#ffffff')
  if (v.image === 'jpeg') return { bytes: await encodeNative(c, 'image/jpeg', q / 100), ext: 'jpg', mime: 'image/jpeg' }
  const bytes = (await canEncodeNatively('image/webp'))
    ? await encodeNative(c, 'image/webp', q / 100)
    : await encodeWebp(rgba(c), q, false)
  return { bytes, ext: 'webp', mime: 'image/webp' }
}

const dataUri = (mime: string, bytes: Uint8Array) => `data:${mime};base64,${bytesToBase64(bytes)}`

/** slides/07.webp — named by the deck's slide number, padded to sort. */
function slidePath(ctx: ExportContext, s: Slide, ext: string): string {
  const width = Math.max(2, String(Math.max(...ctx.slides.map((x) => x.index + 1))).length)
  return `slides/${String(s.index + 1).padStart(width, '0')}.${ext}`
}

function pageTitle(ctx: ExportContext, v: Values): string {
  return String(v.title ?? '').trim() || ctx.deck.title || ctx.base
}

/** Every slide's picture, as a path (files written) or a data: URI. */
async function writePictures(ctx: ExportContext, v: Values, f: Frame, inline: boolean): Promise<{ src: string[] }> {
  const src: string[] = []
  await eachSlide(ctx, async (s, i) => {
    const pic = await picture(ctx, s, f, v)
    if (inline) {
      src[i] = dataUri(pic.mime, pic.bytes)
    } else {
      src[i] = slidePath(ctx, s, pic.ext)
      ctx.add(src[i], pic.bytes)
    }
  })
  checkCancel(ctx.signal)
  return { src }
}

// ---------------------------------------------------------------- media

/** A single file carries embedded video and audio as base64 text, a third
 * bigger than the files and all held in memory when the page opens. Past
 * this much (before encoding) it stops being a sensible web page. */
export const INLINE_MEDIA_CAP = 40 * 1024 * 1024

const MEDIA_TYPES: Record<string, string> = {
  mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', ogv: 'video/ogg',
  mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', ogg: 'audio/ogg', aac: 'audio/aac',
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp',
}

function mediaWriter(ctx: ExportContext, single: boolean): NonNullable<DeckDataInput['media']> {
  const done = new Map<string, { src: string | null; why?: string }>()
  let inlined = 0
  const mb = (n: number) => `${(n / 1048576).toFixed(n < 10485760 ? 1 : 0)} MB`

  const put = (part: string, poster: boolean) => {
    const known = done.get(part)
    if (known) return { ...known, first: false }
    const bytes = ctx.deck.readPart(part)
    const type = ctx.deck.contentTypeOf(part) ?? MEDIA_TYPES[part.split('.').pop()!.toLowerCase()] ?? 'application/octet-stream'
    let r: { src: string | null; why?: string }
    if (!bytes) {
      r = { src: null, why: 'is missing from the deck file' }
    } else if (!single) {
      const path = `media/${part.split('/').pop()}`
      ctx.add(path, bytes)
      r = { src: path }
    } else if (poster || inlined + bytes.length <= INLINE_MEDIA_CAP) {
      if (!poster) inlined += bytes.length
      r = { src: dataUri(type, bytes) }
    } else {
      r = { src: null, why: `(${mb(bytes.length)}) would take the file past the ${mb(INLINE_MEDIA_CAP)} a single file carries — use the folder packaging to keep it` }
    }
    done.set(part, r)
    return { ...r, first: true }
  }

  return (spot, slide) => {
    const m = spot.media
    const what = `Slide ${slide.index + 1}: the ${m.kind} “${spot.label}”`
    if (!m.part) {
      ctx.warn(`${what} is linked rather than embedded${m.external ? ` (${m.external})` : ''}, so it isn't in the deck to copy. Left out.`)
      return null
    }
    const r = put(m.part, false)
    if (!r.src) {
      if (r.first) ctx.warn(`${what} ${r.why}. Left out.`)
      return null
    }
    const poster = m.poster ? put(m.poster, true).src ?? undefined : undefined
    return { src: r.src, poster, type: m.contentType ?? ctx.deck.contentTypeOf(m.part) }
  }
}

// ---------------------------------------------------------------- HTML5 player

function playerOptions(deck: Deck, v: Values): Omit<PlayerOptions, 'presenter'> {
  const transitions = String(v.transitions)
  return {
    controls: !!v.controls,
    progress: !!v.progress,
    numbers: !!v.numbers,
    loop: v.loop === 'deck' ? deck.show.loop : v.loop === 'yes',
    // "As set up in PowerPoint" honours Slide Show › Use Timings.
    auto: v.auto === 'every' ? 'every' : v.auto === 'deck' && deck.show.useTimings ? 'deck' : 'off',
    autoSeconds: Number(v.autoSeconds) || 8,
    transitions: transitions === 'fade' || transitions === 'slide' || transitions === 'none' ? transitions : 'deck',
    transitionMs: Number(v.transitionMs) || 0,
    theme: v.theme === 'light' ? 'light' : 'dark',
    skipHidden: true,
  }
}

/** What a browser without JavaScript shows: every slide as a plain picture. */
function fallbackHtml(data: PlayerDeck, pictures: string[], single: boolean): string {
  const items = data.slides.map((s, i) => {
    const n = i + 1
    const size = `width="${data.width}" height="${data.height}"`
    const alt = esc(s.text || s.alt)
    // In a single file this is the only copy of the picture; the player's
    // data points here ('#s7').
    const img = single
      ? `<img id="s${n}" src="${pictures[i]}" alt="${alt}" ${size}>`
      : `<img src="${esc(s.src)}" alt="${alt}" ${size} loading="lazy">`
    const extra: string[] = []
    for (const l of s.links ?? []) {
      const href = l.href ?? `#f${(l.slide ?? 0) + 1}`
      extra.push(`<li><a href="${esc(href)}">${esc(l.label)}</a></li>`)
    }
    if (!single) for (const m of s.media ?? []) extra.push(`<li><a href="${esc(m.src)}">${m.kind === 'audio' ? 'Audio' : 'Video'}: ${esc(m.label)}</a></li>`)
    const list = extra.length ? `<ul>${extra.join('')}</ul>` : ''
    return `<li><figure id="f${n}">${img}<figcaption>${n}. ${esc(s.title)}</figcaption>${list}</figure></li>`
  })
  const body = `<div class="ap-fallback"><h1>${esc(data.title)}</h1><ol>\n${items.join('\n')}\n</ol></div>`
  return single ? body : `<noscript>${body}</noscript>`
}

export function html5Page(data: PlayerDeck, fallback: string, single: boolean): string {
  return `<!doctype html>
<html${data.lang ? ` lang="${esc(data.lang)}"` : ''}>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="generator" content="Advanced Exporter">
<title>${esc(data.title)}</title>
${single ? '<script>document.documentElement.className += " js"</script>\n' : ''}${single ? `<style>\n${inlineCss(PLAYER_CSS)}\n</style>` : '<link rel="stylesheet" href="player.css">'}
<style>html, body { background: ${data.background}; }</style>
</head>
<body>
${fallback}
<div id="player"></div>
<script type="application/json" id="deck-data">${jsonForScript(data)}</script>
${single ? `<script>\n${inlineJs(PLAYER_JS)}\n</script>` : '<script src="player.js"></script>'}
</body>
</html>
`
}

export const html5: FormatDef = {
  id: 'html5',
  label: 'Web presentation (HTML5)',
  group: 'web',
  blurb: 'Plays in any browser, even straight from disk: keys, clicks and swipes, working links and video, a presenter view with your notes.',
  ext: '.html',
  needs: { images: true },
  options: [
    packagingOption(`A folder is lighter and starts quicker. One file is easier to email, but every picture rides inside it as text, about a third bigger, and video and audio only go in while they total ${INLINE_MEDIA_CAP / 1048576} MB or less.`),
    ...pictureOptions(),
    {
      key: 'transitions', label: 'Transitions', type: 'select', default: 'deck',
      choices: [
        { value: 'deck', label: 'As in the deck' },
        { value: 'fade', label: 'Fade, every slide' },
        { value: 'slide', label: 'Slide, every slide' },
        { value: 'none', label: 'None' },
      ],
      help: 'From the deck, fades and dissolves play as a fade; push, wipe, cover, uncover and split as a slide; anything fancier as a fade. Anyone who asks their system for less motion gets a cut.',
    },
    {
      key: 'transitionMs', label: 'Transition length', type: 'number', min: 0, max: 5000, step: 50, unit: 'ms', default: 0,
      help: '0 keeps each slide’s own.', when: (v) => v.transitions !== 'none',
    },
    {
      key: 'auto', label: 'Advance', type: 'select', default: 'deck',
      choices: [
        { value: 'off', label: 'On a click or key only' },
        { value: 'deck', label: 'Slide timings, as set up in PowerPoint' },
        { value: 'every', label: 'Every few seconds' },
      ],
      help: 'Timings are each slide’s “After” time, and only when Slide Show › Use Timings is on. A playing video holds the slide.',
    },
    { key: 'autoSeconds', label: 'Seconds per slide', type: 'number', min: 1, max: 3600, step: 1, unit: 's', default: 8, when: (v) => v.auto === 'every' },
    {
      key: 'loop', label: 'At the end', type: 'select', default: 'deck',
      choices: [
        { value: 'deck', label: 'As set up in PowerPoint' },
        { value: 'yes', label: 'Loop back to the start' },
        { value: 'no', label: 'Stop on the last slide' },
      ],
    },
    { key: 'controls', label: 'Control bar', type: 'toggle', default: true },
    { key: 'progress', label: 'Progress bar', type: 'toggle', default: true },
    { key: 'numbers', label: 'Slide counter', type: 'toggle', default: true },
    { key: 'notes', label: 'Speaker notes and presenter view', type: 'toggle', default: false, help: NOTES_PRIVATE },
    { key: 'links', label: 'Clickable links', type: 'toggle', default: true },
    {
      key: 'media', label: 'Video and audio', type: 'toggle', default: true,
      help: 'Copies embedded video and audio out of the deck to play over the slide. Linked media was never in the deck file, so it can’t come along.',
    },
    titleOption(),
    { key: 'background', label: 'Background', type: 'color', default: '#000000', help: 'Around the slide when the window is a different shape.' },
    {
      key: 'theme', label: 'Controls', type: 'select', default: 'dark',
      choices: [{ value: 'dark', label: 'Dark' }, { value: 'light', label: 'Light' }],
    },
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, v)
    const single = v.packaging === 'single'
    const pics = await writePictures(ctx, v, f, single)
    const data = buildDeckData({
      deck: ctx.deck,
      slides: ctx.slides,
      title: pageTitle(ctx, v),
      background: String(v.background),
      width: f.width,
      height: f.height,
      src: (_s, i) => (single ? `#s${i + 1}` : pics.src[i]),
      media: v.media ? mediaWriter(ctx, single) : undefined,
      notes: !!v.notes,
      links: !!v.links,
      options: playerOptions(ctx.deck, v),
    })
    const page = html5Page(data, fallbackHtml(data, pics.src, single), single)
    if (single) {
      ctx.add(`${ctx.base}.html`, page)
      return
    }
    ctx.add('index.html', page)
    ctx.add('player.js', PLAYER_JS)
    ctx.add('player.css', PLAYER_CSS)
  },
}

// ---------------------------------------------------------------- reveal.js

export const REVEAL_VERSION = '6.0.2'
const REVEAL_CDN = `https://cdn.jsdelivr.net/npm/reveal.js@${REVEAL_VERSION}/dist`

/** Subresource-integrity hashes of the jsDelivr files, so a tampered CDN copy
 * is refused. Computed from node_modules/reveal.js and checked against the CDN
 * (test/web.test.ts recomputes them — bump REVEAL_VERSION, and it fails until
 * these are updated). */
export const REVEAL_SRI: Record<string, string> = {
  'reveal.css': 'sha384-rKClnwfQQ2b7LYSOAhRvm5m4LS7QKciQTAXBpVvr30y6oGB+ZBBXlpDnvFqlv5Eu',
  'reveal.js': 'sha384-UilXW7W77EPIY5TmQGBrjhBPcfX40NqGwuIC8DGs6OeGTCxKy36B+DiOXtfqgk9S',
  'plugin/notes.js': 'sha384-XUVoTb/UbE06yoCUaqswCPKSUvPq3q4heijGsemnUSS/JEgW1lNnK0rKlNwjuz7O',
}

function sri(base: string, file: string): string {
  return base === REVEAL_CDN ? ` integrity="${REVEAL_SRI[file]}" crossorigin="anonymous"` : ''
}

/** reveal.js's own files, for an export that works offline: loaded only when
 * asked for, so they stay out of the pane's main bundle. A glob rather than
 * plain imports because the package's exports map hides dist/reveal.js, and
 * because a build without reveal.js installed should lose this one option,
 * not fail outright. `exhaustive` lets the glob look inside node_modules. */
const REVEAL_DIST = import.meta.glob<string>(
  [
    '../../node_modules/reveal.js/dist/reveal.js',
    '../../node_modules/reveal.js/dist/reveal.css',
    '../../node_modules/reveal.js/dist/plugin/notes.js',
    '../../node_modules/reveal.js/LICENSE',
  ],
  { query: '?raw', import: 'default', exhaustive: true },
)

async function revealFiles(notes: boolean): Promise<Record<string, string>> {
  const load = (tail: string) => {
    const key = Object.keys(REVEAL_DIST).find((k) => k.endsWith(`/reveal.js/${tail}`))
    return key ? REVEAL_DIST[key]() : null
  }
  const want: [string, string][] = [
    ['reveal/reveal.js', 'dist/reveal.js'],
    ['reveal/reveal.css', 'dist/reveal.css'],
    // The minified files carry no notice of their own, and MIT wants one.
    ['reveal/LICENSE', 'LICENSE'],
  ]
  if (notes) want.push(['reveal/plugin/notes.js', 'dist/plugin/notes.js'])
  const files: Record<string, string> = {}
  for (const [path, tail] of want) {
    const body = load(tail)
    if (!body) throw new Error(`This build of the add-in doesn’t include reveal.js ${REVEAL_VERSION}, so it can’t be copied into the export. Choose jsDelivr instead.`)
    files[path] = await body
  }
  return files
}

const REVEAL_TRANSITIONS = ['none', 'fade', 'slide', 'convex', 'concave', 'zoom'] as const

function revealSpeed(ms: number): string | null {
  if (ms <= 0) return null
  return ms <= 400 ? 'fast' : ms >= 1000 ? 'slow' : 'default'
}

export function revealPage(o: {
  data: PlayerDeck
  base: string
  transition: string
  controls: boolean
  progress: boolean
  slideNumber: boolean
  hash: boolean
  autoSlide: boolean
}): string {
  const { data } = o
  const W = data.width
  const H = data.height
  const bg = data.background
  const ink = isDark(bg) ? '#f2f2f2' : '#1c1c1c'
  const deckMode = o.transition === 'deck'
  const sections = data.slides.map((s, i) => {
    const attrs = [
      `data-background-image="${esc(s.src)}"`,
      'data-background-size="contain"',
      `data-background-color="${esc(bg)}"`,
      `aria-label="${esc(`Slide ${i + 1}${s.title ? `: ${s.title}` : ''}`)}"`,
    ]
    if (deckMode) {
      const t = s.transition ?? { type: 'none', ms: 0 }
      attrs.push(`data-transition="${t.type}"`, `data-background-transition="${t.type}"`)
      const speed = revealSpeed(t.ms)
      if (speed && t.type !== 'none') attrs.push(`data-transition-speed="${speed}"`)
    }
    if (o.autoSlide && s.advanceMs != null) attrs.push(`data-autoslide="${Math.max(250, s.advanceMs)}"`)
    const parts: string[] = []
    const words = s.text.split('\n').filter((l) => l.trim()).map((l) => `<p>${esc(l)}</p>`).join('')
    const links = (s.links ?? []).map((l) => `<li>${revealLink(l, 'tabindex="-1"')}${esc(l.label)}</a></li>`).join('')
    parts.push(`<div class="ae-sr">${words}${links ? `<ul>${links}</ul>` : ''}</div>`)
    const hot = (s.links ?? []).filter((l) => l.w && l.h)
    if (hot.length) {
      const anchors = hot.map((l) => {
        const style = `left:${(l.x! * W).toFixed(1)}px;top:${(l.y! * H).toFixed(1)}px;width:${(l.w! * W).toFixed(1)}px;height:${(l.h! * H).toFixed(1)}px`
        const label = l.href ? `${l.label} (opens in a new tab)` : l.label
        return `${revealLink(l, `style="${style}" aria-label="${esc(label)}"${l.tooltip ? ` title="${esc(l.tooltip)}"` : ''}`)}</a>`
      })
      parts.push(`<div class="ae-hot" style="width:${W}px;height:${H}px">${anchors.join('')}</div>`)
    }
    if (s.notes) parts.push(`<aside class="notes">${s.notes}</aside>`)
    return `<section ${attrs.join(' ')}>\n${parts.join('\n')}\n</section>`
  })
  const config = {
    hash: o.hash,
    hashOneBasedIndex: true,
    controls: o.controls,
    progress: o.progress,
    slideNumber: o.slideNumber ? 'c/t' : false,
    transition: deckMode ? 'none' : o.transition,
    backgroundTransition: deckMode ? 'none' : o.transition,
    loop: data.options.loop,
    width: W,
    height: H,
    margin: 0,
    center: false,
  }
  const notes = data.options.presenter
  return `<!doctype html>
<html${data.lang ? ` lang="${esc(data.lang)}"` : ''}>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Advanced Exporter">
<title>${esc(data.title)}</title>
<link rel="stylesheet" href="${o.base}/reveal.css"${sri(o.base, 'reveal.css')}>
<style>
html, body, .reveal { background: ${bg}; color: ${ink}; }
.reveal .controls, .reveal .progress { color: ${ink}; }
.reveal .slide-number { background: rgba(0, 0, 0, 0.4); color: #fff; }
.ae-sr { position: absolute; width: 1px; height: 1px; margin: -1px; overflow: hidden; clip: rect(0 0 0 0); clip-path: inset(50%); white-space: nowrap; }
.ae-hot { position: absolute; left: 0; top: 0; pointer-events: none; }
.ae-hot a { position: absolute; display: block; pointer-events: auto; }
.ae-hot a:focus-visible { outline: 3px solid #ffbf47; outline-offset: 2px; }
</style>
</head>
<body>
<div class="reveal">
<div class="slides">
${sections.join('\n')}
</div>
</div>
<script src="${o.base}/reveal.js"${sri(o.base, 'reveal.js')}></script>
${notes ? `<script src="${o.base}/plugin/notes.js"${sri(o.base, 'plugin/notes.js')}></script>\n` : ''}<script>
Reveal.initialize(Object.assign(${jsonForScript(config)}, { plugins: ${notes ? 'window.RevealNotes ? [RevealNotes] : []' : '[]'} }));
// Links to other slides go by reveal's own API, whatever the hash setting.
document.addEventListener('click', function (e) {
  var a = e.target && e.target.closest ? e.target.closest('a[data-slide]') : null;
  if (!a || e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey) return;
  e.preventDefault();
  Reveal.slide(parseInt(a.getAttribute('data-slide'), 10));
});
</script>
</body>
</html>
`
}

function revealLink(l: PlayerLink, attrs: string): string {
  if (l.href) return `<a href="${esc(l.href)}" target="_blank" rel="noopener noreferrer" ${attrs}>`
  return `<a href="#/${(l.slide ?? 0) + 1}" data-slide="${l.slide ?? 0}" ${attrs}>`
}

export const reveal: FormatDef = {
  id: 'reveal',
  label: 'reveal.js',
  group: 'web',
  blurb: 'A reveal.js presentation: each slide’s picture as its background, links still clickable, notes in reveal’s speaker view.',
  ext: '.html',
  needs: { images: true },
  options: [
    {
      key: 'source', label: 'reveal.js from', type: 'select', default: 'cdn',
      choices: [
        { value: 'cdn', label: `jsDelivr, pinned to ${REVEAL_VERSION} (needs the internet)` },
        { value: 'offline', label: 'Copied into the export (works offline)' },
      ],
    },
    ...pictureOptions(),
    {
      key: 'transition', label: 'Transition', type: 'select', default: 'deck',
      choices: [
        { value: 'deck', label: 'As in the deck' },
        ...REVEAL_TRANSITIONS.map((t) => ({ value: t, label: t.charAt(0).toUpperCase() + t.slice(1) })),
      ],
    },
    { key: 'controls', label: 'Arrows', type: 'toggle', default: true },
    { key: 'progress', label: 'Progress bar', type: 'toggle', default: true },
    { key: 'slideNumber', label: 'Slide numbers', type: 'toggle', default: true },
    { key: 'hash', label: 'Slide number in the address', type: 'toggle', default: true, help: 'So a reload, or a link someone sends, lands on the same slide.' },
    { key: 'notes', label: 'Speaker notes (S)', type: 'toggle', default: false, help: NOTES_PRIVATE },
    { key: 'links', label: 'Clickable links', type: 'toggle', default: true },
    titleOption(),
    { key: 'background', label: 'Background', type: 'color', default: '#000000' },
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, v)
    const pics = await writePictures(ctx, v, f, false)
    // reveal.js has no "hidden but reachable" — an exported hidden slide is
    // simply one of the slides.
    const data = buildDeckData({
      deck: ctx.deck,
      slides: ctx.slides,
      title: pageTitle(ctx, v),
      background: String(v.background),
      width: f.width,
      height: f.height,
      src: (_s, i) => pics.src[i],
      notes: !!v.notes,
      links: !!v.links,
      options: { ...playerOptions(ctx.deck, { ...v, loop: 'deck', auto: 'deck' }), skipHidden: false },
    })
    let base = REVEAL_CDN
    if (v.source === 'offline') {
      ctx.progress(1, 'Copying reveal.js')
      for (const [path, body] of Object.entries(await revealFiles(!!v.notes))) ctx.add(path, body)
      base = 'reveal'
    }
    ctx.add('index.html', revealPage({
      data,
      base,
      transition: String(v.transition),
      controls: !!v.controls,
      progress: !!v.progress,
      slideNumber: !!v.slideNumber,
      hash: !!v.hash,
      autoSlide: ctx.deck.show.useTimings,
    }))
  },
}

// ---------------------------------------------------------------- scrolling page

export function scrollPageHtml(data: PlayerDeck): string {
  const articles = data.slides.map((s, i) => {
    const n = i + 1
    const links = (s.links ?? []).map((l) => {
      const href = l.href ?? `#slide-${(l.slide ?? 0) + 1}`
      const ext = l.href ? ' target="_blank" rel="noopener noreferrer"' : ''
      return `<li><a href="${esc(href)}"${ext}>${esc(l.label)}</a></li>`
    }).join('')
    return `<article class="slide" id="slide-${n}">
<h2><a class="num" href="#slide-${n}" aria-label="Slide ${n}">${n}</a> ${esc(s.title || `Slide ${s.number}`)}</h2>
<img class="pic" src="${s.src.startsWith('data:') ? s.src : esc(s.src)}" alt="${esc(s.text || s.alt)}" width="${data.width}" height="${data.height}" loading="lazy" decoding="async">
${links ? `<ul class="links" aria-label="Links on this slide">${links}</ul>\n` : ''}${s.notes ? `<section class="notes" aria-label="Notes"><h3>Notes</h3>${s.notes}</section>\n` : ''}</article>`
  })
  return `<!doctype html>
<html${data.lang ? ` lang="${esc(data.lang)}"` : ''}>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Advanced Exporter">
<title>${esc(data.title)}</title>
<style>
:root { color-scheme: light dark; --bg: #ffffff; --fg: #1c1c1c; --muted: #5f5f5f; --line: #dddddd; --card: #f5f5f5; --accent: #0a5fc2; }
@media (prefers-color-scheme: dark) { :root { --bg: #161616; --fg: #ececec; --muted: #a6a6a6; --line: #333333; --card: #202020; --accent: #6cb4ff; } }
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--fg); font: 17px/1.55 -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, Roboto, sans-serif; -webkit-text-size-adjust: 100%; }
header, main { max-width: 1100px; margin: 0 auto; padding: 0 16px; }
header { padding-top: 24px; padding-bottom: 8px; }
h1 { margin: 0 0 4px; font-size: 1.6rem; line-height: 1.25; }
.count { margin: 0; color: var(--muted); }
.slide { margin: 0 0 56px; scroll-margin-top: 8px; }
.slide h2 { margin: 24px 0 10px; font-size: 1.15rem; line-height: 1.3; }
.num { display: inline-block; min-width: 2em; color: var(--muted); font-variant-numeric: tabular-nums; text-decoration: none; }
.pic { display: block; width: 100%; height: auto; border-radius: 6px; background: ${data.background}; box-shadow: 0 1px 4px rgba(0, 0, 0, 0.18); }
.links { margin: 10px 0 0; padding-left: 1.2em; font-size: 0.95rem; }
a { color: var(--accent); }
a:focus-visible { outline: 3px solid #ffbf47; outline-offset: 2px; }
.notes { margin-top: 12px; padding: 12px 16px; border-left: 3px solid var(--line); border-radius: 4px; background: var(--card); overflow-wrap: break-word; }
.notes h3 { margin: 0 0 6px; color: var(--muted); font-size: 0.75rem; letter-spacing: 0.06em; text-transform: uppercase; }
.notes p { margin: 0 0 0.6em; }
.notes p:last-child { margin-bottom: 0; }
@media print {
  @page { margin: 14mm; }
  :root { --bg: #ffffff; --fg: #000000; --muted: #555555; --card: transparent; --line: #999999; --accent: #000000; }
  body { font-size: 11pt; }
  header { padding-top: 0; }
  .slide { margin: 0; break-after: page; page-break-after: always; break-inside: avoid; }
  .slide:last-child { break-after: auto; page-break-after: auto; }
  .slide h2 { margin-top: 0; }
  .pic { width: auto; max-width: 100%; max-height: 120mm; margin: 0 auto; box-shadow: none; border: 1px solid #cccccc; }
  .notes { padding: 8px 0 0 12px; }
  .num { color: inherit; }
}
</style>
</head>
<body>
<header>
<h1>${esc(data.title)}</h1>
<p class="count">${data.slides.length} slide${data.slides.length === 1 ? '' : 's'}</p>
</header>
<main>
${articles.join('\n')}
</main>
<script>
// Lazy pictures that haven't scrolled into view yet would print blank.
addEventListener('beforeprint', function () {
  Array.prototype.forEach.call(document.querySelectorAll('img[loading="lazy"]'), function (i) { i.loading = 'eager'; });
});
</script>
</body>
</html>
`
}

export const scrollPage: FormatDef = {
  id: 'scroll',
  label: 'Scrolling page',
  group: 'web',
  blurb: 'Every slide down one long page with its notes underneath. Reads on a phone; prints a slide a page.',
  ext: '.html',
  needs: { images: true },
  options: [
    packagingOption('A folder is lighter and the pictures load as you scroll. One file is easier to send on, but every picture rides inside it as text, about a third bigger.'),
    ...pictureOptions(),
    { key: 'notes', label: 'Speaker notes', type: 'toggle', default: true, help: NOTES_PRIVATE },
    titleOption(),
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, v)
    const single = v.packaging === 'single'
    const pics = await writePictures(ctx, v, f, single)
    const data = buildDeckData({
      deck: ctx.deck,
      slides: ctx.slides,
      title: pageTitle(ctx, v),
      background: '#000000',
      width: f.width,
      height: f.height,
      src: (_s, i) => pics.src[i],
      notes: !!v.notes,
      links: true,
      options: playerOptions(ctx.deck, v),
    })
    ctx.add(single ? `${ctx.base}.html` : 'index.html', scrollPageHtml(data))
  },
}

export const WEB_FORMATS: FormatDef[] = [html5, reveal, scrollPage]
