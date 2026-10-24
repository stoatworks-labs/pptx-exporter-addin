// Words out: Markdown in three flavours, outlines PowerPoint and outliners can
// read back, the speaker notes on their own, and the deck as CSV or JSON.

import type { Link, Paragraph, Run, Shape, Slide, TableCell } from '../deck/model'
import { paragraphsText, slideText } from '../deck/model'
import { canEncodeNatively, encodeNative } from '../render/canvas'
import type { Frame } from '../render/sizing'
import { encodeWebp } from '../encode/wasm'
import { rgba } from '../render/canvas'
import { xmlEscape, xmlSafe } from '../lib/xml'
import { eachSlide } from './common'
import { cropRegion, passThroughImage } from './crop'
import { safeUrl } from './overlay'
import { structure, type Block } from './structure'
import type { ExportContext, FormatDef, Values } from './types'

// --- Markdown ---------------------------------------------------------------------

/** Escape text so Markdown shows it literally. */
export function mdEscape(text: string): string {
  return text.replace(/([\\`*_[\]<>|~])/g, '\\$1').replace(/^(\s*)([#+-]|\d+\.)(\s)/, '$1\\$2$3')
}

function mdRun(r: Run, link: (l: Link) => string | null): string {
  if (r.br) return '  \n'
  const raw = r.text
  if (!raw) return ''
  // Emphasis markers can't touch spaces, so keep them outside the markers.
  const lead = raw.match(/^\s*/)![0]
  const trail = raw.slice(lead.length).match(/\s*$/)![0]
  const core = raw.slice(lead.length, raw.length - trail.length)
  if (!core) return raw
  let t = r.mono ? `\`${core.replace(/`/g, 'ˋ')}\`` : mdEscape(core)
  if (r.bold && r.italic) t = `***${t}***`
  else if (r.bold) t = `**${t}**`
  else if (r.italic) t = `*${t}*`
  if (r.strike) t = `~~${t}~~`
  if (r.baseline && r.baseline > 0) t = `<sup>${t}</sup>`
  if (r.baseline && r.baseline < 0) t = `<sub>${t}</sub>`
  const href = r.link ? link(r.link) : null
  if (href) t = `[${t}](${href.replace(/[()\s]/g, encodeURIComponent)})`
  return lead + t + trail
}

export function mdParagraphs(ps: Paragraph[], link: (l: Link) => string | null): string {
  const lines: string[] = []
  const counters: number[] = []
  // Each nesting level indents by the width of its parent's marker ("- " or "1. ").
  const widths: number[] = []
  let prevWasList = false
  for (const p of ps) {
    const text = p.runs.map((r) => mdRun(r, link)).join('').trim()
    if (!text) continue
    if (p.bullet === 'none') {
      if (prevWasList) lines.push('')
      lines.push(text, '')
      prevWasList = false
      counters.length = 0
      continue
    }
    counters.length = p.level + 1
    counters[p.level] = (counters[p.level] ?? 0) + 1
    const marker = p.bullet === 'number' ? `${counters[p.level]}.` : '-'
    widths[p.level] = marker.length + 1
    let indent = 0
    for (let l = 0; l < p.level; l++) indent += widths[l] ?? 2
    lines.push(`${' '.repeat(indent)}${marker} ${text}`)
    prevWasList = true
  }
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()
}

function mdTable(rows: TableCell[][], link: (l: Link) => string | null): string {
  const cellText = (c: TableCell) => c.paragraphs.map((p) => p.runs.map((r) => mdRun(r, link)).join('')).join('<br>').replace(/\n/g, ' ').trim()
  const width = Math.max(...rows.map((r) => r.length))
  const grid = rows.map((r) => {
    const cells = r.map((c) => (c.merged ? '' : cellText(c)))
    while (cells.length < width) cells.push('')
    return cells
  })
  const out = [`| ${grid[0].join(' | ')} |`, `|${' --- |'.repeat(width)}`]
  for (const r of grid.slice(1)) out.push(`| ${r.join(' | ')} |`)
  return out.join('\n')
}

export const markdown: FormatDef = {
  id: 'markdown',
  label: 'Markdown',
  group: 'text',
  blurb: 'Plain Markdown for notes and wikis, Marp or Pandoc/Quarto slides — the words rebuilt, the slide pictures alongside, or both.',
  ext: '.md + images',
  needs: { images: true },
  media: 'still',
  options: [
    {
      key: 'flavour', label: 'Flavour', type: 'select', default: 'gfm',
      choices: [
        { value: 'gfm', label: 'Markdown (GitHub-flavoured) document' },
        { value: 'marp', label: 'Marp slides' },
        { value: 'pandoc', label: 'Pandoc / Quarto slides' },
      ],
    },
    {
      key: 'content', label: 'Slides as', type: 'select', default: 'both',
      choices: [
        { value: 'text', label: 'Text only' },
        { value: 'pictures', label: 'Pictures only' },
        { value: 'both', label: 'Text and pictures' },
      ],
    },
    { key: 'notes', label: 'Speaker notes', type: 'toggle', default: true },
    {
      key: 'image', label: 'Pictures as', type: 'select', default: 'png',
      choices: [{ value: 'png', label: 'PNG' }, { value: 'jpeg', label: 'JPEG' }, { value: 'webp', label: 'WebP' }],
      when: (v) => v.content !== 'text',
    },
    { key: 'width', label: 'Picture width', type: 'number', min: 320, max: 7680, step: 10, unit: 'px', default: 1600, when: (v) => v.content !== 'text' },
    { key: 'crops', label: 'Charts and diagrams as pictures', type: 'toggle', default: true, when: (v) => v.content !== 'pictures' },
  ],
  async run(ctx, v) {
    await buildMarkdown(ctx, v)
  },
}

async function slideImage(ctx: ExportContext, s: Slide, v: Values): Promise<string> {
  const w = Number(v.width) || 1600
  const h = Math.round(w * (ctx.deck.height / ctx.deck.width))
  const f: Frame = { width: w, height: h, dx: 0, dy: 0, dw: w, dh: h, renderW: w, renderH: h, dpi: 96 }
  const n = String(s.index + 1).padStart(3, '0')
  if (v.image === 'png') {
    ctx.add(`slides/slide-${n}.png`, await ctx.images.framePng(s.index, f, '#ffffff'))
    return `slides/slide-${n}.png`
  }
  const c = await ctx.images.frame(s.index, f, '#ffffff')
  if (v.image === 'webp') {
    const bytes = (await canEncodeNatively('image/webp')) ? await encodeNative(c, 'image/webp', 0.88) : await encodeWebp(rgba(c), 88, false)
    ctx.add(`slides/slide-${n}.webp`, bytes)
    return `slides/slide-${n}.webp`
  }
  ctx.add(`slides/slide-${n}.jpg`, await encodeNative(c, 'image/jpeg', 0.88))
  return `slides/slide-${n}.jpg`
}

async function buildMarkdown(ctx: ExportContext, v: Values): Promise<void> {
  const flavour = String(v.flavour)
  const exported = new Set(ctx.slides.map((s) => s.index))
  const anchor = (i: number) => `slide-${i + 1}`
  const link = (l: Link): string | null => {
    if (l.url) return safeUrl(l.url)
    if (l.slide != null && exported.has(l.slide)) return flavour === 'gfm' ? `#${anchor(l.slide)}` : null
    return null
  }
  const out: string[] = []
  if (flavour === 'marp') {
    out.push('---', 'marp: true', `title: ${yamlString(ctx.deck.title)}`, 'paginate: true', `size: ${ctx.deck.width / ctx.deck.height < 1.5 ? '4:3' : '16:9'}`, '---', '')
  } else if (flavour === 'pandoc') {
    out.push('---', `title: ${yamlString(ctx.deck.title)}`, ...(ctx.deck.author ? [`author: ${yamlString(ctx.deck.author)}`] : []), `aspectratio: ${ctx.deck.width / ctx.deck.height < 1.5 ? 43 : 169}`, '---', '')
  } else {
    out.push(`# ${mdEscape(ctx.deck.title)}`, '')
  }

  let first = true
  await eachSlide(ctx, async (s) => {
    const title = s.title || `Slide ${s.index + 1}`
    if (flavour === 'marp' && !first) out.push('---', '')
    first = false
    if (flavour === 'gfm') out.push(`<a id="${anchor(s.index)}"></a>`, '', `## ${s.index + 1}. ${mdEscape(title)}`, '')
    else if (flavour === 'pandoc') out.push(`## ${mdEscape(title)} {#${anchor(s.index)}}`, '')
    else if (v.content !== 'pictures') out.push(`## ${mdEscape(title)}`, '')

    if (v.content !== 'text') {
      const src = await slideImage(ctx, s, v)
      if (flavour === 'marp' && v.content === 'pictures') out.push(`![bg contain](${src})`, '')
      else if (flavour === 'marp') out.push(`![bg right:40% contain](${src})`, '')
      else out.push(`![${mdEscape(title)}](${src})`, '')
    }
    if (v.content !== 'pictures') {
      const body = await mdBody(ctx, s, v, link)
      if (body) out.push(body, '')
    }
    if (v.notes && s.notesText.trim()) {
      if (flavour === 'marp') out.push(`<!--\n${s.notesText.replace(/-->/g, '—>')}\n-->`, '')
      else if (flavour === 'pandoc') out.push('::: notes', '', mdEscapeBlock(s.notesText), '', ':::', '')
      else out.push(mdEscapeBlock(s.notesText).split('\n').map((l) => (l ? `> ${l}` : '>')).join('\n'), '')
    }
  })
  ctx.add(`${ctx.base}.md`, out.join('\n').replace(/\n{3,}/g, '\n\n'))
}

function mdEscapeBlock(text: string): string {
  return text.split('\n').map((l) => mdEscape(l)).join('\n')
}

function yamlString(s: string): string {
  return JSON.stringify(s)
}

async function mdBody(ctx: ExportContext, s: Slide, v: Values, link: (l: Link) => string | null): Promise<string> {
  const st = structure(s)
  const parts: string[] = []
  if (st.subtitle?.paragraphs) parts.push(`*${mdParagraphs(st.subtitle.paragraphs, link)}*`)
  for (const col of st.columns) for (const b of col) parts.push(await mdBlock(ctx, s, b, v, link))
  return parts.filter(Boolean).join('\n\n')
}

async function mdBlock(ctx: ExportContext, s: Slide, b: Block, v: Values, link: (l: Link) => string | null): Promise<string> {
  const sh: Shape = b.shape
  switch (b.kind) {
    case 'text':
      return mdParagraphs(sh.paragraphs ?? [], link)
    case 'table':
      return mdTable(sh.table!.rows, link)
    case 'image': {
      const ext = passThroughImage(sh.image!.contentType)
      const bytes = ctx.deck.readPart(sh.image!.part)
      if (bytes && ext) {
        const name = `media/${sh.image!.part.split('/').pop()}`
        ctx.add(name, bytes)
        return `![${mdEscape(sh.description ?? '')}](${name})`
      }
      return mdCrop(ctx, s, b, v)
    }
    case 'crop':
      return mdCrop(ctx, s, b, v)
  }
}

async function mdCrop(ctx: ExportContext, s: Slide, b: Block, v: Values): Promise<string> {
  if (!v.crops || !b.shape.rect) return ''
  const name = `media/slide${s.index + 1}-shape${b.shape.id}.png`
  ctx.add(name, await cropRegion(ctx, s, b.shape.rect, Number(v.width) || 1600))
  return `![${mdEscape(b.shape.description ?? b.shape.name ?? '')}](${name})`
}

// --- Outline -------------------------------------------------------------------------

interface OutlineItem {
  level: number
  text: string
}

/** Title at level 0, then the body text one level down per bullet level. */
export function outlineOf(s: Slide): OutlineItem[] {
  const items: OutlineItem[] = [{ level: 0, text: s.title || `Slide ${s.index + 1}` }]
  const st = structure(s)
  if (st.subtitle?.paragraphs) {
    for (const p of st.subtitle.paragraphs) {
      const t = p.runs.map((r) => (r.br ? ' ' : r.text)).join('').trim()
      if (t) items.push({ level: 1, text: t })
    }
  }
  for (const col of st.columns) {
    for (const b of col) {
      if (b.kind === 'text') {
        for (const p of b.shape.paragraphs ?? []) {
          const t = p.runs.map((r) => (r.br ? ' ' : r.text)).join('').trim()
          if (t) items.push({ level: Math.min(8, p.level + 1), text: t })
        }
      } else if (b.kind === 'table') {
        for (const row of b.shape.table!.rows) {
          const t = row.filter((c) => !c.merged).map((c) => paragraphsText(c.paragraphs).replace(/\n/g, ' ').trim()).join(' · ')
          if (t.replace(/[ ·]/g, '')) items.push({ level: 1, text: t })
        }
      }
    }
  }
  return items
}

/** RTF with PowerPoint's outline levels: Insert ▸ Slides from Outline reads it back. */
export function outlineRtf(items: OutlineItem[]): string {
  const esc = (t: string) => {
    let o = ''
    for (const ch of t) {
      const cp = ch.codePointAt(0)!
      if (ch === '\\' || ch === '{' || ch === '}') o += `\\${ch}`
      else if (cp < 0x80) o += ch
      else if (cp < 0x10000) o += `\\u${cp > 32767 ? cp - 65536 : cp}?`
      else {
        // Outside the BMP: RTF wants the UTF-16 surrogate pair.
        const hi = 0xd800 + ((cp - 0x10000) >> 10)
        const lo = 0xdc00 + ((cp - 0x10000) & 0x3ff)
        o += `\\u${hi - 65536}?\\u${lo - 65536}?`
      }
    }
    return o
  }
  const styles = Array.from({ length: 9 }, (_, i) => `{\\s${i + 1}\\outlinelevel${i} heading ${i + 1};}`).join('')
  const body = items.map((it) => `\\pard\\plain\\s${it.level + 1}\\outlinelevel${it.level}\\li${it.level * 360}${it.level === 0 ? '\\b\\fs32' : '\\fs24'} ${esc(it.text)}\\par`).join('\n')
  return `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1\n{\\fonttbl{\\f0\\fswiss Calibri;}}\n{\\stylesheet${styles}}\n${body}\n}\n`
}

export function outlineOpml(title: string, slides: OutlineItem[][]): string {
  const lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<opml version="2.0">', `  <head><title>${xmlEscape(xmlSafe(title))}</title></head>`, '  <body>']
  for (const items of slides) {
    const stack: number[] = []
    for (const it of items) {
      while (stack.length && stack[stack.length - 1] >= it.level) {
        stack.pop()
        lines.push(`${'  '.repeat(stack.length + 2)}</outline>`)
      }
      lines.push(`${'  '.repeat(stack.length + 2)}<outline text="${xmlEscape(xmlSafe(it.text))}">`)
      stack.push(it.level)
    }
    while (stack.length) {
      stack.pop()
      lines.push(`${'  '.repeat(stack.length + 2)}</outline>`)
    }
  }
  lines.push('  </body>', '</opml>', '')
  // Self-close the leaves.
  return lines.join('\n').replace(/<outline text="([^"]*)">\n\s*<\/outline>/g, '<outline text="$1"/>')
}

export const outline: FormatDef = {
  id: 'outline',
  label: 'Outline',
  group: 'text',
  blurb: "The deck's words as an outline: tab-indented text, RTF that PowerPoint reads straight back in, or OPML for outliners.",
  ext: '.txt / .rtf / .opml',
  needs: {},
  media: 'none',
  options: [
    {
      key: 'as', label: 'Save as', type: 'select', default: 'rtf',
      choices: [
        { value: 'rtf', label: 'RTF (PowerPoint outline)' },
        { value: 'txt', label: 'Text, tab-indented' },
        { value: 'opml', label: 'OPML' },
      ],
    },
  ],
  async run(ctx, v) {
    const all = ctx.slides.map(outlineOf)
    if (v.as === 'rtf') ctx.add(`${ctx.base}-outline.rtf`, outlineRtf(all.flat()))
    else if (v.as === 'opml') ctx.add(`${ctx.base}-outline.opml`, outlineOpml(ctx.deck.title, all))
    else ctx.add(`${ctx.base}-outline.txt`, all.map((items) => items.map((it) => `${'\t'.repeat(it.level)}${it.text}`).join('\n')).join('\n\n') + '\n')
  },
}

// --- Speaker notes -------------------------------------------------------------------

export const speakerNotes: FormatDef = {
  id: 'notes',
  label: 'Speaker notes',
  group: 'text',
  blurb: 'Just the notes, slide by slide — a script to rehearse from, or hand to the autocue.',
  ext: '.txt / .md',
  needs: {},
  media: 'none',
  options: [
    { key: 'as', label: 'Save as', type: 'select', default: 'txt', choices: [{ value: 'txt', label: 'Plain text' }, { value: 'md', label: 'Markdown' }] },
    { key: 'titles', label: 'Slide numbers and titles', type: 'toggle', default: true },
    { key: 'slideText', label: 'Each slide’s own text too', type: 'toggle', default: false },
    { key: 'skipEmpty', label: 'Leave out slides with no notes', type: 'toggle', default: false },
  ],
  async run(ctx, v) {
    const md = v.as === 'md'
    const chunks: string[] = []
    if (md) chunks.push(`# ${mdEscape(ctx.deck.title)} — speaker notes`)
    for (const s of ctx.slides) {
      if (v.skipEmpty && !s.notesText.trim()) continue
      const part: string[] = []
      const head = `${s.index + 1}. ${s.title || 'Untitled'}${s.hidden ? ' (hidden)' : ''}`
      if (v.titles) part.push(md ? `## ${mdEscape(head)}` : `${head}\n${'-'.repeat(Math.min(72, [...head].length))}`)
      if (v.slideText) {
        const t = slideText(s)
        if (t) part.push(md ? t.split('\n').map((l) => `> ${mdEscape(l)}`).join('\n') : t.split('\n').map((l) => `  | ${l}`).join('\n'))
      }
      part.push(s.notesText.trim() ? (md ? mdEscapeBlock(s.notesText) : s.notesText) : md ? '*No notes.*' : '(no notes)')
      chunks.push(part.join('\n\n'))
    }
    ctx.add(`${ctx.base}-notes.${md ? 'md' : 'txt'}`, chunks.join(md ? '\n\n' : '\n\n\n') + '\n')
  },
}

// --- CSV and JSON ----------------------------------------------------------------------

export function csvField(v: string | number | boolean, sep: string): string {
  const s = String(v)
  return s.includes(sep) || /["\n\r]/.test(s) || /^\s|\s$/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
}

export const csv: FormatDef = {
  id: 'csv',
  label: 'CSV',
  group: 'text',
  blurb: 'One row per slide — number, title, section, hidden, timings, word count, notes — for a running order or a spreadsheet.',
  ext: '.csv',
  needs: {},
  media: 'none',
  options: [
    { key: 'sep', label: 'Separator', type: 'select', default: ',', choices: [{ value: ',', label: 'Comma' }, { value: ';', label: 'Semicolon (European Excel)' }, { value: '\t', label: 'Tab' }] },
    { key: 'notes', label: 'Notes column', type: 'toggle', default: true },
    { key: 'text', label: 'Slide text column', type: 'toggle', default: false },
    { key: 'bom', label: 'Mark as UTF-8 for Excel', type: 'toggle', default: true, help: 'Adds a byte-order mark so Excel reads accents and non-Latin text correctly.' },
  ],
  async run(ctx, v) {
    const sep = String(v.sep)
    const head = ['Slide', 'Title', 'Section', 'Hidden', 'Layout', 'Transition', 'Advance after (s)', 'Words']
    if (v.text) head.push('Text')
    if (v.notes) head.push('Notes')
    const rows = [head]
    for (const s of ctx.slides) {
      const text = slideText(s)
      const row: (string | number)[] = [
        s.index + 1, s.title, s.section ?? '', s.hidden ? 'yes' : '', s.layoutName ?? '',
        s.transition && s.transition.type !== 'none' ? s.transition.type : '',
        s.transition?.advanceAfterMs != null ? s.transition.advanceAfterMs / 1000 : '',
        (text.match(/[\p{L}\p{N}]+/gu) ?? []).length,
      ]
      if (v.text) row.push(text)
      if (v.notes) row.push(s.notesText)
      rows.push(row.map(String))
    }
    const body = rows.map((r) => r.map((c) => csvField(c, sep)).join(sep)).join('\r\n') + '\r\n'
    ctx.add(`${ctx.base}.csv`, (v.bom ? '\ufeff' : '') + body)
  },
}

export const json: FormatDef = {
  id: 'json',
  label: 'JSON',
  group: 'text',
  blurb: 'Everything the exporter read from the deck — text with formatting, notes, links, geometry, transitions — for scripts.',
  ext: '.json',
  needs: {},
  media: 'none',
  options: [
    { key: 'geometry', label: 'Shape positions', type: 'toggle', default: true },
    { key: 'rich', label: 'Formatting (runs, levels)', type: 'toggle', default: true },
    { key: 'pretty', label: 'Indented', type: 'toggle', default: true },
  ],
  async run(ctx, v) {
    const emu = { width: ctx.deck.width, height: ctx.deck.height }
    const paras = (ps: Paragraph[]) => (v.rich ? ps : paragraphsText(ps))
    const doc = {
      format: 'advanced-exporter-deck',
      version: 1,
      generator: `Advanced Exporter ${__APP_VERSION__}`,
      title: ctx.deck.title,
      author: ctx.deck.author ?? null,
      size: { emu, inches: { width: emu.width / 914400, height: emu.height / 914400 } },
      show: ctx.deck.show,
      sections: ctx.deck.sections,
      slides: ctx.slides.map((s) => ({
        number: s.index + 1,
        title: s.title,
        hidden: s.hidden,
        section: s.section ?? null,
        layout: s.layoutName ?? null,
        transition: s.transition ?? null,
        hasAnimations: s.hasAnimations,
        notes: v.rich ? s.notes : s.notesText,
        shapes: s.shapes.map((sh) => ({
          id: sh.id,
          name: sh.name,
          kind: sh.kind,
          placeholder: sh.placeholder ?? null,
          ...(v.geometry ? { rect: sh.rect ?? null, rotation: sh.rotation ?? 0 } : {}),
          ...(sh.paragraphs ? { text: paras(sh.paragraphs) } : {}),
          ...(sh.table ? { table: sh.table.rows.map((r) => r.map((c) => (c.merged ? null : { text: paras(c.paragraphs), colSpan: c.colSpan, rowSpan: c.rowSpan }))) } : {}),
          ...(sh.image ? { image: sh.image } : {}),
          ...(sh.media ? { media: sh.media } : {}),
          ...(sh.link ? { link: sh.link } : {}),
          ...(sh.description ? { description: sh.description } : {}),
        })),
      })),
    }
    ctx.add(`${ctx.base}.json`, JSON.stringify(doc, null, v.pretty ? 2 : undefined) + '\n')
  },
}

export const TEXT_FORMATS: FormatDef[] = [markdown, outline, speakerNotes, csv, json]
