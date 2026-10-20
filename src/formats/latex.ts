// LaTeX: a Beamer deck (or an article) that compiles.
//
// Two ways to get there, because they're for different people:
//   pictures   — every frame is the slide exactly as PowerPoint drew it (its
//                vector PDF page, or a picture), links kept as hotspots. For
//                when the slides must look right inside a LaTeX workflow.
//   structured — the slides rebuilt as Beamer: titles, nested bullets, bold
//                and italic, links, tables, pictures, columns. For when you
//                mean to keep editing them in LaTeX. Anything that can't be
//                rebuilt (charts, SmartArt, drawn shapes) comes across as a
//                picture of its own patch of the slide.

import type { Paragraph, Run, Slide } from '../deck/model'
import { EMU_PER_INCH } from '../deck/model'
import { encodeNative } from '../render/canvas'
import { mapPdf, pageForSlide } from '../render/pdfmap'
import type { Frame } from '../render/sizing'
import { eachSlide } from './common'
import { cropRegion, passThroughImage } from './crop'
import { hotspots, safeUrl } from './overlay'
import { structure, widthFraction, type Block } from './structure'
import type { ExportContext, FormatDef, Values } from './types'

// --- Escaping -------------------------------------------------------------------

const SPECIAL: Record<string, string> = {
  '\\': '\\textbackslash{}', '{': '\\{', '}': '\\}', $: '\\$', '&': '\\&', '#': '\\#', '%': '\\%',
  _: '\\_', '^': '\\textasciicircum{}', '~': '\\textasciitilde{}', '<': '\\textless{}', '>': '\\textgreater{}',
  '|': '\\textbar{}',
}

// pdfLaTeX with inputenc knows Latin-1 and most of Latin Extended-A; these
// others turn up in slides often enough to map by hand.
const PDFLATEX_MAP: Record<string, string> = {
  '✓': '\\checkmark{}', '✔': '\\checkmark{}', '→': '\\textrightarrow{}', '←': '\\textleftarrow{}',
  '•': '\\textbullet{}', '…': '\\ldots{}', '™': '\\texttrademark{}', '©': '\\textcopyright{}',
  '®': '\\textregistered{}', '°': '\\textdegree{}', '±': '\\textpm{}', '×': '\\texttimes{}',
  '÷': '\\textdiv{}', '€': '\\texteuro{}', '≤': '\\ensuremath{\\leq}', '≥': '\\ensuremath{\\geq}',
  '≠': '\\ensuremath{\\neq}', '≈': '\\ensuremath{\\approx}', '∞': '\\ensuremath{\\infty}',
  ' ': '~', ' ': '\\,', '​': '',
}

export type Engine = 'xelatex' | 'lualatex' | 'pdflatex'

export function texEscape(text: string, engine: Engine, unsupported?: Set<string>): string {
  let out = ''
  for (const ch of text) {
    if (SPECIAL[ch]) out += SPECIAL[ch]
    else if (engine === 'pdflatex') {
      const cp = ch.codePointAt(0)!
      if (PDFLATEX_MAP[ch] != null) out += PDFLATEX_MAP[ch]
      else if (cp < 0x180 || (cp >= 0x2010 && cp <= 0x201e) || ch === '\n') out += ch
      else {
        unsupported?.add(ch)
        out += '?'
      }
    } else if (ch === ' ') out += '~'
    else out += ch
  }
  // Quotes and dashes LaTeX would otherwise ligature or mangle.
  return out.replace(/\r/g, '')
}

const CJK = /[぀-ヿ㐀-䶿一-鿿가-힯豈-﫿]/

// --- Options --------------------------------------------------------------------

const THEMES = ['default', 'Madrid', 'Boadilla', 'CambridgeUS', 'Singapore', 'Warsaw', 'metropolis']

export const latex: FormatDef = {
  id: 'latex',
  label: 'LaTeX (Beamer)',
  group: 'text',
  blurb: 'A Beamer deck that compiles: either the slides exactly as drawn, or rebuilt as editable LaTeX. Or an article with the notes.',
  ext: '.tex + files',
  needs: {},
  needsFor: (v) => (v.mode === 'pictures' && v.pictures === 'pdf' ? { pdf: true, images: true } : { images: true }),
  options: [
    {
      key: 'mode', label: 'Slides as', type: 'select', default: 'structured',
      choices: [
        { value: 'structured', label: 'Editable LaTeX (rebuilt)' },
        { value: 'pictures', label: 'Pictures of the slides (exact)' },
      ],
    },
    {
      key: 'pictures', label: 'Pictures from', type: 'select', default: 'pdf',
      choices: [{ value: 'pdf', label: "PowerPoint's PDF (vector)" }, { value: 'png', label: 'PNG files' }],
      when: (v) => v.mode === 'pictures',
      help: "Hidden slides aren't in PowerPoint's PDF; any being exported come across as PNGs.",
    },
    {
      key: 'docclass', label: 'Document', type: 'select', default: 'beamer',
      choices: [{ value: 'beamer', label: 'Beamer slides' }, { value: 'article', label: 'Article (slides with notes, for print)' }],
    },
    {
      key: 'engine', label: 'Engine', type: 'select', default: 'xelatex',
      choices: [{ value: 'xelatex', label: 'XeLaTeX' }, { value: 'lualatex', label: 'LuaLaTeX' }, { value: 'pdflatex', label: 'pdfLaTeX' }],
      help: 'XeLaTeX and LuaLaTeX take any script. pdfLaTeX is the oldest and pickiest: characters it has no way to set become ?.',
    },
    { key: 'mainfont', label: 'Main font', type: 'text', default: '', placeholder: 'e.g. Noto Sans', when: (v) => v.engine !== 'pdflatex', help: 'Empty keeps Latin Modern. Pick one you have installed that covers your text.' },
    { key: 'cjkfont', label: 'CJK font', type: 'text', default: '', placeholder: 'e.g. Hiragino Sans', when: (v) => v.engine !== 'pdflatex', help: 'Only matters if the deck has Chinese, Japanese or Korean text. Empty lets xeCJK use TeX Live’s own (Fandol), which compiles anywhere.' },
    { key: 'theme', label: 'Beamer theme', type: 'select', default: 'default', choices: THEMES.map((t) => ({ value: t, label: t })), when: (v) => v.docclass === 'beamer' && v.mode === 'structured' },
    {
      key: 'notes', label: 'Speaker notes', type: 'select', default: 'notes',
      choices: [
        { value: 'none', label: 'Leave out' },
        { value: 'notes', label: '\\note{} (hidden unless you ask)' },
        { value: 'second', label: 'Notes on a second screen' },
      ],
      when: (v) => v.docclass === 'beamer',
    },
    { key: 'articleNotes', label: 'Include speaker notes', type: 'toggle', default: true, when: (v) => v.docclass === 'article' },
    { key: 'sections', label: 'Sections from the deck', type: 'toggle', default: true },
    { key: 'links', label: 'Keep links', type: 'toggle', default: true },
    { key: 'crops', label: 'Charts and diagrams as pictures', type: 'toggle', default: true, when: (v) => v.mode === 'structured' },
    { key: 'resolution', label: 'Picture width', type: 'number', min: 640, max: 7680, step: 10, unit: 'px', default: 1920 },
  ],
  async run(ctx, v) {
    await buildLatex(ctx, v)
  },
}

// --- Build ---------------------------------------------------------------------

interface Build {
  ctx: ExportContext
  v: Values
  engine: Engine
  unsupported: Set<string>
  files: Map<string, Uint8Array | string>
  /** Slides in the export, by deck index — for links between them. */
  exported: Set<number>
}

async function buildLatex(ctx: ExportContext, v: Values): Promise<void> {
  const b: Build = {
    ctx, v,
    engine: v.engine as Engine,
    unsupported: new Set(),
    files: new Map(),
    exported: new Set(ctx.slides.map((s) => s.index)),
  }
  const body: string[] = []
  let pdfPages: ((i: number) => number | null) | null = null
  if (v.mode === 'pictures' && v.pictures === 'pdf') {
    ctx.progress(0, 'Asking PowerPoint for its PDF')
    const bytes = await ctx.pdf()
    const { PDFDocument } = await import('pdf-lib')
    const n = (await PDFDocument.load(bytes, { updateMetadata: false, ignoreEncryption: true })).getPageCount()
    const hidden = ctx.deck.slides.map((s) => s.hidden)
    const map = mapPdf(n, hidden)
    if (map.kind === 'mismatch') ctx.warn(`PowerPoint's PDF has ${n} pages for ${hidden.length} slides; using PNGs instead.`)
    else {
      b.files.set('slides.pdf', bytes)
      pdfPages = (i) => pageForSlide(map, hidden, i)
    }
  }

  let section: string | undefined
  await eachSlide(ctx, async (s) => {
    if (v.sections && s.section && s.section !== section) {
      section = s.section
      body.push(`\\section{${tex(b, s.section)}}`, '')
    }
    if (v.docclass === 'article') body.push(await articleSlide(b, s, pdfPages))
    else if (v.mode === 'pictures') body.push(await pictureFrame(b, s, pdfPages))
    else body.push(await structuredFrame(b, s))
    body.push('')
  })

  const anyCjk = ctx.slides.some((s) => CJK.test(s.title + s.notesText + s.shapes.map((x) => x.paragraphs?.map((p) => p.runs.map((r) => r.text).join('')).join('') ?? '').join('')))
  const main = [preamble(b, anyCjk), '\\begin{document}', '', ...titlePage(b), ...body, '\\end{document}', ''].join('\n')
  b.files.set('main.tex', main)
  b.files.set('latexmkrc', latexmkrc(b.engine))
  b.files.set('README.txt', readme(b))
  if (b.unsupported.size) {
    ctx.warn(`pdfLaTeX can't set ${[...b.unsupported].slice(0, 12).join(' ')}${b.unsupported.size > 12 ? ' …' : ''} — replaced with ?. Choose XeLaTeX to keep them.`)
  }
  for (const [path, data] of b.files) ctx.add(path, data)
}

function tex(b: Build, text: string): string {
  return texEscape(text, b.engine, b.unsupported)
}

function aspectOption(w: number, h: number): string | null {
  const r = w / h
  const known: [number, string][] = [[16 / 9, '169'], [16 / 10, '1610'], [4 / 3, '43'], [3 / 2, '32'], [5 / 4, '54'], [14 / 9, '149'], [1.41, '141']]
  for (const [v, opt] of known) if (Math.abs(r - v) < 0.01) return opt
  return null
}

function preamble(b: Build, anyCjk: boolean): string {
  const { ctx, v } = b
  const lines: string[] = [`% !TEX program = ${b.engine}`, `% ${ctx.deck.title} — exported from PowerPoint by Advanced Exporter`]
  const aspect = aspectOption(ctx.deck.width, ctx.deck.height)
  if (v.docclass === 'article') {
    lines.push('\\documentclass[11pt,a4paper]{article}', '\\usepackage[margin=2cm]{geometry}')
  } else {
    lines.push(`\\documentclass[${aspect ? `aspectratio=${aspect}` : ''}${v.notes === 'second' ? `${aspect ? ',' : ''}notes` : ''}]{beamer}`)
    if (!aspect) {
      const wmm = (ctx.deck.width / EMU_PER_INCH) * 25.4
      const hmm = (ctx.deck.height / EMU_PER_INCH) * 25.4
      // Beamer's frames are about 128 mm wide; scale the slide to match.
      const k = 128 / wmm
      lines.push(`\\geometry{papersize={${(wmm * k).toFixed(1)}mm,${(hmm * k).toFixed(1)}mm}}`)
    }
    if (v.mode === 'structured' && v.theme !== 'default') lines.push(`\\usetheme{${v.theme}}`)
    if (v.mode === 'pictures') lines.push('\\setbeamertemplate{navigation symbols}{}')
    if (v.notes === 'second') lines.push('\\usepackage{pgfpages}', '\\setbeameroption{show notes on second screen=right}')
  }
  if (b.engine === 'pdflatex') {
    lines.push('\\usepackage[utf8]{inputenc}', '\\usepackage[T1]{fontenc}', '\\usepackage{lmodern}', '\\usepackage{textcomp}', '\\usepackage{amssymb}')
  } else {
    lines.push('\\usepackage{fontspec}')
    if (String(v.mainfont).trim()) lines.push(`\\setmainfont{${String(v.mainfont).trim()}}`, `\\setsansfont{${String(v.mainfont).trim()}}`)
    if (anyCjk) {
      const cjk = String(v.cjkfont).trim()
      lines.push(b.engine === 'xelatex' ? '\\usepackage{xeCJK}' : '\\usepackage{luatexja-fontspec}')
      // xeCJK files dashes, curly quotes and the ellipsis as CJK punctuation and
      // eats the spaces round them in Western text ("Fixture —généré").
      if (b.engine === 'xelatex') lines.push('\\xeCJKDeclareCharClass{Default}{"2013,"2014,"2018,"2019,"201C,"201D,"2026,"00B7}')
      if (cjk) {
        lines.push(b.engine === 'xelatex' ? `\\setCJKmainfont{${cjk}}` : `\\setmainjfont{${cjk}}`)
        lines.push(b.engine === 'xelatex' ? `\\setCJKsansfont{${cjk}}` : `\\setsansjfont{${cjk}}`)
      }
    }
  }
  lines.push('\\usepackage{graphicx}', '\\usepackage{booktabs}', '\\usepackage[normalem]{ulem}')
  if (v.mode === 'pictures' && v.links) lines.push('\\usepackage{tikz}')
  // Beamer loads hyperref itself; an article needs it.
  if (v.docclass === 'article') lines.push('\\usepackage[colorlinks=true,urlcolor=blue,linkcolor=black]{hyperref}')
  lines.push(`\\title{${tex(b, ctx.deck.title)}}`)
  if (ctx.deck.author) lines.push(`\\author{${tex(b, ctx.deck.author)}}`)
  lines.push('\\date{}', '')
  return lines.join('\n')
}

function titlePage(b: Build): string[] {
  if (b.v.docclass === 'article') return ['\\maketitle', '']
  return []
}

function latexmkrc(engine: Engine): string {
  const mode = engine === 'xelatex' ? 5 : engine === 'lualatex' ? 4 : 1
  return `# latexmk picks the engine from this: run \`latexmk main.tex\`\n$pdf_mode = ${mode};\n`
}

function readme(b: Build): string {
  return [
    `${b.ctx.deck.title}`,
    '',
    `Exported from PowerPoint by Advanced Exporter as ${b.v.docclass === 'article' ? 'a LaTeX article' : 'Beamer slides'} (${b.v.mode === 'pictures' ? 'pictures of the slides' : 'rebuilt as LaTeX'}).`,
    '',
    'To build it:',
    `    latexmk main.tex          (uses latexmkrc: ${b.engine})`,
    `or  ${b.engine} main.tex      (run twice so links and sections settle)`,
    '',
    b.engine === 'pdflatex' ? '' : 'XeLaTeX and LuaLaTeX need the fonts named in main.tex installed.',
    '',
  ].join('\n')
}

// --- Pictures mode ---------------------------------------------------------------

async function slidePicture(b: Build, s: Slide, pdfPages: ((i: number) => number | null) | null): Promise<string> {
  const page = pdfPages?.(s.index)
  if (page != null) return `[page=${page + 1}]{slides.pdf}`
  const w = Number(b.v.resolution) || 1920
  const h = Math.round(w * (b.ctx.deck.height / b.ctx.deck.width))
  const f: Frame = { width: w, height: h, dx: 0, dy: 0, dw: w, dh: h, renderW: w, renderH: h, dpi: 96 }
  const name = `slides/slide-${String(s.index + 1).padStart(3, '0')}.png`
  b.files.set(name, await b.ctx.images.framePng(s.index, f, '#ffffff'))
  return `{${name}}`
}

async function pictureFrame(b: Build, s: Slide, pdfPages: ((i: number) => number | null) | null): Promise<string> {
  const pic = await slidePicture(b, s, pdfPages)
  const out = [
    `% Slide ${s.index + 1}${s.title ? `: ${s.title.replace(/\n/g, ' ')}` : ''}`,
    '{',
    `\\setbeamertemplate{background}{\\includegraphics[width=\\paperwidth,height=\\paperheight]${pic}}`,
    `\\begin{frame}[plain,label=slide-${s.index + 1}]`,
  ]
  if (b.v.links) {
    const spots = hotspots(b.ctx.deck, s).filter((h) => (h.link.url && safeUrl(h.link.url)) || (h.link.slide != null && b.exported.has(h.link.slide)))
    if (spots.length) {
      out.push('\\begin{tikzpicture}[remember picture,overlay]')
      for (const h of spots) {
        const target = h.link.url ? `\\href{${urlEscape(h.link.url)}}` : `\\hyperlink{slide-${h.link.slide! + 1}}`
        out.push(`  \\node[anchor=north west,inner sep=0pt] at ([xshift=${(h.x).toFixed(4)}\\paperwidth,yshift=-${(h.y).toFixed(4)}\\paperheight]current page.north west) {${target}{\\rule{0pt}{${(h.h).toFixed(4)}\\paperheight}\\rule{${(h.w).toFixed(4)}\\paperwidth}{0pt}}};`)
      }
      out.push('\\end{tikzpicture}')
    }
  }
  out.push('\\end{frame}', '}')
  out.push(...notesFor(b, s))
  return out.join('\n')
}

function urlEscape(url: string): string {
  return url.replace(/\\/g, '/').replace(/[%#]/g, (c) => `\\${c}`).replace(/[{}]/g, '')
}

function notesFor(b: Build, s: Slide): string[] {
  if (b.v.notes === 'none' || !s.notesText.trim()) return []
  return [`\\note{${notesTex(b, s)}}`]
}

function notesTex(b: Build, s: Slide): string {
  return s.notesText
    .split(/\n{2,}/)
    .map((para) => para.split('\n').map((l) => tex(b, l)).join('\\\\\n'))
    .join('\n\n')
}

// --- Structured mode ---------------------------------------------------------------

async function structuredFrame(b: Build, s: Slide): Promise<string> {
  const st = structure(s)
  const out: string[] = [`% Slide ${s.index + 1}`]
  if (st.isTitleSlide && s.index === b.ctx.slides[0]?.index) {
    out.push(`\\begin{frame}[label=slide-${s.index + 1}]`)
    if (st.title) out.push(`\\title{${runsInline(b, st.title.paragraphs ?? [])}}`)
    if (st.subtitle) out.push(`\\subtitle{${runsInline(b, st.subtitle.paragraphs ?? [])}}`)
    out.push('\\titlepage')
  } else {
    const title = st.title?.paragraphs ? runsInline(b, st.title.paragraphs) : ''
    out.push(`\\begin{frame}[label=slide-${s.index + 1}]${title ? `{${title}}` : ''}`)
    if (st.subtitle) out.push(`\\framesubtitle{${runsInline(b, st.subtitle.paragraphs ?? [])}}`)
  }
  const cols = st.columns.filter((c) => c.length)
  if (cols.length > 1) {
    out.push('\\begin{columns}[T]')
    const w = (0.96 / cols.length).toFixed(2)
    for (const col of cols) {
      out.push(`\\begin{column}{${w}\\textwidth}`)
      for (const block of col) out.push(await blockTex(b, s, block, cols.length))
      out.push('\\end{column}')
    }
    out.push('\\end{columns}')
  } else if (cols.length === 1) {
    for (const block of cols[0]) out.push(await blockTex(b, s, block, 1))
  }
  out.push('\\end{frame}')
  out.push(...notesFor(b, s))
  return out.filter((l) => l !== '').join('\n')
}

async function blockTex(b: Build, s: Slide, block: Block, columns: number): Promise<string> {
  const sh = block.shape
  switch (block.kind) {
    case 'text':
      return paragraphsTex(b, sh.paragraphs ?? [])
    case 'table':
      return tableTex(b, sh.table!)
    case 'image': {
      const ext = passThroughImage(sh.image!.contentType)
      const bytes = b.ctx.deck.readPart(sh.image!.part)
      if (!bytes || !ext || ext === 'gif') return cropTex(b, s, block)
      const name = `media/${sh.image!.part.split('/').pop()!.replace(/\.[^.]+$/, '')}.${ext}`
      b.files.set(name, bytes)
      return figure(name, widthFraction(sh, b.ctx.deck.width) * columns)
    }
    case 'crop':
      return cropTex(b, s, block)
  }
}

async function cropTex(b: Build, s: Slide, block: Block): Promise<string> {
  const sh = block.shape
  if (!b.v.crops || !sh.rect) return `% (${sh.kind}${sh.name ? ` “${sh.name}”` : ''} left out)`
  const name = `media/slide${s.index + 1}-shape${sh.id}.png`
  b.files.set(name, await cropRegion(b.ctx, s, sh.rect, Number(b.v.resolution) || 1920))
  return figure(name, widthFraction(sh, b.ctx.deck.width))
}

function figure(path: string, fraction: number): string {
  return `\\begin{center}\\includegraphics[width=${Math.min(1, fraction).toFixed(2)}\\linewidth,height=0.7\\textheight,keepaspectratio]{${path}}\\end{center}`
}

/** Paragraphs → itemize/enumerate, nested by level; plain paragraphs as text. */
export function paragraphsTex(b: Pick<Build, 'engine' | 'unsupported' | 'v' | 'exported'>, ps: Paragraph[]): string {
  const out: string[] = []
  const stack: { env: 'itemize' | 'enumerate'; level: number }[] = []
  const close = (toLevel: number) => {
    while (stack.length && stack[stack.length - 1].level >= toLevel) out.push(`${'  '.repeat(stack.length - 1)}\\end{${stack.pop()!.env}}`)
  }
  for (const p of ps) {
    const text = runsInline(b, [p])
    if (!text.trim()) continue
    if (p.bullet === 'none') {
      close(0)
      out.push(`${text}\\par`)
      continue
    }
    const env = p.bullet === 'number' ? 'enumerate' : 'itemize'
    close(p.level + 1)
    const top = stack[stack.length - 1]
    if (!top || top.level < p.level || top.env !== env) {
      if (top && top.level === p.level && top.env !== env) close(p.level)
      stack.push({ env, level: p.level })
      out.push(`${'  '.repeat(stack.length - 1)}\\begin{${env}}`)
    }
    out.push(`${'  '.repeat(stack.length)}\\item ${text}`)
  }
  close(0)
  return out.join('\n')
}

/** Runs of one or more paragraphs as inline LaTeX (titles join with spaces). */
export function runsInline(b: Pick<Build, 'engine' | 'unsupported' | 'v' | 'exported'>, ps: Paragraph[]): string {
  return ps.map((p) => p.runs.map((r) => runTex(b, r)).join('')).join(' ').replace(/\s+$/, '')
}

function runTex(b: Pick<Build, 'engine' | 'unsupported' | 'v' | 'exported'>, r: Run): string {
  if (r.br) return '\\newline '
  let t = texEscape(r.text, b.engine, b.unsupported)
  if (!t) return ''
  if (r.mono) t = `\\texttt{${t}}`
  if (r.bold) t = `\\textbf{${t}}`
  if (r.italic) t = `\\emph{${t}}`
  if (r.underline && !r.link) t = `\\uline{${t}}`
  if (r.strike) t = `\\sout{${t}}`
  if (r.baseline && r.baseline > 0) t = `\\textsuperscript{${t}}`
  if (r.baseline && r.baseline < 0) t = `\\textsubscript{${t}}`
  if (r.link && b.v.links !== false) {
    if (r.link.url && safeUrl(r.link.url)) t = `\\href{${urlEscape(r.link.url)}}{${t}}`
    else if (r.link.slide != null && b.exported.has(r.link.slide)) t = `\\hyperlink{slide-${r.link.slide + 1}}{${t}}`
  }
  return t
}

function tableTex(b: Build, t: NonNullable<import('../deck/model').Shape['table']>): string {
  const cols = Math.max(t.colWidths.length, ...t.rows.map((r) => r.reduce((n, c) => n + (c.merged ? 0 : c.colSpan), 0)))
  const lines = [`\\begin{center}\\small`, `\\begin{tabular}{${'l'.repeat(cols)}}`, '\\toprule']
  t.rows.forEach((row, ri) => {
    const cells: string[] = []
    let ci = 0
    while (ci < row.length) {
      const c = row[ci]
      if (c.merged) {
        // Covered by a span to its left (already emitted) or from above (leave empty).
        if (!row.slice(0, ci).some((p, k) => !p.merged && k + p.colSpan > ci)) cells.push('')
        ci++
        continue
      }
      const text = c.paragraphs.map((p) => p.runs.map((r) => runTex(b, r)).join('')).filter(Boolean).join(' \\newline ')
      cells.push(c.colSpan > 1 ? `\\multicolumn{${c.colSpan}}{l}{${text}}` : text)
      ci += Math.max(1, c.colSpan)
    }
    lines.push(`${cells.join(' & ')} \\\\`)
    if (ri === 0 && t.rows.length > 1) lines.push('\\midrule')
  })
  lines.push('\\bottomrule', '\\end{tabular}', '\\end{center}')
  return lines.join('\n')
}

// --- Article ---------------------------------------------------------------------

async function articleSlide(b: Build, s: Slide, pdfPages: ((i: number) => number | null) | null): Promise<string> {
  const page = pdfPages?.(s.index)
  let pic: string
  if (page != null) pic = `[page=${page + 1},width=\\linewidth]{slides.pdf}`
  else {
    const w = Number(b.v.resolution) || 1920
    const h = Math.round(w * (b.ctx.deck.height / b.ctx.deck.width))
    const f: Frame = { width: w, height: h, dx: 0, dy: 0, dw: w, dh: h, renderW: w, renderH: h, dpi: 96 }
    const name = `slides/slide-${String(s.index + 1).padStart(3, '0')}.jpg`
    b.files.set(name, await encodeNative(await b.ctx.images.frame(s.index, f, '#ffffff'), 'image/jpeg', 0.9))
    pic = `[width=\\linewidth]{${name}}`
  }
  const out = [
    `\\subsection*{${s.index + 1}. ${tex(b, s.title || `Slide ${s.index + 1}`)}}`,
    `\\hypertarget{slide-${s.index + 1}}{}`,
    `\\begin{center}\\fbox{\\includegraphics${pic}}\\end{center}`,
  ]
  if (b.v.articleNotes && s.notesText.trim()) out.push(notesTex(b, s), '\\par')
  out.push('\\bigskip')
  return out.join('\n')
}

export const LATEX_FORMATS: FormatDef[] = [latex]
