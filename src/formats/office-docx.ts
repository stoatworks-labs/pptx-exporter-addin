// A Word document of slide pictures and speaker notes — PowerPoint's own
// "Create Handouts in Microsoft Word", without needing Word on the machine
// doing the exporting, and with the notes' bold and italics kept.
//
// The smallest package Word opens without complaint: the document, its
// styles and settings, the pictures, and the two property parts. Element
// order inside w:pPr, w:rPr and w:tblPr is the schema's; Word is strict.

import type { Deck, Paragraph, Run, Shape, Slide, TableCell } from '../deck/model'
import type { PictureSlide } from './office-flat'
import { appPropsXml, corePropsXml, CT, esc, NSX, PackageWriter, REL, Rels, XML_DECL } from './ooxml-write'

export type HandoutLayout = 'beside' | 'below' | 'lines'

export interface DocxOptions {
  layout: HandoutLayout
  /** Picture width, in centimetres (clamped to what fits). */
  imageWidthCm: number
  page: 'a4' | 'letter'
  /** Put the slide's own text under its heading. */
  outline: boolean
}

const TWIPS_PER_CM = 1440 / 2.54
const EMU_PER_TWIP = 635
const CELL_MARGIN = 108
const MARGIN = 1134 // 2 cm
const PAGES = { a4: { w: 11906, h: 16838 }, letter: { w: 12240, h: 15840 } }
const LINE = 480 // ruled-line pitch for handwriting: about 8.5 mm

/** Things one document needs to number as it goes. */
class Doc {
  rels = new Rels()
  private drawingId = 0

  constructor(readonly w: PackageWriter) {
    this.rels.add(REL.styles, 'styles.xml')
    this.rels.add(REL.settings, 'settings.xml')
  }

  image(name: string, bytes: Uint8Array): string {
    this.w.add(`word/media/${name}`, bytes)
    return this.rels.add(REL.image, `media/${name}`)
  }

  link(url: string): string {
    return this.rels.add(REL.hyperlink, url, true)
  }

  nextDrawingId(): number {
    return ++this.drawingId
  }
}

export function buildNotesDocx(deck: Deck, pics: PictureSlide[], o: DocxOptions): Uint8Array {
  const w = new PackageWriter()
  w.defaultType('png', 'image/png')
  w.defaultType('jpeg', 'image/jpeg')
  const root = new Rels()
  root.add(REL.officeDocument, 'word/document.xml')
  root.add(REL.coreProps, 'docProps/core.xml')
  root.add(REL.extendedProps, 'docProps/app.xml')
  w.rels('', root)

  const doc = new Doc(w)
  const page = PAGES[o.page] ?? PAGES.a4
  const contentW = page.w - 2 * MARGIN
  const contentH = page.h - 2 * MARGIN

  const body: string[] = [para(run(deck.title), { style: 'Heading1' })]
  const meta = [deck.author, `${pics.length} slide${pics.length === 1 ? '' : 's'}`].filter(Boolean).join(' · ')
  body.push(para(run(meta, { color: '595959' })))

  for (const pic of pics) {
    const s = pic.slide
    const name = `slide-${String(s.index + 1).padStart(3, '0')}.${pic.type === 'png' ? 'png' : 'jpeg'}`
    const rid = doc.image(name, pic.bytes)
    const heading = `Slide ${s.index + 1}${s.title ? ` — ${s.title}` : ''}${s.hidden ? ' (hidden)' : ''}`
    // Beside: one slide after another in the flow, like PowerPoint's three a
    // page. Below and lines: a page each, so the title stands on its own.
    body.push(para(run(heading), { style: 'Heading2', pageBreakBefore: o.layout !== 'beside' }))
    const outline = o.outline ? outlineBlocks(doc, s, contentW) : []
    body.push(...outline)

    if (o.layout === 'beside') {
      const imgTw = Math.min(Math.round(o.imageWidthCm * TWIPS_PER_CM), contentW - Math.round(5 * TWIPS_PER_CM))
      const col1 = imgTw + 2 * CELL_MARGIN + 20
      const picture = drawing(doc, rid, name, s, imgTw * EMU_PER_TWIP, deck)
      const notes = notesParagraphs(doc, s.notes)
      body.push(table([col1, contentW - col1], [[[para(picture)], notes.length ? notes : [para('')]]], true), para(''))
      continue
    }

    const imgTw = Math.min(Math.round(o.imageWidthCm * TWIPS_PER_CM), contentW)
    const cx = imgTw * EMU_PER_TWIP
    body.push(para(drawing(doc, rid, name, s, cx, deck), { keepNext: true, spacing: [120, 240], center: true }))
    if (o.layout === 'below') {
      body.push(...notesParagraphs(doc, s.notes))
    } else {
      // Fill what's left of the page with lines, leaving room for the heading.
      const picH = Math.round((imgTw * deck.height) / deck.width)
      const outlineH = outline.length * 360
      const n = Math.max(3, Math.min(40, Math.floor((contentH - 900 - picH - 360 - outlineH) / LINE) - 1))
      for (let i = 0; i < n; i++) body.push(RULED_LINE)
    }
  }

  const sect = `<w:sectPr><w:pgSz w:w="${page.w}" w:h="${page.h}"/>`
    + `<w:pgMar w:top="${MARGIN}" w:right="${MARGIN}" w:bottom="${MARGIN}" w:left="${MARGIN}" w:header="567" w:footer="567" w:gutter="0"/>`
    + '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
  const document = `${XML_DECL}<w:document xmlns:w="${NSX.w}" xmlns:r="${NSX.r}" xmlns:wp="${NSX.wp}" xmlns:a="${NSX.a}" xmlns:pic="${NSX.pic}">`
    + `<w:body>${body.join('')}${sect}</w:body></w:document>`

  w.add('word/document.xml', document, CT.docx)
  w.rels('word/document.xml', doc.rels)
  w.add('word/styles.xml', STYLES, CT.styles)
  w.add('word/settings.xml', SETTINGS, CT.settings)
  w.add('docProps/core.xml', corePropsXml({ title: deck.title, creator: deck.author }), CT.core)
  w.add('docProps/app.xml', appPropsXml(), CT.app)
  return w.zip()
}

// ---------------------------------------------------------------------------
// paragraphs and runs

interface ParaOptions {
  style?: string
  keepNext?: boolean
  pageBreakBefore?: boolean
  /** Space before and after, in twips. */
  spacing?: [number, number]
  indent?: number
  center?: boolean
}

/** A w:p; pPr children in schema order. */
function para(content: string, p: ParaOptions = {}): string {
  const pPr = (p.style ? `<w:pStyle w:val="${p.style}"/>` : '')
    + (p.keepNext ? '<w:keepNext/>' : '')
    + (p.pageBreakBefore ? '<w:pageBreakBefore/>' : '')
    + (p.spacing ? `<w:spacing w:before="${p.spacing[0]}" w:after="${p.spacing[1]}"/>` : '')
    + (p.indent ? `<w:ind w:left="${p.indent}"/>` : '')
    + (p.center ? '<w:jc w:val="center"/>' : '')
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content}</w:p>`
}

interface RunStyle {
  bold?: boolean
  italic?: boolean
  strike?: boolean
  underline?: boolean
  color?: string
  vertAlign?: 'superscript' | 'subscript'
  link?: boolean
}

/** A w:r; tabs become w:tab. rPr children in schema order. */
function run(text: string, st: RunStyle = {}): string {
  const rPr = (st.link ? '<w:rStyle w:val="Hyperlink"/>' : '')
    + (st.bold ? '<w:b/>' : '')
    + (st.italic ? '<w:i/>' : '')
    + (st.strike ? '<w:strike/>' : '')
    + (st.color ? `<w:color w:val="${st.color}"/>` : '')
    + (st.underline ? '<w:u w:val="single"/>' : '')
    + (st.vertAlign ? `<w:vertAlign w:val="${st.vertAlign}"/>` : '')
  const body = text.split('\t').map((t) => (t ? `<w:t xml:space="preserve">${esc(t)}</w:t>` : '')).join('<w:tab/>')
  return `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${body}</w:r>`
}

function modelRun(doc: Doc, r: Run): string {
  if (r.br) return '<w:r><w:br/></w:r>'
  const st: RunStyle = {
    bold: r.bold, italic: r.italic, strike: r.strike, underline: r.underline,
    vertAlign: r.baseline ? (r.baseline > 0 ? 'superscript' : 'subscript') : undefined,
  }
  if (r.link?.url) return `<w:hyperlink r:id="${doc.link(r.link.url)}" w:history="1">${run(r.text, { ...st, link: true })}</w:hyperlink>`
  return run(r.text, st)
}

/** Paragraphs from the deck: bullets as a bullet character, numbers counted,
 * levels as indents. Blank paragraphs stay blank. */
function modelParagraphs(doc: Doc, ps: Paragraph[], baseIndent = 0): string[] {
  const counters: number[] = []
  return ps.map((p) => {
    let prefix = ''
    if (p.bullet === 'number') {
      // Deeper lists restart; anything else at this level or above breaks the run.
      counters.length = p.level + 1
      counters[p.level] = (counters[p.level] ?? 0) + 1
      prefix = `${counters[p.level]}.\t`
    } else {
      counters.length = Math.min(counters.length, p.level)
      if (p.bullet === 'char') prefix = '•\t'
    }
    const indent = baseIndent + p.level * 360
    const runs = (prefix ? run(prefix) : '') + p.runs.map((r) => modelRun(doc, r)).join('')
    return para(runs, { indent: indent || undefined })
  })
}

function notesParagraphs(doc: Doc, notes: Paragraph[]): string[] {
  return modelParagraphs(doc, notes)
}

/** The slide's text, in shape order, title left out (it's the heading). */
function outlineBlocks(doc: Doc, s: Slide, width: number): string[] {
  const out: string[] = []
  for (const sh of s.shapes) {
    const type = sh.placeholder?.type
    if (type === 'title' || type === 'ctrTitle') continue
    // The model gives every body-family placeholder bullets; a subtitle never shows them.
    const ps = sh.paragraphs && type === 'subTitle' ? sh.paragraphs.map((p) => ({ ...p, bullet: 'none' as const })) : sh.paragraphs
    if (ps) out.push(...modelParagraphs(doc, trimEmpty(ps)))
    else if (sh.table) out.push(outlineTable(doc, sh, width), para(''))
  }
  return out
}

function trimEmpty(ps: Paragraph[]): Paragraph[] {
  const blank = (p: Paragraph) => p.runs.every((r) => !r.text.trim())
  let a = 0
  let b = ps.length
  while (a < b && blank(ps[a])) a++
  while (b > a && blank(ps[b - 1])) b--
  return ps.slice(a, b)
}

// ---------------------------------------------------------------------------
// tables

type Cell = string[] | { paras: string[]; span: number; vMerge?: 'restart' | 'continue' }

function table(cols: number[], rows: Cell[][], cantSplit = false): string {
  const total = cols.reduce((a, b) => a + b, 0)
  const tblPr = `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="${total}" w:type="dxa"/><w:tblLayout w:type="fixed"/>`
    + '<w:tblLook w:val="04A0" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>'
  const grid = `<w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid>`
  const trs = rows.map((cells) => {
    let col = 0
    const tcs = cells.map((c) => {
      const cell = Array.isArray(c) ? { paras: c, span: 1 } : c
      const width = cols.slice(col, col + cell.span).reduce((a, b) => a + b, 0)
      col += cell.span
      const tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`
        + (cell.span > 1 ? `<w:gridSpan w:val="${cell.span}"/>` : '')
        + (cell.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : cell.vMerge === 'continue' ? '<w:vMerge/>' : '')
      // Every cell must end in a paragraph.
      return `<w:tc><w:tcPr>${tcPr}</w:tcPr>${cell.paras.length ? cell.paras.join('') : '<w:p/>'}</w:tc>`
    })
    return `<w:tr>${cantSplit ? '<w:trPr><w:cantSplit/></w:trPr>' : ''}${tcs.join('')}</w:tr>`
  })
  return `<w:tbl>${tblPr}${grid}${trs.join('')}</w:tbl>`
}

/** A slide's table as a Word table, merged cells kept merged. The model
 * flags covered cells without saying which way they're covered, so this
 * tracks the spans itself. */
function outlineTable(doc: Doc, sh: Shape, width: number): string {
  const t = sh.table!
  const n = Math.max(1, t.colWidths.length)
  const sum = t.colWidths.reduce((a, b) => a + b, 0) || n
  const cols = (t.colWidths.length ? t.colWidths : [1]).map((cw) => Math.max(200, Math.round((cw / sum) * width)))
  const below = new Array<number>(n).fill(0) // rows still covered from above, per column
  const spanAt = new Array<number>(n).fill(1) // width of the merge that starts in this column
  const rows: Cell[][] = t.rows.map((row) => {
    const out: Cell[] = []
    let col = 0
    let skip = 0 // cells covered by a colSpan to their left
    for (const c of row) {
      if (col >= n) break
      if (skip > 0) {
        skip--
      } else if (c.merged && below[col] > 0) {
        // A continuation is as wide as the cell that started the merge.
        const span = spanAt[col]
        for (let k = col; k < col + span; k++) below[k]--
        out.push({ paras: [], span, vMerge: 'continue' })
        skip = span - 1
      } else {
        const span = Math.max(1, Math.min(c.colSpan, n - col))
        skip = span - 1
        if (c.rowSpan > 1) {
          for (let k = col; k < col + span; k++) below[k] = c.rowSpan - 1
          spanAt[col] = span
        }
        out.push({ paras: cellParagraphs(doc, c), span, vMerge: c.rowSpan > 1 ? 'restart' : undefined })
      }
      col++
    }
    return out
  })
  return table(cols, rows)
}

function cellParagraphs(doc: Doc, c: TableCell): string[] {
  return c.merged ? [] : modelParagraphs(doc, trimEmpty(c.paragraphs))
}

// ---------------------------------------------------------------------------
// pictures

/** An inline picture: wp:inline children in schema order, extents in EMU. */
function drawing(doc: Doc, rid: string, name: string, s: Slide, cx: number, deck: Deck): string {
  const cy = Math.round((cx * deck.height) / deck.width)
  const id = doc.nextDrawingId()
  const alt = esc(s.title || `Slide ${s.index + 1}`)
  return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:docPr id="${id}" name="Slide ${s.index + 1}" descr="${alt}"/>`
    + `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic><a:graphicData uri="${NSX.pic}"><pic:pic>`
    + `<pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}" descr="${alt}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
    // A hairline, so a white slide doesn't vanish into white paper.
    + '<a:ln w="6350"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></pic:spPr>'
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
}

/** Word groups paragraphs with identical borders and draws only the group's
 * bottom edge, so each line is a "between" border plus the group's bottom. */
const RULED_LINE = '<w:p><w:pPr><w:pBdr>'
  + '<w:bottom w:val="single" w:sz="4" w:space="1" w:color="A6A6A6"/><w:between w:val="single" w:sz="4" w:space="1" w:color="A6A6A6"/>'
  + `</w:pBdr><w:spacing w:before="0" w:after="0" w:line="${LINE}" w:lineRule="exact"/></w:pPr></w:p>`

// ---------------------------------------------------------------------------
// styles and settings

const BORDER = 'w:val="single" w:sz="4" w:space="0" w:color="auto"'

const STYLES = `${XML_DECL}<w:styles xmlns:w="${NSX.w}">`
  + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
  + '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
  + '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + heading(1, 36, 0)
  + heading(2, 28, 360)
  + '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
  + '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'
  + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>'
  + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
  + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/>'
  + '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar>'
  + `<w:top w:w="0" w:type="dxa"/><w:left w:w="${CELL_MARGIN}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${CELL_MARGIN}" w:type="dxa"/>`
  + '</w:tblCellMar></w:tblPr></w:style>'
  + '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/>'
  + '<w:pPr><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>'
  + `<w:top ${BORDER}/><w:left ${BORDER}/><w:bottom ${BORDER}/><w:right ${BORDER}/><w:insideH ${BORDER}/><w:insideV ${BORDER}/>`
  + '</w:tblBorders></w:tblPr></w:style>'
  + '</w:styles>'

function heading(level: number, size: number, before: number): string {
  return `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>`
    + '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>'
    + `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>`
    + `<w:rPr><w:b/><w:bCs/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`
}

/** compatibilityMode 15, or Word opens it in Compatibility Mode. The view is
 * set because without one Word for Mac opened it in Outline view on Allan's
 * Mac (2026-10-20), where the slide pictures don't show. */
const SETTINGS = `${XML_DECL}<w:settings xmlns:w="${NSX.w}"><w:view w:val="print"/><w:zoom w:percent="100"/><w:defaultTabStop w:val="425"/>`
  + '<w:characterSpacingControl w:val="doNotCompress"/>'
  + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
  + '</w:settings>'
