// A new deck in which every slide is one picture of the original.
//
// The point is a file that plays the same on a machine without the fonts,
// the plug-ins or the linked files the original leans on. So it keeps what a
// presenter notices — size, order, hidden slides, notes, transitions and
// timings, clickable links — and nothing else: one theme, one master, one
// blank layout. Element order follows the schema exactly; PowerPoint answers
// anything else with its repair dialog.

import type { Deck, Link, Paragraph, Run, Slide, Transition } from '../deck/model'
import { hotspots } from './overlay'
import { appPropsXml, corePropsXml, CT, esc, NSX, PackageWriter, REL, Rels, XML_DECL } from './ooxml-write'

export interface Picture {
  bytes: Uint8Array
  type: 'png' | 'jpeg'
}

export interface PictureSlide extends Picture {
  slide: Slide
}

export interface FlatOptions {
  notes: boolean
  transitions: boolean
  links: boolean
}

export interface FlatResult {
  bytes: Uint8Array
  /** Links whose target slide isn't in the new deck, and so were dropped. */
  droppedLinks: number
}

const NOTES_W = 6858000
const NOTES_H = 9144000
const NS_DECL = `xmlns:a="${NSX.a}" xmlns:r="${NSX.r}" xmlns:p="${NSX.p}"`
const GROUP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
  + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
const CLR_MAP = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
  + 'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"'

export function buildFlatPptx(deck: Deck, pics: PictureSlide[], o: FlatOptions): FlatResult {
  const w = new PackageWriter()
  w.defaultType('png', 'image/png')
  w.defaultType('jpeg', 'image/jpeg')
  const hasNotes = o.notes && pics.some((p) => p.slide.notes.length > 0)

  const root = new Rels()
  root.add(REL.officeDocument, 'ppt/presentation.xml')
  root.add(REL.coreProps, 'docProps/core.xml')
  root.add(REL.extendedProps, 'docProps/app.xml')
  w.rels('', root)

  const presRels = new Rels()
  const masterRid = presRels.add(REL.slideMaster, 'slideMasters/slideMaster1.xml')
  const notesMasterRid = hasNotes ? presRels.add(REL.notesMaster, 'notesMasters/notesMaster1.xml') : null
  const slideRids = pics.map((_, i) => presRels.add(REL.slide, `slides/slide${i + 1}.xml`))
  presRels.add(REL.presProps, 'presProps.xml')
  presRels.add(REL.viewProps, 'viewProps.xml')
  presRels.add(REL.theme, 'theme/theme1.xml')
  presRels.add(REL.tableStyles, 'tableStyles.xml')

  w.add('ppt/presentation.xml', presentationXml(deck, masterRid, notesMasterRid, slideRids), CT.pptx)
  w.rels('ppt/presentation.xml', presRels)
  w.add('ppt/presProps.xml', presPropsXml(deck.show), CT.presProps)
  w.add('ppt/viewProps.xml', VIEW_PROPS, CT.viewProps)
  w.add('ppt/tableStyles.xml', `${XML_DECL}<a:tblStyleLst xmlns:a="${NSX.a}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`, CT.tableStyles)
  w.add('ppt/theme/theme1.xml', themeXml('Flattened'), CT.theme)

  const masterRels = new Rels()
  const layoutRid = masterRels.add(REL.slideLayout, '../slideLayouts/slideLayout1.xml')
  masterRels.add(REL.theme, '../theme/theme1.xml')
  w.add('ppt/slideMasters/slideMaster1.xml', slideMasterXml(layoutRid), CT.slideMaster)
  w.rels('ppt/slideMasters/slideMaster1.xml', masterRels)
  const layoutRels = new Rels()
  layoutRels.add(REL.slideMaster, '../slideMasters/slideMaster1.xml')
  w.add('ppt/slideLayouts/slideLayout1.xml', BLANK_LAYOUT, CT.slideLayout)
  w.rels('ppt/slideLayouts/slideLayout1.xml', layoutRels)

  if (hasNotes) {
    const nmRels = new Rels()
    nmRels.add(REL.theme, '../theme/theme2.xml')
    w.add('ppt/notesMasters/notesMaster1.xml', notesMasterXml(deck), CT.notesMaster)
    w.rels('ppt/notesMasters/notesMaster1.xml', nmRels)
    w.add('ppt/theme/theme2.xml', themeXml('Notes'), CT.theme)
  }

  // Deck position → number in the new deck, for links between slides.
  const number = new Map(pics.map((p, i) => [p.slide.index, i + 1]))
  let droppedLinks = 0
  pics.forEach((pic, i) => {
    const n = i + 1
    const part = `ppt/slides/slide${n}.xml`
    const rels = new Rels()
    rels.add(REL.slideLayout, '../slideLayouts/slideLayout1.xml')
    const media = `slide${n}.${pic.type === 'png' ? 'png' : 'jpeg'}`
    w.add(`ppt/media/${media}`, pic.bytes)
    const imageRid = rels.add(REL.image, `../media/${media}`)
    const shapes = [titleShape(deck, pic.slide), pictureShape(deck, pic.slide, imageRid)]
    if (o.links) {
      for (const h of hotspots(deck, pic.slide)) {
        const click = hlinkClick(h.link, rels, number)
        if (!click) {
          droppedLinks++
          continue
        }
        const id = shapes.length + 2
        shapes.push(hotspotShape(id, h.label, click, {
          x: h.x * deck.width, y: h.y * deck.height, w: h.w * deck.width, h: h.h * deck.height,
        }))
      }
    }
    if (hasNotes && pic.slide.notes.length) {
      rels.add(REL.notesSlide, `../notesSlides/notesSlide${n}.xml`)
      const nRels = new Rels()
      nRels.add(REL.notesMaster, '../notesMasters/notesMaster1.xml')
      nRels.add(REL.slide, `../slides/slide${n}.xml`)
      w.add(`ppt/notesSlides/notesSlide${n}.xml`, notesSlideXml(pic.slide.notes), CT.notesSlide)
      w.rels(`ppt/notesSlides/notesSlide${n}.xml`, nRels)
    }
    const transition = o.transitions && pic.slide.transition ? transitionXml(pic.slide.transition) : ''
    w.add(part, slideXml(shapes, pic.slide.hidden, transition), CT.slide)
    w.rels(part, rels)
  })

  w.add('docProps/core.xml', corePropsXml({ title: deck.title, creator: deck.author }), CT.core)
  w.add('docProps/app.xml', appPropsXml({
    PresentationFormat: 'Custom',
    Slides: pics.length,
    Notes: hasNotes ? pics.filter((p) => p.slide.notes.length).length : 0,
    HiddenSlides: pics.filter((p) => p.slide.hidden).length,
  }), CT.app)
  return { bytes: w.zip(), droppedLinks }
}

// ---------------------------------------------------------------------------
// presentation-level parts

function presentationXml(deck: Deck, masterRid: string, notesMasterRid: string | null, slideRids: string[]): string {
  const slides = slideRids.map((rid, i) => `<p:sldId id="${256 + i}" r:id="${rid}"/>`).join('')
  return `${XML_DECL}<p:presentation ${NS_DECL} saveSubsetFonts="1">`
    + `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRid}"/></p:sldMasterIdLst>`
    + (notesMasterRid ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRid}"/></p:notesMasterIdLst>` : '')
    + (slides ? `<p:sldIdLst>${slides}</p:sldIdLst>` : '')
    + `<p:sldSz cx="${deck.width}" cy="${deck.height}"/>`
    + `<p:notesSz cx="${NOTES_W}" cy="${NOTES_H}"/>`
    + '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>'
    + '</p:presentation>'
}

function presPropsXml(show: Deck['show']): string {
  return `${XML_DECL}<p:presentationPr ${NS_DECL}>`
    + `<p:showPr loop="${show.loop ? 1 : 0}" useTimings="${show.useTimings ? 1 : 0}"><p:present/><p:sldAll/></p:showPr>`
    + '</p:presentationPr>'
}

const VIEW_PROPS = `${XML_DECL}<p:viewPr ${NS_DECL}>`
  + '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>'
  + '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'

function slideMasterXml(layoutRid: string): string {
  const lvl = (sz: number) => `<a:lvl1pPr><a:defRPr sz="${sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr>`
  return `${XML_DECL}<p:sldMaster ${NS_DECL}>`
    // Black behind the pictures, so a rounding pixel at the edge doesn't show white.
    + '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    + `<p:spTree>${GROUP}</p:spTree></p:cSld>`
    + `<p:clrMap ${CLR_MAP}/>`
    + `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="${layoutRid}"/></p:sldLayoutIdLst>`
    + `<p:txStyles><p:titleStyle>${lvl(4400)}</p:titleStyle><p:bodyStyle>${lvl(2800)}</p:bodyStyle><p:otherStyle>${lvl(1800)}</p:otherStyle></p:txStyles>`
    + '</p:sldMaster>'
}

const BLANK_LAYOUT = `${XML_DECL}<p:sldLayout ${NS_DECL} type="blank" preserve="1">`
  + `<p:cSld name="Blank"><p:spTree>${GROUP}</p:spTree></p:cSld>`
  + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'

/** A plain Office-coloured theme. Nothing on these slides uses it but the
 * hidden titles and the notes; it has to exist and be complete. */
function themeXml(name: string): string {
  const srgb = (v: string) => `<a:srgbClr val="${v}"/>`
  const colours: [string, string][] = [
    ['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'], ['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'],
    ['dk2', srgb('44546A')], ['lt2', srgb('E7E6E6')], ['accent1', srgb('4472C4')], ['accent2', srgb('ED7D31')],
    ['accent3', srgb('A5A5A5')], ['accent4', srgb('FFC000')], ['accent5', srgb('5B9BD5')], ['accent6', srgb('70AD47')],
    ['hlink', srgb('0563C1')], ['folHlink', srgb('954F72')],
  ]
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
  const font = (face: string) => `<a:latin typeface="${face}"/><a:ea typeface=""/><a:cs typeface=""/>`
  return `${XML_DECL}<a:theme xmlns:a="${NSX.a}" name="${name}"><a:themeElements>`
    + `<a:clrScheme name="Office">${colours.map(([k, v]) => `<a:${k}>${v}</a:${k}>`).join('')}</a:clrScheme>`
    + `<a:fontScheme name="Office"><a:majorFont>${font('Calibri Light')}</a:majorFont><a:minorFont>${font('Calibri')}</a:minorFont></a:fontScheme>`
    + '<a:fmtScheme name="Office">'
    + `<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>`
    + `<a:lnStyleLst>${[6350, 12700, 19050].map((wd) => `<a:ln w="${wd}">${fill}</a:ln>`).join('')}</a:lnStyleLst>`
    + `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst>`
    + `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>`
    + '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
}

// ---------------------------------------------------------------------------
// slides

interface Box {
  x: number
  y: number
  w: number
  h: number
}

const xfrm = (b: Box) =>
  `<a:xfrm><a:off x="${Math.round(b.x)}" y="${Math.round(b.y)}"/><a:ext cx="${Math.max(1, Math.round(b.w))}" cy="${Math.max(1, Math.round(b.h))}"/></a:xfrm>`
const RECT = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'

function slideXml(shapes: string[], hidden: boolean, transition: string): string {
  return `${XML_DECL}<p:sld ${NS_DECL}${hidden ? ' show="0"' : ''}>`
    + `<p:cSld><p:spTree>${GROUP}${shapes.join('')}</p:spTree></p:cSld>`
    + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
    + transition
    + '</p:sld>'
}

/** The slide's title as a hidden title placeholder: PowerPoint's outline, its
 * "go to slide" list and screen readers still get a title, and nothing in it
 * needs a font to be installed. */
function titleShape(deck: Deck, s: Slide): string {
  const title = s.title || `Slide ${s.index + 1}`
  return '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title" hidden="1"/>'
    + '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
    + `<p:spPr>${xfrm({ x: deck.width * 0.05, y: deck.height * 0.05, w: deck.width * 0.9, h: deck.height * 0.15 })}${RECT}</p:spPr>`
    + `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(title)}</a:t></a:r></a:p></p:txBody></p:sp>`
}

function pictureShape(deck: Deck, s: Slide, rid: string): string {
  const alt = s.title || `Slide ${s.index + 1}`
  return `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Slide ${s.index + 1}" descr="${esc(alt)}"/>`
    + '<p:cNvPicPr><a:picLocks noGrp="1" noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
    + `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
    + `<p:spPr>${xfrm({ x: 0, y: 0, w: deck.width, h: deck.height })}${RECT}</p:spPr></p:pic>`
}

/** A clear rectangle carrying the link. A 100%-transparent fill rather than
 * no fill: PowerPoint only takes clicks inside a shape that has a fill. */
function hotspotShape(id: number, label: string, click: string, b: Box): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Link ${id - 3}" descr="${esc(label)}">${click}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xfrm(b)}${RECT}<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
}

const JUMP: Record<NonNullable<Link['jump']>, string> = {
  first: 'firstslide', last: 'lastslide', next: 'nextslide', previous: 'previousslide', end: 'endshow',
}

function hlinkClick(link: Link, rels: Rels, number: Map<number, number>): string | null {
  const tip = link.tooltip ? ` tooltip="${esc(link.tooltip)}"` : ''
  if (link.url) return `<a:hlinkClick r:id="${rels.add(REL.hyperlink, link.url, true)}"${tip}/>`
  if (link.slide != null) {
    const n = number.get(link.slide)
    if (n == null) return null
    return `<a:hlinkClick r:id="${rels.add(REL.slide, `slide${n}.xml`)}" action="ppaction://hlinksldjump"${tip}/>`
  }
  if (link.jump) return `<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=${JUMP[link.jump]}"${tip}/>`
  return null
}

// ---------------------------------------------------------------------------
// transitions

const FOUR = ['l', 'u', 'r', 'd']
const EIGHT = [...FOUR, 'lu', 'ru', 'ld', 'rd']
const HV = ['horz', 'vert']
/** PresentationML's own effects and the directions each accepts. Everything
 * else (p14's vortex, prism, ripple…, p15's morph) becomes a fade. */
const BASE_EFFECTS: Record<string, string[]> = {
  blinds: HV, checker: HV, comb: HV, randomBar: HV,
  cover: EIGHT, pull: EIGHT, push: FOUR, wipe: FOUR,
  strips: ['lu', 'ru', 'ld', 'rd'], zoom: ['in', 'out'], split: ['in', 'out'],
  circle: [], diamond: [], dissolve: [], newsflash: [], plus: [], random: [], wedge: [], wheel: [], cut: [], fade: [],
}
const SPEEDS: [string, number][] = [['fast', 500], ['med', 750], ['slow', 1000]]

export function transitionXml(t: Transition): string {
  let effect = ''
  if (t.type !== 'none') {
    const dirs = BASE_EFFECTS[t.type]
    const name = dirs ? t.type : 'fade'
    const dir = dirs && t.dir && dirs.includes(t.dir) ? ` dir="${t.dir}"` : ''
    effect = `<p:${name}${dir}/>`
  }
  const spd = t.durationMs <= 500 ? 'fast' : t.durationMs <= 750 ? 'med' : 'slow'
  const timing = (t.advanceOnClick ? '' : ' advClick="0"')
    + (t.advanceAfterMs != null ? ` advTm="${Math.max(0, Math.round(t.advanceAfterMs))}"` : '')
  const plain = effect ? `<p:transition spd="${spd}"${timing}>${effect}</p:transition>` : `<p:transition${timing}/>`
  // A duration the three speeds can't say goes in p14:dur, in the
  // AlternateContent wrapper PowerPoint 2010+ writes; older readers take the
  // plain fallback.
  if (!effect || SPEEDS.some(([, ms]) => ms === t.durationMs)) return plain
  return `<mc:AlternateContent xmlns:mc="${NSX.mc}"><mc:Choice xmlns:p14="${NSX.p14}" Requires="p14">`
    + `<p:transition spd="${spd}" p14:dur="${Math.max(0, Math.round(t.durationMs))}"${timing}>${effect}</p:transition>`
    + `</mc:Choice><mc:Fallback>${plain}</mc:Fallback></mc:AlternateContent>`
}

// ---------------------------------------------------------------------------
// notes

/** Where the slide picture sits on a notes page: the top half, the slide's shape. */
function notesImageBox(deck: Deck): Box {
  const maxW = 5486400
  const maxH = 3429000
  const s = Math.min(maxW / deck.width, maxH / deck.height)
  const w = deck.width * s
  const h = deck.height * s
  return { x: (NOTES_W - w) / 2, y: 685800, w, h }
}

function notesMasterXml(deck: Deck): string {
  const img = notesImageBox(deck)
  const bodyY = img.y + img.h + 457200
  const body: Box = { x: 685800, y: bodyY, w: 5486400, h: NOTES_H - bodyY - 685800 }
  return `${XML_DECL}<p:notesMaster ${NS_DECL}>`
    + '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>'
    + GROUP
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>'
    + `<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm(img)}${RECT}<a:noFill/>`
    + '<a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
    + `<p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm(body)}${RECT}</p:spPr>`
    + '<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/>'
    + '<a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
    + '</p:spTree></p:cSld>'
    + `<p:clrMap ${CLR_MAP}/>`
    + '<p:notesStyle>'
    + [0, 1, 2, 3, 4].map((l) => `<a:lvl${l + 1}pPr marL="${l * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
      + '<a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
      + `<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${l + 1}pPr>`).join('')
    + '</p:notesStyle></p:notesMaster>'
}

function notesSlideXml(notes: Paragraph[]): string {
  return `${XML_DECL}<p:notes ${NS_DECL}><p:cSld><p:spTree>${GROUP}`
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>'
    + '<p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
    + '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
    + '<p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr/>'
    + `<p:txBody><a:bodyPr/><a:lstStyle/>${drawingParagraphs(notes)}</p:txBody></p:sp>`
    + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
}

export function drawingParagraphs(ps: Paragraph[]): string {
  if (!ps.length) return '<a:p><a:endParaRPr lang="en-US"/></a:p>'
  return ps.map((p) => {
    const pPr = p.level ? `<a:pPr lvl="${Math.min(8, p.level)}"/>` : ''
    const runs = p.runs.map((r) => (r.br ? '<a:br/>' : `<a:r>${runProps(r)}<a:t>${esc(r.text)}</a:t></a:r>`)).join('')
    return `<a:p>${pPr}${runs || '<a:endParaRPr lang="en-US"/>'}</a:p>`
  }).join('')
}

function runProps(r: Run): string {
  const a: string[] = []
  if (r.bold) a.push('b="1"')
  if (r.italic) a.push('i="1"')
  if (r.underline) a.push('u="sng"')
  if (r.strike) a.push('strike="sngStrike"')
  if (r.baseline) a.push(`baseline="${r.baseline}"`)
  return a.length ? `<a:rPr ${a.join(' ')}/>` : ''
}
