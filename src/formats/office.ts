// PowerPoint and Office documents: the deck's own file, trimmed or re-typed;
// one file per slide; a flattened copy that needs no fonts; an OpenDocument
// version; and a Word handout.
//
// The package builders live in office-*.ts and take ready-made picture bytes,
// so they're tested without a canvas. This file is the options and the
// plumbing between the export context and those builders.

import type { Slide } from '../deck/model'
import { encodeNative } from '../render/canvas'
import type { Frame } from '../render/sizing'
import { eachSlide, expandName, frameFromValues, nameOption, sizeOptions } from './common'
import { rewritePptx, type PptxKind } from './office-copy'
import { buildNotesDocx, type HandoutLayout } from './office-docx'
import { buildFlatPptx, type Picture, type PictureSlide } from './office-flat'
import { buildOdp } from './office-odp'
import { checkCancel, type ExportContext, type FormatDef, type OptionDef, type Values } from './types'

const pictureOptions = (): OptionDef[] => [
  {
    key: 'image', label: 'Pictures', type: 'select', default: 'png',
    choices: [{ value: 'png', label: 'PNG (sharp)' }, { value: 'jpeg', label: 'JPEG (smaller)' }],
  },
  { key: 'quality', label: 'Quality', type: 'slider', min: 1, max: 100, step: 1, default: 88, when: (v) => v.image === 'jpeg' },
]

/** One slide's picture, at the frame's size, as the options ask. */
async function picture(ctx: ExportContext, s: Slide, f: Frame, v: Values): Promise<Picture> {
  if (v.image === 'jpeg') {
    const c = await ctx.images.frame(s.index, f, '#ffffff')
    return { bytes: await encodeNative(c, 'image/jpeg', Number(v.quality) / 100), type: 'jpeg' }
  }
  return { bytes: await ctx.images.framePng(s.index, f, '#ffffff'), type: 'png' }
}

async function pictures(ctx: ExportContext, f: Frame, v: Values): Promise<PictureSlide[]> {
  const out: PictureSlide[] = []
  await eachSlide(ctx, async (s) => {
    out.push({ slide: s, ...(await picture(ctx, s, f, v)) })
  })
  return out
}

/** A picture is a picture: say once what doesn't survive being one. (Video and
 * sound get the runner's media note.) */
function warnFlattened(ctx: ExportContext): void {
  const n = (ss: Slide[]) => `slide${ss.length > 1 ? 's' : ''} ${ss.map((s) => s.index + 1).join(', ')}`
  const animated = ctx.slides.filter((s) => s.hasAnimations)
  if (animated.length) ctx.warn(`Animations don't survive as pictures (${n(animated)}): each slide shows as PowerPoint draws it at rest.`)
}

export const pptxCopy: FormatDef = {
  id: 'pptx',
  label: 'PowerPoint file',
  group: 'office',
  blurb: 'The deck itself, as a .pptx, a show that opens straight into presenting, or a template — trimmed to the slides you picked.',
  ext: '.pptx / .ppsx / .potx',
  needs: {},
  media: 'kept',
  options: [
    {
      key: 'type', label: 'Save as', type: 'select', default: 'pptx',
      choices: [
        { value: 'pptx', label: 'Presentation (.pptx)' },
        { value: 'ppsx', label: 'Slide show (.ppsx) — opens straight into the show' },
        { value: 'potx', label: 'Template (.potx)' },
      ],
    },
    { key: 'onlyExported', label: 'Only the slides being exported', type: 'toggle', default: true, help: 'Off keeps every slide in the deck.' },
    { key: 'removeHidden', label: 'Leave out hidden slides', type: 'toggle', default: false },
    { key: 'removeNotes', label: 'Remove speaker notes', type: 'toggle', default: false },
    { key: 'removeComments', label: 'Remove comments', type: 'toggle', default: false, help: 'Old-style and modern comments, and the list of who wrote them.' },
    {
      key: 'stripPersonal', label: 'Remove names from the file', type: 'toggle', default: false,
      help: 'Clears author, last modified by, manager and company. The title stays. Comments keep their authors unless they go too.',
    },
  ],
  async run(ctx, v) {
    const kind = v.type as PptxKind
    const exported = new Set(ctx.slides.map((s) => s.index))
    const remove = ctx.deck.slides
      .filter((s) => (v.onlyExported && !exported.has(s.index)) || (v.removeHidden && s.hidden))
      .map((s) => s.index)
    if (ctx.deck.slides.length && remove.length === ctx.deck.slides.length) ctx.warn('That leaves no slides: the file has the masters and layouts only.')
    checkCancel(ctx.signal)
    ctx.progress(0.1, 'Rewriting the package')
    const out = rewritePptx(ctx.pptx, {
      kind,
      removeSlides: remove,
      removeNotes: !!v.removeNotes,
      removeComments: !!v.removeComments,
      stripPersonal: !!v.stripPersonal,
    })
    for (const w of out.warnings) ctx.warn(w)
    ctx.add(`${ctx.base}.${kind}`, out.bytes)
    ctx.progress(1)
  },
}

export const splitSlides: FormatDef = {
  id: 'split',
  label: 'One file per slide',
  group: 'office',
  blurb: 'Each slide as a .pptx of its own, masters, notes and media included. PowerPoint does the cutting.',
  ext: '.pptx',
  needs: { exportSlide: true },
  media: 'kept',
  options: [nameOption('{deck}-{nn}')],
  async run(ctx, v) {
    await eachSlide(ctx, async (s, i) => {
      ctx.add(`${expandName(String(v.name), ctx, s, i)}.pptx`, await ctx.host.exportSlide(s.index))
    })
  },
}

export const flattenedPptx: FormatDef = {
  id: 'flat-pptx',
  label: 'Flattened deck (pictures)',
  group: 'office',
  blurb: 'Every slide as one picture in a fresh deck, notes, transitions and links kept — it plays the same on a machine without your fonts.',
  ext: '.pptx',
  needs: { images: true },
  media: 'still',
  options: [
    ...sizeOptions({ size: '1080p', noFit: true }),
    ...pictureOptions(),
    { key: 'notes', label: 'Speaker notes', type: 'toggle', default: true },
    { key: 'transitions', label: 'Transitions and timings', type: 'toggle', default: true, help: 'Effects PowerPoint 2007 never had become a fade.' },
    { key: 'links', label: 'Clickable links', type: 'toggle', default: true, help: 'A clear shape over each link, pointing where the original did.' },
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, { ...v, fit: 'contain' })
    warnFlattened(ctx)
    const pics = await pictures(ctx, f, v)
    checkCancel(ctx.signal)
    const out = buildFlatPptx(ctx.deck, pics, { notes: !!v.notes, transitions: !!v.transitions, links: !!v.links })
    if (out.droppedLinks) ctx.warn(`${out.droppedLinks} link${out.droppedLinks === 1 ? '' : 's'} to slides left out of the export had nowhere to go and were dropped.`)
    ctx.add(`${ctx.base}-flattened.pptx`, out.bytes)
  },
}

export const odp: FormatDef = {
  id: 'odp',
  label: 'OpenDocument (pictures)',
  group: 'documents',
  blurb: 'An .odp of slide pictures with the speaker notes, for LibreOffice and anywhere that asks for ODF.',
  ext: '.odp',
  needs: { images: true },
  media: 'still',
  options: [
    ...sizeOptions({ size: '1080p', noFit: true }),
    ...pictureOptions(),
    { key: 'notes', label: 'Speaker notes', type: 'toggle', default: true },
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, { ...v, fit: 'contain' })
    const pics = await pictures(ctx, f, v)
    checkCancel(ctx.signal)
    ctx.add(`${ctx.base}.odp`, buildOdp(ctx.deck, pics, { notes: !!v.notes }))
  },
}

export const docxNotes: FormatDef = {
  id: 'docx',
  label: 'Word document (notes)',
  group: 'documents',
  blurb: 'PowerPoint’s “Create Handouts in Word”: slide pictures with the speaker notes, or with lines to write on.',
  ext: '.docx',
  needs: { images: true },
  media: 'still',
  options: [
    {
      key: 'layout', label: 'Layout', type: 'select', default: 'below',
      choices: [
        { value: 'beside', label: 'Notes next to slides' },
        { value: 'below', label: 'Notes below slides' },
        { value: 'lines', label: 'Blank lines below slides' },
      ],
    },
    { key: 'imageWidth', label: 'Picture width', type: 'number', min: 3, max: 25, step: 0.5, unit: 'cm', default: 15, when: (v) => v.layout !== 'beside' },
    { key: 'thumbWidth', label: 'Picture width', type: 'number', min: 3, max: 12, step: 0.5, unit: 'cm', default: 8, when: (v) => v.layout === 'beside' },
    {
      key: 'pageSize', label: 'Paper', type: 'select', default: 'a4',
      choices: [{ value: 'a4', label: 'A4' }, { value: 'letter', label: 'US Letter' }],
    },
    { key: 'outline', label: 'Slide text too', type: 'toggle', default: false, help: 'The words on each slide under its heading, so the handout can be searched and edited.' },
    ...pictureOptions(),
    { key: 'renderWidth', label: 'Picture resolution', type: 'number', min: 480, max: 3840, step: 10, unit: 'px', default: 1280, help: '1280 pixels is plenty for print at these sizes.' },
  ],
  async run(ctx, v) {
    const f = frameFromValues(ctx, { size: 'width', width: Number(v.renderWidth) || 1280, fit: 'contain' })
    const pics = await pictures(ctx, f, v)
    checkCancel(ctx.signal)
    const layout = v.layout as HandoutLayout
    ctx.add(`${ctx.base}-notes.docx`, buildNotesDocx(ctx.deck, pics, {
      layout,
      imageWidthCm: Number(layout === 'beside' ? v.thumbWidth : v.imageWidth),
      page: v.pageSize === 'letter' ? 'letter' : 'a4',
      outline: !!v.outline,
    }))
  },
}

export const OFFICE_FORMATS: FormatDef[] = [pptxCopy, splitSlides, flattenedPptx, odp, docxNotes]
