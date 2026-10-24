// PDFs: PowerPoint's own, pictures, handouts and notes pages.
//
// PowerPoint's PDF (getFileAsync(Pdf)) is vector and as faithful as it gets,
// but Office.js offers no options for it: no handouts, no notes pages, no
// slide range, and — like every exporter — it leaves hidden slides out. So
// the vector pages are reused where they exist (subset, split, laid out
// several to a sheet), and a slide PowerPoint's PDF hasn't got is drawn from
// a picture instead.
//
// This file is the formats and their options; the work is in
// documents-run.ts, loaded on first use so pdf-lib stays out of the pane's
// first load.

import { nameOption } from './common'
import type { FormatDef, Needs, OptionDef, Values } from './types'

const impl = () => import('./documents-run')

// --- PowerPoint's PDF --------------------------------------------------------

export const pdfVector: FormatDef = {
  id: 'pdf',
  label: 'PDF — PowerPoint (vector)',
  group: 'documents',
  blurb: "PowerPoint's own PDF: sharp at any zoom, text selectable. Trimmed to your slide range, or one file per slide.",
  ext: '.pdf',
  needs: { pdf: true },
  media: 'still',
  options: [
    { key: 'split', label: 'One PDF per slide', type: 'toggle', default: false },
    { ...nameOption('{deck}-{nn}'), when: (v) => !!v.split },
    { key: 'bookmarks', label: 'Bookmarks from slide titles', type: 'toggle', default: true },
    { key: 'info', label: 'Title and author from the deck', type: 'toggle', default: true },
  ],
  async run(ctx, v) {
    return (await impl()).runVector(ctx, v)
  },
}

// --- Pictures -----------------------------------------------------------------

const pictureOptions = (dpi: number): OptionDef[] => [
  { key: 'dpi', label: 'Resolution', type: 'number', min: 36, max: 600, step: 1, unit: 'dpi', default: dpi },
  {
    key: 'image', label: 'Pictures as', type: 'select', default: 'jpeg',
    choices: [{ value: 'jpeg', label: 'JPEG (small)' }, { value: 'png', label: 'PNG (lossless, big)' }],
  },
  { key: 'quality', label: 'JPEG quality', type: 'slider', min: 10, max: 100, step: 1, default: 85, when: (v) => v.image === 'jpeg' },
]

export const pdfPictures: FormatDef = {
  id: 'pdf-pictures',
  label: 'PDF — pictures',
  group: 'documents',
  blurb: 'Every slide as a picture at the resolution you choose. Hidden slides included, size under your control, looks identical everywhere.',
  ext: '.pdf',
  needs: { images: true },
  media: 'still',
  options: [
    ...pictureOptions(150),
    { key: 'links', label: 'Keep links clickable', type: 'toggle', default: true },
    { key: 'bookmarks', label: 'Bookmarks from slide titles', type: 'toggle', default: true },
  ],
  async run(ctx, v) {
    return (await impl()).runPictures(ctx, v)
  },
}

// --- Handouts and notes pages ----------------------------------------------------

const pageOptions: OptionDef[] = [
  {
    key: 'page', label: 'Paper', type: 'select', default: 'a4',
    choices: [
      { value: 'a4', label: 'A4' }, { value: 'letter', label: 'US Letter' }, { value: 'legal', label: 'US Legal' },
      { value: 'a3', label: 'A3' }, { value: 'tabloid', label: 'Tabloid' },
    ],
  },
  {
    key: 'source', label: 'Slides drawn from', type: 'select', default: 'vector',
    choices: [{ value: 'vector', label: "PowerPoint's PDF (vector)" }, { value: 'pictures', label: 'Pictures' }],
    help: "Vector is sharper and smaller. Pictures don't need PowerPoint's PDF and include hidden slides as drawn.",
  },
  { ...pictureOptions(200)[0], when: (v) => v.source === 'pictures' },
  { ...pictureOptions(200)[1], when: (v) => v.source === 'pictures' },
  { ...pictureOptions(200)[2], when: (v) => v.source === 'pictures' && v.image === 'jpeg' },
  { key: 'header', label: 'Header', type: 'text', default: '{deck}', help: '{deck} and {date} are filled in.' },
  { key: 'footer', label: 'Footer', type: 'text', default: '{date}' },
  { key: 'pageNumbers', label: 'Page numbers', type: 'toggle', default: true },
]

function drawnFrom(v: Values): Needs {
  return v.source === 'pictures' ? { images: true } : { pdf: true }
}

export const MARGIN = 40
export const HEAD = 22

interface Box {
  x: number
  y: number
  w: number
  h: number
}

/** Where each slide (and its writing lines) goes on a handout page, in PDF
 * coordinates. Pure — unit tested. */
export function handoutLayout(perPage: number, pageW: number, pageH: number, aspect: number, opts: { header: boolean; footer: boolean; numbers: boolean }): { slides: Box[]; lines: Box[] } {
  const top = pageH - MARGIN - (opts.header ? HEAD : 0)
  const bottom = MARGIN + (opts.footer ? HEAD : 0)
  const left = MARGIN
  const width = pageW - 2 * MARGIN
  const height = top - bottom
  const portrait = pageH >= pageW
  const label = opts.numbers ? 12 : 0
  const gutter = 18
  const grids: Record<number, [number, number]> = portrait
    ? { 1: [1, 1], 2: [1, 2], 3: [1, 3], 4: [2, 2], 6: [2, 3], 9: [3, 3] }
    : { 1: [1, 1], 2: [2, 1], 3: [3, 1], 4: [2, 2], 6: [3, 2], 9: [3, 3] }
  const [cols, rows] = grids[perPage] ?? [2, 3]
  const cellW = (width - gutter * (cols - 1)) / cols
  const cellH = (height - gutter * (rows - 1)) / rows
  const slides: Box[] = []
  const lines: Box[] = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cx = left + c * (cellW + gutter)
      const cyTop = top - r * (cellH + gutter)
      // 3-up: the slide takes part of the cell and ruled lines take the rest.
      const slideAreaW = perPage === 3 && portrait ? cellW * 0.48 : cellW
      const slideAreaH = (perPage === 3 && !portrait ? cellH * 0.45 : cellH) - label
      let w = slideAreaW
      let h = w / aspect
      if (h > slideAreaH) {
        h = slideAreaH
        w = h * aspect
      }
      const x = perPage === 3 && portrait ? cx : cx + (cellW - w) / 2
      const y = perPage === 3 && !portrait ? cyTop - h : cyTop - (slideAreaH - h) / 2 - h
      slides.push({ x, y, w, h })
      if (perPage === 3) {
        lines.push(portrait
          ? { x: cx + slideAreaW + gutter, y: y, w: cellW - slideAreaW - gutter, h }
          : { x: cx, y: cyTop - cellH, w: cellW, h: cellH - h - label - gutter / 2 })
      }
    }
  }
  return { slides, lines }
}

export const handouts: FormatDef = {
  id: 'handouts',
  label: 'Handouts',
  group: 'documents',
  blurb: 'Several slides to a page — 1, 2, 3 with lines for notes, 4, 6 or 9 — as PowerPoint prints them, minus the print dialog.',
  ext: '.pdf',
  needs: {},
  media: 'still',
  needsFor: drawnFrom,
  options: [
    {
      key: 'perPage', label: 'Slides per page', type: 'select', default: '3',
      choices: ['1', '2', '3', '4', '6', '9'].map((n) => ({ value: n, label: n === '3' ? '3, with lines for notes' : n })),
    },
    {
      key: 'orientation', label: 'Orientation', type: 'select', default: 'portrait',
      choices: [{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }],
    },
    ...pageOptions,
    { key: 'frame', label: 'Frame round each slide', type: 'toggle', default: true },
    { key: 'numbers', label: 'Slide numbers', type: 'toggle', default: true },
  ],
  async run(ctx, v) {
    return (await impl()).runLayout(ctx, v, 'handouts')
  },
}

export const notesPages: FormatDef = {
  id: 'notes-pages',
  label: 'Notes pages',
  group: 'documents',
  blurb: 'One slide a page with its speaker notes underneath, running onto more pages when the notes do.',
  ext: '.pdf',
  needs: {},
  media: 'still',
  needsFor: drawnFrom,
  options: [
    ...pageOptions,
    { key: 'fontSize', label: 'Notes size', type: 'number', min: 7, max: 24, step: 1, unit: 'pt', default: 12 },
    { key: 'frame', label: 'Frame round the slide', type: 'toggle', default: true },
    { key: 'skipEmpty', label: 'Leave out slides with no notes', type: 'toggle', default: false },
  ],
  async run(ctx, v) {
    return (await impl()).runLayout(ctx, v, 'notes')
  },
}

export const DOCUMENT_FORMATS: FormatDef[] = [pdfVector, pdfPictures, handouts, notesPages]
