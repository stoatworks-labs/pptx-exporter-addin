// Every format the pane offers, in the order it lists them.

import { DOCUMENT_FORMATS } from './documents'
import { IMAGE_FORMATS } from './images'
import { LATEX_FORMATS } from './latex'
import { mediaBundle } from './media'
import { OFFICE_FORMATS } from './office'
import { TEXT_FORMATS } from './text'
import type { FormatDef } from './types'
import { VIDEO_FORMATS } from './video'
import { WEB_FORMATS } from './web'

export const FORMATS: FormatDef[] = [
  ...IMAGE_FORMATS,
  ...DOCUMENT_FORMATS,
  ...WEB_FORMATS,
  ...VIDEO_FORMATS,
  mediaBundle,
  ...LATEX_FORMATS,
  ...TEXT_FORMATS,
  ...OFFICE_FORMATS,
]

const ids = new Set<string>()
for (const f of FORMATS) {
  if (ids.has(f.id)) throw new Error(`Two formats share the id "${f.id}".`)
  ids.add(f.id)
}

export function formatById(id: string): FormatDef | undefined {
  return FORMATS.find((f) => f.id === id)
}
