// Running an export: every chosen format against the chosen slides, sharing
// one set of slide pictures and one PowerPoint PDF between them, into one
// file (or one .zip) for the user to save.

import type { Deck, Slide } from './deck/model'
import type { Host } from './host/types'
import { Bundle, type OutputFile } from './output/bundle'
import { SlideImages, type Engine } from './render/images'
import { pageForSlide } from './render/pdfmap'
import { noteMedia } from './formats/media'
import { Cancelled, type ExportContext, type FormatDef, type Needs, type Values } from './formats/types'

export interface Job {
  formats: { def: FormatDef; values: Values }[]
  slides: Slide[]
  base: string
  engine: Engine
}

export interface JobReport {
  file: OutputFile | null
  warnings: { format: string; message: string }[]
  errors: { format: string; message: string }[]
  seconds: number
  fileCount: number
}

export interface JobEvents {
  progress(fraction: number, label: string): void
}

export function needsOf(def: FormatDef, v: Values): Needs {
  return def.needsFor ? def.needsFor(v) : def.needs
}

export async function runJob(host: Host, deck: Deck, pptx: Uint8Array, job: Job, ev: JobEvents, signal: AbortSignal): Promise<JobReport> {
  const started = performance.now()
  const bundle = new Bundle()
  const warnings: JobReport['warnings'] = []
  const errors: JobReport['errors'] = []

  let pdf: Promise<Uint8Array> | undefined
  const getPdf = () => {
    pdf ??= host.pdf((p) => ev.progress(0, `Getting PowerPoint's PDF… ${p.done}/${p.total}`))
    return pdf
  }
  const images = new SlideImages(host, deck, job.engine, getPdf, (message) => warnings.push({ format: 'Slide pictures', message }))
  const multi = job.formats.length > 1

  // PowerPoint's PDF has no hidden slides, so on the PDF engine they can't be
  // drawn: formats that need pictures go without them (and say so); the rest
  // keep them.
  let drawable = job.slides
  if (job.engine === 'pdf' && job.slides.some((s) => s.hidden) && job.formats.some(({ def, values }) => needsOf(def, values).images)) {
    const map = await images.pdfMapping()
    const hidden = deck.slides.map((s) => s.hidden)
    drawable = job.slides.filter((s) => pageForSlide(map, hidden, s.index) != null)
    const left = job.slides.filter((s) => !drawable.includes(s)).map((s) => s.index + 1)
    if (left.length) warnings.push({ format: 'Slide pictures', message: `Slide${left.length > 1 ? 's' : ''} ${left.join(', ')} ${left.length > 1 ? 'are' : 'is'} hidden and not in PowerPoint's PDF, so formats made of pictures leave ${left.length > 1 ? 'them' : 'it'} out.` })
  }
  const share = 1 / job.formats.length

  for (let i = 0; i < job.formats.length; i++) {
    const { def, values } = job.formats[i]
    if (signal.aborted) throw new Cancelled()
    const prefix = multi ? `${def.id}/` : ''
    const ctx: ExportContext = {
      deck,
      slides: needsOf(def, values).images ? drawable : job.slides,
      base: job.base,
      host,
      images,
      pptx,
      add: (p, d) => void bundle.add(prefix + p, d),
      pdf: getPdf,
      progress: (f, label) => ev.progress(i * share + Math.min(1, Math.max(0, f)) * share, `${def.label}${label ? ` — ${label}` : ''}`),
      warn: (message) => void warnings.push({ format: def.label, message }),
      signal,
    }
    ctx.progress(0)
    try {
      await def.run(ctx, values)
      noteMedia(ctx, def)
    } catch (e) {
      if (e instanceof Cancelled || signal.aborted) throw new Cancelled()
      errors.push({ format: def.label, message: (e as Error)?.message ?? String(e) })
      console.error(`[${def.id}]`, e)
    }
  }

  ev.progress(1, 'Packing up')
  if (bundle.count > 1 && (warnings.length || errors.length)) {
    bundle.add('export-notes.txt', report(job, warnings, errors))
  }
  const zipName = multi ? `${job.base}-export.zip` : `${job.base}-${job.formats[0].def.id}.zip`
  return {
    file: bundle.finish(zipName),
    warnings,
    errors,
    seconds: (performance.now() - started) / 1000,
    fileCount: bundle.count,
  }
}

function report(job: Job, warnings: JobReport['warnings'], errors: JobReport['errors']): string {
  const lines = [`Exported by Advanced Exporter ${__APP_VERSION__} on ${new Date().toISOString()}`, '']
  if (errors.length) {
    lines.push('Formats that failed:')
    for (const e of errors) lines.push(`  ${e.format}: ${e.message}`)
    lines.push('')
  }
  if (warnings.length) {
    lines.push('Warnings:')
    for (const w of warnings) lines.push(`  ${w.format}: ${w.message}`)
    lines.push('')
  }
  lines.push(`Slides: ${job.slides.map((s) => s.index + 1).join(', ')}`)
  return lines.join('\n') + '\n'
}
