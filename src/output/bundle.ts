// Collecting what an export produces into the one file the user saves.
//
// One output file is handed over as itself. More than one becomes a .zip,
// built as the files arrive (fflate's streaming Zip) so an export of a few
// hundred 4K slides isn't held in memory twice. Pictures and video are
// stored, not deflated — they're compressed already and deflating them only
// costs time.

import { Zip, ZipDeflate, ZipPassThrough } from 'fflate'

export interface OutputFile {
  name: string
  blob: Blob
}

const STORED = /\.(png|jpe?g|webp|avif|gif|mp4|webm|mov|m4v|mp3|m4a|zip|pptx|ppsx|potx|odp|docx|pdf|woff2?)$/i

export class Bundle {
  private first?: { path: string; data: Uint8Array }
  private zip?: Zip
  private chunks: Uint8Array[] = []
  private zipError?: Error
  private paths = new Set<string>()
  count = 0
  bytes = 0

  /** Add a file. Paths use '/', and a repeated path gets " (2)" added. */
  add(path: string, data: Uint8Array | string): string {
    const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data
    path = this.unique(path.replace(/^\/+/, ''))
    this.count++
    this.bytes += bytes.length
    if (!this.first && !this.zip) {
      this.first = { path, data: bytes }
      return path
    }
    if (!this.zip) {
      this.startZip()
      const f = this.first!
      this.first = undefined
      this.push(f.path, f.data)
    }
    this.push(path, bytes)
    return path
  }

  private unique(path: string): string {
    if (!this.paths.has(path)) {
      this.paths.add(path)
      return path
    }
    const dot = path.lastIndexOf('.')
    const stem = dot > path.lastIndexOf('/') ? path.slice(0, dot) : path
    const ext = dot > path.lastIndexOf('/') ? path.slice(dot) : ''
    for (let n = 2; ; n++) {
      const p = `${stem} (${n})${ext}`
      if (!this.paths.has(p)) {
        this.paths.add(p)
        return p
      }
    }
  }

  private startZip(): void {
    this.zip = new Zip((err, chunk) => {
      if (err) this.zipError = err
      else this.chunks.push(chunk)
    })
  }

  private push(path: string, data: Uint8Array): void {
    const entry = STORED.test(path) ? new ZipPassThrough(path) : new ZipDeflate(path, { level: 6 })
    const now = new Date()
    entry.mtime = now
    this.zip!.add(entry)
    entry.push(data, true)
  }

  /** Finish. `zipName` is used when there is more than one file. */
  finish(zipName: string): OutputFile | null {
    if (this.first) {
      const f = this.first
      return { name: f.path.split('/').pop()!, blob: new Blob([f.data as BlobPart], { type: mimeFor(f.path) }) }
    }
    if (!this.zip) return null
    this.zip.end()
    if (this.zipError) throw this.zipError
    return { name: zipName, blob: new Blob(this.chunks as BlobPart[], { type: 'application/zip' }) }
  }
}

export function mimeFor(path: string): string {
  const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
  return ({
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp', avif: 'image/avif',
    gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', tga: 'image/x-tga',
    svg: 'image/svg+xml', pdf: 'application/pdf', mp4: 'video/mp4', webm: 'video/webm',
    html: 'text/html', htm: 'text/html', md: 'text/markdown', txt: 'text/plain', csv: 'text/csv',
    json: 'application/json', tex: 'application/x-tex', rtf: 'application/rtf', opml: 'text/x-opml',
    srt: 'application/x-subrip', vtt: 'text/vtt', zip: 'application/zip',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
    potx: 'application/vnd.openxmlformats-officedocument.presentationml.template',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    odp: 'application/vnd.oasis.opendocument.presentation',
  } as Record<string, string>)[ext] ?? 'application/octet-stream'
}
