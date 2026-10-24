// PowerPoint, through Office.js.

import { base64ToBytes, type FileProgress, type Host, type HostCapabilities } from './types'

const API_VERSIONS = ['1.10', '1.9', '1.8', '1.7', '1.6', '1.5', '1.4', '1.3', '1.2', '1.1']

function supported(v: string): boolean {
  try {
    return Office.context.requirements.isSetSupported('PowerPointApi', v)
  } catch {
    return false
  }
}

export class OfficeHost implements Host {
  readonly kind = 'powerpoint' as const
  readonly label: string
  readonly caps: HostCapabilities

  constructor() {
    const top = API_VERSIONS.find(supported) ?? 'none'
    this.caps = {
      slideImages: supported('1.8'),
      exportSlide: supported('1.8'),
      selection: supported('1.5'),
      // Pdf is a Common API file type; PowerPoint on every platform offers it.
      pdf: true,
      apiVersion: top,
    }
    const p = Office.context.platform
    const where = p === Office.PlatformType.Mac ? 'Mac' : p === Office.PlatformType.PC ? 'Windows'
      : p === Office.PlatformType.OfficeOnline ? 'the web' : p === Office.PlatformType.iOS ? 'iPad' : String(p)
    this.label = `PowerPoint for ${where}`
  }

  documentName(): string {
    const url = Office.context.document.url ?? ''
    const last = decodeURIComponent(url.split(/[\\/]/).pop() ?? '')
    return last.replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm|ppt|pps|pot)$/i, '')
  }

  pptx(onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
    return getFile(Office.FileType.Compressed, onProgress)
  }

  pdf(onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
    return getFile(Office.FileType.Pdf, onProgress)
  }

  async slideCount(): Promise<number> {
    return PowerPoint.run(async (ctx) => {
      const n = ctx.presentation.slides.getCount()
      await ctx.sync()
      return n.value
    })
  }

  async selectedSlides(): Promise<number[]> {
    if (!this.caps.selection) return []
    return PowerPoint.run(async (ctx) => {
      const all = ctx.presentation.slides
      all.load('items/id')
      const sel = ctx.presentation.getSelectedSlides()
      sel.load('items/id')
      await ctx.sync()
      const order = new Map(all.items.map((s, i) => [s.id, i]))
      return sel.items.map((s) => order.get(s.id)).filter((i): i is number => i != null).sort((a, b) => a - b)
    })
  }

  async renderSlide(index: number, width: number, height: number): Promise<Uint8Array> {
    return PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItemAt(index)
      const img = slide.getImageAsBase64({ width: Math.round(width), height: Math.round(height) })
      await ctx.sync()
      return base64ToBytes(img.value)
    })
  }

  async exportSlide(index: number): Promise<Uint8Array> {
    return PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItemAt(index)
      const data = slide.exportAsBase64()
      await ctx.sync()
      return base64ToBytes(data.value)
    })
  }
}

// getFileAsync hands the document back in slices (≤ 4 MB each). Learned on
// real Mac PowerPoint while building the font manager add-in: it is slow
// (~17 s for 9 MB), and it re-serialises the document — a deck PowerPoint
// thinks is slightly corrupt triggers its repair dialog first.
function getFile(type: Office.FileType, onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
  return new Promise((fulfil, reject) => {
    Office.context.document.getFileAsync(type, { sliceSize: 4 * 1024 * 1024 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(res.error.message))
        return
      }
      readSlices(res.value, onProgress).then(fulfil, reject)
    })
  })
}

export function concatSlices(slices: Uint8Array[]): Uint8Array {
  let total = 0
  for (const s of slices) total += s.length
  const out = new Uint8Array(total)
  let off = 0
  for (const s of slices) {
    out.set(s, off)
    off += s.length
  }
  return out
}

function readSlices(file: Office.File, onProgress?: (p: FileProgress) => void): Promise<Uint8Array> {
  return new Promise((fulfil, reject) => {
    const { sliceCount } = file
    // Filled as the slices arrive, rather than keeping every slice and copying
    // them into one at the end: that doubled the peak, and a 460 MB deck of
    // video in a pane that already held one copy wedged WKWebView at 2.7 GB
    // (PowerPoint for Mac 16.112.3, 2026-10-24). The slices are only kept if
    // the file turns out bigger than Office said.
    let whole: Uint8Array | null = file.size > 0 ? new Uint8Array(file.size) : null
    let filled = 0
    const slices: Uint8Array[] = []
    let received = 0
    let settled = false
    const fail = (msg: string) => {
      if (settled) return
      settled = true
      file.closeAsync(() => reject(new Error(msg)))
    }
    // One slice at a time, as Microsoft's own sample reads them. (The font
    // manager asks for all of them at once and that works on Mac; nothing here
    // gains from it, and one at a time keeps memory flat on a big deck.)
    const next = (i: number) => {
      file.getSliceAsync(i, (res) => {
        if (settled) return
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          fail(`Could not read part ${i + 1} of ${sliceCount} of the file: ${res.error.message}`)
          return
        }
        const raw = res.value.data as unknown
        const chunk = raw instanceof Uint8Array ? raw : Uint8Array.from(raw as number[])
        if (whole && filled + chunk.length <= whole.length) whole.set(chunk, filled)
        else {
          if (whole) slices.push(whole.subarray(0, filled))
          whole = null
          slices.push(chunk)
        }
        filled += chunk.length
        received++
        onProgress?.({ done: received, total: sliceCount })
        if (received === sliceCount) {
          settled = true
          const bytes = whole ? (filled === whole.length ? whole : whole.subarray(0, filled)) : concatSlices(slices)
          file.closeAsync(() => fulfil(bytes))
        } else next(i + 1)
      })
    }
    if (sliceCount === 0) {
      settled = true
      file.closeAsync(() => fulfil(new Uint8Array()))
    } else next(0)
  })
}
