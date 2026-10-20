// What the exporter needs from whatever is holding the deck.
//
// OfficeHost is the real thing (PowerPoint, through Office.js). MockHost
// stands in when the pane is opened in an ordinary browser, so the whole
// pipeline can be driven and tested without PowerPoint — its slide pictures
// are a crude drawing of the deck's text, not PowerPoint's rendering.

export interface HostCapabilities {
  /** Slide.getImageAsBase64 — PowerPoint draws the slide (PowerPointApi 1.8). */
  slideImages: boolean
  /** getFileAsync(Pdf) — PowerPoint's own vector PDF. */
  pdf: boolean
  /** Slide.exportAsBase64 — one slide as its own .pptx (PowerPointApi 1.8). */
  exportSlide: boolean
  /** getSelectedSlides (PowerPointApi 1.5). */
  selection: boolean
  /** Highest PowerPointApi set this host reports, for the status line. */
  apiVersion: string
}

export interface FileProgress {
  done: number
  total: number
}

export interface Host {
  readonly kind: 'powerpoint' | 'mock'
  readonly label: string
  readonly caps: HostCapabilities
  /** The document's file name without its extension, or '' when unsaved. */
  documentName(): string
  /** The open deck as .pptx bytes. Slow on big decks (~17 s for 9 MB on Mac). */
  pptx(onProgress?: (p: FileProgress) => void): Promise<Uint8Array>
  /** PowerPoint's PDF of the deck. */
  pdf(onProgress?: (p: FileProgress) => void): Promise<Uint8Array>
  /** Number of slides, hidden ones included. */
  slideCount(): Promise<number>
  /** Zero-based indexes of the slides selected in the thumbnail pane. */
  selectedSlides(): Promise<number[]>
  /** PNG of one slide, scaled to fit width × height. */
  renderSlide(index: number, width: number, height: number): Promise<Uint8Array>
  /** One slide as a stand-alone .pptx. */
  exportSlide(index: number): Promise<Uint8Array>
}

export function base64ToBytes(b64: string): Uint8Array {
  const fromBase64 = (Uint8Array as unknown as { fromBase64?: (s: string) => Uint8Array }).fromBase64
  if (fromBase64) return fromBase64(b64)
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

export function bytesToBase64(bytes: Uint8Array): string {
  const toBase64 = (bytes as unknown as { toBase64?: () => string }).toBase64
  if (toBase64) return toBase64.call(bytes)
  let s = ''
  const CH = 0x8000
  for (let i = 0; i < bytes.length; i += CH) {
    s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH) as unknown as number[])
  }
  return btoa(s)
}
