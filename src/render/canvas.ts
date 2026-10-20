// Canvas plumbing that works in a task pane (WKWebView on Mac, WebView2 on
// Windows, an iframe on the web) and in a plain browser tab.

export type AnyCanvas = HTMLCanvasElement | OffscreenCanvas
export type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D

export function makeCanvas(w: number, h: number): AnyCanvas {
  const W = Math.max(1, Math.round(w))
  const H = Math.max(1, Math.round(h))
  if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(W, H)
  const c = document.createElement('canvas')
  c.width = W
  c.height = H
  return c
}

export function ctx2d(c: AnyCanvas): Ctx2D {
  const ctx = c.getContext('2d', { alpha: true }) as Ctx2D | null
  if (!ctx) throw new Error(`This web view could not make a ${c.width}×${c.height} canvas — try a smaller size.`)
  ctx.imageSmoothingEnabled = true
  ctx.imageSmoothingQuality = 'high'
  return ctx
}

export async function canvasToBlob(c: AnyCanvas, type: string, quality?: number): Promise<Blob> {
  if ('convertToBlob' in c) return c.convertToBlob({ type, quality })
  return new Promise((fulfil, reject) => {
    ;(c as HTMLCanvasElement).toBlob((b) => (b ? fulfil(b) : reject(new Error(`Could not encode ${type}.`))), type, quality)
  })
}

export async function blobBytes(b: Blob): Promise<Uint8Array> {
  return new Uint8Array(await b.arrayBuffer())
}

/** Encode with the web view's own encoder. Throws if it quietly hands back a
 * different type (Safari answers an unsupported image/webp with a PNG). */
export async function encodeNative(c: AnyCanvas, type: string, quality?: number): Promise<Uint8Array> {
  const b = await canvasToBlob(c, type, quality)
  if (b.type && b.type !== type) throw new UnsupportedEncoding(type)
  return blobBytes(b)
}

export class UnsupportedEncoding extends Error {
  constructor(public type: string) {
    super(`This web view cannot encode ${type} itself.`)
  }
}

const nativeSupport = new Map<string, Promise<boolean>>()

/** Whether canvas.toBlob(type) really produces that type here. */
export function canEncodeNatively(type: string): Promise<boolean> {
  let p = nativeSupport.get(type)
  if (!p) {
    p = (async () => {
      try {
        const c = makeCanvas(2, 2)
        ctx2d(c).fillRect(0, 0, 1, 1)
        const b = await canvasToBlob(c, type, 0.8)
        return b.type === type
      } catch {
        return false
      }
    })()
    nativeSupport.set(type, p)
  }
  return p
}

export async function decodeImage(bytes: Uint8Array | Blob, type = 'image/png'): Promise<ImageBitmap | HTMLImageElement> {
  const blob = bytes instanceof Blob ? bytes : new Blob([bytes as BlobPart], { type })
  if (typeof createImageBitmap !== 'undefined') {
    try {
      return await createImageBitmap(blob)
    } catch {
      /* fall back to <img> — some web views refuse certain blobs here */
    }
  }
  const url = URL.createObjectURL(blob)
  try {
    const img = new Image()
    img.decoding = 'async'
    img.src = url
    await img.decode()
    return img
  } finally {
    URL.revokeObjectURL(url)
  }
}

export function imageSize(img: ImageBitmap | HTMLImageElement): { w: number; h: number } {
  return 'naturalWidth' in img ? { w: img.naturalWidth, h: img.naturalHeight } : { w: img.width, h: img.height }
}

export function rgba(c: AnyCanvas): ImageData {
  return ctx2d(c).getImageData(0, 0, c.width, c.height)
}

/** Copy a canvas, optionally scaled. */
export function scaled(src: AnyCanvas, w: number, h: number): AnyCanvas {
  const c = makeCanvas(w, h)
  const ctx = ctx2d(c)
  ctx.drawImage(src as CanvasImageSource, 0, 0, c.width, c.height)
  return c
}

/** A CSS colour as [r, g, b]. Accepts #rgb, #rrggbb. */
export function parseHex(color: string): [number, number, number] {
  let h = color.trim().replace(/^#/, '')
  if (h.length === 3) h = h.split('').map((x) => x + x).join('')
  const n = parseInt(h.slice(0, 6), 16)
  if (!Number.isFinite(n)) return [0, 0, 0]
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
}
