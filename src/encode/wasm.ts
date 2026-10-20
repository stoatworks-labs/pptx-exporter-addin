// WebP and AVIF through jSquash (Squoosh's codecs compiled to WebAssembly),
// for web views that can't encode them natively — which, for AVIF, is all of
// them, and for WebP includes WKWebView. Loaded only when used.

export async function encodeWebp(img: ImageData, quality: number, lossless: boolean): Promise<Uint8Array> {
  const { default: encode } = await import('@jsquash/webp/encode')
  const buf = await encode(img, lossless ? { lossless: 1, quality: 100, exact: 0 } : { quality, method: 4 })
  return new Uint8Array(buf)
}

export async function encodeAvif(img: ImageData, quality: number, speed: number, lossless = false): Promise<Uint8Array> {
  const { default: encode } = await import('@jsquash/avif/encode')
  const buf = await encode(img, { quality, speed, lossless, bitDepth: 8 })
  return new Uint8Array(buf)
}
