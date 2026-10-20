// Handing the finished file to the user from inside a task pane.
//
// VERIFIED on Mac by the font manager add-in (2026-09-12): a blob URL and an
// <a download> click raises a native Save As dialog in PowerPoint for Mac's
// WKWebView, and works in Chromium (Windows WebView2) too. WKWebView blocks
// window.open, so that fallback only matters for a host that blocks the
// anchor instead.
//
// In development (`npm run dev` with ?sink in the URL, or the pane's sink
// switch) the file is POSTed to the dev server, which writes it to .out/ —
// so an export from real PowerPoint can be checked with pdfinfo, ffprobe and
// friends without a Save dialog in the way.

import type { OutputFile } from './bundle'

export type DeliveryOutcome = 'downloaded' | 'opened-in-window' | 'blocked' | 'sunk'

export async function deliver(file: OutputFile, opts: { sink?: boolean } = {}): Promise<DeliveryOutcome> {
  if (import.meta.env.DEV && opts.sink) {
    const res = await fetch(`/__sink?name=${encodeURIComponent(file.name)}`, { method: 'POST', body: file.blob })
    if (!res.ok) throw new Error(`Dev sink refused the file: ${res.status}`)
    return 'sunk'
  }
  const url = URL.createObjectURL(file.blob)
  try {
    const a = document.createElement('a')
    if ('download' in a) {
      a.href = url
      a.download = file.name
      document.body.appendChild(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(url), 60_000)
      return 'downloaded'
    }
  } catch {
    /* fall through to the window path */
  }
  const win = window.open(url, '_blank')
  if (win) {
    setTimeout(() => URL.revokeObjectURL(url), 120_000)
    return 'opened-in-window'
  }
  URL.revokeObjectURL(url)
  return 'blocked'
}
