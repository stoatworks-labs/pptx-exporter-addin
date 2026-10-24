// The deck's video and sound as the font manager's scanner sees them
// (vendor/pptx-font-manager/src/core/media.ts): embedded, linked from the
// author's disk, or online, on the slide or inherited from its layout or
// master, with the playback settings and trims PowerPoint keeps in the timing
// tree and a probe of each file. parse.ts knows where media shapes sit on a
// slide; this knows everything else about them.

import { scanMedia, type MediaPlacement, type MediaScanResult } from '@fm/core/media'
import type { Slide } from './model'

const scans = new WeakMap<Uint8Array, MediaScanResult>()

/** The scan of a deck file, made once however many formats ask for it. */
export function mediaOf(pptx: Uint8Array): MediaScanResult {
  let r = scans.get(pptx)
  if (!r) {
    try {
      r = scanMedia(pptx)
    } catch (e) {
      // A deck the scanner can't read still exports; it just can't be asked
      // about its media.
      r = { slides: [], placements: [], files: [], warnings: [`Couldn't look for video and sound: ${(e as Error)?.message ?? e}`] }
    }
    scans.set(pptx, r)
  }
  return r
}

/**
 * The part of a scan on these slides: their placements (layout and master
 * media included, since the scanner puts those on every slide that shows
 * them) and the files those placements use, each listing only these slides.
 */
export function mediaOn(scan: MediaScanResult, slides: Slide[]): MediaScanResult {
  const nums = new Set(slides.map((s) => s.index + 1))
  const placements = scan.placements.filter((p) => nums.has(p.slide))
  const parts = new Set(placements.flatMap((p) => (p.source === 'embedded' ? [p.part] : [])))
  const files = scan.files.filter((f) => parts.has(f.part)).map((f) => ({ ...f, slides: f.slides.filter((n) => nums.has(n)) }))
  return { ...scan, placements, files }
}

/** "slide 4", "slides 2, 3 and 6". */
export function slidesPhrase(nums: number[]): string {
  const ns = [...new Set(nums)].sort((a, b) => a - b)
  if (ns.length === 1) return `slide ${ns[0]}`
  return `slides ${ns.slice(0, -1).join(', ')} and ${ns[ns.length - 1]}`
}

export type Online = Extract<MediaPlacement, { source: 'online' }>
export type Linked = Extract<MediaPlacement, { source: 'linked' }>
