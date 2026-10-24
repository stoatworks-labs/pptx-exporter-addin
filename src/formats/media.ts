// The deck's own video and sound: the Video bundle format, which saves the
// files, and the note every other format gets saying what it did with them.
//
// Both read the font manager's media scan (src/deck/media.ts), so a video on a
// layout, a sound PowerPoint calls wavAudioFile, an online film and a clip
// linked from the author's disk are all seen, whatever parse.ts makes of the
// shape.

import { mediaBytes, mediaFilenames, type MediaScanResult } from '@fm/core/media'
import { formatBytes, mediaManifest, secs } from '@fm/core/mediamanifest'
import { mediaOf, mediaOn, slidesPhrase, type Linked, type Online } from '../deck/media'
import { Cancelled, type ExportContext, type FormatDef } from './types'

// fflate's Zip, which bundle.ts streams into, has no Zip64.
const ZIP_LIMIT = 4_000_000_000

const cap = (s: string) => s.charAt(0).toUpperCase() + s.slice(1)
const has = (nums: number[]) => (new Set(nums).size > 1 ? 'have' : 'has')
const listOf = (xs: string[]) => {
  const u = [...new Set(xs)]
  return u.length > 1 ? `${u.slice(0, -1).join(', ')} and ${u[u.length - 1]}` : u[0] ?? ''
}

function onlineAndLinked(scan: MediaScanResult) {
  // A plain hyperlink to a video site opens a browser, in any format; it's in
  // MEDIA.txt but isn't media on the slide.
  const online = scan.placements.filter((p): p is Online => p.source === 'online' && p.via !== 'hyperlink')
  const linked = scan.placements.filter((p): p is Linked => p.source === 'linked')
  return { online, linked }
}

/** "Slide 7's video is online (YouTube) and slide 8's is linked from the author's disk". */
function notInDeck(online: Online[], linked: Linked[]): string {
  const bits: string[] = []
  if (online.length) bits.push(`${slidesPhrase(online.map((p) => p.slide))} ${online.length > 1 ? 'play' : 'plays'} online video (${listOf(online.map((p) => p.online.label))})`)
  if (linked.length) bits.push(`${slidesPhrase(linked.map((p) => p.slide))} ${linked.length > 1 ? 'have' : 'has'} media linked from the author’s disk (${listOf(linked.map((p) => p.target.replace(/^file:\/+/, '/').split(/[\\/]/).pop()!))})`)
  return cap(bits.join(', and '))
}

export const mediaBundle: FormatDef = {
  id: 'media',
  label: 'Video bundle',
  group: 'video',
  blurb: 'The deck’s own video and sound files, exactly as embedded, in a .zip with MEDIA.txt: which slide each plays on, and how PowerPoint plays it.',
  ext: '.zip',
  needs: {},
  media: 'files',
  options: [
    {
      key: 'which', label: 'Files', type: 'select', default: 'slides',
      choices: [
        { value: 'slides', label: 'Those on the exported slides' },
        { value: 'deck', label: 'Everything in the deck' },
      ],
      help: 'Everything means hidden slides and slides you didn’t pick as well. Video on a layout or master comes along with any slide that shows it.',
    },
    { key: 'audio', label: 'Sound files too', type: 'toggle', default: true, help: 'Off saves the video and leaves out the sound.' },
    {
      key: 'names', label: 'File names', type: 'select', default: 'slides',
      choices: [
        { value: 'slides', label: 'After slide and shape (Slide 2 - Intro film.mp4)' },
        { value: 'deck', label: 'As stored in the deck (media1.mp4)' },
      ],
    },
  ],
  async run(ctx, v) {
    const all = mediaOf(ctx.pptx)
    let scan = v.which === 'deck' ? all : mediaOn(all, ctx.slides)
    if (!v.audio) scan = { ...scan, files: scan.files.filter((f) => f.kind !== 'audio'), placements: scan.placements.filter((p) => p.kind !== 'audio') }
    for (const w of scan.warnings) ctx.warn(w)
    if (!scan.placements.length) {
      ctx.warn(v.which === 'deck' ? 'There’s no video or sound in this deck.' : 'There’s no video or sound on the exported slides.')
      return
    }

    const total = scan.files.reduce((n, f) => n + f.size, 0)
    if (total > ZIP_LIMIT) {
      throw new Error(`That’s ${formatBytes(total)} of media, more than one .zip can hold here (4 GB). Export it a few slides at a time.`)
    }
    const friendly = mediaFilenames(scan)
    const names = new Map(scan.files.map((f) => [f.part, v.names === 'deck' ? f.part.replace(/^.*\//, '') : friendly.get(f.part)!]))
    for (let i = 0; i < scan.files.length; i++) {
      if (ctx.signal.aborted) throw new Cancelled()
      const f = scan.files[i]
      ctx.progress(i / scan.files.length, names.get(f.part))
      // A view into the deck when the entry is stored, so a big film isn't copied.
      ctx.add(names.get(f.part)!, mediaBytes(ctx.pptx, f.part))
    }
    ctx.add('MEDIA.txt', mediaManifest(scan, names, ctx.base))
    ctx.progress(1)

    const { online, linked } = onlineAndLinked(scan)
    if (online.length || linked.length) {
      ctx.warn(`${notInDeck(online, linked)}. ${online.length + linked.length > 1 ? 'Those aren’t' : 'That isn’t'} in the deck, so ${online.length + linked.length > 1 ? 'they aren’t' : 'it isn’t'} in the zip; MEDIA.txt lists ${online.length + linked.length > 1 ? 'them' : 'it'}.`)
    }
    const trimmed = scan.placements.filter((p) => p.source === 'embedded' && (p.playback.trimStartMs || p.playback.trimEndMs))
    if (trimmed.length) {
      ctx.warn(`${cap(slidesPhrase(trimmed.map((p) => p.slide)))} ${has(trimmed.map((p) => p.slide))} a clip trimmed in PowerPoint. The saved file is the whole clip; MEDIA.txt has the trim points.`)
    }
  },
}

/**
 * The note a format gets about the deck's media on the slides it exported,
 * written by the runner from the format's own `media` declaration so that no
 * format loses a video without saying so.
 */
export function noteMedia(ctx: ExportContext, def: FormatDef): void {
  if (def.media === 'none' || def.media === 'files') return
  const scan = mediaOn(mediaOf(ctx.pptx), ctx.slides)
  const { online, linked } = onlineAndLinked(scan)
  const embedded = scan.placements.filter((p) => p.source === 'embedded')
  const onSlides = [...embedded, ...online, ...linked]
  if (!onSlides.length) return
  const nums = onSlides.map((p) => p.slide)

  if (def.media === 'still') {
    const how = def.group === 'video'
      ? 'It doesn’t play, and there’s no sound: each shows as PowerPoint’s picture of it, its poster frame.'
      : 'Nothing plays here: where this has a picture of it, that’s PowerPoint’s poster frame.'
    ctx.warn(`${cap(slidesPhrase(nums))} ${has(nums)} video or sound. ${how}${embedded.length ? ' Video bundle saves the files.' : ''}`)
    return
  }

  if (def.media === 'kept') {
    // Embedded media travels with the copy, trims and all. Only what was
    // never in the deck needs saying.
    const dead = online.filter((p) => p.via === 'flash' || p.online.retired)
    const live = online.filter((p) => !dead.includes(p))
    if (live.length) ctx.warn(`${notInDeck(live, [])}, which needs an internet connection wherever the copy is shown.`)
    for (const p of dead) ctx.warn(`Slide ${p.slide}’s ${p.via === 'flash' ? 'Flash video' : `${p.online.label} video`} won’t play anywhere: ${p.via === 'flash' ? 'Flash is retired' : `${p.online.label} has been retired`}.`)
    if (linked.length) ctx.warn(`${notInDeck([], linked)}, not embedded, so ${linked.length > 1 ? 'they play' : 'it plays'} only where that file is too.`)
    return
  }

  // 'plays': the HTML5 presentation, which notes each clip it can't copy as
  // it goes. What it can't see from parse.ts, or doesn't honour, is said here.
  const inherited = embedded.filter((p) => p.inheritedFrom)
  if (inherited.length) ctx.warn(`${cap(slidesPhrase(inherited.map((p) => p.slide)))} ${has(inherited.map((p) => p.slide))} video or sound on ${inherited.some((p) => p.inheritedFrom === 'master') ? 'a master or layout' : 'the layout'}, which the page leaves out.`)
  const shown = embedded.filter((p) => !p.inheritedFrom)
  for (const p of shown.filter((p) => p.playback.trimStartMs || p.playback.trimEndMs)) {
    const from = p.playback.trimStartMs ? secs(p.playback.trimStartMs) : '0 s'
    const to = p.playback.trimEndMs ? `${secs(p.playback.trimEndMs)} before the end` : 'the end'
    ctx.warn(`Slide ${p.slide}’s “${p.shape}” is trimmed in PowerPoint (from ${from} to ${to}), but the page plays the whole clip.`)
  }
  // PowerPoint writes vol="80000" on every clip it inserts, so 80% is its
  // default rather than anyone's choice; the page plays at full volume, which
  // only matters for a clip someone turned down.
  const settings = shown.filter((p) => p.playback.start === 'automatically' || p.playback.loop || p.playback.fullScreen || p.playback.hideWhenStopped || p.playback.muted || (p.playback.volume ?? 100) < 80)
  if (settings.length) ctx.warn(`PowerPoint’s playback settings on ${slidesPhrase(settings.map((p) => p.slide))} (starting by itself, looping, volume and the like) aren’t carried over: on the page each clip waits for a click and plays once.`)
}
