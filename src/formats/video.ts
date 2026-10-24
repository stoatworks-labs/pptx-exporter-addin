// Slides into time: MP4 and WebM through WebCodecs, and animated GIF.
//
// The web view's own VideoEncoder does the encoding (hardware where it has
// it); mp4-muxer and webm-muxer only wrap the chunks. Each slide is the
// picture PowerPoint draws of it at rest, so builds and animations don't play,
// and there's no sound — WebCodecs gives us no way to hear the deck.
//
// Timing is worked out in whole frames first (buildTimeline, framePlan), so
// every timestamp lands on a frame boundary and the maths is testable without
// an encoder. A slide's transition is the way it comes *in*, and like
// PowerPoint's "After" timer the hold starts once that transition has finished.

import { GIFEncoder } from 'gifenc'
import type { Slide } from '../deck/model'
import { paletteFor, indexFrame } from '../encode/gif'
import { ctx2d, makeCanvas, rgba, type AnyCanvas } from '../render/canvas'
import type { Frame } from '../render/sizing'
import { background, frameFromValues, sizeOptions, slideSeconds } from './common'
import { checkCancel, type ExportContext, type FormatDef, type OptionDef, type Values } from './types'

// ---------------------------------------------------------------- timing

export interface Shot {
  slide: Slide
  /** First frame of this slide: the first frame of its crossfade in, or its first frame after a cut. */
  start: number
  /** Frames of crossfade from the previous slide at the start; 0 is a cut. */
  fade: number
  /** First frame of the next shot. The slide is alone on screen from start + fade until here. */
  end: number
}

export interface Timeline {
  fps: number
  /** Length of the whole video, in frames. */
  frames: number
  shots: Shot[]
}

/** Seconds of crossfade into the slide at `position` in the export.
 *
 * deck: where the slide has a transition, a crossfade as long as that
 *       transition (whatever the effect — a push or a vortex becomes a fade);
 *       no transition, or a cut, is a cut. `transitionSeconds` is not used: the
 *       parser always gives a transition a length, and a non-zero default
 *       overriding it would leave the deck's own lengths out of reach.
 * fade: every change, `transitionSeconds` long.
 * cut:  never.
 * The first slide never fades in — there's nothing before it. */
export function crossfadeSeconds(slide: Slide, position: number, v: Values): number {
  if (position === 0) return 0
  switch (v.transition) {
    case 'cut':
      return 0
    case 'fade':
      return Math.max(0, Number(v.transitionSeconds) || 0)
    default: {
      const t = slide.transition
      if (!t || t.type === 'none' || t.type === 'cut') return 0
      return Math.max(0, t.durationMs / 1000)
    }
  }
}

export function buildTimeline(slides: Slide[], v: Values, fps: number): Timeline {
  const shots: Shot[] = []
  let at = 0
  slides.forEach((slide, i) => {
    const fade = Math.round(crossfadeSeconds(slide, i, v) * fps)
    const hold = Math.max(1, Math.round(slideSeconds(slide, v) * fps))
    shots.push({ slide, start: at, fade, end: at + fade + hold })
    at += fade + hold
  })
  return { fps, frames: at, shots }
}

/** A frame's timestamp in microseconds, as WebCodecs and the muxers want it. */
export function frameUs(n: number, fps: number): number {
  return Math.round((n * 1_000_000) / fps)
}

export interface PlannedFrame {
  /** Frame number on the timeline. */
  n: number
  /** Timeline frames this encoded frame covers: 1, or a whole hold with minimal framing. */
  length: number
  /** Index into timeline.shots. */
  shot: number
  /** 1 is the shot's slide alone; below that, the slide at this opacity over the previous one. */
  mix: number
  key: boolean
}

/** Every frame to encode, in order. Keyframes at each slide change, where a
 * crossfade settles, and at least every `keyEvery` frames.
 *
 * Minimal framing encodes a hold as one long frame. It still ends with a
 * one-frame repeat of the last picture: WebM times a frame by the start of the
 * next, so without it the last slide would last no time at all. */
export function* framePlan(t: Timeline, minimal: boolean, keyEvery: number): Generator<PlannedFrame> {
  let lastKey = -Infinity
  for (let k = 0; k < t.shots.length; k++) {
    const s = t.shots[k]
    let n = s.start
    while (n < s.end) {
      const fading = n < s.start + s.fade
      let length = 1
      if (minimal && !fading) {
        length = s.end - n
        if (k === t.shots.length - 1 && length > 1) length -= 1
      }
      const key = n === s.start || n === s.start + s.fade || n - lastKey >= keyEvery
      if (key) lastKey = n
      yield { n, length, shot: k, mix: fading ? (n - s.start + 1) / (s.fade + 1) : 1, key }
      n += length
    }
  }
}

// ---------------------------------------------------------------- captions

export interface Cue {
  /** Seconds. */
  start: number
  end: number
  text: string
}

/** One cue per slide with notes, from the moment it starts to appear until the next does. */
export function captionCues(t: Timeline): Cue[] {
  const cues: Cue[] = []
  for (const s of t.shots) {
    const text = cueText(s.slide.notesText)
    if (text) cues.push({ start: s.start / t.fps, end: s.end / t.fps, text })
  }
  return cues
}

/** A blank line ends a cue in both SRT and VTT, so blank lines go; so does an
 * arrow that would read as a timing line. */
function cueText(notes: string): string {
  return (notes ?? '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((l) => l.trim())
    .filter(Boolean)
    .join('\n')
    .replace(/-->/g, '→')
}

export function cueTime(seconds: number, sep: ',' | '.'): string {
  const ms = Math.max(0, Math.round(seconds * 1000))
  const p = (n: number, w = 2) => String(n).padStart(w, '0')
  return `${p(Math.floor(ms / 3_600_000))}:${p(Math.floor(ms / 60_000) % 60)}:${p(Math.floor(ms / 1000) % 60)}${sep}${p(ms % 1000, 3)}`
}

export function formatSrt(cues: Cue[]): string {
  return cues.map((c, i) => `${i + 1}\n${cueTime(c.start, ',')} --> ${cueTime(c.end, ',')}\n${c.text}\n`).join('\n')
}

export function formatVtt(cues: Cue[]): string {
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  return `WEBVTT\n\n${cues.map((c, i) => `${i + 1}\n${cueTime(c.start, '.')} --> ${cueTime(c.end, '.')}\n${esc(c.text)}\n`).join('\n')}`
}

// ---------------------------------------------------------------- size and bitrate

/** 4:2:0 video needs even sides. Round each down; where the slide filled that
 * side exactly it is drawn a pixel smaller rather than cropped, otherwise it is
 * recentred in the new box. */
export function evenFrame(f: Frame): Frame {
  const width = Math.max(2, f.width - (f.width % 2))
  const height = Math.max(2, f.height - (f.height % 2))
  if (width === f.width && height === f.height) return f
  const out: Frame = { ...f, width, height }
  if (width !== f.width) {
    if (f.dx === 0 && f.dw === f.width) {
      out.dw = width
      if (f.renderW === f.dw) out.renderW = width
    } else {
      out.dx = Math.round((width - f.dw) / 2)
    }
  }
  if (height !== f.height) {
    if (f.dy === 0 && f.dh === f.height) {
      out.dh = height
      if (f.renderH === f.dh) out.renderH = height
    } else {
      out.dy = Math.round((height - f.dh) / 2)
    }
  }
  return out
}

export type VideoCodec = 'avc' | 'hevc' | 'av1' | 'vp9' | 'vp8'

export const CODEC_LABELS: Record<VideoCodec, string> = {
  avc: 'H.264', hevc: 'HEVC', av1: 'AV1', vp9: 'VP9', vp8: 'VP8',
}

/** H.264 Mbps at 1080p30 for each quality. */
const QUALITY_MBPS: Record<string, number> = { low: 4, medium: 8, high: 14, veryhigh: 24 }

/** Bits needed for the same quality, relative to H.264. */
const CODEC_EFFICIENCY: Record<VideoCodec, number> = { avc: 1, vp8: 1.1, hevc: 0.7, vp9: 0.7, av1: 0.6 }

/** Bits per second. Scales with pixel count, and with the square root of the
 * frame rate — slides sit still, so doubling the frames costs far less than
 * double the bits. A custom rate is used exactly as given. */
export function bitrateFor(quality: string, customMbps: number, codec: VideoCodec, width: number, height: number, fps: number): number {
  if (quality === 'custom') return Math.max(100_000, Math.round(customMbps * 1_000_000))
  const mbps = (QUALITY_MBPS[quality] ?? QUALITY_MBPS.medium)
    * ((width * height) / (1920 * 1080))
    * Math.sqrt(fps / 30)
    * CODEC_EFFICIENCY[codec]
  return Math.max(250_000, Math.round(mbps * 1000) * 1000)
}

// ---------------------------------------------------------------- codec strings

// Level tables, lowest first: what each level allows of picture size, pixel
// rate and bitrate. The encoder is offered the lowest level that fits, then a
// couple above it — an encoder may refuse a level while accepting a higher one.

// H.264: level_idc, max macroblocks/s, max frame size (macroblocks), max bitrate (kbit/s, Main/Baseline; High is 1.25×).
const AVC_LEVELS: [number, number, number, number][] = [
  [10, 1485, 99, 64], [11, 3000, 396, 192], [12, 6000, 396, 384], [13, 11880, 396, 768],
  [20, 11880, 396, 2000], [21, 19800, 792, 4000], [22, 20250, 1620, 4000],
  [30, 40500, 1620, 10000], [31, 108000, 3600, 14000], [32, 216000, 5120, 20000],
  [40, 245760, 8192, 20000], [41, 245760, 8192, 50000], [42, 522240, 8704, 50000],
  [50, 589824, 22080, 135000], [51, 983040, 36864, 240000], [52, 2073600, 36864, 240000],
  [60, 4177920, 139264, 240000], [61, 8355840, 139264, 480000], [62, 16711680, 139264, 800000],
]

// HEVC Main tier: level × 30, max luma picture size, max luma samples/s, max bitrate (kbit/s).
const HEVC_LEVELS: [number, number, number, number][] = [
  [30, 36864, 552960, 128], [60, 122880, 3686400, 1500], [63, 245760, 7372800, 3000],
  [90, 552960, 16588800, 6000], [93, 983040, 33177600, 10000],
  [120, 2228224, 66846720, 12000], [123, 2228224, 133693440, 20000],
  [150, 8912896, 267386880, 25000], [153, 8912896, 534773760, 40000], [156, 8912896, 1069547520, 60000],
  [180, 35651584, 1069547520, 60000], [183, 35651584, 2139095040, 120000], [186, 35651584, 4278190080, 240000],
]

// AV1 Main tier: seq_level_idx, max picture size, max width, max height, max display rate (samples/s), max Mbit/s.
const AV1_LEVELS: [number, number, number, number, number, number][] = [
  [0, 147456, 2048, 1152, 4423680, 1.5], [1, 278784, 2816, 1584, 8363520, 3],
  [4, 665856, 4352, 2448, 19975680, 6], [5, 1065024, 5504, 3096, 31950720, 10],
  [8, 2359296, 6144, 3456, 70778880, 12], [9, 2359296, 6144, 3456, 141557760, 20],
  [12, 8912896, 8192, 4352, 267386880, 30], [13, 8912896, 8192, 4352, 534773760, 40],
  [14, 8912896, 8192, 4352, 1069547520, 60], [15, 8912896, 8192, 4352, 1069547520, 60],
  [16, 35651584, 16384, 8704, 1069547520, 60], [17, 35651584, 16384, 8704, 2139095040, 100],
  [18, 35651584, 16384, 8704, 4278190080, 160], [19, 35651584, 16384, 8704, 4278190080, 160],
]

// VP9: level × 10, max luma samples/s, max luma picture size, max bitrate (kbit/s), max side.
const VP9_LEVELS: [number, number, number, number, number][] = [
  [10, 829440, 36864, 200, 512], [11, 2764800, 73728, 800, 768],
  [20, 4608000, 122880, 1800, 960], [21, 9216000, 245760, 3600, 1344],
  [30, 20736000, 552960, 7200, 2048], [31, 36864000, 983040, 12000, 2752],
  [40, 83558400, 2228224, 18000, 4160], [41, 160432128, 2228224, 30000, 4160],
  [50, 311951360, 8912896, 60000, 8384], [51, 588251136, 8912896, 120000, 8384], [52, 1176502272, 8912896, 180000, 8384],
  [60, 1176502272, 35651584, 180000, 16832], [61, 2353004544, 35651584, 240000, 16832], [62, 4706009088, 35651584, 480000, 16832],
]

/** Index of the first level that fits, or the last level when none does. */
function firstFit<T>(levels: T[], fits: (l: T) => boolean): number {
  const i = levels.findIndex(fits)
  return i < 0 ? levels.length - 1 : i
}

const hex2 = (n: number) => n.toString(16).padStart(2, '0')

/** Codec strings to offer the encoder, best first. */
export function codecCandidates(codec: VideoCodec, width: number, height: number, fps: number, bitrate: number): string[] {
  const out: string[] = []
  const upTo = <T>(levels: T[], i: number, make: (l: T) => string) => {
    for (let j = i; j < Math.min(levels.length, i + 3); j++) out.push(make(levels[j]))
  }
  switch (codec) {
    case 'avc': {
      const mbW = Math.ceil(width / 16), mbH = Math.ceil(height / 16), fs = mbW * mbH
      // High, then Main, then Constrained Baseline for encoders that only do that.
      for (const [profile, flags, brFactor] of [[0x64, 0x00, 1.25], [0x4d, 0x00, 1], [0x42, 0xe0, 1]]) {
        const i = firstFit(AVC_LEVELS, ([, mbps, maxFs, br]) => {
          const side = Math.sqrt(8 * maxFs)
          return fs <= maxFs && mbW <= side && mbH <= side && fs * fps <= mbps && bitrate <= br * 1000 * brFactor
        })
        upTo(AVC_LEVELS, i, ([idc]) => `avc1.${hex2(profile)}${hex2(flags)}${hex2(idc)}`)
      }
      break
    }
    case 'hevc': {
      const ps = width * height
      const i = firstFit(HEVC_LEVELS, ([, maxPs, maxSr, br]) => {
        const side = Math.sqrt(8 * maxPs)
        return ps <= maxPs && width <= side && height <= side && ps * fps <= maxSr && bitrate <= br * 1000
      })
      // Main profile, Main tier, progressive frame-only constraint flags.
      upTo(HEVC_LEVELS, i, ([l]) => `hvc1.1.6.L${l}.B0`)
      break
    }
    case 'av1': {
      const ps = width * height
      const i = firstFit(AV1_LEVELS, ([, maxPs, maxW, maxH, rate, mbps]) =>
        ps <= maxPs && width <= maxW && height <= maxH && ps * fps <= rate && bitrate <= mbps * 1_000_000)
      // Main profile, Main tier, 8-bit.
      upTo(AV1_LEVELS, i, ([idx]) => `av01.0.${String(idx).padStart(2, '0')}M.08`)
      break
    }
    case 'vp9': {
      const ps = width * height
      const i = firstFit(VP9_LEVELS, ([, rate, maxPs, br, side]) =>
        ps <= maxPs && Math.max(width, height) <= side && ps * fps <= rate && bitrate <= br * 1000)
      // Profile 0 (4:2:0), 8-bit.
      upTo(VP9_LEVELS, i, ([l]) => `vp09.00.${l}.08`)
      break
    }
    case 'vp8':
      out.push('vp8')
      break
  }
  return out
}

function encoderConfig(codec: VideoCodec, codecString: string, width: number, height: number, fps: number, bitrate: number): VideoEncoderConfig {
  const config: VideoEncoderConfig = { codec: codecString, width, height, bitrate, framerate: fps }
  // MP4 wants the parameter sets out of band (avcC), not repeated in the stream.
  if (codec === 'avc') config.avc = { format: 'avc' }
  return config
}

async function supported(config: VideoEncoderConfig): Promise<boolean> {
  try {
    return !!(await VideoEncoder.isConfigSupported(config)).supported
  } catch {
    // A codec string the web view can't parse throws rather than answering no.
    return false
  }
}

/** The first configuration of this codec the web view will take, or null. */
async function findConfig(codec: VideoCodec, width: number, height: number, fps: number, bitrate: number): Promise<VideoEncoderConfig | null> {
  for (const s of codecCandidates(codec, width, height, fps, bitrate)) {
    const config = encoderConfig(codec, s, width, height, fps, bitrate)
    if (await supported(config)) return config
  }
  return null
}

function listOf(items: string[]): string {
  return items.length < 2 ? items.join('') : `${items.slice(0, -1).join(', ')} or ${items[items.length - 1]}`
}

// ---------------------------------------------------------------- containers

type Container = 'mp4' | 'webm'

const CONTAINER_CODECS: Record<Container, VideoCodec[]> = {
  mp4: ['avc', 'hevc', 'av1'],
  webm: ['vp9', 'vp8', 'av1'],
}

interface Sink {
  add(chunk: EncodedVideoChunk, meta: EncodedVideoChunkMetadata | undefined, durationUs: number): void
  finish(): Uint8Array
}

async function openSink(container: Container, codec: VideoCodec, width: number, height: number, fps: number, minimal: boolean): Promise<Sink> {
  if (container === 'mp4') {
    const { Muxer, ArrayBufferTarget } = await import('mp4-muxer')
    const target = new ArrayBufferTarget()
    const track = codec === 'hevc' ? 'hevc' : codec === 'av1' ? 'av1' : 'avc'
    // frameRate makes the track's timescale the frame rate, so every sample
    // lands on a whole frame.
    const muxer = new Muxer({ target, video: { codec: track, width, height, frameRate: fps }, fastStart: 'in-memory' })
    return {
      add(chunk, meta, durationUs) {
        // The raw form, so the duration is ours: a long held frame's length
        // can't depend on whether the web view copied it into the chunk.
        const data = new Uint8Array(chunk.byteLength)
        chunk.copyTo(data)
        muxer.addVideoChunkRaw(data, chunk.type, chunk.timestamp, durationUs, meta)
      },
      finish() {
        muxer.finalize()
        return new Uint8Array(target.buffer)
      },
    }
  }
  const { Muxer, ArrayBufferTarget } = await import('webm-muxer')
  const target = new ArrayBufferTarget()
  const id = codec === 'vp8' ? 'V_VP8' : codec === 'av1' ? 'V_AV1' : 'V_VP9'
  // WebM's frame rate is only a default frame length — a lie for held frames.
  const muxer = new Muxer({ target, video: { codec: id, width, height, frameRate: minimal ? undefined : fps } })
  return {
    add(chunk, meta) {
      muxer.addVideoChunk(chunk, meta)
    },
    finish() {
      muxer.finalize()
      return new Uint8Array(target.buffer)
    },
  }
}

const NO_WEBCODECS = "This web view has no video encoder (WebCodecs), so it can't make video. Recent Edge, Chrome and Safari have one."

const MAX_QUEUE = 6

/** Wait until the encoder has room — each queued frame is a full picture held in memory. */
async function room(enc: VideoEncoder): Promise<void> {
  while (enc.encodeQueueSize > MAX_QUEUE && enc.state === 'configured') {
    await new Promise<void>((resolve) => {
      const done = () => {
        clearTimeout(timer)
        enc.removeEventListener('dequeue', done)
        resolve()
      }
      // Polled as well: older WebKit has no dequeue event.
      const timer = setTimeout(done, 25)
      enc.addEventListener('dequeue', done)
    })
  }
}

const tick = () => new Promise((r) => setTimeout(r, 0))

function messageOf(e: unknown): string {
  return e instanceof Error ? e.message || e.name : String(e)
}

/** Slides whose animations won't appear, said once rather than discovered. (Video
 * and sound get the runner's media note.) */
function warnStill(ctx: ExportContext): void {
  const animated = ctx.slides.filter((s) => s.hasAnimations).map((s) => s.index + 1)
  if (animated.length) ctx.warn(`Slide${animated.length > 1 ? 's' : ''} ${animated.join(', ')} ${animated.length > 1 ? 'have' : 'has'} animations. They don't play — each slide appears as PowerPoint draws it at rest.`)
}

async function encodeVideo(ctx: ExportContext, v: Values, container: Container): Promise<void> {
  if (!ctx.slides.length) throw new Error('There are no slides to put in the video.')
  if (typeof VideoEncoder === 'undefined' || typeof VideoFrame === 'undefined') throw new Error(NO_WEBCODECS)
  const fps = Number(v.fps) || 30
  const codec = String(v.codec) as VideoCodec
  const label = CODEC_LABELS[codec] ?? codec
  const f = evenFrame(frameFromValues(ctx, v))
  const { width, height } = f
  const bitrate = bitrateFor(String(v.quality), Number(v.bitrate), codec, width, height, fps)

  const config = await findConfig(codec, width, height, fps, bitrate)
  if (!config) {
    const others: string[] = []
    for (const c of CONTAINER_CODECS[container]) {
      if (c !== codec && (await findConfig(c, width, height, fps, bitrateFor(String(v.quality), Number(v.bitrate), c, width, height, fps)))) {
        others.push(CODEC_LABELS[c])
      }
    }
    const tried = codecCandidates(codec, width, height, fps, bitrate).join(', ')
    throw new Error(
      `This web view can't encode ${label} at ${width} × ${height}, ${fps} fps, ${(bitrate / 1e6).toFixed(1)} Mbps. ` +
      (others.length ? `It can do ${listOf(others)} at that size — pick that instead.` : `Nor any other ${container.toUpperCase()} codec at that size — try a smaller one.`) +
      ` (Tried ${tried}.)`,
    )
  }

  warnStill(ctx)
  const timeline = buildTimeline(ctx.slides, v, fps)
  const minimal = v.framing === 'minimal'
  const sink = await openSink(container, codec, width, height, fps, minimal)
  const what = `${label} (${config.codec}, ${width} × ${height}, ${fps} fps, ${(bitrate / 1e6).toFixed(1)} Mbps)`

  // Durations by timestamp, so the muxer gets each frame's real length.
  const durations = new Map<number, number>()
  let failure: Error | null = null
  const enc = new VideoEncoder({
    output(chunk, meta) {
      if (failure) return
      try {
        const d = durations.get(chunk.timestamp) ?? chunk.duration ?? frameUs(1, fps)
        durations.delete(chunk.timestamp)
        sink.add(chunk, meta, d)
      } catch (e) {
        // Thrown inside the callback this would only reach the console.
        failure = new Error(`Couldn't add ${what} output to the ${container.toUpperCase()} file: ${messageOf(e)}`)
      }
    },
    error(e) {
      failure ??= new Error(`The ${what} encoder stopped: ${messageOf(e)}`)
    },
  })

  let still: VideoFrame | null = null
  try {
    try {
      enc.configure(config)
    } catch (e) {
      throw new Error(`Couldn't set up the ${what} encoder: ${messageOf(e)}`)
    }
    const comp = makeCanvas(width, height)
    const g = ctx2d(comp)
    let prev: AnyCanvas | null = null
    let cur: AnyCanvas | null = null
    let shot = -1
    let count = 0
    for (const p of framePlan(timeline, minimal, Math.round(2 * fps))) {
      checkCancel(ctx.signal)
      if (failure) throw failure
      if (p.shot !== shot) {
        shot = p.shot
        const s = timeline.shots[shot].slide
        ctx.progress((0.95 * p.n) / timeline.frames, `Slide ${s.index + 1}`)
        still?.close()
        still = null
        prev = cur
        cur = await ctx.images.frame(s.index, f, background(v))
      } else if (count % fps === 0) {
        ctx.progress((0.95 * p.n) / timeline.frames)
      }
      const timestamp = frameUs(p.n, fps)
      const duration = frameUs(p.n + p.length, fps) - timestamp
      let frame: VideoFrame
      try {
        if (p.mix < 1 && prev) {
          g.globalAlpha = 1
          g.drawImage(prev as CanvasImageSource, 0, 0)
          g.globalAlpha = p.mix
          g.drawImage(cur as CanvasImageSource, 0, 0)
          g.globalAlpha = 1
          frame = new VideoFrame(comp as CanvasImageSource, { timestamp, duration })
        } else {
          // One copy of the slide per shot; each frame of a hold shares its pixels.
          still ??= new VideoFrame(cur as CanvasImageSource, { timestamp })
          frame = new VideoFrame(still, { timestamp, duration })
        }
      } catch (e) {
        throw new Error(`Couldn't make a video frame of slide ${timeline.shots[shot].slide.index + 1} (${width} × ${height}): ${messageOf(e)}`)
      }
      durations.set(timestamp, duration)
      try {
        enc.encode(frame, { keyFrame: p.key })
      } catch (e) {
        // A closed encoder throws InvalidStateError here; the reason it closed is in failure.
        throw failure ?? new Error(`The ${what} encoder refused frame ${p.n}: ${messageOf(e)}`)
      } finally {
        frame.close()
      }
      await room(enc)
      // Let the pane repaint and a Cancel click land.
      if (++count % 15 === 0) await tick()
    }
    still?.close()
    still = null

    ctx.progress(0.96, 'Finishing the video')
    try {
      await enc.flush()
    } catch (e) {
      throw failure ?? new Error(`The ${what} encoder failed at the end: ${messageOf(e)}`)
    }
    if (failure) throw failure
    enc.close()

    ctx.progress(0.98, `Writing the ${container.toUpperCase()}`)
    let bytes: Uint8Array
    try {
      bytes = sink.finish()
    } catch (e) {
      throw new Error(`Couldn't finish the ${container.toUpperCase()} file (${what}): ${messageOf(e)}`)
    }
    ctx.add(`${ctx.base}.${container}`, bytes)

    if (v.captions === 'srt' || v.captions === 'vtt') {
      const cues = captionCues(timeline)
      if (!cues.length) ctx.warn('None of these slides have speaker notes, so there are no captions.')
      // The BOM is for Windows players that otherwise read SRT as ANSI.
      else if (v.captions === 'srt') ctx.add(`${ctx.base}.srt`, `\uFEFF${formatSrt(cues)}`)
      else ctx.add(`${ctx.base}.vtt`, formatVtt(cues))
    }
    ctx.progress(1)
  } finally {
    still?.close()
    if (enc.state !== 'closed') enc.close()
  }
}

// ---------------------------------------------------------------- options

const secondsOption: OptionDef = {
  key: 'seconds', label: 'Seconds per slide', type: 'number', min: 0.5, max: 600, step: 0.5, unit: 's', default: 5,
}

const timingsOption: OptionDef = {
  key: 'timings', label: "Use the deck's slide timings where set", type: 'toggle', default: true,
  help: 'Slides with a rehearsed or "After" timing stay up that long; the rest get the seconds above.',
}

function videoOptions(container: Container): OptionDef[] {
  const codecs: OptionDef = container === 'mp4'
    ? {
        key: 'codec', label: 'Codec', type: 'select', default: 'avc',
        choices: [{ value: 'avc', label: 'H.264' }, { value: 'hevc', label: 'HEVC (H.265)' }, { value: 'av1', label: 'AV1' }],
        help: 'H.264 plays everywhere. HEVC and AV1 are smaller for the same quality, but older players and editors may not open them.',
      }
    : {
        key: 'codec', label: 'Codec', type: 'select', default: 'vp9',
        choices: [{ value: 'vp9', label: 'VP9' }, { value: 'vp8', label: 'VP8' }, { value: 'av1', label: 'AV1' }],
        help: 'VP9 plays in every current browser. VP8 for elderly ones; AV1 is smallest and slowest to make.',
      }
  return [
    ...sizeOptions({ size: '1080p', fit: 'pad' }),
    {
      key: 'fps', label: 'Frame rate', type: 'select', default: '30',
      choices: ['24', '25', '30', '50', '60'].map((f) => ({ value: f, label: `${f} fps` })),
      help: 'Only crossfades move, so more frames mostly buys smoother fades.',
    },
    codecs,
    {
      key: 'quality', label: 'Quality', type: 'select', default: 'medium',
      choices: [
        { value: 'low', label: 'Low' },
        { value: 'medium', label: 'Medium' },
        { value: 'high', label: 'High' },
        { value: 'veryhigh', label: 'Very high' },
        { value: 'custom', label: 'Bitrate…' },
      ],
      help: 'Medium is about 8 Mbps for H.264 at 1080p30, scaled for the size, frame rate and codec.',
    },
    {
      key: 'bitrate', label: 'Bitrate', type: 'number', min: 0.2, max: 200, step: 0.1, unit: 'Mbps', default: 8,
      help: 'Used as given, whatever the size or codec.',
      when: (v) => v.quality === 'custom',
    },
    secondsOption,
    timingsOption,
    {
      key: 'transition', label: 'Transitions', type: 'select', default: 'deck',
      choices: [
        { value: 'deck', label: 'Where the deck has them' },
        { value: 'fade', label: 'Crossfade every change' },
        { value: 'cut', label: 'Straight cuts' },
      ],
      help: 'Every transition is played as a crossfade — a push or a vortex included. "Where the deck has them" keeps each one\'s own length.',
    },
    {
      key: 'transitionSeconds', label: 'Crossfade', type: 'number', min: 0, max: 10, step: 0.1, unit: 's', default: 0.5,
      help: 'Added between slides, not taken from their time.',
      when: (v) => v.transition === 'fade',
    },
    {
      key: 'framing', label: 'Frames', type: 'select', default: 'constant',
      choices: [
        { value: 'constant', label: 'Every frame' },
        { value: 'minimal', label: 'One per slide' },
      ],
      help: 'Every frame keeps editors and fussy players happy. One per slide holds each slide as a single long frame — a far smaller file that players are fine with, though editors may complain.',
    },
    {
      key: 'captions', label: 'Captions', type: 'select', default: 'none',
      choices: [
        { value: 'none', label: 'None' },
        { value: 'srt', label: 'Speaker notes as .srt' },
        { value: 'vtt', label: 'Speaker notes as .vtt' },
      ],
      help: 'A subtitle file beside the video: each slide\'s notes for as long as it is up. Slides without notes get no caption.',
    },
  ]
}

function videoUnavailable(container: Container): () => Promise<string | null> {
  return async () => {
    if (typeof VideoEncoder === 'undefined' || typeof VideoFrame === 'undefined') return NO_WEBCODECS
    for (const c of CONTAINER_CODECS[container]) {
      if (await findConfig(c, 1920, 1080, 30, bitrateFor('medium', 8, c, 1920, 1080, 30))) return null
    }
    const names = listOf(CONTAINER_CODECS[container].map((c) => CODEC_LABELS[c]))
    return `This web view's video encoder can't do ${names} at 1080p, so it can't make ${container.toUpperCase()}.`
  }
}

export const mp4: FormatDef = {
  id: 'mp4',
  label: 'MP4',
  group: 'video',
  blurb: 'H.264, HEVC or AV1 video, timed by the deck or by you. Silent, and each slide as it rests — animations don’t play.',
  ext: '.mp4',
  needs: { images: true },
  media: 'still',
  options: videoOptions('mp4'),
  unavailable: videoUnavailable('mp4'),
  run: (ctx, v) => encodeVideo(ctx, v, 'mp4'),
}

export const webm: FormatDef = {
  id: 'webm',
  label: 'WebM',
  group: 'video',
  blurb: 'VP9, VP8 or AV1 for the web. Silent, and each slide as it rests — animations don’t play.',
  ext: '.webm',
  needs: { images: true },
  media: 'still',
  options: videoOptions('webm'),
  unavailable: videoUnavailable('webm'),
  run: (ctx, v) => encodeVideo(ctx, v, 'webm'),
}

// ---------------------------------------------------------------- animated GIF

/** Each crossfade frame's length. Browsers stretch anything under 20 ms to 100 ms. */
const GIF_FADE_MS = 50
/** GIF delays are 16-bit hundredths of a second. */
const GIF_MAX_MS = 655_350

/** 'forever' → 0, 'once' → -1 (no loop block), 'three' → 2: a GIF's loop count is repeats after the first play. */
export function gifRepeat(loop: string): number {
  return loop === 'once' ? -1 : loop === 'three' ? 2 : 0
}

export const animatedGif: FormatDef = {
  id: 'animated-gif',
  label: 'Animated GIF',
  group: 'video',
  blurb: 'The deck as a looping GIF for chat, email and READMEs. 256 colours at most, and no sound, naturally.',
  ext: '.gif',
  needs: { images: true },
  media: 'still',
  options: [
    { key: 'width', label: 'Width', type: 'number', min: 16, max: 1920, step: 1, unit: 'px', default: 960 },
    secondsOption,
    timingsOption,
    { key: 'colors', label: 'Colours', type: 'slider', min: 2, max: 256, step: 1, default: 128 },
    { key: 'dither', label: 'Dither', type: 'toggle', default: true, help: 'Smooths gradients into dots instead of bands.' },
    {
      key: 'loop', label: 'Loop', type: 'select', default: 'forever',
      choices: [{ value: 'forever', label: 'Forever' }, { value: 'once', label: 'Play once' }, { value: 'three', label: 'Three times' }],
    },
    {
      key: 'crossfade', label: 'Crossfade frames', type: 'number', min: 0, max: 12, step: 1, default: 0,
      help: `Frames blended between slides, ${GIF_FADE_MS} ms each. Every one is a whole picture, so they add up.`,
    },
  ],
  async run(ctx, v) {
    if (!ctx.slides.length) throw new Error('There are no slides to put in the GIF.')
    warnStill(ctx)
    const f = frameFromValues(ctx, { size: 'width', width: Number(v.width), fit: 'contain' })
    const bg = background(v)
    const n = ctx.slides.length
    const fades = Math.max(0, Math.round(Number(v.crossfade) || 0))

    // Two passes so only two slides are ever held as pixels: the first samples
    // every slide for the one palette, the second indexes and writes. The
    // second's pictures mostly come from the render cache.
    const pixels = f.width * f.height
    const step = Math.max(1, Math.ceil((n * pixels) / 900_000))
    const perSlide = Math.ceil(pixels / step)
    const sample = new Uint8Array(n * perSlide * 4)
    for (let i = 0; i < n; i++) {
      checkCancel(ctx.signal)
      const s = ctx.slides[i]
      ctx.progress((0.3 * i) / n, `Colours of slide ${s.index + 1}`)
      const px = rgba(await ctx.images.frame(s.index, f, bg)).data
      let o = i * perSlide * 4
      for (let p = 0; p < pixels; p += step, o += 4) {
        sample[o] = px[p * 4]
        sample[o + 1] = px[p * 4 + 1]
        sample[o + 2] = px[p * 4 + 2]
        sample[o + 3] = 255
      }
      await tick()
    }
    const palette = paletteFor([{ width: sample.length / 4, height: 1, rgba: sample }], Number(v.colors))
    const dither = !!v.dither

    const gif = GIFEncoder()
    let first = true
    const write = (rgbaPx: Uint8Array | Uint8ClampedArray, delayMs: number) => {
      const index = indexFrame({ width: f.width, height: f.height, rgba: rgbaPx }, palette, dither)
      gif.writeFrame(index, f.width, f.height, first
        ? { palette, delay: delayMs, repeat: gifRepeat(String(v.loop)) }
        : { delay: delayMs })
      first = false
    }
    let prev: Uint8ClampedArray | null = null
    for (let i = 0; i < n; i++) {
      checkCancel(ctx.signal)
      const s = ctx.slides[i]
      ctx.progress(0.3 + (0.7 * i) / n, `Slide ${s.index + 1}`)
      const cur = rgba(await ctx.images.frame(s.index, f, bg)).data
      if (prev && fades) {
        const mixed = new Uint8ClampedArray(cur.length)
        for (let j = 1; j <= fades; j++) {
          checkCancel(ctx.signal)
          const a = j / (fades + 1)
          for (let k = 0; k < cur.length; k++) mixed[k] = prev[k] + (cur[k] - prev[k]) * a
          write(mixed, GIF_FADE_MS)
          await tick()
        }
      }
      write(cur, Math.min(GIF_MAX_MS, Math.round(slideSeconds(s, v) * 1000)))
      prev = cur
      await tick()
    }
    gif.finish()
    ctx.add(`${ctx.base}.gif`, gif.bytes())
    ctx.progress(1)
  },
}

export const VIDEO_FORMATS: FormatDef[] = [mp4, webm, animatedGif]
