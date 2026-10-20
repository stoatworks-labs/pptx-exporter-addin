import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDeck } from '../src/deck/parse'
import type { Slide, Transition } from '../src/deck/model'
import { frameFor } from '../src/render/sizing'
import { defaults, type ExportContext, type Values } from '../src/formats/types'
import {
  bitrateFor, buildTimeline, captionCues, codecCandidates, crossfadeSeconds, cueTime, evenFrame,
  formatSrt, formatVtt, framePlan, frameUs, gifRepeat, mp4, VIDEO_FORMATS, webm, animatedGif,
} from '../src/formats/video'

const deck = parseDeck(new Uint8Array(readFileSync(resolve(__dirname, 'fixtures', 'sample-16x9.pptx'))))
/** What the pane hands a format by default: the deck without its hidden slide. */
const shown = deck.slides.filter((s) => !s.hidden)

const V = (over: Values = {}): Values => ({ ...defaults(mp4), ...over })

function slide(index: number, extra: Partial<Slide> = {}): Slide {
  return {
    index, sldId: 256 + index, part: `ppt/slides/slide${index + 1}.xml`, hidden: false,
    title: `Slide ${index + 1}`, shapes: [], notes: [], notesText: '', hasAnimations: false, ...extra,
  }
}
const tr = (type: string, durationMs: number, advanceAfterMs?: number): Transition =>
  ({ type, durationMs, advanceOnClick: true, advanceAfterMs })

describe('buildTimeline — the fixture deck', () => {
  it('holds each slide for the default, its rehearsed timing where it has one, and crossfades where the deck has transitions', () => {
    const t = buildTimeline(shown, V(), 30)
    expect(t.shots.map((s) => s.slide.index)).toEqual([0, 1, 2, 3, 4, 6, 7, 8])
    expect(t.shots.map(({ start, fade, end }) => [start, fade, end])).toEqual([
      [0, 0, 150],
      [150, 0, 300],
      [300, 0, 450],
      [450, 0, 600],
      [600, 0, 750],
      // Slide 7: a 750 ms fade (22.5 frames, rounded up) then 3 s of rehearsed timing.
      [750, 23, 863],
      // Slide 8: a 1.6 s push, played as a 48-frame crossfade, then the default 5 s.
      [863, 48, 1061],
      [1061, 0, 1211],
    ])
    expect(t.frames).toBe(1211)
  })

  it('ignores rehearsed timings when asked to', () => {
    const t = buildTimeline(shown, V({ timings: false }), 30)
    const s7 = t.shots.find((s) => s.slide.index === 6)!
    expect(s7.end - s7.start - s7.fade).toBe(150)
    expect(t.frames).toBe(8 * 150 + 23 + 48)
  })

  it('crossfades every change, never into the first slide', () => {
    const t = buildTimeline(shown, V({ transition: 'fade', transitionSeconds: 0.5 }), 30)
    expect(t.shots.map((s) => s.fade)).toEqual([0, 15, 15, 15, 15, 15, 15, 15])
    expect(t.frames).toBe(7 * 150 + 90 + 7 * 15)
  })

  it('cuts throughout, and the total is just the holds', () => {
    const t = buildTimeline(shown, V({ transition: 'cut' }), 30)
    expect(t.shots.every((s) => s.fade === 0)).toBe(true)
    expect(t.frames).toBe(7 * 150 + 90)
    // Shots tile the timeline with no gaps.
    t.shots.forEach((s, i) => expect(s.start).toBe(i ? t.shots[i - 1].end : 0))
  })

  it('scales with the frame rate', () => {
    const t = buildTimeline(shown, V(), 25)
    expect(t.shots.find((s) => s.slide.index === 7)!.fade).toBe(40)
    expect(t.frames).toBe(7 * 125 + 75 + 19 + 40) // 0.75 × 25 = 18.75 → 19
  })
})

describe('crossfadeSeconds', () => {
  it('treats cut and timing-only transitions as cuts, and anything else as a crossfade of its length', () => {
    const v = V({ transition: 'deck' })
    expect(crossfadeSeconds(slide(1, { transition: tr('cut', 500) }), 1, v)).toBe(0)
    expect(crossfadeSeconds(slide(1, { transition: tr('none', 500, 2000) }), 1, v)).toBe(0)
    expect(crossfadeSeconds(slide(1, { transition: tr('vortex', 2000) }), 1, v)).toBe(2)
    expect(crossfadeSeconds(slide(1), 1, v)).toBe(0)
  })

  it("keeps the deck's own lengths in deck mode, whatever transitionSeconds says", () => {
    expect(crossfadeSeconds(slide(1, { transition: tr('fade', 750) }), 1, V({ transition: 'deck', transitionSeconds: 2 }))).toBe(0.75)
  })

  it('never fades into the first slide of the export', () => {
    expect(crossfadeSeconds(slide(0, { transition: tr('fade', 1000) }), 0, V())).toBe(0)
    expect(crossfadeSeconds(slide(4), 0, V({ transition: 'fade' }))).toBe(0)
  })

  it('keeps a held slide on screen for at least one frame', () => {
    const t = buildTimeline([slide(0), slide(1)], V({ seconds: 0.001 }), 30)
    expect(t.shots.map((s) => s.end - s.start)).toEqual([1, 1])
  })
})

describe('framePlan', () => {
  const t = buildTimeline(shown, V(), 30)

  it('constant framing: one frame per tick, keyframes at slide changes, settled crossfades and every 2 s', () => {
    const plan = [...framePlan(t, false, 60)]
    expect(plan.length).toBe(t.frames)
    expect(plan.every((p, i) => p.n === i && p.length === 1)).toBe(true)
    const keys = plan.filter((p) => p.key).map((p) => p.n)
    for (const s of t.shots) {
      expect(keys).toContain(s.start)
      if (s.fade) expect(keys).toContain(s.start + s.fade)
    }
    // Never more than 2 s between keyframes.
    for (let i = 1; i < keys.length; i++) expect(keys[i] - keys[i - 1]).toBeLessThanOrEqual(60)
    expect(keys.slice(0, 4)).toEqual([0, 60, 120, 150])
  })

  it('blends the crossfade frames from just-past the old slide to just-short of the new', () => {
    const plan = [...framePlan(t, false, 60)]
    const s = t.shots[5]
    const fade = plan.filter((p) => p.n >= s.start && p.n < s.start + s.fade)
    expect(fade.length).toBe(23)
    expect(fade[0].mix).toBeCloseTo(1 / 24)
    expect(fade[22].mix).toBeCloseTo(23 / 24)
    expect(plan[s.start + s.fade].mix).toBe(1)
  })

  it('minimal framing: each hold is one long frame, crossfades are frame by frame, and the end has a one-frame tail', () => {
    const plan = [...framePlan(t, true, 60)]
    expect(plan.length).toBe(5 + (23 + 1) + (48 + 1) + 2)
    // Contiguous, and covering the whole timeline.
    for (let i = 1; i < plan.length; i++) expect(plan[i].n).toBe(plan[i - 1].n + plan[i - 1].length)
    expect(plan.reduce((n, p) => n + p.length, 0)).toBe(t.frames)
    expect(plan[0]).toMatchObject({ n: 0, length: 150, key: true, mix: 1 })
    const last = plan[plan.length - 1]
    expect(last).toMatchObject({ n: t.frames - 1, length: 1, shot: t.shots.length - 1 })
    expect(plan[plan.length - 2]).toMatchObject({ n: 1061, length: 149, key: true })
  })

  it('turns frames into whole-microsecond timestamps that never drift', () => {
    expect(frameUs(0, 30)).toBe(0)
    expect(frameUs(1, 30)).toBe(33333)
    expect(frameUs(30, 30)).toBe(1_000_000)
    expect(frameUs(1, 24)).toBe(41667)
    expect(frameUs(24 * 3600, 24)).toBe(3600 * 1_000_000)
  })
})

describe('captions', () => {
  const cues = captionCues(buildTimeline(shown, V(), 30))

  it('has one cue per slide with notes, spanning its time on screen', () => {
    expect(cues.length).toBe(5)
    expect(cues[0]).toEqual({ start: 0, end: 5, text: 'Welcome notes with ünïcödé ✓\nSecond line of the opening notes.' })
    // A blank line would end the cue early, so it goes.
    expect(cues[1]).toMatchObject({ start: 5, end: 10, text: 'Agenda notes.\nA paragraph after a blank line.' })
    expect(cues[3].start).toBeCloseTo(863 / 30)
    expect(cues[3].end).toBeCloseTo(1061 / 30)
  })

  it('formats timestamps for SRT and VTT', () => {
    expect(cueTime(5, ',')).toBe('00:00:05,000')
    expect(cueTime(5, '.')).toBe('00:00:05.000')
    expect(cueTime(3661.5, ',')).toBe('01:01:01,500')
    expect(cueTime(59.9996, '.')).toBe('00:01:00.000')
    expect(cueTime(863 / 30, ',')).toBe('00:00:28,767')
  })

  it('writes SRT', () => {
    const srt = formatSrt(cues)
    expect(srt.startsWith(
      '1\n00:00:00,000 --> 00:00:05,000\nWelcome notes with ünïcödé ✓\nSecond line of the opening notes.\n\n' +
      '2\n00:00:05,000 --> 00:00:10,000\nAgenda notes.\nA paragraph after a blank line.\n\n' +
      '3\n00:00:15,000 --> 00:00:20,000\nNotes in Ελληνικά and Русский.\n',
    )).toBe(true)
    expect(srt).toContain('4\n00:00:28,767 --> 00:00:35,367\nNotes with specials: 100% {sure} & $cash$ #tag ~ ^ _ \\ done.\n')
    expect(srt.endsWith('5\n00:00:35,367 --> 00:00:40,367\n日本語のノート。\n')).toBe(true)
  })

  it('writes VTT, escaping what VTT would read as markup', () => {
    const vtt = formatVtt(cues)
    expect(vtt.startsWith('WEBVTT\n\n1\n00:00:00.000 --> 00:00:05.000\nWelcome notes')).toBe(true)
    expect(vtt).toContain('100% {sure} &amp; $cash$')
  })

  it('skips empty notes and defuses arrows', () => {
    const t = buildTimeline([
      slide(0, { notesText: '  \n ' }),
      slide(1, { notesText: 'a --> b\r\n\r\n<i>c</i>' }),
    ], V({ seconds: 2 }), 30)
    const c = captionCues(t)
    expect(c).toEqual([{ start: 2, end: 4, text: 'a → b\n<i>c</i>' }])
    expect(formatVtt(c)).toBe('WEBVTT\n\n1\n00:00:02.000 --> 00:00:04.000\na → b\n&lt;i&gt;c&lt;/i&gt;\n')
  })
})

describe('evenFrame', () => {
  it('leaves an even frame alone', () => {
    const f = frameFor(deck.width, deck.height, { size: '1080p', fit: 'pad' })
    expect(evenFrame(f)).toBe(f)
  })

  it("shrinks the slide by a pixel where it fills the side, so nothing's cropped", () => {
    const f = frameFor(deck.width, deck.height, { size: 'width', width: 1001 })
    expect([f.width, f.height]).toEqual([1001, 563])
    expect(evenFrame(f)).toMatchObject({ width: 1000, height: 562, dx: 0, dy: 0, dw: 1000, dh: 562, renderW: 1000, renderH: 562 })
  })

  it('recentres a letterboxed slide in the trimmed box', () => {
    // A 4:3 slide in an odd 16:9 box: pillarboxed across, filling top to bottom.
    const f = frameFor(9144000, 6858000, { size: 'box', width: 1921, height: 1081, fit: 'pad' })
    expect(f).toMatchObject({ width: 1921, height: 1081, dw: 1441, dh: 1081, dx: 240, dy: 0 })
    expect(evenFrame(f)).toMatchObject({ width: 1920, height: 1080, dw: 1441, dx: 240, dh: 1080, dy: 0, renderH: 1080 })
  })

  it('recentres a cropped (cover) slide on the side that overhangs', () => {
    // 4:3 into a wide box: fills across exactly, overhangs top and bottom.
    const f = frameFor(9144000, 6858000, { size: 'box', width: 1001, height: 501, fit: 'cover' })
    expect(f).toMatchObject({ dw: 1001, dx: 0, dh: 751 })
    const e = evenFrame(f)
    expect([e.width, e.height]).toEqual([1000, 500])
    expect(e).toMatchObject({ dw: 1000, dx: 0, dh: 751, dy: Math.round((500 - 751) / 2) })
  })
})

describe('bitrateFor', () => {
  it('is about 8 Mbps for medium H.264 at 1080p30', () => {
    expect(bitrateFor('medium', 0, 'avc', 1920, 1080, 30)).toBe(8_000_000)
  })

  it('scales with pixel count', () => {
    expect(bitrateFor('medium', 0, 'avc', 3840, 2160, 30)).toBe(32_000_000)
    expect(bitrateFor('medium', 0, 'avc', 1280, 720, 30)).toBe(3_556_000)
  })

  it('scales with the square root of the frame rate', () => {
    expect(bitrateFor('medium', 0, 'avc', 1920, 1080, 60)).toBe(11_314_000)
  })

  it('asks less of the more efficient codecs', () => {
    expect(bitrateFor('medium', 0, 'hevc', 1920, 1080, 30)).toBe(5_600_000)
    expect(bitrateFor('medium', 0, 'av1', 1920, 1080, 30)).toBe(4_800_000)
    expect(bitrateFor('medium', 0, 'vp8', 1920, 1080, 30)).toBe(8_800_000)
  })

  it('orders the qualities and keeps a floor', () => {
    const q = ['low', 'medium', 'high', 'veryhigh'].map((k) => bitrateFor(k, 0, 'avc', 1920, 1080, 30))
    expect(q).toEqual([...q].sort((a, b) => a - b))
    expect(bitrateFor('low', 0, 'avc', 16, 16, 24)).toBe(250_000)
  })

  it('uses a custom rate exactly as given', () => {
    expect(bitrateFor('custom', 12.5, 'av1', 3840, 2160, 60)).toBe(12_500_000)
  })
})

describe('codecCandidates', () => {
  it('picks H.264 High at the lowest level that fits, Main and Baseline after', () => {
    const c = codecCandidates('avc', 1920, 1080, 30, 8_000_000)
    expect(c[0]).toBe('avc1.640028')
    expect(c).toContain('avc1.4d0028')
    expect(c).toContain('avc1.42e028')
    expect(codecCandidates('avc', 3840, 2160, 30, 32_000_000)[0]).toBe('avc1.640033')
    expect(codecCandidates('avc', 3840, 2160, 60, 45_000_000)[0]).toBe('avc1.640034')
    expect(codecCandidates('avc', 1280, 720, 30, 3_556_000)[0]).toBe('avc1.64001f')
  })

  it('climbs a level for frame rate or bitrate', () => {
    expect(codecCandidates('avc', 1920, 1080, 60, 11_314_000)[0]).toBe('avc1.64002a')
    expect(codecCandidates('avc', 1920, 1080, 30, 30_000_000)[0]).toBe('avc1.640029')
  })

  it('names HEVC, AV1 and VP9 levels', () => {
    expect(codecCandidates('hevc', 1920, 1080, 30, 5_600_000)[0]).toBe('hvc1.1.6.L120.B0')
    expect(codecCandidates('hevc', 3840, 2160, 30, 22_400_000)[0]).toBe('hvc1.1.6.L150.B0')
    expect(codecCandidates('av1', 1920, 1080, 30, 4_800_000)[0]).toBe('av01.0.08M.08')
    expect(codecCandidates('av1', 3840, 2160, 30, 19_200_000)[0]).toBe('av01.0.12M.08')
    expect(codecCandidates('vp9', 1920, 1080, 30, 5_600_000)[0]).toBe('vp09.00.40.08')
    expect(codecCandidates('vp9', 3840, 2160, 30, 22_400_000)[0]).toBe('vp09.00.50.08')
    expect(codecCandidates('vp8', 1920, 1080, 30, 8_000_000)).toEqual(['vp8'])
  })

  it('offers a couple of levels above the minimum, and stops at the top', () => {
    expect(codecCandidates('vp9', 1920, 1080, 30, 5_600_000)).toEqual(['vp09.00.40.08', 'vp09.00.41.08', 'vp09.00.50.08'])
    expect(codecCandidates('av1', 16384, 8704, 60, 1e9)).toEqual(['av01.0.19M.08'])
  })
})

describe('the formats', () => {
  it('are video formats that need slide pictures', () => {
    expect(VIDEO_FORMATS.map((f) => f.id)).toEqual(['mp4', 'webm', 'animated-gif'])
    for (const f of VIDEO_FORMATS) {
      expect(f.group).toBe('video')
      expect(f.needs.images).toBe(true)
      expect(new Set(f.options.map((o) => o.key)).size).toBe(f.options.length)
    }
  })

  it('default to 1080p letterboxed, 30 fps, the usual codec, the deck’s transitions', () => {
    expect(defaults(mp4)).toMatchObject({ size: '1080p', fit: 'pad', fps: '30', codec: 'avc', quality: 'medium', seconds: 5, timings: true, transition: 'deck', transitionSeconds: 0.5, framing: 'constant', captions: 'none' })
    expect(defaults(webm)).toMatchObject({ codec: 'vp9' })
    expect(defaults(animatedGif)).toMatchObject({ width: 960, colors: 128, dither: true, loop: 'forever', crossfade: 0 })
  })

  it('show the crossfade length only for crossfade-everything, and the bitrate only for custom', () => {
    const opt = (k: string) => mp4.options.find((o) => o.key === k)!
    expect(opt('transitionSeconds').when!(V({ transition: 'fade' }))).toBe(true)
    expect(opt('transitionSeconds').when!(V({ transition: 'deck' }))).toBe(false)
    expect(opt('transitionSeconds').when!(V({ transition: 'cut' }))).toBe(false)
    expect(opt('bitrate').when!(V({ quality: 'custom' }))).toBe(true)
  })

  it('cap the GIF at 1920 wide', () => {
    const w = animatedGif.options.find((o) => o.key === 'width')!
    expect(w.type === 'number' && w.max).toBe(1920)
  })

  it('maps GIF loops to loop counts', () => {
    expect([gifRepeat('forever'), gifRepeat('once'), gifRepeat('three')]).toEqual([0, -1, 2])
  })

  it('say plainly when there is no WebCodecs', async () => {
    expect(await mp4.unavailable!()).toMatch(/no video encoder \(WebCodecs\)/)
    expect(await webm.unavailable!()).toMatch(/WebCodecs/)
    const ctx = { slides: shown, deck } as unknown as ExportContext
    await expect(mp4.run(ctx, defaults(mp4))).rejects.toThrow(/WebCodecs/)
  })
})
