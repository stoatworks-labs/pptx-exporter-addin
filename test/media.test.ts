import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'
import { parseDeck } from '../src/deck/parse'
import { FORMATS } from '../src/formats'
import { mediaBundle, noteMedia } from '../src/formats/media'
import type { ExportContext, FormatDef, MediaHandling } from '../src/formats/types'
import { fixture, runFormat, text } from './helpers'

// The font manager's media decks: media.pptx is built to copy what PowerPoint
// writes (embedded, linked, online, Flash, an add-in, a layout video, trims),
// online-video.pptx was made in PowerPoint for Mac with Insert ▸ Online Film.
const FM = '../../vendor/pptx-font-manager/test/fixtures/'
const MEDIA = FM + 'media.pptx'

describe('Video bundle', () => {
  it('saves every embedded file, byte for byte, under slide-and-shape names, with MEDIA.txt', async () => {
    const r = await runFormat(mediaBundle, { which: 'deck' }, { deck: MEDIA })
    const zip = unzipSync(fixture(MEDIA))
    const names = [...r.files.keys()]
    expect(names).toContain('Slide 1 - Intro film.mp4')
    expect(names).toContain('Slide 6 - ProRes master.mov')
    expect(names).toContain('MEDIA.txt')
    expect(names).toHaveLength(8) // seven files and MEDIA.txt
    expect(r.files.get('Slide 1 - Intro film.mp4')).toEqual(zip['ppt/media/media1.mp4'])
    const txt = text(r.files.get('MEDIA.txt'))
    expect(txt).toMatch(/^Media from sample\r\n/)
    expect(txt).toMatch(/NEEDS THE INTERNET/)
    expect(txt).toMatch(/LINKED FROM THE AUTHOR’S DISK/)
  })

  it('says what isn’t in the zip, and that trimmed clips are saved whole', async () => {
    const r = await runFormat(mediaBundle, { which: 'deck' }, { deck: MEDIA })
    expect(r.warnings.join('\n')).toMatch(/online video \(.*YouTube.*\).*linked from the author’s disk.*MEDIA\.txt lists them/s)
    expect(r.warnings.join('\n')).toMatch(/trimmed in PowerPoint\. The saved file is the whole clip/)
  })

  it('keeps to the exported slides, layout video included', async () => {
    const first = await runFormat(mediaBundle, {}, { deck: MEDIA, slides: [0] })
    expect([...first.files.keys()]).toEqual(['Slide 1 - Intro film.mp4', 'MEDIA.txt'])
    // Slide 7 has no video of its own; its layout has a loop.
    const seventh = await runFormat(mediaBundle, {}, { deck: MEDIA, slides: [6] })
    expect([...seventh.files.keys()]).toHaveLength(2)
    expect(text(seventh.files.get('MEDIA.txt'))).toMatch(/slide 7 \(layout\)/)
  })

  it('leaves sound out on request, and can keep the deck’s own file names', async () => {
    const noSound = await runFormat(mediaBundle, { which: 'deck', audio: false }, { deck: MEDIA })
    expect([...noSound.files.keys()].some((n) => n.endsWith('.mp3'))).toBe(false)
    const plain = await runFormat(mediaBundle, { which: 'deck', names: 'deck' }, { deck: MEDIA })
    expect([...plain.files.keys()].sort()).toEqual(['MEDIA.txt', 'media1.mp4', 'media2.mp4', 'media3.mp4', 'media4.mov', 'media5.wmv', 'media6.mp3', 'media7.mp4'])
  })

  it('says so when there’s nothing to save', async () => {
    const r = await runFormat(mediaBundle, {}, { deck: 'sample-4x3.pptx' })
    expect(r.files.size).toBe(0)
    expect(r.warnings).toEqual(['There’s no video or sound on the exported slides.'])
  })

  it('takes PowerPoint’s "NULL" link on a trimmed clip for no link', async () => {
    // PowerPoint for Mac 16.112.3 wrote the trimmed clip's videoFile link as
    // an External video relationship to "NULL", the file through p14:media.
    const zip = unzipSync(fixture('sample-16x9.pptx'))
    const rels = 'ppt/slides/_rels/slide9.xml.rels'
    zip[rels] = strToU8(strFromU8(zip[rels]).replace(
      'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/media1.mp4"',
      'Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="NULL" TargetMode="External"'))
    const bytes = zipSync(zip)
    const deck = parseDeck(bytes)
    const files = new Map<string, Uint8Array | string>()
    const warnings: string[] = []
    const ctx = { deck, pptx: bytes, slides: deck.slides, base: 'sample', add: (p: string, d: Uint8Array | string) => void files.set(p, d), progress: () => {}, warn: (m: string) => void warnings.push(m), signal: new AbortController().signal } as unknown as ExportContext
    await mediaBundle.run(ctx, { which: 'slides', audio: true, names: 'slides' })
    expect([...files.keys()]).toHaveLength(2)
    expect(text(files.get('MEDIA.txt'))).not.toMatch(/LINKED|NULL/)
    expect(warnings).toEqual([])
    expect(deck.slides[8].shapes.find((s) => s.kind === 'media')?.media?.part).toBe('ppt/media/media1.mp4')
  })

  it('lists an online film in MEDIA.txt and nothing else, since nothing is embedded', async () => {
    const r = await runFormat(mediaBundle, {}, { deck: FM + 'online-video.pptx' })
    expect([...r.files.keys()]).toEqual(['MEDIA.txt'])
    expect(text(r.files.get('MEDIA.txt'))).toMatch(/YouTube · Online video · https:\/\/www\.youtube\.com\/embed\//)
  })
})

/** The runner's media note for a format declaring `media`, on a fixture deck. */
function note(media: MediaHandling, deckName: string, group: FormatDef['group'] = 'images', slides?: number[]): string[] {
  const bytes = fixture(deckName)
  const deck = parseDeck(bytes)
  const warnings: string[] = []
  const ctx = { deck, pptx: bytes, slides: slides ? slides.map((i) => deck.slides[i]) : deck.slides, warn: (m: string) => void warnings.push(m) } as unknown as ExportContext
  noteMedia(ctx, { media, group } as FormatDef)
  return warnings
}

describe('media notes', () => {
  it('every format says what it does with video and sound', () => {
    for (const f of FORMATS) expect(['still', 'kept', 'plays', 'files', 'none'], f.id).toContain(f.media)
  })

  it('pictures: nothing plays, and the bundle saves the files', () => {
    expect(note('still', 'sample-16x9.pptx')).toEqual([
      'Slide 9 has video or sound. Nothing plays here: where this has a picture of it, that’s PowerPoint’s poster frame. Video bundle saves the files.',
    ])
    expect(note('still', 'sample-16x9.pptx', 'video')[0]).toMatch(/It doesn’t play, and there’s no sound/)
    expect(note('still', 'sample-16x9.pptx', 'images', [0, 1])).toEqual([])
  })

  it('PowerPoint copies: only what was never in the deck needs saying', () => {
    expect(note('kept', 'sample-16x9.pptx')).toEqual([])
    const w = note('kept', MEDIA)
    expect(w.find((m) => /internet connection/.test(m))).toMatch(/^Slides 3 and 5 play online video \(/)
    expect(w).toContain('Slide 5’s Flash video won’t play anywhere: Flash is retired.')
    expect(w.find((m) => /linked from the author’s disk/.test(m))).toMatch(/^Slide 4 has media linked .*, not embedded, so it plays only where that file is too\.$/)
  })

  it('the HTML5 page: layout video, trims and playback settings it doesn’t honour', () => {
    const w = note('plays', MEDIA)
    expect(w).toContain('Slides 7 and 8 have video or sound on the layout, which the page leaves out.')
    expect(w.find((m) => /trimmed in PowerPoint/.test(m))).toMatch(/^Slide 1’s “.+” is trimmed in PowerPoint \(from 1\.5 s to 0\.3 s before the end\), but the page plays the whole clip\.$/)
    expect(w.find((m) => /playback settings/.test(m))).toMatch(/^PowerPoint’s playback settings on slide 1 /)
  })

  it('takes PowerPoint’s default volume for a default, not a setting', () => {
    // sample-16x9.pptx's clip is vol="80000" and starts on a click, as PowerPoint inserts it.
    expect(note('plays', 'sample-16x9.pptx')).toEqual([])
  })

  it('text formats and the bundle itself say nothing', () => {
    expect(note('none', MEDIA)).toEqual([])
    expect(note('files', MEDIA)).toEqual([])
  })
})

describe('parse', () => {
  /** sample-16x9.pptx with one slide's XML changed. */
  function patched(slide: number, edit: (xml: string) => string) {
    const zip = unzipSync(fixture('sample-16x9.pptx'))
    const part = `ppt/slides/slide${slide}.xml`
    zip[part] = strToU8(edit(strFromU8(zip[part])))
    return parseDeck(zipSync(zip))
  }

  it('a video’s timing node is not an animation', () => {
    expect(parseDeck(fixture('sample-16x9.pptx')).slides[8].hasAnimations).toBe(false)
    const animated = patched(9, (x) => x.replace('<p:cTn id="2" fill="hold" display="0">', '<p:cTn id="2" fill="hold" presetClass="entr" presetID="10">'))
    expect(animated.slides[8].hasAnimations).toBe(true)
  })

  it('reads the older quickTimeFile spelling as video', () => {
    const d = patched(9, (x) => x.replace('<a:videoFile ', '<a:quickTimeFile '))
    expect(d.slides[8].shapes.find((s) => s.kind === 'media')?.media?.kind).toBe('video')
  })
})
