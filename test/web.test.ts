// The web exports: the data the player is given, the pages the formats
// write, and the player itself driven by keys and hash changes under jsdom.
// No canvas here, so the formats run with PNG pictures (PowerPoint's own PNG
// handed straight through) and transitions are cuts (jsdom has no
// Element.animate, which the player treats as "no animation").

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { encode as encodePng } from 'fast-png'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { parseDeck } from '../src/deck/parse'
import type { Deck } from '../src/deck/model'
import { normalise, type ExportContext, type FormatDef, type Values } from '../src/formats/types'
import {
  INLINE_MEDIA_CAP, buildDeckData, html5, jsonForScript, mapTransition, notesHtml, reveal, scrollPage,
  type DeckDataInput, type PlayerDeck,
} from '../src/formats/web'

// Vitest hands CSS imports back empty (it doesn't process CSS unless told
// to), ?raw included; a real build gives the file. Give the tests the file.
vi.mock('../player/player.css?raw', async () => {
  const fs = await import('node:fs')
  const path = await import('node:path')
  return { default: fs.readFileSync(path.resolve(__dirname, '../player/player.css'), 'utf8') }
})

const fixture = (name: string) => new Uint8Array(readFileSync(resolve(__dirname, 'fixtures', name)))
const PLAYER = readFileSync(resolve(__dirname, '../player/player.js'), 'utf8')
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0))

const OPTIONS: DeckDataInput['options'] = {
  controls: true, progress: true, numbers: true, loop: false, auto: 'off', autoSeconds: 5,
  transitions: 'deck', transitionMs: 0, theme: 'dark', skipHidden: true,
}

function data(deck: Deck, keep: number[], over: Partial<DeckDataInput> = {}): PlayerDeck {
  return buildDeckData({
    deck,
    slides: keep.map((i) => deck.slides[i]),
    title: deck.title,
    background: '#000000',
    width: 1920,
    height: 1080,
    src: (s) => `slides/${s.index + 1}.png`,
    notes: true,
    links: true,
    options: OPTIONS,
    ...over,
  })
}

interface Out {
  files: Map<string, Uint8Array | string>
  warnings: string[]
}

async function run(def: FormatDef, values: Partial<Values>, opts: { keep?: number[]; deck?: Deck; picture?: (index: number) => Uint8Array } = {}): Promise<Out> {
  const bytes = fixture('sample-16x9.pptx')
  const deck = opts.deck ?? parseDeck(bytes)
  const files = new Map<string, Uint8Array | string>()
  const warnings: string[] = []
  const ctx: ExportContext = {
    deck,
    slides: (opts.keep ?? deck.slides.map((s) => s.index)).map((i) => deck.slides[i]),
    base: 'sample',
    host: {} as ExportContext['host'],
    images: {
      framePng: async (index: number) => opts.picture?.(index) ?? PNG,
      frame: async () => { throw new Error('no canvas under vitest') },
    } as unknown as ExportContext['images'],
    add: (p, d) => void files.set(p, d),
    pdf: async () => new Uint8Array(),
    pptx: bytes,
    progress: () => {},
    warn: (m) => void warnings.push(m),
    signal: new AbortController().signal,
  }
  await def.run(ctx, normalise(def, values as Values))
  return { files, warnings }
}

const str = (d: Uint8Array | string | undefined) => (typeof d === 'string' ? d : new TextDecoder().decode(d))

function deckDataOf(html: string): PlayerDeck {
  const m = /<script type="application\/json" id="deck-data">([\s\S]*?)<\/script>/.exec(html)
  expect(m).not.toBeNull()
  return JSON.parse(m![1])
}

// ---------------------------------------------------------------- deck data

describe('buildDeckData', () => {
  const deck = parseDeck(fixture('sample-16x9.pptx'))

  it('numbers the export, not the deck, when slides are left out', () => {
    const d = data(deck, [1, 2, 3, 4, 6, 7, 8])
    expect(d.slides.map((s) => s.number)).toEqual([2, 3, 4, 5, 7, 8, 9])
    expect(d.slides[0].src).toBe('slides/2.png')
    expect(d.slides[0].title).toBe('Agenda')
  })

  it('points a slide link at the linked slide’s place in the export', () => {
    // Deck slide 8 links back to deck slide 2, which is first here.
    const d = data(deck, [1, 2, 3, 4, 6, 7, 8])
    expect(d.slides[5].links).toContainEqual(expect.objectContaining({ slide: 0 }))
    // And with everything kept it is second.
    expect(data(deck, deck.slides.map((s) => s.index)).slides[7].links).toContainEqual(expect.objectContaining({ slide: 1 }))
  })

  it('drops a link to a slide that isn’t exported', () => {
    const d = data(deck, [0, 2, 3, 4, 6, 7, 8])
    const links = d.slides.find((s) => s.number === 8)!.links ?? []
    expect(links.some((l) => l.slide != null)).toBe(false)
  })

  it('keeps external links, http(s) only', () => {
    const d = data(deck, [1])
    expect(d.slides[0].links).toContainEqual(expect.objectContaining({ href: 'https://example.com/' }))
  })

  it('writes no notes at all when they are off', () => {
    const d = data(deck, deck.slides.map((s) => s.index), { notes: false })
    expect(d.slides.some((s) => 'notes' in s)).toBe(false)
    expect(d.options.presenter).toBe(false)
    expect(JSON.stringify(d)).not.toContain('Welcome notes')
  })

  it('writes notes as escaped HTML when they are on', () => {
    const d = data(deck, [0, 1])
    expect(d.options.presenter).toBe(true)
    expect(d.slides[0].notes).toBe('<p>Welcome notes with ünïcödé ✓</p><p>Second line of the opening notes.</p>')
    expect(d.slides[1].notes).toBe('<p>Agenda notes.</p><p><br></p><p>A paragraph after a blank line.</p>')
    expect(notesHtml([{ level: 0, bullet: 'none', runs: [{ text: '<b>&"x"</b>', bold: true }] }])).toBe('<p><strong>&lt;b&gt;&amp;&quot;x&quot;&lt;/b&gt;</strong></p>')
    expect(notesHtml([{ level: 0, bullet: 'none', runs: [{ text: 'bad', link: { url: 'javascript:alert(1)' } }] }])).toBe('<p>bad</p>')
  })

  it('marks hidden slides and carries transitions and timings', () => {
    const d = data(deck, deck.slides.map((s) => s.index))
    expect(d.slides[5].hidden).toBe(true)
    expect(d.slides[0].hidden).toBeUndefined()
    expect(d.slides[6].transition).toEqual({ type: 'fade', ms: 750 })
    expect(d.slides[6].advanceMs).toBe(3000)
    expect(d.slides[7].transition).toEqual({ type: 'slide', ms: 1600, dir: 'up' })
    expect(d.slides[0].transition).toBeUndefined()
  })

  it('maps PowerPoint transitions onto fade, slide or none', () => {
    const t = (type: string, dir?: string) => mapTransition({ type, durationMs: 500, dir, advanceOnClick: true })
    expect(t('dissolve').type).toBe('fade')
    expect(t('wipe', 'r')).toEqual({ type: 'slide', ms: 500, dir: 'right' })
    expect(t('cover', 'ld').dir).toBe('left')
    expect(t('split', 'in').dir).toBe('left')
    expect(t('vortex').type).toBe('fade')
    expect(t('cut').type).toBe('none')
    expect(mapTransition(undefined).type).toBe('none')
  })

  it('places embedded video where the deck has it', () => {
    const d = data(deck, [8], { media: () => ({ src: 'media/media1.mp4', poster: 'media/image2.png', type: 'video/mp4' }) })
    const m = d.slides[0].media![0]
    expect(m).toMatchObject({ kind: 'video', src: 'media/media1.mp4', type: 'video/mp4' })
    expect(m.w).toBeGreaterThan(0)
    expect(m.x + m.w).toBeLessThanOrEqual(1.0001)
  })
})

describe('jsonForScript', () => {
  it('cannot close the script element it sits in', () => {
    const value = { title: '</script><script>alert(1)</script> <!-- &   ' }
    const out = jsonForScript(value)
    expect(out).not.toMatch(/<\/script/i)
    expect(out).not.toContain('<')
    expect(out).not.toContain(' ')
    expect(out).not.toContain(' ')
    expect(JSON.parse(out)).toEqual(value)
  })
})

// ---------------------------------------------------------------- formats

describe('html5 format', () => {
  it('writes a folder: page, player, pictures, media', async () => {
    const { files, warnings } = await run(html5, { image: 'png', notes: true })
    expect(warnings).toEqual([])
    expect([...files.keys()].sort()).toEqual([
      'index.html', 'media/image2.png', 'media/media1.mp4', 'player.css', 'player.js',
      'slides/01.png', 'slides/02.png', 'slides/03.png', 'slides/04.png', 'slides/05.png',
      'slides/06.png', 'slides/07.png', 'slides/08.png', 'slides/09.png',
    ])
    const html = str(files.get('index.html'))
    expect(html).toMatch(/^<!doctype html>/)
    expect(html).toContain('<meta charset="utf-8">')
    expect(html).toContain('<title>Advanced Exporter Test Deck</title>')
    expect(html).toContain('<noscript><div class="ap-fallback">')
    expect(html).toContain('<img src="slides/01.png"')
    const d = deckDataOf(html)
    expect(d.slides).toHaveLength(9)
    expect(d.width).toBe(1920)
    expect(d.slides[8].media?.[0].src).toBe('media/media1.mp4')
    expect(d.slides[8].media?.[0].poster).toBe('media/image2.png')
    expect(files.get('media/media1.mp4')).toBeInstanceOf(Uint8Array)
    expect(str(files.get('player.css'))).toContain('.ap-stage {')
    expect(str(files.get('player.js'))).toContain('window.AdvancedPlayer = api')
  })

  it('leaves the notes out of every file when they are off', async () => {
    const { files } = await run(html5, { image: 'png', notes: false })
    for (const [, body] of files) if (typeof body === 'string' && body.includes('deck-data')) expect(body).not.toContain('Welcome notes')
    expect(deckDataOf(str(files.get('index.html'))).options.presenter).toBe(false)
  })

  it('writes one self-contained file, each picture once', async () => {
    const { files, warnings } = await run(html5, { image: 'png', packaging: 'single' }, { keep: [0, 1, 8] })
    expect(warnings).toEqual([])
    expect([...files.keys()]).toEqual(['sample.html'])
    const html = str(files.get('sample.html'))
    expect(html).not.toContain('src="player.js"')
    expect(html).not.toContain('href="player.css"')
    expect(html).toContain('window.AdvancedPlayer')
    expect(html).toContain('.ap-stage {')
    expect(html).toContain('<img id="s1" src="data:image/png;base64,')
    expect(html.match(/data:image\/png;base64,iVBOR/g)!.length).toBe(3 + 1) // three slides + the video's poster
    const d = deckDataOf(html)
    expect(d.slides.map((s) => s.src)).toEqual(['#s1', '#s2', '#s3'])
    expect(d.slides[2].media?.[0].src).toMatch(/^data:video\/mp4;base64,/)
  })

  it('leaves out media past the single-file allowance, and says so', async () => {
    const deck = parseDeck(fixture('sample-16x9.pptx'))
    const read = deck.readPart.bind(deck)
    deck.readPart = (part) => (part.endsWith('.mp4') ? new Uint8Array(INLINE_MEDIA_CAP + 1) : read(part))
    const { files, warnings } = await run(html5, { image: 'png', packaging: 'single' }, { deck, keep: [8] })
    expect(warnings).toHaveLength(1)
    expect(warnings[0]).toMatch(/Slide 9: the video .* use the folder packaging/)
    expect(deckDataOf(str(files.get('sample.html'))).slides[0].media).toBeUndefined()
  })

  it('warns about linked media rather than pretending', async () => {
    const deck = parseDeck(fixture('sample-16x9.pptx'))
    const v = deck.slides[8].shapes.find((s) => s.kind === 'media')!
    v.media = { kind: 'video', external: 'file:///Users/someone/Movies/clip.mp4' }
    const { warnings } = await run(html5, { image: 'png' }, { deck, keep: [8] })
    expect(warnings).toEqual([expect.stringMatching(/linked rather than embedded \(file:\/\/\/Users\/someone\/Movies\/clip\.mp4\)/)])
  })

  it('calls an online film online, not linked', async () => {
    const deck = parseDeck(fixture('sample-16x9.pptx'))
    const v = deck.slides[8].shapes.find((s) => s.kind === 'media')!
    v.media = { kind: 'video', external: 'https://www.youtube.com/embed/abc?feature=oembed' }
    const { warnings } = await run(html5, { image: 'png' }, { deck, keep: [8] })
    expect(warnings).toEqual([expect.stringMatching(/is an online video \(YouTube\), not a file in the deck/)])
  })
})

describe('reveal format', () => {
  it('writes sections with background pictures from the pinned CDN', async () => {
    const { files } = await run(reveal, { image: 'png', notes: true }, { keep: [0, 1, 7] })
    expect([...files.keys()].sort()).toEqual(['index.html', 'slides/01.png', 'slides/02.png', 'slides/08.png'])
    const html = str(files.get('index.html'))
    expect(html).toContain('https://cdn.jsdelivr.net/npm/reveal.js@6.0.2/dist/reveal.js')
    expect(html).toContain('https://cdn.jsdelivr.net/npm/reveal.js@6.0.2/dist/plugin/notes.js')
    expect(html.match(/<section /g)).toHaveLength(3)
    expect(html).toContain('data-background-image="slides/01.png" data-background-size="contain" data-background-color="#000000"')
    expect(html).toContain('<aside class="notes"><p>Welcome notes with ünïcödé ✓</p>')
    expect(html).toContain('Reveal.initialize(')
    expect(html).toContain('"hashOneBasedIndex":true')
    // Slide 8's link back to slide 2 lands on the second section.
    expect(html).toContain('href="#/2" data-slide="1"')
  })

  it('copies reveal.js in for offline use — or says it can’t', async () => {
    const installed = existsSync(resolve(__dirname, '../node_modules/reveal.js/dist/reveal.js'))
    const going = run(reveal, { image: 'png', notes: true, source: 'offline' }, { keep: [0] })
    if (!installed) {
      await expect(going).rejects.toThrow(/doesn’t include reveal\.js 6\.0\.2.*jsDelivr/)
      return
    }
    const { files } = await going
    expect([...files.keys()].sort()).toEqual(['index.html', 'reveal/LICENSE', 'reveal/plugin/notes.js', 'reveal/reveal.css', 'reveal/reveal.js', 'slides/01.png'])
    const html = str(files.get('index.html'))
    expect(html).toContain('<script src="reveal/reveal.js"></script>')
    expect(html).not.toContain('cdn.jsdelivr.net')
  })

  it('leaves the notes and the notes plugin out when notes are off', async () => {
    const { files } = await run(reveal, { image: 'png', notes: false }, { keep: [0] })
    const html = str(files.get('index.html'))
    expect(html).not.toContain('class="notes"')
    expect(html).not.toContain('notes.js')
    expect(html).not.toContain('Welcome notes')
  })
})

describe('scrolling page', () => {
  it('puts every slide on one page with its notes', async () => {
    const { files } = await run(scrollPage, { image: 'png' }, { keep: [0, 1, 2] })
    const html = str(files.get('index.html'))
    expect(html.match(/<article class="slide"/g)).toHaveLength(3)
    expect(html).toContain('loading="lazy"')
    expect(html).toContain('<h3>Notes</h3><p>Welcome notes')
    expect(html).toContain('@media print')
    expect(html).toContain('href="https://example.com/"')
  })

  it('can leave the notes out, and be one file', async () => {
    const { files } = await run(scrollPage, { image: 'png', notes: false, packaging: 'single' }, { keep: [0] })
    expect([...files.keys()]).toEqual(['sample.html'])
    const html = str(files.get('sample.html'))
    expect(html).not.toContain('Welcome notes')
    expect(html).toContain('src="data:image/png;base64,')
  })
})

// ---------------------------------------------------------------- the player

const GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACw='

interface Player {
  mode: string
  index(): number
  count: number
  blank(): string
  overlay(): string
  destroy(): void
}

declare global {
  interface Window {
    AdvancedPlayer?: { current: Player | null }
  }
}

function fakeDeck(over: Partial<PlayerDeck['options']> = {}, slides?: Partial<PlayerDeck['slides'][number]>[]): PlayerDeck {
  const base = slides ?? [{}, {}, {}, {}]
  return {
    version: 1,
    title: 'Fake deck',
    aspect: 16 / 9,
    width: 16,
    height: 9,
    background: '#000000',
    slides: base.map((s, i) => ({
      src: `${GIF}#${i + 1}`,
      alt: ['One', 'Two', 'Three', 'Four'][i] ?? `S${i + 1}`,
      title: ['One', 'Two', 'Three', 'Four'][i] ?? `S${i + 1}`,
      number: i + 1,
      text: `Words on slide ${i + 1}`,
      notes: `<p>Notes for ${i + 1}</p>`,
      ...s,
    })),
    options: { ...OPTIONS, presenter: true, ...over },
  }
}

function boot(deck: PlayerDeck, hash = ''): Player {
  document.body.innerHTML = `<div id="player"></div><script type="application/json" id="deck-data">${jsonForScript(deck)}</script>`
  history.replaceState(null, '', `${location.pathname}${hash}`)
  new Function(PLAYER)()
  const p = window.AdvancedPlayer!.current
  expect(p).not.toBeNull()
  return p!
}

function key(k: string, init: KeyboardEventInit = {}): void {
  document.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init }))
}

const shown = () => document.querySelector<HTMLImageElement>('.ap-frame .ap-slide img')!.getAttribute('src')
const live = () => document.querySelector('.ap-audience > [aria-live]')!.textContent

describe('player', () => {
  afterEach(() => {
    window.AdvancedPlayer?.current?.destroy()
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('starts on the first slide and says where it is', () => {
    const p = boot(fakeDeck())
    expect(p.mode).toBe('audience')
    expect(p.index()).toBe(0)
    expect(location.hash).toBe('#/1')
    expect(live()).toBe('Slide 1 of 4: One')
    expect(shown()).toBe(`${GIF}#1`)
    expect(document.querySelectorAll('.ap-frame .ap-slide')).toHaveLength(1)
    expect(document.querySelector('.ap-counter')!.textContent).toBe('1 / 4')
    // The words of every slide, for screen readers and Find.
    expect(document.querySelector('[aria-label="Text of the slides"]')!.textContent).toContain('Words on slide 3')
  })

  it('steps with the arrow keys, Space, N and P, and keeps the hash in step', () => {
    const p = boot(fakeDeck())
    key('ArrowRight')
    expect(p.index()).toBe(1)
    expect(location.hash).toBe('#/2')
    expect(shown()).toBe(`${GIF}#2`)
    key(' ')
    expect(p.index()).toBe(2)
    key('ArrowLeft')
    key('p')
    expect(p.index()).toBe(0)
    key('n')
    key('PageDown')
    expect(p.index()).toBe(2)
    key('Backspace')
    expect(p.index()).toBe(1)
  })

  it('goes to the ends with Home and End, and stops there without loop', () => {
    const p = boot(fakeDeck())
    key('End')
    expect(p.index()).toBe(3)
    key('ArrowRight')
    expect(p.index()).toBe(3)
    expect(document.querySelector('.ap-btn[aria-label="Next slide"]')!.getAttribute('aria-disabled')).toBe('true')
    key('Home')
    expect(p.index()).toBe(0)
    key('ArrowLeft')
    expect(p.index()).toBe(0)
  })

  it('wraps round with loop', () => {
    const p = boot(fakeDeck({ loop: true }))
    key('ArrowLeft')
    expect(p.index()).toBe(3)
    key('ArrowRight')
    expect(p.index()).toBe(0)
  })

  it('jumps to a typed number on Enter', () => {
    const p = boot(fakeDeck())
    key('3')
    expect(p.index()).toBe(0)
    expect(document.querySelector('.ap-goto')!.textContent).toContain('3')
    key('Enter')
    expect(p.index()).toBe(2)
    key('9')
    key('Enter')
    expect(p.index()).toBe(2)
  })

  it('opens on the slide the hash names, and follows hash changes', async () => {
    const p = boot(fakeDeck(), '#/3')
    expect(p.index()).toBe(2)
    expect(shown()).toBe(`${GIF}#3`)
    const changed = new Promise((r) => window.addEventListener('hashchange', r, { once: true }))
    location.hash = '#/4'
    await changed
    expect(p.index()).toBe(3)
    history.replaceState(null, '', '#/2')
    window.dispatchEvent(new PopStateEvent('popstate'))
    expect(p.index()).toBe(1)
  })

  it('steps over hidden slides, which a number still reaches', () => {
    const p = boot(fakeDeck({}, [{}, { hidden: true }, {}, {}]))
    key('ArrowRight')
    expect(p.index()).toBe(2)
    key('ArrowLeft')
    expect(p.index()).toBe(0)
    key('2')
    key('Enter')
    expect(p.index()).toBe(1)
  })

  it('opens and closes the overview, and jumps from it', () => {
    const p = boot(fakeDeck())
    const ov = document.querySelector<HTMLElement>('.ap-overview')!
    expect(ov.hidden).toBe(true)
    key('o')
    expect(ov.hidden).toBe(false)
    expect(p.overlay()).toBe('overview')
    const thumbs = ov.querySelectorAll<HTMLButtonElement>('.ap-thumb')
    expect(thumbs).toHaveLength(4)
    expect(thumbs[0].getAttribute('aria-current')).toBe('true')
    expect(document.activeElement).toBe(thumbs[0])
    key('ArrowRight') // moves focus in the grid, not the slide
    expect(p.index()).toBe(0)
    expect(document.activeElement).toBe(thumbs[1])
    key('Escape')
    expect(ov.hidden).toBe(true)
    key('Escape')
    expect(ov.hidden).toBe(false)
    thumbs[2].click()
    expect(ov.hidden).toBe(true)
    expect(p.index()).toBe(2)
  })

  it('blanks to black or white, and a key brings the slide back without moving', () => {
    const p = boot(fakeDeck())
    const blank = document.querySelector<HTMLElement>('.ap-blank')!
    key('b')
    expect(blank.hidden).toBe(false)
    expect(blank.dataset.mode).toBe('black')
    key('ArrowRight')
    expect(blank.hidden).toBe(true)
    expect(p.index()).toBe(0)
    key('w')
    expect(blank.dataset.mode).toBe('white')
    key('w')
    expect(blank.hidden).toBe(true)
  })

  it('follows slide links and leaves external ones to the browser', () => {
    const p = boot(fakeDeck({}, [
      { links: [{ x: 0.1, y: 0.1, w: 0.2, h: 0.1, slide: 3, label: 'To four' }, { x: 0.5, y: 0.5, w: 0.2, h: 0.1, href: 'https://example.com/', label: 'Site' }, { href: 'javascript:alert(1)', label: 'Nope' }] },
      {}, {}, {},
    ]))
    const links = document.querySelectorAll<HTMLAnchorElement>('.ap-frame a.ap-link')
    expect(links).toHaveLength(2)
    expect(links[1].target).toBe('_blank')
    expect(links[1].rel).toBe('noopener noreferrer')
    expect(links[0].style.left).toBe('10%')
    links[0].click()
    expect(p.index()).toBe(3)
  })

  it('advances on a click on the slide', () => {
    const p = boot(fakeDeck())
    document.querySelector<HTMLElement>('.ap-stage img')!.click()
    expect(p.index()).toBe(1)
  })

  it('advances on its own every few seconds, and stops at the end', () => {
    vi.useFakeTimers()
    const p = boot(fakeDeck({ auto: 'every', autoSeconds: 2 }))
    vi.advanceTimersByTime(1999)
    expect(p.index()).toBe(0)
    vi.advanceTimersByTime(1)
    expect(p.index()).toBe(1)
    vi.advanceTimersByTime(10000)
    expect(p.index()).toBe(3)
  })

  it('follows the deck’s own timings, and holds while the overview is open', () => {
    vi.useFakeTimers()
    const p = boot(fakeDeck({ auto: 'deck' }, [{ advanceMs: 1000 }, {}, {}, {}]))
    key('o')
    vi.advanceTimersByTime(5000)
    expect(p.index()).toBe(0)
    key('Escape')
    vi.advanceTimersByTime(1000)
    expect(p.index()).toBe(1)
    vi.advanceTimersByTime(60000)
    expect(p.index()).toBe(1) // slide 2 has no timing: it waits for a click
  })

  it('moves when the presenter window says so', () => {
    const deck = fakeDeck()
    const p = boot(deck)
    const key = `${deck.title}|${deck.slides.length}|${deck.slides[0].src.slice(0, 64)}`
    window.dispatchEvent(new MessageEvent('message', { data: { ns: 'advanced-exporter-player', key, from: 'other', seq: 1, type: 'state', body: { index: 2, blank: '' } } }))
    expect(p.index()).toBe(2)
    // Someone else's message (another deck) is ignored.
    window.dispatchEvent(new MessageEvent('message', { data: { ns: 'advanced-exporter-player', key: 'x', from: 'other', seq: 2, type: 'state', body: { index: 0 } } }))
    expect(p.index()).toBe(2)
  })

  it('says so when the presenter window is blocked', () => {
    const open = vi.spyOn(window, 'open').mockReturnValue(null)
    boot(fakeDeck())
    key('s')
    expect(open).toHaveBeenCalledWith(expect.stringMatching(/#presenter$/), expect.any(String), expect.any(String))
    expect(document.querySelector('.ap-toast')!.textContent).toMatch(/blocked/)
    open.mockRestore()
  })

  it('has no presenter view when the notes were left out', () => {
    boot(fakeDeck({ presenter: false }))
    expect(document.querySelector('.ap-btn[aria-label="Presenter view (S)"]')).toBeNull()
  })

  describe('presenter view', () => {
    beforeEach(() => vi.spyOn(window, 'open').mockReturnValue(null))

    it('shows the current and next slides with the notes, and moves', () => {
      const p = boot(fakeDeck(), '#presenter')
      expect(p.mode).toBe('presenter')
      const imgs = document.querySelectorAll<HTMLImageElement>('.ap-presenter .pv-box img')
      expect(imgs[0].getAttribute('src')).toBe(`${GIF}#1`)
      expect(imgs[1].getAttribute('src')).toBe(`${GIF}#2`)
      expect(document.querySelector('.pv-notes-body')!.innerHTML).toBe('<p>Notes for 1</p>')
      key('End')
      expect(p.index()).toBe(3)
      expect(document.querySelector<HTMLElement>('.pv-end')!.hidden).toBe(false)
      expect(document.querySelector('.pv-notes-body')!.innerHTML).toBe('<p>Notes for 4</p>')
    })
  })
})

describe('player on a real export', () => {
  // jsdom has no media playback; the player pauses a slide's video on the way out.
  beforeEach(() => vi.spyOn(HTMLMediaElement.prototype, 'pause').mockImplementation(() => {}))
  afterEach(() => {
    window.AdvancedPlayer?.current?.destroy()
    vi.restoreAllMocks()
  })

  it('plays the html5 export’s own data, links and video included', async () => {
    const { files } = await run(html5, { image: 'png', notes: true })
    const d = deckDataOf(str(files.get('index.html')))
    const p = boot(d)
    expect(p.count).toBe(9)
    key('8')
    key('Enter')
    expect(p.index()).toBe(7)
    // Slide 8's link back to slide 2 is in the text layer as a real link.
    const back = document.querySelector<HTMLAnchorElement>('section[data-index="7"] a[data-slide="1"]')
    expect(back).not.toBeNull()
    back!.click()
    expect(p.index()).toBe(1)
    // Slide 6 is hidden: stepping from 5 goes to 7.
    key('5')
    key('Enter')
    key('ArrowRight')
    expect(p.index()).toBe(6)
    key('ArrowRight')
    key('ArrowRight')
    expect(p.index()).toBe(8)
    const video = document.querySelector<HTMLVideoElement>('.ap-frame video')!
    expect(video.getAttribute('src')).toBe('media/media1.mp4')
    expect(video.controls).toBe(true)
  })
})

// SAMPLES_DIR=/somewhere npx vitest run test/web.test.ts writes each web
// export there, with a coloured stand-in for each slide, to open in a browser.
describe.runIf(!!process.env.SAMPLES_DIR)('samples for a browser', () => {
  /** 480 × 270, a colour per slide, n white bars for slide n, a stripe to show movement. */
  const stand_in = (index: number): Uint8Array => {
    const W = 480, H = 270
    const px = new Uint8Array(W * H * 4)
    const hue = (index * 67) % 360
    const [r, g, b] = [0, 120, 240].map((o) => Math.round(127 + 100 * Math.cos(((hue - o) * Math.PI) / 180)))
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const bar = y > 40 && y < 120 && x > 40 && x < 40 + (index + 1) * 24 && (x - 40) % 24 < 16
        const stripe = Math.abs(x - y * 1.6) < 6
        const v = bar || stripe ? [255, 255, 255] : [r, g, b]
        px.set([...v, 255], (y * W + x) * 4)
      }
    }
    return encodePng({ width: W, height: H, data: px, channels: 4 })
  }
  const write = (dir: string, files: Map<string, Uint8Array | string>) => {
    for (const [name, data] of files) {
      const p = resolve(process.env.SAMPLES_DIR!, dir, name)
      mkdirSync(dirname(p), { recursive: true })
      writeFileSync(p, data)
    }
  }

  it('writes them', async () => {
    // The fixture has no whole-shape links; give slide 5's picture one, so
    // there is a click area to line up against the picture.
    const deck = parseDeck(fixture('sample-16x9.pptx'))
    deck.slides[4].shapes.find((s) => s.kind === 'picture')!.link = { slide: 0, tooltip: 'Back to the start' }
    const all = { picture: stand_in, deck }
    write('html5-folder', (await run(html5, { image: 'png', size: 'native', notes: true, transitions: 'deck', auto: 'off' }, all)).files)
    write('html5-single', (await run(html5, { image: 'png', size: 'native', packaging: 'single', notes: true, theme: 'light', loop: 'yes' }, all)).files)
    write('reveal', (await run(reveal, { image: 'png', size: 'native', notes: true }, all)).files)
    write('scroll', (await run(scrollPage, { image: 'png', size: 'native' }, all)).files)
  })
})

describe('reveal.js integrity hashes', () => {
  it('match the installed reveal.js, which jsDelivr serves byte for byte', async () => {
    const { createHash } = await import('node:crypto')
    const { readFileSync } = await import('node:fs')
    const { resolve } = await import('node:path')
    const { REVEAL_SRI, REVEAL_VERSION } = await import('../src/formats/web')
    const pkg = JSON.parse(readFileSync(resolve(__dirname, '../node_modules/reveal.js/package.json'), 'utf8'))
    expect(pkg.version).toBe(REVEAL_VERSION)
    for (const [file, hash] of Object.entries(REVEAL_SRI)) {
      const bytes = readFileSync(resolve(__dirname, '../node_modules/reveal.js/dist', file))
      expect(`sha384-${createHash('sha384').update(bytes).digest('base64')}`).toBe(hash)
    }
  })
})
