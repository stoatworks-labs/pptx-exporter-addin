import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDeck, resolvePart } from '../src/deck/parse'
import { aspectLabel, slideText } from '../src/deck/model'

const fixture = (name: string) => new Uint8Array(readFileSync(resolve(__dirname, 'fixtures', name)))

describe('resolvePart', () => {
  it('walks relative targets', () => {
    expect(resolvePart('ppt/slides/slide1.xml', '../media/image1.png')).toBe('ppt/media/image1.png')
    expect(resolvePart('ppt/presentation.xml', 'slides/slide1.xml')).toBe('ppt/slides/slide1.xml')
    expect(resolvePart('', 'ppt/presentation.xml')).toBe('ppt/presentation.xml')
    expect(resolvePart('ppt/slides/slide1.xml', '/ppt/media/a.png')).toBe('ppt/media/a.png')
  })
})

describe('parseDeck — sample-16x9', () => {
  const deck = parseDeck(fixture('sample-16x9.pptx'))

  it('reads size, title, author and every slide in order', () => {
    expect(deck.width).toBe(12192000)
    expect(deck.height).toBe(6858000)
    expect(aspectLabel(deck.width, deck.height)).toBe('16:9')
    expect(deck.title).toBe('Advanced Exporter Test Deck')
    expect(deck.author).toBe('Fixture Author')
    expect(deck.slides.map((s) => s.title)).toEqual([
      'Advanced Exporter Test Deck', 'Agenda', 'Two Columns', 'A Table', 'A Picture',
      'Hidden Slide', 'Fade and Advance', 'Specials: 50% & $5 #1', 'A Video',
    ])
  })

  it('knows which slide is hidden', () => {
    expect(deck.slides.filter((s) => s.hidden).map((s) => s.index)).toEqual([5])
  })

  it('reads speaker notes, keeping blank lines and scripts', () => {
    expect(deck.slides[0].notesText).toBe('Welcome notes with ünïcödé ✓\nSecond line of the opening notes.')
    expect(deck.slides[1].notesText).toBe('Agenda notes.\n\nA paragraph after a blank line.')
    expect(deck.slides[8].notesText).toBe('日本語のノート。')
    expect(deck.slides[2].notesText).toBe('')
  })

  it('reads bullet levels and run formatting', () => {
    const body = deck.slides[1].shapes.find((s) => s.placeholder?.type === 'obj')!
    expect(body.paragraphs!.map((p) => p.level)).toEqual([0, 1, 2, 0])
    expect(body.paragraphs!.every((p) => p.bullet === 'char')).toBe(true)
    const runs = body.paragraphs![1].runs
    expect(runs.find((r) => r.text === 'bold')?.bold).toBe(true)
    expect(runs.find((r) => r.text === 'italic')?.italic).toBe(true)
    expect(body.paragraphs![3].runs[0].link).toEqual({ url: 'https://example.com/', tooltip: undefined })
  })

  it('inherits placeholder positions from the layout or master', () => {
    const title = deck.slides[1].shapes.find((s) => s.placeholder?.type === 'title')!
    expect(title.rect).toBeDefined()
    expect(title.rect!.w).toBeGreaterThan(deck.width / 2)
    const [left, right] = deck.slides[2].shapes.filter((s) => s.placeholder && s.placeholder.type !== 'title')
    expect(left.rect!.x).toBeLessThan(right.rect!.x)
  })

  it('reads a table, merged cells included', () => {
    const t = deck.slides[3].shapes.find((s) => s.kind === 'table')!.table!
    expect(t.rows.length).toBe(3)
    expect(t.colWidths.length).toBe(3)
    expect(t.rows[2][1].colSpan).toBe(2)
    expect(t.rows[2][2].merged).toBe(true)
    expect(slideText(deck.slides[3])).toContain('Projector\t2\tLaser')
  })

  it('reads a picture with its alt text and bytes', () => {
    const pic = deck.slides[4].shapes.find((s) => s.kind === 'picture')!
    expect(pic.description).toBe('Diagonal stripes and a circle')
    expect(pic.image!.contentType).toBe('image/png')
    const bytes = deck.readPart(pic.image!.part)!
    expect([...bytes.slice(1, 4)].map((c) => String.fromCharCode(c)).join('')).toBe('PNG')
  })

  it('reads transitions in both the plain and p14 forms', () => {
    expect(deck.slides[6].transition).toMatchObject({ type: 'fade', durationMs: 750, advanceAfterMs: 3000, advanceOnClick: true })
    expect(deck.slides[7].transition).toMatchObject({ type: 'push', durationMs: 1600, dir: 'u' })
    expect(deck.slides[0].transition).toBeUndefined()
  })

  it('resolves a slide-jump link to a slide index', () => {
    const body = deck.slides[7].shapes.find((s) => s.placeholder?.type === 'obj')!
    const linkRun = body.paragraphs![1].runs[0]
    expect(linkRun.link?.slide).toBe(1)
  })

  it('reads embedded video with its poster frame', () => {
    const v = deck.slides[8].shapes.find((s) => s.kind === 'media')!
    expect(v.media).toMatchObject({ kind: 'video', part: 'ppt/media/media1.mp4', contentType: 'video/mp4' })
    expect(v.media!.poster).toMatch(/^ppt\/media\/image\d\.png$/)
    expect(v.rect!.w).toBeGreaterThan(0)
    expect(deck.readPart(v.media!.part!)!.length).toBeGreaterThan(1000)
  })

  it('reads sections', () => {
    expect(deck.sections.map((s) => [s.name, s.slideIndexes])).toEqual([
      ['Intro', [0, 1]], ['Body', [2, 3, 4, 5]], ['Finale', [6, 7, 8]],
    ])
    expect(deck.slides[4].section).toBe('Body')
  })
})

describe('parseDeck — sample-4x3', () => {
  const deck = parseDeck(fixture('sample-4x3.pptx'))
  it('reads the 4:3 size and a loose text box', () => {
    expect(aspectLabel(deck.width, deck.height)).toBe('4:3')
    const tb = deck.slides[1].shapes.find((s) => !s.placeholder && s.kind === 'text')!
    expect(tb.paragraphs![0].bullet).toBe('none')
    expect(tb.paragraphs![0].runs[0].text).toBe('A loose text box')
  })
})
