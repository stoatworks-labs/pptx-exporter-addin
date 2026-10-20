// Where the clickable and playable things sit on a slide, as fractions of the
// slide — for SVG link areas and the HTML player's hotspots and video
// overlays. A slide picture is just pixels; these put the links back.

import type { Deck, Link, MediaRef, Paragraph, Shape, Slide } from '../deck/model'

export interface Box {
  x: number
  y: number
  w: number
  h: number
}

export interface Hotspot extends Box {
  link: Link
  label: string
}

export interface MediaSpot extends Box {
  media: MediaRef
  label: string
}

function box(deck: Deck, s: Shape): Box | null {
  if (!s.rect || s.rect.w <= 0 || s.rect.h <= 0) return null
  return { x: s.rect.x / deck.width, y: s.rect.y / deck.height, w: s.rect.w / deck.width, h: s.rect.h / deck.height }
}

function linksIn(ps: Paragraph[] | undefined): Link[] {
  const out: Link[] = []
  for (const p of ps ?? []) for (const r of p.runs) if (r.link) out.push(r.link)
  return out
}

function sameLink(a: Link, b: Link): boolean {
  return a.url === b.url && a.slide === b.slide && a.jump === b.jump
}

/** Click areas. A shape's own action covers the shape. Linked text makes the
 * whole text box clickable only when every link in it goes to the same place
 * — a picture can't say where inside the box each word landed. */
export function hotspots(deck: Deck, slide: Slide): Hotspot[] {
  const out: Hotspot[] = []
  for (const s of slide.shapes) {
    const b = box(deck, s)
    if (!b) continue
    const label = s.description || s.name || 'Link'
    if (s.link) {
      out.push({ ...b, link: s.link, label })
      continue
    }
    const links = linksIn(s.paragraphs)
    if (links.length && links.every((l) => sameLink(l, links[0]))) {
      const text = (s.paragraphs ?? []).flatMap((p) => p.runs).filter((r) => r.link).map((r) => r.text).join('').trim()
      // Only when the linked text is most of the box — otherwise a click on
      // the unlinked words would follow it too.
      const all = (s.paragraphs ?? []).flatMap((p) => p.runs).map((r) => r.text).join('').trim()
      if (text.length >= all.length * 0.6) out.push({ ...b, link: links[0], label: text || label })
    }
  }
  return out
}

/** Every link on a slide (for link lists and text layers), wherever it is. */
export function allLinks(slide: Slide): { text: string; link: Link }[] {
  const out: { text: string; link: Link }[] = []
  for (const s of slide.shapes) {
    if (s.link) out.push({ text: s.description || s.name, link: s.link })
    const ps = s.paragraphs ?? s.table?.rows.flat().flatMap((c) => c.paragraphs)
    for (const p of ps ?? []) for (const r of p.runs) if (r.link) out.push({ text: r.text, link: r.link })
  }
  return out
}

export function mediaSpots(deck: Deck, slide: Slide): MediaSpot[] {
  const out: MediaSpot[] = []
  for (const s of slide.shapes) {
    if (s.kind !== 'media' || !s.media) continue
    const b = box(deck, s)
    if (b) out.push({ ...b, media: s.media, label: s.description || s.name || s.media.kind })
  }
  return out
}

/** A link as an address relative to a set of per-slide files. */
export function linkHref(link: Link, slideFile: (index: number) => string, currentIndex: number, count: number): string | null {
  if (link.url) return link.url
  if (link.slide != null) return slideFile(link.slide)
  switch (link.jump) {
    case 'first': return slideFile(0)
    case 'last': return slideFile(count - 1)
    case 'next': return currentIndex + 1 < count ? slideFile(currentIndex + 1) : null
    case 'previous': return currentIndex > 0 ? slideFile(currentIndex - 1) : null
    default: return null
  }
}

/** Only http(s), mailto and tel addresses go into generated pages as links. */
export function safeUrl(url: string): string | null {
  return /^(https?:|mailto:|tel:)/i.test(url.trim()) ? url.trim() : null
}
