// Which page of PowerPoint's PDF is which slide. Pure — unit tested.
//
// Exporters drop hidden slides, so a 9-slide deck with one hidden slide may
// come back as an 8-page PDF, and page 6 is then slide 7. (Presentation
// Converter's whole reason for existing; the same reconciliation, here.)

export type PdfMapping =
  | { kind: 'all'; pages: number }
  | { kind: 'visible'; pages: number }
  | { kind: 'mismatch'; pages: number; slides: number; visible: number }

export function mapPdf(pageCount: number, hidden: boolean[]): PdfMapping {
  const visible = hidden.filter((h) => !h).length
  if (pageCount === hidden.length) return { kind: 'all', pages: pageCount }
  if (pageCount === visible) return { kind: 'visible', pages: pageCount }
  return { kind: 'mismatch', pages: pageCount, slides: hidden.length, visible }
}

/** Zero-based PDF page for a slide, or null when that slide has no page. */
export function pageForSlide(m: PdfMapping, hidden: boolean[], slideIndex: number): number | null {
  if (m.kind === 'all') return slideIndex
  if (m.kind === 'visible') {
    if (hidden[slideIndex]) return null
    let page = 0
    for (let i = 0; i < slideIndex; i++) if (!hidden[i]) page++
    return page
  }
  return null
}
