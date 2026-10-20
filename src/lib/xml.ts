// Small helpers over DOMParser for reading OOXML.
//
// Everything matches on localName, never on prefix: the same element turns up
// as <p:sp>, <ns0:sp> or a default-namespaced <sp> depending on which program
// last wrote the file, and PowerPoint is happy with all of them.

export const NS = {
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
} as const

const decoder = new TextDecoder('utf-8')

export function parseXml(src: string | Uint8Array): Document {
  const text = typeof src === 'string' ? src : decoder.decode(src)
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  const err = doc.getElementsByTagName('parsererror')[0]
  if (err) throw new Error(`Malformed XML: ${err.textContent?.slice(0, 200) ?? ''}`)
  return doc
}

/** Direct children with this localName. */
export function kids(el: Element | null | undefined, local: string): Element[] {
  if (!el) return []
  const out: Element[] = []
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.localName === local) out.push(c)
  }
  return out
}

/** First direct child with this localName. */
export function kid(el: Element | null | undefined, local: string): Element | null {
  if (!el) return null
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if (c.localName === local) return c
  }
  return null
}

/** Walk a chain of direct children: path(sp, 'nvSpPr', 'nvPr', 'ph'). */
export function path(el: Element | null | undefined, ...locals: string[]): Element | null {
  let cur: Element | null | undefined = el
  for (const l of locals) {
    cur = kid(cur, l)
    if (!cur) return null
  }
  return cur ?? null
}

/** All descendants with this localName, in document order. */
export function all(el: Element | Document | null | undefined, local: string): Element[] {
  if (!el) return []
  const out: Element[] = []
  const list = el.getElementsByTagName('*')
  for (let i = 0; i < list.length; i++) {
    if (list[i].localName === local) out.push(list[i])
  }
  return out
}

/** First descendant with this localName. */
export function first(el: Element | Document | null | undefined, local: string): Element | null {
  if (!el) return null
  const list = el.getElementsByTagName('*')
  for (let i = 0; i < list.length; i++) {
    if (list[i].localName === local) return list[i]
  }
  return null
}

/** A relationship attribute (r:id, r:embed, r:link), whatever its prefix. */
export function rAttr(el: Element | null | undefined, local: string): string | null {
  if (!el) return null
  const v = el.getAttributeNS(NS.r, local)
  if (v) return v
  for (const a of Array.from(el.attributes)) {
    if (a.localName === local && a.name.includes(':')) return a.value
  }
  return null
}

/** An attribute by localName (ignores any prefix). */
export function attrLocal(el: Element | null | undefined, local: string): string | null {
  if (!el) return null
  if (el.hasAttribute(local)) return el.getAttribute(local)
  for (const a of Array.from(el.attributes)) {
    if (a.localName === local) return a.value
  }
  return null
}

export function intAttr(el: Element | null | undefined, name: string): number | undefined {
  const v = el?.getAttribute(name)
  if (v == null || v === '') return undefined
  const n = parseInt(v, 10)
  return Number.isFinite(n) ? n : undefined
}

/** OOXML booleans: "1", "true", "on" are true. */
export function boolAttr(el: Element | null | undefined, name: string): boolean | undefined {
  const v = el?.getAttribute(name)
  if (v == null) return undefined
  return v === '1' || v === 'true' || v === 'on'
}

/** Escape text for XML content and attribute values. */
export function xmlEscape(s: string): string {
  return s.replace(/[<>&"']/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&apos;' })[c]!)
}

/** Strip characters XML 1.0 cannot carry at all (C0 controls other than tab/LF/CR). */
export function xmlSafe(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
}
