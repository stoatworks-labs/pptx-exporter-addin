// Advanced Exporter — task pane entry point.
//
// Finds out what it's running in (PowerPoint, or a plain browser for
// development), reads the deck, draws the format list and each chosen
// format's options from its schema, and runs the export.

import type { Deck } from './deck/model'
import { aspectLabel, EMU_PER_INCH } from './deck/model'
import { parseDeck } from './deck/parse'
import { runJob } from './export'
import { FORMATS, formatById } from './formats'
import { safeName } from './formats/common'
import { Cancelled, defaults, GROUPS, normalise, type FormatDef, type Values } from './formats/types'
import { MockHost } from './host/mock'
import { OfficeHost } from './host/office'
import type { Host } from './host/types'
import { deliver } from './output/deliver'
import { selectSlides, type Scope } from './selection'
import { renderOptions } from './ui/options'
import type { RemoteJob } from './dev/remote'
import { canEncodeNatively } from './render/canvas'
import { loadPresets, loadState, savePresets, saveState, type PaneState } from './ui/state'

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T

let host: Host | null = null
let deck: Deck | null = null
let selected: number[] = []
let state: PaneState = loadState()
let presets = loadPresets()
let running: AbortController | null = null
const unavailable = new Map<string, string>()

// --- Start-up -------------------------------------------------------------------

async function boot(): Promise<void> {
  $('version').textContent = `v${__APP_VERSION__}`
  if (import.meta.env.DEV) $('sink-wrap').hidden = false
  wireStatic()
  host = await findHost()
  applyTheme()
  if (!host) return
  await checkAvailability()
  renderFormatList()
  renderCards()
  restoreControls()
  await loadDeck()
  if (import.meta.env.DEV) {
    const { startRemote } = await import('./dev/remote')
    startRemote({ info: remoteInfo, exportJob: remoteExport, probe: remoteProbe })
  }
}

async function findHost(): Promise<Host | null> {
  const info = await officeReady()
  if (info?.host === Office.HostType.PowerPoint) {
    const h = new OfficeHost()
    $('host').textContent = `${h.label} · PowerPointApi ${h.caps.apiVersion}${h.caps.slideImages ? '' : ' — slide pictures come from its PDF'}`
    trackSelection()
    return h
  }
  const params = new URLSearchParams(location.search)
  if (import.meta.env.DEV || params.has('mock')) {
    $('mock-open').hidden = false
    const url = params.get('mock')
    if (url) {
      try {
        const m = await MockHost.fromUrl(url)
        $('host').textContent = m.label
        return m
      } catch (e) {
        setStatus((e as Error).message)
      }
    }
    $('deck').textContent = 'No deck open.'
    $<HTMLInputElement>('mock-file').addEventListener('change', async (ev) => {
      const f = (ev.target as HTMLInputElement).files?.[0]
      if (!f) return
      host = new MockHost(new Uint8Array(await f.arrayBuffer()), f.name.replace(/\.pptx$/i, ''))
      $('host').textContent = host.label
      await checkAvailability()
      renderFormatList()
      renderCards()
      restoreControls()
      await loadDeck()
    })
    return null
  }
  $('deck').textContent = 'This page is the add-in’s task pane. Open it from inside PowerPoint.'
  $('host').innerHTML = 'How to install it: <a href="/">the install page</a>.'
  return null
}

/** Office.onReady, but give up if office.js never arrived (offline, blocked). */
function officeReady(): Promise<{ host: Office.HostType | null } | null> {
  return new Promise((resolve) => {
    if (typeof Office === 'undefined') return resolve(null)
    const t = setTimeout(() => resolve(null), 8000)
    Office.onReady((info) => {
      clearTimeout(t)
      resolve(info as { host: Office.HostType | null })
    })
  })
}

function applyTheme(): void {
  try {
    const bg = Office.context?.officeTheme?.bodyBackgroundColor
    if (!bg) return
    const n = parseInt(bg.replace('#', ''), 16)
    const lum = 0.2126 * ((n >> 16) & 255) + 0.7152 * ((n >> 8) & 255) + 0.0722 * (n & 255)
    document.documentElement.dataset.theme = lum < 128 ? 'dark' : 'light'
  } catch {
    /* no Office theme: the system scheme applies */
  }
}

function trackSelection(): void {
  const refresh = async () => {
    if (!host) return
    try {
      selected = await host.selectedSlides()
    } catch {
      selected = []
    }
    updateSelectionNote()
  }
  void refresh()
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => void refresh())
  } catch {
    /* older hosts: the selection is read again at export */
  }
}

async function checkAvailability(): Promise<void> {
  unavailable.clear()
  if (!host) return
  const c = host.caps
  for (const f of FORMATS) {
    let why: string | null = null
    if (f.needs.images && !c.slideImages && !c.pdf) why = "This PowerPoint can't draw slide pictures for add-ins."
    else if (f.needs.pdf && !c.pdf) why = "This PowerPoint can't hand add-ins a PDF."
    else if (f.needs.exportSlide && !c.exportSlide) why = host.kind === 'mock' ? 'Needs PowerPoint (PowerPointApi 1.8).' : 'Needs a newer PowerPoint (PowerPointApi 1.8).'
    if (!why && f.unavailable) {
      try {
        why = await f.unavailable()
      } catch (e) {
        why = (e as Error).message
      }
    }
    if (why) unavailable.set(f.id, why)
  }
}

// --- Deck -------------------------------------------------------------------------

async function loadDeck(): Promise<void> {
  if (!host) return
  setStatus('Reading the deck…')
  try {
    if (host instanceof MockHost) deck = host.deck
    else {
      const bytes = await host.pptx((p) => setStatus(`Reading the deck… ${p.done}/${p.total}`))
      deck = parseDeck(bytes, host.documentName() || 'Presentation')
    }
    if (host.caps.selection) selected = await host.selectedSlides().catch(() => [])
    showDeck()
    setStatus('')
  } catch (e) {
    setStatus(`Couldn't read the deck: ${(e as Error).message}`)
  }
  updateExportButton()
}

function showDeck(): void {
  if (!deck) return
  const hidden = deck.slides.filter((s) => s.hidden).length
  const w = deck.width / EMU_PER_INCH
  const h = deck.height / EMU_PER_INCH
  $('deck').textContent = `${deck.title} — ${deck.slides.length} slide${deck.slides.length === 1 ? '' : 's'}${hidden ? ` (${hidden} hidden)` : ''}, ${aspectLabel(deck.width, deck.height)}, ${w.toFixed(2)} × ${h.toFixed(2)} in`
  const sel = $<HTMLSelectElement>('section')
  sel.innerHTML = ''
  for (const s of deck.sections) sel.add(new Option(`${s.name} (${s.slideIndexes.length})`, s.name, false, s.name === state.section))
  $('scope-section-wrap').hidden = deck.sections.length === 0
  if (!$<HTMLInputElement>('base').value) $<HTMLInputElement>('base').value = safeName(host?.documentName() || deck.title || 'Presentation')
  updateSelectionNote()
}

// --- Slides --------------------------------------------------------------------------

function wireStatic(): void {
  for (const r of document.querySelectorAll<HTMLInputElement>('input[name="scope"]')) {
    r.addEventListener('change', () => {
      state.scope = r.value as Scope
      persist()
      updateScopeRows()
      updateSelectionNote()
    })
  }
  $<HTMLInputElement>('range').addEventListener('input', (e) => {
    state.range = (e.target as HTMLInputElement).value
    persist()
    updateSelectionNote()
  })
  $<HTMLSelectElement>('section').addEventListener('change', (e) => {
    state.section = (e.target as HTMLSelectElement).value
    persist()
    updateSelectionNote()
  })
  $<HTMLInputElement>('hidden').addEventListener('change', (e) => {
    state.includeHidden = (e.target as HTMLInputElement).checked
    persist()
    updateSelectionNote()
  })
  $<HTMLSelectElement>('engine').addEventListener('change', (e) => {
    state.engine = (e.target as HTMLSelectElement).value as PaneState['engine']
    persist()
  })
  $('export').addEventListener('click', () => void onExport())
  $('preset').addEventListener('change', onPresetPick)
  $('preset-save').addEventListener('click', onPresetSave)
  $('preset-delete').addEventListener('click', onPresetDelete)
  renderPresets()
}

function restoreControls(): void {
  for (const r of document.querySelectorAll<HTMLInputElement>('input[name="scope"]')) r.checked = r.value === state.scope
  $<HTMLInputElement>('range').value = state.range
  $<HTMLInputElement>('hidden').checked = state.includeHidden
  $<HTMLSelectElement>('engine').value = state.engine
  updateScopeRows()
}

function updateScopeRows(): void {
  $('range-row').hidden = state.scope !== 'range'
  $('section-row').hidden = state.scope !== 'section'
}

function currentSelection() {
  if (!deck) return null
  if (state.scope === 'section' && !state.section && deck.sections[0]) state.section = deck.sections[0].name
  return selectSlides(deck, { scope: state.scope, range: state.range, section: state.section, selected, includeHidden: state.includeHidden })
}

function updateSelectionNote(): void {
  const sel = currentSelection()
  const note = $('selection-note')
  if (!sel) {
    note.textContent = ''
    return
  }
  if (sel.error) note.textContent = sel.error
  else {
    const nums = sel.slides.map((s) => s.index + 1)
    const list = nums.length > 12 ? `${nums.slice(0, 10).join(', ')} … ${nums[nums.length - 1]}` : nums.join(', ')
    note.textContent = `${sel.slides.length} slide${sel.slides.length === 1 ? '' : 's'}: ${list}.${sel.note ? ` ${sel.note}` : ''}`
  }
  updateExportButton()
}

// --- Formats ----------------------------------------------------------------------

function renderFormatList(): void {
  const root = $('format-list')
  root.innerHTML = ''
  for (const g of GROUPS) {
    const fs = FORMATS.filter((f) => f.group === g.id)
    if (!fs.length) continue
    const box = document.createElement('div')
    box.className = 'group'
    const h = document.createElement('h3')
    h.textContent = g.label
    const chips = document.createElement('div')
    chips.className = 'chips'
    for (const f of fs) chips.append(chip(f))
    box.append(h, chips)
    root.append(box)
  }
}

function chip(f: FormatDef): HTMLElement {
  const why = unavailable.get(f.id)
  const label = document.createElement('label')
  label.className = `chip${why ? ' off' : ''}`
  label.title = why ?? f.blurb
  const cb = document.createElement('input')
  cb.type = 'checkbox'
  cb.checked = !why && state.enabled.includes(f.id)
  cb.disabled = !!why
  cb.addEventListener('change', () => {
    state.enabled = cb.checked ? [...new Set([...state.enabled, f.id])] : state.enabled.filter((id) => id !== f.id)
    persist()
    renderCards()
    updateExportButton()
  })
  const name = document.createElement('span')
  name.textContent = f.label
  label.append(cb, name)
  return label
}

function enabledFormats(): FormatDef[] {
  return FORMATS.filter((f) => state.enabled.includes(f.id) && !unavailable.has(f.id))
}

function renderCards(): void {
  const root = $('cards')
  root.innerHTML = ''
  for (const f of enabledFormats()) {
    state.values[f.id] = normalise(f, state.values[f.id])
    const card = document.createElement('article')
    card.className = `fcard${state.closed.includes(f.id) ? ' closed' : ''}`
    const head = document.createElement('header')
    head.setAttribute('role', 'button')
    head.tabIndex = 0
    head.setAttribute('aria-expanded', String(!state.closed.includes(f.id)))
    head.innerHTML = '<span class="tw" aria-hidden="true">▾</span><h3></h3><span class="ext"></span>'
    head.querySelector('h3')!.textContent = f.label
    head.querySelector('.ext')!.textContent = f.ext
    const toggle = () => {
      const closed = card.classList.toggle('closed')
      head.setAttribute('aria-expanded', String(!closed))
      state.closed = closed ? [...state.closed, f.id] : state.closed.filter((x) => x !== f.id)
      persist()
    }
    head.addEventListener('click', toggle)
    head.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault()
        toggle()
      }
    })
    const body = document.createElement('div')
    body.className = 'body'
    const blurb = document.createElement('p')
    blurb.className = 'blurb'
    blurb.textContent = f.blurb
    body.append(blurb, renderOptions(f, state.values[f.id], (v) => {
      state.values[f.id] = v
      persist()
    }))
    const tools = document.createElement('div')
    tools.className = 'tools'
    const reset = document.createElement('button')
    reset.type = 'button'
    reset.className = 'link'
    reset.textContent = 'Back to defaults'
    reset.addEventListener('click', () => {
      state.values[f.id] = defaults(f)
      persist()
      renderCards()
    })
    tools.append(reset)
    body.append(tools)
    card.append(head, body)
    root.append(card)
  }
}

// --- Presets -------------------------------------------------------------------------

function renderPresets(): void {
  const sel = $<HTMLSelectElement>('preset')
  sel.innerHTML = '<option value="">—</option>'
  for (const p of presets) sel.add(new Option(p.name, p.name))
  $<HTMLButtonElement>('preset-delete').disabled = true
}

function onPresetPick(): void {
  const name = $<HTMLSelectElement>('preset').value
  $<HTMLButtonElement>('preset-delete').disabled = !name
  const p = presets.find((x) => x.name === name)
  if (!p) return
  state.enabled = p.enabled.filter((id) => formatById(id))
  state.values = { ...state.values, ...structuredClone(p.values) }
  persist()
  renderFormatList()
  renderCards()
  updateExportButton()
  $<HTMLInputElement>('preset-name').value = p.name
}

function onPresetSave(): void {
  const name = $<HTMLInputElement>('preset-name').value.trim()
  if (!name) {
    setStatus('Give the preset a name first.')
    return
  }
  const values: Record<string, Values> = {}
  for (const id of state.enabled) if (state.values[id]) values[id] = state.values[id]
  presets = [...presets.filter((p) => p.name !== name), { name, enabled: [...state.enabled], values: structuredClone(values) }]
  presets.sort((a, b) => a.name.localeCompare(b.name))
  savePresets(presets)
  renderPresets()
  $<HTMLSelectElement>('preset').value = name
  $<HTMLButtonElement>('preset-delete').disabled = false
  setStatus(`Saved preset “${name}”.`)
}

function onPresetDelete(): void {
  const name = $<HTMLSelectElement>('preset').value
  if (!name) return
  presets = presets.filter((p) => p.name !== name)
  savePresets(presets)
  renderPresets()
  setStatus(`Deleted preset “${name}”.`)
}

// --- Export ------------------------------------------------------------------------

function updateExportButton(): void {
  const btn = $<HTMLButtonElement>('export')
  if (running) return
  const n = enabledFormats().length
  const sel = currentSelection()
  btn.disabled = !host || !deck || !n || !sel || !!sel.error
  btn.textContent = n ? `Export ${n} format${n === 1 ? '' : 's'}` : 'Choose a format'
}

function setStatus(text: string): void {
  $('status').textContent = text
}

function setProgress(f: number | null): void {
  $('progress').hidden = f == null
  if (f != null) $('bar').style.width = `${Math.round(f * 100)}%`
}

async function onExport(): Promise<void> {
  if (running) {
    running.abort()
    return
  }
  if (!host) return
  const btn = $<HTMLButtonElement>('export')
  const result = $('result')
  result.hidden = true
  running = new AbortController()
  btn.textContent = 'Cancel'
  btn.classList.add('cancel')
  setProgress(0)
  try {
    // Read the deck again: it may have changed since the pane opened. The old
    // one goes first: a parsed deck holds its whole file, and two copies of a
    // deck full of video is more than the web view can take.
    if (!(host instanceof MockHost)) deck = null
    setStatus('Reading the deck from PowerPoint…')
    const bytes = await host.pptx((p) => setStatus(`Reading the deck from PowerPoint… ${p.done}/${p.total}`))
    deck = host instanceof MockHost ? host.deck : parseDeck(bytes, host.documentName() || 'Presentation')
    if (host.caps.selection) selected = await host.selectedSlides().catch(() => selected)
    showDeck()
    const sel = currentSelection()
    if (!sel || sel.error) throw new Error(sel?.error ?? 'No deck.')
    const engine = state.engine === 'auto' ? (host.caps.slideImages ? 'powerpoint' : 'pdf') : state.engine
    const base = safeName($<HTMLInputElement>('base').value || deck.title)
    const formats = enabledFormats().map((def) => ({ def, values: normalise(def, state.values[def.id]) }))
    const report = await runJob(host, deck, bytes, { formats, slides: sel.slides, base, engine }, {
      progress: (f, label) => {
        setProgress(f)
        setStatus(label)
      },
    }, running.signal)
    if (!report.file) throw new Error(report.errors.map((e) => `${e.format}: ${e.message}`).join('\n') || 'Nothing was produced.')
    setStatus('Handing the file over…')
    const outcome = await deliver(report.file, { sink: $<HTMLInputElement>('sink').checked })
    showResult(report, outcome)
    setStatus('')
  } catch (e) {
    if (!deck) void loadDeck() // the read failed after the old deck was let go
    if (e instanceof Cancelled) setStatus('Cancelled.')
    else {
      result.hidden = false
      result.className = 'card bad'
      result.innerHTML = '<h3>Export failed</h3><p></p>'
      result.querySelector('p')!.textContent = (e as Error).message
      setStatus('')
      console.error(e)
    }
  } finally {
    running = null
    btn.classList.remove('cancel')
    setProgress(null)
    updateExportButton()
  }
}

function showResult(report: Awaited<ReturnType<typeof runJob>>, outcome: string): void {
  const r = $('result')
  r.hidden = false
  r.className = `card${report.errors.length ? ' bad' : ''}`
  const size = report.file ? report.file.blob.size : 0
  const mb = size > 1048576 ? `${(size / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(size / 1024))} KB`
  r.innerHTML = '<h3></h3><p></p>'
  r.querySelector('h3')!.textContent = outcome === 'blocked'
    ? 'The web view blocked the download'
    : outcome === 'sunk' ? `Sent ${report.file!.name} to the dev server` : `Exported ${report.file!.name}`
  r.querySelector('p')!.textContent = `${report.fileCount} file${report.fileCount === 1 ? '' : 's'}, ${mb}, in ${report.seconds.toFixed(1)} s.`
  const items = [
    ...report.errors.map((e) => ({ cls: 'err', text: `${e.format} failed: ${e.message}` })),
    ...report.warnings.map((w) => ({ cls: 'wrn', text: `${w.format}: ${w.message}` })),
  ]
  if (items.length) {
    const ul = document.createElement('ul')
    for (const it of items) {
      const li = document.createElement('li')
      li.className = it.cls
      li.textContent = it.text
      ul.append(li)
    }
    r.append(ul)
  }
}

// --- Development: driven from the shell (src/dev/remote.ts) ---------------------

async function remoteInfo(): Promise<unknown> {
  const codecs: Record<string, boolean> = {}
  if (typeof VideoEncoder !== 'undefined') {
    for (const c of ['avc1.640028', 'hvc1.1.6.L120.B0', 'av01.0.08M.08', 'vp09.00.40.08', 'vp8']) {
      try {
        codecs[c] = !!(await VideoEncoder.isConfigSupported({ codec: c, width: 1920, height: 1080, bitrate: 8e6, framerate: 30 })).supported
      } catch {
        codecs[c] = false
      }
    }
  }
  return {
    host: host?.label,
    caps: host?.caps,
    deck: deck && { title: deck.title, slides: deck.slides.length, hidden: deck.slides.filter((x) => x.hidden).map((x) => x.index + 1), sections: deck.sections.map((x) => x.name), width: deck.width, height: deck.height, lang: deck.lang },
    selected,
    unavailable: Object.fromEntries(unavailable),
    nativeEncoders: { webp: await canEncodeNatively('image/webp'), avif: await canEncodeNatively('image/avif'), jpeg: await canEncodeNatively('image/jpeg') },
    videoEncoder: typeof VideoEncoder !== 'undefined',
    codecs,
    offscreenCanvas: typeof OffscreenCanvas !== 'undefined',
    ua: navigator.userAgent,
  }
}

/** Find the limits: how big a slide PowerPoint will draw, how big a canvas
 * this web view will make and encode. */
async function remoteProbe(job: RemoteJob): Promise<unknown> {
  const out: Record<string, unknown> = {}
  for (const w of job.widths ?? []) {
    const h = Math.round((w * 9) / 16)
    const t0 = performance.now()
    try {
      const png = await host!.renderSlide(0, w, h)
      const dv = new DataView(png.buffer, png.byteOffset)
      out[`render ${w}x${h}`] = `ok ${dv.getUint32(16)}x${dv.getUint32(20)} ${png.length} bytes ${Math.round(performance.now() - t0)} ms`
    } catch (e) {
      out[`render ${w}x${h}`] = `FAIL ${(e as Error).message} ${Math.round(performance.now() - t0)} ms`
    }
  }
  for (const [w, h] of job.canvases ?? []) {
    const t0 = performance.now()
    try {
      const c = new OffscreenCanvas(w, h)
      const g = c.getContext('2d')
      if (!g) throw new Error('no 2d context')
      g.fillStyle = '#c33'
      g.fillRect(0, 0, w, h)
      const px = g.getImageData(w - 1, h - 1, 1, 1).data
      const blob = await c.convertToBlob({ type: 'image/png' })
      out[`canvas ${w}x${h}`] = `ok pixel=${px[0]},${px[1]},${px[2]} png=${blob.size} ${Math.round(performance.now() - t0)} ms`
    } catch (e) {
      out[`canvas ${w}x${h}`] = `FAIL ${(e as Error).message} ${Math.round(performance.now() - t0)} ms`
    }
  }
  // How long getFileAsync takes to hand over the deck: the whole of it, every
  // time, before anything that reads media can start.
  for (let i = 1; i <= (job.file ?? 0); i++) {
    const t0 = performance.now()
    let first = 0
    let slices = 0
    try {
      const bytes = await host!.pptx((p) => {
        if (!first) first = performance.now() - t0
        slices = p.total
      })
      const s = (performance.now() - t0) / 1000
      out[`file #${i}`] = `ok ${bytes.length} bytes, ${slices} slices, first slice ${Math.round(first)} ms, ${s.toFixed(1)} s, ${(bytes.length / 1048576 / s).toFixed(1)} MB/s`
    } catch (e) {
      out[`file #${i}`] = `FAIL ${(e as Error).message} ${Math.round(performance.now() - t0)} ms`
    }
  }
  return out
}

async function remoteExport(job: RemoteJob, log: (kind: string, data: unknown) => void): Promise<unknown> {
  if (!host) throw new Error('No host.')
  setStatus('Remote export: reading the deck…')
  if (!(host instanceof MockHost)) deck = null
  let lastSlice = 0
  const bytes = await host.pptx((p) => {
    if (p.done - lastSlice >= 25 || p.done === p.total) log('export:read', `${(lastSlice = p.done)}/${p.total}`)
  })
  deck = host instanceof MockHost ? host.deck : parseDeck(bytes, host.documentName() || 'Presentation')
  if (host.caps.selection) selected = await host.selectedSlides().catch(() => selected)
  const sel = selectSlides(deck, { scope: job.scope ?? 'all', range: job.range ?? '', section: job.section ?? '', selected, includeHidden: job.includeHidden ?? false })
  if (sel.error) throw new Error(sel.error)
  const formats = Object.entries(job.formats ?? {}).map(([id, v]) => {
    const def = formatById(id)
    if (!def) throw new Error(`No format "${id}".`)
    if (unavailable.has(id)) throw new Error(`${def.label} is unavailable here: ${unavailable.get(id)}`)
    return { def, values: normalise(def, v as Values) }
  })
  const engine = !job.engine || job.engine === 'auto' ? (host.caps.slideImages ? 'powerpoint' : 'pdf') : job.engine
  const base = job.base ?? safeName(host.documentName() || deck.title)
  setProgress(0)
  try {
    let lastLabel = ''
    const report = await runJob(host, deck, bytes, { formats, slides: sel.slides, base, engine }, {
      progress: (f, label) => {
        setProgress(f)
        setStatus(label)
        if (label !== lastLabel) log('export:progress', (lastLabel = label))
      },
    }, new AbortController().signal)
    log('export:deliver', { name: report.file?.name, bytes: report.file?.blob.size })
    if (report.file) await deliver(report.file, { sink: true })
    setStatus(`Remote export done: ${report.file?.name ?? 'nothing'}`)
    return { file: report.file?.name, bytes: report.file?.blob.size, files: report.fileCount, seconds: report.seconds, slides: sel.slides.map((x) => x.index + 1), engine, warnings: report.warnings, errors: report.errors }
  } finally {
    setProgress(null)
  }
}

function persist(): void {
  saveState(state)
}

void boot()
