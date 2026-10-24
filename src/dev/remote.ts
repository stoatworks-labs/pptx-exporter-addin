// Development only: drive the pane from the shell.
//
// Testing an exporter in real PowerPoint means dozens of exports with
// different formats and options, and clicking them through a task pane is
// slow and error-prone. In a dev build the pane polls the dev server for a
// queued job (scripts/pane.sh puts one there), runs it with the output sent
// to .out/, and posts the report and any console errors to .out/pane.log.
//
// Imported behind `import.meta.env.DEV`, so none of this is in a production
// build — check with `grep -c __control dist/assets/*.js` (expect 0).

export interface RemoteJob {
  op: 'export' | 'info' | 'probe' | 'reload'
  /** probe: slide widths to ask PowerPoint for (16:9), and canvas sizes to try. */
  widths?: number[]
  canvases?: [number, number][]
  /** probe: read the deck through getFileAsync this many times, timing each. */
  file?: number
  /** Format id → option values (missing values take the defaults). */
  formats?: Record<string, Record<string, string | number | boolean>>
  scope?: 'all' | 'selected' | 'range' | 'section'
  range?: string
  section?: string
  includeHidden?: boolean
  engine?: 'auto' | 'powerpoint' | 'pdf'
  base?: string
}

export interface RemoteHandlers {
  info(): Promise<unknown>
  probe(job: RemoteJob): Promise<unknown>
  /** `log` writes a line to .out/pane.log as the job goes, for finding where a long one stalls. */
  exportJob(job: RemoteJob, log: (kind: string, data: unknown) => void): Promise<unknown>
}

export function log(kind: string, data: unknown): void {
  void fetch('/__log', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ t: new Date().toISOString(), kind, data }),
  }).catch(() => {})
}

export function startRemote(h: RemoteHandlers): void {
  const origError = console.error
  console.error = (...args: unknown[]) => {
    origError(...args)
    log('console.error', args.map((a) => (a instanceof Error ? `${a.name}: ${a.message}\n${a.stack}` : String(a))))
  }
  window.addEventListener('error', (e) => log('window.error', `${e.message} @ ${e.filename}:${e.lineno}`))
  window.addEventListener('unhandledrejection', (e) => log('unhandledrejection', String(e.reason?.stack ?? e.reason)))
  log('hello', { ua: navigator.userAgent, href: location.href })

  let busy = false
  setInterval(async () => {
    if (busy) return
    let job: RemoteJob | null = null
    try {
      const r = await fetch('/__control', { cache: 'no-store' })
      if (r.status !== 200) return
      job = (await r.json()) as RemoteJob
    } catch {
      return
    }
    if (job.op === 'reload') {
      log('reload', location.href)
      location.reload()
      return
    }
    busy = true
    try {
      log(`${job.op}:start`, job)
      const out = job.op === 'info' ? await h.info() : job.op === 'probe' ? await h.probe(job) : await h.exportJob(job, log)
      log(`${job.op}:done`, out)
    } catch (e) {
      log(`${job.op}:failed`, String((e as Error)?.stack ?? e))
    } finally {
      busy = false
    }
  }, 1500)
}
