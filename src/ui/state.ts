// What the pane remembers between sessions: the formats you last used and
// how you set them, and named presets. localStorage only — the pane's own
// origin, this machine — and every read survives it being unavailable
// (private windows, locked-down Office installs).

import type { Scope } from '../selection'
import type { Values } from '../formats/types'
import type { Engine } from '../render/images'

export interface PaneState {
  enabled: string[]
  values: Record<string, Values>
  scope: Scope
  range: string
  section: string
  includeHidden: boolean
  engine: 'auto' | Engine
  closed: string[]
}

export interface Preset {
  name: string
  enabled: string[]
  values: Record<string, Values>
}

const STATE_KEY = 'advx.state.v1'
const PRESETS_KEY = 'advx.presets.v1'

export const DEFAULT_STATE: PaneState = {
  enabled: ['pdf', 'png'],
  values: {},
  scope: 'all',
  range: '',
  section: '',
  includeHidden: false,
  engine: 'auto',
  closed: [],
}

function read<T>(key: string): T | null {
  try {
    const raw = localStorage.getItem(key)
    return raw ? (JSON.parse(raw) as T) : null
  } catch {
    return null
  }
}

function write(key: string, v: unknown): void {
  try {
    localStorage.setItem(key, JSON.stringify(v))
  } catch {
    /* storage unavailable: the pane still works, it just forgets */
  }
}

export function loadState(): PaneState {
  const s = read<Partial<PaneState>>(STATE_KEY)
  return { ...DEFAULT_STATE, ...(s ?? {}), values: s?.values ?? {} }
}

export function saveState(s: PaneState): void {
  write(STATE_KEY, s)
}

export function loadPresets(): Preset[] {
  const p = read<Preset[]>(PRESETS_KEY)
  return Array.isArray(p) ? p.filter((x) => x && typeof x.name === 'string') : []
}

export function savePresets(p: Preset[]): void {
  write(PRESETS_KEY, p)
}
