// Draws a format's options from its schema. One renderer for every format,
// so a new option is a line in the format's definition, not a form to build.

import type { FormatDef, OptionDef, Values } from '../formats/types'

let uid = 0

export function renderOptions(def: FormatDef, values: Values, onChange: (v: Values) => void): HTMLElement {
  const root = document.createElement('div')
  const rows: { def: OptionDef; row: HTMLElement }[] = []

  const refresh = () => {
    for (const { def: o, row } of rows) row.hidden = o.when ? !o.when(values) : false
  }
  const set = (key: string, v: string | number | boolean) => {
    values[key] = v
    refresh()
    onChange(values)
  }

  for (const o of def.options) {
    const id = `opt-${def.id}-${o.key}-${++uid}`
    const row = document.createElement('div')
    row.className = `opt${o.type === 'toggle' ? ' toggle' : ''}`
    const label = document.createElement('label')
    label.htmlFor = id
    label.textContent = o.label
    const ctl = document.createElement('div')
    ctl.className = 'ctl'
    row.append(label, ctl)
    const cur = values[o.key]

    switch (o.type) {
      case 'select': {
        const sel = document.createElement('select')
        sel.id = id
        for (const c of o.choices) sel.add(new Option(c.label, c.value, false, c.value === cur))
        sel.addEventListener('change', () => set(o.key, sel.value))
        ctl.append(sel)
        break
      }
      case 'number': {
        const inp = document.createElement('input')
        inp.type = 'number'
        inp.id = id
        inp.min = String(o.min)
        inp.max = String(o.max)
        inp.step = String(o.step ?? 1)
        inp.value = String(cur)
        inp.addEventListener('change', () => {
          const n = Number(inp.value)
          const clamped = Number.isFinite(n) ? Math.min(o.max, Math.max(o.min, n)) : o.default
          inp.value = String(clamped)
          set(o.key, clamped)
        })
        ctl.append(inp)
        if (o.unit) ctl.append(unit(o.unit))
        break
      }
      case 'slider': {
        const inp = document.createElement('input')
        inp.type = 'range'
        inp.id = id
        inp.min = String(o.min)
        inp.max = String(o.max)
        inp.step = String(o.step ?? 1)
        inp.value = String(cur)
        const out = document.createElement('output')
        out.htmlFor.add(id)
        out.textContent = `${cur}${o.unit ?? ''}`
        inp.addEventListener('input', () => {
          out.textContent = `${inp.value}${o.unit ?? ''}`
        })
        inp.addEventListener('change', () => set(o.key, Number(inp.value)))
        ctl.append(inp, out)
        break
      }
      case 'toggle': {
        const inp = document.createElement('input')
        inp.type = 'checkbox'
        inp.id = id
        inp.checked = !!cur
        inp.addEventListener('change', () => set(o.key, inp.checked))
        ctl.append(inp)
        break
      }
      case 'color': {
        const inp = document.createElement('input')
        inp.type = 'color'
        inp.id = id
        inp.value = String(cur)
        inp.addEventListener('change', () => set(o.key, inp.value))
        ctl.append(inp)
        break
      }
      case 'text': {
        const inp = document.createElement('input')
        inp.type = 'text'
        inp.id = id
        inp.value = String(cur ?? '')
        inp.spellcheck = false
        if (o.placeholder) inp.placeholder = o.placeholder
        inp.addEventListener('change', () => set(o.key, inp.value))
        ctl.append(inp)
        break
      }
    }
    if (o.help) {
      const h = document.createElement('p')
      h.className = 'help'
      h.textContent = o.help
      row.append(h)
    }
    rows.push({ def: o, row })
    root.append(row)
  }
  refresh()
  return root
}

function unit(text: string): HTMLElement {
  const s = document.createElement('span')
  s.className = 'unit'
  s.textContent = text
  return s
}
