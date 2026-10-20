import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const manifest = readFileSync(resolve(__dirname, '../manifest.xml'), 'utf8')
const wrangler = readFileSync(resolve(__dirname, '../wrangler.toml'), 'utf8')

describe('manifest.xml is the production manifest', () => {
  const urls = [...manifest.matchAll(/(?:DefaultValue|Url)="(https?:[^"]+)"/g)].map((m) => m[1])
  it('never points at localhost', () => {
    // (The header comment mentions localhost; only the URLs matter.)
    expect(urls.filter((u) => /localhost|127\.0\.0\.1/.test(u))).toEqual([])
  })
  it('points every pane and icon URL at the hostname wrangler deploys to', () => {
    // A route pattern ("host/*") or a custom domain ("host").
    const host = /pattern\s*=\s*"([^"/]+)/.exec(wrangler)![1]
    const ours = urls.filter((u) => !u.includes('github.com'))
    expect(ours.length).toBeGreaterThan(5)
    for (const u of ours) expect(new URL(u).host).toBe(host)
  })
  it('has a manifest version Office accepts', () => {
    const v = /<Version>([^<]+)<\/Version>/.exec(manifest)![1]
    expect(Number(v.split('.')[0])).toBeGreaterThanOrEqual(1)
  })
})
