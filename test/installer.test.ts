import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

// The installers do one thing each: put manifest.xml where PowerPoint looks,
// and take every copy of it out again. These run the Mac postinstall scripts
// against a pretend home folder, so the folder logic is checked on any machine
// (CI is Linux). Installer.app, Gatekeeper and PowerPoint itself are not — see
// the README's Installers section for what has been run for real.

const ROOT = resolve(__dirname, '..')
const manifest = readFileSync(resolve(ROOT, 'manifest.xml'), 'utf8')
const id = /<Id>([^<]+)<\/Id>/.exec(manifest)![1]!
// Font Manager's Id: the add-in most likely to share the wef folder.
const otherAddin = manifest.replace(id, 'abc06eaf-2365-4fb4-a75d-6c5591af1cf6').replace('Advanced Exporter', 'Another add-in')

let home: string
let wef: string

/** Lay a postinstall out the way Installer does: the script with the manifest beside it. */
function run(kind: 'install' | 'uninstall'): string {
  const scripts = mkdtempSync(join(tmpdir(), `ae-${kind}-`))
  copyFileSync(resolve(ROOT, `installer/mac/${kind}-postinstall`), join(scripts, 'postinstall'))
  copyFileSync(resolve(ROOT, 'manifest.xml'), join(scripts, 'manifest.xml'))
  try {
    return execFileSync('bash', [join(scripts, 'postinstall')], { env: { ...process.env, HOME: home }, encoding: 'utf8' })
  } finally {
    rmSync(scripts, { recursive: true, force: true })
  }
}

beforeEach(() => {
  home = mkdtempSync(join(tmpdir(), 'ae-home-'))
  wef = join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef')
})
afterEach(() => rmSync(home, { recursive: true, force: true }))

describe.runIf(process.platform !== 'win32')('Mac installer scripts', () => {
  it('installs the production manifest, verbatim, into the wef folder', () => {
    run('install')
    expect(readdirSync(wef)).toEqual(['stoatworks-advanced-exporter.xml'])
    expect(readFileSync(join(wef, 'stoatworks-advanced-exporter.xml'), 'utf8')).toBe(manifest)
  })

  it('replaces the sideload script’s copy rather than adding a second one with the same Id', () => {
    mkdirSync(wef, { recursive: true })
    writeFileSync(join(wef, 'pptx-exporter.manifest.xml'), manifest)
    run('install')
    expect(readdirSync(wef)).toEqual(['stoatworks-advanced-exporter.xml'])
  })

  it('never touches another add-in in the folder, installing or uninstalling', () => {
    mkdirSync(wef, { recursive: true })
    writeFileSync(join(wef, 'stoatworks-font-manager.xml'), otherAddin)
    run('install')
    run('uninstall')
    expect(readdirSync(wef)).toEqual(['stoatworks-font-manager.xml'])
    expect(readFileSync(join(wef, 'stoatworks-font-manager.xml'), 'utf8')).toBe(otherAddin)
  })

  it('uninstalls every copy, whatever it was called, and is safe to run twice', () => {
    mkdirSync(wef, { recursive: true })
    writeFileSync(join(wef, 'pptx-exporter.manifest.xml'), manifest)
    run('install')
    writeFileSync(join(wef, 'renamed-by-hand.xml'), manifest)
    run('uninstall')
    expect(readdirSync(wef)).toEqual([])
    expect(run('uninstall')).toMatch(/nothing to remove/)
  })

  it('uninstalls cleanly when PowerPoint was never set up', () => {
    expect(() => run('uninstall')).not.toThrow()
    expect(existsSync(wef)).toBe(false)
  })
})

describe('Mac installer distribution', () => {
  const dist = readFileSync(resolve(ROOT, 'installer/mac/install-distribution.xml'), 'utf8')

  it('checks for PowerPoint by its files, not system.applications', () => {
    // fromIdentifier() only finds a running app, so with PowerPoint closed Font
    // Manager's first build told people PowerPoint wasn't installed (2026-10-24).
    expect(dist).not.toMatch(/system\.applications\.\w+\(/)
    expect(dist).toMatch(/Library\/Containers\/com\.microsoft\.Powerpoint/)
  })
})

describe('Windows installer script', () => {
  const nsi = readFileSync(resolve(ROOT, 'installer/windows/advanced-exporter.nsi'), 'utf8')

  it('registers under the key Microsoft’s own sideload tool uses, per user', () => {
    // office-addin-dev-settings 4.0.1: a value named after the add-in Id, holding
    // the manifest path, under HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer.
    expect(nsi).toMatch(/!define WEF_KEY "Software\\Microsoft\\Office\\16\.0\\Wef\\Developer"/)
    expect(nsi).toMatch(/WriteRegStr HKCU "\$\{WEF_KEY\}" "\$\{ADDIN_ID\}" "\$INSTDIR\\manifest\.xml"/)
    expect(nsi).toMatch(/DeleteRegValue HKCU "\$\{WEF_KEY\}" "\$\{ADDIN_ID\}"/)
    expect(nsi).toMatch(/RequestExecutionLevel user/)
    expect(nsi).not.toMatch(/HKLM/)
  })

  it('has its own Installed apps entry, not Font Manager’s', () => {
    expect(nsi).toMatch(/!define UNINST_KEY "[^"]*\\StoatworksAdvancedExporterForPowerPoint"/)
  })
})
