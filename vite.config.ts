import { defineConfig, type Plugin } from 'vite'
import { homedir } from 'node:os'
import { join, resolve, basename } from 'node:path'
import { readFileSync, existsSync, mkdirSync, createWriteStream, appendFileSync } from 'node:fs'

const ROOT = import.meta.dirname
// The font manager's media scanner, vendored by git subtree (`npm run
// core:pull`) and imported unchanged through this alias.
const FM = resolve(ROOT, 'vendor/pptx-font-manager/src')

// office-addin-dev-certs writes a trusted localhost cert here; Office requires
// HTTPS even in development.
const certDir = join(homedir(), '.office-addin-dev-certs')
const certFile = join(certDir, 'localhost.crt')
const keyFile = join(certDir, 'localhost.key')
// PLAIN_HTTP=1 serves plain http — for driving the pane in a browser that
// doesn't trust the dev cert. PowerPoint itself needs the https server.
const httpsReady = existsSync(certFile) && existsSync(keyFile) && !process.env.PLAIN_HTTP

// The release version, shown in the task pane's footer and on the install
// page. Distinct from the manifest's <Version>, which Office requires to be
// >= 1.0 and which only moves when the manifest itself changes.
const pkg = JSON.parse(readFileSync(resolve(ROOT, 'package.json'), 'utf8')) as { version: string }

// Ship manifest.xml at the root of the built site: it is the file a user
// sideloads, it hard-codes this origin, so the origin is the natural place to
// fetch it from. The repo copy is the source of truth.
function shipManifest(): Plugin {
  return {
    name: 'ship-manifest',
    apply: 'build',
    generateBundle() {
      this.emitFile({ type: 'asset', fileName: 'manifest.xml', source: readFileSync(resolve(ROOT, 'manifest.xml')) })
    },
  }
}

// Development only:
//   POST /__sink?name=x   writes the body to .out/x — the pane's "send output
//                         to the dev server" switch, so an export from real
//                         PowerPoint can be checked with pdfinfo, ffprobe and
//                         friends without a Save dialog in the way.
//   GET  /__fixtures/x    serves test/fixtures/x, for ?mock=/__fixtures/x.
//   /__control, /__log    a job queue and log for driving a dev pane from the
//                         shell — see src/dev/remote.ts and scripts/pane.sh.
function devRoutes(): Plugin {
  return {
    name: 'dev-routes',
    apply: 'serve',
    configureServer(server) {
      const jobs: string[] = []
      server.middlewares.use((req, res, next) => {
        const url = new URL(req.url ?? '/', 'https://localhost')
        if (url.pathname === '/__sink' && req.method === 'POST') {
          const name = basename(url.searchParams.get('name') ?? 'export.bin').replace(/[^\w.\- ()]/g, '_')
          const dir = resolve(ROOT, '.out')
          mkdirSync(dir, { recursive: true })
          const out = createWriteStream(join(dir, name))
          req.pipe(out)
          out.on('finish', () => {
            res.statusCode = 204
            res.end()
            server.config.logger.info(`sink: .out/${name}`)
          })
          out.on('error', (e) => {
            res.statusCode = 500
            res.end(String(e))
          })
          return
        }
        // The shell drives a dev pane through these (src/dev/remote.ts, scripts/pane.sh).
        if (url.pathname === '/__control') {
          if (req.method === 'POST') {
            let body = ''
            req.on('data', (c) => (body += c))
            req.on('end', () => {
              jobs.push(body)
              res.statusCode = 202
              res.end(`queued (${jobs.length})\n`)
            })
            return
          }
          const job = jobs.shift()
          res.statusCode = job ? 200 : 204
          res.setHeader('Cache-Control', 'no-store')
          res.end(job ?? '')
          return
        }
        if (url.pathname === '/__log' && req.method === 'POST') {
          let body = ''
          req.on('data', (c) => (body += c))
          req.on('end', () => {
            const dir = resolve(ROOT, '.out')
            mkdirSync(dir, { recursive: true })
            appendFileSync(join(dir, 'pane.log'), body + '\n')
            res.statusCode = 204
            res.end()
          })
          return
        }
        if (url.pathname.startsWith('/__fixtures/') && req.method === 'GET') {
          const file = resolve(ROOT, 'test/fixtures', basename(url.pathname))
          if (!existsSync(file)) {
            res.statusCode = 404
            res.end()
            return
          }
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation')
          res.end(readFileSync(file))
          return
        }
        next()
      })
    },
  }
}

export default defineConfig({
  resolve: { alias: { '@fm': FM } },
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  plugins: [shipManifest(), devRoutes()],
  // jSquash's codecs find their .wasm next to their own module; pre-bundling
  // them moves the module and loses the file.
  optimizeDeps: { exclude: ['@jsquash/webp', '@jsquash/avif'] },
  worker: { format: 'es' },
  server: {
    port: 3000,
    strictPort: true,
    host: 'localhost',
    https: httpsReady ? { cert: readFileSync(certFile), key: readFileSync(keyFile) } : undefined,
    // NO_HMR=1: an edit no longer reloads the pane mid-export.
    hmr: process.env.NO_HMR ? false : undefined,
  },
  build: {
    outDir: 'dist',
    sourcemap: true,
    target: 'es2022',
    chunkSizeWarningLimit: 1500,
    rollupOptions: {
      input: {
        taskpane: resolve(ROOT, 'taskpane.html'),
        index: resolve(ROOT, 'index.html'),
      },
    },
  },
})
