# Advanced Exporter for PowerPoint

> **AI-assisted project.** This codebase was created with Claude (Anthropic), directed and
> reviewed by a human author. Every format has been exported from a real deck inside
> PowerPoint for Mac and checked with independent tools; PowerPoint for Windows and on the web
> haven't been tried yet. See [Status](#status).

PowerPoint will save a deck as a PDF, a picture per slide or a video, and it asks you
remarkably little about any of them. This add-in does all of those with the options left in
(size, resolution, quality, slide range, letterboxing) and then carries on into the formats
PowerPoint never got round to: a self-contained HTML5 presentation, reveal.js, LaTeX Beamer,
Markdown and Marp, handouts and notes pages, a deck flattened to pictures for a playback machine
that hasn't got your fonts, and a few more.

It's an Office task-pane add-in, so the same thing runs in PowerPoint for Mac, Windows and the
web. PowerPoint draws every slide itself, at exactly the size each format asks for; the add-in
reads the rest (notes, hidden slides, transitions, links, video) out of the deck. Everything
happens in the pane. Nothing is uploaded.

| Group | Formats |
| --- | --- |
| Images | PNG, JPEG, WebP, AVIF, TIFF (one per slide or multi-page; LZW, Deflate or PackBits), BMP, TGA, GIF, SVG with clickable links, contact sheet |
| Documents | PowerPoint's vector PDF (trimmed to a range or split per slide), picture PDF (keeps hidden slides and links), handouts 1/2/3/4/6/9 to a page, notes pages, Word notes document, OpenDocument |
| Web | HTML5 presentation with transitions, presenter view, links and the deck's own video; reveal.js; a scrolling page |
| Video | MP4 (H.264, HEVC, AV1) and WebM (VP9, VP8, AV1) with the deck's timings and captions from the notes; animated GIF |
| Text | LaTeX Beamer (the slides as drawn, or rebuilt as editable LaTeX) or a LaTeX article with notes; Markdown, Marp, Pandoc/Quarto; an outline PowerPoint reads back in (RTF), OPML or text; speaker notes; CSV; JSON |
| PowerPoint | A copy as PPTX, PPSX or POTX, trimmed and cleaned (notes, comments, hidden slides, names); one file per slide; a flattened picture deck |

Every option for every format is in the [user guide](docs/USER-GUIDE.md).

<!-- downloads:start -->

## Download

**[v0.1.0](https://github.com/stoatworks-labs/pptx-exporter-addin/releases/tag/v0.1.0)**

This release contains:

- [`manifest.xml`](https://github.com/stoatworks-labs/pptx-exporter-addin/releases/latest/download/manifest.xml) — Office add-in manifest (sideload into PowerPoint), 5 KB

All builds, checksums and release notes: [github.com/stoatworks-labs/pptx-exporter-addin/releases](https://github.com/stoatworks-labs/pptx-exporter-addin/releases).

<!-- downloads:end -->

## Status

- **Mac: checked.** On PowerPoint for Mac 16.112.3 (macOS 26.4.1, 2026-10-20), a nine-slide
  test deck with a hidden slide, sections, a table, a picture, links, transitions, notes in four
  scripts and an embedded video went through every format at once, inside PowerPoint, in 15.7
  seconds. Every file was checked with something other than the add-in (see
  [Testing](#testing)), and PowerPoint and Word opened the generated Office files without a
  repair prompt.
- **Windows and the web: untested.** Both are in the manifest's host list. Nothing here has run
  on them, because no machine here has desktop PowerPoint for Windows.
- **Hosted** at <https://pptx-exporter-addin.stoatworks-labs.com> since 2026-10-20. The
  production build was checked under its own CSP in a Chromium browser, then loaded in
  PowerPoint for Mac from the production manifest, where it opened and exported. The
  every-format run above used the dev build from `localhost`.

## Install

1. Download [`manifest.xml`](https://pptx-exporter-addin.stoatworks-labs.com/manifest.xml).
2. **PowerPoint for Mac:** copy it into
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`, quit PowerPoint and
   reopen it, then **Home ▸ Add-ins ▸ Developer Add-ins ▸ Advanced Exporter**.
3. **PowerPoint on the web, and Microsoft 365 on Windows:** **Home ▸ Add-ins ▸ More Add-ins ▸ My
   Add-ins ▸ Upload My Add-in**, and pick the file.

## Findings

Measured in PowerPoint for Mac 16.112.3 on 2026-10-20:

- **PowerPoint draws slides for add-ins quickly.** `Slide.getImageAsBase64` (PowerPointApi 1.8,
  and this build reports 1.10) drew nine 1080p slides in under a second; a 4K slide takes about
  0.2 s.
- **It won't draw past about 4,000 pixels.** A 16:9 slide 4032 px wide works; 4040 px answers
  `InvalidArgument`. The add-in finds that ceiling by bisection the first time it hits it, and
  draws anything bigger from PowerPoint's vector PDF through pdf.js, which comes out as sharp at
  8K as at 4K. A hidden slide isn't in that PDF, so at those sizes it is scaled up instead, and
  the export notes say so.
- **PowerPoint's PDF leaves hidden slides out** (`getFileAsync(Pdf)`, eight pages for nine
  slides). Pages are matched to slides around them, and a range that includes a hidden slide
  says it's missing rather than shifting every page after it by one.
- **WKWebView can't encode WebP or AVIF from a canvas.** Asked for WebP it quietly hands back a
  PNG. Both go through Squoosh's encoders compiled to WebAssembly (jSquash) instead.
- **WKWebView's WebCodecs encodes H.264, HEVC, AV1, VP9 and VP8** at 1080p. Video is encoded in
  the pane.
- **A 16384 × 9216 canvas works in the pane** (made, filled, read back and encoded as PNG), so
  that's the size cap, not the 8K I'd guessed.
- **pdf-lib's font subsetting blanks most of Noto Sans.** The text still extracts, but poppler
  draws nothing. The PDF fonts are subset at build time with fontTools instead and embedded
  whole, about 90 KB a PDF.

## Develop

```bash
npm install
npm run certs          # trust a localhost certificate (Office needs HTTPS); lasts 30 days
npm run dev            # https://localhost:3000
npm run sideload:mac   # a dev manifest in PowerPoint's wef folder, then quit and reopen PowerPoint
npm test               # vitest
npm run build          # typecheck and build to dist/
npm run validate       # Microsoft's manifest validator
```

The certificate from `npm run certs` expires after 30 days, and PowerPoint then refuses the dev
pane with "isn't signed by a valid security certificate". Run it again (`--days 90` makes it last
longer).

Outside PowerPoint the pane runs against a stand-in host that draws slides crudely itself, so
every format can be driven in a browser: open
`https://localhost:3000/taskpane.html?mock=/__fixtures/sample-16x9.pptx`, or run
`PLAIN_HTTP=1 npx vite --port 3001` for a browser that doesn't trust the dev certificate.

A dev pane, in PowerPoint or a browser, can be driven from the shell, which is how it was tested:

```bash
scripts/pane.sh info                     # host, API sets, codecs, the deck
scripts/pane.sh export '{"formats":{"png":{"size":"4k"},"pdf":{}},"includeHidden":true}'
scripts/pane.sh probe '{"widths":[3840,4096]}'
scripts/pane.sh reload
```

Output lands in `.out/` and the pane's log in `.out/pane.log`. None of this is in a production
build. Run the dev server with `NO_HMR=1` while doing this, or an edit reloads the pane in the
middle of an export.

### Testing

`npm test` covers the deck parser, the encoders, page layout, the text formats and the Office
packages against the decks in `test/fixtures/` (built by `npm run fixtures` with python-pptx).
Anything that needs a canvas or PowerPoint's renderer is checked in a browser or in PowerPoint,
through `scripts/pane.sh`, with the output checked by: libtiff (`tiffinfo -D`) and Pillow for
images, poppler (`pdfinfo`, `pdftoppm`, `pdffonts`) for PDFs, ffprobe and ffmpeg for video,
Tectonic for LaTeX, and `scripts/check-office.py` (python-pptx, python-docx, odfpy) for Office
files.

## Hosting and deploying

`manifest.xml` is the production manifest: every URL in it points at
`https://pptx-exporter-addin.stoatworks-labs.com`, and `test/manifest.test.ts` fails if one
points anywhere else or stops matching `wrangler.toml`. `npm run build` puts the pane, the
install page, the icons, `_headers` and a copy of the manifest in `dist/`.
[`deploy.yml`](.github/workflows/deploy.yml) publishes that as a Cloudflare Worker with static
assets on every push to main, then checks the live origin is serving the build it just made;
`cf-run npm run deploy` does the same from a laptop. The workflow needs the repo secret
`CLOUDFLARE_API_TOKEN` and the variable `CLOUDFLARE_ACCOUNT_ID`.

Releasing: bump `package.json` (and the lockfile), commit, tag `vX.Y.Z` and push the tag.
[`release.yml`](.github/workflows/release.yml) checks the tag against `package.json`, validates
the manifest and attaches it to a GitHub release. The CSP in
[`public/_headers`](public/_headers) lets Office.js load from Microsoft's CDN and lets PowerPoint
on the web frame the pane; `npm run serve:dist` serves `dist/` with those headers applied.

## Layout

```
manifest.xml         the production manifest
taskpane.html        the pane; index.html is the install page
src/taskpane.ts      start-up, the deck, the format list, export
src/deck/            reading the .pptx: text, notes, hidden slides, transitions, links, media
src/host/            PowerPoint through Office.js, and a stand-in for browsers
src/render/          slide pictures (PowerPoint, or its PDF through pdf.js), sizing, canvas
src/formats/         one module per group of formats; types.ts is the contract
src/encode/          TIFF, BMP, TGA, GIF, WebP/AVIF (jSquash)
src/pdf/             pdf-lib helpers: Unicode text, bookmarks, links
src/output/          zipping and handing the file over
player/              the HTML5 presentation player (plain JS, copied into exports)
src/dev/             the shell-driven harness (development builds only)
fonts-src/           full Noto Sans; scripts/make-pdf-fonts.sh subsets it into src/assets/fonts
test/                vitest; fixtures/ are python-pptx decks
```

## Credits

PDFs are written with [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) and read with
[pdf.js](https://github.com/mozilla/pdf.js) (Apache 2.0). WebP and AVIF use
[jSquash](https://github.com/jamsinclair/jSquash) (Apache 2.0), GIFs
[gifenc](https://github.com/mattdesl/gifenc) (MIT), video
[mp4-muxer](https://github.com/Vanilagy/mp4-muxer) and
[webm-muxer](https://github.com/Vanilagy/webm-muxer) (MIT), zips
[fflate](https://github.com/101arrowz/fflate) (MIT), and the reveal.js export
[reveal.js](https://revealjs.com) (MIT). PDF text is set in
[Noto Sans](https://notofonts.github.io) (SIL Open Font Licence, `src/assets/fonts/OFL.txt`).

MIT licence.

<!-- attributions:start -->
This project is built on other people's work — see [ATTRIBUTIONS.md](ATTRIBUTIONS.md).
<!-- attributions:end -->
