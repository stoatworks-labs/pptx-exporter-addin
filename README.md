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
| Video | MP4 (H.264, HEVC, AV1) and WebM (VP9, VP8, AV1) with the deck's timings and captions from the notes; animated GIF; the deck's own video and sound files in a .zip (Video bundle) |
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
- **Video bundle and media notes: not in PowerPoint yet.** Run in a Chromium browser on a
  460 MB deck (files byte for byte, stored, `MEDIA.txt` right, ffprobe reads them). In
  PowerPoint for Mac the first attempt wedged the pane at 2.7 GB while reading the deck for the
  second time; reading it into one buffer and letting the old copy go first should halve that,
  and hasn't been run there yet.
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
- **PowerPoint hands a deck over at about 1 MB a second.** `getFileAsync` took 385.6 s for a
  461 MB deck full of video (111 slices of 4 MB: 4.3 s to the first, then about 3.5 s each).
  Every export starts by reading the deck, so a big one costs minutes before any format runs.
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
the manifest, builds the Windows installer and attaches both to a GitHub release. Then, on the
Mac with the signing identity, `npm run installer:mac` and `gh release upload vX.Y.Z
release/*.pkg` add the Mac installer and uninstaller (see [Installers](#installers)). The CSP in
[`public/_headers`](public/_headers) lets Office.js load from Microsoft's CDN and lets PowerPoint
on the web frame the pane; `npm run serve:dist` serves `dist/` with those headers applied.

## Installers

People expect an installer, the kind Slido has. Slido can have one because its PowerPoint
add-in is the old kind (VBA plus a native library on the Mac, COM on Windows), which PowerPoint
loads from a folder by design. This is an Office.js add-in, and Microsoft supports exactly two
ways of installing one: Microsoft Marketplace, and an organisation's admin pushing it from the
Microsoft 365 admin centre. Everything else is what Microsoft calls sideloading, "not for
production".

So the installers automate the sideload, and that's all they do. They're the same as Font
Manager's (`pptx-font-manager-addin`), with the names changed:

| | Mac | Windows |
| --- | --- | --- |
| File | `Advanced-Exporter-for-PowerPoint.pkg` | `Advanced-Exporter-for-PowerPoint-Setup.exe` |
| What it does | copies `manifest.xml` into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` as `stoatworks-advanced-exporter.xml` | copies it to `%LOCALAPPDATA%\Programs\Stoatworks\Advanced Exporter for PowerPoint\` and sets `HKCU\Software\Microsoft\Office\16.0\Wef\Developer`, value `<add-in Id>` = that path |
| Same as | Microsoft's `office-addin-dev-settings` on a Mac | the same tool on Windows (read from its 4.0.1 source); not the network-share catalog `sideload-windows.ps1` sets up |
| Rights | your account only, no admin password | your account only, no UAC prompt |
| PowerPoint open? | Installer asks to quit it (`must-close`) | asks you to close it, then Retry |
| Uninstall | `Uninstall-Advanced-Exporter-for-PowerPoint.pkg` | Settings ▸ Apps ▸ Installed apps |
| Signed | Developer ID Installer, notarised and stapled | no: SmartScreen will warn |

Both remove any other copy of the add-in (matched on its `<Id>`, so the sideload script's
`pptx-exporter.manifest.xml` goes, and Font Manager's file beside it doesn't) before adding
theirs, and the Mac one refuses politely if PowerPoint isn't installed or has never been opened
(its sandbox container, where `wef` lives, doesn't exist until then). `scripts/sideload-mac.sh`
does the same in reverse, so a dev copy never sits beside an installed one. Because the
manifest only points at the hosted pane, nobody needs to run an installer again unless
`manifest.xml` itself changes.

```bash
npm run installer:mac            # sign + notarise from ~/.config/stoatworks/release-signing.env
npm run installer:mac:unsigned   # quick local build
npm run installer:windows        # needs makensis (brew install makensis)
```

What has been run, 2026-10-24, on macOS 26.4.1:

- **Mac, from the command line** (`installer -pkg … -target CurrentUserHomeDirectory`, no
  password): install swapped the sideload script's copy for `stoatworks-advanced-exporter.xml`,
  byte for byte the same manifest; Font Manager's manifest beside it was left alone; uninstall
  removed it; a second uninstall said there was nothing to remove. Payload-free packages leave
  no receipt. The installation check logged `advanced-exporter: … container=true app=true` to
  `/var/log/install.log`.
- **Gatekeeper:** both packages are notarised and stapled, and copies carrying a browser's
  quarantine flag are accepted (`source=Notarized Developer ID`). The copy that's installed on
  this Mac now came from one of those, through `installer`.
- **Not run yet:** the packages in Installer.app, and PowerPoint listing and opening the add-in
  from the installer's file name. Font Manager's identical packages were run both ways on the
  same Mac (its README has the details), but these haven't been.
- **Windows: not run at all.** The Setup.exe compiles with warnings as errors, and its payload
  unpacks to the identical manifest.

Two build traps, both found building Font Manager's:

- **makensis needs a UTF-8 locale.** With `LANG` unset (or `LC_ALL=C`), Homebrew's 3.12 dies
  with `std::bad_alloc` before reading a line. `build-exe.sh` sets one.
- **The .pkg carries `._manifest.xml` and `._postinstall`.** macOS 26 stamps every new file with
  `com.apple.provenance`, `xattr -c` cannot remove it, and `pkgbuild` archives it as AppleDouble
  whatever `COPYFILE_DISABLE` says. They're inert: the script only opens `manifest.xml` by name.

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
src/deck/media.ts    video and sound, through the font manager's scanner
vendor/pptx-font-manager/  that scanner and MEDIA.txt, via git subtree (npm run core:pull); don't edit
installer/mac/       the .pkg installer and uninstaller: postinstall scripts, distributions, pages
installer/windows/   the NSIS Setup.exe script, its build script and icon
fonts-src/           full Noto Sans; scripts/make-pdf-fonts.sh subsets it into src/assets/fonts
test/                vitest; fixtures/ are python-pptx decks; installer.test.ts runs the Mac scripts
```

## Credits

PDFs are written with [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) and read with
[pdf.js](https://github.com/mozilla/pdf.js) (Apache 2.0). WebP and AVIF use
[jSquash](https://github.com/jamsinclair/jSquash) (Apache 2.0), GIFs
[gifenc](https://github.com/mattdesl/gifenc) (MIT), video
[mp4-muxer](https://github.com/Vanilagy/mp4-muxer) and
[webm-muxer](https://github.com/Vanilagy/webm-muxer) (MIT), zips
[fflate](https://github.com/101arrowz/fflate) (MIT), and the reveal.js export
[reveal.js](https://revealjs.com) (MIT). Video and sound are found and described by
[pptx-font-manager](https://github.com/stoatworks-labs/pptx-font-manager)'s media scanner (MIT), vendored. PDF text is set in
[Noto Sans](https://notofonts.github.io) (SIL Open Font Licence, `src/assets/fonts/OFL.txt`).

MIT licence.

<!-- attributions:start -->
This project is built on other people's work — see [ATTRIBUTIONS.md](ATTRIBUTIONS.md).
<!-- attributions:end -->
