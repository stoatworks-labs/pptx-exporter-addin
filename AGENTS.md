# AGENTS.md — Advanced Exporter for PowerPoint

Invariants for anyone (human or agent) changing this repo. The README says what it is and how
to run it; this is what not to break.

## Shape of the thing

- An **Office.js task-pane add-in**, static files only, hosted as a Cloudflare Worker with static
  assets at `pptx-exporter-addin.stoatworks-labs.com`. Everything runs in the pane's web view
  (WKWebView on Mac, WebView2 on Windows, an iframe on the web). **Nothing is uploaded** — keep
  it that way; it is a promise on the install page and in the pane.
- `manifest.xml` is the **production** manifest. `test/manifest.test.ts` fails if a URL in it
  points anywhere but the hostname in `wrangler.toml`. Dev sideloading rewrites the origin on the
  way in (`scripts/sideload-mac.sh`); never commit a localhost manifest.
- The hostname is part of the add-in's identity. Moving it breaks every installed copy.

## Where things come from

- **PowerPoint draws the pixels**: `Slide.getImageAsBase64` (PowerPointApi 1.8) at exactly the
  size a format asks for (`src/render/images.ts`). Never resample a PowerPoint render you could
  have asked for at the right size — `frameFor()` works out the render size first.
- **The .pptx supplies the meaning**: text, notes, hidden flags, transitions, timings, links,
  media, sections (`src/deck/parse.ts`). Placeholder positions inherit from layout then master;
  match on `localName`, never on prefix.
- **PowerPoint's PDF** (`getFileAsync(Pdf)`) is vector but leaves hidden slides out. Map pages to
  slides with `src/render/pdfmap.ts` — exact, adjusted for hidden slides, or a refused mismatch.
  **Never assume page n is slide n** (the lesson Presentation Converter exists for).
- Older PowerPoint without 1.8 falls back to rasterising that PDF with pdf.js. Keep that path
  working: it is the only one Office LTSC gets.
- **Video and sound come from the font manager's scanner**, vendored by git subtree at
  `vendor/pptx-font-manager` and imported through the `@fm` alias (`src/deck/media.ts`). It sees
  what parse.ts doesn't: layout and master media, online and linked clips, trims and playback
  settings, codecs. Never edit the vendored tree; change it in stoatworks-labs/pptx-font-manager
  and `npm run core:pull`.

## Formats

- A format is a `FormatDef` (`src/formats/types.ts`): an option schema the pane renders
  generically, and `run(ctx, values)`. **Formats never call Office.js** — only the context — so
  every format runs against `MockHost` in a browser and most run under vitest.
- **Every format declares `media`** (`still`, `kept`, `plays`, `files` or `none`), and the runner
  turns that into one note about the deck's video and sound on the exported slides
  (`noteMedia`). A format must never lose a video without saying so; TypeScript won't let a new
  format leave the field out.
- Option values are saved in presets. Renaming an option key or a format id silently drops
  people's saved settings; `normalise()` copes, but say so in the changelog.
- Multi-file exports are zipped as they are produced (`src/output/bundle.ts`); pictures and video
  are stored, not deflated.
- Everything from the deck that lands in HTML, XML, LaTeX, RTF, Markdown or CSV is escaped by that
  format's own escaper. Generated pages only link `http(s):`, `mailto:` and `tel:` addresses
  (`safeUrl`).

## Traps, found the hard way

- **pdf-lib's font subsetting blanks most of Noto Sans** (text still extracts; poppler and pdf.js
  draw nothing). The PDF fonts are pre-subset with fontTools (`scripts/make-pdf-fonts.sh`, sources
  in `fonts-src/`) and embedded whole. Don't switch `subset: true` back on without rendering the
  result with `pdftoppm` and looking at it.
- Text the bundled font can't set (CJK, emoji, ✓) is drawn by the web view into a picture
  (`src/pdf/raster-text.ts`), so it reads right but can't be selected.
- **xeCJK eats the spaces round dashes and curly quotes** in Western text. `latex.ts` declares
  those characters Default class; keep it.
- **Safari/WKWebView answers `canvas.toBlob('image/webp')` with a PNG.** Encode through
  `encodeNative` (which checks the type it got) or jSquash, never trust the blob.
- Vite's dev server full-reloads the pane when source changes, which kills a running export.
  Run it with `NO_HMR=1` while testing, and `scripts/pane.sh reload` after an edit.
- **PowerPoint won't draw a slide past ~4,000 px** (4032 ok, 4040 InvalidArgument for 16:9 on
  Mac 16.112.3). `SlideImages` bisects for the ceiling on first refusal and draws bigger
  pictures from the vector PDF; don't replace that with a hard-coded number.
- **The Office dev certificate lasts 30 days.** When it lapses PowerPoint shows "isn't signed by
  a valid security certificate" in the pane. Renewing it (`npm run certs`) needs Allan's
  password — it trusts a CA — so ask, don't work round it.
- Quitting PowerPoint to load a sideloaded manifest closes Allan's decks; ask first. A deck
  opened from another app lives in PowerPoint's own tmp folder and is gone after a quit.

## Verify on a real machine

Unit tests cover parsing, encoders, layouts and package structure. They prove nothing about
PowerPoint's renderer, its PDF, WKWebView's codecs or whether PowerPoint opens a generated .pptx
without a repair prompt. What has and hasn't been run in real PowerPoint is recorded in
`docs/USER-GUIDE.md`'s status block — update it when something is actually run, never because
it compiles.

The pane can be driven from the shell in a dev build (`scripts/pane.sh`, `src/dev/remote.ts`):
queue an export with any formats and options, and the output lands in `.out/` without a Save
dialog. Use it rather than clicking through the pane.

Independent checks used so far: libtiff (`tiffinfo -D`) and Pillow for TIFF/BMP/TGA, poppler
(`pdfinfo`, `pdftoppm`, `pdffonts`) for PDFs, Tectonic (XeTeX) for LaTeX, ffprobe for video,
python-pptx / python-docx / odfpy for Office files (`scripts/check-office.py`).
