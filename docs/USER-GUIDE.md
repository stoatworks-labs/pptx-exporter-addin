# Advanced Exporter user guide

Advanced Exporter is a PowerPoint add-in that turns the open deck into other things: pictures
of every slide at whatever size you need, PDFs that PowerPoint won't make (handouts, notes pages,
a range of slides), video, an HTML5 presentation, LaTeX, Markdown, and cleaned-up copies of the
deck itself. It sits in a pane beside the deck. You tick the formats you want, set each one's
options, press **Export**, and PowerPoint asks where to save the file.

> **Released at v0.1.0.** The hosted build, which is what `manifest.xml` installs, opens and
> exports in PowerPoint for Mac. The full run below used the development build of the same
> code.
>
> **Checked in PowerPoint for Mac 16.112.3 on macOS 26.4.1 (2026-10-20):** every format below
> was exported from a real deck inside PowerPoint, and every file was checked with something
> other than the add-in. Pictures were opened in Pillow, TIFFs decoded by libtiff, video
> decoded by ffmpeg, PDFs read by poppler, and the LaTeX compiled with XeTeX (Tectonic). The
> Office files were read by python-pptx, python-docx and odfpy. PowerPoint itself opened the
> flattened deck, the trimmed PPSX (straight into the slide show), the cleaned copy, a
> one-slide file and the OpenDocument file with no repair prompt. Word opened the notes
> document. The **Export** button raised the native Save dialog.
>
> **Not checked:** PowerPoint for Windows, PowerPoint on the web, and PowerPoint older than
> 16.96 on a Mac (the PDF route those would use was run in PowerPoint for Mac). The HTML5
> presentation was played in a Chromium browser, but not in Safari, not opened from a file on
> disk, and its presenter view was only tried in one window. The pdfLaTeX output hasn't been
> compiled (only XeLaTeX), and the transitions and links in the flattened deck were checked in
> its XML, not by playing it.
>
> **Video bundle and the media notes (2026-10-24), not yet released:** exported inside
> PowerPoint for Mac 16.112.3, with PNG and the HTML5 presentation, from a 460 MB deck of three
> films, a WAV, a trim and a hidden slide reusing a film. The four files came out of PowerPoint's
> own copy of the deck byte for byte the same as the originals, stored rather than deflated;
> `unzip -t` and ffprobe read them all, and `MEDIA.txt` and the notes were right once a
> "linked" clip called `NULL` was fixed (PowerPoint writes that for a trimmed clip). It is heavy
> going: the pane reached about 7 GB reading a deck that size, and an earlier attempt, before the
> reading was slimmed down, wedged it at 2.7 GB. The same export ran in a Chromium browser too.

## Install

The add-in installs from one small file, `manifest.xml`, which points PowerPoint at the
add-in's web address. Nothing is installed on your machine.

**PowerPoint for Mac**

1. Download `manifest.xml` from the install page or the latest GitHub release.
2. Copy it into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (make the
   `wef` folder if it isn't there).
3. Quit PowerPoint completely and open it again.
4. Open a deck, then **Home ▸ Add-ins ▸ Developer Add-ins ▸ Advanced Exporter**. That also
   puts an **Advanced Exporter** button on the Home tab until PowerPoint quits; next time,
   start from **Home ▸ Add-ins** again.

**PowerPoint on the web, and Microsoft 365 on Windows**

1. **Home ▸ Add-ins ▸ More Add-ins**, then **My Add-ins ▸ Upload My Add-in**.
2. Pick `manifest.xml`.

## Using it

1. **Slides.** All of them, the ones selected in PowerPoint's thumbnails, a range (`1-3, 7,
   9-` — open ends run to the first or last slide), or one section. Hidden slides are left out
   unless you tick **Include hidden slides**; slides you select by hand always go.
2. **Formats.** Tick as many as you like. Each one gets a card with its options, and
   **Back to defaults** puts that card back how it started.
3. **Output.** The file name defaults to the deck's. One file comes out as itself; several
   come out as one `.zip`, a folder per format.
4. **Export.** The pane reads the deck afresh (so edits since you opened it are included),
   builds everything, and hands PowerPoint the file to save. A note lists anything worth
   knowing: hidden slides that a format couldn't include, animations that don't survive as
   pictures, and so on. In a zip, the same notes are in `export-notes.txt`.

Under **Presets and rendering**:

- **Presets** save the formats you've ticked and all their options under a name. They live in
  the add-in's storage on this machine.
- **Slide pictures from** is normally **Automatic**: PowerPoint draws each slide at the size a
  format asks for. **PowerPoint's PDF, rasterised** draws them from PowerPoint's PDF instead,
  which is what happens anyway on a PowerPoint too old to draw slides for add-ins.

### Sizes

Every picture format has the same size choices: the slide's own size (96 dpi), HD to 8K, a
width, a height, both, a print resolution in dpi, or a percentage. When the size you pick isn't
the slide's shape (a 4:3 deck into 1920 × 1080), **Fit** decides what happens: keep the slide's
shape, letterbox to the exact size with bars in a colour you choose, crop to fill, or stretch.

PowerPoint for Mac won't draw a slide wider than about 4,000 pixels (4032 px for a 16:9 slide
on 16.112). Bigger pictures come from PowerPoint's own vector PDF, drawn at full resolution, so
an 8K export is as sharp as a 4K one. A hidden slide isn't in that PDF, so at those sizes it's
PowerPoint's largest picture scaled up, and the export notes say so.

### Things that don't come along

- **Animations.** Every picture-based format shows a slide as it looks with all its builds
  finished. Video and GIF are slides with transitions between them, not a recording of the
  slide show.
- **Video and sound playing**, except in the HTML5 presentation. See below.
- **Hidden slides in PowerPoint's PDF.** PowerPoint leaves them out of its PDF, so the vector
  PDF formats can't include them. **PDF — pictures** can.

### Video and sound

Every export says in its notes what it did with the video and sound on the slides it exported.
In short:

| Export | Video and sound |
| --- | --- |
| **Video bundle** | The files themselves, exactly as embedded, in a `.zip` with `MEDIA.txt` |
| **Web presentation (HTML5)** | Embedded video and sound play in place when clicked. PowerPoint's trims, looping, volume and "start automatically" don't come along, and nor does video on a layout; the notes name each one. |
| **PowerPoint file**, **one file per slide** | Kept as they are, trims and settings included |
| Everything else | PowerPoint's picture of each, its poster frame. Nothing plays, and MP4, WebM and GIF are silent. |

Online video (YouTube and friends) and video linked from somebody's disk were never in the deck
file, so no export can bring them along. The notes list them anyway, and so does `MEDIA.txt`, so
you find out before the show rather than during it.

**Video bundle** saves the deck's own video and sound files, named after their slide and shape
(`Slide 2 - Intro film.mp4`), from the exported slides or the whole deck. `MEDIA.txt` beside
them says, for each file, which slides play it, how big it is, what's in it, whether PowerPoint
for Windows and for Mac can play it, and how PowerPoint plays it: starting by itself or on a
click, looping, volume, trims. A clip trimmed in PowerPoint is saved whole, because the trim is
a PowerPoint setting rather than part of the file, so `MEDIA.txt` gives the trim points for
whatever plays it next. Video on a slide's layout or master comes along with every slide that
shows it.

**A deck full of video is slow to read.** Office.js hands the add-in the deck at about a
megabyte a second: a 460 MB deck took six and a half minutes in PowerPoint for Mac 16.112.3,
about three and a half seconds for each four-megabyte slice. That happens at the start of every export,
whatever the format, so it isn't the bundle being slow; it's just where you notice.

## Formats and their options

<!-- options:start (generated by test/docs.test.ts — edit the format definitions, not this) -->

### Images

#### PNG (`.png`)

Lossless. What PowerPoint drew, untouched when the size is the slide’s own shape.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### JPEG (`.jpg`)

Small files for email and the web. Quality is yours to pick.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Quality | 1–100, default **90** |  |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### WebP (`.webp`)

Smaller than JPEG at the same quality, or lossless and smaller than PNG.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Quality | 1–100, default **85** |  |
| Lossless | off |  |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### AVIF (`.avif`)

The smallest of the lot. Slow to encode — this runs in the pane, not on a GPU.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Quality | 1–100, default **60** |  |
| Encoder speed | 0–10, default **7** | 0 is smallest and slowest; 10 is quick and a bit bigger. |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### TIFF (`.tif`)

Lossless, with real DPI — for print shops and DTP. One file per slide, or one multi-page file.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Compression | **LZW (reads everywhere)** · Deflate / ZIP (smaller) · PackBits · None |  |
| One multi-page file | off |  |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### BMP (`.bmp`)

24-bit Windows bitmap, uncompressed — for the kit that still insists.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### TGA (`.tga`)

Targa, 24-bit — the still format older media servers and vision mixers load.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Run-length compressed | on | Smaller. Turn it off for kit that only reads raw Targa. |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### GIF (still) (`.gif`)

256 colours at most. Fine for diagrams, unkind to photos.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | **Keep the slide's shape** · Exact size, letterboxed · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Colours | 2–256, default **256** |  |
| Dither | on | Smooths gradients into dots instead of bands. |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### SVG (`.svg`)

The slide as a picture inside an SVG, with its links still clickable and its text still findable.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Picture inside | **PNG (sharp)** · JPEG (small) |  |
| Quality (shown when it applies) | 1–100, default **88** |  |
| Clickable links | on |  |
| Invisible text layer | on | Lets search, screen readers and copy-paste find the words. |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### Contact sheet (`.png / .jpg`)

Every slide as a thumbnail on one big picture — for running orders and sign-off.

| Option | Choices | Notes |
| --- | --- | --- |
| Columns | 1–12, default **4** |  |
| Thumbnail width | 80–1920 px, default **480** |  |
| Gap | 0–200 px, default **24** |  |
| Labels | None · Slide number · **Number and title** |  |
| Background | default `#1f2328` |  |
| Label colour | default `#f0f0f0` |  |
| Mark hidden slides | on |  |
| Save as | **PNG** · JPEG |  |

### Documents

#### PDF — PowerPoint (vector) (`.pdf`)

PowerPoint's own PDF: sharp at any zoom, text selectable. Trimmed to your slide range, or one file per slide.

| Option | Choices | Notes |
| --- | --- | --- |
| One PDF per slide | off |  |
| File names (shown when it applies) | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |
| Bookmarks from slide titles | on |  |
| Title and author from the deck | on |  |

#### PDF — pictures (`.pdf`)

Every slide as a picture at the resolution you choose. Hidden slides included, size under your control, looks identical everywhere.

| Option | Choices | Notes |
| --- | --- | --- |
| Resolution | 36–600 dpi, default **150** |  |
| Pictures as | **JPEG (small)** · PNG (lossless, big) |  |
| JPEG quality (shown when it applies) | 10–100, default **85** |  |
| Keep links clickable | on |  |
| Bookmarks from slide titles | on |  |

#### Handouts (`.pdf`)

Several slides to a page — 1, 2, 3 with lines for notes, 4, 6 or 9 — as PowerPoint prints them, minus the print dialog.

| Option | Choices | Notes |
| --- | --- | --- |
| Slides per page | 1 · 2 · **3, with lines for notes** · 4 · 6 · 9 |  |
| Orientation | **Portrait** · Landscape |  |
| Paper | **A4** · US Letter · US Legal · A3 · Tabloid |  |
| Slides drawn from | **PowerPoint's PDF (vector)** · Pictures | Vector is sharper and smaller. Pictures don't need PowerPoint's PDF and include hidden slides as drawn. |
| Resolution (shown when it applies) | 36–600 dpi, default **200** |  |
| Pictures as (shown when it applies) | **JPEG (small)** · PNG (lossless, big) |  |
| JPEG quality (shown when it applies) | 10–100, default **85** |  |
| Header | default `{deck}` | {deck} and {date} are filled in. |
| Footer | default `{date}` |  |
| Page numbers | on |  |
| Frame round each slide | on |  |
| Slide numbers | on |  |

#### Notes pages (`.pdf`)

One slide a page with its speaker notes underneath, running onto more pages when the notes do.

| Option | Choices | Notes |
| --- | --- | --- |
| Paper | **A4** · US Letter · US Legal · A3 · Tabloid |  |
| Slides drawn from | **PowerPoint's PDF (vector)** · Pictures | Vector is sharper and smaller. Pictures don't need PowerPoint's PDF and include hidden slides as drawn. |
| Resolution (shown when it applies) | 36–600 dpi, default **200** |  |
| Pictures as (shown when it applies) | **JPEG (small)** · PNG (lossless, big) |  |
| JPEG quality (shown when it applies) | 10–100, default **85** |  |
| Header | default `{deck}` | {deck} and {date} are filled in. |
| Footer | default `{date}` |  |
| Page numbers | on |  |
| Notes size | 7–24 pt, default **12** |  |
| Frame round the slide | on |  |
| Leave out slides with no notes | off |  |

#### OpenDocument (pictures) (`.odp`)

An .odp of slide pictures with the speaker notes, for LibreOffice and anywhere that asks for ODF.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Pictures | **PNG (sharp)** · JPEG (smaller) |  |
| Quality (shown when it applies) | 1–100, default **88** |  |
| Speaker notes | on |  |

#### Word document (notes) (`.docx`)

PowerPoint’s “Create Handouts in Word”: slide pictures with the speaker notes, or with lines to write on.

| Option | Choices | Notes |
| --- | --- | --- |
| Layout | Notes next to slides · **Notes below slides** · Blank lines below slides |  |
| Picture width (shown when it applies) | 3–25 cm, default **15** |  |
| Picture width (shown when it applies) | 3–12 cm, default **8** |  |
| Paper | **A4** · US Letter |  |
| Slide text too | off | The words on each slide under its heading, so the handout can be searched and edited. |
| Pictures | **PNG (sharp)** · JPEG (smaller) |  |
| Quality (shown when it applies) | 1–100, default **88** |  |
| Picture resolution | 480–3840 px, default **1280** | 1280 pixels is plenty for print at these sizes. |

### Web

#### Web presentation (HTML5) (`.html`)

Plays in any browser, even straight from disk: keys, clicks and swipes, working links and video, a presenter view with your notes.

| Option | Choices | Notes |
| --- | --- | --- |
| Packaging | **A folder: index.html and its files** · One self-contained .html file | A folder is lighter and starts quicker. One file is easier to email, but every picture rides inside it as text, about a third bigger, and video and audio only go in while they total 40 MB or less. |
| Slide pictures | **WebP (smallest)** · JPEG · PNG (lossless, biggest) | WebP is roughly a third the size of JPEG for slides, and every current browser shows it. |
| Quality (shown when it applies) | 1–100, default **85** |  |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Transitions | **As in the deck** · Fade, every slide · Slide, every slide · None | From the deck, fades and dissolves play as a fade; push, wipe, cover, uncover and split as a slide; anything fancier as a fade. Anyone who asks their system for less motion gets a cut. |
| Transition length (shown when it applies) | 0–5000 ms, default **0** | 0 keeps each slide’s own. |
| Advance | On a click or key only · **Slide timings, as set up in PowerPoint** · Every few seconds | Timings are each slide’s “After” time, and only when Slide Show › Use Timings is on. A playing video holds the slide. |
| Seconds per slide (shown when it applies) | 1–3600 s, default **8** |  |
| At the end | **As set up in PowerPoint** · Loop back to the start · Stop on the last slide |  |
| Control bar | on |  |
| Progress bar | on |  |
| Slide counter | on |  |
| Speaker notes and presenter view | off | Off leaves the notes out of the files altogether, not merely out of sight — worth it when they were for your eyes only and the export is going to anyone else. |
| Clickable links | on |  |
| Video and audio | on | Copies embedded video and audio out of the deck to play over the slide. Linked media was never in the deck file, so it can’t come along. |
| Page title | empty by default |  |
| Background | default `#000000` | Around the slide when the window is a different shape. |
| Controls | **Dark** · Light |  |

#### reveal.js (`.html`)

A reveal.js presentation: each slide’s picture as its background, links still clickable, notes in reveal’s speaker view.

| Option | Choices | Notes |
| --- | --- | --- |
| reveal.js from | **jsDelivr, pinned to 6.0.2 (needs the internet)** · Copied into the export (works offline) |  |
| Slide pictures | **WebP (smallest)** · JPEG · PNG (lossless, biggest) | WebP is roughly a third the size of JPEG for slides, and every current browser shows it. |
| Quality (shown when it applies) | 1–100, default **85** |  |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Transition | **As in the deck** · None · Fade · Slide · Convex · Concave · Zoom |  |
| Arrows | on |  |
| Progress bar | on |  |
| Slide numbers | on |  |
| Slide number in the address | on | So a reload, or a link someone sends, lands on the same slide. |
| Speaker notes (S) | off | Off leaves the notes out of the files altogether, not merely out of sight — worth it when they were for your eyes only and the export is going to anyone else. |
| Clickable links | on |  |
| Page title | empty by default |  |
| Background | default `#000000` |  |

#### Scrolling page (`.html`)

Every slide down one long page with its notes underneath. Reads on a phone; prints a slide a page.

| Option | Choices | Notes |
| --- | --- | --- |
| Packaging | **A folder: index.html and its files** · One self-contained .html file | A folder is lighter and the pictures load as you scroll. One file is easier to send on, but every picture rides inside it as text, about a third bigger. |
| Slide pictures | **WebP (smallest)** · JPEG · PNG (lossless, biggest) | WebP is roughly a third the size of JPEG for slides, and every current browser shows it. |
| Quality (shown when it applies) | 1–100, default **85** |  |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Speaker notes | on | Off leaves the notes out of the files altogether, not merely out of sight — worth it when they were for your eyes only and the export is going to anyone else. |
| Page title | empty by default |  |

### Video & animation

#### MP4 (`.mp4`)

H.264, HEVC or AV1 video, timed by the deck or by you. Silent, and each slide as it rests — animations don’t play.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | Keep the slide's shape · **Exact size, letterboxed** · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Frame rate | 24 fps · 25 fps · **30 fps** · 50 fps · 60 fps | Only crossfades move, so more frames mostly buys smoother fades. |
| Codec | **H.264** · HEVC (H.265) · AV1 | H.264 plays everywhere. HEVC and AV1 are smaller for the same quality, but older players and editors may not open them. |
| Quality | Low · **Medium** · High · Very high · Bitrate… | Medium is about 8 Mbps for H.264 at 1080p30, scaled for the size, frame rate and codec. |
| Bitrate (shown when it applies) | 0.2–200 Mbps, default **8** | Used as given, whatever the size or codec. |
| Seconds per slide | 0.5–600 s, default **5** |  |
| Use the deck's slide timings where set | on | Slides with a rehearsed or "After" timing stay up that long; the rest get the seconds above. |
| Transitions | **Where the deck has them** · Crossfade every change · Straight cuts | Every transition is played as a crossfade — a push or a vortex included. "Where the deck has them" keeps each one's own length. |
| Crossfade (shown when it applies) | 0–10 s, default **0.5** | Added between slides, not taken from their time. |
| Frames | **Every frame** · One per slide | Every frame keeps editors and fussy players happy. One per slide holds each slide as a single long frame — a far smaller file that players are fine with, though editors may complain. |
| Captions | **None** · Speaker notes as .srt · Speaker notes as .vtt | A subtitle file beside the video: each slide's notes for as long as it is up. Slides without notes get no caption. |

#### WebM (`.webm`)

VP9, VP8 or AV1 for the web. Silent, and each slide as it rests — animations don’t play.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Fit (shown when it applies) | Keep the slide's shape · **Exact size, letterboxed** · Exact size, cropped to fill · Exact size, stretched | Only matters when the size isn't the slide's shape — a 4:3 deck into 1920 × 1080, say. |
| Bars (shown when it applies) | default `#000000` |  |
| Frame rate | 24 fps · 25 fps · **30 fps** · 50 fps · 60 fps | Only crossfades move, so more frames mostly buys smoother fades. |
| Codec | **VP9** · VP8 · AV1 | VP9 plays in every current browser. VP8 for elderly ones; AV1 is smallest and slowest to make. |
| Quality | Low · **Medium** · High · Very high · Bitrate… | Medium is about 8 Mbps for H.264 at 1080p30, scaled for the size, frame rate and codec. |
| Bitrate (shown when it applies) | 0.2–200 Mbps, default **8** | Used as given, whatever the size or codec. |
| Seconds per slide | 0.5–600 s, default **5** |  |
| Use the deck's slide timings where set | on | Slides with a rehearsed or "After" timing stay up that long; the rest get the seconds above. |
| Transitions | **Where the deck has them** · Crossfade every change · Straight cuts | Every transition is played as a crossfade — a push or a vortex included. "Where the deck has them" keeps each one's own length. |
| Crossfade (shown when it applies) | 0–10 s, default **0.5** | Added between slides, not taken from their time. |
| Frames | **Every frame** · One per slide | Every frame keeps editors and fussy players happy. One per slide holds each slide as a single long frame — a far smaller file that players are fine with, though editors may complain. |
| Captions | **None** · Speaker notes as .srt · Speaker notes as .vtt | A subtitle file beside the video: each slide's notes for as long as it is up. Slides without notes get no caption. |

#### Animated GIF (`.gif`)

The deck as a looping GIF for chat, email and READMEs. 256 colours at most, and no sound, naturally.

| Option | Choices | Notes |
| --- | --- | --- |
| Width | 16–1920 px, default **960** |  |
| Seconds per slide | 0.5–600 s, default **5** |  |
| Use the deck's slide timings where set | on | Slides with a rehearsed or "After" timing stay up that long; the rest get the seconds above. |
| Colours | 2–256, default **128** |  |
| Dither | on | Smooths gradients into dots instead of bands. |
| Loop | **Forever** · Play once · Three times |  |
| Crossfade frames | 0–12, default **0** | Frames blended between slides, 50 ms each. Every one is a whole picture, so they add up. |

#### Video bundle (`.zip`)

The deck’s own video and sound files, exactly as embedded, in a .zip with MEDIA.txt: which slide each plays on, and how PowerPoint plays it.

| Option | Choices | Notes |
| --- | --- | --- |
| Files | **Those on the exported slides** · Everything in the deck | Everything means hidden slides and slides you didn’t pick as well. Video on a layout or master comes along with any slide that shows it. |
| Sound files too | on | Off saves the video and leaves out the sound. |
| File names | **After slide and shape (Slide 2 - Intro film.mp4)** · As stored in the deck (media1.mp4) |  |

### LaTeX, Markdown & data

#### LaTeX (Beamer) (`.tex + files`)

A Beamer deck that compiles: either the slides exactly as drawn, or rebuilt as editable LaTeX. Or an article with the notes.

| Option | Choices | Notes |
| --- | --- | --- |
| Slides as | **Editable LaTeX (rebuilt)** · Pictures of the slides (exact) |  |
| Pictures from (shown when it applies) | **PowerPoint's PDF (vector)** · PNG files | Hidden slides aren't in PowerPoint's PDF; any being exported come across as PNGs. |
| Document | **Beamer slides** · Article (slides with notes, for print) |  |
| Engine | **XeLaTeX** · LuaLaTeX · pdfLaTeX | XeLaTeX and LuaLaTeX take any script. pdfLaTeX is the oldest and pickiest: characters it has no way to set become ?. |
| Main font (shown when it applies) | empty by default | Empty keeps Latin Modern. Pick one you have installed that covers your text. |
| CJK font (shown when it applies) | empty by default | Only matters if the deck has Chinese, Japanese or Korean text. Empty lets xeCJK use TeX Live’s own (Fandol), which compiles anywhere. |
| Beamer theme (shown when it applies) | **default** · Madrid · Boadilla · CambridgeUS · Singapore · Warsaw · metropolis |  |
| Speaker notes (shown when it applies) | Leave out · **\note{} (hidden unless you ask)** · Notes on a second screen |  |
| Include speaker notes (shown when it applies) | on |  |
| Sections from the deck | on |  |
| Keep links | on |  |
| Charts and diagrams as pictures (shown when it applies) | on |  |
| Picture width | 640–7680 px, default **1920** |  |

#### Markdown (`.md + images`)

Plain Markdown for notes and wikis, Marp or Pandoc/Quarto slides — the words rebuilt, the slide pictures alongside, or both.

| Option | Choices | Notes |
| --- | --- | --- |
| Flavour | **Markdown (GitHub-flavoured) document** · Marp slides · Pandoc / Quarto slides |  |
| Slides as | Text only · Pictures only · **Text and pictures** |  |
| Speaker notes | on |  |
| Pictures as (shown when it applies) | **PNG** · JPEG · WebP |  |
| Picture width (shown when it applies) | 320–7680 px, default **1600** |  |
| Charts and diagrams as pictures (shown when it applies) | on |  |

#### Outline (`.txt / .rtf / .opml`)

The deck's words as an outline: tab-indented text, RTF that PowerPoint reads straight back in, or OPML for outliners.

| Option | Choices | Notes |
| --- | --- | --- |
| Save as | **RTF (PowerPoint outline)** · Text, tab-indented · OPML |  |

#### Speaker notes (`.txt / .md`)

Just the notes, slide by slide — a script to rehearse from, or hand to the autocue.

| Option | Choices | Notes |
| --- | --- | --- |
| Save as | **Plain text** · Markdown |  |
| Slide numbers and titles | on |  |
| Each slide’s own text too | off |  |
| Leave out slides with no notes | off |  |

#### CSV (`.csv`)

One row per slide — number, title, section, hidden, timings, word count, notes — for a running order or a spreadsheet.

| Option | Choices | Notes |
| --- | --- | --- |
| Separator | **Comma** · Semicolon (European Excel) · Tab |  |
| Notes column | on |  |
| Slide text column | off |  |
| Mark as UTF-8 for Excel | on | Adds a byte-order mark so Excel reads accents and non-Latin text correctly. |

#### JSON (`.json`)

Everything the exporter read from the deck — text with formatting, notes, links, geometry, transitions — for scripts.

| Option | Choices | Notes |
| --- | --- | --- |
| Shape positions | on |  |
| Formatting (runs, levels) | on |  |
| Indented | on |  |

### PowerPoint & Office

#### PowerPoint file (`.pptx / .ppsx / .potx`)

The deck itself, as a .pptx, a show that opens straight into presenting, or a template — trimmed to the slides you picked.

| Option | Choices | Notes |
| --- | --- | --- |
| Save as | **Presentation (.pptx)** · Slide show (.ppsx) — opens straight into the show · Template (.potx) |  |
| Only the slides being exported | on | Off keeps every slide in the deck. |
| Leave out hidden slides | off |  |
| Remove speaker notes | off |  |
| Remove comments | off | Old-style and modern comments, and the list of who wrote them. |
| Remove names from the file | off | Clears author, last modified by, manager and company. The title stays. Comments keep their authors unless they go too. |

#### One file per slide (`.pptx`)

Each slide as a .pptx of its own, masters, notes and media included. PowerPoint does the cutting.

| Option | Choices | Notes |
| --- | --- | --- |
| File names | default `{deck}-{nn}` | {deck} the file name · {n} slide number · {nn}/{nnn} zero-padded · {title} slide title · {i} position in the export |

#### Flattened deck (pictures) (`.pptx`)

Every slide as one picture in a fresh deck, notes, transitions and links kept — it plays the same on a machine without your fonts.

| Option | Choices | Notes |
| --- | --- | --- |
| Size | Slide's own size (96 dpi) · HD 1280 × 720 · **Full HD 1920 × 1080** · QHD 2560 × 1440 · 4K UHD 3840 × 2160 · 5K 5120 × 2880 · 8K 7680 × 4320 · Width… · Height… · Width × height… · Print resolution (dpi)… · Scale (% of slide size)… |  |
| Width (shown when it applies) | 16–16384 px, default **1920** |  |
| Height (shown when it applies) | 16–16384 px, default **1080** |  |
| Resolution (shown when it applies) | 36–1200 dpi, default **300** |  |
| Scale (shown when it applies) | 10–800 %, default **200** |  |
| Pictures | **PNG (sharp)** · JPEG (smaller) |  |
| Quality (shown when it applies) | 1–100, default **88** |  |
| Speaker notes | on |  |
| Transitions and timings | on | Effects PowerPoint 2007 never had become a fade. |
| Clickable links | on | A clear shape over each link, pointing where the original did. |

<!-- options:end -->

## The HTML5 presentation

The **Web presentation (HTML5)** export is a folder (or one self-contained `.html` file) that
plays the deck in any browser, offline.

| Key | Does |
| --- | --- |
| → ↓ Space Page Down Enter N, click, swipe left | Next slide |
| ← ↑ Page Up Backspace P, swipe right | Previous slide |
| Home / End | First / last slide |
| A number, then Enter | Go to that slide |
| O or Esc | All slides |
| F | Full screen |
| B or . / W or , | Black / white screen |
| S | Presenter view (current, next, notes, timer) in a second window |
| ? | Help |

`#/7` at the end of the address opens slide 7. Links on the slides are clickable where the
picture can say where they are, and every link is in the page's text for screen readers and
Find. Speaker notes are left out of the files unless you turn **Speaker notes** on, since a web
presentation tends to get sent on.
