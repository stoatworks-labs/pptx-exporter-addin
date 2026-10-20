# Attributions

Advanced Exporter for PowerPoint is built on other people's work. This file lists what that work is, who did
it, and what it is doing here.

It is generated — the master lists live in the `stoatworks-backend` repo and are
pushed out by `scripts/sync-attributions.py`. Edit it there, not here.

## Third-party code this project uses

Libraries, SDKs and frameworks the project is built on or bundles.

### pdf-lib

<https://github.com/Hopding/pdf-lib>  
Licence: MIT  
Copyright: Andrew Dillon

An npm dependency, bundled into the page. Where fonts are embedded, so is @pdf-lib/fontkit, the same author's MIT fork of fontkit.

Writes and edits PDF files in plain JavaScript, with no native library and no PDF engine behind it.

### pdf.js — Mozilla

<https://github.com/mozilla/pdf.js>  
Licence: Apache-2.0  
Copyright: Mozilla Foundation

An npm dependency (pdfjs-dist), bundled into the page with its worker.

Mozilla's PDF renderer, the one in Firefox. It draws a PDF's pages to a canvas at whatever size they are wanted, rather than scaling a picture of them.

### fflate

<https://github.com/101arrowz/fflate>  
Licence: MIT  
Copyright: Arjun Barrett

An npm dependency, bundled into the page.

Zip and deflate in plain JavaScript, fast enough to use in a browser. Office documents, OpenDocument files and most multi-file downloads are zip packages underneath.

### gifenc

<https://github.com/mattdesl/gifenc>  
Licence: MIT  
Copyright: Matt DesLauriers

An npm dependency, bundled into the page.

A small, fast GIF encoder with its own palette quantiser, so animated GIFs are written in the page.

### mp4-muxer

<https://github.com/Vanilagy/mp4-muxer>  
Licence: MIT  
Copyright: Vanilagy

An npm dependency, bundled into the page.

Wraps what the browser's own WebCodecs encoder produces in an MP4 file, so video is made in the page without shipping ffmpeg.

### webm-muxer

<https://github.com/Vanilagy/webm-muxer>  
Licence: MIT  
Copyright: Vanilagy

An npm dependency, bundled into the page.

The same job as mp4-muxer for WebM: the browser's own VP8, VP9 and AV1 encoders, wrapped in a file.

### reveal.js

<https://revealjs.com>  
Licence: MIT  
Copyright: Hakim El Hattab and reveal.js contributors

An npm dependency. Its dist/ files are copied into an export that has to work offline; otherwise the export links them from jsDelivr, pinned to the same version with subresource integrity.

The HTML presentation framework. A reveal.js export is a deck people can go on editing as HTML.

### jSquash

<https://github.com/jamsinclair/jSquash>  
Licence: Apache-2.0  
Copyright: jamsinclair

npm dependencies (@jsquash/webp, @jsquash/avif), bundled with their WebAssembly codecs, which come from the Squoosh app. The WebP codec is libwebp, whose BSD-3-Clause licence (Google Inc.) ships beside it as codec/LICENSE.codec.md; the AVIF codec is libavif v1.0.1, from AOMediaCodec.

WebP and AVIF encoders compiled to WebAssembly. Safari's canvas encodes neither (asked for WebP it quietly returns a PNG), so a tool that has to write them in every webview brings its own.

### Noto Sans

<https://notofonts.github.io>  
Licence: SIL Open Font License 1.1  
Copyright: The Noto Project Authors

Regular and Bold, subset with fontTools and committed with OFL.txt beside them; embedded in the files it is used for.

A typeface a generated PDF can carry inside itself, covering Latin, Greek and Cyrillic. The OFL allows embedding it in any document.

### The npm ecosystem

<https://www.npmjs.com>  
Licence: predominantly MIT  
Copyright: the individual package authors

npm dependencies, resolved and pinned in the lockfile.

Build tooling, test runners and the libraries the front ends are assembled from. The exact set and versions for any build are in that repo's lockfile, which is the authoritative list.

The full transitive dependency set for any build is pinned in this repo's lockfile,
which is the authoritative list. What is named above is the layers a reader would
want to know about, not every package that has ever been resolved.

## Getting this wrong

If your work is here and the description is inaccurate, the licence is wrong, or you would rather not be listed — open an issue and it will be fixed.
