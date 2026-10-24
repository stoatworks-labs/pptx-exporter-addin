# Changelog

## Unreleased

Installers: a signed, notarised Mac `.pkg` with an uninstaller `.pkg`, and a per-user Windows
`Setup.exe` with an Installed apps entry. They put `manifest.xml` where PowerPoint looks for
sideloaded add-ins, which is all an Office.js add-in can have outside Microsoft Marketplace.
The Home-tab button is now described as it behaves: it lasts until PowerPoint quits.

## 0.1.0 — 2026-10-20

First version. 31 formats in six groups, each with its own options, exported from the open deck
inside PowerPoint: pictures (PNG, JPEG, WebP, AVIF, TIFF, BMP, TGA, GIF, SVG, contact sheet),
PDFs (PowerPoint's vector PDF trimmed or split, picture PDF, handouts, notes pages), Word and
OpenDocument, an HTML5 presentation, reveal.js, a scrolling page, MP4, WebM, animated GIF, LaTeX
Beamer and article, Markdown/Marp/Pandoc, outline, speaker notes, CSV, JSON, and PowerPoint
copies (cleaned, per slide, flattened). Checked in PowerPoint for Mac 16.112.3.
