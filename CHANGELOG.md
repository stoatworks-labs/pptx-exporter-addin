# Changelog

## Unreleased

**Transparent backgrounds** for PNG, WebP, AVIF, TIFF and TGA: only what's on the slide, with
an alpha channel, for overlays on video, in media servers and in CG kit. PowerPoint draws a
temporary copy of each slide on black and on white at the end of the deck, and the pane works
the alpha out from the difference, so shadows, glows and see-through fills come out right.
Layout and master graphics (logos, bars) are left out unless you ask for them, and a full-slide
picture at the back can be dropped too. PNG, TIFF and TGA can also write a separate fill and key
for vision mixers. TIFF gains RGBA (unassociated alpha) and TGA a 32-bit form with a TGA 2.0
extension area; without transparency both are as before.

It needs PowerPoint to draw the slides (PowerPointApi 1.8), so it isn't there when slides come
from PowerPoint's PDF or on Office LTSC, and pictures wider than PowerPoint will draw (about
4,000 px) are scaled up, with a note. Your slides aren't changed, but the copies go through
PowerPoint's Undo history and the deck counts as edited: Undo straight after an export brings a
copy back. They're tagged, and the pane deletes any it finds whenever it next reads the deck.

New option keys in saved presets: `transparent`, `masterGraphics`, `backPicture`, `alphaAs`.
They default off, so existing presets export as before.

## 0.2.0 — 2026-11-08

Installers: a signed, notarised Mac `.pkg` with an uninstaller `.pkg`, and a per-user Windows
`Setup.exe` with an Installed apps entry. They put `manifest.xml` where PowerPoint looks for
sideloaded add-ins, which is all an Office.js add-in can have outside Microsoft Marketplace.
The Home-tab button is now described as it behaves: it lasts until PowerPoint quits.

**Video bundle**, a new format: the deck's own video and sound files, exactly as embedded and
named after their slide and shape, in a .zip with MEDIA.txt (slides, codec, whether PowerPoint
for Windows and Mac can play it, start, loop, volume, trims). Online and linked clips are listed,
since they aren't in the deck to save.

Every format now says in its notes what it did with the video and sound on the exported slides.
reveal.js, the scrolling page, pictures, PDFs, OpenDocument and Word used to leave them out
without a word. The HTML5 presentation names trims, playback settings and layout video it
doesn't honour, and calls an online film online rather than "linked". Media is found by the font
manager's scanner (vendored), so video on layouts and older `quickTimeFile` and `wavAudioFile`
shapes count too, and a slide whose only timing is a video no longer reports animations.

## 0.1.0 — 2026-10-20

First version. 31 formats in six groups, each with its own options, exported from the open deck
inside PowerPoint: pictures (PNG, JPEG, WebP, AVIF, TIFF, BMP, TGA, GIF, SVG, contact sheet),
PDFs (PowerPoint's vector PDF trimmed or split, picture PDF, handouts, notes pages), Word and
OpenDocument, an HTML5 presentation, reveal.js, a scrolling page, MP4, WebM, animated GIF, LaTeX
Beamer and article, Markdown/Marp/Pandoc, outline, speaker notes, CSV, JSON, and PowerPoint
copies (cleaned, per slide, flattened). Checked in PowerPoint for Mac 16.112.3.
