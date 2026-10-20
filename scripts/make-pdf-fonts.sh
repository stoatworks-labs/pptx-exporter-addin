#!/bin/bash
# Build the fonts the PDF formats embed (src/assets/fonts/*.ttf) from the
# full Noto Sans files in fonts-src/.
#
# pdf-lib's own subsetter (subset: true) mangles Noto Sans — most glyphs come
# out blank in poppler and pdf.js, though the text still extracts (checked
# 2026-10-20). So the subsetting happens here, once, with fontTools, and
# pdf-lib embeds the result whole. Kept: Latin (incl. Vietnamese), Greek,
# Cyrillic, punctuation, currency, letterlike symbols, arrows, a few shapes.
# Anything else in a deck's notes is drawn by the web view instead (see
# src/pdf/kit.ts).
set -euo pipefail
HERE="$(cd "$(dirname "$0")/.." && pwd)"
[ -x "$HERE/.venv/bin/pyftsubset" ] || { python3 -m venv "$HERE/.venv"; "$HERE/.venv/bin/pip" install -q fonttools; }
RANGES='U+0020-007E,U+00A0-024F,U+0250-02FF,U+0300-036F,U+0370-03FF,U+0400-052F,U+1E00-1EFF,U+2000-206F,U+20A0-20CF,U+2100-214F,U+2190-21FF,U+2200-22FF,U+25A0-25FF,U+FB00-FB06,U+FFFD'
for w in Regular Bold; do
  "$HERE/.venv/bin/pyftsubset" "$HERE/fonts-src/NotoSans-$w.ttf" \
    --unicodes="$RANGES" --no-hinting --layout-features='' --drop-tables+=GSUB,GPOS,GDEF,STAT \
    --name-IDs='*' --notdef-outline --recalc-bounds \
    --output-file="$HERE/src/assets/fonts/NotoSans-$w.ttf"
done
ls -la "$HERE/src/assets/fonts"
