#!/bin/bash
# Build test/fixtures/*.pptx (see make-fixtures.py). Needs python3 and ffmpeg.
set -euo pipefail
HERE="$(cd "$(dirname "$0")/.." && pwd)"
if [ ! -x "$HERE/.venv/bin/python" ]; then
  python3 -m venv "$HERE/.venv"
  "$HERE/.venv/bin/pip" install -q python-pptx pillow
fi
"$HERE/.venv/bin/python" "$HERE/scripts/make-fixtures.py"
