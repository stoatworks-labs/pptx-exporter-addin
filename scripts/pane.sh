#!/bin/bash
# Drive a development pane (npm run dev, sideloaded into PowerPoint or open in
# a browser) from the shell. See src/dev/remote.ts.
#
#   scripts/pane.sh info                       # host, API sets, codecs, the deck
#   scripts/pane.sh export '{"formats":{"png":{"size":"4k"}},"includeHidden":true}'
#   scripts/pane.sh probe '{"widths":[3840,6000],"canvases":[[8000,4500]]}'
#   scripts/pane.sh reload                     # reload the pane (the dev server runs with HMR off)
#   scripts/pane.sh log                        # follow .out/pane.log
#
# Output files land in .out/ (the dev sink). Set ORIGIN for a non-default server.
set -euo pipefail
HERE="$(cd "$(dirname "$0")/.." && pwd)"
ORIGIN="${ORIGIN:-https://localhost:3000}"
case "${1:-}" in
  info) curl -sk -X POST "$ORIGIN/__control" --data '{"op":"info"}' ;;
  reload) curl -sk -X POST "$ORIGIN/__control" --data '{"op":"reload"}' ;;
  probe) curl -sk -X POST "$ORIGIN/__control" --data "$(echo "$2" | python3 -c 'import json,sys; j=json.load(sys.stdin); j["op"]="probe"; print(json.dumps(j))')" ;;
  export) curl -sk -X POST "$ORIGIN/__control" --data "$(echo "$2" | python3 -c 'import json,sys; j=json.load(sys.stdin); j["op"]="export"; print(json.dumps(j))')" ;;
  log) touch "$HERE/.out/pane.log"; tail -f "$HERE/.out/pane.log" ;;
  *) echo "usage: $0 info | export '<json>' | probe '<json>' | reload | log" >&2; exit 2 ;;
esac
