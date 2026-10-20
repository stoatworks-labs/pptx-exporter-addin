#!/bin/bash
# Sideload into Mac PowerPoint: drop a manifest in the wef folder, then quit
# and reopen PowerPoint.
#
#   scripts/sideload-mac.sh            # DEV: points at https://localhost:3000
#   scripts/sideload-mac.sh --hosted   # the production manifest, verbatim
#
# manifest.xml in the repo is the PRODUCTION manifest (see its header). The dev
# copy is derived here by rewriting the hosted origin to the vite dev server,
# so the two can never drift apart in the repo — there is only one file. For
# the dev copy the dev server (npm run dev) must be running and its cert
# trusted (npm run certs).
#
# Both copies carry the same add-in <Id>, so PowerPoint treats them as the same
# add-in: sideload one or the other, never both at once.
#
# NOTE: the custom Home-tab button may only appear after you launch the add-in
# once from  Home ▸ Add-ins ▸ Developer Add-ins ▸ Advanced Exporter  — a known
# wef-sideload quirk. After that first launch the ribbon button sticks.
set -euo pipefail
HERE="$(cd "$(dirname "$0")/.." && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
HOSTED="https://pptx-exporter-addin.stoatworks-labs.com"
DEV="https://localhost:3000"

mode="dev"
for a in "$@"; do
  case "$a" in
    --hosted) mode="hosted" ;;
    *) echo "usage: $0 [--hosted]" >&2; exit 2 ;;
  esac
done

mkdir -p "$WEF"
dst="$WEF/pptx-exporter.manifest.xml"
if [ "$mode" = "hosted" ]; then
  cp "$HERE/manifest.xml" "$dst"
  echo "Sideloaded the PRODUCTION manifest ($HOSTED) → $dst"
else
  grep -q "$HOSTED" "$HERE/manifest.xml" || { echo "manifest.xml does not mention $HOSTED — is it still the production manifest?" >&2; exit 1; }
  sed "s#$HOSTED#$DEV#g" "$HERE/manifest.xml" > "$dst"
  echo "Sideloaded a DEV manifest ($DEV) → $dst"
fi
echo "Quit PowerPoint (Cmd-Q) and reopen it. Then: Home ▸ Add-ins ▸ Developer Add-ins ▸ Advanced Exporter."
