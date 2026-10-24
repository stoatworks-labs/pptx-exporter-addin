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
# add-in: sideload one or the other, never both at once. This takes out any
# other file with that Id first, the installer's stoatworks-advanced-exporter.xml
# included, and nothing else.
#
# NOTE: the custom Home-tab button appears only once the add-in has been opened
# from  Home ▸ Add-ins ▸ Developer Add-ins ▸ Advanced Exporter, and is gone again
# after PowerPoint relaunches (measured with Font Manager, sideloaded the same
# way, on PowerPoint for Mac 16.112.3, 2026-10-24).
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
id="$(sed -n 's:.*<Id>\(.*\)</Id>.*:\1:p' "$HERE/manifest.xml" | head -n 1)"
for f in "$WEF"/*.xml; do
  if [ -f "$f" ] && [ "$f" != "$dst" ] && grep -q "<Id>$id</Id>" "$f"; then
    rm -f "$f"
    echo "Removed another copy: $f"
  fi
done
if [ "$mode" = "hosted" ]; then
  cp "$HERE/manifest.xml" "$dst"
  echo "Sideloaded the PRODUCTION manifest ($HOSTED) → $dst"
else
  grep -q "$HOSTED" "$HERE/manifest.xml" || { echo "manifest.xml does not mention $HOSTED — is it still the production manifest?" >&2; exit 1; }
  sed "s#$HOSTED#$DEV#g" "$HERE/manifest.xml" > "$dst"
  echo "Sideloaded a DEV manifest ($DEV) → $dst"
fi
echo "Quit PowerPoint (Cmd-Q) and reopen it. Then: Home ▸ Add-ins ▸ Developer Add-ins ▸ Advanced Exporter."
