#!/bin/bash
# Build the Windows installer,
# release/Advanced-Exporter-for-PowerPoint-Setup.exe, from manifest.xml and
# installer/windows/advanced-exporter.nsi.
#
#   installer/windows/build-exe.sh
#
# Needs makensis (brew install makensis; apt-get install nsis), so it runs on a
# Mac, on Linux and in CI alike. The installer is not code-signed: Windows
# SmartScreen will say it doesn't recognise it until it has a reputation or a
# certificate.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
HERE="$ROOT/installer/windows"
OUT="$ROOT/release"
HOSTED="https://pptx-exporter-addin.stoatworks-labs.com"

manifest="$ROOT/manifest.xml"
grep -q "$HOSTED" "$manifest" || { echo "manifest.xml does not point at $HOSTED" >&2; exit 1; }
# The header comment mentions localhost by name; only a URL value counts.
if grep -Eq '"https?://(localhost|127\.0\.0\.1)' "$manifest"; then
  echo "manifest.xml has a localhost URL; refusing to package it" >&2
  exit 1
fi
version="$(node -p "require('$ROOT/package.json').version")"
id="$(sed -n 's:.*<Id>\(.*\)</Id>.*:\1:p' "$manifest" | head -n 1)"
[ -n "$id" ] || { echo "no <Id> in manifest.xml" >&2; exit 1; }

mkdir -p "$OUT"
out="$OUT/Advanced-Exporter-for-PowerPoint-Setup.exe"

# makensis dies with std::bad_alloc before reading a line unless the locale is
# UTF-8 — an unset LANG, or LC_ALL=C, crashes Homebrew's 3.12 even on an empty
# script (2026-10-24). macOS has en_US.UTF-8; Linux always has C.UTF-8.
case "${LC_ALL:-${LC_CTYPE:-${LANG:-}}}" in
  *UTF-8* | *utf8*) ;;
  *) if [ "$(uname)" = Darwin ]; then export LC_ALL=en_US.UTF-8; else export LC_ALL=C.UTF-8; fi ;;
esac
# -WX: a warning is an error, so a typo in the script can't quietly ship.
makensis -V2 -WX \
  "-DVERSION=$version" "-DADDIN_ID=$id" "-DMANIFEST=$manifest" "-DOUTFILE=$out" \
  "$HERE/advanced-exporter.nsi"
echo "built $out ($version)"
