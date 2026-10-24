#!/bin/bash
# Build the Mac installer and uninstaller: two flat packages that put
# manifest.xml into, and take it out of, PowerPoint's wef folder for the user
# who runs them.
#
#   installer/mac/build-pkg.sh             # build, sign, notarise and staple
#   installer/mac/build-pkg.sh --unsigned  # build only, for a quick local look
#
# Writes release/Advanced-Exporter-for-PowerPoint.pkg and
# release/Uninstall-Advanced-Exporter-for-PowerPoint.pkg. The names carry no
# version on purpose: gen-downloads links an unversioned asset through
# /releases/latest/download/, which never goes stale.
#
# Signing and notarising use the fleet's settings in
# ~/.config/stoatworks/release-signing.env (or the environment):
#   RL_MAC_INSTALLER_IDENTITY  "Developer ID Installer: NAME (TEAMID)"
#   RL_NOTARY_KEY, RL_NOTARY_KEY_ID, RL_NOTARY_ISSUER  App Store Connect API key, or
#   RL_NOTARY_PROFILE          a notarytool keychain profile
# A signed build with no notary credentials stops rather than ship a package
# Gatekeeper will refuse.
#
# Mac only (pkgbuild, productbuild, notarytool), so it is run by hand at release
# time, not in CI: the signing identity is in this Mac's keychain.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
HERE="$ROOT/installer/mac"
OUT="$ROOT/release"
HOSTED="https://pptx-exporter-addin.stoatworks-labs.com"

sign=1
for a in "$@"; do
  case "$a" in
    --unsigned) sign=0 ;;
    *) echo "usage: $0 [--unsigned]" >&2; exit 2 ;;
  esac
done

manifest="$ROOT/manifest.xml"
grep -q "$HOSTED" "$manifest" || { echo "manifest.xml does not point at $HOSTED" >&2; exit 1; }
# The header comment mentions localhost by name; only a URL value counts.
if grep -Eq '"https?://(localhost|127\.0\.0\.1)' "$manifest"; then
  echo "manifest.xml has a localhost URL; refusing to package it" >&2
  exit 1
fi
version="$(node -p "require('$ROOT/package.json').version")"

if [ "$sign" = 1 ]; then
  env_file="$HOME/.config/stoatworks/release-signing.env"
  # shellcheck disable=SC1090
  [ -f "$env_file" ] && . "$env_file"
  : "${RL_MAC_INSTALLER_IDENTITY:?set RL_MAC_INSTALLER_IDENTITY, or build with --unsigned}"
  if [ -n "${RL_NOTARY_KEY:-}" ] && [ -f "${RL_NOTARY_KEY:-}" ] && [ -n "${RL_NOTARY_KEY_ID:-}" ] && [ -n "${RL_NOTARY_ISSUER:-}" ]; then
    notary=(--key "$RL_NOTARY_KEY" --key-id "$RL_NOTARY_KEY_ID" --issuer "$RL_NOTARY_ISSUER")
  elif [ -n "${RL_NOTARY_PROFILE:-}" ]; then
    notary=(--keychain-profile "$RL_NOTARY_PROFILE")
  else
    echo "no notary credentials (RL_NOTARY_KEY… or RL_NOTARY_PROFILE); build with --unsigned to skip" >&2
    exit 1
  fi
fi

work="$(mktemp -d)"
trap 'rm -rf "$work"' EXIT
mkdir -p "$OUT"

# build <kind> <output name>
build() {
  local kind="$1" name="$2" ident="com.stoatworks.pptx-exporter-addin.$1"
  local scripts="$work/$kind-scripts" dist="$work/$kind-distribution.xml"

  # The scripts folder travels in the package, so the manifest goes in beside
  # the postinstall that copies it (or, for the uninstaller, reads its Id).
  mkdir -p "$scripts"
  cp "$HERE/$kind-postinstall" "$scripts/postinstall"
  chmod 755 "$scripts/postinstall"
  cp "$manifest" "$scripts/manifest.xml"

  # The package's Scripts also get ._manifest.xml and ._postinstall: macOS 26
  # stamps every new file with com.apple.provenance, which xattr -c cannot
  # remove, and pkgbuild archives it as AppleDouble whatever COPYFILE_DISABLE
  # says (tried 2026-10-24). They are inert: the postinstall only ever opens
  # manifest.xml by name.
  pkgbuild --nopayload --scripts "$scripts" --identifier "$ident" --version "$version" \
    "$work/$kind-component.pkg" >/dev/null
  sed "s/@VERSION@/$version/g" "$HERE/$kind-distribution.xml" >"$dist"

  local args=(--distribution "$dist" --resources "$HERE/resources" --package-path "$work")
  [ "$sign" = 1 ] && args+=(--sign "$RL_MAC_INSTALLER_IDENTITY" --timestamp)
  rm -f "$OUT/$name"
  productbuild "${args[@]}" "$OUT/$name" >/dev/null
  echo "built $OUT/$name ($version)"

  [ "$sign" = 1 ] || return 0
  # One retry: the notary service and the keychain profile lookup both fail
  # transiently now and then. A rejection is not transient and is not retried.
  local log="$work/$kind-notary.log" attempt
  for attempt in 1 2; do
    xcrun notarytool submit "$OUT/$name" "${notary[@]}" --wait >"$log" 2>&1 || true
    grep -q 'status: Accepted' "$log" && break
    grep -Eq 'No Keychain password item found|timed out|Could not connect|NSURLErrorDomain' "$log" || break
    sleep 30
  done
  if ! grep -q 'status: Accepted' "$log"; then
    cat "$log" >&2
    id="$(grep -m1 '  id:' "$log" | awk '{print $2}')"
    [ -n "$id" ] && xcrun notarytool log "$id" "${notary[@]}" >&2 || true
    echo "notarisation failed for $name" >&2
    exit 1
  fi
  xcrun stapler staple "$OUT/$name" >/dev/null
  # What a user's Mac will decide when they double-click it.
  spctl --assess --type install -vv "$OUT/$name" 2>&1 | sed 's/^/  /'
}

build install "Advanced-Exporter-for-PowerPoint.pkg"
build uninstall "Uninstall-Advanced-Exporter-for-PowerPoint.pkg"
