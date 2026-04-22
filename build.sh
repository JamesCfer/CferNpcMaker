#!/usr/bin/env bash
# build.sh <ModuleName> [<ModuleName>...]
#
# Assembles each named module into a top-level <ModuleName>/ folder by
# combining shared/ (core code) with modules/<ModuleName>/ (overlay).
#
# Output structure per module:
#   <ModuleName>/
#     module.json                     ← from modules/<Name>/module.json
#     scripts/
#       core/                         ← from shared/scripts/core/
#       entry.js                      ← from modules/<Name>/scripts/
#       adapter.js                    ← from modules/<Name>/scripts/
#       sanitizer.js                  ← from modules/<Name>/scripts/
#     styles/
#       builder.css                   ← shared/styles/core.css + modules/<Name>/styles/accents.css
#     templates/
#       builder.html                  ← shared/templates/shell.html with home + form spliced in
#
# Used by:
#   - Local dev (symlink the output folder into Foundry's Data/modules)
#   - .github/workflows/release.yml (assembles each module before zip/release)

set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
SHARED="${ROOT}/shared"

splice() {
  # splice <marker> <replacement-file> <in-file> → stdout
  local marker="$1" replacement="$2" infile="$3"
  awk -v m="$marker" -v rep="$replacement" '
    index($0, m) > 0 { while ((getline line < rep) > 0) print line; close(rep); next }
    { print }
  ' "$infile"
}

build_module() {
  local name="$1"
  local overlay="${ROOT}/modules/${name}"
  local out="${ROOT}/${name}"

  if [ ! -d "$overlay" ]; then
    echo "ERROR: overlay not found at $overlay" >&2
    return 1
  fi

  rm -rf "$out"
  mkdir -p "$out/scripts/core" "$out/styles" "$out/templates"

  # Shared core scripts (auth, history, feedback, image, update, sidebar, utils, app, adapter)
  cp -r "${SHARED}/scripts/core/." "$out/scripts/core/"

  # Per-module scripts (entry, adapter, sanitizer)
  cp "${overlay}/scripts/"*.js "$out/scripts/"

  # Stylesheet: concat core + accents
  cat "${SHARED}/styles/core.css" "${overlay}/styles/accents.css" > "$out/styles/builder.css"

  # Template: splice home.html and form.html into shell.html
  local tmp
  tmp="$(mktemp)"
  splice "<!-- HOME_PANEL_INSERT -->" "${SHARED}/templates/home.html" "${SHARED}/templates/shell.html" > "$tmp"
  splice "<!-- FORM_INSERT -->"       "${overlay}/templates/form.html" "$tmp" > "$out/templates/builder.html"
  rm -f "$tmp"

  # Manifest (bare — the release workflow overwrites with version/manifest/download)
  cp "${overlay}/module.json" "$out/module.json"

  echo "✓ built $name"
}

if [ "$#" -eq 0 ]; then
  echo "usage: $0 <ModuleName> [<ModuleName>...]" >&2
  echo "       (valid names: $(ls "${ROOT}/modules" | tr '\n' ' '))" >&2
  exit 1
fi

for name in "$@"; do
  build_module "$name"
done
