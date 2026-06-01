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
  mkdir -p "$out/scripts/core" "$out/styles" "$out/templates" "$out/lang"

  # Shared core scripts (auth, history, feedback, image, update, sidebar, utils, app, adapter)
  cp -r "${SHARED}/scripts/core/." "$out/scripts/core/"

  # Per-module scripts (entry, adapter, sanitizer)
  cp "${overlay}/scripts/"*.js "$out/scripts/"

  # Stylesheet: concat shared core + every *.css in the overlay's styles/ (sorted)
  cat "${SHARED}/styles/core.css" > "$out/styles/builder.css"
  for css in "${overlay}/styles/"*.css; do
    [ -e "$css" ] || continue
    cat "$css" >> "$out/styles/builder.css"
  done

  # Template: splice home.html and form.html into shell.html
  local tmp
  tmp="$(mktemp)"
  splice "<!-- HOME_PANEL_INSERT -->" "${SHARED}/templates/home.html" "${SHARED}/templates/shell.html" > "$tmp"
  splice "<!-- FORM_INSERT -->"       "${overlay}/templates/form.html" "$tmp" > "$out/templates/builder.html"
  rm -f "$tmp"

  # Copy any other templates (Handlebars .hbs files, partials, etc.) verbatim.
  # form.html was already consumed by the splice above so skip it.
  if [ -d "${overlay}/templates" ]; then
    find "${overlay}/templates" -type f ! -name 'form.html' -print0 |
      while IFS= read -r -d '' src; do
        rel="${src#${overlay}/templates/}"
        mkdir -p "$out/templates/$(dirname "$rel")"
        cp "$src" "$out/templates/$rel"
      done
  fi

  # Language files
  cp "${SHARED}/lang/en.json" "$out/lang/en.json"

  # Manifest (bare — the release workflow overwrites with version/manifest/download)
  cp "${overlay}/module.json" "$out/module.json"

  # Inline assembled templates as pre-registered Handlebars partials so loadTemplates is a no-op.
  node "${ROOT}/shared/scripts/generate-preload.js" "$out" "$name"

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
