#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
MANIFEST_SRC="${ROOT_DIR}/manifest.dev.xml"
MANIFEST_DST_DIR="${HOME}/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
MANIFEST_DST="${MANIFEST_DST_DIR}/manifest.dev.xml"

mkdir -p "${MANIFEST_DST_DIR}"
cp "${MANIFEST_SRC}" "${MANIFEST_DST}"

open -a "Microsoft Excel"

cd "${ROOT_DIR}"
npm run dev-server
