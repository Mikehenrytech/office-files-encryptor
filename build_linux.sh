#!/usr/bin/env bash
set -euo pipefail

APP_NAME="Office_files_Cypher"
ENTRY_POINT="src/main.py"
SPEC_FILE="Office_files_Cypher.spec"

echo
echo "=========================="
echo "Building ${APP_NAME} (Linux)"
echo "=========================="
echo

# Preferimos usar el .spec si existe (mantienes datas i18n ahí)
if [[ -f "${SPEC_FILE}" ]]; then
  echo "[*] Using spec file: ${SPEC_FILE}"
else
  echo "[!] Spec file not found: ${SPEC_FILE}"
  echo "    Falling back to direct command (you may need to add --add-data manually)."
fi

python -c "import sys; print('Python:', sys.executable)"

# Ensure pip + deps
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

# Clean previous builds
rm -rf dist build "${APP_NAME}.spec" || true

# Build
if [[ -f "${SPEC_FILE}" ]]; then
  python -m PyInstaller --noconfirm --clean "${SPEC_FILE}"
else
  # If you don't use spec, ensure i18n is added:
  python -m PyInstaller --noconfirm --clean \
    --name "${APP_NAME}" \
    --onefile \
    --add-data "i18n: i18n" \
    "${ENTRY_POINT}"
fi

echo
echo "=========================="
echo "Build finished"
echo "Output: dist/${APP_NAME}"
echo "=========================="
