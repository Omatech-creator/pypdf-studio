#!/usr/bin/env bash
# =============================================================================
#  Build a portable PyPDF Studio app on macOS or Linux with PyInstaller.
#
#  Output:
#    Linux  -> dist/PyPDFStudio/        (run dist/PyPDFStudio/PyPDFStudio)
#    macOS  -> dist/PyPDFStudio.app     (double-click, or open dist/PyPDFStudio.app)
#
#  The app stores settings/logs in a "PyPDFStudio-Data" folder next to the
#  executable, so the bundle is portable and leaves no traces on the host.
#
#  Usage:  ./build_portable.sh
# =============================================================================
set -euo pipefail
cd "$(dirname "$0")"

PY="${PYTHON:-python3}"

echo ">> Using interpreter: $($PY --version)"
echo ">> Installing build dependencies..."
"$PY" -m pip install --upgrade pip pyinstaller >/dev/null
"$PY" -m pip install PySide6 PyMuPDF pypdf Pillow reportlab pytesseract \
    python-docx openpyxl python-pptx >/dev/null

echo ">> Building with PyInstaller (a few minutes)..."
"$PY" -m PyInstaller --noconfirm --clean --onedir --windowed \
  --name PyPDFStudio \
  --collect-all PySide6 --collect-all fitz --collect-all pymupdf \
  --hidden-import pypdf --hidden-import PIL --hidden-import reportlab \
  --hidden-import pytesseract --hidden-import docx --hidden-import pptx \
  --hidden-import openpyxl --exclude-module cv2 --exclude-module tkinter \
  main.py

echo ">> Pruning unused Qt modules..."
UNUSED="WebEngineCore WebEngineWidgets WebEngineQuick WebChannel WebSockets WebView \
Quick Qml Qt3DCore Qt3DRender Charts DataVisualization Graphs Multimedia SpatialAudio \
Designer Help Sql Test Bluetooth Nfc SerialPort Positioning Sensors Location \
RemoteObjects Scxml StateMachine TextToSpeech Pdf"

if [[ "$(uname)" == "Darwin" ]]; then
  for m in $UNUSED; do
    find dist -iname "*${m}*.framework" -prune -exec rm -rf {} + 2>/dev/null || true
  done
else
  base="dist/PyPDFStudio/_internal/PySide6"
  for m in $UNUSED; do
    find "$base" -iname "*Qt6${m}*" -delete 2>/dev/null || true
    find "$base" -iname "*Qt${m}*" -delete 2>/dev/null || true
  done
  find "$base" -iname "QtWebEngineProcess*" -delete 2>/dev/null || true
  rm -rf "$base"/Qt/qml "$base"/Qt/translations "$base"/Qt/resources 2>/dev/null || true
fi

echo
echo "============================================================"
if [[ "$(uname)" == "Darwin" && -d dist/PyPDFStudio.app ]]; then
  echo " DONE.  App bundle:  dist/PyPDFStudio.app"
  echo " Run:    open dist/PyPDFStudio.app"
else
  echo " DONE.  Portable folder:  dist/PyPDFStudio"
  echo " Run:    ./dist/PyPDFStudio/PyPDFStudio"
fi
echo "============================================================"
