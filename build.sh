#!/bin/bash
# مذكرتي Pro v17 — Build Script
# No set -e: optional steps (fonts) shouldn't abort the build

echo "==> Python: $(python3 --version)"

echo "==> [1/3] System packages..."
apt-get update -qq 2>/dev/null && \
  apt-get install -y -qq fontconfig fonts-noto-core 2>/dev/null || \
  echo "WARNING: apt-get failed (normal on some platforms)"

echo "==> [2/3] Arabic font..."
FONT_DIR="${HOME}/.fonts/cairo"
mkdir -p "$FONT_DIR" 2>/dev/null || { FONT_DIR="/tmp/fonts/cairo"; mkdir -p "$FONT_DIR"; }

if fc-list 2>/dev/null | grep -qi "cairo"; then
  echo "    ✅ Cairo already present"
else
  FONT_URL="https://github.com/google/fonts/raw/main/ofl/cairo/Cairo%5Bslnt%2Cwght%5D.ttf"
  if curl -fsSL --max-time 30 "$FONT_URL" -o "$FONT_DIR/Cairo.ttf" 2>/dev/null; then
    fc-cache -fv "$FONT_DIR" 2>/dev/null || true
    echo "    ✅ Cairo downloaded"
  else
    curl -fsSL --max-time 30 \
      "https://github.com/google/fonts/raw/main/ofl/amiri/Amiri-Regular.ttf" \
      -o "$FONT_DIR/Amiri-Regular.ttf" 2>/dev/null && \
      echo "    ✅ Amiri fallback installed" || \
      echo "    ⚠️  No Arabic font — Calibri will be used"
    fc-cache -fv "$FONT_DIR" 2>/dev/null || true
  fi
fi

echo "==> [3/3] Python dependencies..."
pip install --no-cache-dir -r requirements.txt

echo ""
echo "════════════════════════════════"
echo "  Build Complete — v17"
echo "  Python : $(python3 --version)"
echo "  Cairo  : $(fc-list 2>/dev/null | grep -ic cairo) file(s)"
echo "  Amiri  : $(fc-list 2>/dev/null | grep -ic amiri) file(s)"
echo "════════════════════════════════"
