#!/usr/bin/env bash
# Register Quizboard with the desktop so it appears in GNOME's dash and
# "Show Applications" grid with the mortarboard icon.
#
# Installs, under ~/.local (no root needed):
#   - PNG icons into  ~/.local/share/icons/hicolor/<size>/apps/quizboard.png
#   - ~/.local/share/applications/quizboard.desktop
#
# Re-run this after make_icon.py to refresh the icon.
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BIN_DIR="${1:-$HOME/.local/bin}"
DATA_DIR="$HOME/.local/share"
ICON_ROOT="$DATA_DIR/icons/hicolor"
APPS_DIR="$DATA_DIR/applications"

mkdir -p "$APPS_DIR" "$ICON_ROOT"
if [ ! -e "$ICON_ROOT/index.theme" ] && [ -f /usr/share/icons/hicolor/index.theme ]; then
  cp /usr/share/icons/hicolor/index.theme "$ICON_ROOT/index.theme"
fi

for size in 512 256 128 64 48; do
  if [ "$size" = 512 ]; then src="$APP_DIR/icons/quizboard.png"; else src="$APP_DIR/icons/quizboard-$size.png"; fi
  [ -f "$src" ] || continue
  dst="$ICON_ROOT/${size}x${size}/apps"
  mkdir -p "$dst"
  cp "$src" "$dst/quizboard.png"
done
mkdir -p "$DATA_DIR/pixmaps"
cp "$APP_DIR/icons/quizboard.png" "$DATA_DIR/pixmaps/quizboard.png"

DESKTOP="$APPS_DIR/quizboard.desktop"
cat > "$DESKTOP" <<EOT
[Desktop Entry]
Type=Application
Version=1.0
Name=Quizboard
GenericName=Quiz Game
Comment=A 6 x 6 board of point values, questions and answers read from an Excel workbook
Keywords=quiz;trivia;game;jeopardy;board;teams;
Exec=$BIN_DIR/quizboard %f
Icon=quizboard
Terminal=false
Categories=Game;
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
StartupNotify=true
StartupWMClass=quizboard
EOT
chmod 644 "$DESKTOP"
echo "Installed: $DESKTOP"

command -v update-desktop-database >/dev/null 2>&1 && { update-desktop-database "$APPS_DIR" || true; }
command -v gtk-update-icon-cache >/dev/null 2>&1 && { gtk-update-icon-cache -f -t "$ICON_ROOT" >/dev/null 2>&1 || true; }
echo "Done. If the icon doesn't appear right away, log out and back in."
