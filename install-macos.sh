#!/usr/bin/env bash
# Install Quizboard on macOS (current user, no admin needed).
#
#   bash install-macos.sh
#
# Creates .venv in this folder with PyQt6, writes the `quizboard` command into
# ~/.local/bin, and puts a Quizboard app (with its icon) in ~/Applications, so it
# opens from Launchpad, Spotlight or the Dock.
set -euo pipefail

APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BIN_DIR="${BIN_DIR:-$HOME/.local/bin}"
APPS_DIR="${APPS_DIR:-$HOME/Applications}"
NAME="Quizboard"
MIN="3.10"            # oldest Python Quizboard supports
MAX=""            # newest (blank = no limit)

if [ "$(uname -s)" != "Darwin" ] && [ -z "${INSTALL_TEST:-}" ]; then
  echo "This installer is for macOS. On Linux use ./install.sh; on Windows install.ps1." >&2
  exit 1
fi

ver_ok() {  # $1 = "3.12"
  python3_ver="$1"
  awk -v v="$python3_ver" -v lo="$MIN" -v hi="$MAX" 'BEGIN {
    split(v, a, "."); split(lo, b, "."); n = a[1]*100 + a[2]; m = b[1]*100 + b[2];
    ok = (n >= m); if (hi != "") { split(hi, c, "."); ok = ok && (n <= c[1]*100 + c[2]) }
    exit ok ? 0 : 1 }'
}

PY=""
for cand in python3.13 python3.12 python3.14 python3.11 python3.10 \
            /opt/homebrew/bin/python3 /usr/local/bin/python3 \
            /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
            python3 /usr/bin/python3; do
  exe="$(command -v "$cand" 2>/dev/null || true)"
  [ -n "$exe" ] || continue
  case "$exe" in */.venv/*|*/venv/*) continue ;; esac     # another program's environment
  v="$("$exe" -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || true)"
  if [ -n "$v" ] && ver_ok "$v"; then PY="$exe"; break; fi
done
if [ -z "$PY" ]; then
  echo "$NAME needs Python 3.10 or newer. Install it, then run this again:" >&2
  echo "    brew install python@3.13        (Homebrew: https://brew.sh)" >&2
  echo "  or the installer from https://www.python.org/downloads/macos/" >&2
  exit 1
fi
echo "Python: $PY"

VENV="$APP_DIR/.venv"
if [ -x "$VENV/bin/python" ]; then
  v="$("$VENV/bin/python" -c 'import sys; print("%d.%d" % sys.version_info[:2])')"
  if ! ver_ok "$v"; then echo "Replacing .venv (it was made with Python $v)."; rm -rf "$VENV"; fi
fi
[ -x "$VENV/bin/python" ] || "$PY" -m venv "$VENV"
if [ -z "${SKIP_PIP:-}" ]; then
  "$VENV/bin/python" -m pip install --quiet --upgrade pip
  "$VENV/bin/python" -m pip install --quiet PyQt6
fi

# `quizboard` command
mkdir -p "$BIN_DIR"
cat > "$BIN_DIR/quizboard" <<EOT
#!/usr/bin/env bash
exec "$VENV/bin/python" "$APP_DIR/quizboard.py" "\$@"
EOT
chmod +x "$BIN_DIR/quizboard"
echo "Installed: $BIN_DIR/quizboard"

# Quizboard.app in ~/Applications
APP="$APPS_DIR/$NAME.app"
rm -rf "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
cp "$APP_DIR/icons/quizboard.icns" "$APP/Contents/Resources/quizboard.icns"
cat > "$APP/Contents/MacOS/quizboard" <<EOT
#!/usr/bin/env bash
cd "$APP_DIR"
exec "$VENV/bin/python" "$APP_DIR/quizboard.py" "\$@"
EOT
chmod +x "$APP/Contents/MacOS/quizboard"
cat > "$APP/Contents/Info.plist" <<EOT
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key><string>$NAME</string>
  <key>CFBundleDisplayName</key><string>$NAME</string>
  <key>CFBundleIdentifier</key><string>local.quizboard</string>
  <key>CFBundleExecutable</key><string>quizboard</string>
  <key>CFBundleIconFile</key><string>quizboard.icns</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>1.0</string>
  <key>LSMinimumSystemVersion</key><string>11.0</string>
  <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
EOT
LSREG=/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister
[ -x "$LSREG" ] && "$LSREG" -f "$APP" >/dev/null 2>&1 || true
echo "App: $APP  (in Launchpad and Spotlight)"

case ":$PATH:" in
  *":$BIN_DIR:"*) echo "Done. Run it with:  quizboard      (or open $NAME from Launchpad)" ;;
  *) echo "Done. Open $NAME from Launchpad, or add $BIN_DIR to your PATH to use \`quizboard\`:"
     echo "    echo 'export PATH=\"\$HOME/.local/bin:\$PATH\"' >> ~/.zprofile" ;;
esac
