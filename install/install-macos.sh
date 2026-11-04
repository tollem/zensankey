#!/bin/bash
# Installs (or removes) the ZenSankey add-ins for Excel and PowerPoint on macOS.
#
# Copies ZenSankey.xlam and ZenSankey.ppam to the Office add-ins folder of the
# current user and removes the quarantine attribute added to downloaded files.
# Office on macOS cannot auto-register add-ins from a script: after running it, tick the
# add-in once in Excel (Tools > Excel Add-ins) and in PowerPoint (Tools > PowerPoint Add-ins,
# "+" button); Office remembers the choice.
#
# Usage:  bash install/install-macos.sh            # install
#         bash install/install-macos.sh --uninstall
set -e
NAME="ZenSankey"
HERE="$(cd "$(dirname "$0")" && pwd)"
DIST="$HERE/../dist"
[ -f "$DIST/$NAME.xlam" ] || DIST="$HERE"
TARGET="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Add-Ins.localized"

if [ "$1" = "--uninstall" ]; then
  rm -f "$TARGET/$NAME.xlam" "$TARGET/$NAME.ppam"
  echo "$NAME removed from: $TARGET"
  echo "Untick the add-in in Excel (Tools > Excel Add-ins) and PowerPoint (Tools > PowerPoint Add-ins) if it is still listed."
  exit 0
fi

if [ ! -f "$DIST/$NAME.xlam" ] || [ ! -f "$DIST/$NAME.ppam" ]; then
  echo "Add-in files not found next to this script or in ../dist ($NAME.xlam, $NAME.ppam)." >&2
  exit 1
fi
mkdir -p "$TARGET"
cp "$DIST/$NAME.xlam" "$DIST/$NAME.ppam" "$TARGET/"
xattr -d com.apple.quarantine "$TARGET/$NAME.xlam" 2>/dev/null || true
xattr -d com.apple.quarantine "$TARGET/$NAME.ppam" 2>/dev/null || true
echo "$NAME copied to: $TARGET"
echo
echo "Next steps (once):"
echo "  Excel:      Tools > Excel Add-ins > tick '$NAME' (use Browse if it is not listed)."
echo "  PowerPoint: Tools > PowerPoint Add-ins > '+' > choose $NAME.ppam in the folder above."
echo "  Choose 'Enable macros' if asked. A 'Sankey' tab then appears in the ribbon."
