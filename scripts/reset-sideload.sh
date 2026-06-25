#!/bin/bash
#
# Forces Excel for Mac to re-scan sideloaded add-ins.
# Quits Excel, clears ONLY the add-in aggregation cache (not your manifests or
# your files), re-copies the MacTrace manifest, and reopens Excel.
#
# Run it:  bash scripts/reset-sideload.sh
# (Save any open Excel work first — this restarts Excel.)

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
CACHE="$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef"

echo "1/5  Quitting Excel (you may get a save prompt)…"
osascript -e 'tell application "Microsoft Excel" to quit' 2>/dev/null || true
for i in $(seq 1 30); do
  pgrep -f "Microsoft Excel" >/dev/null || break
  sleep 1
done
if pgrep -f "Microsoft Excel" >/dev/null; then
  echo "    Excel is still open — please quit it manually (Cmd+Q) and run this again."
  exit 1
fi

echo "2/5  Clearing the add-in aggregation cache (rebuilds automatically)…"
rm -rf "$CACHE/AggregatedCache" "$CACHE/AppCommands" "$CACHE/Resources" "$CACHE/icons"

echo "3/5  Re-installing the MacTrace manifest…"
mkdir -p "$WEF"
cp "$PROJECT_DIR/manifest.xml" "$WEF/MacTrace.manifest.xml"

echo "4/5  Verifying placement…"
ls -la "$WEF/MacTrace.manifest.xml"

echo "5/5  Reopening Excel…"
open -a "Microsoft Excel"

echo ""
echo "Done. When Excel opens, look on the Home tab for the MacTrace group,"
echo "or Insert tab > My Add-ins > Developer Add-ins."
echo "(Keep your 'npm start' server running in another terminal.)"
