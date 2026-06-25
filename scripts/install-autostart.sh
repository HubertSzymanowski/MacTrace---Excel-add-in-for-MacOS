#!/bin/bash
#
# Installs a macOS LaunchAgent so the MacTrace server starts automatically at
# login (and right now), and restarts itself if it ever crashes. After this you
# never need to run `npm start` again.
#
#   bash scripts/install-autostart.sh
#
set -e

PROJECT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
NODE_BIN="$(command -v node)"
LABEL="com.mactrace.server"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

if [ -z "$NODE_BIN" ]; then
  echo "Could not find 'node' on PATH. Install Node.js first." >&2
  exit 1
fi

mkdir -p "$HOME/Library/LaunchAgents"
cat > "$PLIST" <<PLISTEOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>$NODE_BIN</string>
    <string>$PROJECT_DIR/server.js</string>
  </array>
  <key>WorkingDirectory</key><string>$PROJECT_DIR</string>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>$PROJECT_DIR/mactrace-server.log</string>
  <key>StandardErrorPath</key><string>$PROJECT_DIR/mactrace-server.err</string>
</dict>
</plist>
PLISTEOF

# (Re)load it.
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load -w "$PLIST"

echo "Installed: $PLIST"
echo "MacTrace server now starts automatically at login (and is starting now)."
echo "Logs: $PROJECT_DIR/mactrace-server.log"
