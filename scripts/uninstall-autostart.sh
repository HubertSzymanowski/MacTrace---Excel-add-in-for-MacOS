#!/bin/bash
#
# Removes the MacTrace autostart LaunchAgent and stops the background server.
#
#   bash scripts/uninstall-autostart.sh
#
PLIST="$HOME/Library/LaunchAgents/com.mactrace.server.plist"
launchctl unload "$PLIST" 2>/dev/null || true
rm -f "$PLIST"
echo "Removed MacTrace autostart. (Run 'npm start' manually if you still want the server.)"
