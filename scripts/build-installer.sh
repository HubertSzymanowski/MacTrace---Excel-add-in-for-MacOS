#!/bin/bash
#
# Builds the end-user installer: dist/MacTrace-Installer.pkg
#
#   bash scripts/build-installer.sh      (or: npm run build:dist)
#
# The .pkg is a tiny, payload-free, PER-USER package (no admin password):
# its postinstall script quits Excel, clears the add-in cache, copies the
# HOSTED manifest (dist/MacTrace.manifest.xml) into Excel's wef folder and
# reopens Excel. Everything the manual install does, in one double-click.
#
# The built .pkg is committed so GitHub Pages can serve it from the landing
# page (index.html) without a separate release pipeline.
#
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
MANIFEST="$ROOT/dist/MacTrace.manifest.xml"
OUT="$ROOT/dist/MacTrace-Installer.pkg"
VERSION="$(node -p "require('$ROOT/package.json').version")"
IDENTIFIER="com.mactrace.installer"

# The hosted manifest is an input — (re)generate it first.
node "$ROOT/scripts/build-dist.js"

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
mkdir -p "$WORK/scripts" "$WORK/resources"

# ---- postinstall: the entire install, run as the logged-in user ------------
# The manifest is embedded verbatim so the package is self-contained.
{
  cat <<'HEADER'
#!/bin/bash
# MacTrace post-install. Runs as the current user (per-user install).
# MACTRACE_SKIP_EXCEL=1 skips quitting/reopening Excel (used by tests).
set -u

WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
CACHE="$HOME/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef"
SKIP_EXCEL="${MACTRACE_SKIP_EXCEL:-0}"

# 1) Ask Excel to quit so it re-scans add-ins on next launch. If the user is
#    stuck on a save dialog we give up after 30s and finish anyway — the
#    conclusion screen tells them to restart Excel themselves.
if [ "$SKIP_EXCEL" != "1" ]; then
  osascript -e 'tell application "Microsoft Excel" to quit' >/dev/null 2>&1 || true
  for _ in $(seq 1 30); do
    pgrep -f "Microsoft Excel.app" >/dev/null || break
    sleep 1
  done
fi

# 2) Clear the add-in aggregation cache (Excel rebuilds it; a stale cache is
#    the most common reason a sideloaded add-in doesn't show up).
rm -rf "$CACHE/AggregatedCache" "$CACHE/AppCommands" "$CACHE/Resources" "$CACHE/icons" 2>/dev/null || true

# 3) Register the add-in: one manifest file in Excel's wef folder. The
#    filename is deliberately DIFFERENT from the dev sideload's
#    MacTrace.manifest.xml (scripts/sideload.js) so the hosted and dev
#    versions can coexist and neither flow clobbers the other.
mkdir -p "$WEF"
cat > "$WEF/MacTrace.hosted.manifest.xml" <<'MANIFEST_EOF'
HEADER
  cat "$MANIFEST"
  cat <<'FOOTER'
MANIFEST_EOF

# 4) Reopen Excel (no-op if Excel isn't installed yet).
if [ "$SKIP_EXCEL" != "1" ]; then
  open -a "Microsoft Excel" >/dev/null 2>&1 || true
fi

exit 0
FOOTER
} > "$WORK/scripts/postinstall"
chmod +x "$WORK/scripts/postinstall"

# ---- Installer wizard text --------------------------------------------------
cat > "$WORK/resources/welcome.html" <<'EOF'
<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  body { font-family: -apple-system, "Helvetica Neue", sans-serif; font-size: 13px; color: #1f1f1f; }
  h2 { color: #217346; }
</style></head><body>
<h2>MacTrace — formula tracing for Excel on Mac</h2>
<p>This installs MacTrace for <b>your user account only</b> — no administrator
password is needed.</p>
<p>The installer will:</p>
<ul>
  <li>register the MacTrace add-in with Excel,</li>
  <li>clear Excel&rsquo;s add-in cache so the buttons appear,</li>
  <li><b>close and reopen Excel</b> &mdash; save your work before you continue!</li>
</ul>
<p>Nothing else is installed and nothing runs in the background. MacTrace reads
your formulas locally, inside Excel — your spreadsheet data never leaves your Mac.</p>
</body></html>
EOF

cat > "$WORK/resources/conclusion.html" <<'EOF'
<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  body { font-family: -apple-system, "Helvetica Neue", sans-serif; font-size: 13px; color: #1f1f1f; }
  h2 { color: #217346; }
  code { background: #e6f2ec; padding: 1px 4px; border-radius: 3px; }
</style></head><body>
<h2>All done!</h2>
<p>Open Excel and look on the <b>Home</b> tab for the <b>MacTrace</b> buttons
(<b>Trace</b> and <b>Dependents</b>).</p>
<p>If Excel was already open, quit it fully (<code>Cmd+Q</code>) and open it again.</p>
<p>You need an internet connection when the MacTrace pane opens.</p>
<p style="color:#6b6b6b">To uninstall: delete <code>MacTrace.hosted.manifest.xml</code> from
<code>~/Library/Containers/com.microsoft.Excel/Data/Documents/wef</code> and restart Excel.</p>
</body></html>
EOF

# ---- Build ------------------------------------------------------------------
pkgbuild \
  --nopayload \
  --identifier "$IDENTIFIER" \
  --version "$VERSION" \
  --scripts "$WORK/scripts" \
  "$WORK/MacTrace-component.pkg" >/dev/null

cat > "$WORK/distribution.xml" <<EOF
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="1">
  <title>MacTrace for Excel</title>
  <welcome file="welcome.html"/>
  <conclusion file="conclusion.html"/>
  <!-- Per-user install only: no admin password, writes only inside \$HOME. -->
  <domains enable_localSystem="false" enable_anywhere="false" enable_currentUserHome="true"/>
  <options customize="never" rootVolumeOnly="false"/>
  <choices-outline>
    <line choice="default">
      <line choice="$IDENTIFIER"/>
    </line>
  </choices-outline>
  <choice id="default"/>
  <choice id="$IDENTIFIER" visible="false" title="MacTrace">
    <pkg-ref id="$IDENTIFIER"/>
  </choice>
  <pkg-ref id="$IDENTIFIER" version="$VERSION" onConclusion="none">MacTrace-component.pkg</pkg-ref>
</installer-gui-script>
EOF

mkdir -p "$ROOT/dist"
productbuild \
  --distribution "$WORK/distribution.xml" \
  --resources "$WORK/resources" \
  --package-path "$WORK" \
  "$OUT" >/dev/null

echo "Wrote ${OUT#$ROOT/} ($(du -h "$OUT" | cut -f1 | tr -d ' '))"
echo "Test locally without the GUI:  installer -pkg \"$OUT\" -target CurrentUserHomeDirectory"
