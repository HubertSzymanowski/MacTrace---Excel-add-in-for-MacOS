/*
 * Sideloads (or removes) the MacTrace manifest for Excel on Mac by copying it
 * into Excel's add-in folder. Excel must be restarted afterwards.
 *
 *   node scripts/sideload.js            -> install
 *   node scripts/sideload.js --remove   -> uninstall
 */
const fs = require("fs");
const os = require("os");
const path = require("path");

const WEF = path.join(
  os.homedir(),
  "Library/Containers/com.microsoft.Excel/Data/Documents/wef"
);
const SRC = path.join(__dirname, "..", "manifest.xml");
const DEST = path.join(WEF, "MacTrace.manifest.xml");

const remove = process.argv.includes("--remove");

if (remove) {
  if (fs.existsSync(DEST)) {
    fs.unlinkSync(DEST);
    console.log("Removed MacTrace from Excel. Restart Excel to finish.");
  } else {
    console.log("MacTrace was not sideloaded — nothing to remove.");
  }
  process.exit(0);
}

if (!fs.existsSync(SRC)) {
  console.error("manifest.xml not found at " + SRC);
  process.exit(1);
}
fs.mkdirSync(WEF, { recursive: true });
fs.copyFileSync(SRC, DEST);
console.log("Sideloaded MacTrace.");
console.log("  -> " + DEST);
console.log("Now: fully quit and reopen Excel. Look for the MacTrace group on the Home tab.");
