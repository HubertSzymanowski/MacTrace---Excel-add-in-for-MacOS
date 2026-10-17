# MacTrace

A fast, keyboard-driven **formula tracing add-in for Excel on Mac**. Most
formula-auditing tools of this kind are Windows-only; MacTrace brings the same
kind of workflow to macOS for FREE. It's a plain Office.js task-pane add-in, so the same
code runs on Excel for Mac, Windows, and the web.

It does two things, well and quickly:

1. **Trace precedents** — breaks the active cell's formula into a logical tree
   (functions, arguments, references), shows each reference's **live value** and
   **source location**, and lets you "teleport" to any source cell — even on
   another sheet.
2. **Trace dependents** — lists every cell that directly depends on the
   selection, so you can change or delete inputs without triggering cascading
   `#REF!` errors. Click any dependent to drill further.

Navigate entirely from the keyboard: `Backspace` steps back one hop as you drill,
and `Esc` closes the pane, which returns keyboard control to the worksheet.
Reopen it any time with **Cmd+Option+A** (or the **Trace** button).

---

## Install (Excel for Mac) — the easy way

**[→ Download MacTrace-Installer.pkg](https://hubertszymanowski.github.io/MacTrace---Excel-add-in-for-MacOS/dist/MacTrace-Installer.pkg)**
— or visit the [MacTrace page](https://hubertszymanowski.github.io/MacTrace---Excel-add-in-for-MacOS/)
for a friendly walkthrough.

1. **Download** the installer and open it from your Downloads folder.
2. **If macOS says it can't verify the developer** (normal for a free community
   project): System Settings → **Privacy & Security** → scroll down → **Open
   Anyway**. On macOS before Sequoia: right-click the file → **Open** → **Open**.
3. **Follow the installer** — no password needed; it closes and reopens Excel
   for you (save your work first). Then look on Excel's **Home** tab for the
   **Trace** and **Dependents** buttons.

That's it. No terminal, no Node.js, nothing running in the background: the
installer puts a single manifest file into Excel's add-in folder, pointing at
the add-in hosted on GitHub Pages. Your spreadsheet data never leaves your Mac
(see [Design notes](#design-notes-why-its-simple-and-fast)); the connection is
only used to load the add-in's own files.

**Requirements:** Excel for Mac **version 16.55 (late 2021) or newer** — a
**Microsoft 365 subscription is not required**, a perpetual licence on a
current build works. (The ribbon buttons need the shared runtime, Excel for
Mac 16.35+; the Cmd+Option+A shortcuts need 16.55+.) Internet access is needed
when the pane opens.

**Uninstall:** in Finder press `Cmd+Shift+G`, paste
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`, delete
`MacTrace.hosted.manifest.xml`, restart Excel.

---

<details>
<summary><h2>Advanced / for developers</h2></summary>

### What's included

```
MacTrace/
├── manifest.xml            # DEV manifest (localhost:3100, validated)
├── shortcuts.json          # custom keyboard shortcuts
├── package.json
├── server.js               # local HTTPS server for the add-in's files
├── index.html              # installer landing page (served by GitHub Pages)
├── assets/                 # PNG icons
├── dist/                   # GENERATED: hosted manifest + end-user installer
├── scripts/
│   ├── make-icons.js       # regenerates the icons (no dependencies)
│   ├── sideload.js         # installs/removes the DEV add-in for Excel on Mac
│   ├── install-autostart.sh / uninstall-autostart.sh   # run the server at login
│   ├── reset-sideload.sh   # clears Excel's add-in cache + re-installs
│   ├── build-dist.js       # generates dist/MacTrace.manifest.xml (hosted URLs)
│   └── build-installer.sh  # builds dist/MacTrace-Installer.pkg
├── src/
│   ├── taskpane.html/.css/.js   # the pane UI + controller
│   ├── formula-parser.js        # Excel formula tokenizer + parser
│   └── excel.js                 # all Office.js calls (batched for speed)
└── preview/                # optional browser mock to see the UI without Excel
```

### Local dev install (localhost server)

**Requirements:** Excel for Mac and **Node.js 18+**. Sideloading is never
subscription-gated.

```bash
cd MacTrace
npm install            # installs the dev-cert + manifest tooling
npm run autostart      # runs the HTTPS server now + at every login (no more npm start)
npm run sideload       # copies the DEV manifest into Excel's add-in folder
```

Then **fully quit and reopen Excel**. On the **Home** tab you'll see a
**MacTrace** group with **Trace** and **Dependents** buttons.

> First run creates and trusts a `localhost` certificate (macOS may ask for your
> password; it expires after ~30 days — rerun `npm run certs`). If the pane is
> blank, visit `https://localhost:3100/src/taskpane.html` once in Safari to
> accept the certificate.

If a sideloaded add-in doesn't appear (Excel's add-in cache can go stale), run:

```bash
bash scripts/reset-sideload.sh   # quits Excel, clears the cache, re-installs, reopens
```

To remove the dev version: `npm run unsideload` (and `npm run autostart:off`),
then restart Excel.

The dev manifest and the hosted one have **different GUIDs and different
filenames in `wef`** (`MacTrace.manifest.xml` vs `MacTrace.hosted.manifest.xml`),
so both can be installed side by side (you'll see two MacTrace groups on the
ribbon), and `npm run sideload` / `unsideload` never touches an end-user
install.

### The end-user distribution (how the easy install works)

- The add-in's static files are served by **GitHub Pages** from this repo's
  root (`https://hubertszymanowski.github.io/MacTrace---Excel-add-in-for-MacOS/`).
- `npm run build:dist` regenerates `dist/`:
  - `dist/MacTrace.manifest.xml` — the manifest with hosted URLs and its own
    GUID (`scripts/build-dist.js`; validate with `npm run validate:dist`);
  - `dist/MacTrace-Installer.pkg` — a tiny payload-free, **per-user** package
    (no admin password). Its postinstall quits Excel, clears the add-in cache,
    copies the manifest into `wef`, and reopens Excel
    (`scripts/build-installer.sh`).
- `dist/` is committed so Pages can serve the installer directly — **rerun
  `npm run build:dist` and commit whenever `manifest.xml`, the version, or the
  hosting URL changes.** Code-only changes (src/, assets/) need no rebuild:
  users pick them up automatically from Pages.
- Test the pkg without the GUI:
  `installer -pkg dist/MacTrace-Installer.pkg -target CurrentUserHomeDirectory`
  (this really installs it for your user; delete the manifest from `wef` to undo).

**One-time repo setup (maintainer):** GitHub → Settings → Pages → Deploy from a
branch → `main` / `(root)`. If you ever rename the repo or account, every
installed user's manifest keeps the old URL — treat the Pages URL as permanent,
update `PAGES_BASE` in `scripts/build-dist.js`, rebuild, and have users
reinstall.

> The installer is not signed/notarized (no paid Apple Developer ID), so users
> get a one-time Gatekeeper stop — the landing page walks them through it. With
> a Developer ID, add `--sign` to `productbuild` + notarize, and that prompt
> disappears.

</details>

---

## Use it

Select a cell and open the pane with **Trace**, or with the shortcuts below.
MacTrace traces **only when you ask**: clicking around the grid or previewing
references never changes what the pane shows. To audit another cell, select it
and press **Cmd+Option+A** again, or drill into it from the tree.

| Action | How |
| --- | --- |
| Trace **precedents** of the active cell | **Cmd+Option+A**, or the **Trace** button |
| Trace **dependents** of the active cell | **Cmd+Option+Shift+A**, or the **Dependents** button |
| Move through the tree (live-preview the cell) | `↑` / `↓` |
| Expand / collapse a node | `→` / `←` |
| Commit / drill into the selected reference | `Enter`, or click the row |
| Step back one hop | `Backspace` |
| Close the pane (returns control to the grid) | `Esc` |

**About `Esc`:** pressing `Esc` *closes* the pane rather than just moving focus.
On Excel for Mac, a task pane can't reliably hand keyboard focus back to the grid
while staying open, so MacTrace hides the pane instead — once it's gone, Excel
returns control to the worksheet and you can keep typing right away. Reopen with
**Cmd+Option+A** (or the **Trace** button) and your selection is re-traced
instantly. (While editing a formula in place, `Esc` only cancels that edit and
leaves the pane open.)

**Live preview + grid highlight:** as you move `↑`/`↓` over a referenced cell,
MacTrace scrolls to it, **paints it `#27E7F5` in the grid**, and highlights the
matching part of the formula at the top — without rebuilding the tree. The
highlight is a strictly temporary inspection aid: before painting, MacTrace saves
the range's original fill (colour *and* pattern, so "No Fill" cells stay No Fill),
and restores it the moment you move to the next part, click away, hide or close
the pane. The saved state is also mirrored to document settings, so a highlight
is undone even if the pane closes unexpectedly (it's restored on next open).
Press `Enter` (or click) to *commit* and make that cell the new thing you're
auditing; click any blue reference in the formula line to jump straight to it.

**Edit a formula in place:** click anywhere on the formula line (other than a blue
reference) to turn it into an editable field. Type your changes and press `Enter`
(or click away) to write the new formula back to the cell; press `Esc` to cancel.
If the cell was changed in the grid since you traced it, MacTrace won't overwrite
it: it shows the cell's current formula instead, so you can redo the edit on it.

> Note: highlighting physically changes the cell fill (the Excel JS API has no
> overlay layer), so hovering will mark the workbook as modified even though the
> original formatting is always restored.

**Shortcuts & focus:** the shortcuts open the pane, switch to the right mode,
retrace the active cell, and try to grab keyboard focus so arrows work
immediately. They rely on the add-in's *shared runtime*; changing the shortcut
keys (`shortcuts.json`) or the manifest requires a full re-sideload + Excel
restart.

---

## Design notes (why it's simple and fast)

- **No framework, no build step.** Plain HTML/CSS/JS loads instantly.
- **One round-trip per action.** Every Office.js call batches its reads into a
  single `ctx.sync()`. References resolve in a few cheap, phased syncs that never
  pull the values of a huge range (a whole-column `A:A` reference reports its
  size, not a million values).
- **Capabilities are feature-detected.** The add-in loads on any Excel that
  supports add-ins; Dependents (ExcelApi 1.13) degrades with a clear message on
  older builds.

### Deliberately out of scope

MacTrace shows structure, live leaf values, and the final result — it never
recomputes Excel's math, which keeps it fast and non-destructive. Step-by-step
evaluation of intermediate sub-formulas, named-range resolution, and workbook
comparison are not included.

---

## Preview the UI without Excel (optional)

```bash
node preview/serve.js     # http://localhost:4123
```

Loads the real pane against a small mock of Excel so you can see the interface
and try the keyboard navigation in a browser.

## Regenerate icons

```bash
npm run icons
```
