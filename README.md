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

## What's included

```
MacTrace/
├── manifest.xml            # the add-in manifest (validated)
├── shortcuts.json          # custom keyboard shortcuts
├── package.json
├── server.js               # local HTTPS server for the add-in's files
├── assets/                 # PNG icons
├── scripts/
│   ├── make-icons.js       # regenerates the icons (no dependencies)
│   ├── sideload.js         # installs/removes the add-in for Excel on Mac
│   ├── install-autostart.sh / uninstall-autostart.sh   # run the server at login
│   └── reset-sideload.sh   # clears Excel's add-in cache + re-installs
├── src/
│   ├── taskpane.html/.css/.js   # the pane UI + controller
│   ├── formula-parser.js        # Excel formula tokenizer + parser
│   └── excel.js                 # all Office.js calls (batched for speed)
└── preview/                # optional browser mock to see the UI without Excel
```

---

## Install (Excel for Mac)

**Requirements:** Excel for Mac and **Node.js 18+**. A **Microsoft 365
subscription is not required** — Office Add-ins run on perpetual Excel for Mac
(2016 / 2019 / 2021) on a reasonably recent build. Sideloading is never
subscription-gated.

```bash
cd MacTrace
npm install            # installs the dev-cert + manifest tooling
npm run autostart      # runs the HTTPS server now + at every login (no more npm start)
npm run sideload       # copies the manifest into Excel's add-in folder
```

Then **fully quit and reopen Excel**. On the **Home** tab you'll see a
**MacTrace** group with **Trace** and **Dependents** buttons.

> First run creates and trusts a `localhost` certificate (macOS may ask for your
> password). If the pane is blank, visit `https://localhost:3100/src/taskpane.html`
> once in Safari to accept the certificate.

If a sideloaded add-in doesn't appear (Excel's add-in cache can go stale), run:

```bash
bash scripts/reset-sideload.sh   # quits Excel, clears the cache, re-installs, reopens
```

To remove MacTrace: `npm run unsideload` (and `npm run autostart:off`), then
restart Excel.

---

## Use it

Open the pane with **Trace**, or with the shortcuts below. It tracks your
selection automatically — click any cell and its formula breaks down instantly.

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
