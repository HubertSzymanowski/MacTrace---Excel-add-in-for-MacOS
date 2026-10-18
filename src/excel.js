/*
 * MacTrace — Excel interaction layer
 * ----------------------------------
 * All Office.js calls live here so the UI layer (taskpane.js) stays declarative.
 * Every public method runs inside a single Excel.run and batches its loads into
 * one ctx.sync() round-trip, which is the main lever for "max efficiency in app".
 *
 * Exposed as a global: window.MacTraceExcel
 */
(function (global) {
  "use strict";

  // Build tag — bumped on every change to this layer so the pane can show which
  // code is actually running (shared-runtime JS only reloads on a full Excel
  // restart, so a stale build is a real possibility to rule out).
  var BUILD = "grid-v18";

  // ---- Capability detection -------------------------------------------------

  function supports(version) {
    try {
      return Office.context.requirements.isSetSupported("ExcelApi", version);
    } catch (e) {
      return false;
    }
  }
  var CAPS = {
    dependents: function () {
      return supports("1.13");
    }
  };

  // ---- Reference helpers ----------------------------------------------------

  // Split a textual reference into { sheet, body, external }.
  function splitRef(ref) {
    ref = String(ref);
    var external = false;
    var bang = ref.lastIndexOf("!");
    if (bang === -1) {
      return { sheet: null, body: ref, external: false, raw: ref };
    }
    var sheetPart = ref.slice(0, bang);
    var body = ref.slice(bang + 1);
    if (sheetPart.charAt(0) === "'" && sheetPart.charAt(sheetPart.length - 1) === "'") {
      sheetPart = sheetPart.slice(1, -1).replace(/''/g, "'");
    }
    var mb = sheetPart.match(/^\[[^\]]*\]/);
    if (mb) {
      external = true;
      sheetPart = sheetPart.slice(mb[0].length);
    }
    return { sheet: sheetPart || null, body: body, external: external, raw: ref };
  }

  function stripSheet(address) {
    var bang = String(address).lastIndexOf("!");
    return bang === -1 ? address : address.slice(bang + 1);
  }

  // Worksheet.getRange does not want absolute "$" markers.
  function stripDollars(body) {
    return String(body).replace(/\$/g, "");
  }

  // Excel sheet names are case-insensitive, and Excel can return the sheet portion
  // of an address in a different Unicode normalization than worksheet.name — notably
  // for non-ASCII names like 'Sprzedaż po opakowaniu' — which makes a plain
  // getItem(name)/getItemOrNullObject(name) lookup MISS (→ false "missing-sheet").
  // Match defensively: exact, then NFC-normalized, then case-insensitive.
  function sameName(a, b) {
    if (a == null || b == null) return false;
    if (a === b) return true;
    try {
      var na = String(a).normalize("NFC");
      var nb = String(b).normalize("NFC");
      if (na === nb) return true;
      return na.toLowerCase() === nb.toLowerCase();
    } catch (e) { return false; }
  }

  // Given worksheets loaded with "items/name", return the ACTUAL worksheet name
  // matching `wanted` (so a later getItem() uses the exact, guaranteed-valid name),
  // or null if there's genuinely no such sheet.
  function matchSheetName(items, wanted) {
    if (!items) return null;
    for (var i = 0; i < items.length; i++) {
      if (sameName(items[i].name, wanted)) return items[i].name;
    }
    return null;
  }

  // The worksheet a parsed reference lives on: its own sheet (matched as above)
  // or, when the reference names none, `fallback`. Null if the sheet is missing.
  function sheetFor(wsCol, fallback, p) {
    if (!p.sheet) return fallback;
    var name = matchSheetName(wsCol.items, p.sheet);
    return name != null ? wsCol.getItem(name) : null;
  }

  // ---- Selection / audited target -------------------------------------------

  // The active cell as the pane last left it: after a deliberate trace read it,
  // after our own select() (preview, drill), or once a grid-click check adopted
  // yours. A preview about to select its cell compares against it, so it never
  // selects over a click you made in the meantime. null = unknown.
  var lastActive = null;

  // The audited target as the pane uses it: `cell` is the cell whose formula is
  // traced, `range` the range Dependents analyses (both already loaded).
  function describeTarget(cell, range) {
    var formula = cell.formulas[0][0];
    var formulaStr = formula == null ? "" : String(formula);
    return {
      address: cell.address, // e.g. "Calcs!J510"
      local: stripSheet(cell.address),
      sheet: cell.worksheet.name,
      formula: formulaStr,
      isFormula: formulaStr.charAt(0) === "=",
      text: cell.text[0][0],
      valueType: cell.valueTypes[0][0],
      selection: range.address,
      selectionLocal: stripSheet(range.address),
      selectionCount: range.cellCount
    };
  }

  // What Excel has selected right now. The pane reads this ONLY for a deliberate
  // trace (ribbon button, shortcut, first open): live preview moves the selection,
  // so otherwise it is never "the cell being audited".
  function getSelection() {
    return Excel.run(function (ctx) {
      var cell = ctx.workbook.getActiveCell();
      cell.load("address, formulas, text, valueTypes, worksheet/name");
      var sel = ctx.workbook.getSelectedRange();
      sel.load("address, cellCount");
      return ctx.sync().then(function () {
        lastActive = cell.address;
        return describeTarget(cell, sel);
      });
    });
  }

  // Same shape as getSelection(), for an EXPLICIT target instead of the live
  // selection: `rangeRef` is the audited range, `cellRef` the cell whose formula
  // is traced (its top-left cell; defaults to the range's). This is how the pane
  // drills, steps back and re-renders without ever reading the selection.
  function getTargetInfo(rangeRef, cellRef) {
    var pr = splitRef(rangeRef);
    var pc = splitRef(cellRef || rangeRef);
    if (pr.external || pc.external) {
      return Promise.reject(new Error("it's in another workbook"));
    }
    return Excel.run(function (ctx) {
      var wsCol = ctx.workbook.worksheets;
      var active = wsCol.getActiveWorksheet();
      wsCol.load("items/name");
      return ctx.sync().then(function () {
        var wsR = sheetFor(wsCol, active, pr);
        var wsC = sheetFor(wsCol, active, pc);
        if (!wsR || !wsC) throw new Error("sheet '" + (wsR ? pc.sheet : pr.sheet) + "' not found");
        var range = wsR.getRange(stripDollars(pr.body));
        var cell = wsC.getRange(stripDollars(pc.body)).getCell(0, 0);
        range.load("address, cellCount");
        cell.load("address, formulas, text, valueTypes, worksheet/name");
        return ctx.sync().then(function () { return describeTarget(cell, range); });
      });
    });
  }

  // ---- Resolve a list of references to live values --------------------------

  // Resolved in three cheap syncs so we never pull the values of a huge range
  // (e.g. a whole-column reference like A:A) and never call getRange on a sheet
  // that turns out not to exist:
  //   1. confirm which referenced sheets exist
  //   2. read each range's dimensions (metadata only)
  //   3. read the displayed value for single cells only
  // References that name no sheet resolve on `baseSheet` (the audited formula's
  // own sheet, exact name), not on the active one, which preview may have switched.
  function resolveRefs(refTexts, baseSheet) {
    if (!refTexts || !refTexts.length) return Promise.resolve({});
    var unique = [];
    var seen = {};
    refTexts.forEach(function (t) {
      if (!seen[t]) { seen[t] = true; unique.push(t); }
    });

    return Excel.run(function (ctx) {
      var active = ctx.workbook.worksheets.getActiveWorksheet();
      var base = baseSheet ? ctx.workbook.worksheets.getItemOrNullObject(baseSheet) : null;
      if (base) base.load("isNullObject");
      var entries = unique.map(function (text) {
        var p = splitRef(text);
        var e = { text: text, parsed: p, sheetProxy: null, sheetMissing: false, range: null, single: false };
        if (!p.external && p.sheet) {
          e.sheetProxy = ctx.workbook.worksheets.getItemOrNullObject(p.sheet);
          e.sheetProxy.load("isNullObject");
        }
        return e;
      });

      return ctx.sync().then(function () {
        var home = base && !base.isNullObject ? base : active;
        entries.forEach(function (e) {
          if (e.parsed.external) return;
          var ws;
          if (e.parsed.sheet) {
            if (e.sheetProxy.isNullObject) { e.sheetMissing = true; return; }
            ws = e.sheetProxy;
          } else {
            ws = home;
          }
          try {
            e.range = ws.getRange(stripDollars(e.parsed.body));
            e.range.load("address, cellCount, rowCount, columnCount");
          } catch (err) {
            e.range = null;
          }
        });

        return ctx.sync().then(function () {
          entries.forEach(function (e) {
            if (e.range && e.range.cellCount === 1) {
              e.range.load("text, valueTypes");
              e.single = true;
            }
          });

          return ctx.sync().then(function () {
            var out = {};
            entries.forEach(function (e) {
              if (e.parsed.external) {
                out[e.text] = { kind: "external", display: "(other workbook)", address: e.text };
              } else if (e.sheetMissing) {
                out[e.text] = { kind: "missing", display: "(sheet not found)", address: e.text };
              } else if (!e.range) {
                out[e.text] = { kind: "invalid", display: "(unresolved)", address: e.text };
              } else if (e.single) {
                out[e.text] = {
                  kind: "cell",
                  display: e.range.text[0][0],
                  valueType: e.range.valueTypes[0][0],
                  address: e.range.address
                };
              } else {
                out[e.text] = {
                  kind: "range",
                  display: e.range.rowCount + "×" + e.range.columnCount + " range",
                  count: e.range.cellCount,
                  address: e.range.address
                };
              }
            });
            return out;
          });
        });
      });
    });
  }

  // ---- Navigation ("teleport") ----------------------------------------------

  function navigateTo(refText) {
    var p = splitRef(refText);
    if (p.external) {
      return Promise.resolve({ ok: false, reason: "external" });
    }
    var body = stripDollars(p.body);
    return Excel.run(function (ctx) {
      var wsCol = ctx.workbook.worksheets;
      var active = wsCol.getActiveWorksheet();
      active.load("name");
      if (p.sheet) wsCol.load("items/name");
      return ctx.sync().then(function () {
        var ws = null;
        if (!p.sheet) {
          ws = active;
        } else {
          var name = matchSheetName(wsCol.items, p.sheet);
          if (name != null) ws = wsCol.getItem(name);
        }
        if (!ws) return { ok: false, reason: "missing-sheet" };
        ws.activate();
        ws.getRange(body).select();
        var ac = ctx.workbook.getActiveCell();
        ac.load("address");
        return ctx.sync().then(function () {
          lastActive = ac.address;
          return { ok: true };
        });
      });
    }).catch(function (e) {
      return { ok: false, reason: "error", message: e.message };
    });
  }

  // ---- Dependents -----------------------------------------------------------

  // Direct dependents of the audited range, given as `sheet` (an exact worksheet
  // name as Excel returned it, so a plain getItem is safe) + its local address.
  // With no arguments, of the current selection.
  function getDirectDependents(sheet, rangeLocal) {
    if (!CAPS.dependents()) {
      return Promise.resolve({ supported: false, items: [] });
    }
    return Excel.run(function (ctx) {
      var rng = sheet
        ? ctx.workbook.worksheets.getItem(sheet).getRange(stripDollars(rangeLocal))
        : ctx.workbook.getSelectedRange();
      var deps = rng.getDirectDependents();
      // Load BOTH: the per-area addresses (the reliable source on Excel for Mac)
      // AND the combined address as a fallback.
      deps.load("address, areas/address");
      return ctx.sync().then(function () {
        var raw = [];
        var areas = deps.areas && deps.areas.items;
        if (areas && areas.length) {
          areas.forEach(function (a) { if (a && a.address) raw.push(a.address); });
        }
        if (!raw.length && deps.address) raw.push(deps.address);

        // Each raw address may itself be a comma-joined multi-area string, so split
        // every one into clean standalone addresses, then flatten + de-duplicate.
        var seen = {};
        var items = [];
        raw.forEach(function (r) {
          splitAddressList(r).forEach(function (addr) {
            if (addr && !seen[addr]) {
              seen[addr] = true;
              items.push({ address: addr, local: stripSheet(addr) });
            }
          });
        });
        return { supported: true, items: items };
      });
    }).catch(function (e) {
      // getDirectDependents REJECTS both when a cell genuinely feeds nothing
      // downstream (ItemNotFound) and on real errors — distinguish them so a true
      // failure doesn't silently read as "no dependents".
      var code = (e && (e.code || e.name)) || "";
      if (code === "ItemNotFound") return { supported: true, items: [] };
      return { supported: true, items: [], error: (e && e.message) ? e.message : String(e) };
    });
  }

  // Split a multi-area address like "'My Sheet'!E12,'My Sheet'!E40" (or
  // "Sheet1!A1,B2") into clean, individually-resolvable single-area addresses.
  // CRITICAL: each piece must be a valid standalone address before it reaches
  // getRange()/select() — the old code passed the whole comma-joined string as if
  // it were one cell, so splitRef() treated "'S'!E12,'S'" as the sheet name.
  //   - split only on commas OUTSIDE single quotes (quoted sheet names may contain
  //     commas/spaces, and '' is an escaped quote within a name)
  //   - carry the sheet prefix forward to any area that omits it
  function splitAddressList(combined) {
    var s = String(combined || "");
    var parts = [];
    var cur = "";
    var inQuote = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === "'") {
        if (inQuote && s.charAt(i + 1) === "'") { cur += "''"; i++; continue; } // escaped ''
        inQuote = !inQuote;
        cur += ch;
      } else if (ch === "," && !inQuote) {
        if (cur.trim()) parts.push(cur.trim());
        cur = "";
      } else {
        cur += ch;
      }
    }
    if (cur.trim()) parts.push(cur.trim());

    var lastSheet = null;
    return parts.map(function (addr) {
      var bang = addr.lastIndexOf("!");
      if (bang === -1) {
        return lastSheet ? lastSheet + "!" + addr : addr; // inherit the sheet prefix
      }
      lastSheet = addr.slice(0, bang);
      return addr;
    });
  }

  // ---- Temporary highlight (inspection aid) ---------------------------------
  //
  // Paints the inspected range #27E7F5, having first SAVED its original fill so
  // the highlight is strictly reversible. Uses the universally-supported
  // range.format.fill API (ExcelApi 1.1) — fill.color to read/set, fill.clear()
  // to restore "No Fill" — rather than get/setCellProperties (1.9), which is
  // unreliable on Excel for Mac. Only one range is highlighted at a time, and the
  // saved state lives in memory AND in document settings, so the highlight can be
  // undone even after the pane closes unexpectedly or Excel is relaunched.

  var HL_COLOR = "#27E7F5";
  var HL_KEY = "MacTrace.tempHighlight"; // settings key for crash/close recovery
  // Upper bound on what we'll paint. Filling/clearing a range is a SINGLE Excel
  // operation regardless of cell count, so whole rows (16,384) and whole columns
  // (1,048,576) are cheap — this just guards against pathological multi-column /
  // whole-sheet references. Sized to comfortably cover one whole column.
  var HL_MAX_CELLS = 1200000;
  var hl = null; // { sheet, address, color }  (color "" / "#FFFFFF" == No Fill)

  // Mirror the in-memory record into document settings (fire-and-forget) so
  // cleanup can happen even if this runtime goes away before clearing.
  function persistHighlight() {
    try {
      var settings = Office.context.document.settings;
      if (hl) settings.set(HL_KEY, hl);
      else settings.remove(HL_KEY);
      settings.saveAsync(function () {});
    } catch (e) {/* settings unavailable — non-fatal */}
  }

  // Restore a range's saved fill. Empty / white originals are treated as No Fill.
  function restoreFill(range, color) {
    if (!color || color === "" || String(color).toUpperCase() === "#FFFFFF") {
      range.format.fill.clear();
    } else {
      range.format.fill.color = color;
    }
  }

  // Restore the currently-saved highlight (from memory, or from settings if this
  // is a fresh runtime recovering a dangling highlight), then forget it.
  function clearTemporaryHighlight() {
    var rec = hl;
    if (!rec) {
      try { rec = Office.context.document.settings.get(HL_KEY); } catch (e) { rec = null; }
    }
    hl = null;
    if (!rec) {
      persistHighlight();
      return Promise.resolve({ ok: true, cleared: false });
    }
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItemOrNullObject(rec.sheet);
      ws.load("isNullObject");
      return ctx.sync().then(function () {
        if (!ws.isNullObject) restoreFill(ws.getRange(rec.address), rec.color);
        return ctx.sync();
      });
    }).then(function () {
      persistHighlight();
      return { ok: true, cleared: true };
    }).catch(function (e) {
      persistHighlight();
      return { ok: false, reason: "error", message: e.message };
    });
  }

  // A selection change in the grid. The pane runs this on its grid queue, after
  // any preview in flight, so it sees the settled selection. Whatever the active
  // cell now is becomes the baseline (see lastActive). The preview fill is kept
  // only while the selection is EXACTLY what our own select() left. That is
  // recorded rather than recomputed because Excel widens it to whole merged
  // areas; any other selection, even one that merely contains the cell, restores
  // the fill, so it can't leak into a copy, your own formatting or an inserted
  // row. A multi-area selection is never ours. If Excel can't be read at all
  // (e.g. you're typing in a cell), nothing changes and the next change re-checks.
  function reconcileSelection() {
    return Excel.run(function (ctx) {
      var active = ctx.workbook.getActiveCell(); // readable even with several areas selected
      active.load("address");
      return ctx.sync().then(function () {
        lastActive = active.address;
        var rec = hl;
        if (!rec) return { rec: null };
        var sel = ctx.workbook.getSelectedRange();
        sel.load("address, worksheet/name");
        return ctx.sync().then(function () {
          var onIt = rec.selected
            ? sameName(sel.address, rec.selected)
            : sameName(sel.worksheet.name, rec.sheet) &&
              stripSheet(sel.address).toUpperCase() === String(rec.address).toUpperCase();
          return { rec: rec, onIt: onIt };
        }, function () {
          return { rec: rec, onIt: false }; // getSelectedRange rejects a multi-area selection
        });
      });
    }).then(function (r) {
      if (!r.rec || r.onIt || hl !== r.rec) return { ok: true, cleared: false };
      return clearTemporaryHighlight();
    }, function () {
      return { ok: true, cleared: false };
    });
  }

  // Resolve a reference to a concrete { sheet name, local address } and guard the
  // size — the reliable part (loading name/address/cellCount rarely fails).
  function resolveRange(p, body) {
    return Excel.run(function (ctx) {
      var wsCol = ctx.workbook.worksheets;
      var active = wsCol.getActiveWorksheet();
      active.load("name");
      if (p.sheet) wsCol.load("items/name");
      return ctx.sync().then(function () {
        var ws = null, name = null;
        if (!p.sheet) {
          ws = active; name = active.name;
        } else {
          name = matchSheetName(wsCol.items, p.sheet);
          if (name != null) ws = wsCol.getItem(name);
        }
        if (!ws) return { ok: false, reason: "missing-sheet", sheet: p.sheet };
        var range = ws.getRange(body);
        range.load("address, cellCount");
        return ctx.sync().then(function () {
          if (range.cellCount > HL_MAX_CELLS) {
            return { ok: false, reason: "too-large", count: range.cellCount };
          }
          return { ok: true, sheet: name, address: stripSheet(range.address) };
        });
      });
    }).catch(function (e) {
      return { ok: false, reason: "error", message: e && e.message ? e.message : String(e) };
    });
  }

  // Read a range's current fill colour. This is the FRAGILE operation — loading
  // fill.color can make ctx.sync() reject on Excel for Mac (ranges / mixed fills),
  // so it lives in its OWN Excel.run and any failure degrades to "" (= No Fill),
  // which simply means we restore by clearing the fill later. It can never block
  // the paint, because the paint is a separate Excel.run.
  function readFill(sheet, address) {
    return Excel.run(function (ctx) {
      var fill = ctx.workbook.worksheets.getItem(sheet).getRange(address).format.fill;
      fill.load("color");
      return ctx.sync().then(function () {
        var c = "";
        try { c = fill.color || ""; } catch (e) { c = ""; }
        return c;
      });
    }).catch(function () { return ""; });
  }

  // Restore any previous highlight, then paint the new range #27E7F5.
  //
  // Each phase runs in its OWN Excel.run so a failure in one can't sabotage the
  // next — this is the fix for "paint throws, nothing colours": the old code read
  // the fill and painted in the SAME batch, so a fill.color load that rejected on
  // Mac took the paint down with it. Now:
  //   1. resolveRange — sheet + address + size guard (reliable)
  //   2. readFill     — original colour, best-effort, isolated (can't block paint)
  //   3. PAINT        — minimal isolated batch: set fill.color and sync
  //   4. SCROLL       — activate + select, best-effort (paint already committed)
  // `stillWanted` (optional) is asked before the paint and before the scroll; once
  // the pane has moved on, a stale preview stops there and returns
  // { superseded: true }, so it can't select its cell after you've moved on.
  function applyTemporaryHighlight(refText, stillWanted) {
    var wanted = function () { return !stillWanted || stillWanted(); };
    var SUPERSEDED = { ok: false, reason: "superseded", superseded: true };
    var p = splitRef(refText);
    if (p.external) return clearTemporaryHighlight().then(function () {
      return wanted() ? { ok: false, reason: "external" } : SUPERSEDED;
    });
    var body = stripDollars(p.body);
    return clearTemporaryHighlight()
      .then(function () { return resolveRange(p, body); })
      .then(function (r) {
        if (!wanted()) return SUPERSEDED;
        if (!r.ok) return r; // missing-sheet / too-large / invalid range
        return readFill(r.sheet, r.address).then(function (original) {
          if (!wanted()) return SUPERSEDED;
          // --- PAINT: the operation that actually colours the cell -----------
          return Excel.run(function (ctx) {
            ctx.workbook.worksheets.getItem(r.sheet).getRange(r.address)
              .format.fill.color = HL_COLOR;
            return ctx.sync();
          }).then(function () {
            hl = { sheet: r.sheet, address: r.address, color: original };
            persistHighlight(); // highlight is live; record it for recovery
            // Painted but no longer wanted: skip the select. Whatever superseded
            // it is queued next and restores this fill first.
            if (!wanted()) return { ok: true, scrolled: false, superseded: true };
            // --- SCROLL into view (best-effort; paint is already committed) ---
            // Not if you've clicked into the grid since the pane last moved the
            // selection (the active cell is no longer where we left it): the
            // preview must not undo your click or pull focus back. The grid-click
            // check queued behind it restores this fill.
            var rec = hl;
            return Excel.run(function (ctx) {
              var before = ctx.workbook.getActiveCell();
              before.load("address");
              return ctx.sync().then(function () {
                if (lastActive != null && !sameName(before.address, lastActive)) return { moved: true };
                var ws = ctx.workbook.worksheets.getItem(r.sheet);
                ws.activate();
                return ctx.sync().then(function () {
                  ws.getRange(r.address).select();
                  var sel = ctx.workbook.getSelectedRange(); // widened to whole merged areas
                  var after = ctx.workbook.getActiveCell();
                  sel.load("address");
                  after.load("address");
                  return ctx.sync().then(function () {
                    lastActive = after.address;
                    rec.selected = sel.address; // exactly what our select left
                    return { ok: true, scrolled: true };
                  });
                });
              });
            }).then(
              function (res) { return res.moved ? { ok: true, scrolled: false, superseded: true } : res; },
              function () { return { ok: true, scrolled: false }; }
            );
          });
        });
      })
      .catch(function (e) {
        try { console.error("MacTrace applyTemporaryHighlight failed:", e); } catch (_) {}
        if (!wanted()) return SUPERSEDED;
        return { ok: false, reason: "error", message: e && e.message ? e.message : String(e) };
      });
  }

  // Write a new formula (or literal value) back into a single cell, then sync so
  // Excel recalculates. Uses the same robust sheet resolution as everything else.
  // With `expected`, it's a compare-and-set: the write happens only if the cell
  // still holds that formula, else { ok: false, reason: "changed" }. The pane
  // shows a snapshot, and the cell may have been edited in the grid (or shifted
  // by an inserted row) since it was traced.
  function setFormula(refText, formula, expected) {
    var p = splitRef(refText);
    if (p.external) return Promise.resolve({ ok: false, reason: "external" });
    var body = stripDollars(p.body);
    return Excel.run(function (ctx) {
      var wsCol = ctx.workbook.worksheets;
      var active = wsCol.getActiveWorksheet();
      active.load("name");
      if (p.sheet) wsCol.load("items/name");
      return ctx.sync().then(function () {
        var ws = null;
        if (!p.sheet) {
          ws = active;
        } else {
          var name = matchSheetName(wsCol.items, p.sheet);
          if (name != null) ws = wsCol.getItem(name);
        }
        if (!ws) return { ok: false, reason: "missing-sheet" };
        var range = ws.getRange(body);
        if (expected == null) {
          range.formulas = [[formula]];
          return ctx.sync().then(function () { return { ok: true }; });
        }
        range.load("formulas");
        return ctx.sync().then(function () {
          var live = range.formulas[0][0];
          if (String(live == null ? "" : live) !== String(expected)) return { ok: false, reason: "changed" };
          range.formulas = [[formula]];
          return ctx.sync().then(function () { return { ok: true }; });
        });
      });
    }).catch(function (e) {
      return { ok: false, reason: "error", message: e && e.message ? e.message : String(e) };
    });
  }

  global.MacTraceExcel = {
    BUILD: BUILD,
    CAPS: CAPS,
    HL_COLOR: HL_COLOR,
    getSelection: getSelection,
    getTargetInfo: getTargetInfo,
    resolveRefs: resolveRefs,
    navigateTo: navigateTo,
    getDirectDependents: getDirectDependents,
    applyTemporaryHighlight: applyTemporaryHighlight,
    clearTemporaryHighlight: clearTemporaryHighlight,
    reconcileSelection: reconcileSelection,
    setFormula: setFormula
  };
})(window);
