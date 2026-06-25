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
  var BUILD = "hl-v12";

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

  // ---- Selection ------------------------------------------------------------

  function getSelection() {
    return Excel.run(function (ctx) {
      var cell = ctx.workbook.getActiveCell();
      cell.load("address, formulas, text, valueTypes, worksheet/name");
      var sel = ctx.workbook.getSelectedRange();
      sel.load("address, cellCount, rowCount, columnCount");
      return ctx.sync().then(function () {
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
          selection: sel.address,
          selectionLocal: stripSheet(sel.address),
          selectionCount: sel.cellCount
        };
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
  function resolveRefs(refTexts) {
    if (!refTexts || !refTexts.length) return Promise.resolve({});
    var unique = [];
    var seen = {};
    refTexts.forEach(function (t) {
      if (!seen[t]) { seen[t] = true; unique.push(t); }
    });

    return Excel.run(function (ctx) {
      var active = ctx.workbook.worksheets.getActiveWorksheet();
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
        entries.forEach(function (e) {
          if (e.parsed.external) return;
          var ws;
          if (e.parsed.sheet) {
            if (e.sheetProxy.isNullObject) { e.sheetMissing = true; return; }
            ws = e.sheetProxy;
          } else {
            ws = active;
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
        return ctx.sync().then(function () {
          return { ok: true };
        });
      });
    }).catch(function (e) {
      return { ok: false, reason: "error", message: e.message };
    });
  }

  // ---- Dependents -----------------------------------------------------------

  function getDirectDependents() {
    if (!CAPS.dependents()) {
      return Promise.resolve({ supported: false, items: [] });
    }
    return Excel.run(function (ctx) {
      var rng = ctx.workbook.getSelectedRange();
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
  function applyTemporaryHighlight(refText) {
    var p = splitRef(refText);
    if (p.external) return clearTemporaryHighlight().then(function () {
      return { ok: false, reason: "external" };
    });
    var body = stripDollars(p.body);
    return clearTemporaryHighlight()
      .then(function () { return resolveRange(p, body); })
      .then(function (r) {
        if (!r.ok) return r; // missing-sheet / too-large / invalid range
        return readFill(r.sheet, r.address).then(function (original) {
          // --- PAINT: the operation that actually colours the cell -----------
          return Excel.run(function (ctx) {
            ctx.workbook.worksheets.getItem(r.sheet).getRange(r.address)
              .format.fill.color = HL_COLOR;
            return ctx.sync();
          }).then(function () {
            hl = { sheet: r.sheet, address: r.address, color: original };
            persistHighlight(); // highlight is live; record it for recovery
            // --- SCROLL into view (best-effort; paint is already committed) ---
            return Excel.run(function (ctx) {
              var ws = ctx.workbook.worksheets.getItem(r.sheet);
              ws.activate();
              return ctx.sync().then(function () {
                ws.getRange(r.address).select();
                return ctx.sync();
              });
            }).then(
              function () { return { ok: true, scrolled: true }; },
              function () { return { ok: true, scrolled: false }; }
            );
          });
        });
      })
      .catch(function (e) {
        try { console.error("MacTrace applyTemporaryHighlight failed:", e); } catch (_) {}
        return { ok: false, reason: "error", message: e && e.message ? e.message : String(e) };
      });
  }

  // Hand keyboard control back to the worksheet. Re-selecting the active cell
  // forces the Excel host to pull focus out of the task pane and onto the grid —
  // the only reliable way to "exit" the pane via the keyboard on Mac.
  function returnFocusToGrid() {
    return Excel.run(function (ctx) {
      ctx.workbook.getActiveCell().select();
      return ctx.sync();
    }).then(function () {
      return { ok: true };
    }).catch(function (e) {
      return { ok: false, message: e && e.message ? e.message : String(e) };
    });
  }

  // Write a new formula (or literal value) back into a single cell, then sync so
  // Excel recalculates. Uses the same robust sheet resolution as everything else.
  function setFormula(refText, formula) {
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
        ws.getRange(body).formulas = [[formula]];
        return ctx.sync().then(function () { return { ok: true }; });
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
    resolveRefs: resolveRefs,
    navigateTo: navigateTo,
    getDirectDependents: getDirectDependents,
    applyTemporaryHighlight: applyTemporaryHighlight,
    clearTemporaryHighlight: clearTemporaryHighlight,
    returnFocusToGrid: returnFocusToGrid,
    setFormula: setFormula
  };
})(window);
