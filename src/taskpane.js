/*
 * MacTrace — task pane controller
 * -------------------------------
 * Owns the UI: tracks the Excel selection, renders the precedents tree or the
 * dependents list, drives keyboard navigation, and manages the "rubber band"
 * (snap back to where you started auditing).
 */
(function () {
  "use strict";

  var P = window.MacTraceParser;
  var X = window.MacTraceExcel;

  var state = {
    mode: "precedents", // "precedents" | "dependents"
    selection: null, // last getSelection() result
    origin: null, // { address } — rubber-band start cell
    navStack: [], // visited addresses, for Back
    cursor: 0, // index into the currently focusable rows
    busy: false,
    pending: false,
    programmatic: false, // true while WE drive the selection (drilling), not the user
    suppressUntil: 0, // ignore selection-changed events we caused (preview/teleport)
    visible: true // is the pane actually shown? (shared runtime loads it hidden)
  };

  var dom = {};

  // Serialises ALL grid-highlight work (apply + clear) through one promise chain so
  // two Excel.run batches never overlap. Initialised here (before Office.onReady,
  // which can fire synchronously) so queueHL never sees it undefined.
  var hlQueue = Promise.resolve();

  // Formula-editor state: whether the in-pane formula field is open for editing,
  // and the value it held when editing began (to detect a real change on commit).
  var editing = false;
  var formulaOriginal = "";

  // ---- Boot -----------------------------------------------------------------

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      document.getElementById("mt-boot").textContent =
        "MacTrace runs inside Excel. Open it from the Excel ribbon.";
      return;
    }
    cacheDom();
    wireEvents();
    // Stamp the running build into the footer so we can confirm — from inside
    // Excel — that the latest pane code actually loaded (shared-runtime JS only
    // reloads on a full Excel restart).
    var hint = document.querySelector(".mt-hint");
    if (hint) hint.textContent = "build " + (X.BUILD || "?") + " · ↑↓ move · → / ← expand · Enter go · Esc exit";
    document.getElementById("mt-boot").hidden = true;
    document.getElementById("mt-app").hidden = false;
    subscribeSelection();
    setupCommands(); // ribbon/keyboard handlers + visibility tracking
    // Recovery: undo any highlight left painted by a previous session that closed
    // or crashed before it could clean up (state survives in document settings).
    X.clearTemporaryHighlight();
    if (state.visible) { refresh(); focusPane(); }
  });

  // Register the shared-runtime command handlers (used by the ribbon buttons and
  // the Cmd+Option+A / Cmd+Option+Shift+A shortcuts) and track pane visibility.
  function setupCommands() {
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("TracePrecedents", function (event) { activateTrace("precedents", event); });
      Office.actions.associate("TraceDependents", function (event) { activateTrace("dependents", event); });
    }
    if (Office.addin && typeof Office.addin.onVisibilityModeChanged === "function") {
      state.visible = false; // shared runtime starts hidden; wait until shown
      Office.addin.onVisibilityModeChanged(function (args) {
        var shown = args && String(args.visibilityMode) === "Taskpane";
        state.visible = shown;
        if (shown) { refresh(); focusPane(); }
        else clearInspect(); // pane hidden — drop the highlight
      });
    }

    // Clean up the highlight if the pane is closed/unloaded or the document is
    // backgrounded. Best-effort (unload can't await async); document-settings
    // recovery on next load is the backstop for a hard close.
    window.addEventListener("pagehide", function () { clearInspect(); });
    window.addEventListener("beforeunload", function () { clearInspect(); });
    document.addEventListener("visibilitychange", function () {
      if (document.hidden) clearInspect();
    });
  }

  // Invoked by a ribbon command or keyboard shortcut: switch to the requested
  // mode, retrace the active cell, ensure the pane is visible, and grab keyboard
  // focus so the arrow keys work immediately.
  //
  // The retrace (applyMode + refresh) runs FIRST and unconditionally — it must
  // never be gated behind the show/hide sequence below. The old code ran it only
  // inside the .then() of addin.hide().then(() => addin.showAsTaskpane()), with a
  // silent .catch(function(){}) around the whole thing: when the pane was already
  // open and that hide-then-show sequence rejected (flaky on Mac), the retrace
  // simply never happened — the ribbon button "did nothing" while the pane stayed
  // open on the OLD mode. Now the trace always updates; only the extra focus-grab
  // dance is allowed to fail quietly.
  function activateTrace(mode, event) {
    var wasVisible = state.visible;
    state.visible = true;
    applyMode(mode);
    refresh();
    focusPane();

    var addin = Office.addin;
    var seq;
    if (addin && typeof addin.showAsTaskpane === "function") {
      // If the pane was already open, a plain showAsTaskpane() leaves keyboard
      // focus on the grid. Hide then re-show to force the SAME hidden->shown
      // transition that grabs focus on first open. This is now best-effort only.
      if (wasVisible && typeof addin.hide === "function") {
        seq = addin.hide().then(function () { return addin.showAsTaskpane(); });
      } else {
        seq = addin.showAsTaskpane();
      }
    } else {
      seq = Promise.resolve();
    }
    Promise.resolve(seq)
      .then(focusPane, focusPane) // succeed or fail, take one more shot at focus
      .then(function () {
        if (event && event.completed) event.completed();
      });
  }

  // Best-effort: pull keyboard focus into the pane so arrows/Enter/Esc work
  // without the user clicking in first. Excel for Mac may keep focus on the grid.
  function focusPane() {
    // Reinforce focus a few times: the pane needs a moment to settle after the
    // show transition before it will accept focus.
    grabFocus([0, 60, 180]);
  }

  // Lighter-weight focus re-grab used after a live-preview scroll/select, which
  // momentarily hands keyboard focus to the grid. Two quick passes are enough to
  // win it back without the visible flicker of the full open-time sequence.
  function refocusPane() {
    grabFocus([0, 50]);
  }

  function grabFocus(delays) {
    delays.forEach(function (d) {
      setTimeout(function () {
        try { window.focus(); } catch (e) {}
        try { if (dom["mt-content"]) dom["mt-content"].focus(); } catch (e) {}
      }, d);
    });
  }

  function cacheDom() {
    [
      "mt-addr", "mt-value", "mt-sheet", "mt-formula", "mt-content",
      "mt-tab-prec", "mt-tab-dep", "mt-status"
    ].forEach(function (id) {
      dom[id] = document.getElementById(id);
    });
  }

  function wireEvents() {
    dom["mt-tab-prec"].addEventListener("click", function () { setMode("precedents"); });
    dom["mt-tab-dep"].addEventListener("click", function () { setMode("dependents"); });

    // Keydown at the HIGHEST level (window, capture phase) so the pane reacts to
    // arrows/Enter/Esc no matter which descendant has focus — and so Esc is caught
    // as early as possible. The handler bows out while the formula editor is active.
    window.addEventListener("keydown", onKeyDown, true);

    // Formula line: when NOT editing, click a coloured reference to teleport, or
    // click elsewhere to start editing in place. While editing it's contentEditable
    // (clicks position the caret); Enter commits, Esc cancels, blur commits.
    var f = dom["mt-formula"];
    f.addEventListener("click", function (e) {
      if (editing) return;
      var t = e.target;
      if (t && t.classList.contains("tok-ref")) { teleport(t.dataset.nav); return; }
      startEdit();
    });
    f.addEventListener("keydown", onFormulaEditKey);
    f.addEventListener("blur", function () { if (editing) finishEdit(true); });
  }

  // ---- Selection tracking ---------------------------------------------------

  function subscribeSelection() {
    Excel.run(function (ctx) {
      ctx.workbook.onSelectionChanged.add(onSelectionChanged);
      return ctx.sync();
    }).catch(function () {/* non-fatal */});
  }

  var debounceTimer = null;
  function onSelectionChanged() {
    // Ignore selection changes we triggered ourselves (arrow preview / teleport);
    // only rebuild for genuine, user-driven changes while the pane is visible.
    if (Date.now() < state.suppressUntil) return Promise.resolve();
    if (!state.visible) return Promise.resolve();
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(refresh, 70);
    return Promise.resolve();
  }

  // ---- Core refresh ---------------------------------------------------------

  function refresh() {
    if (state.busy) { state.pending = true; return; }
    state.busy = true;
    var wasProgrammatic = state.programmatic;
    state.programmatic = false;
    clearInspect(); // a new audited cell / mode invalidates any inspection highlight
    X.getSelection()
      .then(function (sel) {
        state.selection = sel;
        if (!wasProgrammatic) {
          // A manual selection change starts a fresh audit: reset the rubber band.
          state.origin = { address: sel.address };
          state.navStack = [];
        } else if (!state.origin) {
          state.origin = { address: sel.address };
        }
        renderHeader(sel);
        return state.mode === "precedents" ? renderPrecedents(sel) : renderDependents(sel);
      })
      .catch(function (e) {
        setStatus("Couldn't read the selection: " + (e.message || e), true);
      })
      .then(finishRefresh, finishRefresh);
  }

  function finishRefresh() {
    state.busy = false;
    if (state.pending) { state.pending = false; refresh(); }
  }

  // ---- Header ---------------------------------------------------------------

  function renderHeader(sel) {
    dom["mt-addr"].textContent = sel.local;
    dom["mt-value"].textContent = sel.text != null ? sel.text : "";
    dom["mt-sheet"].textContent = sel.sheet;
    renderFormulaLine(sel);
  }

  function renderFormulaLine(sel) {
    exitEditMode(); // a re-render always returns the formula line to display mode
    var box = dom["mt-formula"];
    box.textContent = "";
    box.title = "Click to edit this formula";
    if (!sel.isFormula) {
      box.appendChild(elText("span", "tok-num", sel.text != null && sel.text !== "" ? sel.text : "(empty cell)"));
      return;
    }
    var tokens = P.tokenize(sel.formula);
    box.appendChild(document.createTextNode("="));
    tokens.forEach(function (tok) {
      var cls = "tok-punct";
      if (tok.type === "ref") cls = "tok-ref";
      else if (tok.type === "function") cls = "tok-fn";
      else if (tok.type === "number") cls = "tok-num";
      else if (tok.type === "string") cls = "tok-str";
      else if (tok.type === "op") cls = "tok-op";
      var span = elText("span", cls, tok.value);
      span.dataset.start = tok.pos;
      span.dataset.end = tok.pos + tok.value.length;
      if (tok.type === "ref") span.dataset.nav = tok.value;
      box.appendChild(span);
    });
  }

  // ---- Editable formula bar -------------------------------------------------
  //
  // The coloured formula line (display + inspect-highlight + click-to-teleport) is
  // preserved; clicking elsewhere on it swaps in a <textarea> pre-filled with the
  // raw formula. Enter (or blur) writes it back via X.setFormula → range.formulas;
  // Esc cancels and drops to the grid.

  // Edit the formula IN PLACE: the same #mt-formula line becomes contentEditable
  // (no separate input box). We swap its coloured tokens for the plain raw formula
  // so editing is clean, then restore the coloured render when done.
  function startEdit() {
    if (!state.selection || editing) return;
    editing = true;
    var box = dom["mt-formula"];
    formulaOriginal = state.selection.formula || "";
    box.textContent = formulaOriginal;
    box.contentEditable = "true";
    box.spellcheck = false;
    box.classList.add("mt-formula--editing");
    box.focus();
    placeCaretEnd(box);
  }

  function exitEditMode() {
    editing = false;
    var box = dom["mt-formula"];
    if (box) {
      box.contentEditable = "false";
      box.classList.remove("mt-formula--editing");
    }
  }

  // Leave edit mode. On commit, if the text actually changed, write it back via
  // X.setFormula → range.formulas and refresh; otherwise just restore the coloured
  // display. Idempotent (a follow-up blur after Enter is a no-op).
  function finishEdit(commit) {
    if (!editing) return;
    var box = dom["mt-formula"];
    var val = box.textContent;
    var changed = commit && state.selection && val !== formulaOriginal;
    if (changed) {
      formulaOriginal = val;
      exitEditMode(); // refresh() will redraw the coloured line with the new formula
      setStatus("Writing formula…");
      X.setFormula(state.selection.address, val).then(function (res) {
        if (res && res.ok) { setStatus("Formula updated."); refresh(); }
        else {
          setStatus("Couldn't write formula: " + (res && (res.message || res.reason) || "error"), true);
          if (state.selection) renderFormulaLine(state.selection);
        }
      });
    } else if (state.selection) {
      renderFormulaLine(state.selection); // restores the coloured tokens + exits edit mode
    } else {
      exitEditMode();
    }
  }

  function onFormulaEditKey(e) {
    if (!editing) return;
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault(); // don't insert a newline
      finishEdit(true);
      if (dom["mt-content"]) dom["mt-content"].focus(); // back to the tree
    } else if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      finishEdit(false); // discard the edit, restore the coloured display
      exitToGrid();       // Esc always ends with keyboard control on the grid
    }
    // all other keys (typing, arrows, Backspace) edit the formula text
  }

  // Put the text caret at the end of a contentEditable element.
  function placeCaretEnd(el) {
    try {
      var range = document.createRange();
      range.selectNodeContents(el);
      range.collapse(false);
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    } catch (e) {}
  }

  // ---- Precedents (formula tree) --------------------------------------------

  function renderPrecedents(sel) {
    var content = dom["mt-content"];
    content.textContent = "";
    document.getElementById("mt-colhead").style.display = "";

    if (!sel.isFormula) {
      var root = makeRow({
        element: sel.local, role: "value", value: sel.text, valueClass: typeClass(sel.valueType),
        location: sel.address, depth: 0, nav: null, expandable: false
      });
      root.row.dataset.origin = sel.address; // arrowing here returns the viewport home
      content.appendChild(root.row);
      content.appendChild(note("This is an input (no formula). Switch to Dependents to see what uses it."));
      setCursor(0, { silent: true });
      return Promise.resolve();
    }

    var parsed = P.parse(sel.formula);
    var refTexts = P.collectRefs(parsed.ast).map(function (n) { return n.text; });

    return X.resolveRefs(refTexts)
      .catch(function (e) {
        // Value-resolution failed — refs still navigate/highlight (see describeNode),
        // they just won't show a live value. Log it so we know why values are blank.
        try { console.error("MacTrace resolveRefs failed:", e); } catch (_) {}
        return {};
      })
      .then(function (refMap) {
        content.textContent = "";
        // Root row = the audited cell.
        var rootInfo = {
          element: sel.local, role: "", value: sel.text, valueClass: typeClass(sel.valueType),
          location: sel.address, depth: 0, nav: null, expandable: true
        };
        var rootRow = makeRow(rootInfo);
        rootRow.row.dataset.origin = sel.address; // arrowing here returns the viewport home
        content.appendChild(rootRow.row);
        var kids = elDiv("mt-children");
        content.appendChild(kids);
        wireTwisty(rootRow, kids);
        renderNode(kids, parsed.ast, "formula", 1, refMap);
        setCursor(0, { silent: true });
      });
  }

  // Recursively render an AST node and its children.
  function renderNode(container, node, role, depth, refMap) {
    var info = describeNode(node, role, refMap);
    info.depth = depth;
    if (node.pos != null) { info.fstart = node.pos; info.fend = node.end; }
    var built = makeRow(info);
    container.appendChild(built.row);

    if (info.children && info.children.length) {
      var kids = elDiv("mt-children");
      container.appendChild(kids);
      wireTwisty(built, kids);
      info.children.forEach(function (ch) {
        renderNode(kids, ch.node, ch.role, depth + 1, refMap);
      });
    }
  }

  // Turn an AST node into row fields + child descriptors.
  function describeNode(node, role, refMap) {
    switch (node.kind) {
      case "func":
        return {
          element: node.name, role: role, value: "", location: "", nav: null,
          func: true, expandable: node.args.length > 0,
          children: node.args.map(function (a, i) {
            return { node: a, role: P.argLabel(node.name, i, node.args.length) };
          })
        };
      case "ref": {
        var rv = refMap[node.text] || {};
        // Navigable unless we KNOW the target is unreachable. We deliberately
        // default to navigable even when resolveRefs returned nothing for this ref
        // (rv == {}), so the live highlight + scroll works off the reference text
        // directly and never silently does nothing just because value-resolution
        // hiccuped. applyTemporaryHighlight / navigateTo re-resolve the range
        // themselves and report a clear reason if it really is unreachable.
        var unreachable = rv.kind === "external" || rv.kind === "missing" || rv.kind === "invalid";
        return {
          element: node.text, role: role, value: rv.display || "",
          valueClass: rv.kind === "cell" ? typeClass(rv.valueType) : "v-" + (rv.kind || ""),
          location: rv.address || "", nav: unreachable ? null : (rv.address || node.text),
          expandable: false, children: []
        };
      }
      case "name":
        return { element: node.text, role: role || "name", value: "", location: "", nav: null, expandable: false, children: [] };
      case "binary":
      case "unary":
      case "group": {
        var operands = P.collectOperands(node);
        return {
          element: P.nodeToText(node), role: role, value: "", location: "", nav: null,
          expandable: operands.length > 0,
          children: operands.map(function (o) { return { node: o, role: "" }; })
        };
      }
      case "number":
        return { element: node.text, role: role, value: node.text, valueClass: "v-number", expandable: false, children: [] };
      case "string":
        return { element: node.text, role: role, value: node.text, valueClass: "v-string", expandable: false, children: [] };
      case "bool":
        return { element: node.text, role: role, value: node.text, expandable: false, children: [] };
      case "error":
        return { element: node.text, role: role, value: node.text, valueClass: "v-error", expandable: false, children: [] };
      case "missing":
        return { element: "‹missing›", role: role, value: "", expandable: false, children: [] };
      default:
        return { element: node.text || "", role: role, value: "", expandable: false, children: [] };
    }
  }

  // ---- Dependents -----------------------------------------------------------

  // Mirrors renderPrecedents exactly: an ORIGIN root row (the cell being analysed,
  // carrying dataset.origin so arrowing back to it scrolls + highlights it) with
  // the dependents as its children. Dependent rows get a `nav` address, so the
  // shared setCursor → inspect flow applies the #27E7F5 highlight and .select()
  // auto-scroll — including across sheets — identically to precedents.
  function renderDependents(sel) {
    var content = dom["mt-content"];
    content.textContent = "";
    document.getElementById("mt-colhead").style.display = "";

    return X.getDirectDependents().then(function (res) {
      content.textContent = "";
      if (!res.supported) {
        content.appendChild(note(
          "Dependents need a newer Excel (ExcelApi 1.13+). Update Microsoft 365 to enable this."
        ));
        return;
      }
      var count = res.items.length;
      // Origin row = the inspected cell (same role as the precedents root).
      var originRow = makeRow({
        element: sel.local, role: "", value: sel.text, valueClass: typeClass(sel.valueType),
        location: sel.address, depth: 0, nav: null, expandable: count > 0
      });
      originRow.row.dataset.origin = sel.address; // arrowing here returns the viewport home
      content.appendChild(originRow.row);

      if (!count) {
        content.appendChild(note(res.error
          ? "Couldn't read dependents — " + res.error
          : "No cells depend on " + sel.selectionLocal + " — safe to change or clear."));
        setCursor(0, { silent: true });
        return;
      }

      content.appendChild(note("Direct dependents (" + count + "). Arrow to preview, Enter to drill in."));
      var kids = elDiv("mt-children");
      content.appendChild(kids);
      wireTwisty(originRow, kids);
      res.items.forEach(function (item) {
        var built = makeRow({
          element: item.local, role: "dependent", value: "", location: item.address,
          depth: 1, nav: item.address, expandable: false
        });
        kids.appendChild(built.row);
      });
      setCursor(0, { silent: true }); // origin row
    });
  }

  // ---- Row factory ----------------------------------------------------------

  function makeRow(info) {
    var row = elDiv("mt-row");
    row.style.setProperty("--depth", info.depth || 0);
    if (info.func) row.classList.add("mt-func");
    if (info.nav) {
      row.classList.add("mt-nav");
      row.dataset.nav = info.nav;
    }
    if (info.fstart != null) {
      row.dataset.fstart = info.fstart;
      row.dataset.fend = info.fend != null ? info.fend : info.fstart;
    }

    var elCol = elDiv("mt-el");
    var twisty = elText("span", "mt-twisty" + (info.expandable ? " mt-open" : " mt-leaf"), "▸");
    elCol.appendChild(twisty);
    elCol.appendChild(elText("span", "mt-el-text", info.element || ""));

    var roleCol = elText("span", "mt-info", info.role || "");
    var valCol = elText("span", "mt-val " + (info.valueClass || ""), info.value != null ? info.value : "");
    var locCol = elText("span", "mt-loc", info.location ? info.location : "");

    row.appendChild(elCol);
    row.appendChild(roleCol);
    row.appendChild(valCol);
    row.appendChild(locCol);

    if (info.nav) {
      row.addEventListener("click", function (e) {
        if (e.target === twisty) return;
        teleport(info.nav);
      });
    }
    return { row: row, twisty: twisty };
  }

  function wireTwisty(built, childrenContainer) {
    var twisty = built.twisty;
    twisty.classList.remove("mt-leaf");
    twisty.classList.add("mt-open");
    twisty.addEventListener("click", function (e) {
      e.stopPropagation();
      toggleCollapse(built.row, childrenContainer, twisty);
    });
  }

  function toggleCollapse(row, childrenContainer, twisty) {
    var collapsed = childrenContainer.classList.toggle("mt-collapsed");
    twisty.classList.toggle("mt-open", !collapsed);
  }

  // ---- Keyboard navigation --------------------------------------------------

  function focusableRows() {
    var rows = dom["mt-content"].querySelectorAll(".mt-row");
    var visible = [];
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].offsetParent !== null) visible.push(rows[i]);
    }
    return visible;
  }

  function setCursor(i, opts) {
    var rows = focusableRows();
    if (!rows.length) return;
    state.cursor = Math.max(0, Math.min(i, rows.length - 1));
    rows.forEach(function (r) { r.classList.remove("mt-cursor"); });
    var cur = rows[state.cursor];
    cur.classList.add("mt-cursor");
    cur.scrollIntoView({ block: "nearest" });
    highlightFormula(cur);
    // Live inspection: on a referenced cell, scroll to + highlight it in the grid;
    // on a non-reference row (a function/operator), drop the highlight. Suppressed
    // on the initial render so opening the pane doesn't move/paint anything.
    if (!(opts && opts.silent)) {
      // Precedent rows scroll + highlight their referenced cell; the ORIGIN row
      // (the formula cell being audited) scrolls + highlights ITSELF, so arrowing
      // back up to the top returns the viewport to where you started. Everything
      // else (functions, operators) just drops the highlight.
      var target = cur.dataset.nav || cur.dataset.origin;
      if (target) inspect(target);
      else clearInspect();
    }
  }

  // Highlight the slice of the formula line that this row corresponds to.
  function highlightFormula(row) {
    var box = dom["mt-formula"];
    var active = box.querySelectorAll(".tok-active");
    for (var i = 0; i < active.length; i++) active[i].classList.remove("tok-active");
    if (row.dataset.fstart == null) return;
    var start = Number(row.dataset.fstart);
    var end = Number(row.dataset.fend);
    var spans = box.children;
    for (var j = 0; j < spans.length; j++) {
      var s = spans[j];
      if (s.dataset && s.dataset.start != null) {
        if (Number(s.dataset.start) >= start && Number(s.dataset.end) <= end) {
          s.classList.add("tok-active");
        }
      }
    }
  }

  function moveCursor(delta) {
    var rows = focusableRows();
    var idx = rows.findIndex(function (r) { return r.classList.contains("mt-cursor"); });
    if (idx === -1) idx = 0;
    setCursor(idx + delta);
  }

  function currentRow() {
    return focusableRows()[state.cursor] || null;
  }

  function onKeyDown(e) {
    // While the formula is being edited in place, let it own every key (typing,
    // arrows, Enter, Backspace, Esc) — don't hijack them for tree navigation.
    if (editing) return;
    switch (e.key) {
      case "ArrowDown": e.preventDefault(); moveCursor(1); break;
      case "ArrowUp": e.preventDefault(); moveCursor(-1); break;
      case "ArrowRight": e.preventDefault(); expandCurrent(true); break;
      case "ArrowLeft": e.preventDefault(); expandCurrent(false); break;
      case "Enter": {
        e.preventDefault();
        var row = currentRow();
        if (row && row.dataset.nav) teleport(row.dataset.nav);
        else if (row) toggleRowFromKeyboard(row);
        break;
      }
      case "Escape":
        // Diagnostic: confirms the webview actually captured the Esc keystroke.
        try { console.log("Escape fired"); } catch (_) {}
        // Deliberately DO NOT preventDefault: on Excel for Mac, the host's own
        // Escape handling is one of the few things that can move keyboard focus
        // out of a task pane, so we let it run alongside our cleanup.
        exitToGrid();
        break;
      case "Backspace": e.preventDefault(); goBack(); break;
      default: break;
    }
  }

  function rowChildren(row) {
    var next = row.nextElementSibling;
    return next && next.classList.contains("mt-children") ? next : null;
  }
  function toggleRowFromKeyboard(row) {
    var kids = rowChildren(row);
    var twisty = row.querySelector(".mt-twisty");
    if (kids && twisty) toggleCollapse(row, kids, twisty);
  }
  function expandCurrent(open) {
    var row = currentRow();
    if (!row) return;
    var kids = rowChildren(row);
    var twisty = row.querySelector(".mt-twisty");
    if (!kids || !twisty) return;
    var collapsed = kids.classList.contains("mt-collapsed");
    if (open && collapsed) toggleCollapse(row, kids, twisty);
    if (!open && !collapsed) toggleCollapse(row, kids, twisty);
  }

  // ---- Navigation (rubber band) ---------------------------------------------

  // Inspect: as you arrow over a referenced cell, bring it into view and paint it
  // #27E7F5 (restoring the previously inspected range first), WITHOUT rebuilding
  // the tree. Debounced so holding an arrow key doesn't thrash the grid. The
  // applyTemporaryHighlight call also selects the range, so it scrolls into view.
  // Serialise grid-highlight work through hlQueue (declared up top): two Excel.run
  // batches never overlap — concurrent fill operations on the grid can throw
  // "object is in an invalid state" errors when arrow keys fire quickly.
  function queueHL(task) {
    hlQueue = hlQueue.then(task, task);
    return hlQueue;
  }

  var inspectTimer = null;
  function inspect(address) {
    if (!address) return;
    state.suppressUntil = Date.now() + 700; // ignore the resulting selection event
    state.programmatic = true;
    clearTimeout(inspectTimer);
    inspectTimer = setTimeout(function () {
      state.suppressUntil = Date.now() + 700;
      queueHL(function () {
        return X.applyTemporaryHighlight(address).then(function (res) {
          // DIAGNOSTIC: surface the full outcome of every highlight attempt so we
          // can see, from inside Excel, whether it painted, why it didn't, or threw.
          reportInspect(address, res);
          // applyTemporaryHighlight calls range.select() to scroll the cell into the
          // viewport, which makes Excel grab keyboard focus. Pull it straight back
          // so the arrow keys keep driving the pane without a manual click.
          refocusPane();
        });
      });
    }, 70);
  }
  function clearInspect() {
    clearTimeout(inspectTimer);
    queueHL(function () { return X.clearTemporaryHighlight(); });
  }

  // Esc: bail out of the inspection flow. Restore any temporary highlight, then
  // hand keyboard control back to the worksheet (selecting the active cell pulls
  // focus off the pane and onto the grid). We do NOT refocus the pane — the whole
  // point is to give the grid control.
  //
  // Runs DIRECTLY rather than through queueHL: Esc is a terminal action and must
  // never be left waiting behind a slow/blocked highlight task. The status line is
  // updated synchronously first, so even if the focus transfer is silently blocked
  // by the host we still get visible proof the handler fired, plus any error text.
  function exitToGrid() {
    setStatus("Esc — returning control to the grid…");
    clearTimeout(inspectTimer);
    state.suppressUntil = Date.now() + 500; // ignore the self-caused selection event
    X.clearTemporaryHighlight()
      .then(function () { return X.returnFocusToGrid(); }) // (re)select the active cell
      .then(function (res) {
        // The cell is now selected, but Excel-for-Mac keeps keyboard focus in the
        // webview unless we actively give it up. Blur the focused element AND the
        // pane window so the OS hands keyboard control to the (already-selected)
        // grid. Repeat once after a tick — the host sometimes needs a beat.
        blurWebview();
        setTimeout(blurWebview, 60);
        if (res && res.ok) setStatus("Returned to the grid.");
        else setStatus("Esc: couldn't move focus" + (res && res.message ? " — " + res.message : ""), true);
      })
      .catch(function (e) {
        setStatus("Esc error: " + (e && e.message ? e.message : e), true);
      });
  }

  // Forcibly relinquish keyboard focus from the task-pane webview so Excel can
  // take it back to the grid. window.blur() is the part that actually drops the
  // webview's hold on the OS keyboard focus on Mac.
  function blurWebview() {
    try { if (document.activeElement && document.activeElement.blur) document.activeElement.blur(); } catch (e) {}
    try { if (dom["mt-content"]) dom["mt-content"].blur(); } catch (e) {}
    try { window.blur(); } catch (e) {}
    try { if (top && top !== window && top.blur) top.blur(); } catch (e) {}
  }

  // DIAGNOSTIC readout for the highlight: writes a compact summary to the status
  // line, e.g. "G518 → painted" / "G518 → FAIL: missing-sheet" / "G518 → ERROR …".
  function reportInspect(address, res) {
    var bang = String(address).lastIndexOf("!");
    var local = bang === -1 ? address : String(address).slice(bang + 1);
    var r = res || {};
    var msg;
    var err = false;
    if (r.ok) {
      msg = local + " → painted" + (r.scrolled === false ? " (no scroll)" : "");
    } else if (r.reason === "error") {
      msg = local + " → ERROR: " + (r.message || "unknown");
      err = true;
    } else {
      msg = local + " → FAIL: " + (r.reason || "unknown") + (r.sheet ? " ['" + r.sheet + "']" : "");
      err = true;
    }
    setStatus(msg, err);
  }

  // Commit: make a referenced cell the new audited cell and rebuild the tree there.
  function teleport(address) {
    if (!address) return;
    if (state.selection) state.navStack.push(state.selection.address);
    navigateAndRebuild(address, function (res) {
      if (!res.ok) {
        if (res.reason === "external") setStatus("That reference is in another workbook — open it to follow.", true);
        else if (res.reason === "missing-sheet") setStatus("Source sheet not found.", true);
        else setStatus("Couldn't navigate there.", true);
        state.navStack.pop();
        return false; // don't rebuild
      }
      setStatus("");
      return true;
    });
  }

  // Navigate Excel and explicitly rebuild the pane at the destination, ignoring
  // the selection event we cause (so we rebuild even if we'd pre-previewed there).
  function navigateAndRebuild(address, onResult) {
    state.programmatic = true;
    state.suppressUntil = Date.now() + 700;
    clearTimeout(inspectTimer); // cancel any pending live-preview highlight
    X.navigateTo(address).then(function (res) {
      var ok = onResult ? onResult(res) : res.ok;
      if (ok) {
        state.programmatic = true; // refresh() consumes this; keeps the origin
        refresh();
      }
    });
  }

  // Backspace steps back one hop through the drill history (no button — keyboard
  // only). navStack is still maintained by teleport, so this keeps working.
  function goBack() {
    if (!state.navStack.length) return;
    var addr = state.navStack.pop();
    navigateAndRebuild(addr);
  }

  // ---- Mode ------------------------------------------------------------------

  function applyMode(mode) {
    state.mode = mode;
    dom["mt-tab-prec"].classList.toggle("mt-tab--active", mode === "precedents");
    dom["mt-tab-dep"].classList.toggle("mt-tab--active", mode === "dependents");
  }
  function setMode(mode) {
    if (state.mode === mode) return;
    applyMode(mode);
    refresh();
  }

  // ---- Small DOM/util helpers ----------------------------------------------

  function elDiv(cls) {
    var d = document.createElement("div");
    if (cls) d.className = cls;
    return d;
  }
  function elText(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function note(text) {
    return elText("div", "mt-note", text);
  }
  function typeClass(valueType) {
    switch (valueType) {
      case "Double": return "v-number";
      case "String": return "v-string";
      case "Error": return "v-error";
      default: return "";
    }
  }
  function setStatus(text, isErr) {
    dom["mt-status"].textContent = text || "";
    dom["mt-status"].classList.toggle("mt-err", !!isErr);
  }
})();
