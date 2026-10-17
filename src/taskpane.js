/*
 * MacTrace — task pane controller
 * -------------------------------
 * Owns the UI: renders the precedents tree or the dependents list for the
 * audited cell, and drives keyboard navigation, live preview and drilling. A
 * cell is traced only when you ask (ribbon button, shortcut, first open) or
 * drill / step back, never because Excel's selection moved. See trace().
 */
(function () {
  "use strict";

  var P = window.MacTraceParser;
  var X = window.MacTraceExcel;

  var state = {
    mode: "precedents", // "precedents" | "dependents"
    selection: null, // the audited target (shape of X.getSelection / X.getTargetInfo)
    navStack: [], // previously audited targets { address, selection }, for Back
    cursor: 0, // index into the currently focusable rows
    busy: false,
    pending: null, // trace request queued while busy (see trace())
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
  var writing = false; // an in-place formula write (and its re-render) is in flight
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
    if (hint) hint.textContent = "build " + (X.BUILD || "?") + " · ⌘⌥A trace · ↑↓ move · → / ← expand · Enter go · Esc close";
    document.getElementById("mt-boot").hidden = true;
    document.getElementById("mt-app").hidden = false;
    watchSelection();
    setupCommands(); // ribbon/keyboard handlers + visibility tracking
    // Recovery: undo any highlight left painted by a previous session that closed
    // or crashed before it could clean up (state survives in document settings).
    X.clearTemporaryHighlight();
    if (state.visible) { trace("active"); focusPane(); }
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
        if (shown) {
          // Being shown is not a request to retrace: the ribbon button / shortcut
          // (activateTrace) has already traced the active cell. Only an empty
          // pane, with nothing audited yet, traces once so it never opens blank.
          if (!state.selection && !state.busy) trace("active");
          focusPane();
        } else clearInspect(); // pane hidden — drop the highlight
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
  // The retrace (applyMode + trace) runs FIRST and unconditionally — it must
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
    setStatus(""); // clear any stale status (e.g. "Closing MacTrace…" from a prior Esc)
    applyMode(mode);
    trace("active");
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

    // Formula line: a click anywhere, references included, edits the formula in
    // place with the caret where you clicked. It deliberately does NOT jump to a
    // clicked reference: that moved the grid and retraced the pane when you only
    // meant to edit (drill from the tree instead: Enter or click a row). While
    // editing it's contentEditable; Enter commits, Esc cancels, blur commits.
    var f = dom["mt-formula"];
    f.addEventListener("click", function (e) {
      if (editing) return;
      startEdit(formulaOffsetAt(e.clientX, e.clientY)); // read before the line is re-rendered
    });
    f.addEventListener("keydown", onFormulaEditKey);
    f.addEventListener("blur", function () { if (editing) finishEdit(true); });
  }

  // ---- Grid clicks ----------------------------------------------------------
  //
  // Excel's selection never retraces the pane (see trace()). It's watched only so
  // that a click into the grid (a) restores the preview fill instead of leaving it
  // on the cell and (b) becomes the baseline that stops a preview still in flight
  // from selecting its cell over your click. Each check runs on the grid queue,
  // behind any preview, and compares against the exact selection the pane itself
  // last left (X.reconcileSelection). No time window involved.
  function watchSelection() {
    Excel.run(function (ctx) {
      ctx.workbook.onSelectionChanged.add(onSelectionChanged);
      return ctx.sync();
    }).catch(function () {/* non-fatal: the fill is still restored on the next move or close */});
  }

  var selTimer = null;
  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(function () {
      queueHL(function () { return X.reconcileSelection(); });
    }, 100);
    return Promise.resolve();
  }

  // ---- Core trace -----------------------------------------------------------
  //
  // The pane audits an explicit target and never re-reads Excel's selection on
  // its own. It used to rebuild on every selection change behind a 700 ms
  // "ignore our own events" window, but live preview selects each referenced
  // cell to scroll it into view, and whenever that select landed after the
  // window (easy, with several Excel round-trips queued) the tree jumped to the
  // precedent's formula. Now:
  //
  //   trace("active")               ribbon button / shortcut / first open: the
  //                                 active cell, read once any preview in flight
  //                                 has settled
  //   trace({ address, selection }) drill (Enter, click) and Back: that cell
  //   trace("current")              re-render the audited cell (tab switch, edit)
  var renderedMode = null; // the mode #mt-content was last rendered in
  function trace(req) {
    if (state.busy) {
      // Coalesce: a re-render never displaces a queued change of target.
      if (!(req === "current" && state.pending)) state.pending = req;
      return;
    }
    if (req === "current" && !state.selection) req = "active"; // nothing audited yet
    state.busy = true;
    clearInspect(); // a new audited cell / mode invalidates any inspection highlight
    var prev = state.selection;
    var read;
    if (req === "active") read = queueHL(function () { return X.getSelection(); });
    else if (req === "current") read = X.getTargetInfo(prev.selection, prev.address);
    else read = X.getTargetInfo(req.selection || req.address, req.address);
    read
      .then(function (sel) {
        // Deliberately tracing a different cell starts a fresh audit, so the Back
        // history goes. (Drill and Back maintain navStack themselves.)
        if (req === "active" && !(prev && prev.address === sel.address)) state.navStack = [];
        state.selection = sel;
        renderedMode = state.mode;
        renderHeader(sel);
        return state.mode === "precedents" ? renderPrecedents(sel) : renderDependents(sel);
      })
      .catch(function (e) {
        var msg = e.message || e;
        if (req === "active") {
          setStatus("Couldn't read the selection: " + msg, true);
          // The ribbon button / shortcut may have switched the tab first: show the
          // audited cell in that view rather than leave the old one under it.
          if (state.selection && renderedMode !== state.mode && !state.pending) state.pending = "current";
          return;
        }
        renderedMode = state.mode;
        // The named cell can't be read (e.g. its sheet was renamed or deleted).
        // Clear the view rather than leave the old one under a newly active tab.
        var what = req === "current" ? prev.local : String(req.address).split("!").pop();
        dom["mt-content"].textContent = "";
        dom["mt-content"].appendChild(note("Couldn't read " + what + ". Select a cell and press ⌘⌥A to trace it."));
        setStatus("Couldn't read " + what + ": " + msg, true);
      })
      .then(finishTrace, finishTrace);
  }

  function finishTrace() {
    state.busy = false;
    var next = state.pending;
    state.pending = null;
    if (next) trace(next);
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
      box.appendChild(span);
    });
  }

  // ---- Editable formula bar -------------------------------------------------
  //
  // The coloured formula line (display + inspect-highlight) is preserved; a click
  // anywhere on it turns it into an editable field holding the raw formula, with
  // the caret where you clicked. Enter (or blur) writes it back via X.setFormula
  // → range.formulas; Esc cancels.

  // Edit the formula IN PLACE: the same #mt-formula line becomes contentEditable
  // (no separate input box). We swap its coloured tokens for the plain raw formula
  // so editing is clean, then restore the coloured render when done.
  function startEdit(caretAt) {
    if (!state.selection || editing || writing || state.busy) return; // let a pending write / re-render land
    editing = true;
    var box = dom["mt-formula"];
    formulaOriginal = state.selection.formula || "";
    box.textContent = formulaOriginal;
    box.contentEditable = "true";
    box.spellcheck = false;
    box.classList.add("mt-formula--editing");
    box.focus();
    if (caretAt == null) placeCaretEnd(box);
    else placeCaretAt(box, caretAt);
  }

  // Where in the raw formula a click on the coloured line landed (index into the
  // formula including its "="), or null if unknown. The coloured line drops
  // whitespace, so this maps through each token's data-start (its offset after
  // the "=") instead of counting the characters on screen.
  function formulaOffsetAt(x, y) {
    var box = dom["mt-formula"];
    var node = null, off = 0;
    try {
      if (document.caretRangeFromPoint) {
        var r = document.caretRangeFromPoint(x, y);
        if (r) { node = r.startContainer; off = r.startOffset; }
      } else if (document.caretPositionFromPoint) {
        var p = document.caretPositionFromPoint(x, y);
        if (p) { node = p.offsetNode; off = p.offset; }
      }
    } catch (e) { return null; }
    if (!node || node.nodeType !== 3) return null; // padding / past the end
    if (node === box.firstChild) return Math.min(off, 1); // the leading "="
    var el = node.parentNode;
    if (!el || el.parentNode !== box || el.dataset.start == null) return null;
    return 1 + Number(el.dataset.start) + off;
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
  // X.setFormula → range.formulas and re-render; otherwise just restore the
  // coloured display. Idempotent (a follow-up blur after Enter is a no-op).
  // The pane is a snapshot (it doesn't follow the grid), so the write only goes
  // through if the cell still holds the formula the editor started from.
  function finishEdit(commit) {
    if (!editing) return;
    var box = dom["mt-formula"];
    var val = box.textContent;
    var changed = commit && state.selection && val !== formulaOriginal;
    if (changed) {
      var base = formulaOriginal; // what the cell held when the editor opened
      formulaOriginal = val;
      exitEditMode(); // trace() will redraw the coloured line with the new formula
      setStatus("Writing formula…");
      writing = true;
      X.setFormula(state.selection.address, val, base).then(function (res) {
        writing = false; // trace() below keeps startEdit waiting until it has re-rendered
        if (res && res.ok) { setStatus("Formula updated."); trace("current"); }
        else if (res && res.reason === "changed") {
          // Edited in the grid (or shifted by an inserted row) since it was
          // traced: don't overwrite it. Show what it holds now instead.
          setStatus("Not written: this cell changed in the grid since it was traced. Showing it as it is now.", true);
          trace("current");
        } else {
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
      if (dom["mt-content"]) dom["mt-content"].focus(); // back to the tree; pane stays open
      // Note: Esc here only cancels the in-place edit. Press Esc again (not
      // editing) to close the pane.
    }
    // all other keys (typing, arrows, Backspace) edit the formula text
  }

  // Put the text caret at character `offset` of a contentEditable element that
  // holds a single text node (clamped to its length).
  function placeCaretAt(el, offset) {
    var text = el.firstChild;
    if (!text || text.nodeType !== 3) return placeCaretEnd(el);
    try {
      var range = document.createRange();
      range.setStart(text, Math.max(0, Math.min(offset, text.length)));
      range.collapse(true);
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    } catch (e) { placeCaretEnd(el); }
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

    return X.resolveRefs(refTexts, sel.sheet)
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

    return X.getDirectDependents(sel.sheet, sel.selectionLocal).then(function (res) {
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
        // Esc closes the pane entirely (host-level hide). We don't preventDefault
        // so the host's native Escape can run alongside our cleanup.
        closePane();
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

  // ---- Navigation (preview, drill, back) ------------------------------------

  // Inspect: as you arrow over a referenced cell, bring it into view and paint it
  // #27E7F5 (restoring the previously inspected range first), WITHOUT rebuilding
  // the tree. Debounced so holding an arrow key doesn't thrash the grid. The
  // applyTemporaryHighlight call also selects the range, so it scrolls into view;
  // that moves Excel's selection, which is harmless: the pane reads the selection
  // only for a deliberate trace.
  // Serialise grid work through hlQueue (declared up top): two Excel.run batches
  // never overlap (concurrent fill operations on the grid can throw "object is in
  // an invalid state" when arrow keys fire quickly), and anything that must see
  // the settled selection (a trace, a grid-click check) queues behind it.
  function queueHL(task) {
    hlQueue = hlQueue.then(task, task);
    return hlQueue;
  }

  // Each preview carries a generation number. A newer preview, or anything that
  // moves on (a trace, a drill, closing the pane), bumps it, so previews still
  // waiting on the grid queue become no-ops and the one in flight stops before
  // it selects its cell. Otherwise a backlog built up by tapping ↓ could select
  // a precedent AFTER you clicked another cell, and a trace would read that.
  var inspectTimer = null;
  var inspectGen = 0;
  function inspect(address) {
    // While a trace is reading, the rows on screen belong to the tree it is about
    // to replace: never preview (select) one of their cells.
    if (!address || state.busy) return;
    address = qualify(address);
    var gen = ++inspectGen;
    clearTimeout(inspectTimer);
    inspectTimer = setTimeout(function () {
      queueHL(function () {
        if (gen !== inspectGen) return; // superseded while it waited
        var current = function () { return gen === inspectGen; };
        return X.applyTemporaryHighlight(address, current).then(function (res) {
          if (res && res.superseded) return;
          // applyTemporaryHighlight calls range.select() to scroll the cell into the
          // viewport, which makes Excel grab keyboard focus. Pull it straight back
          // so the arrow keys keep driving the pane without a manual click. Not
          // while the pane is closing (Esc hands focus to the grid on purpose).
          if (state.visible && (current() || (res && res.scrolled))) refocusPane();
          // DIAGNOSTIC: surface the full outcome of every highlight attempt so we
          // can see, from inside Excel, whether it painted, why it didn't, or threw.
          // Only for a preview the pane hasn't moved past since.
          if (current()) reportInspect(address, res);
        });
      });
    }, 70);
  }
  function clearInspect() {
    clearTimeout(inspectTimer);
    inspectGen++; // cancel previews that haven't selected their cell yet
    queueHL(function () { return X.clearTemporaryHighlight(); });
  }

  // Esc — close the pane entirely. Hiding is a host-level operation through the
  // shared-runtime visibility API; when the pane disappears the grid is the only
  // surface left, so keyboard focus lands there by default. This sidesteps the
  // webview->grid focus transfer that Excel-for-Mac otherwise blocks (the reason
  // the old "return focus to grid on Esc" never worked). The highlight is
  // restored on the grid queue, so it never overlaps a preview's Excel batch;
  // hide() doesn't wait for it (the runtime lives on, and the Hidden handler
  // queues the same cleanup).
  function closePane() {
    setStatus("Closing MacTrace…");
    state.visible = false; // a late preview must not pull focus back into a closing pane
    clearInspect();
    return Promise.resolve()
      .then(function () {
        var addin = Office.addin;
        if (addin && typeof addin.hide === "function") return addin.hide();
        // No shared-runtime hide() (unsupported host): best-effort blur so focus
        // at least leaves the webview.
        blurWebview();
      })
      .catch(function (e) {
        state.visible = true;
        setStatus("Esc: couldn't close the pane" + (e && e.message ? " — " + e.message : ""), true);
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
    address = qualify(address);
    // Record where we came from right away, so a Backspace pressed mid-jump steps
    // back to it; but not twice in a row (a double click), and take back exactly
    // this entry if the jump fails.
    var from = null;
    var top = state.navStack[state.navStack.length - 1];
    if (state.selection && !(top && top.address === state.selection.address)) {
      from = { address: state.selection.address, selection: state.selection.selection };
      state.navStack.push(from);
    }
    navigateAndRebuild({ address: address }, function (res) {
      if (!res.ok) {
        setStatus(jumpError(res), true);
        var i = from ? state.navStack.lastIndexOf(from) : -1;
        if (i !== -1) state.navStack.splice(i, 1);
        return false; // don't rebuild
      }
      setStatus("");
      return true;
    });
  }

  // Drill / Back: move Excel to the target so the grid follows you, then audit
  // that exact target, never whatever Excel reports as selected afterwards. The
  // jump queues behind any preview still in flight, so a late preview select
  // can't land after it.
  function navigateAndRebuild(target, onResult) {
    clearTimeout(inspectTimer); // cancel any pending live-preview highlight
    inspectGen++;
    queueHL(function () { return X.navigateTo(target.address); }).then(function (res) {
      var ok = onResult ? onResult(res) : res.ok;
      if (!ok) return;
      trace(target);
      if (state.visible) refocusPane(); // the jump's select() handed keyboard focus to the grid
    });
  }

  function jumpError(res) {
    if (res.reason === "external") return "That reference is in another workbook — open it to follow.";
    if (res.reason === "missing-sheet") return "Source sheet not found.";
    return "Couldn't navigate there.";
  }

  // Backspace steps back one hop through the drill history (no button — keyboard
  // only). navStack is maintained by teleport. An entry that can't be reached any
  // more (its sheet was renamed or deleted) is reported and dropped, so the next
  // Backspace goes further back instead of failing on it forever.
  function goBack() {
    if (!state.navStack.length) return;
    var entry = state.navStack.pop();
    navigateAndRebuild(entry, function (res) {
      if (res.ok) { setStatus(""); return true; }
      setStatus("Couldn't go back to " + String(entry.address).split("!").pop() + ": " + jumpError(res), true);
      return false;
    });
  }

  // A bare reference in a formula lives on the formula's own sheet. Pin it there:
  // preview may have activated another sheet, and the Excel layer resolves an
  // unqualified reference against the ACTIVE one.
  function qualify(ref) {
    if (!ref || String(ref).indexOf("!") !== -1 || !state.selection) return ref;
    return "'" + String(state.selection.sheet).replace(/'/g, "''") + "'!" + ref;
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
    trace("current"); // same audited cell, other view; never the live selection
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
