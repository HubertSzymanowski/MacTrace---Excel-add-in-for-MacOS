/*
 * MacTrace formula parser
 * -----------------------
 * A small, dependency-free tokenizer + recursive-descent parser that turns an
 * Excel formula string into a tree we can render and navigate. It is built for
 * auditing, not evaluation: it never tries to compute results (Excel already did
 * that). Its job is to expose the formula's logical structure and, crucially,
 * every cell/range reference so we can resolve live values and "teleport" to them.
 *
 * Exposed as a global: window.MacTraceParser
 */
(function (global) {
  "use strict";

  // ---- Token scanner --------------------------------------------------------

  // A worksheet/workbook prefix, e.g.  Sheet1!  'My Sheet'!  [Book.xlsx]Sheet1!  [1]Sheet1!
  var SHEET_PREFIX =
    "(?:'(?:[^']|'')*'|(?:\\[[^\\]]+\\])?[A-Za-z_][A-Za-z0-9_.]*)!";
  // The reference body: a cell, a cell range, a whole-column range, or a whole-row range.
  var CELL = "\\$?[A-Za-z]{1,3}\\$?\\d+";
  var REF_BODY =
    "(?:" +
    CELL + "(?::" + CELL + ")?" + // A1  or  A1:B2
    "|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}" + // A:A   (whole columns)
    "|\\$?\\d+:\\$?\\d+" + // 1:1   (whole rows)
    ")";

  var RULES = [
    ["ws", /^\s+/],
    ["string", /^"(?:[^"]|"")*"/],
    ["error", /^#(?:REF!|DIV\/0!|N\/A|NAME\?|NULL!|NUM!|VALUE!|GETTING_DATA|SPILL!|CALC!)/i],
    // A function is an identifier immediately followed by "(".
    ["function", /^(?:_xlfn\.)?[A-Za-z_][A-Za-z0-9_.]*(?=\s*\()/],
    // A reference (optionally sheet/workbook qualified). The whole sheet prefix
    // is optional as a unit — note the wrapping (?:...)? around SHEET_PREFIX.
    ["ref", new RegExp("^(?:" + SHEET_PREFIX + ")?" + REF_BODY + "(?![A-Za-z0-9_])")],
    // A sheet-qualified defined name, e.g. Sheet1!MyName
    ["name", new RegExp("^" + SHEET_PREFIX + "[A-Za-z_\\\\][A-Za-z0-9_.]*")],
    ["number", /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/],
    ["bool", /^(?:TRUE|FALSE)(?![A-Za-z0-9_(])/i],
    ["name", /^[A-Za-z_\\][A-Za-z0-9_.]*/], // bare defined name
    ["op", /^(?:<=|>=|<>|[+\-*/^&=<>%])/],
    ["punct", /^[(),{};:]/]
  ];

  function tokenize(formula) {
    var s = String(formula);
    if (s.charAt(0) === "=") s = s.slice(1);
    var tokens = [];
    var pos = 0;
    while (pos < s.length) {
      var rest = s.slice(pos);
      var matched = false;
      for (var i = 0; i < RULES.length; i++) {
        var m = RULES[i][1].exec(rest);
        if (m && m[0].length > 0) {
          var type = RULES[i][0];
          if (type !== "ws") {
            tokens.push({ type: type, value: m[0], pos: pos });
          }
          pos += m[0].length;
          matched = true;
          break;
        }
      }
      if (!matched) {
        // Unknown character: emit it so the renderer can still show the formula.
        tokens.push({ type: "unknown", value: s.charAt(pos), pos: pos });
        pos += 1;
      }
    }
    return tokens;
  }

  // ---- Parser ---------------------------------------------------------------
  //
  // Produces an AST of nodes. Node shapes:
  //   { kind: "func", name, args: [node,...] }
  //   { kind: "binary", op, left, right }
  //   { kind: "unary", op, operand }      // leading +/-  or trailing %
  //   { kind: "ref",  text }              // cell or range reference
  //   { kind: "name", text }              // defined name
  //   { kind: "number"|"string"|"bool"|"error", text }
  //   { kind: "array", rows: [[node,...],...] }

  function endOf(tok) {
    return tok.pos + String(tok.value).length;
  }

  function Parser(tokens) {
    this.t = tokens;
    this.i = 0;
    this.lastTok = null;
  }
  Parser.prototype.peek = function () {
    return this.t[this.i];
  };
  Parser.prototype.next = function () {
    this.lastTok = this.t[this.i++];
    return this.lastTok;
  };
  Parser.prototype.isPunct = function (v) {
    var tok = this.peek();
    return tok && tok.type === "punct" && tok.value === v;
  };
  Parser.prototype.isOp = function (set) {
    var tok = this.peek();
    return tok && tok.type === "op" && set.indexOf(tok.value) !== -1;
  };

  Parser.prototype.parse = function () {
    if (!this.peek()) return { kind: "empty" };
    var node = this.parseComparison();
    return node;
  };

  Parser.prototype.parseComparison = function () {
    var node = this.parseConcat();
    while (this.isOp(["=", "<>", "<", ">", "<=", ">="])) {
      var op = this.next().value;
      node = { kind: "binary", op: op, left: node, right: this.parseConcat() };
    }
    return node;
  };
  Parser.prototype.parseConcat = function () {
    var node = this.parseAdditive();
    while (this.isOp(["&"])) {
      var op = this.next().value;
      node = { kind: "binary", op: op, left: node, right: this.parseAdditive() };
    }
    return node;
  };
  Parser.prototype.parseAdditive = function () {
    var node = this.parseMultiplicative();
    while (this.isOp(["+", "-"])) {
      var op = this.next().value;
      node = { kind: "binary", op: op, left: node, right: this.parseMultiplicative() };
    }
    return node;
  };
  Parser.prototype.parseMultiplicative = function () {
    var node = this.parseExponent();
    while (this.isOp(["*", "/"])) {
      var op = this.next().value;
      node = { kind: "binary", op: op, left: node, right: this.parseExponent() };
    }
    return node;
  };
  Parser.prototype.parseExponent = function () {
    var node = this.parseUnary();
    while (this.isOp(["^"])) {
      var op = this.next().value;
      node = { kind: "binary", op: op, left: node, right: this.parseUnary() };
    }
    return node;
  };
  Parser.prototype.parseUnary = function () {
    if (this.isOp(["-", "+"])) {
      var opTok = this.next();
      var operand = this.parseUnary();
      return { kind: "unary", op: opTok.value, operand: operand, pos: opTok.pos, end: operand ? operand.end : endOf(opTok) };
    }
    return this.parsePostfix();
  };
  Parser.prototype.parsePostfix = function () {
    var node = this.parsePrimary();
    while (this.isOp(["%"])) {
      var pTok = this.next();
      node = { kind: "unary", op: "%", operand: node, postfix: true, pos: node.pos, end: endOf(pTok) };
    }
    // Range/intersection operators between references are kept as binary nodes
    // so their reference leaves are still discoverable.
    while (this.isPunct(":") || (this.peek() && this.peek().type === "op" && false)) {
      this.next();
      node = { kind: "binary", op: ":", left: node, right: this.parsePrimary() };
    }
    return node;
  };
  Parser.prototype.parsePrimary = function () {
    var tok = this.peek();
    if (!tok) return { kind: "empty" };

    if (tok.type === "function") {
      this.next(); // function name
      this.next(); // "("
      var args = this.parseArgs();
      var fnode = { kind: "func", name: tok.value, args: args, pos: tok.pos, end: endOf(tok) };
      if (this.lastTok) fnode.end = endOf(this.lastTok); // extend to the ")"
      return fnode;
    }
    if (this.isPunct("(")) {
      var openTok = this.next();
      var inner = this.parseComparison();
      var gnode = { kind: "group", expr: inner, pos: openTok.pos, end: inner ? inner.end : endOf(openTok) };
      if (this.isPunct(")")) gnode.end = endOf(this.next());
      return gnode;
    }
    if (this.isPunct("{")) {
      return this.parseArray();
    }
    this.next();
    var p = tok.pos, e = endOf(tok);
    switch (tok.type) {
      case "ref":
        return { kind: "ref", text: tok.value, pos: p, end: e };
      case "name":
        return { kind: "name", text: tok.value, pos: p, end: e };
      case "number":
        return { kind: "number", text: tok.value, pos: p, end: e };
      case "string":
        return { kind: "string", text: tok.value, pos: p, end: e };
      case "bool":
        return { kind: "bool", text: tok.value, pos: p, end: e };
      case "error":
        return { kind: "error", text: tok.value, pos: p, end: e };
      default:
        return { kind: "raw", text: tok.value, pos: p, end: e };
    }
  };
  Parser.prototype.parseArgs = function () {
    var args = [];
    if (this.isPunct(")")) {
      this.next();
      return args;
    }
    // Handle leading/empty arguments such as OFFSET(A1,,2)
    while (true) {
      if (this.isPunct(",") || this.isPunct(")")) {
        args.push({ kind: "missing" });
      } else {
        args.push(this.parseComparison());
      }
      if (this.isPunct(",")) {
        this.next();
        continue;
      }
      break;
    }
    if (this.isPunct(")")) this.next();
    return args;
  };
  Parser.prototype.parseArray = function () {
    this.next(); // "{"
    var rows = [[]];
    while (this.peek() && !this.isPunct("}")) {
      if (this.isPunct(";")) {
        this.next();
        rows.push([]);
      } else if (this.isPunct(",")) {
        this.next();
      } else {
        rows[rows.length - 1].push(this.parseComparison());
      }
    }
    if (this.isPunct("}")) this.next();
    return { kind: "array", rows: rows };
  };

  // ---- Argument labels for common functions ---------------------------------
  // Mirrors the "logical_test / value_if_true" style labels seen in the screenshot.
  var ARG_NAMES = {
    IF: ["logical_test", "value_if_true", "value_if_false"],
    IFS: ["condition", "value"],
    IFERROR: ["value", "value_if_error"],
    IFNA: ["value", "value_if_na"],
    VLOOKUP: ["lookup_value", "table_array", "col_index_num", "range_lookup"],
    HLOOKUP: ["lookup_value", "table_array", "row_index_num", "range_lookup"],
    XLOOKUP: ["lookup_value", "lookup_array", "return_array", "if_not_found", "match_mode", "search_mode"],
    LOOKUP: ["lookup_value", "lookup_vector", "result_vector"],
    INDEX: ["array", "row_num", "column_num"],
    MATCH: ["lookup_value", "lookup_array", "match_type"],
    OFFSET: ["reference", "rows", "cols", "height", "width"],
    INDIRECT: ["ref_text", "a1"],
    SUMIF: ["range", "criteria", "sum_range"],
    SUMIFS: ["sum_range", "criteria_range", "criteria"],
    COUNTIF: ["range", "criteria"],
    COUNTIFS: ["criteria_range", "criteria"],
    AVERAGEIF: ["range", "criteria", "average_range"],
    ROUND: ["number", "num_digits"],
    ROUNDUP: ["number", "num_digits"],
    ROUNDDOWN: ["number", "num_digits"],
    SUMPRODUCT: ["array"],
    CHOOSE: ["index_num", "value"],
    NPV: ["rate", "value"],
    IRR: ["values", "guess"],
    XNPV: ["rate", "values", "dates"],
    XIRR: ["values", "dates", "guess"],
    PMT: ["rate", "nper", "pv", "fv", "type"],
    PV: ["rate", "nper", "pmt", "fv", "type"],
    FV: ["rate", "nper", "pmt", "pv", "type"]
  };

  function argLabel(funcName, index, total) {
    var key = String(funcName).toUpperCase().replace(/^_XLFN\./, "");
    var labels = ARG_NAMES[key];
    if (!labels) return "arg " + (index + 1);
    if (index < labels.length) return labels[index];
    // For variadic tails (e.g. SUMIFS criteria pairs) reuse the last labels.
    var tail = labels.slice(Math.max(1, labels.length - 1));
    return tail[(index - labels.length) % tail.length] || "arg " + (index + 1);
  }

  // ---- Public helpers -------------------------------------------------------

  // Collect every reference node in the tree, in reading order.
  function collectRefs(node, out) {
    out = out || [];
    if (!node || typeof node !== "object") return out;
    if (node.kind === "ref") out.push(node);
    var children = childrenOf(node);
    for (var i = 0; i < children.length; i++) collectRefs(children[i], out);
    return out;
  }

  // Collect the meaningful operands of a compound (non-function) expression:
  // descend through operators/groups, but stop at — and surface — function calls
  // and references so each becomes its own navigable/expandable tree node.
  function collectOperands(node, out) {
    out = out || [];
    if (!node || typeof node !== "object") return out;
    if (node.kind === "func" || node.kind === "ref" || node.kind === "name") {
      out.push(node);
      return out;
    }
    var children = childrenOf(node);
    for (var i = 0; i < children.length; i++) collectOperands(children[i], out);
    return out;
  }

  function childrenOf(node) {
    switch (node.kind) {
      case "func":
        return node.args;
      case "binary":
        return [node.left, node.right];
      case "unary":
        return [node.operand];
      case "group":
        return [node.expr];
      case "array":
        return node.rows.reduce(function (a, r) {
          return a.concat(r);
        }, []);
      default:
        return [];
    }
  }

  // Render an AST node back to a compact formula string (for sub-expression labels).
  function nodeToText(node) {
    if (!node) return "";
    switch (node.kind) {
      case "func":
        return (
          node.name +
          "(" +
          node.args.map(nodeToText).join(",") +
          ")"
        );
      case "binary":
        return nodeToText(node.left) + node.op + nodeToText(node.right);
      case "unary":
        return node.postfix
          ? nodeToText(node.operand) + node.op
          : node.op + nodeToText(node.operand);
      case "group":
        return "(" + nodeToText(node.expr) + ")";
      case "array":
        return (
          "{" +
          node.rows
            .map(function (r) {
              return r.map(nodeToText).join(",");
            })
            .join(";") +
          "}"
        );
      case "missing":
        return "";
      case "ref":
      case "name":
      case "number":
      case "string":
      case "bool":
      case "error":
      case "raw":
        return node.text;
      default:
        return "";
    }
  }

  // Give every node a source span [pos, end) covering itself and its children,
  // so the UI can highlight the matching slice of the formula text.
  function computeSpans(node) {
    if (!node || typeof node !== "object") return null;
    var lo = node.pos, hi = node.end;
    var kids = childrenOf(node);
    for (var i = 0; i < kids.length; i++) {
      var s = computeSpans(kids[i]);
      if (s) {
        if (lo == null || s.pos < lo) lo = s.pos;
        if (hi == null || s.end > hi) hi = s.end;
      }
    }
    if (lo != null) {
      node.pos = lo;
      node.end = hi;
    }
    return lo != null ? { pos: lo, end: hi } : null;
  }

  function parse(formula) {
    var tokens = tokenize(formula);
    var ast = new Parser(tokens).parse();
    computeSpans(ast);
    return { tokens: tokens, ast: ast };
  }

  global.MacTraceParser = {
    tokenize: tokenize,
    parse: parse,
    collectRefs: collectRefs,
    collectOperands: collectOperands,
    childrenOf: childrenOf,
    nodeToText: nodeToText,
    argLabel: argLabel
  };
})(window);
