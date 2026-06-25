/*
 * Generates MacTrace's PNG icons with zero dependencies (Node's built-in zlib).
 * Produces a "magnifying glass" Trace icon and a "grid" Formula Map icon at the
 * sizes the manifest references. Run with:  node scripts/make-icons.js
 */
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

const ASSETS = path.join(__dirname, "..", "assets");

// ---- Tiny RGBA canvas -------------------------------------------------------

function canvas(n) {
  return { n: n, data: new Uint8Array(n * n * 4) };
}
function setPx(c, x, y, r, g, b, a) {
  if (x < 0 || y < 0 || x >= c.n || y >= c.n) return;
  const i = (y * c.n + x) * 4;
  const ia = a / 255;
  c.data[i] = Math.round(c.data[i] * (1 - ia) + r * ia);
  c.data[i + 1] = Math.round(c.data[i + 1] * (1 - ia) + g * ia);
  c.data[i + 2] = Math.round(c.data[i + 2] * (1 - ia) + b * ia);
  c.data[i + 3] = Math.max(c.data[i + 3], a);
}
function hex(h) {
  return [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
}
function roundedBg(c, color, radiusFrac) {
  const n = c.n;
  const r = n * radiusFrac;
  const [R, G, B] = hex(color);
  for (let y = 0; y < n; y++) {
    for (let x = 0; x < n; x++) {
      // Skip pixels outside the rounded corners.
      const cx = Math.min(x + 0.5, n - x - 0.5);
      const cy = Math.min(y + 0.5, n - y - 0.5);
      if (cx < r && cy < r) {
        const dx = r - cx, dy = r - cy;
        if (dx * dx + dy * dy > r * r) continue;
      }
      setPx(c, x, y, R, G, B, 255);
    }
  }
}
function disc(c, cx, cy, rad, color, fill) {
  const [R, G, B] = hex(color);
  const inner = fill ? 0 : rad - Math.max(1.4, c.n * 0.07);
  for (let y = Math.floor(cy - rad - 1); y <= cy + rad + 1; y++) {
    for (let x = Math.floor(cx - rad - 1); x <= cx + rad + 1; x++) {
      const d = Math.hypot(x + 0.5 - cx, y + 0.5 - cy);
      if (d <= rad && d >= inner) setPx(c, x, y, R, G, B, 255);
    }
  }
}
function thickLine(c, x0, y0, x1, y1, w, color) {
  const [R, G, B] = hex(color);
  const steps = Math.ceil(Math.hypot(x1 - x0, y1 - y0)) * 2;
  for (let s = 0; s <= steps; s++) {
    const t = s / steps;
    const px = x0 + (x1 - x0) * t;
    const py = y0 + (y1 - y0) * t;
    disc(c, px, py, w / 2, color, true);
  }
}
function fillRect(c, x0, y0, w, h, color) {
  const [R, G, B] = hex(color);
  for (let y = Math.round(y0); y < Math.round(y0 + h); y++) {
    for (let x = Math.round(x0); x < Math.round(x0 + w); x++) setPx(c, x, y, R, G, B, 255);
  }
}

// ---- PNG encoder ------------------------------------------------------------

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, "ascii");
  const body = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([len, body, crc]);
}
function encodePng(c) {
  const n = c.n;
  const sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(n, 0);
  ihdr.writeUInt32BE(n, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  // raw: filter byte + row bytes
  const raw = Buffer.alloc(n * (n * 4 + 1));
  for (let y = 0; y < n; y++) {
    raw[y * (n * 4 + 1)] = 0;
    c.data.subarray(y * n * 4, (y + 1) * n * 4).forEach((v, i) => {
      raw[y * (n * 4 + 1) + 1 + i] = v;
    });
  }
  const idat = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([
    sig,
    chunk("IHDR", ihdr),
    chunk("IDAT", idat),
    chunk("IEND", Buffer.alloc(0))
  ]);
}

// ---- Icon artwork -----------------------------------------------------------

const ACCENT = "#217346";
const WHITE = "#ffffff";
const MAP = ["#b6d7a8", "#cfe2f3", "#f9cb9c", "#efefef"];

function traceIcon(n) {
  const c = canvas(n);
  roundedBg(c, ACCENT, 0.2);
  const cx = n * 0.43, cy = n * 0.43, rad = n * 0.22;
  disc(c, cx, cy, rad, WHITE, false); // ring
  const ex = cx + rad * 0.72, ey = cy + rad * 0.72;
  thickLine(c, ex, ey, n * 0.78, n * 0.78, Math.max(2, n * 0.1), WHITE); // handle
  return c;
}
function mapIcon(n) {
  const c = canvas(n);
  roundedBg(c, ACCENT, 0.2);
  const pad = n * 0.18, gap = n * 0.06;
  const cell = (n - pad * 2 - gap) / 2;
  const pos = [
    [pad, pad], [pad + cell + gap, pad],
    [pad, pad + cell + gap], [pad + cell + gap, pad + cell + gap]
  ];
  pos.forEach((p, i) => fillRect(c, p[0], p[1], cell, cell, MAP[i]));
  return c;
}

// ---- Emit -------------------------------------------------------------------

if (!fs.existsSync(ASSETS)) fs.mkdirSync(ASSETS, { recursive: true });
[16, 32, 64, 80, 128].forEach((n) =>
  fs.writeFileSync(path.join(ASSETS, "icon-" + n + ".png"), encodePng(traceIcon(n)))
);
[16, 32, 80].forEach((n) =>
  fs.writeFileSync(path.join(ASSETS, "map-" + n + ".png"), encodePng(mapIcon(n)))
);
console.log("Wrote icons to " + ASSETS);
