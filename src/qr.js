/* 最小 QR 產生器（位元組模式，容錯等級 M，版本 1–40）。演算法依 ISO/IEC 18004，參考 Nayuki QR Code generator（MIT）的結構。
   qrMatrix(text) → 二維布林陣列；qrSvg(text, px) → SVG 字串（含 4 格白邊）。 */
const QR = (() => {
  const ECC = [-1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28];
  const BLK = [-1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49];
  const FORMAT_M = 0;
  const rawModules = v => { let r = (16 * v + 128) * v + 64; if (v >= 2) { const n = Math.floor(v / 7) + 2; r -= (25 * n - 10) * n - 55; if (v >= 7) r -= 36; } return r; };
  const dataCodewords = v => Math.floor(rawModules(v) / 8) - ECC[v] * BLK[v];
  const mul = (x, y) => { let z = 0; for (let i = 7; i >= 0; i--) { z = (z << 1) ^ ((z >>> 7) * 0x11D); z ^= ((y >>> i) & 1) * x; } return z & 0xFF; };
  function rsDivisor(deg) { const r = new Array(deg).fill(0); r[deg - 1] = 1; let root = 1; for (let i = 0; i < deg; i++) { for (let j = 0; j < r.length; j++) { r[j] = mul(r[j], root); if (j + 1 < r.length) r[j] ^= r[j + 1]; } root = mul(root, 0x02); } return r; }
  function rsRemainder(data, div) { const r = div.map(() => 0); for (const b of data) { const f = b ^ r.shift(); r.push(0); div.forEach((c, i) => { r[i] ^= mul(c, f); }); } return r; }
  function matrix(text) {
    const bytes = Array.from(new TextEncoder().encode(text));
    let ver = 1, bits;
    for (; ver <= 40; ver++) { const ccBits = ver < 10 ? 8 : 16; if (4 + ccBits + bytes.length * 8 <= dataCodewords(ver) * 8) break; }
    if (ver > 40) throw new Error('QR 內容太長');
    bits = []; const put = (val, n) => { for (let i = n - 1; i >= 0; i--) bits.push((val >>> i) & 1); };
    put(4, 4); put(bytes.length, ver < 10 ? 8 : 16); bytes.forEach(b => put(b, 8));
    const cap = dataCodewords(ver) * 8;
    put(0, Math.min(4, cap - bits.length)); put(0, (8 - bits.length % 8) % 8);
    for (let pad = 0xEC; bits.length < cap; pad ^= 0xEC ^ 0x11) put(pad, 8);
    const data = []; for (let i = 0; i < bits.length; i += 8) data.push(bits.slice(i, i + 8).reduce((a, b) => (a << 1) | b, 0));
    // 分塊＋錯誤更正＋交錯
    const nb = BLK[ver], el = ECC[ver], raw = Math.floor(rawModules(ver) / 8), nShort = nb - raw % nb, shortLen = Math.floor(raw / nb), div = rsDivisor(el), blocks = [];
    for (let i = 0, k = 0; i < nb; i++) { const dat = data.slice(k, k + shortLen - el + (i < nShort ? 0 : 1)); k += dat.length; const ecc = rsRemainder(dat, div); if (i < nShort) dat.push(0); blocks.push(dat.concat(ecc)); }
    const cw = []; for (let i = 0; i < blocks[0].length; i++) blocks.forEach((b, j) => { if (i !== shortLen - el || j >= nShort) cw.push(b[i]); });
    // 畫面
    const size = ver * 4 + 17, mod = [...Array(size)].map(() => Array(size).fill(false)), fn = [...Array(size)].map(() => Array(size).fill(false));
    const setF = (x, y, d) => { mod[y][x] = d; fn[y][x] = true; };
    for (let i = 0; i < size; i++) { setF(6, i, i % 2 === 0); setF(i, 6, i % 2 === 0); }
    for (const [cx, cy] of [[3, 3], [size - 4, 3], [3, size - 4]]) for (let dy = -4; dy <= 4; dy++) for (let dx = -4; dx <= 4; dx++) { const d = Math.max(Math.abs(dx), Math.abs(dy)), x = cx + dx, y = cy + dy; if (x >= 0 && x < size && y >= 0 && y < size) setF(x, y, d !== 2 && d !== 4); }
    const ap = (() => { if (ver === 1) return []; const n = Math.floor(ver / 7) + 2, step = ver === 32 ? 26 : Math.ceil((ver * 4 + 4) / (n * 2 - 2)) * 2, r = [6]; for (let p = size - 7; r.length < n; p -= step) r.splice(1, 0, p); return r; })();
    ap.forEach((a, i) => ap.forEach((b, j) => { if ((i === 0 && j === 0) || (i === 0 && j === ap.length - 1) || (i === ap.length - 1 && j === 0)) return; for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) setF(a + dx, b + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1); }));
    const fmt = mask => { const d = FORMAT_M << 3 | mask; let r = d; for (let i = 0; i < 10; i++) r = (r << 1) ^ ((r >>> 9) * 0x537); const b = (d << 10 | r) ^ 0x5412, g = i => ((b >>> i) & 1) === 1;
      for (let i = 0; i <= 5; i++) setF(8, i, g(i)); setF(8, 7, g(6)); setF(8, 8, g(7)); setF(7, 8, g(8)); for (let i = 9; i < 15; i++) setF(14 - i, 8, g(i));
      for (let i = 0; i < 8; i++) setF(size - 1 - i, 8, g(i)); for (let i = 8; i < 15; i++) setF(8, size - 15 + i, g(i)); setF(8, size - 8, true); };
    fmt(0);
    if (ver >= 7) { let r = ver; for (let i = 0; i < 12; i++) r = (r << 1) ^ ((r >>> 11) * 0x1F25); const b = ver << 12 | r; for (let i = 0; i < 18; i++) { const bit = ((b >>> i) & 1) === 1, a = size - 11 + i % 3, c = Math.floor(i / 3); setF(a, c, bit); setF(c, a, bit); } }
    let bi = 0;
    for (let right = size - 1; right >= 1; right -= 2) { if (right === 6) right = 5; for (let vert = 0; vert < size; vert++) for (let j = 0; j < 2; j++) { const x = right - j, up = ((right + 1) & 2) === 0, y = up ? size - 1 - vert : vert; if (!fn[y][x] && bi < cw.length * 8) { mod[y][x] = ((cw[bi >>> 3] >>> (7 - (bi & 7))) & 1) === 1; bi++; } } }
    const maskFn = [(x, y) => (x + y) % 2 === 0, (x, y) => y % 2 === 0, x => x % 3 === 0, (x, y) => (x + y) % 3 === 0, (x, y) => (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0, (x, y) => x * y % 2 + x * y % 3 === 0, (x, y) => (x * y % 2 + x * y % 3) % 2 === 0, (x, y) => ((x + y) % 2 + x * y % 3) % 2 === 0];
    const apply = m => { for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) if (!fn[y][x] && maskFn[m](x, y)) mod[y][x] = !mod[y][x]; };
    const penalty = () => { let p = 0, dark = 0; const line = get => { let s = ''; for (let i = 0; i < size; i++) s += get(i) ? '1' : '0'; return s; };
      const lines = []; for (let i = 0; i < size; i++) { lines.push(line(k => mod[i][k])); lines.push(line(k => mod[k][i])); }
      for (const l of lines) { for (const run of l.match(/0{5,}|1{5,}/g) || []) p += 3 + run.length - 5; p += 40 * ((('0000' + l + '0000').match(/(?=00001011101|10111010000)/g) || []).length); }
      for (let y = 0; y < size - 1; y++) for (let x = 0; x < size - 1; x++) { const c = mod[y][x]; if (c === mod[y][x + 1] && c === mod[y + 1][x] && c === mod[y + 1][x + 1]) p += 3; }
      mod.forEach(r => r.forEach(c => { if (c) dark++; })); const total = size * size; p += 10 * Math.floor(Math.abs(dark * 20 - total * 10) / total); return p; };
    let best = 0, min = Infinity;
    for (let m = 0; m < 8; m++) { apply(m); fmt(m); const pp = penalty(); if (pp < min) { min = pp; best = m; } apply(m); }
    apply(best); fmt(best);
    return mod;
  }
  function svg(text, px = 8) { const m = matrix(text), n = m.length, q = 4, w = (n + q * 2) * px; let d = ''; m.forEach((r, y) => r.forEach((c, x) => { if (c) d += `M${(x + q) * px} ${(y + q) * px}h${px}v${px}h-${px}z`; })); return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${w}" width="${w}" height="${w}" shape-rendering="crispEdges"><rect width="${w}" height="${w}" fill="#fff"/><path d="${d}" fill="#000"/></svg>`; }
  return { matrix, svg };
})();
