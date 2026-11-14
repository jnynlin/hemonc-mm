/* 血腫科 M&M — SharePoint 站台內工具（收案＋開會）  v1 2026-11-14
   由書籤載入：書籤從本站「文件」庫讀 mm_app.js 後執行。資料只在瀏覽器 ↔ 本站 SharePoint 之間，不經其他伺服器。
   本檔由 build_app.py 產生：轉換核心取自 mm_convert.html、儀表板取自 mm_dashboard.html，請勿手改。 */
(() => {
  'use strict';
  if (window.__MM_APP) { window.__MM_APP.show(); return; }
  const VERSION = 'v1 2026-11-14';
  const base = location.origin + ((location.pathname.match(/^\/(sites|teams)\/[^/]+/) || [''])[0]);

  /* ================= SharePoint REST ================= */
  let digest = null, digestAt = 0;
  async function getDigest() {
    if (digest && Date.now() - digestAt < 15 * 60e3) return digest;
    const r = await fetch(base + '/_api/contextinfo', { method: 'POST', headers: { Accept: 'application/json;odata=verbose' } });
    if (!r.ok) throw new Error('拿不到寫入權杖（' + r.status + '），請重新整理頁面再試');
    digest = (await r.json()).d.GetContextWebInformation.FormDigestValue; digestAt = Date.now(); return digest;
  }
  async function rq(path, { method = 'GET', body = null, headers = {} } = {}) {
    const url = /^https?:/.test(path) ? path : base + '/_api/' + path;
    const h = { Accept: method === 'GET' ? 'application/json;odata=nometadata' : 'application/json;odata=verbose', ...headers };
    if (method !== 'GET') { h['X-RequestDigest'] = await getDigest(); if (body) h['Content-Type'] = 'application/json;odata=verbose'; }
    const r = await fetch(url, { method, headers: h, body: body ? JSON.stringify(body) : null });
    if (r.status === 204) return null;
    const t = await r.text(); let j = null; try { j = t ? JSON.parse(t) : null; } catch (_) { j = { raw: t }; }
    if (!r.ok) { const m = (j && ((j['odata.error'] && j['odata.error'].message.value) || (j.error && j.error.message && j.error.message.value))) || t.slice(0, 200); const e = new Error(m); e.status = r.status; throw e; }
    return j && j.d !== undefined ? j.d : j;
  }
  const q = s => encodeURIComponent(String(s).replace(/'/g, "''"));
  const LI = {};
  async function listInfo(title) {
    if (LI[title]) return LI[title];
    const l = await rq(`web/lists/getbytitle('${q(title)}')?$select=Id,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
    const fs = (await rq(`web/lists(guid'${l.Id}')/fields?$filter=Hidden eq false`)).value;
    const ours = fs.filter(f => f.InternalName === 'Title' || /^mm/.test(f.InternalName));
    const f = {}; for (const x of ours) f[x.Title] = x; const tf = ours.find(x => x.InternalName === 'Title'); if (tf) f['標題'] = tf;
    return (LI[title] = { title, id: l.Id, type: l.ListItemEntityTypeFullName, url: location.origin + l.RootFolder.ServerRelativeUrl, f, ours });
  }
  async function canAdd(title) {
    try { const li = await listInfo(title); const p = await rq(`web/lists(guid'${li.id}')/EffectiveBasePermissions`); return (Number(p.Low) & 2) !== 0; }
    catch (_) { return false; }
  }
  async function readAll(title) {
    const li = await listInfo(title); const sel = ['Id'], exp = [];
    for (const f of li.ours) { if (/^Lookup/.test(f.TypeAsString)) { sel.push(f.InternalName + '/Title', f.InternalName + '/Id'); exp.push(f.InternalName); } else sel.push(f.EntityPropertyName); }
    let url = `web/lists(guid'${li.id}')/items?$top=5000&$select=${sel.join(',')}${exp.length ? '&$expand=' + exp.join(',') : ''}`, out = [];
    while (url) { const j = await rq(url); out = out.concat(j.value || []); url = j['odata.nextLink'] || null; }
    return out;
  }
  const ymd = v => { const d = new Date(v); if (isNaN(d)) return ''; return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  function toDisplay(li, it) {
    const o = { _id: it.Id };
    for (const f of li.ours) {
      const t = f.InternalName === 'Title' ? '標題' : f.Title, ty = f.TypeAsString;
      let v = /^Lookup/.test(ty) ? it[f.InternalName] : it[f.EntityPropertyName];
      if (ty === 'Lookup') v = v && v.Title ? v.Title : '';
      else if (ty === 'LookupMulti') v = (v || []).map(x => x.Title).join(';');
      else if (ty === 'MultiChoice') v = (v || []).join(';');
      else if (ty === 'DateTime') v = v ? ymd(v) : '';
      else if (ty === 'Boolean') v = v ? '是' : '否';
      else v = v == null ? '' : String(v);
      o[t] = v;
    }
    return o;
  }
  function toBody(li, row, lookupIds) { // row: 顯示名稱 → 字串
    const body = { __metadata: { type: li.type } }, warn = [];
    const split = v => String(v).split(/;#|;|；/).map(s => s.trim()).filter(Boolean);
    for (const [h, raw] of Object.entries(row)) {
      const f = li.f[h]; if (!f || f.ReadOnlyField || raw == null || String(raw).trim() === '') continue;
      const v = String(raw).trim(), k = f.EntityPropertyName, ch = f.Choices || [];
      switch (f.TypeAsString) {
        case 'Number': { const n = parseFloat(v.replace(/[^\d.-]/g, '')); if (isNaN(n)) warn.push(`${h}「${v}」不是數字`); else body[k] = n; break; }
        case 'Choice': if (ch.length && !ch.includes(v)) warn.push(`${h}「${v}」不在選項內`); else body[k] = v; break;
        case 'MultiChoice': { const vals = split(v).filter(x => !ch.length || ch.includes(x)); if (vals.length) body[k] = { __metadata: { type: 'Collection(Edm.String)' }, results: vals }; break; }
        case 'DateTime': { const d = new Date(v.replace(/[./]/g, '-')); if (!isNaN(d)) body[k] = ymd(d) + 'T00:00:00Z'; break; }
        case 'Lookup': { const id = lookupIds && lookupIds[v]; if (id) body[k + 'Id'] = id; else warn.push(`${h}「${v}」對不到根因分類`); break; }
        case 'LookupMulti': { const ids = split(v).map(x => lookupIds && lookupIds[x]).filter(Boolean); if (ids.length) body[k + 'Id'] = { __metadata: { type: 'Collection(Edm.Int32)' }, results: ids }; break; }
        default: body[k] = v;
      }
    }
    return { body, warn };
  }
  const addItem = (li, body) => rq(`web/lists(guid'${li.id}')/items`, { method: 'POST', body });
  const delItem = (li, id) => rq(`web/lists(guid'${li.id}')/items(${id})`, { method: 'POST', headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' } });
  const mergeItem = (li, id, body) => rq(`web/lists(guid'${li.id}')/items(${id})`, { method: 'POST', body, headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' } });

  /* ================= 轉換核心（mm_convert.html） ================= */
  /*CORE-START*/
const MM = (() => {
const LIST_COLS = ['標題','年月','案例類型','年齡層','性別','診斷群','治療階段','詳細診斷','治療內容','臨床經過','住院天數','可避免天數','死亡預期性','Morbidity等級',
  '初評主要根因','初評促成因素','初評可預防性','主治醫師回覆','主治認為主要根因','主治認為可預防性','主治醫師說明','五個為什麼','學習重點','參考指引文獻','狀態','病人簡稱','年齡'];
const KEY_COLS = ['標題','病歷號','姓名','床號','主治醫師','原主治','住院日期','死亡日期','住院日數原文'];
const DX_RULES = [
  [/\bAML\b|acute myelo|acute promyelocytic|\bAPL\b|\bALL\b|acute lymphoblastic|plasma cell leukemia/i, '急性白血病'],
  [/\bCML\b|\bCLL\b|\bCMML\b|chronic (myeloid|lymphocytic|myelomonocytic)/i, '慢性白血病'],
  [/lymphoma|\bDLBCL\b|hodgkin|\bMEITL\b/i, '淋巴瘤'],
  [/myeloma|plasmacytoma/i, '骨髓瘤'],
  [/\bMDS\b|myelodysplastic|myeloproliferative|myelofibrosis|\bMPN\b/i, 'MDS／MPN'],
  [/\blung\b|\bNSCLC\b|\bSCLC\b|thymoma/i, '肺癌'],
  [/breast/i, '乳癌'],
  [/esophag|\bESCC\b|\bEGJ\b/i, '食道癌'],
  [/\bgastric\b|stomach/i, '胃癌'],
  [/\bcolon|rectal|rectum|cecum|colorectal|sigmoid/i, '大腸直腸癌'],
  [/hepatocellular|\bHCC\b|cholangio|pancrea|biliary|gallbladder/i, '肝膽胰癌'],
  [/\bgum\b|\boral\b|tongue|buccal|nasopharyn|oropharyn|hypopharyn|laryn|head and neck|\bHNSCC\b|neck squamous/i, '頭頸癌'],
  [/prostate|bladder|urothelial|renal cell|\bRCC\b|kidney/i, '泌尿道／攝護腺癌'],
  [/ovar|cervical|endometri|uterine/i, '婦科癌'],
  [/sarcoma/i, '肉瘤'],
  [/melanoma/i, '其他固態瘤'],
];

// ---- zip 讀取（只支援 stored／deflate；用瀏覽器內建 DecompressionStream）----
async function zipRead(buf, wantName) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf);
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) throw new Error('不是有效的 .docx（找不到 zip 目錄）');
  const n = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
  const td = new TextDecoder();
  for (let k = 0; k < n; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('zip 目錄損壞');
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
    const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
    if (name === wantName) {
      const lnlen = dv.getUint16(off + 26, true), lelen = dv.getUint16(off + 28, true);
      const start = off + 30 + lnlen + lelen, data = u8.slice(start, start + csize);
      if (method === 0) return td.decode(data);
      if (method !== 8) throw new Error('不支援的壓縮方式 ' + method);
      const ds = new DecompressionStream('deflate-raw');
      const out = await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer();
      return td.decode(out);
    }
    p += 46 + nlen + elen + clen;
  }
  throw new Error('docx 裡沒有 ' + wantName);
}

// ---- zip 寫入（stored，不壓縮；xlsx 很小）----
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zipWrite(files) { // files: [[name, string]]
  const te = new TextEncoder(), parts = [], cd = []; let off = 0;
  const le16 = v => [v & 255, (v >> 8) & 255], le32 = v => [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255];
  for (const [name, content] of files) {
    const nm = te.encode(name), data = te.encode(content), crc = crc32(data);
    const hdr = new Uint8Array([...le32(0x04034b50), ...le16(20), ...le16(0x0800), ...le16(0), ...le16(0), ...le16(0x21), ...le32(crc), ...le32(data.length), ...le32(data.length), ...le16(nm.length), ...le16(0)]);
    parts.push(hdr, nm, data);
    cd.push(new Uint8Array([...le32(0x02014b50), ...le16(20), ...le16(20), ...le16(0x0800), ...le16(0), ...le16(0), ...le16(0x21), ...le32(crc), ...le32(data.length), ...le32(data.length), ...le16(nm.length), ...le16(0), ...le16(0), ...le16(0), ...le16(0), ...le32(0), ...le32(off)]), nm);
    off += hdr.length + nm.length + data.length;
  }
  const cdLen = cd.reduce((a, b) => a + b.length, 0);
  const eocd = new Uint8Array([...le32(0x06054b50), ...le16(0), ...le16(0), ...le16(files.length), ...le16(files.length), ...le32(cdLen), ...le32(off), ...le16(0)]);
  return new Blob([...parts, ...cd, eocd], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
const xesc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function colName(i) { let r = ''; for (;;) { r = String.fromCharCode(65 + i % 26) + r; i = Math.floor(i / 26) - 1; if (i < 0) return r; } }
function writeXlsx(header, rows) {
  const cell = (ri, ci, v) => { const ref = colName(ci) + ri; if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`; if (v == null || v === '') return ''; return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xesc(v)}</t></is></c>`; };
  const data = [header, ...rows].map((r, i) => `<row r="${i + 1}">` + r.map((v, ci) => cell(i + 1, ci, v)).join('') + '</row>').join('');
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${data}</sheetData></worksheet>`;
  return zipWrite([
    ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'],
    ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="案例" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
    ['xl/worksheets/sheet1.xml', sheet],
  ]);
}

// ---- Word XML：第一個表格 → 列 → 兩欄 → 段落文字 ----
const unx = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');
function paraTexts(xml) { // 段落 → 文字（空段落略過）
  const out = []; const re = /<w:p(?:\s[^>]*[^\/])?>([\s\S]*?)<\/w:p>/g; let m;
  while ((m = re.exec(xml))) { const t = [...m[1].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(x => unx(x[1])).join('').trim(); if (t) out.push(t); }
  return out;
}
function firstTable(xml) {
  const s = xml.indexOf('<w:tbl>'); if (s < 0) throw new Error('文件裡沒有表格');
  let depth = 0, i = s; const re = /<w:tbl>|<\/w:tbl>/g; re.lastIndex = s; let m;
  while ((m = re.exec(xml))) { depth += m[0] === '<w:tbl>' ? 1 : -1; if (depth === 0) return xml.slice(s, m.index + 8); }
  throw new Error('表格未閉合');
}
function tableRows(tbl) { return [...tbl.matchAll(/<w:tr(?:\s[^>]*)?>([\s\S]*?)<\/w:tr>/g)].map(r => [...r[1].matchAll(/<w:tc(?:\s[^>]*)?>([\s\S]*?)<\/w:tc>/g)].map(c => paraTexts(c[1]))); }

// ---- 與 mm_import.py 相同的解析邏輯 ----
function parseLeft(lines) {
  let joined = lines.join(' / ').replace(/y\/o/g, 'yo');
  const prev = [...joined.matchAll(/[（(]\s*原\s*([一-鿿]{2,4})\s*[）)]/g)].map(m => m[1]);
  joined = joined.replace(/[（(]\s*原[^）)]*[）)]/g, ' / ').replace(/VS\.?/g, ' / ');
  const toks = joined.split(/\s*\/\s*/).map(t => t.trim()).filter(Boolean);
  const flat = [];
  for (const t of toks) { const m = t.match(/^([一-鿿]{2,4})\s*(\d{6,})$/) || t.match(/^(\d{6,})\s*([一-鿿]{2,4})$/); if (m) flat.push(m[1], m[2]); else flat.push(t); }
  const bed = flat.find(t => /^[A-Z]\d{1,2}-\d{1,3}$/.test(t)) || '';
  const chart = flat.find(t => /^\d{6,}$/.test(t)) || '';
  const agesex = flat.find(t => /^\d{1,3}\s*(yo)?\s*[MF]$/i.test(t)) || '';
  const names = flat.filter(t => /^[一-鿿]{2,4}$/.test(t));
  const m = agesex.match(/^(\d{1,3})\s*(?:yo)?\s*([MF])$/i);
  return { bed, chart, name: names[0] || '', attending: names.slice(1).join('／'), prevAttending: prev.join('／'), age: m ? +m[1] : null, sex: m ? ({ M: '男', F: '女' })[m[2].toUpperCase()] : '' };
}
const ageBand = a => a == null ? '' : a < 40 ? '<40' : a < 50 ? '40–49' : a < 60 ? '50–59' : a < 70 ? '60–69' : a < 80 ? '70–79' : '≥80';
function parseDate(s) {
  s = s.trim(); let m = s.match(/^(\d{4})(\d{2})(\d{2})$/) || s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(1[01]\d)(\d{2})(\d{2})$/) || s.match(/^(1[01]\d)[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (m) return Date.UTC(+m[1] + 1911, +m[2] - 1, +m[3]);
  return null;
}
const ymOf = ts => { const d = new Date(ts); return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'); };
function parseRight(lines, kind) {
  const text = lines.join('\n'); const out = { admission: '', death: '', losRaw: '', los: null, ym: null }; let dx;
  if (kind === 'mortality') {
    const first = lines[0] || '';
    const m = first.match(/^\s*([0-9./-]+?)\s*[－~～至]\s*([0-9./-]+)/) || first.match(/^\s*(\d{7,8})\s*-\s*(\d{7,8})/);
    const parts = m ? [m[1], m[2]] : [first];
    const d1 = parts.length >= 2 ? parseDate(parts[0]) : null, d2 = parts.length >= 2 ? parseDate(parts[1]) : null;
    out.admission = parts[0] || ''; out.death = parts[1] || '';
    if (d1 != null && d2 != null) out.los = Math.round((d2 - d1) / 86400000) + 1;
    if (d2 != null) out.ym = ymOf(d2);
    dx = (d1 != null || d2 != null) ? lines.slice(1).join('\n') : text;
  } else {
    const adm = text.match(/住院日期[:：]\s*([^\n]+)/), losM = text.match(/住院日數[:：]\s*([^\n]+)/);
    out.admission = adm ? adm[1].trim() : ''; out.losRaw = losM ? losM[1].trim() : '';
    const nums = (out.losRaw.split(/[(（]/)[0].match(/\d+/g) || []).map(Number);
    out.los = nums.length ? nums.reduce((a, b) => a + b, 0) : null;
    dx = text.includes('診斷') ? text.split('診斷').pop() : text;
  }
  out.dx = dx.replace(/^[:：]\s*/, '').trim();
  return out;
}
function nameShort(n) { n = (n || '').trim(); if (!n) return ''; const two = ['歐陽','司徒','司馬','諸葛','上官','端木','夏侯','皇甫','尉遲','公孫','長孫','令狐','張簡','范姜']; return (two.includes(n.slice(0, 2)) ? n.slice(0, 2) : n.slice(0, 1)) + 'OO'; }
// 研究醫師的「診斷」其實是問題清單：第 1 行＝主診斷＋診療史，第 2 行起＝其他診斷／併發症。
// 回傳 [詳細診斷, 治療內容, 臨床經過]：主診斷＝第 1 行切在第一個診療史子句之前（255 字、≥12 字）；
// 治療內容＝第 1 行其餘；臨床經過＝第 2 行起以「｜」相接（格線貼上不能有換行）。與 mm_import.py 同步。
const STRONG = '(?:s/p|status post|currently|underwent|received|receiving|treated|diagnosed|接受|目前|曾|使用|化療|治療|已完成)';
const WEAK = '(?:now |under |post |after |dx |on [A-Za-z-]+(?:mab|nib|platin|taxel|cycle)|with (?:relapse|progression|recurren|refractory|PD\\b))';
const TX_CUT = new RegExp('(?:[,;，；]\\s*|\\s+)(?=' + STRONG + ')|[,;，；]\\s*(?=' + WEAK + ')', 'i');
function splitDx(t, limit = 250) {
  const lines = t.split('\n').map(l => l.trim()).filter(Boolean);
  if (!lines.length) return ['', '', ''];
  const first = lines[0], m = TX_CUT.exec(first);
  let dx = first, tx = '';
  if (m) { dx = first.slice(0, m.index).trim(); tx = first.slice(m.index + m[0].length).trim(); }
  if (dx.length < 12) { dx = first; tx = ''; }
  if (dx.length > limit) {
    const cut = Math.max(...[...',;，；. '].map(c => dx.lastIndexOf(c, limit - 1)));
    dx = (cut >= 12 ? dx.slice(0, cut) : dx.slice(0, limit)).replace(/[,;，； .]+$/, '');
  }
  dx = dx.replace(/[,;，； ]+$/, '');
  return [dx, tx, lines.slice(1).join('｜')];
}
function deidentify(s) {
  return s.replace(/\b(20\d{2})(0[1-9]|1[0-2])([0-3]\d)\b/g, '$1/$2')
    .replace(/(20\d{2})[/\-.](\d{1,2})[/\-.](\d{1,2})/g, '$1/$2')
    .replace(/\b(1[01]\d)[/\-.](\d{1,2})[/\-.](\d{1,2})\b/g, '$1/$2')
    .replace(/\b(1[01]\d)(0[1-9]|1[0-2])([0-3]\d)\b/g, '$1/$2')
    .replace(/\d{6,}/g, '[已移除]');
}
function dxGroup(dx) { let best = null; for (const [rx, g] of DX_RULES) { const m = dx.match(rx); if (m && (best === null || m.index < best[0])) best = [m.index, g]; } return best ? best[1] : ''; }

async function convert(buf, defaultYm) {
  const xml = await zipRead(buf, 'word/document.xml');
  const firstP = (xml.match(/<w:p(?:\s[^>]*[^\/])?>([\s\S]*?)<\/w:p>/) || [, ''])[1];
  const head = [...firstP.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(x => unx(x[1])).join('');
  const kind = /mortality|死亡/i.test(head) ? 'mortality' : /morbidity|併發/i.test(head) ? 'morbidity' : 'los';
  const ctype = { mortality: 'Mortality', morbidity: 'Morbidity', los: '超長住院' }[kind], prefix = { mortality: 'MOR', morbidity: 'MOB', los: 'LOS' }[kind];
  if (kind === 'los' && !defaultYm) throw new Error('超長住院清單需要指定討論年月（YYYY-MM）');
  const rows = tableRows(firstTable(xml)); const key = [], drafts = [], counters = {};
  for (const tcs of rows) {
    if (tcs.length !== 2) continue;
    const [left, right] = tcs;
    if (!left.length || /床號|姓名/.test(left[0])) continue;
    const L = parseLeft(left), R = parseRight(right, kind);
    if (!(L.chart || L.age != null)) continue;
    const ym = R.ym || defaultYm || '0000-00'; counters[ym] = (counters[ym] || 0) + 1;
    const code = `${prefix}-${ym}-${String(counters[ym]).padStart(2, '0')}`;
    key.push({ code, bed: L.bed, chart: L.chart, name: L.name, attending: L.attending, prevAttending: L.prevAttending, admission: R.admission, death: R.death, losRaw: R.losRaw });
    const full = deidentify(R.dx), [dxDetail, treatment, course] = splitDx(full);
    drafts.push({ code, ym, types: [ctype], ageBand: ageBand(L.age), age: L.age, nameShort: nameShort(L.name), sex: L.sex, dxGroup: dxGroup(full), dxDetail, treatment, course, los: R.los, deathExpected: kind === 'mortality' ? '' : '不適用', morbGrade: '不適用' });
  }
  const caseRows = drafts.map(d => [d.code, d.ym, d.types.join(';'), d.ageBand, d.sex, d.dxGroup, '', d.dxDetail, d.treatment, d.course, d.los ?? '', '', d.deathExpected, d.morbGrade, '', '', '', '未徵詢', '', '', '', '', '', '', '整理中', d.nameShort, d.age ?? '']);
  const keyRows = key.map(k => [k.code, k.chart, k.name, k.bed, k.attending, k.prevAttending, k.admission, k.death, k.losRaw]);
  const leak = drafts.filter(d => /\d{6,}/.test(JSON.stringify(d))).map(d => d.code);
  return { kind, ctype, head, drafts, key, leak, caseRows, keyRows, caseXlsx: () => writeXlsx(LIST_COLS, caseRows), keyXlsx: () => writeXlsx(KEY_COLS, keyRows) };
}
return { convert, LIST_COLS, KEY_COLS };
})();
/*CORE-END*/

  /* ================= 外框 ================= */
  const host = document.createElement('div'); host.id = 'mm-app-host';
  host.style.cssText = 'position:fixed;inset:0;z-index:2147483000;';
  const sh = host.attachShadow({ mode: 'open' });
  sh.innerHTML = `<style>
:host{all:initial}
.root{--bg:#fbfbf9;--fg:#1f2428;--muted:#5d6670;--line:#d9ddd6;--soft:#eef0ea;--accent:#1c5cab;--accent-soft:#e3ecf8;--ok:#2e7d4f;--bad:#b4322e;--warn:#a86a00;
  position:absolute;inset:0;display:flex;flex-direction:column;background:var(--bg);color:var(--fg);font:15px/1.6 system-ui,"Noto Sans TC","Microsoft JhengHei",sans-serif}
@media (prefers-color-scheme:dark){.root{--bg:#15181b;--fg:#e6e8e3;--muted:#9aa3ab;--line:#333a40;--soft:#1e2328;--accent:#7fb0ee;--accent-soft:#1d2a3c;--ok:#6fcf97;--bad:#ef8a86;--warn:#e6b35a;color-scheme:dark}}
.bar{display:flex;align-items:center;gap:6px;padding:6px 14px;border-bottom:1px solid var(--line);flex:none}
.bar b{margin-right:10px}.bar .sp{flex:1}.bar .v{color:var(--muted);font-size:12px;margin-right:6px}
.tab{font:inherit;background:none;border:0;border-bottom:2px solid transparent;padding:6px 12px;color:var(--muted);cursor:pointer}
.tab[aria-current=true]{color:var(--fg);border-bottom-color:var(--accent);font-weight:600}
button.b{font:inherit;font-size:14px;padding:5px 12px;border:1px solid var(--line);background:none;color:var(--fg);border-radius:6px;cursor:pointer}
button.p{background:var(--accent);border-color:var(--accent);color:#fff}button:disabled{opacity:.45;cursor:default}
.body{flex:1;min-height:0;position:relative}
.pane{position:absolute;inset:0;overflow:auto}.pane[hidden]{display:none}
iframe{border:0;width:100%;height:100%;display:block;background:var(--bg)}
.w{max-width:1000px;margin:0 auto;padding:18px 20px 60px}
h2{font-size:20px;margin:0 0 4px}h3{font-size:16px;margin:22px 0 6px}
.note{color:var(--muted);font-size:13.5px}.ok{color:var(--ok)}.bad{color:var(--bad)}.warn{color:var(--warn)}
.banner{padding:10px 14px;border-left:3px solid var(--bad);background:var(--soft);margin:12px 0}
.drop{border:2px dashed var(--line);padding:26px;text-align:center;cursor:pointer;margin:14px 0 8px;border-radius:8px}
.drop.over{border-color:var(--accent);background:var(--accent-soft)}.drop.has{border-style:solid;border-color:var(--ok);text-align:left;padding:12px 16px}
.row{display:flex;flex-wrap:wrap;gap:10px 18px;align-items:center;margin:8px 0}
input[type=text]{font:inherit;padding:5px 9px;border:1px solid var(--line);background:var(--bg);color:var(--fg);width:7.5em;border-radius:6px}
label.ck{display:inline-flex;gap:6px;align-items:center;cursor:pointer}
.tw{overflow-x:auto;margin-top:8px}table{border-collapse:collapse;width:100%;font-size:13.5px}
th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}th{color:var(--muted);font-weight:500;font-size:12px;white-space:nowrap}th small{display:block;font-weight:400}
td.c{white-space:nowrap;font-family:ui-monospace,Consolas,monospace;font-size:12.5px}td.x{max-width:240px}tr.dup td{color:var(--muted)}
pre.log{white-space:pre-wrap;font:12.5px/1.5 ui-monospace,Consolas,monospace;background:var(--soft);padding:10px 12px;border-radius:6px;max-height:40vh;overflow:auto}
.load{padding:60px;text-align:center;color:var(--muted)}
</style>
<div class="root"><div class="bar" id="bar"><b>M&amp;M</b><button class="tab" data-t="intake">收案</button><button class="tab" data-t="meet">開會</button><span class="sp"></span><span class="v">${VERSION}</span><button class="b" id="fs" title="全螢幕（投影用）">全螢幕</button><button class="b" id="close">關閉</button></div>
<div class="body"><div class="pane" id="p-intake"></div><div class="pane" id="p-meet" hidden><div class="load">讀取 SharePoint 清單中…</div></div></div></div>`;
  const $ = s => sh.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
  const clip = (t, n) => { const a = [...String(t || '')]; return esc(a.slice(0, n).join('')) + (a.length > n ? '…' : ''); };
  let tab = null;
  function show(t) {
    document.body.appendChild(host); host.style.display = '';
    t = t || tab || 'intake'; tab = t;
    sh.querySelectorAll('.tab').forEach(b => b.setAttribute('aria-current', b.dataset.t === t));
    $('#p-intake').hidden = t !== 'intake'; $('#p-meet').hidden = t !== 'meet';
    if (t === 'intake' && !intakeReady) intakeInit();
    if (t === 'meet' && !meetReady) meetInit();
    if (t === 'meet' && frame) frame.contentWindow.focus();
  }
  sh.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => show(b.dataset.t)));
  $('#close').addEventListener('click', () => { if (document.fullscreenElement) document.exitFullscreen(); host.style.display = 'none'; });
  $('#fs').addEventListener('click', () => { if (document.fullscreenElement) document.exitFullscreen(); else host.requestFullscreen().catch(() => {}); });

  /* ================= 收案 ================= */
  let intakeReady = false;
  async function intakeInit() {
    intakeReady = true; const P = $('#p-intake');
    P.innerHTML = '<div class="load">檢查權限中…</div>';
    const [okCase, okKey] = await Promise.all([canAdd('案例'), canAdd('對照表')]);
    if (!okCase || !okKey) {
      P.innerHTML = `<div class="w"><h2>收案</h2><div class="banner">${!okCase ? '你沒有新增「案例」清單項目的權限。' : '你看不到或不能寫入「對照表」清單（需在「M&amp;M 對照表」群組）。'}<br>案例與對照表必須一起寫入，所以這裡不會寫任何資料。請主任把你加進群組後，重新整理頁面再按書籤。</div></div>`;
      return;
    }
    P.innerHTML = `<div class="w"><h2>收案：Word → SharePoint</h2>
      <p class="note">把研究醫師的討論會 Word 檔拖進來 → 自動去識別、編碼、拆欄 → 確認後直接寫入「案例」與「對照表」。Word 檔不會上傳，只在這台電腦的瀏覽器解析。</p>
      <div class="drop" id="drop" tabindex="0"><b id="dt">把 .docx 拖到這裡，或點一下選檔</b><div class="note" id="ds">超長住院／Morbidity／Mortality 討論會清單</div><input type="file" id="file" accept=".docx" hidden></div>
      <div class="row"><span id="ymw">討論年月 <input type="text" id="ym" placeholder="2026-10"> <span class="note" id="ymh">超長住院／Morbidity 需要；檔名有「2026年10月」會自動帶入。</span></span>
        <label class="ck"><input type="checkbox" id="demo">演練模式（代碼加 TEST-，可一鍵刪除）</label></div>
      <div id="res"></div></div>`;
    let file = null, R = null;
    const drop = $('#drop'), fileIn = $('#file'), ymIn = $('#ym');
    const isMort = () => file && /mortality|死亡/i.test(file.name);
    async function go() {
      const res = $('#res');
      if (!file) return;
      $('#ymw').style.display = isMort() ? 'none' : '';
      const ym = ymIn.value.trim();
      if (!isMort() && !/^\d{4}-(0[1-9]|1[0-2])$/.test(ym)) { res.innerHTML = '<p class="warn">請填討論年月（YYYY-MM），填好會自動轉換。</p>'; return; }
      res.innerHTML = '<div class="load">轉換並比對 SharePoint 既有資料…</div>';
      try { R = await plan(file, isMort() ? null : ym, $('#demo').checked); render(); }
      catch (e) { res.innerHTML = `<p class="bad">失敗：${esc(e.message)}</p>`; }
    }
    function pick(f) {
      if (!f) return; if (!/\.docx$/i.test(f.name)) { $('#ds').innerHTML = '<span class="bad">請選 .docx</span>'; return; }
      file = f; drop.classList.add('has'); $('#dt').textContent = f.name; $('#ds').textContent = '（點這裡換檔）';
      const m = f.name.match(/(\d{4})年\s*(\d{1,2})月/); if (m) ymIn.value = m[1] + '-' + m[2].padStart(2, '0');
      go();
    }
    drop.addEventListener('click', () => fileIn.click());
    fileIn.addEventListener('change', () => pick(fileIn.files[0]));
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => pick(e.dataTransfer.files[0]));
    ymIn.addEventListener('change', go); $('#demo').addEventListener('change', go);

    function render() {
      const res = $('#res'), d = R.items, n = d.filter(x => !x.dup).length, dup = d.length - n;
      if (R.leak.length) { res.innerHTML = `<p class="bad">✗ 去識別未通過：仍含 6 碼以上數字（${esc(R.leak.join('、'))}）。不寫入，請回報主任。</p>`; return; }
      const unG = d.filter(x => !x.draft.dxGroup).length;
      res.innerHTML = `<h3><span class="ok">✓</span> ${esc(R.ctype)} ${d.length} 例，去識別通過${dup ? `　<span class="warn">其中 ${dup} 例已收過，預設不寫</span>` : ''}</h3>
        ${unG ? `<p class="note warn">▲ ${unG} 例診斷群沒自動歸到，寫入後在清單手選。</p>` : ''}
        ${R.demo ? '<p class="note">演練模式：代碼前加 TEST-。演練完按最下面「刪除演練資料」。</p>' : ''}
        <div class="row"><button class="b p" id="write">寫入勾選的 <span id="wn">${n}</span> 例</button><span class="note">每例寫「案例」＋「對照表」各一筆；對照表失敗會把該案例退回，不留半套。</span></div>
        <div class="tw"><table><tr><th></th><th>代碼</th><th>病人</th><th>診斷群</th><th>天數</th><th>主診斷<small>→ 詳細診斷</small></th><th>診療史<small>→ 治療內容</small></th><th>其他診斷・併發症<small>→ 臨床經過</small></th></tr>
        ${d.map((x, i) => `<tr class="${x.dup ? 'dup' : ''}"><td><input type="checkbox" data-i="${i}" ${x.dup ? 'disabled title="已收過，不重複寫入"' : 'checked'}></td><td class="c">${esc(x.code)}${x.dup ? `<br><span class="warn">已收過：${esc(x.dup)}</span>` : ''}</td><td>${esc(x.draft.nameShort)} ${esc(x.draft.age ?? '?')}${esc(x.draft.sex)}</td><td>${esc(x.draft.dxGroup) || '<span class="warn">未歸類</span>'}</td><td>${esc(x.draft.los ?? '')}</td><td class="x">${clip(x.draft.dxDetail, 46)}</td><td class="x">${x.draft.treatment ? clip(x.draft.treatment, 38) : '—'}</td><td class="x">${x.draft.course ? clip(x.draft.course.replace(/｜/g, ' ｜ '), 38) : '—'}</td></tr>`).join('')}</table></div>
        <pre class="log" id="log" hidden></pre>
        <h3>演練資料</h3><p class="note">刪除兩個清單裡所有代碼以 TEST- 開頭的項目。</p><button class="b" id="purge">刪除演練資料（TEST-）</button>`;
      const cnt = () => { $('#wn').textContent = res.querySelectorAll('input[data-i]:checked').length; };
      res.querySelectorAll('input[data-i]').forEach(c => c.addEventListener('change', cnt));
      $('#write').addEventListener('click', () => write([...res.querySelectorAll('input[data-i]:checked')].map(c => R.items[+c.dataset.i])));
      $('#purge').addEventListener('click', purge);
    }
    async function write(sel) {
      const log = $('#log'), btn = $('#write'); log.hidden = false; log.textContent = ''; btn.disabled = true;
      const say = s => { log.textContent += s + '\n'; log.scrollTop = log.scrollHeight; };
      const cl = await listInfo('案例'), kl = await listInfo('對照表'); let ok = 0, fail = 0;
      for (const x of sel) {
        let caseId = null;
        try {
          const c = toBody(cl, x.caseRow); const k = toBody(kl, x.keyRow);
          const cr = await addItem(cl, c.body); caseId = cr.Id;
          try { await addItem(kl, k.body); }
          catch (e) { await delItem(cl, caseId).catch(() => {}); throw new Error('對照表寫入失敗，已退回案例：' + e.message); }
          ok++; say(`${x.code} ✓${c.warn.length ? '（略過：' + c.warn.join('；') + '）' : ''}`);
        } catch (e) { fail++; say(`${x.code} ✗ ${e.message}`); }
      }
      say(`=== 完成：新增 ${ok}、失敗 ${fail} ===`);
      if (ok) say(`下一步：到「案例」清單逐案補初評、徵詢主治，完成的狀態改「待討論」。\n${cl.url}`);
      R = null; file = null; drop.classList.remove('has'); $('#dt').textContent = '把 .docx 拖到這裡，或點一下選檔'; $('#ds').textContent = '寫完了；要收下一個檔再拖進來';
      meetReady = false; // 開會頁下次重讀
    }
    async function purge() {
      const res = $('#res'); let log = $('#log'); if (!log) { res.insertAdjacentHTML('beforeend', '<pre class="log" id="log"></pre>'); log = $('#log'); }
      log.hidden = false; log.textContent = '刪除演練資料中…\n';
      let n = 0;
      for (const t of ['案例', '對照表']) {
        const li = await listInfo(t); const j = await rq(`web/lists(guid'${li.id}')/items?$top=5000&$select=Id,Title&$filter=startswith(Title,'TEST-')`);
        for (const it of j.value) { await delItem(li, it.Id); n++; }
      }
      log.textContent += `已刪除 ${n} 筆（兩個清單合計）。`; meetReady = false;
    }
  }
  // 轉換 → 依 SharePoint 既有資料重新編號、標出重複
  async function plan(file, ym, demo) {
    const r = await MM.convert(await file.arrayBuffer(), ym);
    const cl = await listInfo('案例'), kl = await listInfo('對照表'), chartF = kl.f['病歷號'];
    const E = t => cl.f[t].EntityPropertyName;
    const cases = (await rq(`web/lists(guid'${cl.id}')/items?$top=5000&$select=Title,${E('病人簡稱')},${E('年齡')},${E('詳細診斷')}`)).value;
    const titles = cases.map(x => x.Title || '');
    const fp = (stem, s, a, dx) => [stem.replace(/^TEST-/, ''), s || '', a ?? '', String(dx || '').slice(0, 40)].join('|');
    const byFp = {}; for (const c of cases) { const m = (c.Title || '').match(/^(.*-\d{4}-\d{2}-)\d+$/); if (m) byFp[fp(m[1], c[E('病人簡稱')], c[E('年齡')], c[E('詳細診斷')])] = c.Title; }
    const keys = (await rq(`web/lists(guid'${kl.id}')/items?$top=5000&$select=Title,${chartF.EntityPropertyName}`)).value;
    const byChart = {}; for (const k of keys) { const c = String(k[chartF.EntityPropertyName] || '').trim(); if (c) (byChart[c] = byChart[c] || []).push(k.Title || ''); }
    const pre = demo ? 'TEST-' : '', next = {};
    const items = r.drafts.map((d, i) => {
      const prefix = d.code.replace(/-\d{4}-\d{2}-\d+$/, ''), stem = `${pre}${prefix}-${d.ym}-`;
      const chart = String(r.keyRows[i][1] || '').trim();
      let dup = (byChart[chart] || []).find(t => t.startsWith(stem) && /^\d+$/.test(t.slice(stem.length))) || null;
      if (!dup) { const hit = byFp[fp(stem, d.nameShort, d.age, d.dxDetail)]; if (hit && hit.startsWith(stem)) dup = hit; } // 案例在、對照表沒有的舊資料
      return { d, i, stem, dup };
    });
    for (const x of items) {
      if (x.dup) { x.code = x.dup; continue; }
      if (next[x.stem] == null) next[x.stem] = Math.max(0, ...titles.filter(t => t.startsWith(x.stem)).map(t => parseInt(t.slice(x.stem.length), 10)).filter(n => !isNaN(n)));
      x.code = x.stem + String(++next[x.stem]).padStart(2, '0');
    }
    const obj = (cols, row) => { const o = {}; cols.forEach((c, j) => o[c] = row[j]); return o; };
    return {
      ctype: r.ctype, leak: r.leak, demo,
      items: items.map(x => {
        const caseRow = obj(MM.LIST_COLS, r.caseRows[x.i]), keyRow = obj(MM.KEY_COLS, r.keyRows[x.i]);
        caseRow['標題'] = x.code; keyRow['標題'] = x.code;
        return { code: x.code, dup: x.dup, draft: x.d, caseRow, keyRow };
      })
    };
  }

  /* ================= 開會（儀表板＋定案寫回） ================= */
  let meetReady = false, frame = null, caseRaw = {}, caseDisp = {}, facIds = {};
  const DASH_CSS = "\n/* 離線版：讀 SharePoint List 匯出的 CSV／Excel，不連任何資料庫。版面同案例庫：頂列導覽＋平面分區。 */\n:root{\n  --bg:#f9f9f7; --surface:#fcfcfb; --fg:#0b0b0b; --fg2:#52514e; --muted:#898781;\n  --rule:#e1e0d9; --axis:#c3c2b7; --accent:#1c5cab; --accent-soft:#e3edf9;\n  --s1:#2a78d6; --s2:#eb6834; --s3:#1baf7a; --good:#0ca30c; --critical:#d03b3b; --serious:#ec835a;\n  --seq1:#cde2fb; --seq2:#9ec5f4; --seq3:#5598e7; --seq4:#1c5cab;\n  --font:\"Noto Sans TC\",system-ui,-apple-system,\"Segoe UI\",\"PingFang TC\",\"Microsoft JhengHei\",sans-serif;\n}\n@media (prefers-color-scheme: dark){ :root:not([data-theme=\"light\"]){ --bg:#0d0d0d; --surface:#1a1a19; --fg:#f0efec; --fg2:#c3c2b7; --rule:#2c2c2a; --axis:#383835; --accent:#6da7ec; --accent-soft:#17293f; --s1:#3987e5; --s2:#d95926; --s3:#199e70; --seq1:#184f95; --seq2:#256abf; --seq3:#5598e7; --seq4:#b7d3f6; color-scheme:dark } }\n:root[data-theme=\"dark\"]{ --bg:#0d0d0d; --surface:#1a1a19; --fg:#f0efec; --fg2:#c3c2b7; --rule:#2c2c2a; --axis:#383835; --accent:#6da7ec; --accent-soft:#17293f; --s1:#3987e5; --s2:#d95926; --s3:#199e70; --seq1:#184f95; --seq2:#256abf; --seq3:#5598e7; --seq4:#b7d3f6; color-scheme:dark }\n*{box-sizing:border-box}\nbody{margin:0;background:var(--bg);color:var(--fg);font-family:var(--font);font-size:15px;line-height:1.55}\nbody::before{content:\"\";position:fixed;inset:0 0 auto;height:env(safe-area-inset-top,0px);z-index:10;background:var(--bg)}\nh1,h2,h3,h4{margin:0;font-weight:700;text-wrap:balance;line-height:1.3} h1{font-size:1.15rem} h2{font-size:1.05rem} h3{font-size:.95rem} h4{font-size:.85rem;color:var(--fg2);font-weight:500}\np{margin:0} button,select,input{font:inherit;color:inherit}\n.top{position:sticky;top:env(safe-area-inset-top,0px);z-index:5;background:var(--bg);border-bottom:1px solid var(--rule);padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));padding-block:10px}\n.top-in{max-width:1100px;margin:0 auto;display:flex;flex-wrap:wrap;align-items:center;gap:8px 20px}\n.dept{color:var(--fg2);font-size:.85rem}\nnav{display:flex;flex-wrap:wrap;gap:4px;margin-left:auto}\nnav button{background:none;border:0;border-bottom:2px solid transparent;padding:6px 10px;cursor:pointer;color:var(--fg2)}\nnav button[aria-current=\"page\"]{color:var(--fg);border-bottom-color:var(--accent);font-weight:500}\nbutton:focus-visible,select:focus-visible,input:focus-visible{outline:2px solid var(--accent);outline-offset:2px}\nmain{max-width:1100px;margin:0 auto;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));padding-block:20px 60px}\nsection{padding-block:22px;border-bottom:1px solid var(--rule)} section:last-child{border-bottom:0}\n.sec-head{display:flex;flex-wrap:wrap;align-items:baseline;gap:6px 16px;margin-bottom:14px} .note{color:var(--fg2);font-size:.85rem}\n.row{display:flex;flex-wrap:wrap;gap:10px 16px;align-items:center}\n.btn{border:1px solid var(--axis);background:var(--surface);border-radius:6px;padding:6px 12px;cursor:pointer} .btn.primary{background:var(--accent);border-color:var(--accent);color:#fff} .btn.sm{padding:3px 8px;font-size:.85rem}\nselect,input[type=text]{border:1px solid var(--axis);border-radius:6px;background:var(--surface);padding:6px 8px}\nlabel.f{display:flex;flex-direction:column;gap:6px;font-size:.85rem;color:var(--fg2);border-top:2px solid var(--rule);padding-top:10px} label.f input[type=file]{font-size:.85rem;margin-top:4px}\nlabel.f .t{font-size:1rem;color:var(--fg);font-weight:600;display:flex;gap:8px;align-items:baseline} label.f .step{font-family:var(--mono,ui-monospace,monospace);font-size:.75rem;color:var(--accent)}\nlabel.f .tag{font-size:.7rem;padding:0 6px;border-radius:9px;background:var(--accent-soft);color:var(--fg2)} label.f .tag.req{background:var(--accent);color:#fff}\nlabel.f .d{color:var(--fg2);line-height:1.55;min-height:3em;font-size:.92rem} label.f .when{color:var(--muted);font-size:.85rem} label.f .pick{display:flex;gap:10px;align-items:center;margin-top:6px} label.f .fn{color:var(--muted);font-size:.85rem} label.f.over{background:var(--accent-soft)} .loads .ok,.loads .no{font-size:.88rem!important}\n.loads{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:14px 28px}\n.loads .ok{color:var(--good);font-size:.8rem} .loads .no{color:var(--muted);font-size:.8rem}\n.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px 24px} .tile .v{font-size:2rem;font-weight:700;line-height:1.1} .tile .l{color:var(--fg2);font-size:.85rem} .tile .sub{color:var(--muted);font-size:.8rem}\n.tbl{overflow-x:auto} table{border-collapse:collapse;width:100%;font-size:.9rem} th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--rule);vertical-align:top} th{color:var(--fg2);font-weight:500;font-size:.8rem;white-space:nowrap} td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}\ntd.nw{white-space:nowrap} tr.click{cursor:pointer} tr.click:hover td,tr.cur td{background:var(--accent-soft)}\n.pill{display:inline-block;padding:1px 8px;border-radius:999px;font-size:.78rem;border:1px solid var(--axis);color:var(--fg2);white-space:nowrap} .pill.t1{border-color:var(--s1);color:var(--s1)} .pill.t2{border-color:var(--s2);color:var(--s2)} .pill.t3{border-color:var(--s3);color:var(--s3)}\n.dot{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px;vertical-align:middle}\n.flag{color:var(--critical);font-weight:500} .flag::before{content:\"▲ \";font-size:.7em} .agree{color:var(--good)} .disagree{color:var(--serious)}\n.charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(440px,1fr));gap:30px 48px} @media (max-width:520px){ .charts{grid-template-columns:1fr} } .chart h3{margin-bottom:4px} .chart .cap{color:var(--fg2);font-size:.82rem;margin-bottom:8px}\nsvg{display:block;width:100%;height:auto;max-width:100%} svg text{font-family:var(--font);fill:var(--fg2);font-size:14px} svg .grid line{stroke:var(--rule)} svg .axis{stroke:var(--axis)} svg .lbl{fill:var(--fg);font-variant-numeric:tabular-nums}\n.legend{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:.82rem;color:var(--fg2);margin-bottom:8px} .empty{color:var(--muted);padding:18px 0} .small{font-size:.85rem;color:var(--fg2)} .muted{color:var(--muted)}\n.tip{position:fixed;pointer-events:none;background:var(--surface);border:1px solid var(--axis);border-radius:6px;padding:6px 10px;font-size:.82rem;z-index:20;max-width:260px}\n.xtab td.cell{text-align:center;font-variant-numeric:tabular-nums}\n.card{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px 32px} .card .kv{display:flex;flex-direction:column;gap:2px;min-width:0} .card .kv .k{font-size:.78rem;letter-spacing:.06em;color:var(--muted)} .card .kv .v{white-space:pre-wrap;word-break:break-word} .card .kv.wide{grid-column:1/-1}\n.proj{border-left:3px solid var(--accent);padding-left:14px}\n.decide{display:flex;flex-direction:column;gap:10px;margin-top:14px;padding-top:14px;border-top:2px solid var(--rule);font-size:1rem} .decide label{display:flex;flex-direction:column;gap:3px;color:var(--fg2);font-size:.9rem} .decide select,.decide textarea,.decide input[type=text],.decide input[type=number]{font:inherit;font-size:1rem;color:var(--fg);border:1px solid var(--axis);border-radius:6px;background:var(--surface);padding:6px 8px}\n.decide .lbl{color:var(--fg2);font-size:.9rem} .decide .seg{display:flex;flex-wrap:wrap;gap:6px 16px} .decide .seg label{flex-direction:row;align-items:center;gap:5px;color:var(--fg);font-size:1rem}\n.decide .vote{display:flex;flex-wrap:wrap;gap:8px 12px} .decide .vote label{width:6em} .decide .vote label.w{flex:1;min-width:10em;width:auto}\n.decide details summary{cursor:pointer;color:var(--fg2);font-size:.9rem} .decide .fac{max-height:240px;overflow:auto;padding:6px 0;font-size:.9rem} .decide .fac .g{color:var(--muted);font-size:.8rem;margin-top:6px} .decide .ck{flex-direction:row!important;align-items:center;gap:6px!important;color:var(--fg)!important}\n.decide .go{display:flex;align-items:center;gap:12px;flex-wrap:wrap} .pres .flash{background:var(--accent-soft);padding:8px 12px;border-radius:6px;margin-bottom:10px;font-size:1rem;max-width:40em}\n.flow{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:18px 32px} .flow>div{border-top:2px solid var(--rule);padding-top:10px;min-width:0} .flow .ph{font-weight:600;margin-bottom:6px} .flow ol{margin:0;padding-left:1.3em;font-size:.9rem;line-height:1.6} .flow li{margin:3px 0} .flow a,.small a{color:var(--accent)} code{font-family:var(--mono,ui-monospace,monospace);font-size:.9em;background:var(--accent-soft);padding:0 4px;border-radius:3px}\n.banner{background:var(--accent-soft);padding:10px 14px;border-radius:6px;font-size:.9rem}\nbody.present .top{display:none} body.present main{max-width:1400px;padding-block:28px}\n.pres{font-size:1.25rem;line-height:1.6} .pres h2{font-size:2rem;margin-bottom:6px} .pres .meta{color:var(--fg2);font-size:1rem;margin-bottom:18px} .pres .card{grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:18px 48px} .pres .pill{font-size:1rem;padding:3px 12px}\n.pres .cols{display:grid;grid-template-columns:2fr 1fr;gap:32px 56px;align-items:start} .pres .wait{font-size:1.6rem;color:var(--muted);padding:80px 0;text-align:center}\n.pres .ctl{position:fixed;top:max(12px,env(safe-area-inset-top,0px));right:max(16px,env(safe-area-inset-right,0px));display:flex;gap:8px;align-items:center;background:var(--bg);padding:4px 0 4px 10px;z-index:5} .pres .ctl .pos{font-variant-numeric:tabular-nums;color:var(--fg2);font-size:1rem;min-width:3.5em;text-align:center} .pres .ctl .hint{color:var(--muted);font-size:.8rem;margin-left:6px}\n.pres .meta.above{margin-bottom:4px;font-size:1rem} .pres .dx{font-size:1.5rem;line-height:1.4;margin:0 0 22px;max-width:50em;text-wrap:balance} .pres .todo{margin:0;padding-left:1.4em;line-height:1.7} @media (max-width:700px){ .pres .ctl{position:static;padding:0;margin-bottom:10px;flex-wrap:wrap} .pres .ctl .hint{display:none} }\n@media (max-width:900px){ .pres .cols{grid-template-columns:1fr} } @media (max-width:560px){ .dept{display:none} nav{margin-left:0} }\n", DASH_BODY = "<div class=\"top\"><div class=\"top-in\">\n  <div><h1>M&amp;M 會議儀表板</h1><div class=\"dept\">離線版：載入 SharePoint 匯出檔 → 會前看待討論案例、投影 → 會後看紀錄與分析</div></div>\n  <nav id=\"nav\"><button data-view=\"load\">載入</button><button data-view=\"prep\">待討論案例</button><button data-view=\"records\">紀錄與分析</button><button data-view=\"about\">說明</button></nav>\n</div></div>\n<main>\n  <div id=\"v-load\" hidden></div><div id=\"v-records\" hidden></div><div id=\"v-prep\" hidden></div><div id=\"v-about\" hidden></div><div id=\"v-present\" hidden></div>\n</main>\n<div id=\"tip\" class=\"tip\" hidden></div>", DASH_JS = "\n'use strict';\nconst SP=window.MM_SP||null; /* 連線版：由 SharePoint 書籤注入資料與寫回函式 */\nconst VIEWS=['load','records','prep','about','present'];\nconst TYPES=['超長住院','Morbidity','Mortality']; const TYPE_CLASS={'超長住院':'t1','Morbidity':'t2','Mortality':'t3'}; const TYPE_VAR={'超長住院':'--s1','Morbidity':'--s2','Mortality':'--s3'};\nconst PREV=['不可預防','可能可預防','很可能可預防','明確可預防']; const DMODE=['三方一致','主席裁定','投票'];\nconst state={view:'load',records:[],drafts:[],factors:[],actions:[],loaded:{},recYm:null,paretoRange:'12',paretoType:'全部',open:null,presIdx:0,thr:30,preview:false};\nconst $=(s,el=document)=>el.querySelector(s); const esc=s=>String(s??'').replace(/[&<>\"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'}[c]));\nconst nowYm=()=>{const d=new Date();return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');};\nconst addMonths=(ym,n)=>{let [y,m]=ym.split('-').map(Number);m+=n;y+=Math.floor((m-1)/12);m=((m-1)%12+12)%12+1;return y+'-'+String(m).padStart(2,'0');};\nconst ymLabel=ym=>ym?ym.replace('-',' 年 ')+' 月':''; const rangeYms=(e,n)=>{const a=[];for(let i=n-1;i>=0;i--)a.push(addMonths(e,-i));return a;};\nconst hasType=(c,t)=>(c.types||[]).includes(t); const facGroup=n=>{const f=state.factors.find(x=>x.name===n);return f?f.group:'未分類';}; const groupsOf=()=>[...new Set(state.factors.map(f=>f.group))];\n\n/* ---------- 檔案解析 ---------- */\nfunction parseCSV(text){ if(text.charCodeAt(0)===0xFEFF) text=text.slice(1); const rows=[]; let row=[],cell='',q=false; for(let i=0;i<text.length;i++){ const ch=text[i]; if(q){ if(ch==='\"'){ if(text[i+1]==='\"'){cell+='\"';i++;} else q=false; } else cell+=ch; } else { if(ch==='\"') q=true; else if(ch===','){row.push(cell);cell='';} else if(ch==='\\n'||ch==='\\r'){ if(ch==='\\r'&&text[i+1]==='\\n') i++; row.push(cell); rows.push(row); row=[];cell=''; } else cell+=ch; } } if(cell!==''||row.length){row.push(cell);rows.push(row);} return rows.filter(r=>r.some(c=>c.trim()!=='')); }\nfunction rowsToObjects(rows){ const h=rows[0].map(x=>x.trim()); return rows.slice(1).map(r=>{const o={};h.forEach((k,i)=>o[k]=(r[i]??'').toString().trim());return o;}); }\nasync function readFile(file){ const name=file.name.toLowerCase(); if(name.endsWith('.csv')||name.endsWith('.txt')) return rowsToObjects(parseCSV(await file.text()));\n  if(!window.XLSX) throw new Error('沒有網路載入 Excel 解析器，請改用「匯出至 CSV」'); const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]]; return rowsToObjects(XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''}).map(r=>r.map(x=>String(x??'')))); }\nconst col=(o,...names)=>{ for(const n of names){ for(const k of Object.keys(o)){ if(k.replace(/\\s/g,'')===n) return o[k]; } } return ''; };\nconst multi=s=>String(s||'').split(/;#|;|；|\\n/).map(x=>x.trim()).filter(x=>x&&!/^\\d+$/.test(x));\nconst num=s=>{const n=parseFloat(String(s).replace(/[^\\d.-]/g,''));return isNaN(n)?null:n;};\nconst ymOf=(s)=>{ const m=String(s).match(/(\\d{4})[-/.](\\d{1,2})/); return m?`${m[1]}-${m[2].padStart(2,'0')}`:''; };\nfunction mapCase(o){ return { code:col(o,'標題','案例代碼','Title'), ym:ymOf(col(o,'年月')), types:multi(col(o,'案例類型')), ageBand:col(o,'年齡層'), age:col(o,'年齡'), nameShort:col(o,'病人簡稱'), sex:col(o,'性別'), dxGroup:col(o,'診斷群'), phase:col(o,'治療階段'),\n  dxDetail:col(o,'詳細診斷'), treatment:col(o,'治療內容'), course:col(o,'臨床經過'), los:num(col(o,'住院天數')), avoidableDays:num(col(o,'可避免天數')), deathExpected:col(o,'死亡預期性'), morbGrade:col(o,'Morbidity等級'),\n  initial:{mainFactor:col(o,'初評主要根因'),factors:multi(col(o,'初評促成因素')),preventability:col(o,'初評可預防性')},\n  attending:{status:col(o,'主治醫師回覆')||'未徵詢',mainFactor:col(o,'主治認為主要根因'),preventability:col(o,'主治認為可預防性'),comment:col(o,'主治醫師說明')},\n  whys:col(o,'五個為什麼','5Whys'), teachingPoints:col(o,'學習重點'), refs:col(o,'參考指引文獻','參考'), status:col(o,'狀態') }; }\nfunction mapRecord(o){ const r=mapCase(o); r.consensus={mainFactor:col(o,'共識主要根因'),factors:multi(col(o,'共識促成因素')),preventability:col(o,'共識可預防性'),notes:col(o,'會議結論'),voteCount:num(col(o,'票數'))||0,dist:col(o,'投票分布')};\n  r.decisionMode=col(o,'決策方式'); r.chiefComment=col(o,'長官意見'); r.meetingDate=col(o,'會議日期'); r.attendees=num(col(o,'出席人數'))||0; r.amendments=col(o,'修正附註');\n  const ag=col(o,'初評一致'); r.agreeMain = ag?/是|yes|true/i.test(ag):(r.initial.mainFactor===r.consensus.mainFactor); if(!r.ym) r.ym=ymOf(r.meetingDate); return r; }\nfunction mapFactor(o){ return {name:col(o,'標題','因素名稱','Title'),group:col(o,'層級','群組'),order:num(col(o,'排序'))||0}; }\nfunction mapAction(o){ return {title:col(o,'標題','行動','Title'),owner:col(o,'負責人'),due:col(o,'到期日'),status:col(o,'狀態'),factor:col(o,'針對因素'),cases:multi(col(o,'關聯紀錄','關聯案例')),note:col(o,'備註')}; }\nasync function load(kind,file){ const objs=await readFile(file); if(kind==='records') state.records=objs.map(mapRecord).filter(r=>r.code); if(kind==='drafts') state.drafts=objs.map(mapCase).filter(r=>r.code); if(kind==='factors') state.factors=objs.map(mapFactor).filter(f=>f.name).sort((a,b)=>a.order-b.order); if(kind==='actions') state.actions=objs.map(mapAction).filter(a=>a.title);\n  state.loaded[kind]={name:file.name,n:objs.length,at:new Date().toLocaleString()}; if(kind==='records'){ const l=[...state.records].sort((a,b)=>b.ym.localeCompare(a.ym))[0]; state.recYm=l?l.ym:nowYm(); } persist(); }\nfunction persist(){ if(SP) return; try{ localStorage.setItem('mm_dash',JSON.stringify({records:state.records,drafts:state.drafts,factors:state.factors,actions:state.actions,loaded:state.loaded,thr:state.thr})); }catch(e){} }\nfunction restore(){ try{ const s=JSON.parse(localStorage.getItem('mm_dash')||'null'); if(s){ Object.assign(state,s); const l=[...state.records].sort((a,b)=>b.ym.localeCompare(a.ym))[0]; state.recYm=l?l.ym:nowYm(); } }catch(e){} }\n\nfunction spApply(d){ const at=new Date().toLocaleString(); state.drafts=d.drafts.map(mapCase).filter(r=>r.code); state.records=d.records.map(mapRecord).filter(r=>r.code); state.factors=d.factors.map(mapFactor).filter(f=>f.name).sort((a,b)=>a.order-b.order); state.actions=d.actions.map(mapAction).filter(a=>a.title);\n  state.loaded={drafts:{name:'SharePoint',n:state.drafts.length,at},records:{name:'SharePoint',n:state.records.length,at},factors:{name:'SharePoint',n:state.factors.length,at},actions:{name:'SharePoint',n:state.actions.length,at}};\n  const l=[...state.records].sort((a,b)=>b.ym.localeCompare(a.ym))[0]; state.recYm=l?l.ym:nowYm(); }\nif(SP) SP.dashboard={ reload:d=>{ spApply(d); renderAll(); } };\n\n/* ---------- 範例資料（示範用） ---------- */\nconst SAMPLE_FACTORS=[['病人因素','疾病嚴重度／多重共病'],['病人因素','病情急速惡化'],['工作任務與流程','檢驗檢查結果延遲或未追蹤'],['工作任務與流程','用藥流程（處方／調劑／給藥）'],['工作任務與流程','交班資訊不全'],['團隊與溝通','跨科會診延遲'],['團隊與溝通','團隊內溝通不良'],['工作環境','人力不足'],['工作環境','夜間／假日資源不足'],['超長住院因素','等待檢查／處置'],['超長住院因素','感染併發症'],['超長住院因素','化療／治療併發症'],['超長住院因素','安置／家屬因素'],['超長住院因素','轉院／安寧銜接等待']];\nconst SAMPLE_RECORDS=[\n ['DEMO-04','2026-10','Mortality','≥80','男','淋巴瘤','緩和／支持性','範例：套細胞淋巴瘤，TP53 突變，第三線 BTK 抑制劑後惡化，ECOG 3，安寧共照。','停止抗癌治療，症狀控制。','住院第 3 天安寧共照介入，第 12 天往生。',12,'','預期（疾病進展）','不適用','疾病嚴重度／多重共病','','不可預防','同意初評','','','','','TP53 突變 MCL 預後極差，第二線後即應討論照護目標。','','疾病嚴重度／多重共病','','不可預防','三方一致','','預期死亡。','2026-10-20',11,0,'','是',''],\n ['DEMO-05','2026-10','超長住院','40–49','女','急性白血病','鞏固／維持','範例：AML CBFB-MYH11，ELN 良好風險，HiDAC 鞏固第 2 週期。','HiDAC；發燒 cefepime → 依培養改 meropenem。','第 10 天發燒，第 12 天培養陽性但第 14 天才調整抗生素；第 20 天退燒；床位安排延至第 55 天。',55,7,'不適用','不適用','感染併發症','檢驗檢查結果延遲或未追蹤;人力不足','可能可預防','不同意初評','檢驗檢查結果延遲或未追蹤','','感染是結果，根因在報告沒人追。','','培養陽性到抗生素調整的時間差要定期稽核。','','檢驗檢查結果延遲或未追蹤','感染併發症;人力不足','可能可預防','投票','兩案連續出現報告追蹤問題。','共識認為根因在報告追蹤。','2026-10-20',11,10,'檢驗檢查結果延遲或未追蹤 7／感染併發症 3','否',''],\n ['DEMO-06','2026-10','Morbidity','60–69','女','大腸直腸癌','第一線（轉移／晚期）','範例：右側大腸癌 pT3N1bM1a，RAS/BRAF 野生型，MSS，FOLFOX + cetuximab 第 6 週期。','未依腎功能調劑量。','第 2 天 eGFR 降至 45 交班未註明，第 3 天原劑量給藥，第 6 天 grade 3 血小板低下。',9,'','不適用','II','用藥流程（處方／調劑／給藥）','交班資訊不全','很可能可預防','同意初評','','','','','給藥前固定核對當日腎功能。','','用藥流程（處方／調劑／給藥）','交班資訊不全;檢驗檢查結果延遲或未追蹤','很可能可預防','主席裁定','核對流程補腎功能項目。','','2026-10-20',11,0,'','是',''],\n ['DEMO-07','2026-09','超長住院','70–79','男','MDS／MPN','緩和／支持性','範例：MDS-IB2，IPSS-M 高風險，azacitidine 第 5 週期，輸血依賴。','azacitidine；輸血；轉安寧病房。','第 15 天起病情穩定但安置意見不一；第 40 天社工介入；第 74 天轉安寧病房。',74,20,'不適用','不適用','安置／家屬因素','轉院／安寧銜接等待;檢驗檢查結果延遲或未追蹤','可能可預防','同意初評','','','','','高風險 MDS 入院即啟動出院計畫。','','安置／家屬因素','轉院／安寧銜接等待','可能可預防','三方一致','','社工介入太晚。','2026-09-22',9,0,'','是',''],\n ['DEMO-08','2026-09','Mortality','60–69','男','急性白血病','復發／難治','範例：復發 AML，IDH2 R140Q，第二線 azacitidine + venetoclax 第 2 週期。','第 5 天腹痛請外科會診。','第 5 天腹痛，CT 當晚完成但報告隔日下午才被看到（穿孔）；第 6 天手術；第 9 天敗血症死亡。',30,'','非預期','不適用','跨科會診延遲','檢驗檢查結果延遲或未追蹤;團隊內溝通不良','可能可預防','不同意初評','檢驗檢查結果延遲或未追蹤','很可能可預防','CT 當晚就有穿孔。','','中性球低下病人腹痛要當急症；影像危急值要通知。','','檢驗檢查結果延遲或未追蹤','跨科會診延遲;團隊內溝通不良','很可能可預防','投票','影像危急值通知要跟影像科談。','危急值未通知為根因。','2026-09-22',9,8,'檢驗檢查結果延遲或未追蹤 5／跨科會診延遲 3','否','2026-10-01 補充：已與影像科協調危急值流程。']];\nconst REC_HEAD=['標題','年月','案例類型','年齡層','性別','診斷群','治療階段','詳細診斷','治療內容','臨床經過','住院天數','可避免天數','死亡預期性','Morbidity等級','初評主要根因','初評促成因素','初評可預防性','主治醫師回覆','主治認為主要根因','主治認為可預防性','主治醫師說明','五個為什麼','學習重點','參考指引文獻','共識主要根因','共識促成因素','共識可預防性','決策方式','長官意見','會議結論','會議日期','出席人數','票數','投票分布','初評一致','修正附註'];\nfunction loadSample(){ const toObj=(h,r)=>{const o={};h.forEach((k,i)=>o[k]=String(r[i]??''));return o;};\n  state.factors=SAMPLE_FACTORS.map(([g,n],i)=>({name:n,group:g,order:i})); state.records=SAMPLE_RECORDS.map(r=>mapRecord(toObj(REC_HEAD,r)));\n  state.drafts=[mapCase(toObj(REC_HEAD.slice(0,24).concat(['狀態']),['DEMO-01','2026-11','超長住院','60–69','男','淋巴瘤','鞏固／維持','範例：瀰漫性大 B 細胞淋巴瘤，non-GCB，Ann Arbor IV 期（骨髓侵犯），IPI 4，R-CHOP 第 4 週期後。','R-CHOP；發燒 cefepime；G-CSF。','第 1–5 天嗜中性白血球低下發燒；第 7 天起等待 PET-CT 排程，第 14 天檢查，第 18 天出院。',48,10,'不適用','不適用','等待檢查／處置','檢驗檢查結果延遲或未追蹤','可能可預防','同意初評','','','下次期中評估改門診排程。','','期中評估影像可在門診完成。','','待討論'])),\n   mapCase(toObj(REC_HEAD.slice(0,24).concat(['狀態']),['DEMO-02','2026-11','Mortality','70–79','女','急性白血病','誘導','範例：AML，NPM1 突變、FLT3-ITD 陽性，ELN 2022 中度風險，7+3 合併 midostaurin 誘導第 1 週期。','7+3 + midostaurin；第 9 天發燒改 piperacillin-tazobactam。','第 9 天發燒抽血培養；第 11 天（週六）培養陽性報告未被看到；第 13 天才調整抗生素並轉 ICU；第 15 天死亡。',22,'','非預期','不適用','檢驗檢查結果延遲或未追蹤','跨科會診延遲;夜間／假日資源不足','很可能可預防','不同意初評','夜間／假日資源不足','可能可預防','假日微生物室人力不足，報告本身就慢半天。','為什麼延遲？假日無主動通知。','培養陽性必須有主動通知路徑。','IDSA 發燒性嗜中性白血球低下指引','待討論']))];\n  state.actions=[mapAction({'標題':'範例：培養／檢驗異常結果 24 小時內追蹤回報','負責人':'研究醫師','到期日':'2026-12-31','狀態':'進行中','針對因素':'檢驗檢查結果延遲或未追蹤','關聯紀錄':'DEMO-05;DEMO-08'})];\n  state.loaded={records:{name:'範例資料',n:5,at:new Date().toLocaleString()},drafts:{name:'範例資料',n:2},factors:{name:'範例資料',n:14},actions:{name:'範例資料',n:1}}; state.recYm='2026-10'; persist(); }\n\n/* ---------- 導覽 ---------- */\nfunction setView(v){ state.view=v; document.querySelectorAll('#nav button').forEach(b=>b.setAttribute('aria-current',b.dataset.view===v?'page':'false')); VIEWS.forEach(k=>$('#v-'+k).hidden=k!==v); document.body.classList.toggle('present',v==='present'); if(!SP&&location.hash.slice(1)!==v){ try{ history.replaceState(null,'','#'+v); }catch(e){} } if(SP&&SP.onView) SP.onView(v); renderAll(); }\ndocument.querySelectorAll('#nav button').forEach(b=>b.addEventListener('click',()=>setView(b.dataset.view)));\nwindow.addEventListener('hashchange',()=>{const h=location.hash.slice(1); if(VIEWS.includes(h)) setView(h);});\nfunction renderAll(){ ({load:renderLoad,records:renderRecords,prep:renderPrep,about:renderAbout,present:renderPresent})[state.view](); }\n\n/* ---------- tooltip & charts ---------- */\nconst tip=$('#tip'); function bindTips(root){ root.querySelectorAll('[data-tip]').forEach(el=>{ el.addEventListener('pointerenter',e=>{tip.textContent=el.dataset.tip;tip.hidden=false;moveTip(e);}); el.addEventListener('pointermove',moveTip); el.addEventListener('pointerleave',()=>tip.hidden=true); }); } function moveTip(e){ tip.style.left=Math.min(e.clientX+12,window.innerWidth-270)+'px'; tip.style.top=(e.clientY+14)+'px'; }\nconst recIn=ym=>state.records.filter(r=>r.ym===ym); const cMain=r=>r.consensus.mainFactor; const cFactors=r=>[...new Set([cMain(r),...r.consensus.factors].filter(Boolean))]; const cPrev=r=>r.consensus.preventability;\nfunction factorCounts(recs,mode){ const m=new Map(); recs.forEach(r=>(mode==='main'?(cMain(r)?[cMain(r)]:[]):cFactors(r)).forEach(n=>m.set(n,(m.get(n)||0)+1))); return [...m.entries()].map(([name,n])=>({name,n,group:facGroup(name)})).sort((a,b)=>b.n-a.n||a.name.localeCompare(b.name,'zh-Hant')); }\nfunction recurring(endYm){ const yms=rangeYms(endYm,6); const m=new Map(); state.records.filter(r=>yms.includes(r.ym)).forEach(r=>cFactors(r).forEach(n=>{ if(!m.has(n)) m.set(n,new Set()); m.get(n).add(r.ym); })); return [...m.entries()].filter(([,s])=>s.size>=3).map(([name,s])=>({name,months:[...s].sort()})).sort((a,b)=>b.months.length-a.months.length); }\nfunction stackedMonthly(yms){ const W=640,H=230,L=36,R=12,T=14,B=34,iw=W-L-R,ih=H-T-B; const data=yms.map(ym=>({ym,v:TYPES.map(t=>recIn(ym).filter(c=>hasType(c,t)).length)})); const max=Math.max(1,...data.map(d=>d.v.reduce((a,b)=>a+b,0))); const step=max<=5?1:max<=10?2:max<=20?5:10; const top=Math.ceil(max/step)*step; const y=v=>T+ih-(v/top)*ih; const bw=Math.min(36,iw/yms.length*0.6); const xs=i=>L+iw*(i+0.5)/yms.length;\n  let g='<g class=\"grid\">'; for(let v=0;v<=top;v+=step) g+=`<line x1=\"${L}\" x2=\"${W-R}\" y1=\"${y(v)}\" y2=\"${y(v)}\"/><text x=\"${L-6}\" y=\"${y(v)+4}\" text-anchor=\"end\">${v}</text>`; g+='</g>'; let bars='';\n  data.forEach((d,i)=>{ let acc=0; const tot=d.v.reduce((a,b)=>a+b,0); d.v.forEach((n,k)=>{ if(!n) return; const y1=y(acc+n),h=y(acc)-y1-(acc?2:0); bars+=`<rect x=\"${xs(i)-bw/2}\" y=\"${y1}\" width=\"${bw}\" height=\"${Math.max(0,h)}\" rx=\"${acc+n===tot?4:0}\" fill=\"var(${TYPE_VAR[TYPES[k]]})\" data-tip=\"${esc(ymLabel(d.ym))} ${TYPES[k]}：${n} 件\"/>`; acc+=n; }); if(tot) bars+=`<text class=\"lbl\" x=\"${xs(i)}\" y=\"${y(tot)-5}\" text-anchor=\"middle\">${tot}</text>`; bars+=`<text x=\"${xs(i)}\" y=\"${H-B+16}\" text-anchor=\"middle\" ${d.ym===state.recYm?'font-weight=\"700\" fill=\"var(--fg)\"':''}>${d.ym.slice(2).replace('-','/')}</text>`; });\n  return `<svg viewBox=\"0 0 ${W} ${H}\" role=\"img\" aria-label=\"每月定案案例數\">${g}<line class=\"axis\" x1=\"${L}\" x2=\"${W-R}\" y1=\"${y(0)}\" y2=\"${y(0)}\"/>${bars}</svg>`; }\nfunction paretoChart(rows){ if(!rows.length) return '<div class=\"empty\">此範圍內尚無紀錄。</div>'; rows=rows.slice(0,10); const total=rows.reduce((a,r)=>a+r.n,0); const W=640,rowH=28,L=200,R=70,T=6,H=T+rows.length*rowH+8,iw=W-L-R,max=rows[0].n; let acc=0,s=''; rows.forEach((r,i)=>{ acc+=r.n; const w=Math.max(2,iw*r.n/max),y=T+i*rowH; s+=`<text x=\"${L-8}\" y=\"${y+18}\" text-anchor=\"end\">${esc(r.name.length>13?r.name.slice(0,12)+'…':r.name)}</text><rect x=\"${L}\" y=\"${y+6}\" width=\"${w}\" height=\"16\" rx=\"4\" fill=\"var(--s1)\" data-tip=\"${esc(r.name)}（${esc(r.group)}）：${r.n} 件，累積 ${Math.round(acc/total*100)}%\"/>${w>iw-30?`<text class=\"lbl\" x=\"${L+w-6}\" y=\"${y+18}\" text-anchor=\"end\" fill=\"#fff\" style=\"fill:#fff\">${r.n}</text>`:`<text class=\"lbl\" x=\"${L+w+6}\" y=\"${y+18}\">${r.n}</text>`}<text x=\"${W-R+10}\" y=\"${y+18}\" fill=\"var(--muted)\">${Math.round(acc/total*100)}%</text>`; }); return `<svg viewBox=\"0 0 ${W} ${H}\" role=\"img\" aria-label=\"共識主要根因柏拉圖\">${s}</svg>`; }\nfunction prevChart(recs){ const counts=PREV.map(p=>recs.filter(r=>cPrev(r)===p).length); const total=counts.reduce((a,b)=>a+b,0); if(!total) return '<div class=\"empty\">尚無資料。</div>'; const W=640,H=64; let x=0,s='',lbl=''; const cols=['var(--seq1)','var(--seq2)','var(--seq3)','var(--seq4)']; counts.forEach((n,i)=>{ if(!n) return; const w=W*n/total; s+=`<rect x=\"${x}\" y=\"8\" width=\"${Math.max(0,w-2)}\" height=\"22\" rx=\"3\" fill=\"${cols[i]}\" data-tip=\"${PREV[i]}：${n} 件（${Math.round(n/total*100)}%）\"/>`; if(w>70) lbl+=`<text x=\"${x+w/2}\" y=\"48\" text-anchor=\"middle\">${PREV[i]} ${n}</text>`; x+=w; }); return `<svg viewBox=\"0 0 ${W} ${H}\" role=\"img\" aria-label=\"可預防性分布\">${s}${lbl}</svg>`; }\nfunction losChart(recs){ const thr=state.thr||30; const bins=[[thr,thr+14],[thr+15,thr+29],[thr+30,thr+59],[thr+60,Infinity]]; const labels=bins.map(([a,b])=>b===Infinity?`≥${a} 天`:`${a}–${b} 天`); const counts=bins.map(([a,b])=>recs.filter(c=>c.los>=a&&c.los<=b).length); if(!counts.some(Boolean)) return '<div class=\"empty\">此範圍內無超長住院紀錄。</div>'; const W=640,H=170,L=30,T=12,B=30,ih=H-T-B,iw=W-L-10,max=Math.max(...counts); const y=v=>T+ih-(v/max)*ih; const bw=Math.min(70,iw/4*0.6); let s=''; counts.forEach((n,i)=>{ const cx=L+iw*(i+0.5)/4; s+=`<rect x=\"${cx-bw/2}\" y=\"${y(n)}\" width=\"${bw}\" height=\"${y(0)-y(n)}\" rx=\"4\" fill=\"var(--s1)\" data-tip=\"${labels[i]}：${n} 件\"/><text class=\"lbl\" x=\"${cx}\" y=\"${y(n)-5}\" text-anchor=\"middle\">${n||''}</text><text x=\"${cx}\" y=\"${H-B+16}\" text-anchor=\"middle\">${labels[i]}</text>`; }); return `<svg viewBox=\"0 0 ${W} ${H}\" role=\"img\" aria-label=\"超長住院天數分布\"><line class=\"axis\" x1=\"${L}\" x2=\"${W-10}\" y1=\"${y(0)}\" y2=\"${y(0)}\"/>${s}</svg>`; }\nfunction crossTab(recs){ const groups=[...new Set([...groupsOf(),...recs.map(r=>facGroup(cMain(r)))])]; const rows=groups.map(g=>({g,v:TYPES.map(t=>recs.filter(r=>hasType(r,t)&&facGroup(cMain(r))===g).length)})).filter(r=>r.v.some(Boolean)); if(!rows.length) return '<div class=\"empty\">尚無資料。</div>'; const max=Math.max(...rows.flatMap(r=>r.v)); const colr=n=>n===0?'transparent':n<=max*0.25?'var(--seq1)':n<=max*0.5?'var(--seq2)':n<=max*0.75?'var(--seq3)':'var(--seq4)'; const ink=n=>n>max*0.5?'#fff':'var(--fg)'; return `<div class=\"tbl\"><table class=\"xtab\"><thead><tr><th>共識主要根因層級</th>${TYPES.map(t=>`<th class=\"num\">${t}</th>`).join('')}<th class=\"num\">合計</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${esc(r.g)}</td>${r.v.map(n=>`<td class=\"cell\" style=\"background:${colr(n)};color:${ink(n)}\">${n||''}</td>`).join('')}<td class=\"num\">${r.v.reduce((a,b)=>a+b,0)}</td></tr>`).join('')}</tbody></table></div>`; }\n\n/* ---------- 共用片段 ---------- */\nconst typePills=c=>(c.types||[]).map(t=>`<span class=\"pill ${TYPE_CLASS[t]||''}\">${esc(t)}</span>`).join(' ');\nconst kv=(k,v,wide)=>v?`<div class=\"kv${wide?' wide':''}\"><div class=\"k\">${k}</div><div class=\"v\">${esc(v)}</div></div>`:'';\nconst items=v=>String(v||'').split(/｜|\\n/).map(x=>x.trim()).filter(Boolean);\nconst kvList=(k,v)=>{const L=items(v); return L.length?`<div class=\"kv wide\"><div class=\"k\">${k}</div><div class=\"v\">${L.length>1?L.map(x=>'・'+esc(x)).join('\\n'):esc(L[0])}</div></div>`:'';};\nfunction teachingCard(c){ return `<div class=\"card\">${kv('案例代碼',c.code)}${kv('年月',ymLabel(c.ym))}<div class=\"kv\"><div class=\"k\">類型</div><div class=\"v\">${typePills(c)}</div></div>${kv('病人',[c.nameShort,c.age?c.age+' 歲':c.ageBand,c.sex].filter(Boolean).join('　'))}${kv('診斷群／治療階段',[c.dxGroup,c.phase].filter(Boolean).join('　'))}${kv('住院天數',c.los!=null?c.los+' 天'+(c.avoidableDays!=null?'（估可避免 '+c.avoidableDays+' 天）':''):'')}${kv('死亡預期性',c.deathExpected&&c.deathExpected!=='不適用'?c.deathExpected:'')}${kv('Morbidity 等級',c.morbGrade&&c.morbGrade!=='不適用'?'Clavien-Dindo '+c.morbGrade:'')}${kv('主診斷（詳細診斷）',c.dxDetail,true)}${kvList('診療史（治療內容）',c.treatment)}${kvList('其他診斷・併發症／住院經過（臨床經過）',c.course)}${kv('5 Whys',c.whys,true)}${kv('學習重點',c.teachingPoints,true)}${kv('參考指引／文獻',c.refs,true)}</div>`; }\nconst attCell=c=>{const st=c.attending.status; return st==='同意初評'?'<span class=\"agree\">同意</span>':st==='不同意初評'?'<span class=\"disagree\">不同意</span>':'<span class=\"muted\">未徵詢</span>';};\n\n/* ---------- 載入 ---------- */\nfunction nextBlock(){ const L=state.loaded; const q=queue().length, prep=state.drafts.filter(d=>d.status==='整理中').length, btns=[];\n  if(L.drafts){ btns.push(`<button class=\"btn primary\" data-go=\"prep\">待討論案例</button>`); if(q) btns.push(`<button class=\"btn\" data-go=\"present\">直接投影</button>`); }\n  if(L.records) btns.push(`<button class=\"btn ${L.drafts?'':'primary'}\" data-go=\"records\">紀錄與分析</button>`);\n  if(!btns.length) return '';\n  const hint=L.drafts?(q?`${q} 件待討論${prep?`、${prep} 件整理中`:''}。`:(prep?`${prep} 件都還是「整理中」（轉換器產出的檔都是），尚未排入會議；要預覽請在「待討論案例」勾「包含整理中」。`:'這個檔裡沒有待討論或整理中的案例。')):'';\n  return `<div class=\"proj row\" id=\"next\" style=\"margin-bottom:18px\"><span><b>已載入，下一步 →</b></span>${btns.join('')}<span class=\"small\">${hint} 資料留在這台瀏覽器，下次開啟無需重載；要換檔再回這頁。</span></div>`; }\nfunction renderLoad(){ const el=$('#v-load'); const L=state.loaded; const st=k=>L[k]?`<span class=\"ok\">已載入 ${esc(L[k].name)}（${L[k].n} 列${L[k].at?'，'+esc(L[k].at):''}）</span>`:'<span class=\"no\">尚未載入</span>';\n  el.innerHTML=`<section><div class=\"sec-head\"><h2>載入 SharePoint 匯出檔</h2><span class=\"note\">每個清單在 SharePoint 右上「匯出 → 匯出至 CSV」各存一個檔，對應拖進下面的格子。檔案只在這台電腦的瀏覽器處理，不會上傳。第一次用只要前兩個。</span></div>\n    ${nextBlock()}\n    <div class=\"loads\">\n      <label class=\"f\"><span class=\"t\"><span class=\"step\">1</span>案例 <span class=\"tag req\">會前必要</span></span><span class=\"d\">本月要討論的案例與初評、主治意見。給「待討論案例」與「投影模式」用。</span><span class=\"when\">來源：SharePoint 清單「案例」（也可直接用轉換器產出的 案例_匯入_*.xlsx 先看）</span><span class=\"pick\"><button type=\"button\" class=\"btn\" data-pick=\"drafts\">選擇檔案</button><span class=\"fn\">或把檔案拖到這一格</span></span><input type=\"file\" id=\"f-drafts\" accept=\".csv,.xlsx,.txt\" hidden>${st('drafts')}</label>\n      <label class=\"f\"><span class=\"t\"><span class=\"step\">2</span>正式紀錄 <span class=\"tag req\">分析必要</span></span><span class=\"d\">歷月已定案的共識根因、可預防性、決策方式。給「紀錄與分析」的所有圖表用；第一個月還沒有紀錄可先略過。</span><span class=\"when\">來源：SharePoint 清單「正式紀錄」</span><span class=\"pick\"><button type=\"button\" class=\"btn\" data-pick=\"records\">選擇檔案</button><span class=\"fn\">或把檔案拖到這一格</span></span><input type=\"file\" id=\"f-records\" accept=\".csv,.xlsx,.txt\" hidden>${st('records')}</label>\n      <label class=\"f\"><span class=\"t\"><span class=\"step\">3</span>根因分類 <span class=\"tag\">選用</span></span><span class=\"d\">27 項因素各屬哪一層（病人／個人／流程／團隊／環境／組織／超長住院）。載入後根因圖會依層分組並顯示層級交叉表；不載入也能畫圖，只是不分層。</span><span class=\"when\">來源：SharePoint 清單「根因分類」（內容固定，匯出一次留著重複用）</span><span class=\"pick\"><button type=\"button\" class=\"btn\" data-pick=\"factors\">選擇檔案</button><span class=\"fn\">或把檔案拖到這一格</span></span><input type=\"file\" id=\"f-factors\" accept=\".csv,.xlsx,.txt\" hidden>${st('factors')}</label>\n      <label class=\"f\"><span class=\"t\"><span class=\"step\">4</span>改善行動 <span class=\"tag\">選用</span></span><span class=\"d\">上月定案的改善行動與進度，會中回顧「做了沒」。</span><span class=\"when\">來源：SharePoint 清單「改善行動」</span><span class=\"pick\"><button type=\"button\" class=\"btn\" data-pick=\"actions\">選擇檔案</button><span class=\"fn\">或把檔案拖到這一格</span></span><input type=\"file\" id=\"f-actions\" accept=\".csv,.xlsx,.txt\" hidden>${st('actions')}</label>\n    </div>\n    <div class=\"row\" style=\"margin-top:16px\"><label class=\"row small\">超長住院門檻（天） <input type=\"text\" id=\"thr\" value=\"${state.thr}\" style=\"width:5em\"></label><button class=\"btn\" id=\"sample\">載入範例資料</button><button class=\"btn\" id=\"clear\">清除已載入資料</button><span class=\"small\" id=\"load-msg\"></span></div>\n    </section>`;\n  for(const k of ['records','drafts','factors','actions']) $('#f-'+k).addEventListener('change',async e=>{ const f=e.target.files[0]; if(!f) return; try{ await load(k,f); $('#load-msg').textContent=''; }catch(err){ $('#load-msg').textContent='讀取失敗：'+err.message; } renderLoad(); const n=$('#next'); if(n) n.scrollIntoView({block:'nearest'}); });\n  el.querySelectorAll('[data-pick]').forEach(b=>b.addEventListener('click',e=>{ e.preventDefault(); $('#f-'+b.dataset.pick).click(); }));\n  el.querySelectorAll('label.f').forEach(card=>{ const inp=card.querySelector('input[type=file]'); ['dragenter','dragover'].forEach(ev=>card.addEventListener(ev,e=>{e.preventDefault();card.classList.add('over');})); ['dragleave','drop'].forEach(ev=>card.addEventListener(ev,e=>{e.preventDefault();card.classList.remove('over');})); card.addEventListener('drop',e=>{ const f=e.dataTransfer.files[0]; if(!f) return; const dt=new DataTransfer(); dt.items.add(f); inp.files=dt.files; inp.dispatchEvent(new Event('change')); }); });\n  el.querySelectorAll('[data-go]').forEach(b=>b.addEventListener('click',()=>{ if(b.dataset.go==='present') state.presIdx=0; setView(b.dataset.go); }));\n  $('#thr').addEventListener('change',e=>{ state.thr=Number(e.target.value)||30; persist(); });\n  $('#sample').addEventListener('click',()=>{ loadSample(); setView('records'); });\n  $('#clear').addEventListener('click',()=>{ state.records=[];state.drafts=[];state.factors=[];state.actions=[];state.loaded={}; try{localStorage.removeItem('mm_dash');}catch(e){} renderLoad(); }); }\n\n/* ---------- 紀錄與分析 ---------- */\nfunction renderRecords(){ const el=$('#v-records'); if(!state.records.length){ el.innerHTML='<section><div class=\"empty\">尚未載入「正式紀錄」。到「載入」頁選檔，或按「載入範例資料」看看畫面。</div></section>'; return; }\n  const ym=state.recYm||nowYm(); const yms=[...new Set([...state.records.map(r=>r.ym).filter(Boolean),nowYm()])].sort().reverse(); const rangeN=state.paretoRange==='all'?null:Number(state.paretoRange); const rYms=rangeN?rangeYms(ym,rangeN):null;\n  const pool=state.records.filter(r=>(!rYms||rYms.includes(r.ym))&&(state.paretoType==='全部'||hasType(r,state.paretoType))); const cs=recIn(ym); const n=t=>cs.filter(c=>hasType(c,t)).length; const unexpected=cs.filter(c=>c.deathExpected==='非預期').length; const prev=cs.filter(c=>['很可能可預防','明確可預防'].includes(cPrev(c))).length; const agreeN=pool.filter(r=>r.agreeMain).length; const rec=recurring(ym); const open=state.open?state.records.find(r=>r.code===state.open):null;\n  el.innerHTML=`<section><div class=\"sec-head\"><h2>正式紀錄</h2><label class=\"row small\">月份 <select id=\"r-ym\">${yms.map(y=>`<option value=\"${y}\" ${y===ym?'selected':''}>${ymLabel(y)}</option>`).join('')}</select></label><span class=\"note\">資料來源：${esc((state.loaded.records||{}).name||'')}</span></div>\n    <div class=\"tiles\"><div class=\"tile\"><div class=\"v\" style=\"color:var(--s1)\">${n('超長住院')}</div><div class=\"l\">超長住院</div><div class=\"sub\">門檻 &gt;${state.thr} 天</div></div><div class=\"tile\"><div class=\"v\" style=\"color:var(--s2)\">${n('Morbidity')}</div><div class=\"l\">Morbidity</div></div><div class=\"tile\"><div class=\"v\" style=\"color:var(--s3)\">${n('Mortality')}</div><div class=\"l\">Mortality</div><div class=\"sub\">其中非預期 ${unexpected}</div></div><div class=\"tile\"><div class=\"v\">${prev}</div><div class=\"l\">很可能／明確可預防</div></div><div class=\"tile\"><div class=\"v\">${pool.length?Math.round(100*agreeN/pool.length)+'%':'—'}</div><div class=\"l\">初評＝共識</div><div class=\"sub\">分析範圍 ${pool.length} 件</div></div><div class=\"tile\" style=\"grid-column:span 2;min-width:0\"><div class=\"v\" style=\"display:flex;flex-wrap:wrap;gap:4px 16px;align-items:baseline\">${DMODE.map(d=>`<span style=\"white-space:nowrap\">${pool.filter(r=>r.decisionMode===d).length}<span class=\"sub\" style=\"font-size:.8rem;font-weight:400;margin-left:4px\">${d}</span></span>`).join('')}</div><div class=\"l\">決策方式</div><div class=\"sub\">分析範圍內</div></div></div>\n    <div id=\"rec-detail\" style=\"margin-top:18px\"></div>\n    ${cs.length?`<div class=\"tbl\" style=\"margin-top:14px\"><table><thead><tr><th>代碼</th><th>類型</th><th>主診斷</th><th class=\"num\">住院天數</th><th>共識主要根因</th><th>可預防性</th><th>決策</th><th>初評一致</th></tr></thead><tbody>${cs.map(r=>`<tr class=\"click ${r.code===state.open?'cur':''}\" data-id=\"${esc(r.code)}\"><td class=\"nw\"><b>${esc(r.code)}</b></td><td>${typePills(r)}</td><td class=\"small\">${esc(r.dxDetail.slice(0,50))}…</td><td class=\"num\">${r.los??''}</td><td>${esc(cMain(r))}</td><td>${esc(cPrev(r))}</td><td class=\"small\">${esc(r.decisionMode)}</td><td>${r.agreeMain?'<span class=\"agree\">一致</span>':'<span class=\"disagree\">不一致</span>'}</td></tr>`).join('')}</tbody></table></div>`:'<div class=\"empty\">本月尚無紀錄。</div>'}</section>\n  <section><div class=\"sec-head\"><h2>分析</h2><div class=\"row small\">範圍 <select id=\"p-range\"><option value=\"3\" ${state.paretoRange==='3'?'selected':''}>近 3 個月</option><option value=\"6\" ${state.paretoRange==='6'?'selected':''}>近 6 個月</option><option value=\"12\" ${state.paretoRange==='12'?'selected':''}>近 12 個月</option><option value=\"all\" ${state.paretoRange==='all'?'selected':''}>全部</option></select> 類型 <select id=\"p-type\"><option>全部</option>${TYPES.map(t=>`<option ${state.paretoType===t?'selected':''}>${t}</option>`).join('')}</select></div></div>\n    <div class=\"charts\"><div class=\"chart\"><h3>每月定案案例數（近 12 個月）</h3><div class=\"legend\">${TYPES.map(t=>`<span><span class=\"dot\" style=\"background:var(${TYPE_VAR[t]})\"></span>${t}</span>`).join('')}</div>${stackedMonthly(rangeYms(ym,12))}</div>\n      <div class=\"chart\"><h3>共識主要根因柏拉圖</h3><div class=\"cap\">右欄為累積百分比。</div>${paretoChart(factorCounts(pool,'main'))}</div>\n      <div class=\"chart\"><h3>主要根因層級 × 案例類型</h3><div class=\"cap\">${state.factors.length?'':'未載入根因分類，層級顯示為未分類。'}</div>${crossTab(pool)}</div>\n      <div class=\"chart\"><h3>可預防性分布（共識）</h3>${prevChart(pool)}</div>\n      <div class=\"chart\"><h3>超長住院天數分布</h3>${losChart(pool.filter(c=>hasType(c,'超長住院')&&c.los))}</div>\n      <div class=\"chart\"><h3>反覆出現的因素</h3><div class=\"cap\">近 6 個月出現於 ≥3 個不同月份。</div>${rec.length?`<div class=\"tbl\"><table><thead><tr><th>因素</th><th class=\"num\">月數</th><th>出現月份</th><th class=\"num\">進行中行動</th></tr></thead><tbody>${rec.map(r=>`<tr><td class=\"flag\">${esc(r.name)}</td><td class=\"num\">${r.months.length}</td><td class=\"small\">${r.months.map(x=>x.slice(2).replace('-','/')).join('、')}</td><td class=\"num\">${state.actions.filter(a=>a.factor===r.name&&!['完成','取消'].includes(a.status)).length||'<span class=\"disagree\">無</span>'}</td></tr>`).join('')}</tbody></table></div>`:'<div class=\"empty\">近 6 個月沒有反覆出現的因素。</div>'}</div></div></section>`;\n  $('#r-ym').addEventListener('change',e=>{state.recYm=e.target.value;state.open=null;renderRecords();}); $('#p-range').addEventListener('change',e=>{state.paretoRange=e.target.value;renderRecords();}); $('#p-type').addEventListener('change',e=>{state.paretoType=e.target.value;renderRecords();});\n  el.querySelectorAll('tr.click').forEach(tr=>tr.addEventListener('click',()=>{state.open=state.open===tr.dataset.id?null:tr.dataset.id;renderRecords();}));\n  if(open){ const k=open.consensus; const acts=state.actions.filter(a=>a.cases.includes(open.code)); $('#rec-detail').innerHTML=`<div class=\"proj\"><div class=\"eyebrow small\">教學案例卡</div><h3>${esc(open.code)}　<span class=\"small\">定案 ${esc(open.meetingDate)}，出席 ${open.attendees} 人${k.voteCount?'，'+k.voteCount+' 票':''}</span></h3>${teachingCard(open)}<div class=\"card\" style=\"margin-top:12px\"><div class=\"kv\"><div class=\"k\">共識主要根因</div><div class=\"v\"><b>${esc(k.mainFactor)}</b> <span class=\"small\">（${esc(facGroup(k.mainFactor))}）</span></div></div><div class=\"kv\"><div class=\"k\">共識促成因素</div><div class=\"v\">${k.factors.join('、')||'—'}</div></div><div class=\"kv\"><div class=\"k\">共識可預防性</div><div class=\"v\">${esc(k.preventability)}</div></div><div class=\"kv\"><div class=\"k\">初評</div><div class=\"v\">${esc(open.initial.mainFactor||'—')}／${esc(open.initial.preventability||'未評')}　${open.agreeMain?'<span class=\"agree\">一致</span>':'<span class=\"disagree\">與共識不同</span>'}</div></div><div class=\"kv\"><div class=\"k\">主治醫師</div><div class=\"v\">${attCell(open)}${open.attending.comment?'：'+esc(open.attending.comment):''}</div></div><div class=\"kv\"><div class=\"k\">決策方式</div><div class=\"v\">${esc(open.decisionMode)}${k.dist?'　'+esc(k.dist):''}</div></div>${kv('長官意見',open.chiefComment,true)}${kv('會議結論',k.notes,true)}<div class=\"kv wide\"><div class=\"k\">改善行動</div><div class=\"v\">${acts.length?acts.map(a=>`${esc(a.title)}（${esc(a.owner||'未指定')}，${esc(a.due||'無到期')}，${esc(a.status)}）`).join('<br>'):'—'}</div></div>${kv('修正附註',open.amendments,true)}</div><div class=\"row\" style=\"margin-top:8px\"><button class=\"btn sm\" id=\"rd-close\">收合</button></div></div>`; $('#rd-close').addEventListener('click',()=>{state.open=null;renderRecords();}); }\n  bindTips(el); }\n\n/* ---------- 待討論案例 ---------- */\nconst ATT_RANK={'不同意初評':0,'未徵詢':1,'同意初評':2};\nfunction queue(){ return state.drafts.filter(d=>d.status==='待討論'||(state.preview&&d.status==='整理中')).map((d,i)=>[d,i]).sort((x,y)=>((ATT_RANK[x[0].attending.status]??1)-(ATT_RANK[y[0].attending.status]??1))||(x[1]-y[1])).map(x=>x[0]); }\nfunction renderPrep(){ const el=$('#v-prep'); const q=queue(); const all=state.drafts; const prep=all.filter(d=>d.status==='整理中').length, done=all.filter(d=>d.status==='已定案').length;\n  const why=!all.length?'還沒有案例。到「載入」頁載入「案例」清單匯出檔（或轉換器產出的 案例_匯入_*.xlsx）。':(prep&&!state.preview?`${prep} 件狀態還是「整理中」：轉換器剛產出、或研究醫師還沒把初評寫完。在 SharePoint 把狀態改成「待討論」再匯出才會排進會議；只是想先看內容，勾上面的「包含整理中」。`:'這個檔裡沒有待討論的案例。');\n  const groups=[['不同意初評','初評與主治不一致 → 逐案討論、投票'],['未徵詢','還沒問主治 → 會中先問意見'],['同意初評','主治同意初評 → 無異議可批次定案']];\n  const grouped=groups.map(([g,d])=>[g,d,q.map((c,i)=>[c,i]).filter(([c])=>(c.attending.status||'未徵詢')===g)]).filter(([, ,L])=>L.length);\n  const row=(d,i)=>`<tr class=\"click\" data-i=\"${i}\"><td class=\"nw\"><b>${esc(d.code)}</b></td><td class=\"small nw\">${esc([d.nameShort,d.age?d.age+'歲':d.ageBand,d.sex].filter(Boolean).join(' '))}</td><td class=\"small nw ${d.status==='整理中'?'muted':''}\">${esc(d.status||'')}</td><td>${typePills(d)}</td><td class=\"small\">${esc([...d.dxDetail].slice(0,60).join(''))}${[...d.dxDetail].length>60?'…':''}</td><td class=\"num\">${d.los??''}</td><td>${esc(d.initial.mainFactor||'—')}<br><span class=\"muted small\">${esc(d.initial.preventability)}</span></td></tr>`;\n  el.innerHTML=`<section><div class=\"sec-head\"><h2>待討論案例</h2><span class=\"note\">待討論 ${all.filter(d=>d.status==='待討論').length} 件、整理中 ${prep} 件、已定案 ${done} 件（全部 ${all.length} 件）。</span><label class=\"row small\" style=\"gap:4px\"><input type=\"checkbox\" id=\"pv-prep\" ${state.preview?'checked':''}>包含整理中（預覽用）</label><button class=\"btn primary sm\" id=\"go-present\" ${q.length?'':'disabled'}>投影模式（從第一案）</button></div>\n    ${q.length?grouped.map(([g,d,L])=>`<h3 style=\"margin:18px 0 2px\">${g==='同意初評'?'<span class=\"agree\">主治同意初評</span>':g==='不同意初評'?'<span class=\"disagree\">主治不同意初評</span>':'未徵詢主治'}　<span class=\"small muted\">${L.length} 件 · ${d}</span></h3><div class=\"tbl\"><table><thead><tr><th>代碼</th><th>病人</th><th>狀態</th><th>類型</th><th>主診斷</th><th class=\"num\">住院天數</th><th>初評主要根因／可預防性</th></tr></thead><tbody>${L.map(([c,i])=>row(c,i)).join('')}</tbody></table></div>`).join('')+`<p class=\"small\" style=\"margin-top:14px\">點任一列從該案開始投影。順序建議：先討論「不同意」與「未徵詢」，最後把「同意初評」整批過一遍定案；投票用 Microsoft Forms 或舉手。</p>`:`<div class=\"empty\">${why}</div>`}</section>`;\n  $('#pv-prep').addEventListener('change',e=>{ state.preview=e.target.checked; renderPrep(); });\n  const gp=$('#go-present'); if(gp) gp.addEventListener('click',()=>{state.presIdx=0;setView('present');});\n  el.querySelectorAll('tr.click').forEach(tr=>tr.addEventListener('click',()=>{state.presIdx=Number(tr.dataset.i);setView('present');})); }\n\n/* ---------- 投影模式（本機翻頁） ---------- */\nfunction presCard(c){ return `<div class=\"card\">${kv('住院天數',c.los!=null?c.los+' 天'+(c.avoidableDays!=null?'（估可避免 '+c.avoidableDays+' 天）':''):'')}${kv('診斷群／治療階段',[c.dxGroup,c.phase].filter(Boolean).join('　'))}${kv('死亡預期性',c.deathExpected&&c.deathExpected!=='不適用'?c.deathExpected:'')}${kv('Morbidity 等級',c.morbGrade&&c.morbGrade!=='不適用'?'Clavien-Dindo '+c.morbGrade:'')}${kvList('診療史',c.treatment)}${kvList('其他診斷・併發症／住院經過',c.course)}${kvList('五個為什麼',c.whys)}${kv('學習重點',c.teachingPoints,true)}${kv('參考指引／文獻',c.refs,true)}</div>`; }\nfunction decideBlock(c){ const i=c.initial,a=c.attending; const att=a.status==='同意初評'?'<span class=\"agree\">同意初評</span>':a.status==='不同意初評'?`<span class=\"disagree\">不同意</span>：${esc(a.mainFactor||'（根因同初評）')}／${esc(a.preventability||'可預防性同初評')}`:'<span class=\"muted\">未徵詢</span>';\n  const rows=[['主要根因',esc(i.mainFactor||'<span class=\"muted\">未評</span>')],['促成因素',i.factors.length?i.factors.map(esc).join('、'):'<span class=\"muted\">—</span>'],['可預防性',esc(i.preventability||'<span class=\"muted\">未評</span>')],['主治醫師',att]];\n  const next=a.status==='同意初評'?'主治同意初評 → 無異議即照初評定案（三方一致）；有人提出異議才投票。':a.status==='不同意初評'?'初評與主治不一致 → 兩方各說一分鐘 → 全場投票或主席裁定。':'尚未徵詢主治 → 先問在場主治意見，再決定是否投票。';\n  return `<div class=\"card\" style=\"margin-top:12px\">${rows.map(([k,v])=>`<div class=\"kv\"><div class=\"k\">${k}</div><div class=\"v\">${v}</div></div>`).join('')}${a.comment?`<div class=\"kv wide\"><div class=\"k\">主治醫師說明</div><div class=\"v\">${esc(a.comment)}</div></div>`:''}</div>\n  <h3 style=\"margin-top:22px\">本案要定的三件事</h3>${SP?'':'<ol class=\"todo\"><li>共識主要根因（可與初評不同）</li><li>可預防性：不可／可能／很可能／明確</li><li>決策方式：三方一致／主席裁定／投票</li></ol>'}<p class=\"small\">${next}</p>${SP?decideForm(c):''}`; }\nfunction decideForm(c){ if(!SP.canDecide) return '<p class=\"small muted\">你沒有寫入「正式紀錄」的權限（需在「M&M 主席」群組），只能檢視。</p>';\n  const i=c.initial, grp=groupsOf(); const opt=(sel)=>grp.map(g=>`<optgroup label=\"${esc(g)}\">${state.factors.filter(f=>f.group===g).map(f=>`<option ${f.name===sel?'selected':''}>${esc(f.name)}</option>`).join('')}</optgroup>`).join('');\n  const mode=c.attending.status==='同意初評'?'三方一致':'投票';\n  return `<form class=\"decide\" id=\"decide\" autocomplete=\"off\">\n    <label>共識主要根因<select name=\"main\" required><option value=\"\">（選擇）</option>${opt(i.mainFactor)}</select></label>\n    <details><summary>共識促成因素（已選 <b id=\"fc\">${i.factors.length}</b>）</summary><div class=\"fac\">${grp.map(g=>`<div class=\"g\">${esc(g)}</div>`+state.factors.filter(f=>f.group===g).map(f=>`<label class=\"ck\"><input type=\"checkbox\" name=\"f\" value=\"${esc(f.name)}\" ${i.factors.includes(f.name)?'checked':''}>${esc(f.name)}</label>`).join('')).join('')}</div></details>\n    <label>共識可預防性<select name=\"prev\" required><option value=\"\">（選擇）</option>${PREV.map(p=>`<option ${p===i.preventability?'selected':''}>${p}</option>`).join('')}</select></label>\n    <div class=\"lbl\">決策方式</div><div class=\"seg\">${DMODE.map(d=>`<label><input type=\"radio\" name=\"mode\" value=\"${d}\" ${d===mode?'checked':''} required>${d}</label>`).join('')}</div>\n    <div class=\"vote\"><label>出席<input type=\"number\" name=\"att\" min=\"0\" inputmode=\"numeric\"></label><label>票數<input type=\"number\" name=\"votes\" min=\"0\" inputmode=\"numeric\"></label><label class=\"w\">投票分布<input type=\"text\" name=\"dist\" placeholder=\"例：同意 8、不同意 2\"></label></div>\n    <label>會議結論<textarea name=\"concl\" rows=\"2\"></textarea></label>\n    <label>長官意見<textarea name=\"chief\" rows=\"2\"></textarea></label>\n    <div class=\"go\"><button class=\"btn primary\" type=\"submit\">定案，寫入正式紀錄</button><span id=\"dmsg\" class=\"small\"></span></div>\n  </form>`; }\nfunction bindDecide(c){ const f=$('#decide'); if(!f) return; const upd=()=>{ $('#fc').textContent=f.querySelectorAll('input[name=f]:checked').length; }; f.querySelectorAll('input[name=f]').forEach(x=>x.addEventListener('change',upd));\n  f.addEventListener('submit',async e=>{ e.preventDefault(); const fd=new FormData(f), btn=f.querySelector('button[type=submit]'), msg=$('#dmsg');\n    const p={ main:fd.get('main'), factors:fd.getAll('f'), prev:fd.get('prev'), mode:fd.get('mode'), attendees:num(fd.get('att')), votes:num(fd.get('votes')), dist:fd.get('dist').trim(), conclusion:fd.get('concl').trim(), chief:fd.get('chief').trim() };\n    if(!p.main||!p.prev||!p.mode){ msg.innerHTML='<span class=\"disagree\">主要根因、可預防性、決策方式都要選。</span>'; return; }\n    btn.disabled=true; msg.textContent='寫入中…';\n    try{ const rec=mapRecord(await SP.decide(c.code,p)); state.records.push(rec); if(rec.ym) state.recYm=rec.ym; const d=state.drafts.find(x=>x.code===c.code); if(d) d.status='已定案'; state.flash=`${c.code} 已定案 ✓（已寫入正式紀錄）`; renderPresent(); }\n    catch(err){ btn.disabled=false; msg.innerHTML=`<span class=\"disagree\">寫入失敗：${esc(err.message)}</span>`; } }); }\nfunction renderPresent(){ const el=$('#v-present'); const q=queue();\n  const ctl=(i,n)=>`<div class=\"ctl\"><button class=\"btn sm\" id=\"pv\" ${i<=0?'disabled':''}>◀ 上一案</button><span class=\"pos\">${n?`${i+1}／${n}`:''}</span><button class=\"btn sm\" id=\"nx\" ${i>=n-1?'disabled':''}>下一案 ▶</button><button class=\"btn sm\" id=\"ex\">離開</button><span class=\"hint\">鍵盤 ← →　Esc 離開</span></div>`;\n  if(!q.length){ const fl=state.flash; state.flash=null; el.innerHTML=`<div class=\"pres\">${fl?`<div class=\"flash\">${esc(fl)}</div>`:''}<div class=\"wait\">沒有待討論案例${SP?'了 — 本次全部定案完成':''}</div>${ctl(0,0)}</div>`; } else { state.presIdx=Math.max(0,Math.min(q.length-1,state.presIdx)); const c=q[state.presIdx];\n    const who=[c.nameShort,c.age?c.age+' 歲':c.ageBand,c.sex].filter(Boolean).join('　');\n    const fl=state.flash; state.flash=null; el.innerHTML=`<div class=\"pres\">${ctl(state.presIdx,q.length)}${fl?`<div class=\"flash\">${esc(fl)}</div>`:''}<div class=\"meta above\">${esc(c.code)}　${esc(ymLabel(c.ym))}　${typePills(c)}${c.status==='整理中'?'　<span class=\"muted\">（整理中・預覽）</span>':''}</div><h2>${esc(who)||esc(c.code)}</h2><p class=\"dx\">${esc(c.dxDetail)||'<span class=\"muted\">（未填主診斷）</span>'}</p><div class=\"cols\"><div>${presCard(c)}</div><div><h3>初評與主治意見</h3>${decideBlock(c)}</div></div></div>`; }\n  if(q.length) bindDecide(q[state.presIdx]); $('#ex').addEventListener('click',()=>setView('prep')); $('#pv').addEventListener('click',()=>{state.presIdx--;renderPresent();}); $('#nx').addEventListener('click',()=>{state.presIdx++;renderPresent();}); }\ndocument.addEventListener('keydown',e=>{ if(state.view!=='present') return; if(/^(INPUT|SELECT|TEXTAREA)$/.test((e.target&&e.target.tagName)||'')) return; if(e.key==='ArrowRight'||e.key==='PageDown'){state.presIdx++;renderPresent();} if(e.key==='ArrowLeft'||e.key==='PageUp'){state.presIdx--;renderPresent();} if(e.key==='Escape') setView('prep'); });\n\n/* ---------- 說明 ---------- */\nfunction renderAbout(){ $('#v-about').innerHTML=`<section><div class=\"sec-head\"><h2>每月怎麼跑</h2><span class=\"note\">四個階段。主要用 SharePoint 書籤「M&amp;M」；這個網頁版是沒有書籤時的備用（要自己匯出 CSV 載入）。</span></div>\n  <div class=\"flow\">\n    <div><div class=\"ph\">① 會前兩週　研究醫師</div><ol><li>SharePoint 站台按書籤「M&amp;M」→ 收案 → 拖入 Word →「寫入」（案例＋對照表一起進清單）。</li><li>在「案例」逐案補：診療史、住院經過、初評主要根因／促成因素／可預防性、五個為什麼；徵詢主治 → 填「主治醫師回覆」。</li><li>完成的改狀態「待討論」。</li></ol></div>\n    <div><div class=\"ph\">② 會前　主席</div><ol><li>站台按「M&amp;M」→ 開會。資料直接讀清單，不用匯出。</li><li>「待討論案例」看三組：不同意／未徵詢／同意。</li></ol></div>\n    <div><div class=\"ph\">③ 會議中　筆電接投影</div><ol><li>「投影模式」，← → 翻案；工具列「全螢幕」。</li><li>右側直接選共識主要根因、可預防性、決策方式 →「定案」：寫進正式紀錄，案例自動改「已定案」。</li><li>有分歧先投票（Forms／舉手），票數填進表單再定案。</li></ol></div>\n    <div><div class=\"ph\">④ 會後</div><ol><li>正式紀錄已在會中寫好，無需再抄。</li><li>需要時新增「改善行動」（負責人、到期日）。</li><li>「紀錄與分析」看趨勢；按「重新讀取」拿最新資料。</li></ol></div>\n  </div>\n  <h3 style=\"margin-top:26px\">資料與隱私</h3>\n  <p class=\"small\">載入的資料只存在這台電腦的瀏覽器裡（下次開啟仍在），「載入」頁的「清除已載入資料」可移除。「案例」已去識別（姓＋OO、年齡），但仍是科內資料，不要在公用電腦載入；「對照表」清單不要匯出、不要載入這裡。</p>\n  <h3 style=\"margin-top:18px\">工具</h3>\n  <p class=\"small\"><a href=\"mm_convert.html\">案例轉換</a>（Word → xlsx，純瀏覽器）　·　<a href=\"handbook.html\">研究醫師手冊</a>　·　<a href=\"bookmarks.html\">書籤安裝</a>　·　<code>create_lists.js</code>（院內保管）</p></section>`; }\n\nif(SP){ spApply(SP.data); const lb=document.querySelector('#nav button[data-view=\"load\"]'); if(lb) lb.remove(); const dp=document.querySelector('.dept'); if(dp) dp.textContent='連線版：資料直接讀自 SharePoint；會中定案寫回「正式紀錄」'; setView('prep'); }\nelse { restore(); (function(){ const h=location.hash.slice(1); setView(VIEWS.includes(h)&&h!=='present'?h:(state.records.length?'records':'load')); })(); }\n";
  async function loadData() {
    const [cl, rl, fl, al] = await Promise.all(['案例', '正式紀錄', '根因分類', '改善行動'].map(listInfo));
    const [cs, rs, fs, as] = await Promise.all([readAll('案例'), readAll('正式紀錄'), readAll('根因分類'), readAll('改善行動')]);
    caseRaw = {}; caseDisp = {}; facIds = {};
    const drafts = cs.map(it => { const o = toDisplay(cl, it); caseRaw[o['標題']] = it; caseDisp[o['標題']] = o; return o; });
    for (const it of fs) facIds[(it.Title || '').trim()] = it.Id;
    return { drafts, records: rs.map(it => toDisplay(rl, it)), factors: fs.map(it => toDisplay(fl, it)), actions: as.map(it => toDisplay(al, it)) };
  }
  async function decide(code, p) {
    const cl = await listInfo('案例'), rl = await listInfo('正式紀錄'), raw = caseRaw[code];
    if (!raw) throw new Error('找不到案例 ' + code);
    const ex = await rq(`web/lists(guid'${rl.id}')/items?$select=Id&$filter=Title eq '${q(code)}'`);
    if (ex.value.length) throw new Error('正式紀錄已經有 ' + code + '，不重複寫入；要修正請到正式紀錄用「修正附註」。');
    const body = { __metadata: { type: rl.type } };
    for (const f of rl.ours) { // 複製案例欄位
      const cf = cl.f[f.InternalName === 'Title' ? '標題' : f.Title]; if (!cf || f.ReadOnlyField) continue;
      const ty = f.TypeAsString, v = /^Lookup/.test(ty) ? raw[cf.InternalName] : raw[cf.EntityPropertyName], k = f.EntityPropertyName;
      if (ty === 'Lookup') { if (v && v.Id) body[k + 'Id'] = v.Id; }
      else if (ty === 'LookupMulti') { const ids = (v || []).map(x => x.Id).filter(Boolean); if (ids.length) body[k + 'Id'] = { __metadata: { type: 'Collection(Edm.Int32)' }, results: ids }; }
      else if (ty === 'MultiChoice') { if (v && v.length) body[k] = { __metadata: { type: 'Collection(Edm.String)' }, results: v }; }
      else if (v != null && v !== '') body[k] = v;
    }
    const F = t => rl.f[t] && rl.f[t].EntityPropertyName, set = (t, v) => { if (F(t) && v != null && v !== '') body[F(t)] = v; };
    const mid = facIds[p.main]; if (!mid) throw new Error('根因分類找不到「' + p.main + '」');
    body[F('共識主要根因') + 'Id'] = mid;
    const fids = p.factors.map(x => facIds[x]).filter(Boolean); if (fids.length) body[F('共識促成因素') + 'Id'] = { __metadata: { type: 'Collection(Edm.Int32)' }, results: fids };
    set('共識可預防性', p.prev); set('決策方式', p.mode); set('會議結論', p.conclusion); set('長官意見', p.chief); set('投票分布', p.dist);
    set('出席人數', p.attendees); set('票數', p.votes);
    const today = ymd(new Date()); set('會議日期', today + 'T00:00:00Z');
    const initMain = caseDisp[code]['初評主要根因'] || ''; set('初評一致', initMain === p.main ? '是' : '否');
    await addItem(rl, body);
    try { await mergeItem(cl, raw.Id, { __metadata: { type: cl.type }, [cl.f['狀態'].EntityPropertyName]: '已定案' }); }
    catch (e) { throw new Error('正式紀錄已寫入，但案例狀態沒改成「已定案」：' + e.message + '（請到案例清單手動改）'); }
    return { ...caseDisp[code], '共識主要根因': p.main, '共識促成因素': p.factors.join(';'), '共識可預防性': p.prev, '決策方式': p.mode, '會議結論': p.conclusion, '長官意見': p.chief, '會議日期': today, '出席人數': p.attendees ?? '', '票數': p.votes ?? '', '投票分布': p.dist, '初評一致': initMain === p.main ? '是' : '否', '狀態': '已定案' };
  }
  async function meetInit() {
    meetReady = true; const P = $('#p-meet'); P.innerHTML = '<div class="load">讀取 SharePoint 清單中…</div>';
    try {
      const [data, can] = await Promise.all([loadData(), canAdd('正式紀錄')]);
      P.innerHTML = '<iframe title="M&M 會議儀表板"></iframe>'; frame = P.querySelector('iframe');
      const w = frame.contentWindow, doc = frame.contentDocument;
      doc.open(); doc.write(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><style>${DASH_CSS}</style></head><body>${DASH_BODY}</body></html>`); doc.close();
      w.MM_SP = { data, canDecide: can, decide, onView: v => { $('#bar').style.display = v === 'present' ? 'none' : ''; } };
      w.eval(DASH_JS);
      const rb = doc.createElement('button'); rb.textContent = '重新讀取'; rb.className = 'btn sm'; rb.style.marginLeft = '10px'; rb.title = '從 SharePoint 重讀最新資料';
      rb.addEventListener('click', async () => { rb.disabled = true; rb.textContent = '讀取中…'; try { w.MM_SP.dashboard.reload(await loadData()); } catch (e) { alert('讀取失敗：' + e.message); } rb.disabled = false; rb.textContent = '重新讀取'; });
      const nav = doc.getElementById('nav'); if (nav) nav.appendChild(rb);
      w.focus();
    } catch (e) { meetReady = false; P.innerHTML = `<div class="w"><p class="bad">讀取失敗：${esc(e.message)}</p><button class="b" id="retry">再試一次</button></div>`; P.querySelector('#retry').onclick = meetInit; }
  }
  // 投影時 Esc 先交給儀表板；儀表板外按 Esc 不關閉，避免誤觸
  window.__MM_APP = { show, version: VERSION };
  let last = 'intake'; try { last = localStorage.getItem('mm-app-tab') || 'intake'; } catch (_) {}
  sh.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => { try { localStorage.setItem('mm-app-tab', b.dataset.t); } catch (_) {} }));
  show(last === 'meet' ? 'meet' : 'intake');
})();
