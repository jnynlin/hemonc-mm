/* 血腫科 M&M — SharePoint 站台內工具（收案＋開會）  __VERSION__
   由書籤載入：書籤從本站「文件」庫讀 mm_app.js 後執行。資料只在瀏覽器 ↔ 本站 SharePoint 之間，不經其他伺服器。
   本檔由 build_app.py 產生：轉換核心取自 mm_convert.html、儀表板取自 mm_dashboard.html，請勿手改。 */
(() => {
  'use strict';
  if (window.__MM_APP) { window.__MM_APP.show(); return; }
  const VERSION = '__VERSION__';
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
  __CORE__

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
  const DASH_CSS = __DASH_CSS__, DASH_BODY = __DASH_BODY__, DASH_JS = __DASH_JS__;
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
