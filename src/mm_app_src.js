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
    if (!r.ok) throw new Error('連線逾時，請重新整理頁面後再按一次「M&M」（代碼 ' + r.status + '）');
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
        case 'Choice': if (ch.length && !ch.includes(v)) warn.push(`${h}「${v}」不是清單選項`); else body[k] = v; break;
        case 'MultiChoice': { const vals = split(v).filter(x => !ch.length || ch.includes(x)); if (vals.length) body[k] = { __metadata: { type: 'Collection(Edm.String)' }, results: vals }; break; }
        case 'DateTime': { const d = new Date(v.replace(/[./]/g, '-')); if (!isNaN(d)) body[k] = ymd(d) + 'T00:00:00Z'; break; }
        case 'Lookup': { const id = lookupIds && lookupIds[v]; if (id) body[k + 'Id'] = id; else warn.push(`${h}「${v}」不在根因分類中`); break; }
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
<div class="body"><div class="pane" id="p-intake"></div><div class="pane" id="p-meet" hidden><div class="load">讀取中…</div></div></div></div>`;
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
      P.innerHTML = `<div class="w"><h2>收案</h2><div class="banner">你的帳號目前還不能收案${okCase ? '（尚未加入「M&amp;M 對照表」群組）' : ''}。<br>請主任開通權限後，重新整理頁面再試。「開會」頁不受影響，可以照常使用。</div></div>`;
      return;
    }
    P.innerHTML = `<div class="w"><h2>收案</h2>
      <p class="note">把研究醫師整理的討論會 Word 檔拖進來。系統會隱去病歷號、全名與床號，編好案例代碼，確認後寫入「案例」清單；可識別資料另存在只有主任與研究醫師看得到的「對照表」。Word 檔不會離開這台電腦。</p>
      <div class="drop" id="drop" tabindex="0"><b id="dt">把 Word 檔拖到這裡，或點一下選擇（可一次多個）</b><div class="note" id="ds">超長住院、Morbidity、Mortality 討論會清單皆可</div><input type="file" id="file" accept=".docx" multiple hidden></div>
      <div class="row"><span id="ymw">討論年月 <input type="text" id="ym" placeholder="2026-10"> <span class="note" id="ymh">檔名沒有年月時才需要填。</span></span>
        <label class="ck"><input type="checkbox" id="demo">演練模式（寫入的資料可一次刪除，不影響正式案例）</label></div>
      <div id="res"></div></div>`;
    let files = [], R = null;
    const drop = $('#drop'), fileIn = $('#file'), ymIn = $('#ym');
    const isMortF = f => /mortality|死亡/i.test(f.name);
    const ymFromName = f => { const m = f.name.match(/(\d{4})年\s*(\d{1,2})月/); return m ? m[1] + '-' + m[2].padStart(2, '0') : ''; };
    async function go() {
      const res = $('#res'); if (!files.length) return;
      const needYm = files.some(f => !isMortF(f) && !ymFromName(f));
      $('#ymw').style.display = needYm ? '' : 'none';
      const shared = ymIn.value.trim();
      if (needYm && !/^\d{4}-(0[1-9]|1[0-2])$/.test(shared)) { res.innerHTML = '<p class="warn">有檔案看不出是哪個月份，請填討論年月（例如 2026-10）。</p>'; return; }
      res.innerHTML = '<div class="load">整理中，並檢查是否已經收過…</div>';
      try { R = await plan(files.map(f => ({ file: f, ym: isMortF(f) ? null : (ymFromName(f) || shared) })), $('#demo').checked); render(); }
      catch (e) { res.innerHTML = `<p class="bad">處理失敗：${esc(e.message)}</p>`; }
    }
    function pick(list) {
      const fs = [...(list || [])].filter(f => /\.docx$/i.test(f.name));
      if (!fs.length) { $('#ds').innerHTML = '<span class="bad">請選擇 Word 檔（.docx）</span>'; return; }
      files = fs; drop.classList.add('has'); $('#dt').textContent = fs.map(f => f.name).join('、');
      $('#ds').textContent = `已選 ${fs.length} 個檔，點這裡可重新選擇`; go();
    }
    drop.addEventListener('click', () => fileIn.click());
    fileIn.addEventListener('change', () => pick(fileIn.files));
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => pick(e.dataTransfer.files));
    ymIn.addEventListener('change', go); $('#demo').addEventListener('change', go);

    function render() {
      const res = $('#res'), d = R.items, n = d.filter(x => x.ok).length, dup = d.filter(x => x.dup).length;
      const unG = d.filter(x => x.ok && !x.draft.dxGroup).length;
      res.innerHTML = `${R.groups.map(g => g.leak.length
          ? `<p class="bad">${esc(g.name)}：內容裡還有疑似病歷號的長串數字（${esc(g.leak.join('、'))}），這個檔不會寫入，請告訴主任。</p>`
          : `<h3>${esc(g.ctype)} ${g.n} 例，可識別資料已隱去${g.dup ? `　<span class="warn">其中 ${g.dup} 例之前已收過，會略過</span>` : ''}<br><span class="note">${esc(g.name)}</span></h3>`).join('')}
        ${unG ? `<p class="note warn">${unG} 例無法自動判斷診斷群，寫入後請在清單補選。</p>` : ''}
        ${R.demo ? '<p class="note">演練模式：寫入的案例代碼以 TEST- 開頭，練習完按最下方「刪除演練資料」。</p>' : ''}
        <div class="row"><button class="b p" id="write" ${n ? '' : 'disabled'}>寫入勾選的 <span id="wn">${n}</span> 例</button><span class="note">每一例會同時寫入「案例」與「對照表」；任一邊沒寫成功，該例會整筆取消。</span></div>
        <div class="tw"><table><tr><th></th><th>代碼</th><th>病人</th><th>診斷群</th><th>天數</th><th>主診斷<small>→ 詳細診斷</small></th><th>診療史<small>→ 治療內容</small></th><th>其他診斷・併發症<small>→ 臨床經過</small></th></tr>
        ${R.groups.map(g => `<tr><td colspan="8" style="background:var(--soft);font-weight:600">${esc(g.ctype)}　<span class="note">${esc(g.name)}</span></td></tr>` + d.map((x, i) => [x, i]).filter(([x]) => x.g === g).map(([x, i]) => `<tr class="${x.ok ? '' : 'dup'}"><td><input type="checkbox" data-i="${i}" ${x.ok ? 'checked' : 'disabled title="不寫入"'}></td><td class="c">${esc(x.code)}${x.dup ? `<br><span class="warn">已收過：${esc(x.dup)}</span>` : ''}</td><td>${esc(x.draft.nameShort)} ${esc(x.draft.age ?? '?')}${esc(x.draft.sex)}</td><td>${esc(x.draft.dxGroup) || '<span class="warn">未歸類</span>'}</td><td>${esc(x.draft.los ?? '')}</td><td class="x">${clip(x.draft.dxDetail, 46)}</td><td class="x">${x.draft.treatment ? clip(x.draft.treatment, 38) : '—'}</td><td class="x">${x.draft.course ? clip(x.draft.course.replace(/｜/g, ' ｜ '), 38) : '—'}</td></tr>`).join('')).join('')}</table></div>
        <pre class="log" id="log" hidden></pre>
        <h3>演練資料</h3><p class="note">刪除所有 TEST- 開頭的練習案例（含對照表）。正式案例不受影響。</p><button class="b" id="purge">刪除演練資料</button>`;
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
          catch (e) { await delItem(cl, caseId).catch(() => {}); throw new Error('對照表沒寫成功，這一例已整筆取消：' + e.message); }
          ok++; say(`${x.code} 已寫入${c.warn.length ? '（未填：' + c.warn.join('；') + '）' : ''}`);
        } catch (e) { fail++; say(`${x.code} 未寫入：${e.message}`); }
      }
      say(`完成：新增 ${ok} 例${fail ? `，${fail} 例未寫入（原因見上方）` : ''}。`);
      if (ok) say(`接下來請到「案例」清單逐案填寫初評、徵詢主治，完成後把狀態改成「待討論」。\n${cl.url}`);
      R = null; files = []; drop.classList.remove('has'); $('#dt').textContent = '把 .docx 拖到這裡（可一次多個），或點一下選檔'; $('#ds').textContent = '已完成。要收下一批，再把檔案拖進來。';
      meetReady = false; // 開會頁下次重讀
    }
    async function purge() {
      const res = $('#res'); let log = $('#log'); if (!log) { res.insertAdjacentHTML('beforeend', '<pre class="log" id="log"></pre>'); log = $('#log'); }
      log.hidden = false; log.textContent = '刪除中…\n';
      let n = 0;
      for (const t of ['案例', '對照表']) {
        const li = await listInfo(t); const j = await rq(`web/lists(guid'${li.id}')/items?$top=5000&$select=Id,Title&$filter=startswith(Title,'TEST-')`);
        for (const it of j.value) { await delItem(li, it.Id); if (t === '案例') n++; }
      }
      log.textContent += `已刪除 ${n} 例練習資料。`; meetReady = false;
    }
  }
  // 轉換 → 依 SharePoint 既有資料重新編號、標出重複
  async function plan(specs, demo) {
    const cl = await listInfo('案例'), kl = await listInfo('對照表'), chartF = kl.f['病歷號'], E = t => cl.f[t].EntityPropertyName;
    const cases = (await rq(`web/lists(guid'${cl.id}')/items?$top=5000&$select=Title,${E('病人簡稱')},${E('年齡')},${E('詳細診斷')}`)).value;
    const titles = cases.map(x => x.Title || '');
    const fp = (stem, s, a, dx) => [stem.replace(/^TEST-/, ''), s || '', a ?? '', String(dx || '').slice(0, 40)].join('|');
    const byFp = {}; for (const c of cases) { const m = (c.Title || '').match(/^(.*-\d{4}-\d{2}-)\d+$/); if (m) byFp[fp(m[1], c[E('病人簡稱')], c[E('年齡')], c[E('詳細診斷')])] = c.Title; }
    const keys = (await rq(`web/lists(guid'${kl.id}')/items?$top=5000&$select=Title,${chartF.EntityPropertyName}`)).value;
    const byChart = {}; for (const k of keys) { const c = String(k[chartF.EntityPropertyName] || '').trim(); if (c) (byChart[c] = byChart[c] || []).push(k.Title || ''); }
    const pre = demo ? 'TEST-' : '', next = {}, groups = [], items = [];
    const obj = (cols, row) => { const o = {}; cols.forEach((c, j) => o[c] = row[j]); return o; };
    for (const sp of specs) {
      const r = await MM.convert(await sp.file.arrayBuffer(), sp.ym);
      const g = { name: sp.file.name, ctype: r.ctype, leak: r.leak, n: r.drafts.length, dup: 0 }; groups.push(g);
      r.drafts.forEach((d, i) => {
        const prefix = d.code.replace(/-\d{4}-\d{2}-\d+$/, ''), stem = `${pre}${prefix}-${d.ym}-`, chart = String(r.keyRows[i][1] || '').trim();
        let dup = (byChart[chart] || []).find(t => t.startsWith(stem) && /^\d+$/.test(t.slice(stem.length))) || null;
        if (!dup) { const hit = byFp[fp(stem, d.nameShort, d.age, d.dxDetail)]; if (hit && hit.startsWith(stem)) dup = hit; } // 案例在、對照表沒有的舊資料
        let code = dup;
        if (!dup) {
          if (next[stem] == null) next[stem] = Math.max(0, ...titles.filter(t => t.startsWith(stem)).map(t => parseInt(t.slice(stem.length), 10)).filter(n => !isNaN(n)));
          code = stem + String(++next[stem]).padStart(2, '0');
          if (chart) (byChart[chart] = byChart[chart] || []).push(code); // 同一批兩個檔重複的病人
        } else g.dup++;
        const caseRow = obj(MM.LIST_COLS, r.caseRows[i]), keyRow = obj(MM.KEY_COLS, r.keyRows[i]);
        caseRow['標題'] = code; keyRow['標題'] = code;
        items.push({ g, code, dup, ok: !dup && !r.leak.length, draft: d, caseRow, keyRow });
      });
    }
    return { demo, groups, items };
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
    if (ex.value.length) throw new Error(code + ' 已經定案過。如需修正，請到「正式紀錄」該筆的「修正附註」補充。');
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
    const mid = facIds[p.main]; if (!mid) throw new Error('找不到根因「' + p.main + '」，請按「重新讀取」後再試。');
    body[F('共識主要根因') + 'Id'] = mid;
    const fids = p.factors.map(x => facIds[x]).filter(Boolean); if (fids.length) body[F('共識促成因素') + 'Id'] = { __metadata: { type: 'Collection(Edm.Int32)' }, results: fids };
    set('共識可預防性', p.prev); set('決策方式', p.mode); set('會議結論', p.conclusion); set('長官意見', p.chief); set('投票分布', p.dist);
    set('出席人數', p.attendees); set('票數', p.votes);
    const today = ymd(new Date()); set('會議日期', today + 'T00:00:00Z');
    const initMain = caseDisp[code]['初評主要根因'] || ''; set('初評一致', initMain === p.main ? '是' : '否');
    await addItem(rl, body);
    try { await mergeItem(cl, raw.Id, { __metadata: { type: cl.type }, [cl.f['狀態'].EntityPropertyName]: '已定案' }); }
    catch (e) { throw new Error('定案已寫入，但案例狀態沒有更新。請到「案例」清單把這一例改成「已定案」。（' + e.message + '）'); }
    return { ...caseDisp[code], '共識主要根因': p.main, '共識促成因素': p.factors.join(';'), '共識可預防性': p.prev, '決策方式': p.mode, '會議結論': p.conclusion, '長官意見': p.chief, '會議日期': today, '出席人數': p.attendees ?? '', '票數': p.votes ?? '', '投票分布': p.dist, '初評一致': initMain === p.main ? '是' : '否', '狀態': '已定案' };
  }
  async function meetInit() {
    meetReady = true; const P = $('#p-meet'); P.innerHTML = '<div class="load">讀取中…</div>';
    try {
      const [data, can] = await Promise.all([loadData(), canAdd('正式紀錄')]);
      P.innerHTML = '<iframe title="M&M 會議儀表板"></iframe>'; frame = P.querySelector('iframe');
      const w = frame.contentWindow, doc = frame.contentDocument;
      doc.open(); doc.write(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><style>${DASH_CSS}</style></head><body>${DASH_BODY}</body></html>`); doc.close();
      w.MM_SP = { data, canDecide: can, decide, onView: v => { $('#bar').style.display = v === 'present' ? 'none' : ''; } };
      w.eval(DASH_JS);
      const rb = doc.createElement('button'); rb.textContent = '重新讀取'; rb.className = 'btn sm'; rb.style.marginLeft = '10px'; rb.title = '讀取清單的最新內容';
      rb.addEventListener('click', async () => { rb.disabled = true; rb.textContent = '讀取中…'; try { w.MM_SP.dashboard.reload(await loadData()); } catch (e) { alert('讀取失敗，請稍後再試：' + e.message); } rb.disabled = false; rb.textContent = '重新讀取'; });
      const nav = doc.getElementById('nav'); if (nav) nav.appendChild(rb);
      w.focus();
    } catch (e) { meetReady = false; P.innerHTML = `<div class="w"><p class="bad">讀取清單失敗，請稍後再試。（${esc(e.message)}）</p><button class="b" id="retry">再試一次</button></div>`; P.querySelector('#retry').onclick = meetInit; }
  }
  // 投影時 Esc 先交給儀表板；儀表板外按 Esc 不關閉，避免誤觸
  window.__MM_APP = { show, version: VERSION };
  let last = 'intake'; try { last = localStorage.getItem('mm-app-tab') || 'intake'; } catch (_) {}
  sh.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => { try { localStorage.setItem('mm-app-tab', b.dataset.t); } catch (_) {} }));
  show(last === 'meet' ? 'meet' : 'intake');
})();
