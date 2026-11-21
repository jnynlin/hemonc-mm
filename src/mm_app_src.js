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
  const wait = ms => new Promise(r => setTimeout(r, ms));
  async function rq(path, { method = 'GET', body = null, headers = {} } = {}) {
    const url = /^https?:/.test(path) || path.includes('/_api/') ? path : base + '/_api/' + path;
    const xm = headers['X-HTTP-Method'] || '', idem = method === 'GET' || xm === 'DELETE' || xm === 'MERGE';
    for (let attempt = 0; ; attempt++) {
      const h = { Accept: method === 'GET' ? 'application/json;odata=nometadata' : 'application/json;odata=verbose', ...headers };
      if (method !== 'GET') { h['X-RequestDigest'] = await getDigest(); if (body) h['Content-Type'] = 'application/json;odata=verbose'; }
      let r;
      try { r = await fetch(url, { method, headers: h, body: body ? JSON.stringify(body) : null }); }
      catch (e) { if (idem && attempt < 3) { await wait(1000 * (attempt + 1)); continue; } throw new Error('網路連線中斷，請稍後再試'); }
      // SharePoint 流量限制（429／503）一律等候後重試；讀取與刪除、修改遇到暫時性錯誤也重試
      const retryable = r.status === 429 || r.status === 503 || (idem && [500, 502, 504].includes(r.status));
      if (retryable && attempt < 6) { const ra = parseInt(r.headers.get('Retry-After') || '0', 10); await wait(Math.min(30000, (ra > 0 ? ra * 1000 : 1000 * 2 ** attempt))); continue; }
      if (r.status === 204) return null;
      const t = await r.text(); let j = null; try { j = t ? JSON.parse(t) : null; } catch (_) { j = { raw: t }; }
      if (!r.ok) { const m = (j && ((j['odata.error'] && j['odata.error'].message.value) || (j.error && j.error.message && j.error.message.value))) || t.slice(0, 200); const e = new Error(m); e.status = r.status; throw e; }
      return j && j.d !== undefined ? j.d : j;
    }
  }
  async function getAll(path) { // 跟著 nextLink 讀完全部（清單超過 5000 筆時分頁）
    let url = path, out = [];
    while (url) { const j = await rq(url); out = out.concat(j.value || []); url = j['odata.nextLink'] || null; }
    return out;
  }
  const q = s => encodeURIComponent(String(s).replace(/'/g, "''"));
  const LI = {};
  async function listInfo(title) {
    if (LI[title]) return LI[title];
    const BYURL = { '投票': '/Lists/mmvote' }; // 用網址找：多語系站台改名後，依名稱可能找不到
    const sel = '$select=Id,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder';
    const l = BYURL[title] ? await rq(`web/GetList(@u)?${sel}&@u='${encodeURIComponent(base.replace(location.origin, '') + BYURL[title])}'`) : await rq(`web/lists/getbytitle('${q(title)}')?${sel}`);
    const fs = (await rq(`web/lists(guid'${l.Id}')/fields?$filter=Hidden eq false`)).value;
    const ours = fs.filter(f => f.InternalName === 'Title' || /^mm/.test(f.InternalName));
    const f = {}; for (const x of ours) f[x.Title] = x; const tf = ours.find(x => x.InternalName === 'Title'); if (tf) f['標題'] = tf;
    return (LI[title] = { title, id: l.Id, type: l.ListItemEntityTypeFullName, url: location.origin + l.RootFolder.ServerRelativeUrl, f, ours });
  }
  async function canAdd(title) {
    try { const li = await listInfo(title); const p = await rq(`web/lists(guid'${li.id}')/EffectiveBasePermissions`); return (Number(p.Low) & 2) !== 0; }
    catch (_) { return false; }
  }
  async function canEdit(title) { // 能不能修改清單項目（EditListItems）
    try { const li = await listInfo(title); const p = await rq(`web/lists(guid'${li.id}')/EffectiveBasePermissions`); return (Number(p.Low) & 4) !== 0; }
    catch (_) { return false; }
  }
  function selExp(li) { const sel = ['Id'], exp = []; for (const f of li.ours) { if (/^(Lookup|User)/.test(f.TypeAsString)) { sel.push(f.InternalName + '/Title', f.InternalName + '/Id'); exp.push(f.InternalName); } else sel.push(f.EntityPropertyName); } return `$select=${sel.join(',')}${exp.length ? '&$expand=' + exp.join(',') : ''}`; }
  async function readOne(title, id) { const li = await listInfo(title); return rq(`web/lists(guid'${li.id}')/items(${id})?${selExp(li)}`); }
  async function readAll(title) {
    const li = await listInfo(title); const sel = ['Id'], exp = [];
    for (const f of li.ours) { if (/^(Lookup|User)/.test(f.TypeAsString)) { sel.push(f.InternalName + '/Title', f.InternalName + '/Id'); exp.push(f.InternalName); } else sel.push(f.EntityPropertyName); }
    let url = `web/lists(guid'${li.id}')/items?$top=5000&$select=${sel.join(',')}${exp.length ? '&$expand=' + exp.join(',') : ''}`, out = [];
    while (url) { const j = await rq(url); out = out.concat(j.value || []); url = j['odata.nextLink'] || null; }
    return out;
  }
  const ymd = v => { const d = new Date(v); if (isNaN(d)) return ''; return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  function toDisplay(li, it) {
    const o = { _id: it.Id };
    for (const f of li.ours) {
      const t = f.InternalName === 'Title' ? '標題' : f.Title, ty = f.TypeAsString;
      let v = /^(Lookup|User)/.test(ty) ? it[f.InternalName] : it[f.EntityPropertyName];
      if (ty === 'Lookup' || ty === 'User') v = v && v.Title ? v.Title : '';
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
        case 'User': { const id = parseInt(v, 10); if (id) body[k + 'Id'] = id; break; }
        case 'Lookup': { const id = lookupIds && lookupIds[v]; if (id) body[k + 'Id'] = id; else warn.push(`${h}「${v}」不在根因分類中`); break; }
        case 'LookupMulti': { const ids = split(v).map(x => lookupIds && lookupIds[x]).filter(Boolean); if (ids.length) body[k + 'Id'] = { __metadata: { type: 'Collection(Edm.Int32)' }, results: ids }; break; }
        default: body[k] = v;
      }
    }
    return { body, warn };
  }
  const userCache = {};
  async function resolveUser(name) { // Word 上的主治姓名 → 站台使用者 Id；同名或找不到回傳 null
    name = String(name || '').split(/[／/、,，\s]/).filter(Boolean)[0] || '';
    if (!name) return { id: null, why: 'Word 沒寫主治' };
    if (userCache[name]) return userCache[name];
    let out = { id: null, why: '找不到「' + name + '」' };
    try {
      const r = await rq('SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser', { method: 'POST', body: { queryParams: { __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' }, AllowEmailAddresses: true, AllowMultipleEntities: false, AllUrlZones: false, MaximumEntitySuggestions: 10, PrincipalSource: 15, PrincipalType: 1, QueryString: name } } });
      const hits = JSON.parse(r.ClientPeoplePickerSearchUser || '[]').filter(e => String(e.DisplayText || '').startsWith(name));
      if (hits.length > 1) out = { id: null, why: '「' + name + '」有 ' + hits.length + ' 位同名' };
      else if (hits.length === 1) { const u = await rq('web/ensureuser', { method: 'POST', body: { logonName: hits[0].Key } }); out = { id: u.Id, name: hits[0].DisplayText }; }
    } catch (e) { out = { id: null, why: '查詢帳號失敗' }; }
    return (userCache[name] = out);
  }
  const addItem = (li, body) => rq(`web/lists(guid'${li.id}')/items`, { method: 'POST', body });
  const delItem = (li, id) => rq(`web/lists(guid'${li.id}')/items(${id})`, { method: 'POST', headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' } });
  const mergeItem = (li, id, body) => rq(`web/lists(guid'${li.id}')/items(${id})`, { method: 'POST', body, headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' } });

  /* ================= 轉換核心（mm_convert.html） ================= */
  __CORE__

  /* ================= QR code（投票用） ================= */
  __QR__

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
      <p class="note">把討論會的 Word 檔拖進來，確認內容後寫入「案例」清單。「案例」只顯示姓氏與年齡；病歷號、姓名、床號另存於「對照表」，僅主任與研究醫師可查。Word 檔不會上傳。</p>
      <div class="drop" id="drop" tabindex="0"><b id="dt">拖曳 Word 檔到這裡，或點選檔案</b><div class="note" id="ds">可一次選多個檔；超長住院、Morbidity、Mortality 均可</div><input type="file" id="file" accept=".docx" multiple hidden></div>
      <div class="row"><span id="ymw" style="display:none">討論年月 <input type="text" id="ym" placeholder="例：2026-10"> <span class="note" id="ymh">檔名中找不到月份，請填寫。</span></span>
</div>
      <div id="testbox"></div><div id="attbox"></div>
      <div id="res"></div></div>`;
    let files = [], R = null;
    const drop = $('#drop'), fileIn = $('#file'), ymIn = $('#ym');
    const isMortF = f => /mortality|死亡/i.test(f.name);
    const ymFromName = f => { const m = f.name.match(/(\d{4})年\s*(\d{1,2})月/); return m ? m[1] + '-' + m[2].padStart(2, '0') : ''; };
    async function go() {
      const res = $('#res'); if (!files.length) return;
      for (const f of files) { // 先確認是真的 Word 檔（.docx 是 zip，開頭為 PK）
        const sig = new Uint8Array(await f.slice(0, 4).arrayBuffer());
        if (f.size === 0 || sig[0] !== 0x50 || sig[1] !== 0x4B) { res.innerHTML = `<p class="bad">${esc(f.name)}：不是有效的 Word 檔。請用 Word 開啟後「另存新檔」成 .docx 再試。</p>`; return; }
      }
      const needYm = files.some(f => !isMortF(f) && !ymFromName(f));
      $('#ymw').style.display = needYm ? '' : 'none';
      const shared = ymIn.value.trim();
      if (needYm && !/^\d{4}-(0[1-9]|1[0-2])$/.test(shared)) { res.innerHTML = '<p class="warn">請在上方填寫討論年月。</p>'; return; }
      res.innerHTML = '<div class="load">處理中…</div>';
      try { R = await plan(files.map(f => ({ file: f, ym: isMortF(f) ? null : (ymFromName(f) || shared) })), false); render(); }
      catch (e) { res.innerHTML = `<p class="bad">處理失敗：${esc(e.message)}</p>`; }
    }
    function pick(list) {
      const fs = [...(list || [])].filter(f => /\.docx$/i.test(f.name));
      if (!fs.length) { $('#ds').innerHTML = '<span class="bad">請選擇 Word 檔（.docx）</span>'; return; }
      files = fs; drop.classList.add('has'); $('#dt').textContent = fs.map(f => f.name).join('、');
      $('#ds').textContent = `已選 ${fs.length} 個檔（點此可重新選擇）`; go();
    }
    drop.addEventListener('click', () => fileIn.click());
    fileIn.addEventListener('change', () => pick(fileIn.files));
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => pick(e.dataTransfer.files));
    ymIn.addEventListener('change', go);

    function render() {
      const res = $('#res'), d = R.items, n = d.filter(x => x.ok).length, dup = d.filter(x => x.dup).length;
      const unG = d.filter(x => x.ok && !x.draft.dxGroup).length;
      res.innerHTML = `${R.groups.map(g => g.bad
          ? `<p class="bad">${esc(g.name)}：Word 格式和以往不同（${esc(g.bad)}），這個檔不會寫入。請對照上個月的 Word，確認表格欄位順序與寫法後再收。</p>`
          : g.leak.length
          ? `<p class="bad">${esc(g.name)}：內容裡還有疑似病歷號的長串數字（${esc(g.leak.join('、'))}），這個檔不會寫入，請告訴主任。</p>`
          : `<h3>${esc(g.ctype)} ${g.n} 例，可識別資料已隱去${g.dup ? `　<span class="warn">其中 ${g.dup} 例之前已收過，會略過</span>` : ''}<br><span class="note">${esc(g.name)}</span></h3>`).join('')}
        ${unG ? `<p class="note warn">${unG} 例無法自動判斷診斷群，寫入後請在清單補選。</p>` : ''}
        ${(() => { const n = d.filter(x => x.ok && x.att && !x.att.id).length; return n ? `<p class="note warn">${n} 例沒有自動帶入主治醫師，寫入後請在清單的「主治醫師」欄選取。</p>` : ''; })()}
        <div class="row"><button class="b p" id="write" ${n ? '' : 'disabled'}>寫入勾選的 <span id="wn">${n}</span> 例</button><span class="note">每一例會同時寫入「案例」與「對照表」；任一邊沒寫成功，該例會整筆取消。</span></div>
        <div class="tw"><table><tr><th></th><th>代碼</th><th>病人</th><th>主治</th><th>診斷群</th><th>天數</th><th>主診斷<small>→ 詳細診斷</small></th><th>診療史<small>→ 治療內容</small></th><th>其他診斷・併發症<small>→ 臨床經過</small></th></tr>
        ${R.groups.map(g => `<tr><td colspan="9" style="background:var(--soft);font-weight:600">${esc(g.ctype)}　<span class="note">${esc(g.name)}</span></td></tr>` + d.map((x, i) => [x, i]).filter(([x]) => x.g === g).map(([x, i]) => `<tr class="${x.ok ? '' : 'dup'}"><td><input type="checkbox" data-i="${i}" ${x.ok ? 'checked' : 'disabled title="不寫入"'}></td><td class="c">${esc(x.code)}${x.dup ? `<br><span class="warn">已收過：${esc(x.dup)}</span>` : ''}${x.noDate ? '<br><span class="warn">找不到日期，不寫入</span>' : ''}</td><td>${esc(x.draft.nameShort)} ${esc(x.draft.age ?? '?')}${esc(x.draft.sex)}</td><td>${!x.att ? '—' : x.att.id ? esc(String(x.att.name).split(' ')[0]) : `<span class="warn">${esc(x.att.why)}</span>`}</td><td>${esc(x.draft.dxGroup) || '<span class="warn">未歸類</span>'}</td><td>${esc(x.draft.los ?? '')}</td><td class="x">${clip(x.draft.dxDetail, 46)}</td><td class="x">${x.draft.treatment ? clip(x.draft.treatment, 38) : '—'}</td><td class="x">${x.draft.course ? clip(x.draft.course.replace(/｜/g, ' ｜ '), 38) : '—'}</td></tr>`).join('')).join('')}</table></div>
        <pre class="log" id="log" hidden></pre>
`;
      const cnt = () => { $('#wn').textContent = res.querySelectorAll('input[data-i]:checked').length; };
      res.querySelectorAll('input[data-i]').forEach(c => c.addEventListener('change', cnt));
      $('#write').addEventListener('click', () => write([...res.querySelectorAll('input[data-i]:checked')].map(c => R.items[+c.dataset.i])));
    }
    async function write(sel) {
      const log = $('#log'), btn = $('#write'); log.hidden = false; log.textContent = ''; btn.disabled = true;
      const say = s => { log.textContent += s + '\n'; log.scrollTop = log.scrollHeight; };
      const cl = await listInfo('案例'), kl = await listInfo('對照表'); let ok = 0, fail = 0; const orphan = [];
      for (const x of sel) {
        let caseId = null;
        try {
          const c = toBody(cl, x.caseRow); const k = toBody(kl, x.keyRow);
          const cr = await addItem(cl, c.body); caseId = cr.Id;
          try { await addItem(kl, k.body); }
          catch (e) { let undone = true; await delItem(cl, caseId).catch(() => { undone = false; }); if (!undone) orphan.push(x.code); throw new Error((undone ? '對照表沒寫成功，這一例已整筆取消：' : '對照表沒寫成功，而且案例沒能自動取消，請到「案例」清單刪除這一例：') + e.message); }
          ok++; say(`${x.code} 已寫入${c.warn.length ? '（未填：' + c.warn.join('；') + '）' : ''}`);
        } catch (e) { fail++; say(`${x.code} 未寫入：${e.message}`); }
      }
      say(`完成：新增 ${ok} 例${fail ? `，${fail} 例未寫入（原因見上方）` : ''}。`);
      if (orphan.length) say(`請注意：${orphan.join('、')} 只寫入了案例、沒有對照表，請到「案例」清單刪除後重新收案。`);
      if (fail) say('未寫入的例子可以直接再拖一次同一個 Word 檔：已寫入的會標「已收過」略過，只補寫未寫入的。');
      if (ok) say(`接下來請到「案例」清單逐案填寫初評、徵詢主治，完成後把狀態改成「待討論」。\n${cl.url}`);
      R = null; files = []; drop.classList.remove('has'); $('#dt').textContent = '拖曳 Word 檔到這裡，或點選檔案'; $('#ds').textContent = '已完成。如要收下一批，請再拖入檔案。';
      meetReady = false; // 開會頁下次重讀
    }
    async function purge() {
      const box = $('#testbox'); box.innerHTML = '<p class="note">刪除中…</p>'; let n = 0;
      try {
        for (const t of ['案例', '對照表']) {
          const li = await listInfo(t); const all = await getAll(`web/lists(guid'${li.id}')/items?$top=5000&$select=Id,Title`);
          for (const it of all.filter(x => /^TEST-/.test(x.Title || ''))) { await delItem(li, it.Id); if (t === '案例') n++; }
        }
        box.innerHTML = `<p class="note">已刪除 ${n} 例測試資料。</p>`; meetReady = false;
      } catch (e) { box.innerHTML = `<p class="bad">刪除失敗：${esc(e.message)}</p>`; }
    }
    async function backfill(missing) {
      const box = $('#attbox'); box.innerHTML = '<p class="note">補填中…</p>';
      try {
        const cl = await listInfo('案例'), kl = await listInfo('對照表'), kf = kl.f['主治醫師'].EntityPropertyName, uf = cl.f['主治醫師'].EntityPropertyName;
        const keys = {}; for (const k of await getAll(`web/lists(guid'${kl.id}')/items?$top=5000&$select=Title,${kf}`)) keys[k.Title] = k[kf];
        let ok = 0; const miss = [];
        for (const c of missing) { const r = await resolveUser(keys[c.Title]); if (r.id) { await mergeItem(cl, c.Id, { __metadata: { type: cl.type }, [uf + 'Id']: r.id }); ok++; } else miss.push(`${c.Title}（${r.why}）`); }
        box.innerHTML = `<p class="note">已補填 ${ok} 例主治醫師。${miss.length ? `<br><span class="warn">${miss.length} 例需手動選取：${esc(miss.join('、'))}</span>` : ''}</p>`; meetReady = false;
      } catch (e) { box.innerHTML = `<p class="bad">補填失敗：${esc(e.message)}</p>`; }
    }
    (async () => { // 已收的案例若還沒有主治醫師，提供依對照表補填
      try {
        const cl = await listInfo('案例'); if (!cl.f['主治醫師']) return;
        const uf = cl.f['主治醫師'].EntityPropertyName;
        const missing = (await getAll(`web/lists(guid'${cl.id}')/items?$top=5000&$select=Id,Title,${uf}Id`)).filter(c => !c[uf + 'Id'] && !/^TEST-/.test(c.Title || ''));
        if (missing.length) { $('#attbox').innerHTML = `<p class="note warn">「案例」清單有 ${missing.length} 例還沒有主治醫師。 <button class="b" id="bf">依對照表補填</button></p>`; $('#bf').addEventListener('click', () => backfill(missing)); }
      } catch (_) {}
    })();
    (async () => { // 清單裡若有舊的測試資料（代碼 TEST- 開頭）才提示
      try {
        const li = await listInfo('案例'); const tv = (await getAll(`web/lists(guid'${li.id}')/items?$top=5000&$select=Id,Title`)).filter(x => /^TEST-/.test(x.Title || ''));
        if (tv.length) { $('#testbox').innerHTML = `<p class="note warn">「案例」清單裡有 ${tv.length} 例測試資料（代碼 TEST- 開頭）。 <button class="b" id="purge">刪除測試資料</button></p>`; $('#purge').addEventListener('click', purge); }
      } catch (_) {}
    })();
  }
  // 轉換 → 依 SharePoint 既有資料重新編號、標出重複
  async function plan(specs, demo) {
    const cl = await listInfo('案例'), kl = await listInfo('對照表'), chartF = kl.f['病歷號'], E = t => cl.f[t].EntityPropertyName;
    const cases = await getAll(`web/lists(guid'${cl.id}')/items?$top=5000&$select=Title,${E('病人簡稱')},${E('年齡')},${E('詳細診斷')}`);
    const titles = cases.map(x => x.Title || '');
    const fp = (stem, s, a, dx) => [stem.replace(/^TEST-/, ''), s || '', a ?? '', String(dx || '').slice(0, 40)].join('|');
    const byFp = {}; for (const c of cases) { const m = (c.Title || '').match(/^(.*-\d{4}-\d{2}-)\d+$/); if (m) byFp[fp(m[1], c[E('病人簡稱')], c[E('年齡')], c[E('詳細診斷')])] = c.Title; }
    const keys = await getAll(`web/lists(guid'${kl.id}')/items?$top=5000&$select=Title,${chartF.EntityPropertyName}`);
    const byChart = {}; for (const k of keys) { const c = String(k[chartF.EntityPropertyName] || '').trim(); if (c) (byChart[c] = byChart[c] || []).push(k.Title || ''); }
    const pre = demo ? 'TEST-' : '', next = {}, groups = [], items = [];
    const obj = (cols, row) => { const o = {}; cols.forEach((c, j) => o[c] = row[j]); return o; };
    for (const sp of specs) {
      const r = await MM.convert(await sp.file.arrayBuffer(), sp.ym);
      const g = { name: sp.file.name, ctype: r.ctype, leak: r.leak, n: r.drafts.length, dup: 0, bad: '' }; groups.push(g);
      { // Word 格式檢查：異常比例偏高就整個檔不寫入
        const n = r.drafts.length, pct = k => n ? k / n : 1, miss = [];
        if (!n) g.bad = '讀不到任何案例';
        else {
          const noAge = r.drafts.filter(d => d.age == null).length, noDx = r.drafts.filter(d => !d.dxDetail).length, noYm = r.drafts.filter(d => !d.ym || d.ym === '0000-00').length;
          const noLos = r.kind === 'mortality' ? 0 : r.drafts.filter(d => d.los == null).length;
          if (pct(noAge) > 0.3) miss.push(`${noAge} 例沒有年齡`); if (pct(noDx) > 0.3) miss.push(`${noDx} 例沒有診斷`);
          if (pct(noYm) > 0.3) miss.push(`${noYm} 例沒有日期`); if (pct(noLos) > 0.3) miss.push(`${noLos} 例沒有住院日數`);
          if (miss.length) g.bad = miss.join('、');
        }
      }
      for (let i = 0; i < r.drafts.length; i++) { const d = r.drafts[i];
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
        const att = cl.f['主治醫師'] ? await resolveUser(keyRow['主治醫師']) : null; if (att && att.id) caseRow['主治醫師'] = att.id;
        const noDate = !d.ym || d.ym === '0000-00';
        items.push({ g, att, code: noDate ? '（無日期）' : code, dup, noDate, ok: !dup && !noDate && !r.leak.length && !g.bad, draft: d, caseRow, keyRow });
      }
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
    // 病歷號：只有讀得到「對照表」的帳號才有；另外放，不併入案例資料（定案寫回時不會帶到正式紀錄）
    let charts = null;
    try {
      const kl = await listInfo('對照表'), cf = kl.f['病歷號'].EntityPropertyName; charts = {};
      for (const k of await getAll(`web/lists(guid'${kl.id}')/items?$top=5000&$select=Title,${cf}`)) { const v = String(k[cf] || '').trim(); if (k.Title && v) charts[k.Title] = v; }
    } catch (_) { charts = null; }
    return { charts, drafts, records: rs.map(it => toDisplay(rl, it)), factors: fs.map(it => toDisplay(fl, it)), actions: as.map(it => toDisplay(al, it)) };
  }
  async function decide(code, p) {
    const cl = await listInfo('案例'), rl = await listInfo('正式紀錄'); if (!caseRaw[code]) throw new Error('找不到案例 ' + code);
    const raw = await readOne('案例', caseRaw[code].Id); caseRaw[code] = raw; caseDisp[code] = toDisplay(cl, raw); // 定案前重讀最新內容
    if (caseDisp[code]['狀態'] === '已定案') throw new Error(code + ' 已經定案過。');
    const recTitles = () => getAll(`web/lists(guid'${rl.id}')/items?$top=5000&$select=Id,Title`);
    const ex = { value: (await recTitles()).filter(x => x.Title === code) };
    if (ex.value.length) throw new Error(code + ' 已經定案過。如需修正，請到「正式紀錄」該筆的「修正附註」補充。');
    const body = { __metadata: { type: rl.type } };
    for (const f of rl.ours) { // 複製案例欄位
      const cf = cl.f[f.InternalName === 'Title' ? '標題' : f.Title]; if (!cf || f.ReadOnlyField) continue;
      if (f.TypeAsString === 'User') continue;
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
    const recItem = await addItem(rl, body);
    { // 兩台電腦同時定案同一案：保留最早的一筆，較晚寫入的刪除並告知
      const same = (await recTitles()).filter(x => x.Title === code).sort((a, b) => a.Id - b.Id);
      if (same.length > 1 && recItem && same[0].Id !== recItem.Id) { await delItem(rl, recItem.Id).catch(() => {}); throw new Error(code + ' 剛剛已由另一台電腦定案，這次不重複寫入。請按「重新讀取」。'); }
    }
    try { await mergeItem(cl, raw.Id, { __metadata: { type: cl.type }, [cl.f['狀態'].EntityPropertyName]: '已定案' }); }
    catch (e) { throw new Error('定案已寫入，但案例狀態沒有更新。請到「案例」清單把這一例改成「已定案」。（' + e.message + '）'); }
    let _action = null, _actionError = null;
    if (p.action && p.action.title) { // 會中決議的改善行動：建一筆，連到這筆正式紀錄與主要根因
      try {
        const al = await listInfo('改善行動'), A = t => al.f[t] && al.f[t].EntityPropertyName, ab = { __metadata: { type: al.type }, Title: p.action.title };
        if (A('負責人') && p.action.owner) ab[A('負責人')] = p.action.owner;
        if (A('到期日') && p.action.due) ab[A('到期日')] = p.action.due + 'T00:00:00Z';
        if (A('狀態')) ab[A('狀態')] = '待辦';
        if (A('針對因素')) ab[A('針對因素') + 'Id'] = mid;
        if (A('關聯紀錄') && recItem && recItem.Id) ab[A('關聯紀錄') + 'Id'] = { __metadata: { type: 'Collection(Edm.Int32)' }, results: [recItem.Id] };
        await addItem(al, ab);
        _action = { '標題': p.action.title, '負責人': p.action.owner, '到期日': p.action.due, '狀態': '待辦', '針對因素': p.main, '關聯紀錄': code };
      } catch (e) { _actionError = e.message; }
    }
    return { _action, _actionError, ...caseDisp[code], '共識主要根因': p.main, '共識促成因素': p.factors.join(';'), '共識可預防性': p.prev, '決策方式': p.mode, '會議結論': p.conclusion, '長官意見': p.chief, '會議日期': today, '出席人數': p.attendees ?? '', '票數': p.votes ?? '', '投票分布': p.dist, '初評一致': initMain === p.main ? '是' : '否', '狀態': '已定案' };
  }
  /* ---------- 會前圈選「會中報告」：只改案例清單的一個欄位 ---------- */
  async function mark(code, on) {
    const cl = await listInfo('案例'), f = cl.f['會中報告']; if (!f) throw new Error('站台還沒有「會中報告」欄位');
    if (!caseRaw[code]) throw new Error('找不到案例 ' + code);
    await mergeItem(cl, caseRaw[code].Id, { __metadata: { type: cl.type }, [f.EntityPropertyName]: !!on });
    caseRaw[code][f.EntityPropertyName] = !!on; if (caseDisp[code]) caseDisp[code]['會中報告'] = on ? '是' : '否';
    return !!on;
  }
  /* ---------- 會中匿名投票：手機掃 QR 開「投票」清單的表單；每人以最後一票為準 ---------- */
  let voteLi = null;
  const vote = {
    async start() {
      const top = (await rq(`web/lists(guid'${voteLi.id}')/items?$select=Id&$orderby=Id desc&$top=1`)).value[0];
      const url = voteLi.url + '/NewForm.aspx';
      return { startId: top ? top.Id : 0, url, svg: QR.svg(url, 8) };
    },
    async tally(startId) {
      const pf = voteLi.f['可預防性'].EntityPropertyName;
      const rows = (await rq(`web/lists(guid'${voteLi.id}')/items?$top=5000&$select=Id,AuthorId,${pf}&$filter=Id gt ${startId}&$orderby=Id asc`)).value;
      const last = {}; for (const r of rows) last[r.AuthorId] = r;
      const dist = {}; for (const r of Object.values(last)) if (r[pf]) dist[r[pf]] = (dist[r[pf]] || 0) + 1;
      return { n: Object.keys(last).length, dist, rows };
    },
    async stop(startId, code) {
      const t = await this.tally(startId), cf = voteLi.f['案例代碼'].EntityPropertyName, sf = voteLi.f['場次'].EntityPropertyName;
      for (const r of t.rows) { try { await mergeItem(voteLi, r.Id, { __metadata: { type: voteLi.type }, [cf]: code, [sf]: String(startId) }); } catch (_) {} }
      return { n: t.n, dist: t.dist };
    }
  };
  async function meetInit() {
    meetReady = true; const P = $('#p-meet'); P.innerHTML = '<div class="load">讀取中…</div>';
    try {
      const [data, can, canMark] = await Promise.all([loadData(), canAdd('正式紀錄'), canEdit('案例')]);
      const hasReport = !!(await listInfo('案例')).f['會中報告'];
      try { voteLi = await listInfo('投票'); } catch (_) { voteLi = null; }
      P.innerHTML = '<iframe title="M&M 會議儀表板"></iframe>'; frame = P.querySelector('iframe');
      const w = frame.contentWindow, doc = frame.contentDocument;
      doc.open(); doc.write(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><style>${DASH_CSS}</style></head><body>${DASH_BODY}</body></html>`); doc.close();
      w.MM_SP = { data, canDecide: can, decide, hasReport, mark: hasReport && canMark ? mark : null, vote: can && voteLi ? vote : null, actionsUrl: (await listInfo('改善行動')).url, onView: v => { $('#bar').style.display = v === 'present' ? 'none' : ''; } };
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
