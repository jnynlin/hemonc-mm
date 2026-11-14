// 壓力測試用：包在 mock.js 外層，模擬真實 SharePoint 的延遲、錯誤、流量限制、分頁與 5000 筆門檻
(() => {
  const S = window.STRESS = { latency: 0, failRate: 0, throttleEvery: 0, n: 0, calls: 0, posts: 0, throttled: 0, failed: 0, thresholdHits: 0, indexed: {}, pageSize: 5000 };
  const inner = window.fetch;
  const J = (o, s, h = {}) => new Response(JSON.stringify(o), { status: s, headers: { 'Content-Type': 'application/json', ...h } });
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  window.fetch = async (url, opt = {}) => {
    const u = String(url); if (!u.includes('/_api/')) return inner(url, opt);
    S.calls++; const m = (opt.method || 'GET').toUpperCase(); if (m === 'POST') S.posts++;
    if (S.latency) await sleep(S.latency * (0.5 + Math.random()));
    const path = decodeURIComponent(u.split('/_api/')[1]);
    if (path !== 'contextinfo') {
      if (S.throttleEvery && ++S.n % S.throttleEvery === 0) { S.throttled++; return J({ 'odata.error': { message: { value: 'The request has been throttled' } } }, 429, { 'Retry-After': '1' }); }
      if (S.failRate && Math.random() < S.failRate) { S.failed++; return J({ 'odata.error': { message: { value: 'Internal Server Error (injected)' } } }, 500); }
    }
    // 5000 筆門檻：清單超過 5000 筆時，對未建索引的 Title 篩選會失敗
    const lm = path.match(/^web\/lists\(guid'(g\d)'\)\/items\?(.*)$/);
    if (lm && m === 'GET') {
      const names = ['案例', '正式紀錄', '根因分類', '改善行動', '對照表', '投票'];
      const t = names[+lm[1].slice(1)]; const rows = (MOCK.DB[t] || []);
      const f = (lm[2].match(/\$filter=([^&]*)/) || [])[1] || '';
      if (rows.length > 5000 && /Title/.test(f) && !S.indexed[t]) { S.thresholdHits++; return J({ 'odata.error': { message: { value: 'The attempted operation is prohibited because it exceeds the list view threshold.' } } }, 500); }
      // 分頁：每頁最多 pageSize 筆，附 odata.nextLink
      const big = +((lm[2].match(/\$top=(\d+)/) || [])[1] || 0) >= 1000; const r = await inner(big ? url.replace(/([?&])\$top=\d+&?/, '$1') : url, opt); const j = await r.json(); const all = j.value || [];
      const skip = +((lm[2].match(/\$skiptoken=(\d+)/) || [])[1] || 0);
      const page = all.slice(skip, skip + S.pageSize);
      const out = { value: page };
      if (skip + S.pageSize < all.length) out['odata.nextLink'] = u.replace(/&\$skiptoken=\d+/, '') + '&$skiptoken=' + (skip + S.pageSize);
      return J(out, 200);
    }
    return inner(url, opt);
  };
})();
