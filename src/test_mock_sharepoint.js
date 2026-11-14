// 模擬 SharePoint REST（只實作 mm_app.js 用到的部分）
(() => {
  const T = (Title, InternalName, TypeAsString, extra = {}) => ({ Title, InternalName, EntityPropertyName: InternalName, TypeAsString, ReadOnlyField: false, ...extra });
  const PREV = ['不可預防', '可能可預防', '很可能可預防', '明確可預防'];
  const caseF = () => [T('標題', 'Title', 'Text'), T('年月', 'mmYearMonth', 'Text'), T('案例類型', 'mmCaseType', 'MultiChoice', { Choices: ['超長住院', 'Morbidity', 'Mortality'] }), T('年齡層', 'mmAgeBand', 'Choice', { Choices: ['<40', '40–49', '50–59', '60–69', '70–79', '≥80'] }), T('性別', 'mmSex', 'Choice', { Choices: ['男', '女'] }), T('診斷群', 'mmDxGroup', 'Text'), T('治療階段', 'mmPhase', 'Text'), T('詳細診斷', 'mmDxDetail', 'Text'), T('治療內容', 'mmTreatment', 'Note'), T('臨床經過', 'mmCourse', 'Note'), T('住院天數', 'mmLOS', 'Number'), T('可避免天數', 'mmAvoidableDays', 'Number'), T('死亡預期性', 'mmDeathExpect', 'Choice', { Choices: ['不適用', '預期（疾病進展）', '非預期'] }), T('Morbidity等級', 'mmMorbGrade', 'Choice', { Choices: ['不適用', 'I', 'II'] }), T('初評主要根因', 'mmInitMainFactor', 'Lookup'), T('初評促成因素', 'mmInitFactors', 'LookupMulti'), T('初評可預防性', 'mmInitPrev', 'Choice', { Choices: PREV }), T('主治醫師回覆', 'mmAttStatus', 'Choice', { Choices: ['未徵詢', '同意初評', '不同意初評'] }), T('主治醫師說明', 'mmAttNote', 'Note'), T('五個為什麼', 'mmFiveWhys', 'Note'), T('學習重點', 'mmLessons', 'Note'), T('參考指引文獻', 'mmRefs', 'Note'), T('病人簡稱', 'mmPatientShort', 'Text'), T('年齡', 'mmAge', 'Number')];
  const L = {
    '案例': [...caseF(), T('狀態', 'mmStatus', 'Choice', { Choices: ['整理中', '待討論', '已定案'] }), T('主治醫師', 'mmAttending', 'User')],
    '正式紀錄': [...caseF(), T('共識主要根因', 'mmConsMainFactor', 'Lookup'), T('共識促成因素', 'mmConsFactors', 'LookupMulti'), T('共識可預防性', 'mmConsPrev', 'Choice', { Choices: PREV }), T('決策方式', 'mmDecisionMode', 'Choice', { Choices: ['三方一致', '主席裁定', '投票'] }), T('長官意見', 'mmChiefComment', 'Note'), T('會議結論', 'mmConclusion', 'Note'), T('會議日期', 'mmMeetingDate', 'DateTime'), T('出席人數', 'mmAttendees', 'Number'), T('票數', 'mmVotes', 'Number'), T('投票分布', 'mmVoteDist', 'Note'), T('初評一致', 'mmConcordant', 'Choice', { Choices: ['是', '否'] })],
    '根因分類': [T('標題', 'Title', 'Text'), T('層級', 'mmLevel', 'Choice', { Choices: ['病人因素', '工作任務與流程', '超長住院因素'] }), T('排序', 'mmSortOrder', 'Number')],
    '改善行動': [T('標題', 'Title', 'Text'), T('負責人', 'mmOwner', 'Text'), T('到期日', 'mmDueDate', 'DateTime'), T('狀態', 'mmStatus', 'Choice', { Choices: ['待辦', '進行中', '完成', '取消'] }), T('備註', 'mmRemarks', 'Note')],
    '投票': [T('標題', 'Title', 'Text'), T('可預防性', 'mmVotePrev', 'Choice', { Choices: PREV }), T('案例代碼', 'mmVoteCase', 'Text'), T('場次', 'mmVoteSession', 'Text')],
    '對照表': [T('標題', 'Title', 'Text'), T('病歷號', 'mmChartNo', 'Text'), T('姓名', 'mmPatientName', 'Text'), T('床號', 'mmBed', 'Text'), T('主治醫師', 'mmAttending', 'Text'), T('原主治', 'mmPrevAttending', 'Text'), T('住院日期', 'mmAdmissionDate', 'Text'), T('死亡日期', 'mmDeathDate', 'Text'), T('住院日數原文', 'mmLOSRaw', 'Text')],
  };
  const ids = {}, names = {}; Object.keys(L).forEach((t, i) => { ids[t] = 'g' + i; names['g' + i] = t; });
  const DB = { '案例': [], '正式紀錄': [], '根因分類': [], '改善行動': [], '對照表': [], '投票': [] }; let seq = 1;
  [['病人因素', '疾病嚴重度／多重共病'], ['工作任務與流程', '檢驗檢查結果延遲或未追蹤'], ['超長住院因素', '等待檢查／處置'], ['超長住院因素', '感染併發症'], ['臨床決策', '診斷假設未及時修正（定錨／過早結案）'], ['臨床決策', '值班負荷與疲勞']].forEach(([g, n], i) => DB['根因分類'].push({ Id: seq++, Title: n, mmLevel: g, mmSortOrder: i }));
  // 先放一筆「已收過」：讓 10 月第 3 例的病歷號被視為重複
  window.MOCK = { DB, L, log: [], perms: { '對照表': true } };
  const fac = id => { const f = DB['根因分類'].find(x => x.Id === id); return f ? { Id: f.Id, Title: f.Title } : null; };
  const J = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { 'Content-Type': 'application/json' } });
  const realFetch = window.fetch.bind(window);
  window.fetch = async (url, opt = {}) => {
    url = String(url); if (!url.includes('/_api/')) return realFetch(url, opt);
    const path = decodeURIComponent(url.split('/_api/')[1]); const m = (opt.method || 'GET').toUpperCase(), xm = (opt.headers || {})['X-HTTP-Method'];
    MOCK.log.push(m + ' ' + (xm || '') + ' ' + path.slice(0, 90));
    if (path === 'contextinfo') return J({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } });
    if (path.startsWith('SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser')) { const qs = JSON.parse(opt.body).queryParams.QueryString; const hits = (MOCK.users || []).filter(u => u.DisplayText.includes(qs)).map(u => ({ Key: u.Key, DisplayText: u.DisplayText })); return J({ d: { ClientPeoplePickerSearchUser: JSON.stringify(hits) } }); }
    if (path === 'web/ensureuser') { const k = JSON.parse(opt.body).logonName; const u = (MOCK.users || []).find(x => x.Key === k); return u ? J({ d: { Id: u.Id, Title: u.DisplayText } }) : J({ 'odata.error': { message: { value: 'user not found' } } }, 404); }
    if (/^web\/lists\/getbytitle\('投票'\)/.test(path) && MOCK.voteTitleBroken) return J({ 'odata.error': { message: { value: 'List does not exist' } } }, 404);
    let mt = path.match(/^web\/lists\/getbytitle\('(.+?)'\)/);
    if (mt) { const t = mt[1].replace(/''/g, "'"); if (!L[t] || (t === '對照表' && !MOCK.perms['對照表'])) return J({ 'odata.error': { message: { value: 'List does not exist' } } }, 404); return J({ Id: ids[t], ListItemEntityTypeFullName: 'SP.Data.' + ids[t], RootFolder: { ServerRelativeUrl: '/sites/x/Lists/' + (t === '投票' ? 'mmvote' : t) } }); }
    let gl = path.match(/^web\/GetList\(@u\)\?.*@u='([^']*)'/);
    if (gl) { const u = gl[1]; const tt = /\/Lists\/mmvote$/.test(u) ? '投票' : null; if (!tt) return J({ 'odata.error': { message: { value: 'not found' } } }, 404); return J({ Id: ids[tt], ListItemEntityTypeFullName: 'SP.Data.' + ids[tt], RootFolder: { ServerRelativeUrl: '/sites/x/Lists/mmvote' } }); }
    // 模擬多語系：依名稱「投票」找不到
    if (/^web\/lists\/getbytitle\('投票'\)/.test(path) && MOCK.voteTitleBroken) return J({ 'odata.error': { message: { value: 'List does not exist' } } }, 404);
    mt = path.match(/^web\/lists\(guid'(g\d)'\)\/(.*)$/); if (!mt) return J({ error: 'unknown ' + path }, 400);
    const t = names[mt[1]], rest = mt[2], F = L[t];
    if (rest.startsWith('fields')) return J({ value: F });
    if (rest === 'EffectiveBasePermissions') return J({ Low: '4294967295', High: '2147483647' });
    let im = rest.match(/^items\((\d+)\)(?:\?.*)?$/);
    if (im) { const id = +im[1], i = DB[t].findIndex(x => x.Id === id); if (m === 'GET') return i < 0 ? J({ 'odata.error': { message: { value: 'Item does not exist' } } }, 404) : J({ ...DB[t][i] }); if (xm === 'DELETE') { DB[t].splice(i, 1); return new Response(null, { status: 204 }); } if (xm === 'MERGE') { const b = JSON.parse(opt.body); delete b.__metadata; Object.assign(DB[t][i], b); if (b.mmAttendingId) { const u = (MOCK.users || []).find(x => x.Id === b.mmAttendingId); DB[t][i].mmAttending = u ? { Id: u.Id, Title: u.DisplayText } : null; } return new Response(null, { status: 204 }); } }
    if (rest.startsWith('items') && m === 'GET') {
      let rows = DB[t]; const f = rest.match(/\$filter=([^&]*)/);
      if (f) { let a; if ((a = f[1].match(/startswith\(Title,'(.*)'\)/))) rows = rows.filter(r => (r.Title || '').startsWith(a[1])); if ((a = f[1].match(/Title eq '(.*)'/))) rows = rows.filter(r => r.Title === a[1].replace(/''/g, "'")); if ((a = f[1].match(/Id gt (\d+)/))) rows = rows.filter(r => r.Id > +a[1]); }
      if (/\$orderby=Id desc/.test(rest)) rows = [...rows].sort((x, y) => y.Id - x.Id); const tp = rest.match(/\$top=(\d+)/); if (tp) rows = rows.slice(0, +tp[1]);
      return J({ value: rows.map(r => ({ ...r })) });
    }
    if (rest === 'items' && m === 'POST') {
      const b = JSON.parse(opt.body); delete b.__metadata; const it = { Id: seq++ };
      for (const [k, v] of Object.entries(b)) {
        if (k === 'mmAttendingId' && F.find(x => x.InternalName === 'mmAttending')) { const u = (MOCK.users || []).find(x => x.Id === v); it.mmAttendingId = v; it.mmAttending = u ? { Id: u.Id, Title: u.DisplayText } : null; continue; }
        const fl = F.find(x => x.EntityPropertyName + 'Id' === k && /^Lookup/.test(x.TypeAsString));
        if (fl) { it[fl.InternalName] = fl.TypeAsString === 'Lookup' ? fac(v) : v.results.map(fac); continue; }
        if (!F.find(x => x.EntityPropertyName === k)) return J({ 'odata.error': { message: { value: 'property ' + k + ' does not exist on ' + t } } }, 400);
        it[k] = v && v.results ? v.results : v;
      }
      DB[t].push(it); return J({ d: { Id: it.Id } }, 201);
    }
    return J({ error: 'unhandled ' + rest }, 400);
  };
})();
