#!/usr/bin/env python3
"""M&M 書籤工具壓力測試：模擬 SharePoint（mock.js＋stress_mock.js），在無頭 Edge 逐一跑情境，輸出 JSON。
用法：python3 stress_run.py <scratchpad> [情境…]"""
import json, base64, subprocess, sys, re, html, os, time
SP = sys.argv[1]; only = sys.argv[2:]
WIN = '/mnt/c/Users/u9754/Downloads/MM_SharePoint'
EDGE = '/mnt/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe'
rd = lambda p: open(p, encoding='utf8').read()
B = lambda p: base64.b64encode(open(p, 'rb').read()).decode()
mock, stress, app = rd(f'{SP}/mmapp/test/mock.js'), rd(f'{SP}/mmapp/test/stress_mock.js'), rd(f'{SP}/mmapp/mm_app.js')
users = rd(f'{SP}/mmapp/test/users.json')
files = {k: B(f'{SP}/mm_in/{v}') for k, v in {'OCT': 'oct.docx', 'MORT': 'mort.docx'}.items()}
for f in os.listdir(f'{SP}/mm_in/stress'): files['S_' + f.split('.')[0].upper()] = B(f'{SP}/mm_in/stress/{f}')

COMMON = r'''
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const until=async(fn,ms=60000)=>{const t=Date.now();while(Date.now()-t<ms){try{const v=fn();if(v)return v;}catch(e){}await sleep(50);}throw new Error('timeout');};
const F=(k,n)=>new File([Uint8Array.from(atob(FILES[k]),c=>c.charCodeAt(0))],n);
const OUT={};
async function boot(){ try{localStorage.removeItem('mm-app-tab')}catch(e){} MOCK.users=UJ.users; (0,eval)(APP); const sh=document.getElementById('mm-app-host').shadowRoot; await until(()=>sh.querySelector('#file')||sh.querySelector('.banner')); return sh; }
async function intake(sh,fs){ const dt=new DataTransfer(); fs.forEach(f=>dt.items.add(f)); const i=sh.querySelector('#file'); i.files=dt.files; i.dispatchEvent(new Event('change')); await until(()=>{const r=sh.querySelector('#res'); return sh.querySelector('#write')||(r&&/失敗|不會寫入|請選擇|不是有效|不是 Word/.test(r.textContent))||/請選擇/.test(sh.querySelector('#ds').textContent);},180000); }
async function write(sh){ const t=performance.now(); sh.querySelector('#write').click(); await until(()=>/完成：新增/.test((sh.querySelector('#log')||{}).textContent||''),600000); return {ms:Math.round(performance.now()-t), log:sh.querySelector('#log').textContent}; }
async function meet(sh){ sh.querySelector('.tab[data-t=meet]').click(); const fr=await until(()=>sh.querySelector('iframe'),120000); const d=await until(()=>fr.contentDocument&&fr.contentDocument.getElementById('go-present')&&fr.contentDocument,180000); return {d,w:fr.contentWindow}; }
function consistency(){ const c=new Set(MOCK.DB['案例'].map(x=>x.Title)), k=new Set(MOCK.DB['對照表'].map(x=>x.Title)); const dupC=MOCK.DB['案例'].length-c.size; return {cases:MOCK.DB['案例'].length,keys:MOCK.DB['對照表'].length,caseOnly:[...c].filter(t=>!k.has(t)).length,keyOnly:[...k].filter(t=>!c.has(t)).length,dupCodes:dupC}; }
function ready(c){ c.mmStatus='待討論'; c.mmAttStatus='同意初評'; c.mmInitMainFactor={Id:3,Title:'等待檢查／處置'}; c.mmInitPrev='可能可預防'; }
function seed(nCase,nRec,nAct){ const DB=MOCK.DB; let id=100000; const yms=[]; for(let y=2019;y<=2026;y++)for(let m=1;m<=12;m++) yms.push(y+'-'+String(m).padStart(2,'0'));
  for(let i=0;i<nCase;i++){ const ym=yms[Math.floor(i/(nCase/yms.length))%yms.length]; const n=(i%Math.ceil(nCase/yms.length))+1; const code=`LOS-${ym}-${String(n).padStart(2,'0')}`; id++;
    DB['案例'].push({Id:id,Title:code,mmYearMonth:ym,mmCaseType:['超長住院'],mmPatientShort:'王OO',mmAge:60,mmDxDetail:'Diffuse large B cell lymphoma stage IV '+i,mmLOS:40,mmStatus:'已定案',mmAttStatus:'同意初評',mmInitMainFactor:{Id:3,Title:'等待檢查／處置'},mmInitPrev:'可能可預防'});
    DB['對照表'].push({Id:id+500000,Title:code,mmChartNo:String(70000000+i),mmAttending:'某醫師'}); }
  for(let i=0;i<nRec;i++){ const c=DB['案例'][i]; DB['正式紀錄'].push({...c,Id:300000+i,mmConsMainFactor:{Id:3,Title:'等待檢查／處置'},mmConsFactors:[],mmConsPrev:'可能可預防',mmDecisionMode:'三方一致',mmMeetingDate:c.mmYearMonth+'-20T00:00:00Z',mmConcordant:'是'}); }
  for(let i=0;i<nAct;i++) DB['改善行動'].push({Id:400000+i,Title:'行動 '+i,mmOwner:'會後指派',mmDueDate:'2026-0'+(1+i%9)+'-15T00:00:00Z',mmStatus:['待辦','進行中','完成'][i%3]});
}
'''
SCEN = {
# 1 大量資料：案例 6000、正式紀錄 3000（超過 5000 筆門檻與分頁）
'volume': r'''seed(6000,3000,300); const sh=await boot();
  await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx')]);
  const codes=[...sh.querySelectorAll('td.c')].map(t=>t.textContent); OUT.intakeFirstCode=codes[0]; OUT.intakeErr=(sh.querySelector('#res p.bad')||{}).textContent||'';
  const existingOct=MOCK.DB['案例'].filter(c=>/^LOS-2026-10-/.test(c.Title)).map(c=>c.Title).sort().pop(); OUT.maxExistingOct=existingOct;
  OUT.codeCollision=codes.some(c=>MOCK.DB['案例'].some(x=>x.Title===c));
  OUT.backfillNotice=(sh.querySelector('#attbox')||{}).textContent||'';
  MOCK.DB['案例'].slice(-25).forEach(ready);
  let t=performance.now(); const {d,w}=await meet(sh); OUT.meetLoadMs=Math.round(performance.now()-t);
  OUT.prepRows=d.querySelectorAll('#v-prep tr.click').length;
  t=performance.now(); d.querySelector('#nav button[data-view=records]').click(); await until(()=>d.getElementById('v-records').textContent.length>50); OUT.recordsRenderMs=Math.round(performance.now()-t);
  d.querySelector('#nav button[data-view=prep]').click(); d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide'));
  f.querySelector('button[type=submit]').click(); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000);
  OUT.decide=(d.querySelector('.flash')||d.getElementById('dmsg')).textContent.slice(0,120);
  OUT.thresholdHits=STRESS.thresholdHits; OUT.calls=STRESS.calls;''',
# 1b 正式紀錄也超過 5000：定案時查重複會碰門檻
'volume_records': r'''seed(6000,6000,10); MOCK.DB['案例'].slice(-3).forEach(ready); const sh=await boot();
  const {d}=await meet(sh); d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide'));
  f.querySelector('button[type=submit]').click(); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000);
  OUT.decide=(d.querySelector('.flash')||d.getElementById('dmsg')).textContent.slice(0,140); OUT.thresholdHits=STRESS.thresholdHits;''',
# 2 流量限制：每 4 個請求就回 429
'throttle': r'''const sh=await boot(); STRESS.throttleEvery=4;
  await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx'),F('MORT','血液腫瘤科2026年9&10月份Mortality討論會.docx')]);
  OUT.planErr=(sh.querySelector('#res p.bad')||{}).textContent||''; if(!sh.querySelector('#write')) return;
  const r=await write(sh); OUT.result=r.log.split('\n').filter(l=>/完成/.test(l))[0]; OUT.notWritten=(r.log.match(/未寫入/g)||[]).length; OUT.throttled=STRESS.throttled; OUT.consistency=consistency();''',
# 3 隨機 8% 伺服器錯誤
'faults': r'''const sh=await boot(); STRESS.failRate=0.08;
  await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx'),F('MORT','血液腫瘤科2026年9&10月份Mortality討論會.docx')]);
  OUT.planErr=(sh.querySelector('#res p.bad')||{}).textContent||''; if(!sh.querySelector('#write')) return;
  const r=await write(sh); OUT.result=r.log.split('\n').filter(l=>/完成/.test(l))[0]; OUT.failedInjected=STRESS.failed; OUT.consistency=consistency();''',
'faults_redrop': r'''const sh=await boot(); STRESS.failRate=0.1; const fs=()=>[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx'),F('MORT','血液腫瘤科2026年9&10月份Mortality討論會.docx')];
  await intake(sh,fs()); const r1=await write(sh); OUT.first=r1.log.split('\n').filter(l=>/完成/.test(l))[0]; OUT.afterFirst=consistency();
  STRESS.failRate=0; sh.querySelector('#res').innerHTML=''; await intake(sh,fs()); OUT.secondWn=sh.querySelector('#wn').textContent; const r2=await write(sh); OUT.second=r2.log.split('\n').filter(l=>/完成/.test(l))[0]; OUT.final=consistency();''',
# 4 慢網路：每個請求 0.25–0.75 秒
'latency': r'''STRESS.latency=500; const sh=await boot();
  let t=performance.now(); await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx')]); OUT.planMs=Math.round(performance.now()-t);
  const r=await write(sh); OUT.writeMs=r.ms; OUT.consistency=consistency();
  MOCK.DB['案例'].slice(0,3).forEach(ready); t=performance.now(); await meet(sh); OUT.meetLoadMs=Math.round(performance.now()-t);''',
# 5 重複按定案、兩台電腦同時定案同一案
'concurrent_decide': r'''const sh=await boot(); await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx')]); await write(sh);
  MOCK.DB['案例'].slice(0,2).forEach(ready); STRESS.latency=200; const {d,w}=await meet(sh);
  d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide')); const b=f.querySelector('button[type=submit]'); b.click(); b.click(); b.click();
  await until(()=>d.querySelector('.flash'),60000); await sleep(1500); const c1=MOCK.DB['案例'][0].Title; OUT.doubleClickRecords=MOCK.DB['正式紀錄'].filter(r=>r.Title===c1).length;
  const c2=MOCK.DB['案例'][1].Title; const p={main:'等待檢查／處置',factors:[],prev:'可能可預防',mode:'三方一致',attendees:null,votes:null,dist:'',conclusion:'',chief:'',action:null};
  const rs=await Promise.allSettled([w.MM_SP.decide(c2,p),w.MM_SP.decide(c2,p)]); OUT.parallel=rs.map(r=>r.status+(r.reason?': '+r.reason.message.slice(0,40):'')); OUT.parallelRecords=MOCK.DB['正式紀錄'].filter(r=>r.Title===c2).length;''',
# 6 投票：40 人、10 人改投、結束當下仍有人送出
'votes': r'''const sh=await boot(); await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx')]); await write(sh);
  ready(MOCK.DB['案例'][0]); STRESS.latency=150; const {d}=await meet(sh); d.getElementById('go-present').click(); await until(()=>d.getElementById('decide'));
  d.getElementById('vstart').click(); await until(()=>d.querySelector('.vote-ov .vqr svg')); const P=['不可預防','可能可預防','很可能可預防','明確可預防']; let vid=20000;
  for(let a=1;a<=40;a++){ MOCK.DB['投票'].push({Id:vid++,AuthorId:a,mmVotePrev:P[a%4]}); if(a%8===0) await sleep(400); }
  for(let a=1;a<=10;a++) MOCK.DB['投票'].push({Id:vid++,AuthorId:a,mmVotePrev:'明確可預防'});
  await until(()=>d.getElementById('vn').textContent==='40',30000); OUT.live=d.getElementById('vn').textContent;
  d.getElementById('vend').click(); for(let a=41;a<=45;a++) MOCK.DB['投票'].push({Id:vid++,AuthorId:a,mmVotePrev:'不可預防'});
  await until(()=>d.getElementById('vapply'),30000); OUT.result=d.querySelector('.vote-ov').textContent.replace(/\s+/g,' ').slice(0,140);
  const tagged=MOCK.DB['投票'].filter(v=>v.mmVoteCase).length; OUT.tagged=tagged; OUT.expect='40 人（改投以最後一票：明確可預防 10＋其他 30），結束後才送出的 5 票可能算進或不算，但標記數應與計票一致';''',
# 7 異常 Word 檔與惡意字串
'bad_files': r'''const sh=await boot(); const res={};
  for(const [k,n] of [['S_EMPTY','2026年11月empty.docx'],['S_TEXT','2026年11月text.docx'],['S_NOTWORD','2026年11月notword.docx'],['S_NOTABLE','血液腫瘤科2026年11月份超長住院討論會.docx']]){
    await intake(sh,[F(k,n)]); res[k]=(sh.querySelector('#res').textContent||'').replace(/\s+/g,' ').slice(0,90)+' | 寫入鈕：'+(sh.querySelector('#write')?(sh.querySelector('#write').disabled?'停用':'可按 '+sh.querySelector('#wn').textContent):'無'); sh.querySelector('#res').innerHTML=''; }
  OUT.files=res; let t=performance.now(); await intake(sh,[F('S_BIG300','血液腫瘤科2026年10月份超長住院討論會.docx')]); OUT.big={planMs:Math.round(performance.now()-t),rows:sh.querySelectorAll('td.c').length,wn:(sh.querySelector('#wn')||{}).textContent};
  t=performance.now(); const r=await write(sh); OUT.big.writeMs=r.ms; OUT.big.consistency=consistency();
  sh.querySelector('#res').innerHTML=''; await intake(sh,[F('S_XSS','血液腫瘤科2026年10月份超長住院討論會.docx')]); OUT.xssPreviewLiteral=/&lt;img|<img src=x/.test(sh.querySelector('#res').innerHTML)&&!sh.querySelector('#res img');
  MOCK.DB['案例'].length=0; MOCK.DB['對照表'].length=0; await sleep(100); sh.querySelector('#res').innerHTML=''; await intake(sh,[F('S_XSS','血液腫瘤科2026年10月份超長住院討論會.docx')]); await write(sh);
  ready(MOCK.DB['案例'][0]); const {d}=await meet(sh); d.getElementById('go-present').click(); await until(()=>d.getElementById('decide')); await sleep(500);
  OUT.xssExecuted=!!(window.PWNED||(d.defaultView&&d.defaultView.PWNED)); OUT.xssImgInProjection=!!d.querySelector('.pres img');''',
# 8 v1.19 開會頁：病歷號（取自對照表）、會前勾選會中報告、住院天數尺；定案寫回不得帶出病歷號
'meet_v119': r'''const sh=await boot(); await intake(sh,[F('OCT','血液腫瘤科2026年10月份超長住院討論會.docx')]); await write(sh);
  MOCK.DB['案例'].slice(0,4).forEach(ready); const {d}=await meet(sh); const key=t=>(MOCK.DB['對照表'].find(k=>k.Title===t)||{}).mmChartNo;
  OUT.chartCol=[...d.querySelectorAll('#v-prep th')].some(th=>th.textContent==='病歷號');
  const rows=()=>[...d.querySelectorAll('#v-prep tr.click')]; OUT.prepRows=rows().length;
  OUT.chartAllMatch=rows().every(tr=>tr.querySelector('td.chartno').textContent===key(tr.querySelector('td.nw b').textContent));
  OUT.repBoxes=d.querySelectorAll('input[data-rep]').length; const target=rows()[2].querySelector('td.nw b').textContent, tch=key(target);
  d.querySelector('input[data-rep="'+target+'"]').click(); await until(()=>/會中報告/.test((d.querySelector('#v-prep h3')||{}).textContent||''));
  OUT.dbFlag=MOCK.DB['案例'].find(c=>c.Title===target).mmReport===true; OUT.stayedOnPrep=!d.getElementById('v-prep').hidden;
  OUT.firstIsTarget=rows()[0].querySelector('td.nw b').textContent===target; OUT.onlyRepToggle=!!d.getElementById('only-rep');
  d.getElementById('go-present').click(); await until(()=>d.querySelector('.pres h2'));
  OUT.presFirstIsTarget=d.querySelector('.pres .meta.above').textContent.includes(target); OUT.presChart=d.querySelector('.pres h2 .chartno').textContent==='病歷號 '+tch;
  OUT.rail=d.querySelectorAll('.rail button').length; OUT.railRep=d.querySelectorAll('.rail button.rep').length; OUT.ruler=!!d.querySelector('.stay svg');
  d.querySelectorAll('.rail button')[1].click(); await sleep(100); OUT.railJump=d.querySelector('.rail button.cur').dataset.j==='1'; d.querySelectorAll('.rail button')[0].click(); await sleep(100);
  const f=await until(()=>d.getElementById('decide')); f.querySelector('button[type=submit]').click(); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000);
  OUT.decide=(d.querySelector('.flash')||d.getElementById('dmsg')).textContent.slice(0,60); const rec=MOCK.DB['正式紀錄'].find(r=>r.Title===target)||{};
  OUT.recExtraKeys=Object.keys(rec).filter(k=>/Chart|Report|PatientName|Bed/.test(k)); OUT.recHasChart=JSON.stringify(rec).includes(tch);
  d.getElementById('ex').click(); await sleep(100); d.getElementById('show-chart').click(); await sleep(100); OUT.hideChartWorks=!d.querySelector('#v-prep td.chartno');
  const t2=rows()[0].querySelector('td.nw b').textContent; d.querySelector('input[data-rep="'+t2+'"]').click(); await until(()=>d.getElementById('only-rep')); d.getElementById('only-rep').click(); await sleep(100); OUT.onlyRepRows=rows().length;
  d.getElementById('go-present').click(); const f2=await until(()=>d.getElementById('decide')); f2.querySelector('button[type=submit]').click(); await until(()=>d.querySelector('.flash'),60000); await sleep(100);
  OUT.afterLastRep=d.querySelector('.flash').textContent.slice(-22); OUT.falseAllDone=/已全部定案$/.test((d.querySelector('.pres .wait')||{}).textContent||''); OUT.restShown=d.querySelectorAll('.rail button').length;''',
# 9 沒有對照表權限、站台還沒有「會中報告」欄位、案例唯讀：開會頁照常可用，只是不顯示對應功能
'meet_limited': r'''seed(6,0,0); MOCK.DB['案例'].forEach(ready); MOCK.perms['對照表']=false; MOCK.noReportField=true; const sh=await boot();
  const {d}=await meet(sh); OUT.prepRows=d.querySelectorAll('#v-prep tr.click').length; OUT.chartCol=!!d.querySelector('#v-prep td.chartno'); OUT.repCol=[...d.querySelectorAll('#v-prep th')].some(th=>th.textContent==='報告'); OUT.showChartToggle=!!d.getElementById('show-chart');
  d.getElementById('go-present').click(); await until(()=>d.querySelector('.pres h2')); OUT.presChart=!!d.querySelector('.pres .chartno'); OUT.presOk=d.querySelector('.pres h2').textContent.length>0; OUT.ruler=!!d.querySelector('.stay svg');''',
'meet_readonly': r'''seed(6,0,0); MOCK.DB['案例'].forEach(ready); MOCK.DB['案例'][1].mmReport=true; MOCK.permLow={'案例':1,'正式紀錄':1}; const sh=await boot();
  const {d}=await meet(sh); OUT.prepRows=d.querySelectorAll('#v-prep tr.click').length; OUT.boxes=d.querySelectorAll('input[data-rep]').length; OUT.ticks=[...d.querySelectorAll('#v-prep td.rep')].filter(t=>t.textContent==='✓').length;
  d.getElementById('go-present').click(); await until(()=>d.querySelector('.pres h2')); OUT.btnLabel=d.getElementById('dopen').textContent.replace('D',''); d.getElementById('dopen').click(); await sleep(100);
  OUT.drawerOpens=!d.getElementById('drawer').hidden; OUT.hasForm=!!d.getElementById('decide'); OUT.notice=d.getElementById('drawer').textContent.includes('定案由主任執行');''',
}
results = {}
for name, body in SCEN.items():
    if only and name not in only: continue
    scn = COMMON + '(async()=>{try{' + body + '}catch(e){OUT.FAIL=String(e&&e.stack||e).slice(0,300);} document.getElementById("result").textContent=JSON.stringify(OUT);})();'
    page = '<!doctype html><html><head><meta charset="utf-8"></head><body><pre id="result">running</pre><script>' + mock + '</script><script>' + stress + '</script><script>const APP=' + json.dumps(app) + ';const UJ=' + users + ';const FILES=' + json.dumps(files) + ';' + scn + '</script></body></html>'
    open(f'{WIN}/_st.html', 'w', encoding='utf8').write(page)
    t = time.time()
    out = subprocess.run([EDGE, '--headless=new', '--disable-gpu', '--no-first-run', '--user-data-dir=C:\\Users\\u9754\\Downloads\\MM_SharePoint\\_stp', '--virtual-time-budget=900000', '--dump-dom', 'file:///C:/Users/u9754/Downloads/MM_SharePoint/_st.html'], capture_output=True, timeout=900).stdout.decode('utf8', 'replace')
    m = re.search(r'<pre id="result">([\s\S]*?)</pre>', out)
    try: results[name] = json.loads(html.unescape(m.group(1))) if m else {'FAIL': 'no result'}
    except Exception as e: results[name] = {'FAIL': 'parse ' + (m.group(1)[:200] if m else '')}
    results[name]['_wall_s'] = round(time.time() - t, 1)
    print(name, json.dumps(results[name], ensure_ascii=False), flush=True)
os.remove(f'{WIN}/_st.html')
json.dump(results, open(f'{SP}/mmapp/test/stress_results.json', 'w', encoding='utf8'), ensure_ascii=False, indent=1)
