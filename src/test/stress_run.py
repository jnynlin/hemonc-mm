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
const decideNow=async(d,f)=>{ f.querySelector('button[type=submit]').click(); const ok=await until(()=>{ const b=d.getElementById('cfok'), c=d.getElementById('cf'); return ((b&&c&&!c.hidden)||/失敗|請/.test(d.getElementById('dmsg').textContent))&&(b||true); }); if(ok&&ok.click) ok.click(); return ok; };
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
  await decideNow(d,f); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000);
  OUT.decide=(d.querySelector('.flash')||d.getElementById('dmsg')).textContent.slice(0,120);
  OUT.thresholdHits=STRESS.thresholdHits; OUT.calls=STRESS.calls;''',
# 1b 正式紀錄也超過 5000：定案時查重複會碰門檻
'volume_records': r'''seed(6000,6000,10); MOCK.DB['案例'].slice(-3).forEach(ready); const sh=await boot();
  const {d}=await meet(sh); d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide'));
  await decideNow(d,f); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000);
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
  d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide')); const b=await decideNow(d,f); b.click(); b.click();
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
  const f=await until(()=>d.getElementById('decide')); await decideNow(d,f); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000);
  OUT.decide=(d.querySelector('.flash')||d.getElementById('dmsg')).textContent.slice(0,60); const rec=MOCK.DB['正式紀錄'].find(r=>r.Title===target)||{};
  OUT.recExtraKeys=Object.keys(rec).filter(k=>/Chart|Report|PatientName|Bed/.test(k)); OUT.recHasChart=JSON.stringify(rec).includes(tch);
  d.getElementById('ex').click(); await sleep(100); d.getElementById('show-chart').click(); await sleep(100); OUT.hideChartWorks=!d.querySelector('#v-prep td.chartno');
  const t2=rows()[0].querySelector('td.nw b').textContent; d.querySelector('input[data-rep="'+t2+'"]').click(); await until(()=>d.getElementById('only-rep')); d.getElementById('only-rep').click(); await sleep(100); OUT.onlyRepRows=rows().length;
  d.getElementById('go-present').click(); const f2=await until(()=>d.getElementById('decide')); await decideNow(d,f2); await until(()=>d.querySelector('.flash'),60000); await sleep(100);
  OUT.afterLastRep=d.querySelector('.flash').textContent.slice(-22); OUT.falseAllDone=/已全部定案$/.test((d.querySelector('.pres .wait')||{}).textContent||''); OUT.restShown=d.querySelectorAll('.rail button').length;''',
# 9 沒有對照表權限、站台還沒有「會中報告」欄位、案例唯讀：開會頁照常可用，只是不顯示對應功能
'meet_limited': r'''seed(6,0,0); MOCK.DB['案例'].forEach(ready); MOCK.perms['對照表']=false; MOCK.noReportField=true; const sh=await boot();
  const {d}=await meet(sh); OUT.prepRows=d.querySelectorAll('#v-prep tr.click').length; OUT.chartCol=!!d.querySelector('#v-prep td.chartno'); OUT.repCol=[...d.querySelectorAll('#v-prep th')].some(th=>th.textContent==='報告'); OUT.showChartToggle=!!d.getElementById('show-chart');
  d.getElementById('go-present').click(); await until(()=>d.querySelector('.pres h2')); OUT.presChart=!!d.querySelector('.pres .chartno'); OUT.presOk=d.querySelector('.pres h2').textContent.length>0; OUT.ruler=!!d.querySelector('.stay svg');''',
'meet_readonly': r'''seed(6,0,0); MOCK.DB['案例'].forEach(ready); MOCK.DB['案例'][1].mmReport=true; MOCK.permLow={'案例':1,'正式紀錄':1}; const sh=await boot();
  const {d}=await meet(sh); OUT.prepRows=d.querySelectorAll('#v-prep tr.click').length; OUT.boxes=d.querySelectorAll('input[data-rep]').length; OUT.ticks=[...d.querySelectorAll('#v-prep td.rep')].filter(t=>t.textContent==='✓').length;
  d.getElementById('go-present').click(); await until(()=>d.querySelector('.pres h2')); OUT.btnLabel=d.getElementById('dopen').textContent.replace('D',''); d.getElementById('dopen').click(); await sleep(100);
  OUT.drawerOpens=!d.getElementById('drawer').hidden; OUT.hasForm=!!d.getElementById('decide'); OUT.notice=d.getElementById('drawer').textContent.includes('定案由主任執行');''',
# 10 v1.20 根因選單：依案例類型只列對應層級，其餘收在「顯示…」裡；初評或主治已選的層級一律顯示；主要單選、促成複選
'rootcause_scope': r'''seed(3,0,0); const C=MOCK.DB['案例']; C.forEach(ready); C[1].mmCaseType=['Mortality']; C[1].mmInitMainFactor={Id:1,Title:'疾病嚴重度／多重共病'}; C[2].mmCaseType=['超長住院','Morbidity'];
  C[0].mmInitFactors=[{Id:2,Title:'檢驗檢查結果延遲或未追蹤'}]; const sh=await boot(); const {d}=await meet(sh); d.getElementById('go-present').click(); await until(()=>d.getElementById('decide'));
  const snap=()=>{ const f=d.getElementById('decide'), m=d.getElementById('rmore'); const vis=[...f.querySelectorAll('.rc:not([hidden]) details.rg summary')].map(s=>s.firstChild.textContent); return {code:d.querySelector('.drawer .dhead b').textContent.slice(-14), visible:vis, hidden:m?[...m.querySelectorAll('summary')].map(s=>s.firstChild.textContent):[], btn:(d.getElementById('rmore-btn')||{}).textContent||''}; };
  const out=[]; for(let k=0;k<3;k++){ out.push(snap()); if(k<2){ d.getElementById('nx').click(); await sleep(100); } } OUT.cases=out;
  d.getElementById('pv').click(); await sleep(50); d.getElementById('pv').click(); await sleep(100); const f=d.getElementById('decide');
  d.getElementById('rmore-btn').click(); OUT.revealed=!d.getElementById('rmore').hidden; const r=[...f.querySelectorAll('input[name=main]')]; OUT.radios=r.length; r[r.length-1].click(); OUT.oneMain=f.querySelectorAll('input[name=main]:checked').length;
  const cb=[...f.querySelectorAll('input[name=f]')].filter(x=>!x.disabled); cb[0].checked=true; cb[1].checked=true; f.dispatchEvent(new Event('change')); OUT.multiContrib=f.querySelectorAll('input[name=f]:checked').length>=2; OUT.mainBoxDisabled=f.querySelector('input[name=f][value="'+f.main.value+'"]').disabled;
  await decideNow(d,f); await until(()=>d.querySelector('.flash')||/失敗/.test(d.getElementById('dmsg').textContent),60000); const rec=MOCK.DB['正式紀錄'][0]||{}; OUT.savedMain=(rec.mmConsMainFactor||{}).Title; OUT.savedContrib=(rec.mmConsFactors||[]).length;''',
# 11 v1.21 柏拉圖切換：只算主要根因（每案一項）／所有因素（主要＋促成，同一案同一因素不重複計）
'pareto_toggle': r'''seed(12,12,0); const R=MOCK.DB['正式紀錄']; const now=new Date(), ym=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
  R.forEach((r,i)=>{ r.mmYearMonth=ym; r.mmMeetingDate=ym+'-15T00:00:00Z'; r.mmConsMainFactor=i<8?{Id:3,Title:'等待檢查／處置'}:{Id:4,Title:'感染併發症'}; r.mmConsFactors=i<6?[{Id:2,Title:'檢驗檢查結果延遲或未追蹤'},{Id:3,Title:'等待檢查／處置'}]:[]; });
  const sh=await boot(); const {d}=await meet(sh); d.querySelector('#nav button[data-view=records]').click(); await until(()=>d.getElementById('pareto'));
  const read=()=>[...d.querySelectorAll('#pareto svg rect')].map(x=>x.dataset.tip.replace(/（.*?）/,'').replace(/，累積.*$/,''));
  OUT.defaultMode=d.querySelector('[data-pm][aria-pressed=true]').dataset.pm; OUT.main=read(); OUT.mainCap=d.getElementById('pareto-cap').textContent.slice(0,14);
  d.querySelector('[data-pm=all]').click(); await sleep(100); OUT.all=read(); OUT.allCap=d.getElementById('pareto-cap').textContent; OUT.allPressed=d.querySelector('[data-pm=all]').getAttribute('aria-pressed');
  d.querySelector('[data-pm=main]').click(); await sleep(100); OUT.backSame=JSON.stringify(read())===JSON.stringify(OUT.main);''',
# 12 v1.22 定案確認：按「定案」只顯示確認內容不寫入；返回修改、改表單都會回到上一步；確認後才寫入
'confirm_step': r'''seed(3,0,0); MOCK.DB['案例'].forEach(ready); const sh=await boot(); const {d}=await meet(sh); d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide')); const cf=()=>d.getElementById('cf');
  f.act.click(); f.atitle.value='測試行動'; f.concl.value='結論一行'; f.querySelector('button[type=submit]').click(); await sleep(150);
  OUT.confirmShown=!cf().hidden; OUT.notWrittenYet=MOCK.DB['正式紀錄'].length===0&&MOCK.DB['改善行動'].length===0; OUT.summary=cf().textContent.replace(/\s+/g,' ').slice(0,150); OUT.submitHidden=f.querySelector('.go').hidden;
  d.getElementById('cfback').click(); await sleep(50); OUT.backHides=cf().hidden&&!f.querySelector('.go').hidden&&MOCK.DB['正式紀錄'].length===0;
  f.querySelector('button[type=submit]').click(); await sleep(100); f.concl.value='改過的結論'; f.concl.dispatchEvent(new Event('input',{bubbles:true})); await sleep(50); OUT.editHides=cf().hidden;
  f.querySelector('button[type=submit]').click(); await sleep(100); OUT.summaryUpdated=cf().textContent.includes('改過的結論'); d.getElementById('cfok').click(); await until(()=>d.querySelector('.flash'),60000);
  OUT.written=MOCK.DB['正式紀錄'].length; OUT.concl=MOCK.DB['正式紀錄'][0].mmConclusion; OUT.actions=MOCK.DB['改善行動'].length; OUT.flash=d.querySelector('.flash').textContent.slice(0,50);''',
# 13 v1.23 定案按鈕不可無聲無息：臨床決策層缺其他層級因素、漏選項目，都要在按鈕旁說明；補齊後可定案
'decide_feedback': r'''seed(3,0,0); const C=MOCK.DB['案例']; C.forEach(ready); C[0].mmInitMainFactor={Id:6,Title:'值班負荷與疲勞'}; C[0].mmInitFactors=[]; C[1].mmInitMainFactor=null; C[1].mmInitPrev=null;
  const sh=await boot(); const {d}=await meet(sh); d.getElementById('go-present').click(); let f=await until(()=>d.getElementById('decide')); const T=id=>d.getElementById(id).textContent; const sb=()=>f.querySelector('button[type=submit]');
  OUT.jcBtnEnabled=!sb().disabled; OUT.jcHintBeforeClick=T('dneed'); sb().click(); await sleep(100); OUT.jcNoConfirm=d.getElementById('cf').hidden; OUT.jcWritten=MOCK.DB['正式紀錄'].length;
  const other=[...f.querySelectorAll('input[name=f]')].find(x=>x.value==='等待檢查／處置'); other.click(); await sleep(50); OUT.jcHintAfterFix=T('dneed'); sb().click(); await sleep(100); OUT.jcConfirmAfterFix=!d.getElementById('cf').hidden;
  d.getElementById('cfback').click(); d.getElementById('dclose').click(); d.getElementById('nx').click(); await sleep(100); f=d.getElementById('decide');
  OUT.case2=d.querySelector('.drawer .dhead b').textContent.slice(-14); sb().click(); await sleep(100); OUT.missMsg=T('dmsg'); OUT.missNoConfirm=d.getElementById('cf').hidden;
  f.querySelector('input[name=main][value="感染併發症"]').click(); sb().click(); await sleep(100); OUT.missMsg2=T('dmsg'); f.prev.value='可能可預防'; f.prev.dispatchEvent(new Event('change',{bubbles:true}));
  await decideNow(d,f); await until(()=>d.querySelector('.flash'),60000); OUT.written=MOCK.DB['正式紀錄'].length; OUT.novalidate=d.getElementById('decide')?d.getElementById('decide').noValidate:'next-case-form';''',
# 14 v1.24 改善行動追蹤：在頁面更新狀態與進度、完成時填完成日期與成效、已完成區、尚未指派排最前、連結只收 https
'actions_track': r'''seed(3,3,0); MOCK.DB['案例'].forEach(ready); const A=MOCK.DB['改善行動']; const mk=(i,t,o,due,st,x={})=>A.push(Object.assign({Id:900+i,Title:t,mmOwner:o,mmDueDate:due+'T00:00:00Z',mmStatus:st},x));
  mk(1,'甲 已指派近期','本科','2099-01-10','待辦'); mk(2,'乙 尚未指派','會後指派','2099-06-01','待辦'); mk(3,'丙 已逾期','藥劑部','2020-01-01','進行中'); mk(4,'丁 已完成未確認','本科','2025-01-01','完成');
  const sh=await boot(); const {d}=await meet(sh); OUT.prepBtn=(d.getElementById('go-actions')||{}).textContent; d.getElementById('go-actions').click(); await until(()=>d.querySelector('#v-actions table'));
  const V=d.getElementById('v-actions'); const openRows=()=>[...V.querySelectorAll('section:first-child tbody tr:not(.aed)')].map(r=>r.cells[0].textContent.slice(0,7)+'|'+r.cells[1].textContent);
  OUT.order=openRows(); OUT.groups=[...V.querySelectorAll('h3.a-g')].map(h=>h.textContent.replace(/\s+/g,' ')); OUT.needRows=V.querySelectorAll('table.alist.need tbody tr:not(.aed)').length; OUT.calmRows=V.querySelectorAll('table.alist.calm tbody tr:not(.aed)').length; OUT.why=[...V.querySelectorAll('table.alist.need td.a-why')].map(t=>t.textContent.replace(/\d+ 天/,'N 天')); OUT.calmWhy=V.querySelector('table.alist.calm td.a-why').textContent.replace(/\d+/,'N'); OUT.head=V.querySelector('.sec-head .note').textContent; OUT.doneSummary=V.querySelector('#act-done summary').textContent.replace(/\s+/g,' ').trim();
  V.querySelector('[data-edit="901"]').click(); await sleep(50); let f=V.querySelector('form.aform'); OUT.doneFieldsHiddenAtFirst=f.querySelector('.adone').hidden;
  f.st.value='進行中'; f.st.dispatchEvent(new Event('change')); f.nd.value='已排入科會'; f.ow.value='病房護理'; f.querySelector('button[type=submit]').click(); await until(()=>/已儲存/.test((d.getElementById('act-msg')||{}).textContent||''));
  const r1=A.find(x=>x.Id===901); OUT.save1={st:r1.mmStatus,note:r1.mmRemarks,ow:r1.mmOwner,due:String(r1.mmDueDate).slice(0,10)};
  V.querySelector('[data-edit="901"]').click(); await sleep(50); f=V.querySelector('form.aform'); f.st.value='完成'; f.st.dispatchEvent(new Event('change')); OUT.doneFieldsShown=!f.querySelector('.adone').hidden;
  f.ef.value='有效'; f.en.value='三個月未再發生'; f.ev.value='javascript:alert(1)'; f.querySelector('button[type=submit]').click(); await sleep(100); OUT.badLinkMsg=f.querySelector('.amsg').textContent; OUT.badLinkNotSaved=r1.mmStatus==='進行中';
  f.ev.value='https://example.org/sop.pdf'; f.querySelector('button[type=submit]').click(); await until(()=>r1.mmStatus==='完成'); await sleep(100);
  OUT.save2={st:r1.mmStatus,eff:r1.mmEffect,done:String(r1.mmDoneDate).slice(0,4)===String(new Date().getFullYear()),ev:r1.mmEvidence}; OUT.movedToDone=!openRows().some(x=>x.startsWith('甲'))&&V.querySelector('#act-done').open;
  OUT.doneLink=!!V.querySelector('#act-done a[href="https://example.org/sop.pdf"]'); OUT.pendingCount=V.querySelector('#act-done summary').textContent.replace(/\s+/g,' ').trim();''',
# 15 改善行動：站台還沒有追蹤欄位（未跑 create_lists v15）仍可改狀態；唯讀帳號沒有「更新」鈕
'actions_limited': r'''seed(2,0,0); const A=MOCK.DB['改善行動']; A.push({Id:901,Title:'甲',mmOwner:'本科',mmDueDate:'2099-01-10T00:00:00Z',mmStatus:'待辦'}); MOCK.noActionV15=true; const sh=await boot(); const {d}=await meet(sh);
  d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.querySelector('#v-actions table')); const V=d.getElementById('v-actions'); V.querySelector('[data-edit="901"]').click(); await sleep(50); const f=V.querySelector('form.aform');
  f.st.value='完成'; f.st.dispatchEvent(new Event('change')); OUT.noExtraInputs=!f.dd&&!f.ef&&!f.ev; f.querySelector('button[type=submit]').click(); await until(()=>A[0].mmStatus==='完成'); OUT.saved=A[0].mmStatus; await sleep(100);
  OUT.doneCols=[...V.querySelectorAll('#act-done th')].map(t=>t.textContent).join('|');''',
'actions_readonly': r'''seed(2,0,0); MOCK.DB['改善行動'].push({Id:901,Title:'甲',mmOwner:'本科',mmDueDate:'2099-01-10T00:00:00Z',mmStatus:'待辦'}); MOCK.permLow={'改善行動':1}; const sh=await boot(); const {d}=await meet(sh);
  d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.querySelector('#v-actions table')); OUT.rows=d.querySelectorAll('#v-actions tbody tr').length; OUT.editButtons=d.querySelectorAll('#v-actions [data-edit]').length;''',
# 16 v1.24 每月存檔匯出：含當月定案與行動，不含病歷號；逾期提醒信草稿
'export_month': r'''seed(6,6,0); const R=MOCK.DB['正式紀錄']; const ym='2026-10'; R.forEach((r,i)=>{ r.mmYearMonth=i<4?ym:'2026-09'; r.Title=(i<4?'LOS-2026-10-0':'LOS-2026-09-0')+(i+1); r.mmMeetingDate=(i<4?'2026-10-20':'2026-09-22')+'T00:00:00Z'; r.mmConclusion='結論'+i; r.mmPatientShort='王OO'; });
  MOCK.DB['改善行動'].push({Id:901,Title:'行動甲',mmOwner:'本科',mmDueDate:'2020-01-01T00:00:00Z',mmStatus:'進行中',mmRelatedRecords:[{Id:1,Title:'LOS-2026-10-01'}]},{Id:902,Title:'行動乙<img src=x onerror=PWNED=1>',mmOwner:'會後指派',mmDueDate:'2099-01-01T00:00:00Z',mmStatus:'待辦',mmRelatedRecords:[{Id:2,Title:'LOS-2026-09-05'}]});
  const sh=await boot(); const {d,w}=await meet(sh); d.querySelector('#nav button[data-view=records]').click(); await until(()=>d.getElementById('r-export')); OUT.btn=d.getElementById('r-export').textContent; OUT.btnEnabled=!d.getElementById('r-export').disabled;
  const h=w.MM_SP.dashboard.monthReport(ym); const doc=new DOMParser().parseFromString(h,'text/html'); OUT.title=doc.querySelector('h1').textContent; OUT.summary=doc.querySelector('p').textContent.slice(0,60);
  const tabs=doc.querySelectorAll('table'); OUT.recRows=tabs[0].querySelectorAll('tr').length-1; OUT.onlyThisMonth=!h.includes('LOS-2026-09-06'); OUT.monthActions=tabs[1].querySelectorAll('tr').length-1; OUT.openActions=tabs[2].querySelectorAll('tr').length-1;
  OUT.noChartNo=MOCK.DB['對照表'].every(k=>!h.includes(k.mmChartNo)); OUT.escaped=!doc.querySelector('img')&&h.includes('&lt;img'); OUT.hasStyle=!!doc.querySelector('style');
  d.getElementById('r-export').click(); await sleep(100); OUT.afterClick=d.getElementById('r-export-msg').textContent.slice(0,12);
  d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.getElementById('act-mail')); const href=decodeURIComponent(d.getElementById('act-mail').getAttribute('href')); OUT.mail=href.slice(0,30); OUT.mailListsOverdueOnly=href.includes('行動甲')&&!href.includes('行動乙'); OUT.pwned=!!(window.PWNED||w.PWNED);''',
# 17 佐證上傳：檔案進「改善行動佐證」的該行動資料夾，連結填入並可儲存；沒有文件庫時不顯示上傳；上傳失敗有訊息
'evidence_upload': r'''seed(2,0,0); const A=MOCK.DB['改善行動']; A.push({Id:901,Title:'甲',mmOwner:'本科',mmDueDate:'2099-01-10T00:00:00Z',mmStatus:'進行中'}); const sh=await boot(); const {d}=await meet(sh);
  d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.querySelector('#v-actions table')); const V=d.getElementById('v-actions'); V.querySelector('[data-edit="901"]').click(); await sleep(50); const f=V.querySelector('form.aform');
  f.st.value='完成'; f.st.dispatchEvent(new Event('change')); OUT.hasUpload=!!f.evf; const put=async(name,bytes)=>{ const dt=new DataTransfer(); dt.items.add(new File([new Uint8Array(bytes)],name)); const m=f.querySelector('.amsg'); m.textContent=''; f.evf.files=dt.files; f.evf.dispatchEvent(new Event('change')); await until(()=>/已上傳|沒有上傳成功/.test(m.textContent)&&!f.querySelector('button[type=submit]').disabled,30000); };
  MOCK.uploadFail=true; await put('a.pdf',10); OUT.failMsg=f.querySelector('.amsg').textContent.slice(0,14); OUT.failNoLink=f.ev.value===''; MOCK.uploadFail=false;
  await put('修訂 SOP #2: 草稿?.pdf',2048); OUT.file=MOCK.files.map(x=>({folder:x.folder,name:x.name.replace(/^\d{8}-\d{6}_/,'T_'),size:x.size})); OUT.linkShape=f.ev.value.replace(location.origin,'ORIGIN').replace(/\d{8}-\d{6}_/,'T_'); OUT.noSpaceInLink=!/\s/.test(f.ev.value); await put('b.docx',100); OUT.twoLinks=f.ev.value.split(/\s+/).length;
  f.ev.value=f.ev.value.split(location.origin).join('https://x.sharepoint.com'); /* 測試頁是 file://，實際站台是 https */ f.ef.value='有效'; f.querySelector('button[type=submit]').click(); await until(()=>A[0].mmStatus==='完成'); await sleep(100); OUT.savedLinks=String(A[0].mmEvidence).split('\n').length; OUT.doneLinks=[...V.querySelectorAll('#act-done td a')].map(a=>a.textContent);''',
'evidence_nolib': r'''seed(2,0,0); MOCK.DB['改善行動'].push({Id:901,Title:'甲',mmOwner:'本科',mmDueDate:'2099-01-10T00:00:00Z',mmStatus:'完成'}); MOCK.noEvidenceLib=true; const sh=await boot(); const {d}=await meet(sh);
  d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.querySelector('#act-done')); d.getElementById('act-done').open=true; d.querySelector('[data-edit="901"]').click(); await sleep(50); const f=d.querySelector('form.aform'); OUT.hasLinkInput=!!f.ev; OUT.hasUpload=!!f.evf;''',
# 18 v1.25 進度結構化：一筆一筆累加（已做／下一步／追蹤日），舊的自由文字保留，到追蹤日標出，分隔符號不可被使用者文字破壞
'progress_log': r'''seed(2,0,0); const A=MOCK.DB['改善行動']; A.push({Id:901,Title:'甲',mmOwner:'本科',mmDueDate:'2099-01-10T00:00:00Z',mmStatus:'進行中',mmRemarks:'舊的自由文字備註'},{Id:902,Title:'乙',mmOwner:'本科',mmDueDate:'2099-02-10T00:00:00Z',mmStatus:'進行中',mmRemarks:'2020-01-05｜已做：開過會｜下一步：寫草稿｜追蹤：2020-02-01'});
  const sh=await boot(); const {d}=await meet(sh); d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.querySelector('#v-actions table')); const V=d.getElementById('v-actions'); const cell=t=>[...V.querySelectorAll('section:first-child tbody tr:not(.aed)')].find(r=>r.cells[0].textContent.startsWith(t)).cells[0].textContent.replace(/\s+/g,' ');
  OUT.legacyShown=cell('甲'); OUT.dueReview=cell('乙'); OUT.head=V.querySelector('.sec-head .note').textContent;
  const save=async(id,nd,nn,nr)=>{ V.querySelector('[data-edit="'+id+'"]').click(); await sleep(50); const f=V.querySelector('form.aform'); f.nd.value=nd; f.nn.value=nn; f.nr.value=nr; const before=A.find(x=>x.Id===id).mmRemarks; f.querySelector('button[type=submit]').click(); await until(()=>A.find(x=>x.Id===id).mmRemarks!==before||/存成功/.test(f.querySelector('.amsg').textContent)); await sleep(100); };
  await save(901,'完成 SOP｜初稿','送科會審','2099-03-01'); const L=String(A[0].mmRemarks).split('\n'); OUT.lines=L.length; OUT.keepsLegacy=L[0]==='舊的自由文字備註'; OUT.newLine=L[1].replace(/^\d{4}-\d{2}-\d{2}/,'DATE');
  await save(901,'','只填下一步',''); OUT.lines2=String(A[0].mmRemarks).split('\n').length; OUT.third=String(A[0].mmRemarks).split('\n')[2].replace(/^\d{4}-\d{2}-\d{2}/,'DATE');
  V.querySelector('[data-edit="901"]').click(); await sleep(50); const f=V.querySelector('form.aform'); const b4=A[0].mmRemarks; f.st.value='待辦'; f.querySelector('button[type=submit]').click(); await until(()=>A[0].mmStatus==='待辦'); OUT.noEntryWhenBlank=A[0].mmRemarks===b4;
  { V.querySelector('[data-edit="902"]').click(); await sleep(50); const g=V.querySelector('form.aform'); OUT.stepOptions=g.ndk.options.length; const R2=()=>A.find(x=>x.Id===902).mmRemarks, b0=R2();
    g.ndk.value='其他'; g.ndk.dispatchEvent(new Event('change')); OUT.otherPlaceholder=g.nd.placeholder; g.querySelector('button[type=submit]').click(); await sleep(100); OUT.otherNeedsText=g.querySelector('.amsg').textContent; OUT.otherNotSaved=R2()===b0;
    g.nd.value='自訂內容'; g.nnk.value='教育訓練'; g.nn.value='晨會宣導'; g.querySelector('button[type=submit]').click(); await until(()=>R2()!==b0); OUT.composed=R2().split('\n').pop().replace(/^\d{4}-\d{2}-\d{2}/,'DATE'); await sleep(100);
    V.querySelector('[data-edit="902"]').click(); await sleep(50); const h=V.querySelector('form.aform'), b1=R2(); h.ndk.value='召開討論會'; h.querySelector('button[type=submit]').click(); await until(()=>R2()!==b1); OUT.optionOnly=R2().split('\n').pop().replace(/^\d{4}-\d{2}-\d{2}/,'DATE'); await sleep(100); OUT.shown=cell('乙').slice(0,40); }
  OUT.afterCell=cell('甲'); OUT.reviewKeptAfterLaterEntry=/下次追蹤 2099-03-01/.test(cell('甲')); await save(902,'已補追蹤','',''); OUT.overdueReviewClearedByNewEntry=!/該追蹤了/.test(cell('乙')); OUT.historyCount=(V.querySelector('.nlog summary')||{}).textContent; OUT.headAfter=V.querySelector('.sec-head .note').textContent;''',
# 19 v1.28 主持投票不需定案權限：沒有定案權限但讀得到全部選票的帳號可開始與結束投票，結果留在畫面；讀不到全部選票的帳號不給投票鈕
'vote_host': r'''seed(2,0,0); MOCK.DB['案例'].forEach(ready); MOCK.permLow={'正式紀錄':1}; const sh=await boot(); const {d,w}=await meet(sh); OUT.voteState=w.MM_SP.voteState; OUT.canDecide=w.MM_SP.canDecide;
  d.getElementById('go-present').click(); await until(()=>d.getElementById('dopen')); d.getElementById('dopen').click(); await sleep(50); OUT.noForm=!d.getElementById('decide'); OUT.hasStart=!!d.getElementById('vstart');
  d.getElementById('vstart').click(); await until(()=>d.querySelector('.vote-ov .vqr svg')); const P=['不可預防','可能可預防','很可能可預防','明確可預防']; let vid=20000; for(let a=1;a<=9;a++) MOCK.DB['投票'].push({Id:vid++,AuthorId:a,mmVotePrev:P[a<=5?1:2]});
  await until(()=>d.getElementById('vn').textContent==='9',30000); d.getElementById('vend').click(); await until(()=>d.getElementById('vx2'),30000); OUT.noApplyButton=!d.getElementById('vapply'); OUT.result=d.querySelector('.vote-ov').textContent.replace(/\s+/g,' ').slice(0,60);
  d.getElementById('vx2').click(); await sleep(50); OUT.vres=d.getElementById('vres').textContent; OUT.tagged=MOCK.DB['投票'].filter(v=>v.mmVoteCase).length; OUT.records=MOCK.DB['正式紀錄'].length;
  d.getElementById('dclose').click(); d.getElementById('nx').click(); await sleep(50); d.getElementById('pv').click(); await sleep(50); d.getElementById('dopen').click(); await sleep(50); OUT.resultKept=d.getElementById('vres').textContent.includes('共 9 票');''',
'vote_noperm': r'''seed(2,0,0); MOCK.DB['案例'].forEach(ready); MOCK.permLow={'投票':3}; const sh=await boot(); const {d,w}=await meet(sh); OUT.voteState=w.MM_SP.voteState; OUT.canDecide=w.MM_SP.canDecide;
  d.getElementById('go-present').click(); const f=await until(()=>d.getElementById('decide')); OUT.hasStart=!!d.getElementById('vstart'); OUT.msg=[...f.querySelectorAll('p.small.muted')].map(p=>p.textContent).filter(t=>/投票/.test(t))[0]; OUT.canStillDecide=!!f.querySelector('button[type=submit]');''',
# 26 v1.29 行動列顯示相關案例的診斷；附加檔案不必等到「完成」才看得到，進行中就能上傳並存下
'actions_case_attach': r'''seed(3,3,0); const A=MOCK.DB['改善行動'], R=MOCK.DB['正式紀錄']; R[0].mmDxDetail='Acute myeloid leukemia with <b>x</b> '+'很長的診斷'.repeat(20);
  A.push({Id:901,Title:'回顧並更新 SOP',mmOwner:'本科',mmDueDate:'2099-01-10T00:00:00Z',mmStatus:'待辦',mmRelatedRecords:[{Id:R[0].Id,Title:R[0].Title},{Id:R[1].Id,Title:R[1].Title}]},{Id:902,Title:'沒有連到案例',mmOwner:'本科',mmDueDate:'2099-02-10T00:00:00Z',mmStatus:'進行中',mmEvidence:'https://x.sharepoint.com/mmevidence/A902/20261101-101010_a.pdf\nhttps://x.sharepoint.com/mmevidence/A902/20261101-101011_b.pdf'},{Id:903,Title:'連到查不到的代碼',mmOwner:'本科',mmDueDate:'2099-03-10T00:00:00Z',mmStatus:'進行中',mmRelatedRecords:[{Id:1,Title:'GONE-01'}]});
  const sh=await boot(); const {d}=await meet(sh); d.querySelector('#nav button[data-view=actions]').click(); await until(()=>d.querySelector('#v-actions table')); const V=d.getElementById('v-actions'); const row=t=>[...V.querySelectorAll('section:first-child tbody tr:not(.aed)')].find(r=>r.querySelector('.a-t').textContent===t);
  const r1=row('回顧並更新 SOP'); OUT.caseLines=[...r1.querySelectorAll('.a-c > div')].map(x=>x.querySelector('.a-code').textContent+'|'+(x.querySelector('.a-dx')||{textContent:''}).textContent.slice(0,30)); OUT.dxCut=[...r1.querySelector('.a-dx').textContent].length; OUT.noHtml=!r1.querySelector('.a-dx b'); OUT.codeNotInMeta=!r1.querySelector('.a-m').textContent.includes(R[0].Title);
  OUT.noCaseLine=!row('沒有連到案例').querySelector('.a-c'); OUT.unknownCode=row('連到查不到的代碼').querySelector('.a-c').textContent; OUT.rowAttach=[...row('沒有連到案例').querySelectorAll('.a-ev a')].map(a=>a.textContent);
  V.querySelector('[data-edit="902"]').click(); await sleep(50); let f=V.querySelector('form.aform'); OUT.reeditKeepsTwo=f.ev.value.split(/\s+/).length; const e0=A[1].mmEvidence; f.nd.value='只加進度'; f.querySelector('button[type=submit]').click(); await until(()=>A[1].mmRemarks); await sleep(100); OUT.evidenceUntouched=A[1].mmEvidence===e0;
  V.querySelector('[data-edit="901"]').click(); await sleep(50); f=V.querySelector('form.aform'); OUT.status=f.st.value; OUT.doneHidden=f.querySelector('.adone').hidden; OUT.attachVisible=!!(f.evf&&f.evf.offsetParent&&f.ev.offsetParent);
  const dt=new DataTransfer(); dt.items.add(new File([new Uint8Array(500)],'SOP 草稿.docx')); const m=f.querySelector('.amsg'); f.evf.files=dt.files; f.evf.dispatchEvent(new Event('change')); await until(()=>/已上傳|沒有上傳成功/.test(m.textContent)&&!f.querySelector('button[type=submit]').disabled,30000); OUT.upMsg=m.textContent;
  f.ev.value=f.ev.value.split(location.origin).join('https://x.sharepoint.com'); f.querySelector('button[type=submit]').click(); await until(()=>A[0].mmEvidence); await sleep(100); OUT.savedWhileOpen={status:A[0].mmStatus,links:String(A[0].mmEvidence).split('\n').length,done:A[0].mmDoneDate||null}; OUT.rowAfter=[...row('回顧並更新 SOP').querySelectorAll('.a-ev a')].map(a=>a.textContent);
  V.querySelector('[data-edit="901"]').click(); await sleep(50); f=V.querySelector('form.aform'); f.ev.value='not-a-link'; f.querySelector('button[type=submit]').click(); await sleep(100); OUT.badLink=f.querySelector('.amsg').textContent;''',
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
