// ===== 데이터 소스 =====
const SHEET_CSV_URL="https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=408307557&single=true&output=csv";   // events(students)
const BASELINE_CSV_URL="https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=1708903695&single=true&output=csv"; // baseline('학년도' 열)

// ===== 학년도(3월 시작) =====
function academicYearFromDate(d){ const y=d.getFullYear(), m=d.getMonth()+1; return (m>=3)?y:(y-1); }

// ===== 유틸 =====
const qs=s=>document.querySelector(s), fmt=n=>Number(n||0).toLocaleString();
const asOf=qs('#asOfDate'), yearSel=qs('#yearSelect'), gSel=qs('#gradeFilter'), sexSel=qs('#genderFilter'), resetBtn=qs('#resetBtn');
const statE=qs('#statEnrolled'), statI=qs('#statIn'), statO=qs('#statOut');
const tb=document.querySelector('#dataTable tbody'), pv=document.querySelector('#pivotTable tbody');
const form=qs('#transferForm'), localInfo=qs('#localInfo');
const themeBtn=document.getElementById('themeToggle');
let rows=[], base=[], charts={bar:null,pie:null};

// 번역
function tStatus(s){ if(s==='transfer_in'||s==='전입') return '전입'; if(s==='transfer_out'||s==='전출') return '전출'; if(s==='enrolled') return '재학'; return s||''; }
function tGender(g){ const v=(g||'').toString().toUpperCase(); if(v==='M')return '남'; if(v==='F')return '여'; return v; }

function parseDate(s){ if(!s) return null; const [y,m,d]=String(s).split('-').map(Number); return (!y||!m||!d)?null:new Date(y,m-1,d); }
function within(r,asOf){ if(!asOf) return true; const d=parseDate(r.status_date); return !d||d<=asOf; }
function loadCSV(u){ return new Promise((res,rej)=>Papa.parse(u,{download:true,header:true,skipEmptyLines:true,complete:r=>res(r.data),error:rej})); }

function applyTheme(initial=false){ const saved=localStorage.getItem('theme')||'dark';
  document.body.classList.remove('light','dark'); document.body.classList.add(saved);
  themeBtn.textContent = saved==='dark' ? '🌙' : '☀️';
}
themeBtn.addEventListener('click',()=>{ const cur=document.body.classList.contains('dark')?'dark':'light'; const next=cur==='dark'?'light':'dark'; localStorage.setItem('theme', next); applyTheme(); });

async function init(){
  applyTheme(true);
  rows = await loadCSV(SHEET_CSV_URL+'&t='+Date.now()) || [];
  base = await loadCSV(BASELINE_CSV_URL+'&t='+Date.now()) || [];
  let years = base.map(r=>Number(r['학년도']||r['year'])).filter(Boolean);
  if(years.length===0){ const ds=rows.map(r=>parseDate(r.status_date)).filter(Boolean); years = ds.map(academicYearFromDate); }
  if(years.length===0) years=[new Date().getFullYear()];
  years=[...new Set(years)].sort((a,b)=>a-b);
  yearSel.innerHTML = years.map(y=>`<option value="${y}">${y}학년도</option>`).join('');
  const today = new Date(); asOf.valueAsDate=today; yearSel.value = String(academicYearFromDate(today));
  setLocal(getLocal()); render();
}

function filters(){ return { asOf: asOf.value?new Date(asOf.value):null, year:Number(yearSel.value)||null, grade:gSel.value, gender:sexSel.value }; }
function eventsFor(f){ return rows.filter(r=>within(r,f.asOf) && academicYearFromDate(parseDate(r.status_date)||new Date())===f.year && (!f.grade||String(r.grade)===String(f.grade)) && (!f.gender||r.gender===f.gender)); }
function normB(r){ const y=Number(r['학년도']||r.year||0), g=Number(r['학년']||r.grade||0), s=(r['성별']||r.gender||'').toString().toUpperCase().trim(), c=Number(r['인원수']||r.count||0); return {year:y,grade:g,gender:s,count:c}; }
function baselineFor(f){ const b={1:{M:0,F:0},2:{M:0,F:0},3:{M:0,F:0},4:{M:0,F:0},5:{M:0,F:0},6:{M:0,F:0}}; for(const r of base){ const x=normB(r); if(x.year===f.year && b[x.grade] && (x.gender==='M'||x.gender==='F')) b[x.grade][x.gender]+=x.count; } return b; }
function getLocal(){ try{return JSON.parse(localStorage.getItem('local_submissions_v1')||'[]')}catch(e){return[]} }
function setLocal(a){ localStorage.setItem('local_submissions_v1',JSON.stringify(a)); if(localInfo) localInfo.textContent=`로컬 임시저장 ${a.length}건`; }

function compute(f){
  const b=baselineFor(f); let i=0,o=0;
  for(const r of eventsFor(f)){ if(r.status==='transfer_in'){i++; if(b[r.grade]&&(r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]+=1; }
                               else if(r.status==='transfer_out'){o++; if(b[r.grade]&&(r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]-=1; } }
  const gc={1:0,2:0,3:0,4:0,5:0,6:0}, gC={M:0,F:0};
  for(let g=1; g<=6; g++){ gc[g]=b[g].M+b[g].F; gC.M+=b[g].M; gC.F+=b[g].F; }
  return {e:Object.values(gc).reduce((a,v)=>a+v,0), i, o, gc, gC, pv:b};
}

function render(){
  const f=filters(), s=compute(f);
  statE.textContent=fmt(s.e); statI.textContent=fmt(s.i); statO.textContent=fmt(s.o);
  renderCharts(s); renderPivot(s.pv); renderTable(eventsFor(f));
}

function renderPivot(p){ pv.innerHTML=''; let m=0,f=0; for(let g=1; g<=6; g++){ const M=p[g]?.M||0, F=p[g]?.F||0; m+=M; f+=F;
  pv.innerHTML+=`<tr><td>${g}학년</td><td>${fmt(M)}</td><td>${fmt(F)}</td><td>${fmt(M+F)}</td></tr>` } pv.innerHTML+=`<tr><td><b>합계</b></td><td><b>${fmt(m)}</b></td><td><b>${fmt(f)}</b></td><td><b>${fmt(m+f)}</b></td></tr>`; }

function renderTable(rs){ tb.innerHTML=''; const s=[...rs].sort((a,b)=>{ const da=parseDate(a.status_date)||new Date(0), db=parseDate(b.status_date)||new Date(0);
  if(db-da!==0) return db-da; if(Number(a.grade)-Number(b.grade)!==0) return Number(a.grade)-Number(b.grade); return Number(a.class)-Number(b.class); });
  for(const r of s) tb.innerHTML+=`<tr><td>${r.grade??''}</td><td>${r.class??''}</td><td>${tGender(r.gender)}</td><td>${tStatus(r.status)}</td><td>${r.status_date??''}</td><td>${r.name??''}</td><td>${r.note??''}</td></tr>`; }

function renderCharts(s){ if(charts.bar) charts.bar.destroy(); if(charts.pie) charts.pie.destroy();
  const color = getComputedStyle(document.body).getPropertyValue('--fg').trim() || '#e5e5e5';
  charts.bar=new Chart(document.getElementById('gradeBar'),{type:'bar',data:{labels:['1학년','2학년','3학년','4학년','5학년','6학년'],datasets:[{label:'재학생 수',data:[1,2,3,4,5,6].map(g=>s.gc[g]||0)}]},options:{responsive:true,plugins:{legend:{display:false}},scales:{x:{ticks:{color:color}},y:{ticks:{color:color}}}}});
  charts.pie=new Chart(document.getElementById('genderPie'),{type:'pie',data:{labels:['남','여'],datasets:[{data:[s.gC.M||0,s.gC.F||0]}]},options:{responsive:true,plugins:{legend:{labels:{color:color}}}}});
}

// ====== 내보내기 ======
function toAOA_Summary(s,f){ return [['학년도', f.year],['기준일', f.asOf?f.asOf.toISOString().slice(0,10):'전체'],['학년 필터', f.grade||'전체'],['성별 필터', f.gender? tGender(f.gender):'전체'],[],['총 재학생', s.e],['전입(누계)', s.i],['전출(누계)', s.o]]; }
function toAOA_GradeCounts(s){ return [['학년','인원'],...[1,2,3,4,5,6].map(g=>[g+'학년', s.gc[g]||0])]; }
function toAOA_GenderCounts(s){ return [['성별','인원'],['남', s.gC.M||0],['여', s.gC.F||0]]; }
function toAOA_Pivot(s){ const rows=[['학년','남','여','합계']]; for(let g=1; g<=6; g++){ const M=s.pv[g]?.M||0,F=s.pv[g]?.F||0; rows.push([g+'학년',M,F,M+F]); } return rows; }
function toAOA_Events(rs){ const head=['학년','반','성별','상태','일자','이름','비고']; const rows=[head]; for(const r of rs) rows.push([r.grade||'',r.class||'',tGender(r.gender)||'',tStatus(r.status)||'',r.status_date||'',r.name||'',r.note||'']); return rows; }
function downloadExcel(){ const f=filters(), s=compute(f), rs=eventsFor(f); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(toAOA_Summary(s,f)),'요약'); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(toAOA_GradeCounts(s)),'학년별'); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(toAOA_GenderCounts(s)),'성별'); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(toAOA_Pivot(s)),'피벗'); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(toAOA_Events(rs)),'전출입상세'); const ts=new Date().toISOString().replace(/[:T]/g,'-').slice(0,16); XLSX.writeFile(wb, `dashboard_${f.year}_${ts}.xlsx`); }
function downloadCSV(){ const f=filters(), s=compute(f), rs=eventsFor(f); const lines=[]; const push=a=>lines.push(a.map(v=>`"${(v??'').toString().replace(/"/g,'""')}"`).join(',')); push(['학년도', f.year]); push(['기준일', f.asOf?f.asOf.toISOString().slice(0,10):'전체']); lines.push(''); push(['총 재학생', s.e]); push(['전입(누계)', s.i]); push(['전출(누계)', s.o]); lines.push(''); push(['학년','남','여','합계']); for(let g=1; g<=6; g++){ const M=s.pv[g]?.M||0,F=s.pv[g]?.F||0; push([g+'학년',M,F,M+F]); } lines.push(''); push(['학년','반','성별','상태','일자','이름','비고']); for(const r of rs) push([r.grade||'',r.class||'',tGender(r.gender)||'',tStatus(r.status)||'',r.status_date||'',r.name||'',r.note||'']); const blob=new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); const ts=new Date().toISOString().replace(/[:T]/g,'-').slice(0,16); a.download=`dashboard_${f.year}_${ts}.csv`; a.click(); URL.revokeObjectURL(a.href); }

document.getElementById('btnExportExcel').addEventListener('click', downloadExcel);
document.getElementById('btnExportCSV').addEventListener('click', downloadCSV);

// ===== CORS-safe GAS 전송 (x-www-form-urlencoded) =====
function encodeForm(d){ return Object.keys(d).map(k=>encodeURIComponent(k)+'='+encodeURIComponent(d[k]??'')).join('&'); }
async function postToGAS(row){ const url = window.GAS_URL; if(!url) return {ok:false, reason:'NO_URL'};
  try{ const res = await fetch(url, { method:'POST', headers:{'Content-Type':'application/x-www-form-urlencoded'}, body: encodeForm(row) });
       return { ok: res.ok, status: res.status, text: await res.text() }; }
  catch(e){ return { ok:false, reason:String(e) }; } }

// ===== 이벤트 =====
[asOf,yearSel,gSel,sexSel].forEach(el=>el.addEventListener('change',render));
resetBtn.addEventListener('click',()=>{ asOf.value=''; gSel.value=''; sexSel.value=''; render(); });

form.addEventListener('submit', async e=>{ e.preventDefault(); const fd=new FormData(form);
  const row={student_id:'LOCAL-'+Date.now(), name:fd.get('name')||'', grade:Number(fd.get('grade')), class:Number(fd.get('class')), gender:fd.get('gender'),
              status:fd.get('status'), status_date:fd.get('status_date'), note:fd.get('note')||''};
  const a=getLocal(); a.push(row); setLocal(a); render();
  const r = await postToGAS(row);
  if(r.ok) alert('구글시트 반영 성공'); else alert('구글시트 반영 실패: '+(r.reason||r.status));
  if(r.ok) e.target.reset();
});

init();
