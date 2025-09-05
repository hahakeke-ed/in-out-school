// ===== 설정 =====
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=408307557&single=true&output=csv"; // students 탭 CSV
const BASELINE_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=1708903695&single=true&output=csv"; // baseline 탭 CSV

// ===== 유틸 =====
const qs = s=>document.querySelector(s), fmt=n=>Number(n||0).toLocaleString();
const encode = data => Object.keys(data).map(k=>encodeURIComponent(k)+'='+encodeURIComponent(data[k]??'')).join('&');
const asOfInput=qs('#asOfDate'), gradeSel=qs('#gradeFilter'), genderSel=qs('#genderFilter'), csvInput=qs('#csvFile'), resetBtn=qs('#resetBtn');
const statEnrolled=qs('#statEnrolled'), statIn=qs('#statIn'), statOut=qs('#statOut');
const tbody=document.querySelector('#dataTable tbody'), pivotBody=document.querySelector('#pivotTable tbody');
const form=qs('#transferForm'), clearLocalBtn=qs('#clearLocal'), localInfo=qs('#localInfo');
let csvRows=[], baselineRows=[];
function getLocal(){ try{return JSON.parse(localStorage.getItem('local_submissions_v1')||'[]')}catch(e){return[]}}
function setLocal(a){ localStorage.setItem('local_submissions_v1',JSON.stringify(a)); if(localInfo) localInfo.textContent=`로컬 임시저장 ${a.length}건`; }
function parseDate(s){ if(!s) return null; const [y,m,d]=String(s).split('-').map(Number); return (!y||!m||!d)?null:new Date(y,m-1,d); }
function withinAsOf(r,asOf){ if(!asOf) return true; const d=parseDate(r.status_date); return !d||d<=asOf; }
function loadCSV(url){ return new Promise((res,rej)=>{ Papa.parse(url,{download:true,header:true,skipEmptyLines:true,complete:r=>res(r.data),error:rej}); }); }
async function tryLoad(url){ try{ return await loadCSV(url+(url.includes('?')?'&':'?')+'t='+Date.now()); }catch(e){ return null; } }

// ===== 초기화 =====
async function init(){ csvRows = (await tryLoad(SHEET_CSV_URL)) || []; baselineRows = (await tryLoad(BASELINE_CSV_URL)) || []; asOfInput.valueAsDate = new Date(); setLocal(getLocal()); render(); }
function getFilters(){ return { asOf: asOfInput.value?new Date(asOfInput.value):null, grade: gradeSel.value, gender: genderSel.value }; }
function mergedRows(){ return [...csvRows, ...getLocal()]; }
function applyFilters(rows,{asOf,grade,gender}){ return rows.filter(r=>withinAsOf(r,asOf)&&(!grade||String(r.grade)===String(grade))&&(!gender||r.gender===gender)); }

// ===== baseline 파싱 =====
function normalizeBaselineRow(r){ const g = r.grade ?? r['학년']; const s = (r.gender ?? r['성별'] ?? '').toString().toUpperCase().trim(); const c = r.count ?? r['인원수'] ?? r['인원'] ?? 0; return { grade: Number(g), gender: s, count: Number(c) }; }

// ===== 통계 계산 =====
function computeWithBaseline(rows, baseRows, f){ const b = {1:{M:0,F:0},2:{M:0,F:0},3:{M:0,F:0},4:{M:0,F:0},5:{M:0,F:0},6:{M:0,F:0}};
  for(const raw of baseRows){ const {grade,gender,count} = normalizeBaselineRow(raw);
    if(b[grade] && (gender==='M'||gender==='F')){ if((!f.grade||String(f.grade)===String(grade)) && (!f.gender||f.gender===gender)){ b[grade][gender]+=Number(count||0); } }
  }
  let inCnt=0,outCnt=0; const evts=applyFilters(rows,f);
  for(const r of evts){ if(r.status==='transfer_in'){ inCnt++; if(b[r.grade] && (r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]+=1; }
                        else if(r.status==='transfer_out'){ outCnt++; if(b[r.grade] && (r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]-=1; } }
  const gradeCounts={1:0,2:0,3:0,4:0,5:0,6:0}, genderCounts={M:0,F:0};
  for(let g=1; g<=6; g++){ gradeCounts[g]=b[g].M+b[g].F; genderCounts.M+=b[g].M; genderCounts.F+=b[g].F; }
  const enrolledCount=Object.values(gradeCounts).reduce((a,v)=>a+v,0);
  return { enrolledCount, transferInCount: inCnt, transferOutCount: outCnt, gradeCounts, genderCounts, pivot:b };
}

// ===== 렌더링 =====
function render(){ const f=getFilters(), rows=mergedRows();
  const st = (baselineRows && baselineRows.length>0) ? computeWithBaseline(rows, baselineRows, f)
                                                     : { enrolledCount:0, transferInCount:0, transferOutCount:0, gradeCounts:{1:0,2:0,3:0,4:0,5:0,6:0}, genderCounts:{M:0,F:0}, pivot:{1:{M:0,F:0},2:{M:0,F:0},3:{M:0,F:0},4:{M:0,F:0},5:{M:0,F:0},6:{M:0,F:0}} };
  statEnrolled.textContent = fmt(st.enrolledCount); statIn.textContent=fmt(st.transferInCount); statOut.textContent=fmt(st.transferOutCount);
  renderCharts(st); renderPivot(st.pivot); renderTable(applyFilters(rows,f)); }

function renderPivot(pv){ pivotBody.innerHTML=''; let m=0,f=0;
  for(let g=1; g<=6; g++){ const M=pv[g]?.M||0, F=pv[g]?.F||0; m+=M; f+=F;
    pivotBody.innerHTML += `<tr><td>${g}학년</td><td>${fmt(M)}</td><td>${fmt(F)}</td><td>${fmt(M+F)}</td></tr>`; }
  pivotBody.innerHTML += `<tr><td><b>합계</b></td><td><b>${fmt(m)}</b></td><td><b>${fmt(f)}</b></td><td><b>${fmt(m+f)}</b></td></tr>`; }

function renderCharts(st){ const gradeCtx=document.getElementById('gradeBar'), genderCtx=document.getElementById('genderPie');
  const labels=['1학년','2학년','3학년','4학년','5학년','6학년']; const data=[1,2,3,4,5,6].map(g=>st.gradeCounts[g]||0);
  if(window.gradeChart) window.gradeChart.destroy();
  window.gradeChart=new Chart(gradeCtx,{type:'bar',data:{labels,datasets:[{label:'재학생 수',data}]},options:{responsive:true,plugins:{legend:{display:false}}}});
  if(window.genderChart) window.genderChart.destroy();
  window.genderChart=new Chart(genderCtx,{type:'pie',data:{labels:['남(M)','여(F)'],datasets:[{data:[st.genderCounts.M||0, st.genderCounts.F||0]}]},options:{responsive:true}}); }

function renderTable(rows){ const tb=document.querySelector('#dataTable tbody'); tb.innerHTML='';
  const sorted=[...rows].sort((a,b)=>{ const da=parseDate(a.status_date)||new Date(0), db=parseDate(b.status_date)||new Date(0);
    if(db-da!==0) return db-da; if(Number(a.grade)-Number(b.grade)!==0) return Number(a.grade)-Number(b.grade); return Number(a.class)-Number(b.class); });
  for(const r of sorted){ const tr=document.createElement('tr');
    tr.innerHTML=`<td>${r.grade??''}</td><td>${r.class??''}</td><td>${r.gender??''}</td><td>${r.status??''}</td><td>${r.status_date??''}</td><td>${r.name??''}</td><td>${r.note??''}</td>`;
    tb.appendChild(tr); } }

[asOfInput, gradeSel, genderSel].forEach(el=>el&&el.addEventListener('change',render));
resetBtn && resetBtn.addEventListener('click',()=>{ asOfInput.value=''; gradeSel.value=''; genderSel.value=''; render(); });
csvInput && csvInput.addEventListener('change',e=>{ const f=e.target.files?.[0]; if(!f) return; Papa.parse(f,{header:true,skipEmptyLines:true,complete:r=>{csvRows=r.data; render();}}) });
form && form.addEventListener('submit', async e=>{ e.preventDefault();
  const fd=new FormData(form);
  const row={student_id:'LOCAL-'+Date.now(),name:fd.get('name')||'',grade:Number(fd.get('grade')),class:Number(fd.get('class')),gender:fd.get('gender'),status:fd.get('status'),status_date:fd.get('status_date'),note:fd.get('note')||''};
  const a=getLocal(); a.push(row); setLocal(a); render();
  try{ await fetch('/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:encode({'form-name':'transfer',...Object.fromEntries(fd.entries())})}); alert('등록 완료! (대시보드 반영 + Netlify Forms 전송)'); form.reset(); }
  catch(err){ console.error(err); alert('Netlify 전송은 실패했지만, 이 기기에서는 통계가 반영되었습니다.'); }
});
init();
