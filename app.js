const SHEET_CSV_URL="https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=408307557&single=true&output=csv"; // events (students 탭)
const BASELINE_CSV_URL="https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=1708903695&single=true&output=csv"; // baseline 탭 (학년도 열 지원)

// ===== 학년도 계산 (KR: 3월~다음해 2월) =====
function academicYearFromDate(d){ const y=d.getFullYear(); const m=d.getMonth()+1; return (m>=3)?y:(y-1); }

// ===== 유틸 =====
const qs=s=>document.querySelector(s), fmt=n=>Number(n||0).toLocaleString();
const asOf=qs('#asOfDate'), yearSel=qs('#yearSelect'), grade=qs('#gradeFilter'), gender=qs('#genderFilter');
const statE=qs('#statEnrolled'), statI=qs('#statIn'), statO=qs('#statOut');
const tb=document.querySelector('#dataTable tbody'), pv=document.querySelector('#pivotTable tbody');
let rows=[], base=[], charts={bar:null,pie:null};

function parseDate(s){ if(!s) return null; const [y,m,d]=String(s).split('-').map(Number); return (!y||!m||!d)?null:new Date(y,m-1,d); }
function within(r,asOf){ if(!asOf) return true; const d=parseDate(r.status_date); return !d||d<=asOf; }
function loadCSV(u){ return new Promise((res,rej)=>Papa.parse(u,{download:true,header:true,skipEmptyLines:true,complete:r=>res(r.data),error:rej})); }
function uniq(a){ return [...new Set(a)]; }

async function init(){
  rows = await loadCSV(SHEET_CSV_URL+'&t='+Date.now()) || [];
  base = await loadCSV(BASELINE_CSV_URL+'&t='+Date.now()) || [];
  // 학년도 목록 (baseline의 '학년도'가 있으면 사용, 없으면 rows의 status_date에서 계산)
  let years = base.map(r=>Number(r['학년도']||r['year'])).filter(Boolean);
  if(years.length===0){ years = rows.map(r=>parseDate(r.status_date)).filter(Boolean).map(academicYearFromDate); }
  if(years.length===0){ years=[new Date().getFullYear()]; }
  years = years.sort((a,b)=>a-b);
  yearSel.innerHTML = years.map(y=>`<option value="${y}">${y}학년도</option>`).join('');
  const today = new Date(); asOf.valueAsDate=today; yearSel.value = String(academicYearFromDate(today));
  render();
}

function filters(){ return { asOf: asOf.value?new Date(asOf.value):null, year: Number(yearSel.value)||null, grade: grade.value, gender: gender.value }; }
function applyEvents(rs,f){ return rs.filter(r=>within(r,f.asOf) && academicYearFromDate(parseDate(r.status_date)||new Date())===f.year && (!f.grade||String(r.grade)===String(f.grade)) && (!f.gender||r.gender===f.gender)); }
function normBaseline(r){ const y = Number(r['학년도']||r['year']||0); const g = Number(r['학년']||r.grade||0); const s = (r['성별']||r.gender||'').toString().toUpperCase().trim(); const c = Number(r['인원수']||r.count||0); return {year:y, grade:g, gender:s, count:c}; }
function baselineForYear(f){ const b={1:{M:0,F:0},2:{M:0,F:0},3:{M:0,F:0},4:{M:0,F:0},5:{M:0,F:0},6:{M:0,F:0}}; for(const r of base){ const x=normBaseline(r); if(x.year===f.year && b[x.grade] && (x.gender==='M'||x.gender==='F')) b[x.grade][x.gender]+=x.count; } return b; }

function compute(f){
  const b = baselineForYear(f);
  let i=0,o=0;
  for(const r of applyEvents(rows,f)){ if(r.status==='transfer_in'){i++; if(b[r.grade]&&(r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]+=1; }
                                      else if(r.status==='transfer_out'){o++; if(b[r.grade]&&(r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]-=1; } }
  const gc={1:0,2:0,3:0,4:0,5:0,6:0}, gC={M:0,F:0};
  for(let g=1; g<=6; g++){ gc[g]=b[g].M+b[g].F; gC.M+=b[g].M; gC.F+=b[g].F; }
  return { e:Object.values(gc).reduce((a,v)=>a+v,0), i, o, gc, gC, pv:b };
}

function render(){
  const f=filters(), s=compute(f);
  statE.textContent=fmt(s.e); statI.textContent=fmt(s.i); statO.textContent=fmt(s.o);
  renderCharts(s); renderPivot(s.pv); renderTable(applyEvents(rows,f));
}

function renderPivot(p){ pv.innerHTML=''; let m=0,f=0; for(let g=1; g<=6; g++){ const M=p[g]?.M||0, F=p[g]?.F||0; m+=M; f+=F;
  pv.innerHTML+=`<tr><td>${g}학년</td><td>${fmt(M)}</td><td>${fmt(F)}</td><td>${fmt(M+F)}</td></tr>` } pv.innerHTML+=`<tr><td><b>합계</b></td><td><b>${fmt(m)}</b></td><td><b>${fmt(f)}</b></td><td><b>${fmt(m+f)}</b></td></tr>`; }

function renderTable(rs){ tb.innerHTML=''; const s=[...rs].sort((a,b)=>{ const da=parseDate(a.status_date)||new Date(0), db=parseDate(b.status_date)||new Date(0);
  if(db-da!==0) return db-da; if(Number(a.grade)-Number(b.grade)!==0) return Number(a.grade)-Number(b.grade); return Number(a.class)-Number(b.class); });
  for(const r of s) tb.innerHTML+=`<tr><td>${r.grade??''}</td><td>${r.class??''}</td><td>${r.gender??''}</td><td>${r.status??''}</td><td>${r.status_date??''}</td><td>${r.name??''}</td><td>${r.note??''}</td></tr>`; }

function renderCharts(s){ if(charts.bar) charts.bar.destroy(); if(charts.pie) charts.pie.destroy();
  charts.bar=new Chart(document.getElementById('gradeBar'),{type:'bar',data:{labels:['1학년','2학년','3학년','4학년','5학년','6학년'],datasets:[{label:'재학생 수',data:[1,2,3,4,5,6].map(g=>s.gc[g]||0)}]},options:{responsive:true,plugins:{legend:{display:false}}}});
  charts.pie=new Chart(document.getElementById('genderPie'),{type:'pie',data:{labels:['남(M)','여(F)'],datasets:[{data:[s.gC.M||0,s.gC.F||0]}]},options:{responsive:true}});
}

[asOf,yearSel,grade,gender].forEach(el=>el.addEventListener('change',render));
init();
