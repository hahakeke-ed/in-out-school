const SHEET_CSV_URL="https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=408307557&single=true&output=csv";
const BASELINE_CSV_URL="https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?gid=1708903695&single=true&output=csv";

// ===== 공통 유틸 =====
const qs=s=>document.querySelector(s), fmt=n=>Number(n||0).toLocaleString();
const encode=d=>Object.keys(d).map(k=>encodeURIComponent(k)+'='+encodeURIComponent(d[k]??'')).join('&');
const asOf=qs('#asOfDate'), grade=qs('#gradeFilter'), gender=qs('#genderFilter'), resetBtn=qs('#resetBtn');
const statE=qs('#statEnrolled'), statI=qs('#statIn'), statO=qs('#statOut');
const tb=document.querySelector('#dataTable tbody'), pv=document.querySelector('#pivotTable tbody');
const form=qs('#transferForm'), localInfo=qs('#localInfo');
let charts={bar:null,pie:null}, _rows=[], _base=[];

function getLocal(){ try{return JSON.parse(localStorage.getItem('local_submissions_v1')||'[]')}catch(e){return[]} }
function setLocal(a){ localStorage.setItem('local_submissions_v1',JSON.stringify(a)); if(localInfo) localInfo.textContent=`로컬 임시저장 ${a.length}건`; }
function parseDate(s){ if(!s)return null; const [y,m,d]=String(s).split('-').map(Number); return (!y||!m||!d)?null:new Date(y,m-1,d); }
function within(r,d){ if(!d)return true; const t=parseDate(r.status_date); return !t||t<=d; }
function loadCSV(u){ return new Promise((res,rej)=>Papa.parse(u,{download:true,header:true,skipEmptyLines:true,complete:r=>res(r.data),error:rej})); }

async function init(){
  try{ _rows=await loadCSV(SHEET_CSV_URL+'&t='+Date.now()); }catch(e){ _rows=[]; }
  try{ _base=await loadCSV(BASELINE_CSV_URL+'&t='+Date.now()); }catch(e){ _base=[]; }
  asOf.valueAsDate=new Date(); setLocal(getLocal()); render();
}

function filters(){ return {asOf:asOf.value?new Date(asOf.value):null, grade:grade.value, gender:gender.value}; }
function merged(){ return [..._rows, ...getLocal()]; }
function apply(rs,f){ return rs.filter(r=>within(r,f.asOf)&&(!f.grade||String(r.grade)===String(f.grade))&&(!f.gender||r.gender===f.gender)); }
function normB(r){ const g=r.grade??r['학년']; const s=(r.gender??r['성별']??'').toString().toUpperCase().trim(); const c=r.count??r['인원수']??r['인원']??0; return {grade:Number(g),gender:s,count:Number(c)}; }

function compute(f){
  const b={1:{M:0,F:0},2:{M:0,F:0},3:{M:0,F:0},4:{M:0,F:0},5:{M:0,F:0},6:{M:0,F:0}};
  for(const r of _base){ const x=normB(r); if(b[x.grade]&&(x.gender==='M'||x.gender==='F')){ if((!f.grade||String(f.grade)===String(x.grade))&&(!f.gender||f.gender===x.gender)) b[x.grade][x.gender]+=x.count; } }
  let i=0,o=0; for(const r of apply(merged(),f)){ if(r.status==='transfer_in'){i++; if(b[r.grade]&&(r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]+=1; }
                                               else if(r.status==='transfer_out'){o++; if(b[r.grade]&&(r.gender==='M'||r.gender==='F')) b[r.grade][r.gender]-=1; } }
  const gc={1:0,2:0,3:0,4:0,5:0,6:0}, gC={M:0,F:0};
  for(let g=1; g<=6; g++){ gc[g]=b[g].M+b[g].F; gC.M+=b[g].M; gC.F+=b[g].F; }
  return { e:Object.values(gc).reduce((a,v)=>a+v,0), i, o, gc, gC, pv:b };
}

// ===== 렌더링 =====
function render(){
  const f=filters(), s=compute(f);
  statE.textContent=fmt(s.e); statI.textContent=fmt(s.i); statO.textContent=fmt(s.o);
  renderCharts(s); renderPivot(s.pv); renderTable(apply(merged(),f));
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

// ====== 내보내기 ======
function toAOA_Summary(s, f){ return [
  ['스냅샷 생성시각', new Date().toISOString()],
  ['기준일', f.asOf ? f.asOf.toISOString().slice(0,10) : '전체'],
  ['학년 필터', f.grade||'전체'],
  ['성별 필터', f.gender||'전체'],
  [],
  ['총 재학생', s.e],
  ['전입(누계)', s.i],
  ['전출(누계)', s.o],
]; }
function toAOA_GradeCounts(s){ return [['학년','인원'], ...[1,2,3,4,5,6].map(g=>[g+'학년', s.gc[g]||0])]; }
function toAOA_GenderCounts(s){ return [['성별','인원'], ['M', s.gC.M||0], ['F', s.gC.F||0]]; }
function toAOA_Pivot(s){ const rows=[['학년','남','여','합계']]; for(let g=1; g<=6; g++){ const M=s.pv[g]?.M||0, F=s.pv[g]?.F||0; rows.push([g+'학년', M, F, M+F]); } return rows; }
function toAOA_Events(rs){ const head=['학년','반','성별','상태','일자','이름','비고']; const rows=[head]; for(const r of rs) rows.push([r.grade||'',r.class||'',r.gender||'',r.status||'',r.status_date||'',r.name||'',r.note||'']); return rows; }
function downloadExcel(){ const f=filters(), s=compute(f), rs=apply(merged(),f);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA_Summary(s,f)), '요약');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA_GradeCounts(s)), '학년별');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA_GenderCounts(s)), '성별');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA_Pivot(s)), '피벗');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA_Events(rs)), '전출입상세');
  const ts=new Date().toISOString().replace(/[:T]/g,'-').slice(0,16);
  XLSX.writeFile(wb, `dashboard_snapshot_${ts}.xlsx`);
}
function downloadCSV(){ const f=filters(), s=compute(f), rs=apply(merged(),f);
  const lines=[]; const push=(arr)=>lines.push(arr.map(v=>`"${(v??'').toString().replace(/"/g,'""')}"`).join(','));
  push(['스냅샷 생성시각', new Date().toISOString()]); push(['기준일', f.asOf?f.asOf.toISOString().slice(0,10):'전체']); lines.push('');
  push(['총 재학생', s.e]); push(['전입(누계)', s.i]); push(['전출(누계)', s.o]); lines.push('');
  push(['학년','남','여','합계']); for(let g=1; g<=6; g++){const M=s.pv[g]?.M||0,F=s.pv[g]?.F||0; push([g+'학년',M,F,M+F]);} lines.push('');
  push(['학년','반','성별','상태','일자','이름','비고']); for(const r of rs) push([r.grade||'',r.class||'',r.gender||'',r.status||'',r.status_date||'',r.name||'',r.note||'']);
  const blob=new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  const ts=new Date().toISOString().replace(/[:T]/g,'-').slice(0,16); a.download=`dashboard_snapshot_${ts}.csv`; a.click(); URL.revokeObjectURL(a.href);
}

document.getElementById('btnExportExcel').addEventListener('click', downloadExcel);
document.getElementById('btnExportCSV').addEventListener('click', downloadCSV);

// ===== 이벤트 =====
[asOf,grade,gender].forEach(el=>el.addEventListener('change',render));
resetBtn.addEventListener('click',()=>{ asOf.value=''; grade.value=''; gender.value=''; render(); });

form.addEventListener('submit',async e=>{ e.preventDefault(); const fd=new FormData(form);
  const row={student_id:'LOCAL-'+Date.now(), name:fd.get('name')||'', grade:Number(fd.get('grade')), class:Number(fd.get('class')), gender:fd.get('gender'),
              status:fd.get('status'), status_date:fd.get('status_date'), note:fd.get('note')||''};
  const a=getLocal(); a.push(row); setLocal(a); render();
  try{ await fetch('/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:encode({'form-name':'transfer',...Object.fromEntries(fd.entries())})}); alert('등록 완료!'); form.reset(); }
  catch(err){ console.error(err); alert('Netlify 전송 실패(로컬 통계는 반영됨)'); }
});

init();
