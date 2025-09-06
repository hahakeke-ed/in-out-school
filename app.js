const form=document.getElementById('transferForm'); const log=document.getElementById('log');
function encodeForm(d){ return Object.keys(d).map(k=>encodeURIComponent(k)+'='+encodeURIComponent(d[k]??'')).join('&'); }
form.addEventListener('submit', async (e)=>{
  e.preventDefault();
  const fd=new FormData(form); const row=Object.fromEntries(fd.entries());
  row.student_id = 'LOCAL-'+Date.now(); // optional id
  // Send as x-www-form-urlencoded to avoid CORS preflight
  try{
    const res = await fetch(window.GAS_URL, {
      method:'POST',
      headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body: encodeForm(row)
    });
    const txt = await res.text();
    log.textContent = 'status: '+res.status+'\n'+txt;
    if(res.ok) alert('구글시트 반영 성공');
    else alert('구글시트 반영 실패: '+res.status);
  }catch(err){
    alert('구글시트 반영 실패: '+err);
    log.textContent = String(err);
  }
});