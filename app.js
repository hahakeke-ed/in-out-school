
fetch("https://docs.google.com/spreadsheets/d/e/2PACX-1vSW-NQvM_hYkajd1ZYyzWWGbx0qnUFMC9rlWeVLwSla5Xwovivh4Xp9y_bN73fg4Ab4-uidhJP63rzx/pub?output=csv")
  .then(r=>r.text())
  .then(t=>{ document.body.innerHTML = '<pre>'+t.replace(/</g,'&lt;')+'</pre>'; })
  .catch(e=>{ document.body.textContent='CSV 로드 실패: '+e; });
