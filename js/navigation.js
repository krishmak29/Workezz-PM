// ── NAV ──
function goTo(n) {
  document.querySelectorAll('.screen').forEach(function(s,i){ s.classList.toggle('active', i===n-1); });
  document.querySelectorAll('.step').forEach(function(s,i){
    s.classList.remove('active','done');
    if(i+1===n) s.classList.add('active');
    else if(i+1<n) s.classList.add('done');
  });
  if (n===3) renderExportSummary();
}