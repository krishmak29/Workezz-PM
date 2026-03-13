// ── DRAG & DROP ──
function dzOver(e, zoneId) {
  e.preventDefault(); e.stopPropagation();
  document.getElementById(zoneId).classList.add('drag-over');
}
function dzLeave(zoneId) {
  document.getElementById(zoneId).classList.remove('drag-over');
}
function dzDrop(e, type) {
  e.preventDefault(); e.stopPropagation();
  var zoneId = type === 'rfm' ? 'rfmEmptyZone' : 'poEmptyZone';
  document.getElementById(zoneId).classList.remove('drag-over');
  var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
  if (!file) return;
  var ext = file.name.split('.').pop().toLowerCase();
  if (!['xls','xlsx','xlsb'].includes(ext)) {
    showFileMsg(type, 'error', '⚠ Invalid file type. Please upload .xls, .xlsx, or .xlsb');
    return;
  }
  if (type === 'rfm') processRFMFile(file);
  else processPOFile(file);
}

function showFileMsg(type, kind, msg) {
  var errEl  = document.getElementById(type + 'Error');
  var warnEl = document.getElementById(type + 'Warn');
  errEl.style.display  = 'none';
  warnEl.style.display = 'none';
  if (!msg) return;
  if (kind === 'error') { errEl.innerHTML = msg;  errEl.style.display  = 'flex'; }
  else                  { warnEl.innerHTML = msg; warnEl.style.display = 'flex'; }
}
