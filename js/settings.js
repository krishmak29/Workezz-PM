// ── DETECTION SETTINGS ──


function toggleSetting(key) {
  detectionSettings[key] = !detectionSettings[key];
  var on = detectionSettings[key];
  var ids = { autoRFM: 'rfm', autoPO: 'po' };
  var id = ids[key];
  var row   = document.getElementById('tog-' + id + '-row');
  var badge = document.getElementById('tog-' + id + '-badge');
  var track = document.getElementById('tog-' + id + '-track');
  row.classList.toggle('tog-on', on);
  track.classList.toggle('track-on', on);
  badge.textContent = on ? 'ON' : 'OFF';
  badge.className = 'tog-badge ' + (on ? 'badge-on' : 'badge-off');
  try { localStorage.setItem('wz-detect', JSON.stringify(detectionSettings)); } catch(e){}
}

function loadDetectionSettings() {
  try {
    var saved = localStorage.getItem('wz-detect');
    if (!saved) return;
    var obj = JSON.parse(saved);
    Object.keys(obj).forEach(function(k) {
      if (k in detectionSettings) {
        detectionSettings[k] = obj[k];
        if (!obj[k]) {
          var ids = { autoRFM: 'rfm', autoPO: 'po' };
          var id = ids[k];
          if (!id) return;
          document.getElementById('tog-' + id + '-row').classList.remove('tog-on');
          document.getElementById('tog-' + id + '-track').classList.remove('track-on');
          document.getElementById('tog-' + id + '-badge').textContent = 'OFF';
          document.getElementById('tog-' + id + '-badge').className = 'tog-badge badge-off';
        }
      }
    });
  } catch(e){}
}

// ── PERSIST ──
var PKEYS = ['inputPM','inputHWTL','inputClient','inputOrderValue','inputProjCost','inputActualCost'];
function saveSettings() {
  try {
    var obj = {};
    PKEYS.forEach(function(k){ var el=document.getElementById(k); if(el) obj[k]=el.value; });
    localStorage.setItem('wz-rfm-v2', JSON.stringify(obj));
  } catch(e){}
}
function loadSettings() {
  try {
    var saved = localStorage.getItem('wz-rfm-v2');
    if (!saved) return;
    var obj = JSON.parse(saved);
    PKEYS.forEach(function(k){ var el=document.getElementById(k); if(el&&obj[k]!==undefined&&obj[k]!=='') el.value=obj[k]; });
  } catch(e){}
}
function clearSavedData() {
  if (!confirm('Clear all saved project details (PM, client, costs)?')) return;
  try { localStorage.removeItem('wz-rfm-v2'); } catch(e){}
  PKEYS.forEach(function(k){ var el=document.getElementById(k); if(el) el.value=''; });
}
