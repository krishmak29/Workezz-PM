// ── FILE HANDLERS ──
function handleRFM(input) {
  if (!input.files.length) return;
  processRFMFile(input.files[0]);
  input.value = '';
}
function processRFMFile(file) {
  rfmFile = file;
  var ext = rfmFile.name.split('.').pop().toUpperCase();
  document.getElementById('rfmEmptyZone').style.display = 'none';
  document.getElementById('rfmChip').style.display = 'flex';
  document.getElementById('rfmExt').textContent = ext;
  document.getElementById('rfmName').textContent = rfmFile.name;
  document.getElementById('rfmSize').textContent = fmtSz(rfmFile.size);
  showFileMsg('rfm', null, null);
  var bare = rfmFile.name.replace(/\.(xlsx|xls|xlsb)$/i,'');
  document.getElementById('exportFileName').value = 'RFM_' + bare.replace(/[^a-zA-Z0-9_\-]/g,'_').replace(/_+/g,'_').replace(/^_|_$/g,'');

  readExcelSheet(rfmFile, 'Return Mat Slip').then(function(res) {
    // Sheet validation
    var hasSheet = res.foundSheet === 'Return Mat Slip';
    if (!hasSheet) {
      var names = res.sheetNames.join(', ');
      if (res.sheetNames.some(function(n){ return n.toLowerCase().includes('po') || n.toLowerCase().includes('purchase'); })) {
        showFileMsg('rfm', 'error', '⚠ This looks like a PO file — sheets found: <strong>' + names + '</strong>. Expected "Return Mat Slip".');
      } else {
        showFileMsg('rfm', 'warn', '⚠ Sheet "Return Mat Slip" not found. Using largest sheet instead. Found: <strong>' + names + '</strong>');
      }
    }
    // Header auto-detect
    var keywords = ['part','return','make','section','scope','qty'];
    var hdrRow = detectHeaderRow(res.data, keywords);
    var count = hdrRow >= 0
      ? res.data.slice(hdrRow + 1).filter(function(r){return r.some(function(v){return v!=='';});}).length
      : res.data.slice(12).filter(function(r){return r.some(function(v){return v!=='';});}).length;
    document.getElementById('rfmPreview').style.display = '';
    document.getElementById('rfmPreview').innerHTML =
      '✓ <strong>' + count + ' RFM parts</strong> detected' +
      (hdrRow >= 0 ? ' · headers at row ' + (hdrRow+1) : '') +
      (hasSheet ? ' · sheet: Return Mat Slip' : ' · using fallback sheet');
  });
  checkReady();
}

function clearRFM() {
  rfmFile = null;
  document.getElementById('rfmEmptyZone').style.display = '';
  document.getElementById('rfmChip').style.display = 'none';
  document.getElementById('rfmPreview').style.display = 'none';
  document.getElementById('rfmInput').value = '';
  showFileMsg('rfm', null, null);
  checkReady();
}

function handlePO(input) {
  if (!input.files.length) return;
  processPOFile(input.files[0]);
  input.value = '';
}
function processPOFile(file) {
  poFile = file;
  var ext = poFile.name.split('.').pop().toUpperCase();
  document.getElementById('poEmptyZone').style.display = 'none';
  document.getElementById('poChip').style.display = 'flex';
  document.getElementById('poExt').textContent = ext;
  document.getElementById('poName').textContent = poFile.name;
  document.getElementById('poSize').textContent = fmtSz(poFile.size);
  showFileMsg('po', null, null);

  readExcelSheet(poFile, 'PO_Status').then(function(res) {
    var hasSheet = res.foundSheet === 'PO_Status';
    if (!hasSheet) {
      var names = res.sheetNames.join(', ');
      if (res.sheetNames.some(function(n){ return n.toLowerCase().includes('return') || n.toLowerCase().includes('rfm'); })) {
        showFileMsg('po', 'error', '⚠ This looks like an RFM file — sheets found: <strong>' + names + '</strong>. Expected "PO_Status".');
      } else {
        showFileMsg('po', 'warn', '⚠ Sheet "PO_Status" not found. Using largest sheet instead. Found: <strong>' + names + '</strong>');
      }
    }
    var keywords = ['part','po','qty','supplier','price','customer'];
    var hdrRow = detectHeaderRow(res.data, keywords);
    var count = hdrRow >= 0
      ? res.data.slice(hdrRow + 1).filter(function(r){return r.some(function(v){return v!=='';});}).length
      : res.data.slice(4).filter(function(r){return r.some(function(v){return v!=='';});}).length;
    document.getElementById('poPreview').style.display = '';
    document.getElementById('poPreview').innerHTML =
      '✓ <strong>' + count + ' PO entries</strong> detected' +
      (hdrRow >= 0 ? ' · headers at row ' + (hdrRow+1) : '') +
      (hasSheet ? ' · sheet: PO_Status' : ' · using fallback sheet');
  });
  checkReady();
}

function clearPO() {
  poFile = null;
  document.getElementById('poEmptyZone').style.display = '';
  document.getElementById('poChip').style.display = 'none';
  document.getElementById('poPreview').style.display = 'none';
  document.getElementById('poInput').value = '';
  showFileMsg('po', null, null);
  checkReady();
}

function checkReady() {
  var ok = !!(rfmFile && poFile);
  document.getElementById('processBtn').disabled = !ok;
  document.getElementById('s1hint').textContent =
    !rfmFile ? 'Upload RFM file to continue' :
    !poFile  ? 'Upload PO file to continue'  :
               'Both files ready — click Compare';
}

// ── READ EXCEL (xlsx / xls / xlsb — all via SheetJS) ──
function readExcelSheet(file, preferredSheet) {
  return new Promise(function(resolve) {
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var wb = XLSX.read(e.target.result, {type:'array', cellDates:false});
        var foundSheet = null;
        if (preferredSheet && wb.SheetNames.indexOf(preferredSheet) !== -1) {
          foundSheet = preferredSheet;
        }
        var ws = null;
        if (foundSheet) {
          ws = wb.Sheets[foundSheet];
        } else {
          var best = null, bestN = 0;
          for (var i = 0; i < wb.SheetNames.length; i++) {
            var s = wb.Sheets[wb.SheetNames[i]];
            var d = XLSX.utils.sheet_to_json(s,{header:1,defval:'',raw:true});
            var n = d.filter(function(r){return r.some(function(v){return v!=='';});}).length;
            if (n > bestN) { bestN = n; best = s; }
          }
          ws = best;
        }
        var data = XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});
        resolve({data: data, sheetNames: wb.SheetNames, foundSheet: foundSheet});
      } catch(err) { resolve({data:[], sheetNames:[], foundSheet:null}); }
    };
    reader.readAsArrayBuffer(file);
  });
}

