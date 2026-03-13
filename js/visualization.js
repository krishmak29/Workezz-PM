// ── VIZ (Screen 2 preview table) ──
function renderViz() {
  var p = processedData;
  var totalAmt = p.rfmRows.reduce(function(s,r){return s+r.totalReturned;},0);

  // SVG stat cards
  document.getElementById('statsBar').innerHTML =
    statCard(SVG.parts,   p.total,  'Parts',         '') +
    statCard(SVG.matched, p.exact,  'Matched',       'color:var(--success)') +
    statCard(SVG.excess,  p.excess, 'More Returned', 'color:#0369a1') +
    statCard(SVG.short,   p.short,  'Less Returned', 'color:var(--danger)') +
    statCard(SVG.nopo,    p.noPo,   'No PO',         'color:var(--muted)') +
    (p.orNumber ? statCard(SVG.tag, p.orNumber, 'OR Number', 'font-size:.78rem;letter-spacing:-.01em;') : '') +
    (totalAmt > 0 ? statCard(SVG.money, totalAmt.toFixed(0), 'Total Returned', 'font-size:.88rem;') : '');

  // Init filter counts
  document.getElementById('fc-all').textContent   = p.total;
  document.getElementById('fc-exact').textContent = p.exact;
  document.getElementById('fc-short').textContent = p.short;
  document.getElementById('fc-excess').textContent= p.excess;
  document.getElementById('fc-nopo').textContent  = p.noPo;
  document.getElementById('filterBar').style.display = '';

  renderTable();
}

function statCard(iconSvg, val, label, valStyle) {
  return '<div class="stat-card"><span class="stat-icon" style="color:var(--accent);">' + iconSvg + '</span>'
       + '<div><div class="stat-val"' + (valStyle ? ' style="' + valStyle + '"' : '') + '>' + val + '</div>'
       + '<div class="stat-label">' + label + '</div></div></div>';
}

function renderTable() {
  var p   = processedData;
  var q   = filterQuery.toLowerCase().trim();
  var rows = p.rfmRows.filter(function(r) {
    if (activeFilter !== 'all' && r.status !== activeFilter) return false;
    if (q && !(r.typeNo.toLowerCase().includes(q) || r.desc.toLowerCase().includes(q) ||
               r.make.toLowerCase().includes(q)   || r.poNo.toLowerCase().includes(q))) return false;
    return true;
  });

  var html = '<div class="section-block">'
    + '<div class="section-head"><h3>RFM Report</h3><span class="section-badge">' + rows.length + ' / ' + p.total + ' parts</span></div>'
    + '<div style="overflow-x:auto;"><table class="viz-table"><thead><tr>'
    + '<th>#</th><th>Item Req No</th><th>Make</th><th>Part No</th><th>Description</th>'
    + '<th>Scope</th><th>Section</th>'
    + '<th style="background:#0369a1;">Customer</th>'
    + '<th style="background:#0369a1;">OR No.</th>'
    + '<th>Proj Qty</th>'
    + '<th style="background:#7c3aed;">PO Qty</th>'
    + '<th>PO No.</th><th>Supplier</th>'
    + '<th style="background:#b45309;">Used Qty</th>'
    + '<th style="background:#15803d;">Return Qty</th>'
    + '<th style="background:#15803d;">% Ret</th>'
    + '<th>Unit Price</th>'
    + '<th style="background:#1e3a5f;">Total Amt Returned</th>'
    + '<th>Status</th>'
    + '</tr></thead><tbody>';

  if (!rows.length) {
    html += '<tr><td colspan="19" style="text-align:center;padding:32px;color:var(--muted);font-size:.82rem;">No rows match the current filter</td></tr>';
  }
  rows.forEach(function(r, i) {
    var noPo = r.status === 'no-po';
    var pctFmt   = r.poQty > 0    ? r.pctReturned.toFixed(1)+'%'  : '—';
    var totalFmt = r.unitPrice > 0 ? r.totalReturned.toFixed(2)    : '—';
    var usedFmt  = noPo ? '—' : r.usedQty;
    var poFmt    = noPo ? '—' : r.poQty;
    var upFmt    = r.unitPrice > 0 ? r.unitPrice.toFixed(2)        : '—';
    var rowOp    = noPo ? ' style="opacity:.75;"' : '';
    var pill = r.status === 'exact'  ? '<span class="pill pill-exact">Exact</span>'
             : r.status === 'short'  ? '<span class="pill pill-short">Short</span>'
             : r.status === 'excess' ? '<span class="pill pill-excess">Excess</span>'
             :                         '<span class="pill pill-no-po">No PO</span>';
    html += '<tr'+rowOp+'>'
      + '<td style="color:var(--muted);font-size:.68rem;">'+(i+1)+'</td>'
      + '<td class="td-mono" style="font-size:.68rem;color:var(--muted);">'+(r.itemReqNo||'—')+'</td>'
      + '<td style="font-size:.73rem;font-weight:600;">'+(r.make||'—')+'</td>'
      + '<td class="td-mono">'+r.typeNo+'</td>'
      + '<td style="font-size:.72rem;max-width:200px;word-break:break-word;">'+(r.desc||'—')+'</td>'
      + '<td class="td-center" style="font-size:.68rem;">'+(r.scope||'—')+'</td>'
      + '<td class="td-center" style="font-size:.68rem;">'+(r.section||'—')+'</td>'
      + '<td style="font-size:.71rem;font-weight:600;color:#0369a1;">'+(r.customer||'—')+'</td>'
      + '<td class="td-mono td-center" style="font-size:.68rem;color:#0369a1;">'+(r.orNumber||'—')+'</td>'
      + '<td class="td-center" style="color:var(--muted);">—</td>'
      + '<td class="td-center" style="font-weight:700;color:#7c3aed;">'+poFmt+'</td>'
      + '<td style="font-size:.68rem;font-family:\'IBM Plex Mono\',monospace;">'+(r.poNo||'—')+'</td>'
      + '<td style="font-size:.68rem;">'+(r.supplier||'—')+'</td>'
      + '<td class="td-center" style="font-weight:700;color:#b45309;">'+usedFmt+'</td>'
      + '<td class="td-center" style="font-weight:700;color:#15803d;">'+r.returnQty+'</td>'
      + '<td class="td-center" style="font-size:.72rem;">'+pctFmt+'</td>'
      + '<td class="td-center" style="font-size:.72rem;">'+upFmt+'</td>'
      + '<td class="td-center" style="font-weight:700;">'+totalFmt+'</td>'
      + '<td class="td-center">'+pill+'</td>'
      + '</tr>';
  });
  html += '</tbody></table></div></div>';
  document.getElementById('vizContent').innerHTML = html;
}

function renderExportSummary() {
  var p = processedData;
  var totalAmt = p.rfmRows.reduce(function(s,r){return s+r.totalReturned;},0);
  document.getElementById('exportSummary').innerHTML =
    '<div class="exp-item"><div class="exp-item-val">'+p.total+'</div><div class="exp-item-label">Parts</div></div>'+
    '<div class="exp-item"><div class="exp-item-val" style="color:var(--danger)">'+p.short+'</div><div class="exp-item-label">Short</div></div>'+
    '<div class="exp-item"><div class="exp-item-val" style="color:#0369a1">'+p.excess+'</div><div class="exp-item-label">Excess</div></div>'+
    '<div class="exp-item"><div class="exp-item-val" style="font-size:.9rem;">'+totalAmt.toFixed(0)+'</div><div class="exp-item-label">Total Returned</div></div>';
}
