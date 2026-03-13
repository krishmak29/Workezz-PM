
// ── PROCESS ──
async function processData() {
  document.getElementById('processBtn').disabled = true;
  document.getElementById('processBtn').textContent = 'Processing...';

  var skipVals = ['fl','type/order no','make','description','item req id','section','customer',
                  'part no','part number','timewk','request','std'];
  function isSkip(s) {
    if (!s) return true;
    var sl = s.toLowerCase().trim();
    return skipVals.some(function(sv){ return sl === sv || sl.startsWith(sv); });
  }

  var res = await Promise.all([
    readExcelSheet(rfmFile, 'Return Mat Slip'),
    readExcelSheet(poFile,  'PO_Status')
  ]);
  var rfmData = res[0].data;
  var poData  = res[1].data;

 

  // ── RFM COLUMNS ──
  var rfmHdrIdx, rfmHdrRow;
  if (detectionSettings.autoRFM) {
    rfmHdrIdx = detectHeaderRow(rfmData, ['item req', 'make', 'part', 'description', 'scope', 'section', 'qty'], 4);
    rfmHdrRow = rfmHdrIdx >= 0 ? rfmData[rfmHdrIdx].map(function(v){ return String(v||'').toLowerCase().trim(); }) : null;
  } else {
    rfmHdrIdx = -1; rfmHdrRow = null; // force fixed-position mode
  }

  var rfmCI  = rfmHdrRow ? findCol(rfmHdrRow, 'item req')  : 0;  // Item Req No  → col A (0)
  var rfmCM  = rfmHdrRow ? findCol(rfmHdrRow, 'make')      : 1;  // Make         → col B (1)
  var rfmCT  = rfmHdrRow ? findCol(rfmHdrRow, 'part no')   : 2;  // Part No      → col C (2)
  var rfmCD  = rfmHdrRow ? findCol(rfmHdrRow, 'desc')      : 3;  // Description  → col D (3)
  var rfmCSc = rfmHdrRow ? findCol(rfmHdrRow, 'scope')     : 4;  // Scope        → col E (4)
  var rfmCSe = rfmHdrRow ? findCol(rfmHdrRow, 'section')   : 5;  // Section      → col F (5)

  // Return Qty: look for 'return qty' specifically, then last 'qty' col (avoids PO Qty / Project Qty)
  var rfmCQ = -1;
  if (rfmHdrRow) {
    rfmCQ = findCol(rfmHdrRow, 'return qty');
    if (rfmCQ < 0) rfmCQ = findCol(rfmHdrRow, 'return');
    if (rfmCQ < 0) {
      for (var qi = rfmHdrRow.length - 1; qi >= 0; qi--) {
        if (rfmHdrRow[qi].includes('qty') || rfmHdrRow[qi].includes('quantity')) { rfmCQ = qi; break; }
      }
    }
  }
  if (rfmCI  < 0) rfmCI  = 0;
  if (rfmCM  < 0) rfmCM  = 1;
  if (rfmCT  < 0) rfmCT  = 2;
  if (rfmCD  < 0) rfmCD  = 3;
  if (rfmCSc < 0) rfmCSc = 4;
  if (rfmCSe < 0) rfmCSe = 5;
  if (rfmCQ  < 0) rfmCQ  = 6;  // col G — original fixed position

  // Data start: skip sub-headers (FL rows) after header. Fixed fallback = row 13 (0-based 12)
  var rfmDR = rfmHdrIdx >= 0 ? skipSubHeaders(rfmData, rfmHdrIdx) + 1 : 13;

  // OR Number: scan first 10 rows
  var orNumber = '';
  for (var oi = 0; oi < Math.min(10, rfmData.length); oi++) {
    var or = rfmData[oi];
    for (var oc = 0; oc < or.length; oc++) {
      if (String(or[oc]||'').toLowerCase().includes('or no') || String(or[oc]||'').toLowerCase().includes('or number')) {
        orNumber = String(or[oc+1] || or[oc+2] || '').trim();
        break;
      }
    }
    if (orNumber) break;
  }
  if (!orNumber) orNumber = String((rfmData[4] || [])[1] || '').trim();

  // ── PO COLUMNS ──
  var poHdrIdx, poHdrRow;
  if (detectionSettings.autoPO) {
    poHdrIdx = detectHeaderRow(poData, ['part', 'po no', 'qty', 'supplier', 'price', 'customer'], 3);
    poHdrRow = poHdrIdx >= 0 ? poData[poHdrIdx].map(function(v){ return String(v||'').toLowerCase().trim(); }) : null;
  } else {
    poHdrIdx = -1; poHdrRow = null; // force fixed-position mode
  }

  var poCT    = poHdrRow ? findCol(poHdrRow, 'part no')    : 3;   // Part No       → col D (3)
  var poCCust = poHdrRow ? findCol(poHdrRow, 'customer')   : 11;  // Customer      → col L (11)
  var poCQ    = poHdrRow ? (findCol(poHdrRow,'po qty') >= 0 ? findCol(poHdrRow,'po qty') : findCol(poHdrRow,'qty')) : 14;
  var poCPN   = poHdrRow ? findCol(poHdrRow, 'po no')      : 15;  // PO No         → col P (15)
  var poCSup  = poHdrRow ? findCol(poHdrRow, 'supplier')   : 17;  // Supplier      → col R (17)
  var poCUP   = poHdrRow ? (findCol(poHdrRow,'unit price') >= 0 ? findCol(poHdrRow,'unit price') : findCol(poHdrRow,'price')) : 32;

  if (poCT    < 0) poCT    = 3;
  if (poCCust < 0) poCCust = 11;
  if (poCQ    < 0) poCQ    = 14;
  if (poCPN   < 0) poCPN   = 15;
  if (poCSup  < 0) poCSup  = 17;
  if (poCUP   < 0) poCUP   = 32;

  // Data start: skip FL sub-headers. Fixed fallback = row 5 (0-based 4)
  var poDR = poHdrIdx >= 0 ? skipSubHeaders(poData, poHdrIdx) + 1 : 5;

  // Build PO map — dual keys: exact-normalised AND fuzzy-normalised
  // Structure: { exactKey -> entry, fuzzyKey -> entry }
  var poMapExact = new Map();
  var poMapFuzzy = new Map();

  for (var pi = poDR - 1; pi < poData.length; pi++) {
    var pr = poData[pi];
    if (!pr || !pr.some(function(v){ return v !== ''; })) continue;
    var rawPT = String(pr[poCT] !== undefined && pr[poCT] !== null ? pr[poCT] : '');
    var ptExact = normPNExact(rawPT);
    var ptFuzzy = normPN(rawPT);
    if (!ptExact || isSkip(ptExact)) continue;

    var pq = pr[poCQ];
    pq = typeof pq === 'number' ? pq : (parseFloat(String(pq||'').replace(/,/g,''))||0);
    var up = pr[poCUP];
    up = typeof up === 'number' ? up : (parseFloat(String(up||'').replace(/,/g,''))||0);
    var customer = String(pr[poCCust] || '').trim();
    var poNo     = String(pr[poCPN]   || '').trim();
    var supplier = String(pr[poCSup]  || '').trim();

    // Store in exact map
    if (!poMapExact.has(ptExact)) poMapExact.set(ptExact, {customer:'', poQty:0, poNo:'', supplier:'', unitPrice:0});
    var e = poMapExact.get(ptExact);
    e.poQty += pq;
    if (!e.customer  && customer)  e.customer  = customer;
    if (!e.poNo      && poNo)      e.poNo      = poNo;
    if (!e.supplier  && supplier)  e.supplier  = supplier;
    if (!e.unitPrice && up > 0)    e.unitPrice = up;

    // Store in fuzzy map (only if key differs from exact, to avoid double-counting)
    if (ptFuzzy && ptFuzzy !== ptExact) {
      if (!poMapFuzzy.has(ptFuzzy)) poMapFuzzy.set(ptFuzzy, poMapExact.get(ptExact));
    }
  }

  // Build RFM rows
  var rfmRows = [];
  for (var ri = rfmDR - 1; ri < rfmData.length; ri++) {
    var rr = rfmData[ri];
    if (!rr || !rr.some(function(v){ return v !== ''; })) continue;
    var rawPN = String(rr[rfmCT] !== undefined && rr[rfmCT] !== null ? rr[rfmCT] : '').trim();
    if (!rawPN) continue;

    var retQty = rr[rfmCQ];
    retQty = typeof retQty === 'number' ? retQty : (parseFloat(String(retQty||'').replace(/,/g,''))||0);

    // ── MULTI-STRATEGY MATCHING (best-effort, in order of strictness) ──
    var pnExact = normPNExact(rawPN);
    var pnFuzzy = normPN(rawPN);

    var poEntry =
      // 1. Exact normalised (trim+lower)
      poMapExact.get(pnExact) ||
      // 2. Strip trailing .0 artefacts
      poMapExact.get(pnExact.replace(/\.0+$/, '')) ||
      // 3. Fuzzy (strip all non-alphanumeric)
      poMapFuzzy.get(pnFuzzy) ||
      poMapExact.get(pnFuzzy) ||
      // 4. Numeric parse (handles "123.0" ↔ "123")
      (!isNaN(pnFuzzy.replace(/[.\-]/g,'')) && (
        poMapExact.get(String(parseInt(pnExact))) ||
        poMapExact.get(String(parseFloat(pnExact)))
      )) || null;

    var poQty     = poEntry ? poEntry.poQty     : 0;
    var customer  = poEntry ? poEntry.customer  : '';
    var poNo      = poEntry ? poEntry.poNo      : '';
    var supplier  = poEntry ? poEntry.supplier  : '';
    var unitPrice = poEntry ? poEntry.unitPrice : 0;
    var usedQty       = poQty - retQty;
    var pctReturned   = poQty > 0 ? (retQty / poQty) * 100 : 0;
    var totalReturned = unitPrice * retQty;
    var status = !poEntry ? 'no-po' : retQty > poQty ? 'excess' : retQty < poQty ? 'short' : 'exact';

    rfmRows.push({
      itemReqNo: String(rr[rfmCI]  || '').trim(),
      make:      String(rr[rfmCM]  || '').trim(),
      typeNo:    rawPN,
      desc:      String(rr[rfmCD]  || '').trim(),
      scope:     String(rr[rfmCSc] || '').trim(),
      section:   String(rr[rfmCSe] || '').trim(),
      customer, orNumber, poQty, poNo, supplier, unitPrice,
      usedQty, returnQty:retQty, pctReturned, totalReturned, status
    });
  }

  var total  = rfmRows.length;
  var exact  = rfmRows.filter(function(r){return r.status==='exact';}).length;
  var excess = rfmRows.filter(function(r){return r.status==='excess';}).length;
  var short  = rfmRows.filter(function(r){return r.status==='short';}).length;
  var noPo   = rfmRows.filter(function(r){return r.status==='no-po';}).length;

  var projDetails = {
    pm:         (document.getElementById('inputPM').value         || '').trim(),
    hwTL:       (document.getElementById('inputHWTL').value       || '').trim(),
    clientName: (document.getElementById('inputClient').value     || '').trim(),
    orderValue: parseFloat(document.getElementById('inputOrderValue').value)  || null,
    projCost:   parseFloat(document.getElementById('inputProjCost').value)    || null,
    actualCost: parseFloat(document.getElementById('inputActualCost').value) || null
  };

  processedData = {rfmRows, orNumber, total, exact, excess, short, noPo, projDetails};
  document.getElementById('processBtn').disabled = false;
  document.getElementById('processBtn').textContent = 'Compare ›';
  renderViz();
  goTo(2);
}
