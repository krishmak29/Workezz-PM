// ── EXPORT ──
async function exportExcel() {
  var btn = document.getElementById('exportBtn');
  btn.textContent = 'Generating...'; btn.disabled = true;
  await new Promise(function(r){setTimeout(r,50);});

  try {
    var wb = new ExcelJS.Workbook();
    var p  = processedData;
    var pd = p.projDetails;

    var C = {
      DB:'FF1E3A5F', WW:'FFFFFFFF',
      BL:'FF0369A1', PU:'FF7C3AED',
      AM:'FF92400E', GR:'FF15803D',
      GD:'FFFFE598'   // gold — Summary header fill (Book3 style)
    };

    // ══════════════════════════════════════
    // SHEET 1 — RFM Report
    // Row 1: OR Number | Row 2: Total Return Cost | Row 3: Headers | Row 4+: Data
    // ══════════════════════════════════════
    var ws = wb.addWorksheet('RFM Report');
    [5,13,16,24,38,8,12,20,20,10,10,20,24,10,10,9,12,18]
      .forEach(function(w,i){ ws.getColumn(i+1).width = w; });
    ws.views = [{state:'frozen', xSplit:4, ySplit:3}];

    var totalReturnCost = p.rfmRows.reduce(function(s,r){return s+r.totalReturned;},0);

    // Row 1 — OR Number
    var r1 = ws.getRow(1); r1.height = 18;
    r1.getCell(3).value = 'OR Number';
    r1.getCell(3).font  = {bold:true, size:10, name:'Calibri', color:{argb:C.DB}};
    r1.getCell(3).alignment = {vertical:'middle'};
    r1.getCell(4).value = p.orNumber || '';
    r1.getCell(4).font  = {bold:true, size:12, name:'Calibri', color:{argb:C.BL}};
    r1.getCell(4).alignment = {vertical:'middle',horizontal: 'center'};

    // Row 2 — Total Return Cost
    var r2 = ws.getRow(2); r2.height = 18;
    r2.getCell(3).value = 'Total Return Cost';
    r2.getCell(3).font  = {bold:true, size:10, name:'Calibri', color:{argb:C.DB}};
    r2.getCell(3).alignment = {vertical:'middle'};
    r2.getCell(4).value = totalReturnCost > 0 ? parseFloat(totalReturnCost.toFixed(2)) : null;
    r2.getCell(4).font  = {bold:true, size:14, name:'Calibri', color:{argb:'FFDC2626'}};
    r2.getCell(4).numFmt = '#,##0.00';
    r2.getCell(4).alignment = {vertical:'middle',horizontal: 'center'};

    // Row 3 — Headers
    var hdrs = ['SR','Item Req No','Make','Part No','Description','Scope','Section',
                'Customer','OR Number','Project Qty','PO Qty','PO No.','Supplier',
                'Used Qty','Return Qty','% Returned','Unit Price','Total Amt Returned'];
    var hdrBg = [C.DB,C.DB,C.DB,C.DB,C.DB, C.DB,C.DB,
                 C.BL,C.BL,C.DB, C.PU,C.DB,C.DB,
                 C.AM,C.GR,C.GR, C.DB,C.DB];
    var hdrRow = ws.getRow(3); hdrRow.height = 24;
    hdrs.forEach(function(v,i){
      var c = hdrRow.getCell(i+1);
      c.value = v;
      c.font  = {bold:true, size:9, color:{argb:C.WW}, name:'Calibri'};
      c.fill  = {type:'pattern', pattern:'solid', fgColor:{argb:hdrBg[i]}};
      c.alignment = {vertical:'middle', horizontal:'center', wrapText:true};
      c.border = {bottom:{style:'medium',color:{argb:'FFD1D5DB'}}, right:{style:'thin',color:{argb:'FFD1D5DB'}}};
    });

    // Rows 4+ — Data
    p.rfmRows.forEach(function(r, ri) {
      var dr   = ws.addRow([]); dr.height = 17;
      var noPo = r.status === 'no-po';
      var rbg  = ri%2===0 ? 'FFFFFFFF' : 'FFF8FAFC';
      var pctVal   = r.poQty > 0    ? parseFloat(r.pctReturned.toFixed(2))    : null;
      var totalVal = r.unitPrice > 0 ? parseFloat(r.totalReturned.toFixed(2)) : null;
      var upVal    = r.unitPrice > 0 ? parseFloat(r.unitPrice.toFixed(2))     : null;
      var vals = [ri+1, r.itemReqNo||'', r.make||'', r.typeNo, r.desc||'',
                  r.scope||'', r.section||'', r.customer||'', r.orNumber||'', null,
                  noPo?null:r.poQty, r.poNo||'', r.supplier||'',
                  noPo?null:r.usedQty, r.returnQty, pctVal, upVal, totalVal];
      vals.forEach(function(v, ci) {
        var c = dr.getCell(ci+1);
        c.fill  = {type:'pattern', pattern:'solid', fgColor:{argb:rbg}};
        c.value = (v===''||v===undefined||v===null) ? null : v;
        var fg = 'FF1F2937';
        if (ci===3)  fg = C.DB;
        if (ci===7)  fg = C.BL;
        if (ci===8)  fg = C.BL;
        if (ci===10) fg = C.PU;
        if (ci===13) fg = C.AM;
        if (ci===14) fg = C.GR;
        if (ci===15) fg = C.GR;
        c.font      = {size:9, name:ci===3?'Courier New':'Calibri',
                       bold:(ci===3||ci===10||ci===13||ci===14||ci===17),
                       color:{argb:fg}};
        c.alignment = {vertical:'middle',
                       horizontal:(ci===0||ci===8||ci===9||ci===10||ci===13||ci===14||ci===15||ci===16||ci===17)?'center':'left',
                       wrapText:ci===4};
        if (ci===15 && v!==null) c.numFmt = '0.00"%"';
        if ((ci===16||ci===17) && v!==null) c.numFmt = '#,##0.00';
        c.border = {bottom:{style:'thin',color:{argb:'FFE5E7EB'}}, right:{style:'thin',color:{argb:'FFE5E7EB'}}};
      });
    });
    if (!p.rfmRows.length) {
      ws.addRow(['No RFM data found']).getCell(1).font = {italic:true, size:9, color:{argb:'FF64748B'}};
    }

    // ══════════════════════════════════════
    // SHEET 2 — Summary (Book3 gold style)
    // ══════════════════════════════════════
    var ws2 = wb.addWorksheet('Summary');
    var sumCols = [
      {r1:'Sr No',                      r2:null,                             w:8,  m:true},
      {r1:'PM',                          r2:null,                             w:16, m:true},
      {r1:'HW TL',                       r2:null,                             w:16, m:true},
      {r1:'Client name',                 r2:null,                             w:20, m:true},
      {r1:'Order ID',                    r2:null,                             w:18, m:true},
      {r1:'Order Value',                 r2:null,                             w:16, m:true},
      {r1:'Projected Material Cost',     r2:null,                             w:22, m:true},
      {r1:'Actual Material Cost',        r2:null,                             w:22, m:true},
      {r1:'% age Act Material cost',     r2:'wrt to Projected Material Cost', w:22, m:false},
      {r1:'Mfg. Return Material Cost',   r2:null,                             w:24, m:true},
      {r1:'% age Return Material Cost',  r2:'wrt to Actual Material Cost',    w:22, m:false},
      {r1:'% age Return Material Cost',  r2:'wrt to Projected Material Cost', w:22, m:false},
      {r1:'% age Return Material Cost',  r2:'wrt to Order Value',             w:22, m:false},
      {r1:'Remarks',                     r2:null,                             w:20, m:true}
    ];
    sumCols.forEach(function(col,i){ ws2.getColumn(i+1).width = col.w; });
    ws2.views = [{state:'frozen', xSplit:0, ySplit:2}];

    function hdrS(cell) {
      cell.font      = {bold:true, size:11, name:'Calibri', color:{argb:'FF000000'}};
      cell.fill      = {type:'pattern', pattern:'solid', fgColor:{argb:C.GD}};
      cell.alignment = {horizontal:'center', vertical:'middle', wrapText:true};
      cell.border    = {top:{style:'thin',color:{argb:'FF000000'}},bottom:{style:'thin',color:{argb:'FF000000'}},
                        left:{style:'thin',color:{argb:'FF000000'}},right:{style:'thin',color:{argb:'FF000000'}}};
    }
    var sh1 = ws2.getRow(1); sh1.height = 32;
    var sh2 = ws2.getRow(2); sh2.height = 32;
    sumCols.forEach(function(col,i){
      var c1 = sh1.getCell(i+1); c1.value = col.r1; hdrS(c1);
      if (col.m) {
        ws2.mergeCells(1, i+1, 2, i+1);
      } else {
        var c2 = sh2.getCell(i+1); c2.value = col.r2; hdrS(c2);
      }
    });

    // Data row 3
    var ov  = pd.orderValue || null;
    var pc  = pd.projCost   || null;
    var ac  = pd.actualCost || null;
    var mrc = totalReturnCost > 0 ? parseFloat(totalReturnCost.toFixed(2)) : null;
    var pctActWrtProj  = (ac && pc && pc>0)  ? parseFloat((ac  / pc  * 100).toFixed(2)) : null;
    var pctRetWrtAct   = (mrc && ac && ac>0) ? parseFloat((mrc / ac  * 100).toFixed(2)) : null;
    var pctRetWrtProj  = (mrc && pc && pc>0) ? parseFloat((mrc / pc  * 100).toFixed(2)) : null;
    var pctRetWrtOrder = (mrc && ov && ov>0) ? parseFloat((mrc / ov  * 100).toFixed(2)) : null;
    var sumVals = [1, pd.pm||'', pd.hwTL||'', pd.clientName||'', p.orNumber||'',
                   ov, pc, ac, pctActWrtProj, mrc,
                   pctRetWrtAct, pctRetWrtProj, pctRetWrtOrder, ''];
    var dr3 = ws2.getRow(3); dr3.height = 20;
    sumVals.forEach(function(v,i){
      var c = dr3.getCell(i+1);
      c.value = (v===null||v==='') ? null : v;
      c.font  = {size:11, name:'Calibri', color:{argb:'FF1F2937'}};
      c.alignment = {vertical:'middle', horizontal:(i===0||i>=5)?'center':'left'};
      c.border = {top:{style:'thin',color:{argb:'FFCCCCCC'}},bottom:{style:'thin',color:{argb:'FFCCCCCC'}},
                  left:{style:'thin',color:{argb:'FFCCCCCC'}},right:{style:'thin',color:{argb:'FFCCCCCC'}}};
      if ((i===5||i===6||i===7||i===9) && v!==null) c.numFmt = '#,##0.00';
      if ((i===8||i===10||i===11||i===12) && v!==null) c.numFmt = '0.00"%"';
    });

    // ══════════════════════════════════════
    // SHEET 3 — Unmatched Parts (only if any exist)
    // Parts present in RFM but not found in PO
    // ══════════════════════════════════════
    var unmatchedRows = p.rfmRows.filter(function(r){ return r.status === 'no-po'; });

    if (unmatchedRows.length > 0) {
      var ws3 = wb.addWorksheet('Unmatched Parts');

      // Column widths
      [5, 13, 18, 28, 40, 10, 14, 12]
        .forEach(function(w, i){ ws3.getColumn(i+1).width = w; });

      // Info note row 1
      var noteRow = ws3.getRow(1); noteRow.height = 20;
      var noteCell = noteRow.getCell(1);
      noteCell.value = 'These ' + unmatchedRows.length + ' part(s) were present in the RFM slip but had NO matching entry in the PO file. They may be unordered or incorrectly listed.';
      noteCell.font  = {italic: true, size: 9, name: 'Calibri', color: {argb: 'FF92400E'}};
      noteCell.fill  = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFF7ED'}};
      noteCell.alignment = {vertical: 'middle'};
      ws3.mergeCells(1, 1, 1, 8);
      noteCell.border = {bottom: {style: 'thin', color: {argb: 'FFFED7AA'}}};

      // Header row 2
      var u_hdrs = ['SR', 'Item Req No', 'Make', 'Part No', 'Description', 'Scope', 'Section', 'Return Qty'];
      var uHdrRow = ws3.getRow(2); uHdrRow.height = 22;
      u_hdrs.forEach(function(v, i) {
        var hc = uHdrRow.getCell(i+1);
        hc.value = v;
        hc.font      = {bold: true, size: 9, color: {argb: 'FFFFFFFF'}, name: 'Calibri'};
        hc.fill      = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFDC2626'}};
        hc.alignment = {vertical: 'middle', horizontal: 'center', wrapText: true};
        hc.border    = {bottom: {style: 'medium', color: {argb: 'FFB91C1C'}}, right: {style: 'thin', color: {argb: 'FFB91C1C'}}};
      });

      // Data rows — alternating light pink / white
      unmatchedRows.forEach(function(r, ri) {
        var udr = ws3.addRow([]); udr.height = 16;
        var ubg = ri % 2 === 0 ? 'FFFFFFFF' : 'FFFFF5F5';
        var uvals = [ri+1, r.itemReqNo||'', r.make||'', r.typeNo, r.desc||'', r.scope||'', r.section||'', r.returnQty];
        uvals.forEach(function(v, ci) {
          var uc = udr.getCell(ci+1);
          uc.value = (v === '' || v === null || v === undefined) ? null : v;
          uc.fill  = {type: 'pattern', pattern: 'solid', fgColor: {argb: ubg}};
          var isMono = ci === 3;
          var isCtr  = ci === 0 || ci === 5 || ci === 6 || ci === 7;
          var isBold = ci === 3 || ci === 7;
          uc.font  = {size: 9, name: isMono ? 'Courier New' : 'Calibri', bold: isBold, color: {argb: ci === 3 ? 'FFDC2626' : 'FF1F2937'}};
          uc.alignment = {vertical: 'middle', horizontal: isCtr ? 'center' : 'left', wrapText: ci === 4};
          uc.border = {bottom: {style: 'thin', color: {argb: 'FFFEE2E2'}}, right: {style: 'thin', color: {argb: 'FFFEE2E2'}}};
        });
      });

      // Totals row
      var uTotalRow = ws3.addRow([]); uTotalRow.height = 18;
      var uTotalQty = unmatchedRows.reduce(function(s, r){ return s + r.returnQty; }, 0);
      ['', '', '', '', '', '', 'Total Return Qty', uTotalQty].forEach(function(v, ci) {
        var tc = uTotalRow.getCell(ci+1);
        tc.value = v || null;
        tc.font  = {bold: true, size: 9, name: 'Calibri', color: {argb: ci === 7 ? 'FFDC2626' : 'FF374151'}};
        tc.fill  = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFEF2F2'}};
        tc.alignment = {vertical: 'middle', horizontal: ci === 6 ? 'right' : ci === 7 ? 'center' : 'left'};
        tc.border = {top: {style: 'medium', color: {argb: 'FFFCA5A5'}}, bottom: {style: 'thin', color: {argb: 'FFFCA5A5'}}};
      });
    }

    var buffer = await wb.xlsx.writeBuffer();
    var blob   = new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var fname  = (document.getElementById('exportFileName').value||'RFM_Report').trim().replace(/[^a-zA-Z0-9_\-. ]/g,'_');
    saveAs(blob, fname + '.xlsx');

    btn.textContent = '↓ Download Excel'; btn.disabled = false;
  } catch(err) {
    console.error(err);
    alert('Export failed: ' + err.message);
    btn.textContent = '↓ Download Excel'; btn.disabled = false;
  }
}
