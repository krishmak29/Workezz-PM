// ── AUTO-DETECT HEADER ROW ──
// Requires the row to match at least `minScore` of the keywords, AND
// at least one cell must be a SHORT label (≤ 30 chars) to avoid matching
// free-text metadata rows that happen to contain a keyword word.
function detectHeaderRow(data, keywords, minScore) {
  minScore = minScore || Math.ceil(keywords.length * 0.55); // need >55% match
  var best = {row: -1, score: 0};
  for (var i = 0; i < Math.min(25, data.length); i++) {
    var row = data[i].map(function(v){ return String(v||'').toLowerCase().trim(); });
    // Row must have at least 3 non-empty cells that are short labels (header-like)
    var shortCells = row.filter(function(c){ return c.length > 0 && c.length <= 35; }).length;
    if (shortCells < 3) continue;
    var score = keywords.filter(function(kw){
      return row.some(function(cell){ return cell === kw || cell.includes(kw); });
    }).length;
    if (score >= minScore && score > best.score) {
      best = {row: i, score: score};
    }
  }
  return best.row;
}

// Find column index in a header row — prefers EXACT match over partial
function findCol(headerRow, keyword) {
  keyword = keyword.toLowerCase().trim();
  // 1. Exact match first
  for (var i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]||'').toLowerCase().trim() === keyword) return i;
  }
  // 2. Starts-with match
  for (var i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]||'').toLowerCase().trim().startsWith(keyword)) return i;
  }
  // 3. Contains match
  for (var i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]||'').toLowerCase().includes(keyword)) return i;
  }
  return -1;
}

 // ── FL SUB-HEADER SKIPPER ──
  // After the detected header row, skip any rows where every non-empty cell
  // is a short (≤6 chars) non-numeric label — e.g. "FL", "FL FL", separators.
  // These are always treated as part of the header block, never as data.
  function skipSubHeaders(data, hdrIdx) {
    var dataStart = hdrIdx + 1;
    while (dataStart < data.length) {
      var row = data[dataStart];
      var nonEmpty = row.filter(function(v){ return String(v||'').trim() !== ''; });
      if (nonEmpty.length === 0) { dataStart++; continue; }
      var isSubHdr = nonEmpty.every(function(v) {
        var s = String(v).trim();
        return s.length <= 6 && isNaN(s);
      });
      if (isSubHdr) { dataStart++; } else { break; }
    }
    return dataStart;
  }