// Handles single-letter cols (A=0) and two-letter cols (AG=32, etc.)
function colIdx(l) {
  l = (l || 'A').toUpperCase().trim();
  if (l.length === 1) return l.charCodeAt(0) - 65;
  return (l.charCodeAt(0) - 64) * 26 + (l.charCodeAt(1) - 65);
}
function fmtSz(b) { if(b<1024)return b+' B'; if(b<1048576)return (b/1024).toFixed(1)+' KB'; return (b/1048576).toFixed(1)+' MB'; }
function todayStr() { return new Date().toISOString().slice(0,10); }

// ── PART NUMBER NORMALIZATION — strip non-alphanumeric, lowercase ──
function normPN(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/^\^+/, '')   // SheetJS text-forced prefix
    .replace(/[^a-z0-9]/gi, '')
    .toLowerCase()
    .trim();
}
// Keep original normalisation (just trim/lower) as fallback key
function normPNExact(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/^\^+/, '')
    .trim()
    .toLowerCase();
}
