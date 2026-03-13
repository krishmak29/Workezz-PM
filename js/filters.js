// ── FILTER ──
function setFilter(f) {
  activeFilter = f;
  filterQuery  = document.getElementById('filterSearch').value;
  document.querySelectorAll('.filter-btn').forEach(function(b){
    b.classList.toggle('active', b.dataset.filter === f);
  });
  renderTable();
}

function applyFilter() {
  filterQuery = document.getElementById('filterSearch').value;
  renderTable();
}

function resetFilters() {
  activeFilter = 'all';
  filterQuery  = '';
  document.getElementById('filterSearch').value = '';
  document.querySelectorAll('.filter-btn').forEach(function(b){
    b.classList.toggle('active', b.dataset.filter === 'all');
  });
  renderTable();
}
