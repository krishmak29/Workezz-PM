function toggleMode() {
  var isLight = document.body.classList.toggle('light');
  document.getElementById('modeIcon').textContent = isLight ? '🌙' : '☀️';
  document.getElementById('modeLbl').textContent  = isLight ? 'Dark' : 'Light';
  try { localStorage.setItem('wz-theme', isLight ? 'light' : 'dark'); } catch(e) {}
}
(function() {
  try {
    var saved = localStorage.getItem('wz-theme');
    if (saved === 'light') {
      document.body.classList.add('light');
      document.getElementById('modeIcon').textContent = '🌙';
      document.getElementById('modeLbl').textContent  = 'Dark';
    }
  } catch(e) {}
})();