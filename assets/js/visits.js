/* ============ Kullanıcı sayacı (anonim, çerez tabanlı) ============ */

(function(){
  var el = document.getElementById('visitor-stats');
  if(!el) return;

  fetch('/api/visit', { credentials: 'same-origin' })
    .then(function(res){ return res.ok ? res.json() : null; })
    .then(function(data){
      if(!data) return;
      el.textContent = 'Bugün ' + data.today.toLocaleString('tr-TR') +
        ' · Toplam ' + data.total.toLocaleString('tr-TR') + ' kullanıcı bu aracı kullandı';
    })
    .catch(function(){ /* sessizce geç, sayaç kritik değil */ });
})();
