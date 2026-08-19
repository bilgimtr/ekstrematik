/* ============ Çerez / reklam onayı ============ */
/* Kullanıcı "Kabul Ediyorum" demeden AdSense scripti hiç yüklenmez. */

(function(){
  var ADSENSE_CLIENT = 'ca-pub-6046256779379664';
  var STORAGE_KEY = 'ekstrematik_consent';

  function loadAdsense(){
    if(document.getElementById('adsbygoogle-script')) return;
    var s = document.createElement('script');
    s.id = 'adsbygoogle-script';
    s.async = true;
    s.src = 'https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=' + ADSENSE_CLIENT;
    s.crossOrigin = 'anonymous';
    document.head.appendChild(s);
  }

  function buildBanner(){
    var el = document.createElement('div');
    el.className = 'consent-banner';
    el.id = 'consent-banner';
    el.setAttribute('role', 'dialog');
    el.setAttribute('aria-label', 'Çerez onayı');
    el.innerHTML =
      '<div class="consent-inner">' +
        '<div class="consent-text">' +
          '<strong>Çerezler hakkında.</strong> Sitenin temel işlevleri (dosya okuma, mutabakat hesaplama) için çerez kullanmıyoruz. ' +
          'Reklam gösterimi için Google AdSense çerezleri kullanılabilir; bunlara onay vermeniz reklamların size göre kişiselleştirilmesini sağlar. ' +
          'Detaylar için <a href="/gizlilik.html">Gizlilik Politikası</a> sayfamıza bakabilirsiniz.' +
        '</div>' +
        '<div class="consent-actions">' +
          '<button type="button" class="btn btn-reject" id="consent-reject">Reddediyorum</button>' +
          '<button type="button" class="btn btn-accept" id="consent-accept">Kabul Ediyorum</button>' +
        '</div>' +
      '</div>';
    document.body.appendChild(el);
    requestAnimationFrame(function(){ el.classList.add('show'); });

    document.getElementById('consent-accept').addEventListener('click', function(){
      setConsent('accepted');
      hideBanner();
    });
    document.getElementById('consent-reject').addEventListener('click', function(){
      setConsent('rejected');
      hideBanner();
    });
    return el;
  }

  function hideBanner(){
    var el = document.getElementById('consent-banner');
    if(!el) return;
    el.classList.remove('show');
    setTimeout(function(){ el.remove(); }, 350);
  }

  function setConsent(value){
    try{ localStorage.setItem(STORAGE_KEY, value); }catch(e){}
    if(value === 'accepted') loadAdsense();
  }

  function getConsent(){
    try{ return localStorage.getItem(STORAGE_KEY); }catch(e){ return null; }
  }

  function init(){
    var consent = getConsent();
    if(consent === 'accepted'){
      loadAdsense();
    } else if(consent !== 'rejected'){
      buildBanner();
    }

    var prefLinks = document.querySelectorAll('.consent-pref-link');
    prefLinks.forEach(function(link){
      link.addEventListener('click', function(e){
        e.preventDefault();
        if(!document.getElementById('consent-banner')) buildBanner();
      });
    });
  }

  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
