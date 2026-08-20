/* ============ İş mantığı (format algılama, eşleştirme, gruplama) ============ */

const FORMAT_TANIMLARI = {
  hesap_karti: {
    gerekli: ['TARİH','EVRAK TİPİ','ANA DÖVİZ BORÇ','ANA DÖVİZ ALACAK'],
    tarih:'TARİH', tip:'EVRAK TİPİ', borc:'ANA DÖVİZ BORÇ', alacak:'ANA DÖVİZ ALACAK',
    evrak:'GIB FATURA NO', ek:'SORUMLU İSMİ'
  },
  cari_ekstre: {
    gerekli: ['İşlem Tarihi','İşlem Türü','Borç Tutarı','Alacak Tutarı'],
    tarih:'İşlem Tarihi', tip:'İşlem Türü', borc:'Borç Tutarı', alacak:'Alacak Tutarı',
    evrak:'Evrak No', ek:'Açıklama'
  },
  logo_ekstre: {
    gerekli: ['TARIH','BELGE NO','BELGETURU','BORC','ALACAK'],
    tarih:'TARIH', tip:'BELGETURU', borc:'BORC', alacak:'ALACAK',
    evrak:'BELGE NO', ek:'ACIKLAMA'
  },
  doviz_ekstre: {
    gerekli: ['Tarih','Fiş No','İşlem Türü','Borç','Alacak','PB','Kur','İşlem TL'],
    tarih:'Tarih', tip:'İşlem Türü', borc:'Borç', alacak:'Alacak',
    evrak:'Fiş No', ek:null,
    // Borç/Alacak dövizli sütunlar sadece YÖN belirlemek için kullanılır;
    // gerçek karşılaştırma tutarı her zaman TL karşılığından alınır —
    // aksi halde döviz tutarı TL ekstresiyle asla eşleşmez.
    tutarKaynak:'İşlem TL'
  }
};

function round2(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }

function parseSayi(val){
  if(val===null || val===undefined || val==='') return 0;
  if(typeof val === 'number') return val;
  let s = String(val).trim();
  if(s==='') return 0;
  // Para birimi sembolü/etiketi varsa temizle (ör. "64.595,20₺", "$1.234,56", "1.234,56 TL")
  s = s.replace(/[^\d.,-]/g,'');
  if(s==='') return 0;
  // Türkçe biçim: binlik nokta, ondalık virgül (ör. "1.234,56") -> 1234.56
  if(/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)){
    s = s.replace(/\./g,'').replace(',', '.');
  } else if(s.includes(',') && !s.includes('.')){
    // sadece virgüllü ondalık (ör. "1234,56")
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function normalizeEvrak(val){
  if(val===null || val===undefined) return null;
  const s = String(val).trim();
  return s==='' ? null : s;
}

function parseTarih(val){
  if (val instanceof Date) return val;
  if (typeof val === 'number'){
    const d = XLSX.SSF.parse_date_code(val);
    if(!d) return null;
    return new Date(d.y, d.m-1, d.d, d.H||0, d.M||0, d.S||0);
  }
  if (typeof val === 'string'){
    const m = val.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?/);
    if (m){
      const [,d,mo,y,h='0',mi='0',s='0'] = m;
      return new Date(+y, +mo-1, +d, +h, +mi, +s);
    }
  }
  return null;
}

function fmtTarih(d){
  if(!d) return '';
  const p = n => String(n).padStart(2,'0');
  return `${p(d.getDate())}.${p(d.getMonth()+1)}.${d.getFullYear()}`;
}
function fmtSayi(n){
  if(n===null||n===undefined||n===0) return '';
  return n.toLocaleString('tr-TR',{minimumFractionDigits:2,maximumFractionDigits:2});
}

/* ============ Esnek sütun eşleştirme (bilinen formatlardan hiçbiri tutmazsa) ============ */

function normalizeHeader(s){
  return String(s)
    .toLocaleLowerCase('tr')
    .replace(/ı/g,'i').replace(/i̇/g,'i')
    .replace(/ş/g,'s').replace(/ç/g,'c').replace(/ğ/g,'g').replace(/ö/g,'o').replace(/ü/g,'u')
    .replace(/[_.]/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}

const ALAN_ANAHTAR_KELIMELERI = {
  tarih: ['tarih'],
  tip: ['evrak tipi','islem turu','belge turu','belgeturu','hareket tipi','tip','tur'],
  borc: ['borc'],
  alacak: ['alacak'],
  evrak: ['evrak no','belge no','fatura no','evrakno','belgeno','evrak','belge'],
  ek: ['aciklama','sorumlu','ek bilgi','not']
};

function esleKolonBul(normCols, anahtarlar){
  for(const anahtar of anahtarlar){
    const bulunan = normCols.find(c => c.norm === anahtar);
    if(bulunan) return bulunan.orig;
  }
  for(const anahtar of anahtarlar){
    const bulunan = normCols.find(c => c.norm.includes(anahtar));
    if(bulunan) return bulunan.orig;
  }
  return null;
}

function enIyiTahminler(cols){
  const normCols = cols.map(c => ({orig:c, norm:normalizeHeader(c)}));
  return {
    tarih: esleKolonBul(normCols, ALAN_ANAHTAR_KELIMELERI.tarih),
    tip: esleKolonBul(normCols, ALAN_ANAHTAR_KELIMELERI.tip),
    borc: esleKolonBul(normCols, ALAN_ANAHTAR_KELIMELERI.borc),
    alacak: esleKolonBul(normCols, ALAN_ANAHTAR_KELIMELERI.alacak),
    evrak: esleKolonBul(normCols, ALAN_ANAHTAR_KELIMELERI.evrak),
    ek: esleKolonBul(normCols, ALAN_ANAHTAR_KELIMELERI.ek),
  };
}

function detectFormatEsnek(rows){
  if(!rows.length) return null;
  const def = enIyiTahminler(Object.keys(rows[0]));
  if(!def.tarih || !def.borc || !def.alacak) return null;
  return {name:'esnek', def};
}

/* ============ Kullanıcının elle eşleştirdiği formatı hatırlama ============ */

function kolonImzasi(cols){
  return cols.slice().sort().join('|');
}

function hatirlananFormatGetir(cols){
  try{
    const raw = localStorage.getItem('ekstrematik_format_' + kolonImzasi(cols));
    return raw ? JSON.parse(raw) : null;
  }catch(e){ return null; }
}

function hatirlananFormatKaydet(cols, def){
  try{
    localStorage.setItem('ekstrematik_format_' + kolonImzasi(cols), JSON.stringify(def));
  }catch(e){}
}

function detectFormat(rows){
  if(!rows.length) return null;
  const cols = Object.keys(rows[0]);
  const hatirlanan = hatirlananFormatGetir(cols);
  if(hatirlanan) return {name:'hatirlanan', def:hatirlanan};
  const colSet = new Set(cols);
  for(const [name, def] of Object.entries(FORMAT_TANIMLARI)){
    if(def.gerekli.every(c=>colSet.has(c))) return {name, def};
  }
  return detectFormatEsnek(rows);
}

function extractMovements(rows, def){
  const kayitlar = [];
  let sid = 0;
  for(const r of rows){
    const tip = def.tip ? r[def.tip] : '';
    if(def.tip && tip==null) continue;
    const tipStr = String(tip||'').toUpperCase();
    if(tipStr.includes('DEVİR') || tipStr.includes('DEVIR')) continue;
    const tarih = parseTarih(r[def.tarih]);
    if(!tarih) continue;
    const borcYon = parseSayi(r[def.borc]);
    const alacakYon = parseSayi(r[def.alacak]);
    let borc, alacak;
    if(def.tutarKaynak){
      // Borç/Alacak sadece yön belirtir; gerçek tutar ayrı bir sütundan (ör. TL karşılığı) alınır
      const tutar = Math.abs(parseSayi(r[def.tutarKaynak]));
      borc = borcYon!==0 ? tutar : 0;
      alacak = alacakYon!==0 ? tutar : 0;
    } else {
      borc = borcYon;
      alacak = alacakYon;
    }
    borc = round2(borc);
    alacak = round2(alacak);
    const evrak = def.evrak ? normalizeEvrak(r[def.evrak]) : null;
    kayitlar.push({sid, tarih, tip, borc, alacak, evrak, ek: def.ek ? r[def.ek] : null});
    sid++;
  }
  const hareketler = [];
  for(const k of kayitlar){
    if(k.borc!==0) hareketler.push({satirId:k.sid, tarih:k.tarih, tip:k.tip, yon:'BORC', tutar:k.borc, evrak:k.evrak});
    if(k.alacak!==0) hareketler.push({satirId:k.sid, tarih:k.tarih, tip:k.tip, yon:'ALACAK', tutar:k.alacak, evrak:k.evrak});
  }
  return {kayitlar, hareketler};
}

function findSubset(target, pool, tol, maxSize){
  const sorted = pool.slice().sort((a,b)=>a.tutar-b.tutar);
  const n = sorted.length;
  function dfs(start, chosen, total){
    if(chosen.length>=2 && Math.abs(total-target)<=tol) return chosen.slice();
    if(chosen.length>=maxSize || start>=n) return null;
    for(let i=start;i<n;i++){
      const c=sorted[i]; const nt=total+c.tutar;
      if(nt>target+tol) continue;
      chosen.push(c);
      const res=dfs(i+1,chosen,nt);
      if(res) return res;
      chosen.pop();
    }
    return null;
  }
  return dfs(0,[],0);
}

function grupEslestir(h1, h2, tol=0.05, maxSize=5, dateWindowDays=60){
  const ters = {BORC:'ALACAK', ALACAK:'BORC'};
  let grupNo = 0;
  function tekYondenAra(hedefListe, havuzListe){
    const hedefler = hedefListe.filter(r=>!r.matched).sort((a,b)=>a.tarih-b.tarih);
    for(const hedef of hedefler){
      if(hedef.matched) continue;
      const havuz = havuzListe.filter(r=>!r.matched && r.yon===ters[hedef.yon] &&
        Math.abs((r.tarih - hedef.tarih)/86400000) <= dateWindowDays);
      if(havuz.length<2) continue;
      const combo = findSubset(hedef.tutar, havuz, tol, maxSize);
      if(combo){
        grupNo++;
        hedef.matched=true; hedef.status='Gruplu Eşleşme'; hedef.grup=grupNo;
        for(const c of combo){ c.matched=true; c.status='Gruplu Eşleşme'; c.grup=grupNo; }
      }
    }
  }
  tekYondenAra(h1, h2);
  tekYondenAra(h2, h1);
}

function esletir(h1in, h2in, evrakTol=1.0, tutarTol=0.05){
  const h1 = h1in.map(r=>({...r, matched:false, matchId2:null, status:null, grup:null}));
  const h2 = h2in.map(r=>({...r, matched:false, matchId1:null, grup:null}));
  const ters = {BORC:'ALACAK', ALACAK:'BORC'};

  for(const r1 of h1){
    if(r1.matched || r1.evrak==null || r1.evrak==='') continue;
    let best=null, bestDiff=Infinity;
    for(const r2 of h2){
      if(r2.matched || r2.evrak!==r1.evrak || r2.yon!==ters[r1.yon]) continue;
      const diff = Math.abs(r2.tutar - r1.tutar);
      if(diff<bestDiff){ bestDiff=diff; best=r2; }
    }
    if(best){
      r1.matched=true; r1.matchId2=best.satirId;
      r1.status = bestDiff<=evrakTol ? 'Eşleşti' : 'Tutar Farklı';
      best.matched=true; best.matchId1=r1.satirId;
    }
  }
  const kalan = h1.filter(r=>!r.matched).sort((a,b)=>a.tarih-b.tarih);
  for(const r1 of kalan){
    let best=null, bestGap=Infinity;
    for(const r2 of h2){
      if(r2.matched || r2.yon!==ters[r1.yon]) continue;
      if(Math.abs(r2.tutar-r1.tutar)>tutarTol) continue;
      const gap = Math.abs(r2.tarih-r1.tarih);
      if(gap<bestGap){ bestGap=gap; best=r2; }
    }
    if(best){
      r1.matched=true; r1.matchId2=best.satirId; r1.status='Eşleşti';
      best.matched=true; best.matchId1=r1.satirId;
    }
  }
  return {h1, h2};
}

function readWorkbookFile(file){
  return new Promise((resolve,reject)=>{
    const reader = new FileReader();
    reader.onload = e=>{
      try{
        const wb = XLSX.read(e.target.result, {type:'array', cellDates:true, raw:true});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {defval:null, raw:true});
        resolve(rows);
      }catch(err){ reject(err); }
    };
    reader.onerror = ()=>reject(new Error('Dosya okunamadı'));
    reader.readAsArrayBuffer(file);
  });
}
