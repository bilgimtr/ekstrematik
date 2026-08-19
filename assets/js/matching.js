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
  }
};

function round2(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }

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

function detectFormat(rows){
  if(!rows.length) return null;
  const cols = new Set(Object.keys(rows[0]));
  for(const [name, def] of Object.entries(FORMAT_TANIMLARI)){
    if(def.gerekli.every(c=>cols.has(c))) return {name, def};
  }
  return null;
}

function extractMovements(rows, def){
  const kayitlar = [];
  let sid = 0;
  for(const r of rows){
    const tip = r[def.tip];
    if(tip==null) continue;
    const tipStr = String(tip).toUpperCase();
    if(tipStr.includes('DEVİR') || tipStr.includes('DEVIR')) continue;
    const tarih = parseTarih(r[def.tarih]);
    if(!tarih) continue;
    const borc = round2(Number(r[def.borc])||0);
    const alacak = round2(Number(r[def.alacak])||0);
    const evrak = def.evrak ? r[def.evrak] : null;
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
