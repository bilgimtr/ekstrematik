/* ============ Adım (stepper) yönetimi ============ */

const state = { a: null, b: null };
let currentStep = 1;

function goToPanel(id){
  document.querySelectorAll('.step-panel').forEach(p=>p.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.getElementById('wrap').scrollIntoView({behavior:'smooth', block:'start'});
}

function goToStep(n){
  goToPanel('panel-'+n);
  document.querySelectorAll('.step-item').forEach(si=>{
    const s = Number(si.dataset.step);
    si.classList.toggle('active', s===n);
    si.classList.toggle('done', s<n);
  });
  document.querySelectorAll('.step-connector').forEach(c=>{
    c.classList.toggle('done', Number(c.dataset.conn) < n);
  });
  currentStep = n;
  if(n===3) buildConfirmSummary();
}

function buildConfirmSummary(){
  const firmaA = isimA.value.trim();
  const firmaB = isimB.value.trim();
  const el = document.getElementById('confirm-summary');
  el.innerHTML = `
    <div class="confirm-card">
      <div>
        <div class="cc-name">${firmaA}</div>
        <div class="cc-file">${state.a.file.name}</div>
      </div>
      <div class="cc-count">${state.a.kayitlar.length} hareket</div>
    </div>
    <div class="confirm-card">
      <div>
        <div class="cc-name">${firmaB}</div>
        <div class="cc-file">${state.b.file.name}</div>
      </div>
      <div class="cc-count">${state.b.kayitlar.length} hareket</div>
    </div>
  `;
}

/* ============ Dosya yükleme alanları ============ */

function setupDrop(slot, onChange){
  const dropEl = document.getElementById('drop-'+slot);
  const inputEl = document.getElementById('file-'+slot);
  const fnameEl = document.getElementById('fname-'+slot);
  const rowsEl = document.getElementById('rows-'+slot);
  const titleEl = document.getElementById('title-'+slot);
  const clearEl = document.getElementById('clear-'+slot);

  function openPicker(){ inputEl.click(); }
  dropEl.addEventListener('click', openPicker);
  dropEl.addEventListener('dragover', e=>{e.preventDefault(); dropEl.classList.add('drag');});
  dropEl.addEventListener('dragleave', ()=>dropEl.classList.remove('drag'));
  dropEl.addEventListener('drop', e=>{
    e.preventDefault(); dropEl.classList.remove('drag');
    if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  inputEl.addEventListener('change', e=>{
    if(e.target.files.length) handleFile(e.target.files[0]);
  });
  clearEl.addEventListener('click', e=>{
    e.stopPropagation();
    state[slot]=null;
    inputEl.value='';
    dropEl.classList.remove('filled','error');
    titleEl.textContent = 'Ekstre dosyasını sürükle veya seç';
    fnameEl.textContent=''; rowsEl.textContent='';
    onChange();
  });

  async function handleFile(file){
    dropEl.classList.remove('filled','error');
    titleEl.textContent = 'Okunuyor...';
    fnameEl.textContent=''; rowsEl.textContent='';
    try{
      const rows = await readWorkbookFile(file);
      const fmt = detectFormat(rows);
      if(!fmt){
        dropEl.classList.add('error');
        titleEl.textContent = 'Format tanınamadı';
        fnameEl.textContent = file.name;
        rowsEl.textContent = 'Bu dosya biçimi desteklenmiyor';
        state[slot]=null;
        onChange();
        return;
      }
      const {kayitlar, hareketler} = extractMovements(rows, fmt.def);
      state[slot] = {file, rows, fmt, kayitlar, hareketler};
      dropEl.classList.add('filled');
      titleEl.textContent = 'Dosya yüklendi';
      fnameEl.textContent = file.name;
      rowsEl.textContent = `${kayitlar.length} hareket bulundu`;
      onChange();
      const isimEl = document.getElementById('isim-'+slot);
      if(isimEl.value.trim()===''){
        isimEl.focus();
      }
    }catch(err){
      dropEl.classList.add('error');
      titleEl.textContent = 'Okuma hatası';
      fnameEl.textContent = file.name;
      rowsEl.textContent = String(err.message||err);
      state[slot]=null;
      onChange();
    }
  }
}

const isimA = document.getElementById('isim-a');
const isimB = document.getElementById('isim-b');
const hintA = document.getElementById('hint-a');
const hintB = document.getElementById('hint-b');
const next1 = document.getElementById('next-1');
const next2 = document.getElementById('next-2');

function isimKontrol(inputEl, hintEl){
  const bos = inputEl.value.trim() === '';
  const dokunuldu = inputEl.dataset.touched === '1';
  inputEl.classList.toggle('empty', bos && dokunuldu && document.activeElement!==inputEl);
  hintEl.classList.toggle('show', bos && dokunuldu && document.activeElement!==inputEl);
  return !bos;
}

function updateStep1(){
  const ok = isimKontrol(isimA, hintA) && !!state.a;
  next1.disabled = !ok;
}
function updateStep2(){
  const ok = isimKontrol(isimB, hintB) && !!state.b;
  next2.disabled = !ok;
}

isimA.addEventListener('input', ()=>{ isimA.dataset.touched='1'; updateStep1(); });
isimA.addEventListener('blur', ()=>{ isimA.dataset.touched='1'; updateStep1(); });
isimB.addEventListener('input', ()=>{ isimB.dataset.touched='1'; updateStep2(); });
isimB.addEventListener('blur', ()=>{ isimB.dataset.touched='1'; updateStep2(); });

setupDrop('a', updateStep1);
setupDrop('b', updateStep2);

document.getElementById('next-1').addEventListener('click', ()=>{ if(!next1.disabled) goToStep(2); });
document.getElementById('back-2').addEventListener('click', ()=> goToStep(1));
document.getElementById('next-2').addEventListener('click', ()=>{ if(!next2.disabled) goToStep(3); });
document.getElementById('back-3').addEventListener('click', ()=> goToStep(2));

document.getElementById('restart-btn').addEventListener('click', ()=>{
  ['a','b'].forEach(slot=>{
    state[slot]=null;
    document.getElementById('file-'+slot).value='';
    document.getElementById('drop-'+slot).classList.remove('filled','error');
    document.getElementById('title-'+slot).textContent = 'Ekstre dosyasını sürükle veya seç';
    document.getElementById('fname-'+slot).textContent='';
    document.getElementById('rows-'+slot).textContent='';
  });
  isimA.value=''; isimB.value='';
  isimA.dataset.touched=''; isimB.dataset.touched='';
  isimA.classList.remove('empty'); isimB.classList.remove('empty');
  hintA.classList.remove('show'); hintB.classList.remove('show');
  next1.disabled = true; next2.disabled = true;
  currentRows = []; currentFilter = 'all';
  goToStep(1);
});

/* ============ Mutabakatı çalıştır ============ */

let currentRows = [];
let currentFilter = 'all';

const runBtn = document.getElementById('run-btn');
runBtn.addEventListener('click', ()=>{
  runBtn.classList.remove('stamped');
  void runBtn.offsetWidth;
  runBtn.classList.add('stamped');
  runMutabakat();
  setTimeout(()=>showAdGate(), 260);
});

function runMutabakat(){
  const firmaA = isimA.value.trim() || 'Kendi Firmanız';
  const firmaB = isimB.value.trim() || 'Karşı Firma';

  const {h1, h2} = esletir(state.a.hareketler, state.b.hareketler);
  grupEslestir(h1, h2);
  const k1 = new Map(state.a.kayitlar.map(k=>[k.sid,k]));
  const k2 = new Map(state.b.kayitlar.map(k=>[k.sid,k]));

  const rows = [];

  for(const r1 of h1){
    if(r1.matched && r1.matchId2!=null){
      const a = k1.get(r1.satirId);
      const b = k2.get(r1.matchId2);
      rows.push({
        aT:a.tarih, aTip:a.tip, aBorc:a.borc, aAlacak:a.alacak, aEvrak:a.evrak, aEk:a.ek,
        bT:b.tarih, bTip:b.tip, bBorc:b.borc, bAlacak:b.alacak, bEvrak:b.evrak, bEk:b.ek,
        durum:r1.status, grup:null
      });
    } else if(r1.status==='Gruplu Eşleşme'){
      const a = k1.get(r1.satirId);
      rows.push({
        aT:a.tarih, aTip:a.tip, aBorc:a.borc, aAlacak:a.alacak, aEvrak:a.evrak, aEk:a.ek,
        bT:null,bTip:null,bBorc:null,bAlacak:null,bEvrak:null,bEk:null,
        durum:'Gruplu Eşleşme', grup:r1.grup
      });
    } else if(!r1.matched){
      const a = k1.get(r1.satirId);
      rows.push({
        aT:a.tarih, aTip:a.tip, aBorc:a.borc, aAlacak:a.alacak, aEvrak:a.evrak, aEk:a.ek,
        bT:null,bTip:null,bBorc:null,bAlacak:null,bEvrak:null,bEk:null,
        durum:'only-a', grup:null
      });
    }
  }
  for(const r2 of h2){
    if(r2.matched && r2.matchId1!=null){
      continue;
    } else if(r2.status==='Gruplu Eşleşme'){
      const b = k2.get(r2.satirId);
      rows.push({
        aT:null,aTip:null,aBorc:null,aAlacak:null,aEvrak:null,aEk:null,
        bT:b.tarih, bTip:b.tip, bBorc:b.borc, bAlacak:b.alacak, bEvrak:b.evrak, bEk:b.ek,
        durum:'Gruplu Eşleşme', grup:r2.grup
      });
    } else if(!r2.matched){
      const b = k2.get(r2.satirId);
      rows.push({
        aT:null,aTip:null,aBorc:null,aAlacak:null,aEvrak:null,aEk:null,
        bT:b.tarih, bTip:b.tip, bBorc:b.borc, bAlacak:b.alacak, bEvrak:b.evrak, bEk:b.ek,
        durum:'only-b', grup:null
      });
    }
  }
  rows.sort((x,y)=>{
    const dx = x.aT||x.bT, dy=y.aT||y.bT;
    return dx-dy;
  });

  currentRows = rows;
  renderSummary(rows, firmaA, firmaB);
  renderTable(rows, firmaA, firmaB);
}

/* ============ Reklam bekleme kapısı ============ */
/* Sonuç yukarıda arka planda zaten hesaplandı (runMutabakat). Kullanıcı bir reklam
   alanı görür; "Sonucu Gör" butonu sadece süre dolunca aktifleşir — reklamla
   etkileşim (tıklama/hover) şartı YOK, sadece zamana bağlı bir bekleme. */
const AD_WAIT_SECONDS = 18;
let adCountdownTimer = null;

function showAdGate(){
  goToPanel('panel-3-5');
  const revealBtn = document.getElementById('reveal-btn');
  const countdownEl = document.getElementById('ad-countdown');
  revealBtn.disabled = true;
  let remaining = AD_WAIT_SECONDS;
  countdownEl.textContent = remaining;
  if(adCountdownTimer) clearInterval(adCountdownTimer);
  adCountdownTimer = setInterval(()=>{
    remaining--;
    if(remaining<=0){
      clearInterval(adCountdownTimer);
      adCountdownTimer = null;
      revealBtn.disabled = false;
      countdownEl.textContent = 'Hazır';
    } else {
      countdownEl.textContent = remaining;
    }
  }, 1000);
}

document.getElementById('reveal-btn').addEventListener('click', ()=>{
  goToStep(4);
});

function renderSummary(rows, firmaA, firmaB){
  const counts = {'Eşleşti':0,'Tutar Farklı':0,'Gruplu Eşleşme':0,'only-a':0,'only-b':0};
  rows.forEach(r=>counts[r.durum]++);
  const cards = [
    {key:'Eşleşti', label:'Eşleşen hareket', num:counts['Eşleşti'], color:'var(--green)'},
    {key:'Gruplu Eşleşme', label:'Gruplu eşleşme', num:counts['Gruplu Eşleşme'], color:'var(--purple)'},
    {key:'only-a', label:`Yalnız ${firmaA}'de`, num:counts['only-a'], color:'var(--blue)'},
    {key:'only-b', label:`Yalnız ${firmaB}'de`, num:counts['only-b'], color:'var(--red)'},
    {key:'Tutar Farklı', label:'Tutar farklı', num:counts['Tutar Farklı'], color:'var(--amber)'},
  ];
  const el = document.getElementById('summary');
  el.innerHTML = '';
  cards.forEach(c=>{
    const div = document.createElement('div');
    div.className = 'metric';
    div.style.setProperty('--metric-color', c.color);
    div.dataset.key = c.key;
    div.innerHTML = `<div class="num">${c.num}</div><div class="lbl">${c.label}</div>`;
    div.addEventListener('click', ()=>{
      currentFilter = (currentFilter===c.key) ? 'all' : c.key;
      document.querySelectorAll('.metric').forEach(m=>m.classList.toggle('active', m.dataset.key===currentFilter));
      applyFilter();
    });
    el.appendChild(div);
  });
}

function renderTable(rows, firmaA, firmaB){
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  rows.forEach(r=>{
    const tr = document.createElement('tr');
    tr.dataset.status = r.durum;
    const durumText = r.durum==='only-a' ? `Yalnız ${firmaA}'de var`
      : r.durum==='only-b' ? `Yalnız ${firmaB}'de var`
      : r.durum==='Gruplu Eşleşme' ? `Gruplu Eşleşme · Grup ${r.grup}`
      : r.durum;
    tr.innerHTML = `
      <td>${r.aT?fmtTarih(r.aT):'<span class="empty-cell">—</span>'}</td>
      <td>${r.aTip||''}</td>
      <td class="num">${fmtSayi(r.aBorc)}</td>
      <td class="num">${fmtSayi(r.aAlacak)}</td>
      <td>${r.aEvrak||''}</td>
      <td>${r.aEk||''}</td>
      <td>${r.bT?fmtTarih(r.bT):'<span class="empty-cell">—</span>'}</td>
      <td>${r.bTip||''}</td>
      <td class="num">${fmtSayi(r.bBorc)}</td>
      <td class="num">${fmtSayi(r.bAlacak)}</td>
      <td>${r.bEvrak||''}</td>
      <td>${r.bEk||''}</td>
      <td class="status">${durumText}</td>
    `;
    tbody.appendChild(tr);
  });
}

function applyFilter(){
  document.querySelectorAll('#tbody tr').forEach(tr=>{
    const show = currentFilter==='all' || tr.dataset.status===currentFilter;
    tr.classList.toggle('hidden', !show);
  });
}

document.getElementById('download-btn').addEventListener('click', ()=>{
  indirExcel();
});

async function indirExcel(){
  const firmaA = isimA.value.trim() || 'Kendi Firmanız';
  const firmaB = isimB.value.trim() || 'Karşı Firma';
  const visible = currentRows.filter(r=> currentFilter==='all' || r.durum===currentFilter);

  const RENK = {
    beyaz:{argb:'FFFFFFFF'},
    lacivert:{argb:'FF2F5496'},
    yesil:{argb:'FF548235'},
    gri:{argb:'FF7F7F7F'},
    eslesti:{argb:'FFE2EFDA'},
    onlyA:{argb:'FFFCE4D6'},
    onlyB:{argb:'FFDDEBF7'},
    fark:{argb:'FFFFF2CC'},
    grup:{argb:'FFEAE6F5'},
  };

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Karşılaştırma');

  const kolonlar = ['Tarih','Evrak Tipi','Borç (TL)','Alacak (TL)','Evrak No','Ek Bilgi'];
  const toplamKolon = kolonlar.length*2 + 1;

  ws.mergeCells(1,1,1,toplamKolon);
  const baslikH = ws.getCell(1,1);
  baslikH.value = `${firmaA.toUpperCase()} ile ${firmaB.toUpperCase()} MUTABAKATI`;
  baslikH.font = {name:'Arial', size:13, bold:true};

  const sayac = {'Eşleşti':0,'Tutar Farklı':0,'Gruplu Eşleşme':0,'only-a':0,'only-b':0};
  visible.forEach(r=>sayac[r.durum]++);
  ws.mergeCells(2,1,2,toplamKolon);
  const altH = ws.getCell(2,1);
  altH.value = `Toplam ${visible.length} satır  |  Eşleşen: ${sayac['Eşleşti']}  |  Gruplu: ${sayac['Gruplu Eşleşme']}  |  Yalnız ${firmaA}'de: ${sayac['only-a']}  |  Yalnız ${firmaB}'de: ${sayac['only-b']}  |  Tutar Farklı: ${sayac['Tutar Farklı']}`;
  altH.font = {name:'Arial', size:10, italic:true, color:{argb:'FF555555'}};

  ws.mergeCells(4,1,4,kolonlar.length);
  const grupA = ws.getCell(4,1);
  grupA.value = firmaA.toUpperCase();
  grupA.font = {name:'Arial', size:10, bold:true, color:RENK.beyaz};
  grupA.fill = {type:'pattern', pattern:'solid', fgColor:RENK.lacivert};
  grupA.alignment = {horizontal:'center'};

  ws.mergeCells(4,kolonlar.length+1,4,kolonlar.length*2);
  const grupB = ws.getCell(4,kolonlar.length+1);
  grupB.value = firmaB.toUpperCase();
  grupB.font = {name:'Arial', size:10, bold:true, color:RENK.beyaz};
  grupB.fill = {type:'pattern', pattern:'solid', fgColor:RENK.yesil};
  grupB.alignment = {horizontal:'center'};

  const grupD = ws.getCell(4,toplamKolon);
  grupD.value = 'DURUM';
  grupD.font = {name:'Arial', size:10, bold:true, color:RENK.beyaz};
  grupD.fill = {type:'pattern', pattern:'solid', fgColor:RENK.gri};
  grupD.alignment = {horizontal:'center'};

  const baslikSatiri = [...kolonlar, ...kolonlar, 'Durum'];
  baslikSatiri.forEach((ad, i)=>{
    const c = ws.getCell(5, i+1);
    c.value = ad;
    c.font = {name:'Arial', size:10, bold:true, color:RENK.beyaz};
    c.fill = {type:'pattern', pattern:'solid', fgColor: i<kolonlar.length ? RENK.lacivert : (i<kolonlar.length*2 ? RENK.yesil : RENK.gri)};
    c.alignment = {horizontal:'center', wrapText:true};
  });

  const durumRenk = {
    'Eşleşti': RENK.eslesti,
    'only-a': RENK.onlyA,
    'only-b': RENK.onlyB,
    'Tutar Farklı': RENK.fark,
    'Gruplu Eşleşme': RENK.grup,
  };

  visible.forEach((r, idx)=>{
    const rowNo = 6+idx;
    const durumText = r.durum==='only-a' ? `Yalnız ${firmaA}'de var`
      : r.durum==='only-b' ? `Yalnız ${firmaB}'de var`
      : r.durum==='Gruplu Eşleşme' ? `Gruplu Eşleşme · Grup ${r.grup}`
      : r.durum;
    const degerler = [
      r.aT?fmtTarih(r.aT):'', r.aTip||'', r.aBorc||'', r.aAlacak||'', r.aEvrak||'', r.aEk||'',
      r.bT?fmtTarih(r.bT):'', r.bTip||'', r.bBorc||'', r.bAlacak||'', r.bEvrak||'', r.bEk||'',
      durumText
    ];
    const fill = durumRenk[r.durum];
    degerler.forEach((v, i)=>{
      const c = ws.getCell(rowNo, i+1);
      c.value = v;
      c.font = {name:'Arial', size:10, bold: i===toplamKolon-1};
      if(fill) c.fill = {type:'pattern', pattern:'solid', fgColor:fill};
      if([2,3,8,9].includes(i)) c.numFmt = '#,##0.00';
    });
  });

  ws.columns = [
    {width:12},{width:20},{width:13},{width:13},{width:17},{width:22},
    {width:12},{width:20},{width:13},{width:13},{width:17},{width:22},
    {width:22},
  ];
  ws.views = [{state:'frozen', ySplit:5}];

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'mutabakat_raporu.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
