# Ekstrematik — Ekstre Mutabakat Aracı

İki firmanın cari hesap ekstresini (Mikro muhasebe programı formatında) karşılaştırıp
eşleşen / eşleşmeyen / gruplu eşleşen hareketleri tek bir raporda gösteren araç.

## Dosyalar

- **`index.html`** — Tamamen istemci taraflı (client-side) çalışan web arayüzü.
  Sunucu gerektirmez; doğrudan tarayıcıda açılabilir ya da herhangi bir statik
  hosting'e (Netlify, Vercel, GitHub Pages) yüklenebilir. Adım adım (4 aşamalı)
  bir akışla çalışır:
  1. Kendi firmanızın adı + ekstre dosyası
  2. Karşı firmanın adı + ekstre dosyası
  3. Onay ve "Mutabakatı Çalıştır"
  4. Renkli, filtrelenebilir rapor + Excel indirme

- **`ekstre_mutabakat.py`** — Aynı mantığın komut satırı (CLI) sürümü. Toplu/otomatik
  işler için ya da web arayüzü olmadan kullanmak isteyenler için.

## Kullanılan kütüphaneler (index.html içinde CDN üzerinden yükleniyor)

- [SheetJS (xlsx)](https://sheetjs.com/) — yüklenen .xlsx/.xls dosyalarını okumak için
- [ExcelJS](https://github.com/exceljs/exceljs) — renkli/biçimli Excel raporu **yazmak**
  için (SheetJS'in ücretsiz sürümü stil/renk yazamıyor, bu yüzden export için ayrı
  kütüphane kullanılıyor)

## Eşleştirme mantığı

1. **Evrak no eşleşmesi** — aynı evrak numarasına sahip, yönü (borç/alacak) ters olan
   hareketler eşleştirilir. Tutar farklıysa "Tutar Farklı" olarak işaretlenir.
2. **Birebir tutar eşleşmesi** — evrak no ile eşleşmeyenler için, tutarı aynı, yönü
   ters, tarihi en yakın olan hareketle eşleştirme yapılır.
3. **Gruplu eşleşme (N'e 1)** — bir tarafın tek satırının karşı tarafta birden fazla
   satıra bölünmüş olabileceği durumları (subset-sum / alt küme toplamı algoritması ile)
   yakalar. Örn: A'da 3.000 TL tek satır, B'de 1.000+1.000+1.000 üç satır.

## Desteklenen ekstre formatları

Şu an üç ekstre formatı tanınıyor (`FORMAT_TANIMLARI` altında tanımlı,
hem `assets/js/matching.js` içinde hem `ekstre_mutabakat.py` içinde):

- **hesap_karti**: `TARİH, EVRAK TİPİ, ANA DÖVİZ BORÇ, ANA DÖVİZ ALACAK, GIB FATURA NO, SORUMLU İSMİ`
- **cari_ekstre**: `İşlem Tarihi, İşlem Türü, Borç Tutarı, Alacak Tutarı, Evrak No, Açıklama`
- **logo_ekstre**: `TARIH, BELGE NO, BELGETURU, BORC, ALACAK, ACIKLAMA`

Yeni bir formatla karşılaşılırsa `FORMAT_TANIMLARI` sözlüğüne birkaç satırlık yeni bir
tanım eklemek yeterli.

## CLI kullanımı

```bash
python3 ekstre_mutabakat.py dosya1.xlsx dosya2.xls \
  -o rapor.xlsx \
  --firma1 "Kendi Firmanız" \
  --firma2 "Karşı Firma"
```

Gerekli paketler: `pandas`, `openpyxl`, `xlrd` (xls dosyaları için).

```bash
pip install pandas openpyxl xlrd
```

## Yapılabilecek geliştirmeler

- Daha fazla muhasebe programı formatı (Logo, Netsis vb.) desteği
- Kur farkı toleransı / farklı para birimleri
- Basit bir backend + kullanıcı hesabı eklenip gerçek bir SaaS'a dönüştürülmesi
- KVKK açısından, şu an tüm işlem tarayıcıda yapılıyor (veri hiçbir sunucuya
  gitmiyor) — bu backend eklenirse yeniden değerlendirilmeli
