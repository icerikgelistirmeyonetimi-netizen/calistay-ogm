/* Excel İndir eklentisi
 * Vue paketinden bağımsız çalışır: her sayfanın başlığına «Excel İndir» düğmesi
 * koyar, tıklanınca o sayfada gösterilen veriyi Supabase'den çekip .xlsx üretir.
 *  - Ana sayfa      → ders listesi (GM, ders, tamamlanma); «Tümünü İndir» → her ders ayrı sayfa
 *  - GM sayfası     → o GM'nin her dersi ayrı sayfa, önce «Özet»; «Tümünü İndir» → tüm GM'ler
 *  - Ders sayfası   → sınıf başına bir sayfa, ünite/öğrenme çıktısı/öncelikler
 *  - Ayarlar        → ders moderatörleri + içerik türleri
 */
(function () {
  'use strict';

  var SUPABASE_URL = 'https://ykmrystcfwjrrgkglyzr.supabase.co/rest/v1/';
  var SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlrbXJ5c3RjZndqcnJna2dseXpyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjQ5MzQyODIsImV4cCI6MjA4MDUxMDI4Mn0.T4yy_DqaDxilkAW8NMRbYiXiEyQNWknr04TD0sDC_R8';
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  var KUTU_ID = 'excel-indir-kutu';

  /* ---------- Supabase REST ---------- */
  async function sorgu(yol) {
    var tum = [];
    var baslangic = 0, adim = 1000;
    for (;;) {
      var r = await fetch(SUPABASE_URL + yol, {
        headers: {
          apikey: SUPABASE_KEY,
          Authorization: 'Bearer ' + SUPABASE_KEY,
          Range: baslangic + '-' + (baslangic + adim - 1)
        }
      });
      if (!r.ok && r.status !== 206) throw new Error('Supabase ' + r.status);
      var parca = await r.json();
      tum = tum.concat(parca);
      if (parca.length < adim) return tum;
      baslangic += adim;
    }
  }

  async function turHaritalari() {
    var turler = await sorgu('icerik_turleri?select=id,icerik_turu');
    var altTurler = await sorgu('icerik_alt_turleri?select=id,alt_tur_adi');
    var tur = {}, alt = {};
    turler.forEach(function (t) { tur[t.id] = t.icerik_turu; });
    altTurler.forEach(function (a) { alt[a.id] = a.alt_tur_adi; });
    return { tur: tur, alt: alt };
  }

  /* ---------- Sıralama yardımcıları (sayfadaki gruplama ile aynı) ---------- */
  function sinifSirala(a, b) {
    if (a === 'Belirtilmemiş') return 1;
    if (b === 'Belirtilmemiş') return -1;
    var na = parseInt(a, 10), nb = parseInt(b, 10);
    if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
    return a.localeCompare(b, 'tr');
  }

  function uniteSirala(a, b) {
    if (a === 'Ünite Belirtilmemiş') return 1;
    if (b === 'Ünite Belirtilmemiş') return -1;
    var na = a.match(/\d+/), nb = b.match(/\d+/);
    if (na && nb) return parseInt(na[0], 10) - parseInt(nb[0], 10);
    return a.localeCompare(b, 'tr');
  }

  function trSirala(a, b) { return String(a).localeCompare(String(b), 'tr'); }

  function grupla(dizi, anahtarFn) {
    var g = {};
    dizi.forEach(function (x) {
      var k = anahtarFn(x);
      (g[k] = g[k] || []).push(x);
    });
    return g;
  }

  /* ---------- Ad yardımcıları ---------- */
  function sayfaAdiYap(kullanilan, ad) {
    var temiz = String(ad).replace(/[\\/*?:\[\]]/g, '').trim().substring(0, 31) || 'Sayfa';
    var aday = temiz, n = 2;
    while (kullanilan[aday.toLowerCase()]) {
      var ek = ' (' + n + ')';
      aday = temiz.substring(0, 31 - ek.length) + ek;
      n++;
    }
    kullanilan[aday.toLowerCase()] = true;
    return aday;
  }

  function dosyaAdi(s) {
    var tr = { 'ç': 'c', 'Ç': 'C', 'ğ': 'g', 'Ğ': 'G', 'ı': 'i', 'İ': 'I', 'ö': 'o', 'Ö': 'O', 'ş': 's', 'Ş': 'S', 'ü': 'u', 'Ü': 'U' };
    return String(s).replace(/[çÇğĞıİöÖşŞüÜ]/g, function (c) { return tr[c]; })
      .replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
  }

  function bugun() { return new Date().toISOString().slice(0, 10); }

  /* ---------- Ders satırları (ders sayfasındaki tablo ile aynı düzen) ---------- */
  var ONCELIKLER = ['1. Öncelik', '2. Öncelik', '3. Öncelik', 'Zenginleştirme', 'Destekleme'];
  var ACIKLAMA_ALANLARI = ['tur_aciklama_1', 'tur_aciklama_2', 'tur_aciklama_3', 'zenginlestirme_aciklama', 'destekleme_aciklama'];

  function dersSatiri(k, unite, h, sinifEkle) {
    var s = {};
    if (sinifEkle) s['Sınıf'] = k.sinif || 'Belirtilmemiş';
    s['Ünite'] = unite;
    s['Öğrenme Çıktısı'] = k.kazanim || '';
    s['Süreç Bileşenleri'] = k.aciklama || '';
    ONCELIKLER.forEach(function (b, i) {
      var n = i + 1;
      s[b + ' Tür'] = h.tur[k['tur_' + n + '_id']] || '';
      s[b + ' Alt Tür'] = h.alt[k['alt_tur_' + n + '_id']] || '';
      s[b + ' Açıklama'] = k[ACIKLAMA_ALANLARI[i]] || '';
    });
    s['Diğer'] = k.diger_aciklama || '';
    return s;
  }

  function dersSutunGenislikleri(sinifEkle) {
    var g = sinifEkle ? [{ wch: 8 }] : [];
    return g.concat([{ wch: 36 }, { wch: 60 }, { wch: 50 }])
      .concat(ONCELIKLER.map(function () { return [{ wch: 18 }, { wch: 24 }, { wch: 40 }]; }).flat())
      .concat([{ wch: 40 }]);
  }

  // Bir sınıfın kayıtlarını sayfadaki sırayla satırlara çevirir
  function sinifSatirlari(kayitlar, h, sinifEkle) {
    var satirlar = [];
    var uniteler = grupla(kayitlar, function (k) {
      return (k.unite || 'Ünite Belirtilmemiş').replace(/·/g, '').trim();
    });
    Object.keys(uniteler).sort(uniteSirala).forEach(function (unite) {
      var grup = uniteler[unite].slice().sort(function (a, b) { return a.id - b.id; });
      var kazanimlilar = grup.filter(function (k) { return k.kazanim; });
      // Sayfa: ünitede kazanım varsa kazanım satırları; yoksa ünitenin kendisi tek satır
      (kazanimlilar.length ? kazanimlilar : [grup[0]]).forEach(function (k) {
        satirlar.push(dersSatiri(k, unite, h, sinifEkle));
      });
    });
    return satirlar;
  }

  // Bir dersin tüm kayıtları → sınıf sırasıyla tek liste (Sınıf sütunlu)
  function dersTumSatirlari(kayitlar, h) {
    var satirlar = [];
    var siniflar = grupla(kayitlar, function (k) { return k.sinif || 'Belirtilmemiş'; });
    Object.keys(siniflar).sort(sinifSirala).forEach(function (sinif) {
      satirlar = satirlar.concat(sinifSatirlari(siniflar[sinif], h, true));
    });
    return satirlar;
  }

  // Ders başına ayrı sayfa ekler
  function dersSayfalariniEkle(wb, kayitlar, h, kullanilan) {
    var dersler = grupla(kayitlar, function (k) { return k.ders_adi || 'Bilinmeyen'; });
    Object.keys(dersler).sort(trSirala).forEach(function (ders) {
      var ws = XLSX.utils.json_to_sheet(dersTumSatirlari(dersler[ders], h));
      ws['!cols'] = dersSutunGenislikleri(true);
      XLSX.utils.book_append_sheet(wb, ws, sayfaAdiYap(kullanilan, ders));
    });
  }

  function ozetSayfasi(dersler, gmSutunu) {
    var satirlar = dersler.map(function (d) {
      var s = {};
      if (gmSutunu) s['Genel Müdürlük'] = d.gm || '';
      s['Ders'] = d.ders_adi || '';
      s['Toplam İçerik'] = d.toplam_icerik || 0;
      s['Tamamlanan İçerik'] = d.tamamlanan_icerik || 0;
      s['Tamamlanma (%)'] = d.tamamlanma_yuzdesi == null ? '' : Number(d.tamamlanma_yuzdesi);
      s['Moderatör'] = d.moderator || '';
      s['Katılımcı'] = d.katilimci || '';
      return s;
    });
    var ws = XLSX.utils.json_to_sheet(satirlar);
    ws['!cols'] = (gmSutunu ? [{ wch: 44 }] : []).concat([{ wch: 48 }, { wch: 14 }, { wch: 18 }, { wch: 15 }, { wch: 36 }, { wch: 30 }]);
    return { ws: ws, satirlar: satirlar };
  }

  /* ---------- Sayfa bazlı çalışma kitapları ---------- */
  async function dersKitabi(dersAdi) {
    var kayitlar = await sorgu('icerik_kayitlari?select=*&ders_adi=eq.' + encodeURIComponent(dersAdi) + '&order=sinif,id');
    var h = await turHaritalari();
    var wb = XLSX.utils.book_new();
    var kullanilan = {};
    var siniflar = grupla(kayitlar, function (k) { return k.sinif || 'Belirtilmemiş'; });
    Object.keys(siniflar).sort(sinifSirala).forEach(function (sinif) {
      var ws = XLSX.utils.json_to_sheet(sinifSatirlari(siniflar[sinif], h, false));
      ws['!cols'] = dersSutunGenislikleri(false);
      XLSX.utils.book_append_sheet(wb, ws, sayfaAdiYap(kullanilan, 'Sınıf ' + sinif));
    });
    if (!wb.SheetNames.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Bu ders için içerik bulunamadı.']]), 'Bilgi');
    }
    return { wb: wb, ad: 'EBA_' + dosyaAdi(dersAdi) + '_' + bugun() + '.xlsx' };
  }

  async function gmKitabi(gm) {
    var ozet = await sorgu('ders_ozet_view?select=*&gm=eq.' + encodeURIComponent(gm) + '&order=ders_adi');
    var kayitlar = await sorgu('icerik_kayitlari?select=*&gm=eq.' + encodeURIComponent(gm) + '&order=ders_adi,sinif,id');
    var h = await turHaritalari();
    var wb = XLSX.utils.book_new();
    var kullanilan = {};
    XLSX.utils.book_append_sheet(wb, ozetSayfasi(ozet, false).ws, sayfaAdiYap(kullanilan, 'Özet'));
    dersSayfalariniEkle(wb, kayitlar, h, kullanilan);
    return { wb: wb, ad: 'EBA_' + dosyaAdi(gm) + '_' + bugun() + '.xlsx' };
  }

  async function tumuKitabi() {
    var ozet = await sorgu('ders_ozet_view?select=*&order=gm,ders_adi');
    var kayitlar = await sorgu('icerik_kayitlari?select=*&order=gm,ders_adi,sinif,id');
    var h = await turHaritalari();
    var wb = XLSX.utils.book_new();
    var kullanilan = {};
    XLSX.utils.book_append_sheet(wb, ozetSayfasi(ozet, true).ws, sayfaAdiYap(kullanilan, 'Özet'));
    // GM sırasıyla, her GM içinde ders adına göre
    var gmler = grupla(kayitlar, function (k) { return k.gm || 'Diğer'; });
    Object.keys(gmler).sort(trSirala).forEach(function (gm) {
      dersSayfalariniEkle(wb, gmler[gm], h, kullanilan);
    });
    return { wb: wb, ad: 'EBA_Tum_Dersler_' + bugun() + '.xlsx' };
  }

  async function anaSayfaKitabi() {
    var dersler = await sorgu('ders_ozet_view?select=*&order=gm,ders_adi');
    var wb = XLSX.utils.book_new();
    var kullanilan = {};
    var o = ozetSayfasi(dersler, true);
    XLSX.utils.book_append_sheet(wb, o.ws, sayfaAdiYap(kullanilan, 'Dersler'));
    // Sayfadaki gibi GM başına ayrı sayfa
    var gmler = grupla(o.satirlar, function (s) { return s['Genel Müdürlük'] || 'Diğer'; });
    Object.keys(gmler).sort(trSirala).forEach(function (gm) {
      var w = XLSX.utils.json_to_sheet(gmler[gm]);
      w['!cols'] = o.ws['!cols'];
      XLSX.utils.book_append_sheet(wb, w, sayfaAdiYap(kullanilan, gm));
    });
    return { wb: wb, ad: 'EBA_Ders_Listesi_' + bugun() + '.xlsx' };
  }

  async function ayarlarKitabi() {
    var dersler = await sorgu('ders_ozet_view?select=ders_adi,moderator,gm&order=ders_adi');
    var altTurler = await sorgu('icerik_alt_turleri?select=*&order=alt_tur_adi');
    var h = await turHaritalari();
    var wb = XLSX.utils.book_new();
    var wsM = XLSX.utils.json_to_sheet(dersler.map(function (d) {
      return { 'Ders Adı': d.ders_adi || '', 'Moderatör': d.moderator || '', 'Genel Müdürlük': d.gm || '' };
    }));
    wsM['!cols'] = [{ wch: 48 }, { wch: 40 }, { wch: 44 }];
    XLSX.utils.book_append_sheet(wb, wsM, 'Ders Moderatörleri');
    var wsT = XLSX.utils.json_to_sheet(altTurler.map(function (a) {
      return { 'İçerik Türü': h.tur[a.icerik_turu_id] || '', 'Alt Tür': a.alt_tur_adi || '', 'Açıklama': a.aciklama || '' };
    }).sort(function (a, b) {
      return trSirala(a['İçerik Türü'], b['İçerik Türü']) || trSirala(a['Alt Tür'], b['Alt Tür']);
    }));
    wsT['!cols'] = [{ wch: 24 }, { wch: 36 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, wsT, 'İçerik Türleri');
    return { wb: wb, ad: 'EBA_Ayarlar_' + bugun() + '.xlsx' };
  }

  function rotaBilgisi() {
    var hash = decodeURIComponent(location.hash || '#/').replace(/\?.*$/, '');
    var m = hash.match(/^#\/ders\/(.+)$/);
    if (m) return { tur: 'ders', dersAdi: m[1] };
    m = hash.match(/^#\/gm\/(.+)$/);
    if (m) return { tur: 'gm', gm: m[1] };
    if (/^#\/tur-yonetimi/.test(hash)) return { tur: 'ayarlar' };
    return { tur: 'ana' };
  }

  async function kitapOlustur(tumu) {
    if (tumu) return tumuKitabi();
    var r = rotaBilgisi();
    if (r.tur === 'ders') return dersKitabi(r.dersAdi);
    if (r.tur === 'gm') return gmKitabi(r.gm);
    if (r.tur === 'ayarlar') return ayarlarKitabi();
    return anaSayfaKitabi();
  }

  /* ---------- XLSX kütüphanesini gerektiğinde yükle ---------- */
  var xlsxSozu = null;
  function xlsxYukle() {
    if (window.XLSX) return Promise.resolve();
    if (!xlsxSozu) {
      xlsxSozu = new Promise(function (coz, reddet) {
        var s = document.createElement('script');
        s.src = XLSX_URL;
        s.onload = function () { coz(); };
        s.onerror = function () { xlsxSozu = null; reddet(new Error('xlsx kütüphanesi yüklenemedi')); };
        document.head.appendChild(s);
      });
    }
    return xlsxSozu;
  }

  /* ---------- Düğmeler ---------- */
  var calisiyor = false;
  async function indir(dugme, tumu) {
    if (calisiyor) return;
    calisiyor = true;
    var eskiMetin = dugme.innerHTML;
    dugme.disabled = true;
    dugme.innerHTML = ikon() + ' Hazırlanıyor…';
    try {
      await xlsxYukle();
      var k = await kitapOlustur(tumu);
      XLSX.writeFile(k.wb, k.ad);
    } catch (e) {
      console.error('Excel indirme hatası:', e);
      alert('Excel dosyası hazırlanamadı: ' + (e && e.message ? e.message : e));
    } finally {
      dugme.disabled = false;
      dugme.innerHTML = eskiMetin;
      calisiyor = false;
    }
  }

  function ikon() {
    return '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';
  }

  function dugmeOlustur(metin, ipucu, renk, koyu, tumu) {
    var b = document.createElement('button');
    b.type = 'button';
    b.title = ipucu;
    b.className = 'px-4 py-2 text-white text-sm font-medium rounded-lg transition-colors shadow-sm';
    b.style.cssText = 'display:inline-flex;align-items:center;gap:6px;white-space:nowrap;background-color:' + renk + ';color:#fff;border:0;cursor:pointer;';
    b.innerHTML = ikon() + ' ' + metin;
    b.addEventListener('mouseenter', function () { b.style.backgroundColor = koyu; });
    b.addEventListener('mouseleave', function () { b.style.backgroundColor = renk; });
    b.addEventListener('click', function () { indir(b, tumu); });
    return b;
  }

  function kutuOlustur() {
    var r = rotaBilgisi();
    var kutu = document.createElement('div');
    kutu.id = KUTU_ID;
    kutu.setAttribute('data-rota', r.tur);
    kutu.style.cssText = 'display:inline-flex;align-items:center;gap:8px;';
    var ipucu = {
      ana: 'Ders listesini Excel olarak indir',
      gm: 'Bu genel müdürlüğün her dersini ayrı sayfada Excel olarak indir',
      ders: 'Bu dersi sınıf başına ayrı sayfada Excel olarak indir',
      ayarlar: 'Moderatör ve içerik türü listelerini Excel olarak indir'
    }[r.tur];
    kutu.appendChild(dugmeOlustur('Excel İndir', ipucu, '#16a34a', '#15803d', false));
    if (r.tur === 'ana' || r.tur === 'gm') {
      kutu.appendChild(dugmeOlustur('Tümünü İndir', 'Bütün genel müdürlüklerin her dersini ayrı sayfada tek Excel dosyasında indir', '#2563eb', '#1d4ed8', true));
    }
    return kutu;
  }

  function yerlestir() {
    var header = document.querySelector('#app header');
    if (!header) return;
    var mevcut = header.querySelector('#' + KUTU_ID);
    if (mevcut && mevcut.getAttribute('data-rota') === rotaBilgisi().tur) return;
    var eski = document.getElementById(KUTU_ID);
    if (eski) eski.remove();

    var satir = header.querySelector('.flex.items-center');
    if (!satir) return;
    var kutu = kutuOlustur();
    // Ana sayfa: sağdaki giriş kutusunun önüne; ayarlar: «Ana Sayfa» bağlantısının önüne; ders/GM: satırın sonuna
    var sag = satir.querySelector(':scope > .flex.items-center');
    if (sag) {
      sag.insertBefore(kutu, sag.firstChild);
    } else if (satir.classList.contains('justify-between') && satir.children.length > 1) {
      var sarmal = document.createElement('div');
      sarmal.style.cssText = 'display:flex;align-items:center;gap:8px;';
      var son = satir.lastElementChild;
      satir.insertBefore(sarmal, son);
      sarmal.appendChild(kutu);
      sarmal.appendChild(son);
    } else {
      kutu.style.marginLeft = '8px';
      satir.appendChild(kutu);
    }
  }

  function baslat() {
    yerlestir();
    var kok = document.getElementById('app') || document.body;
    new MutationObserver(function () { yerlestir(); }).observe(kok, { childList: true, subtree: true });
    window.addEventListener('hashchange', function () { setTimeout(yerlestir, 50); });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', baslat);
  else baslat();

  // Sınama için
  window.__excelIndir = { kitapOlustur: kitapOlustur, rotaBilgisi: rotaBilgisi, xlsxYukle: xlsxYukle };
})();
