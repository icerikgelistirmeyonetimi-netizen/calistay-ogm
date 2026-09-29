/* Excel İndir eklentisi
 * Vue paketinden bağımsız çalışır: her sayfanın başlığına «Excel İndir» düğmesi
 * koyar, tıklanınca o sayfada gösterilen veriyi Supabase'den çekip .xlsx üretir.
 *  - Ana sayfa      → ders listesi (GM, ders, tamamlanma)
 *  - Ders sayfası   → sınıf başına bir sayfa, ünite/öğrenme çıktısı/öncelikler
 *  - Ayarlar        → ders moderatörleri + içerik türleri
 */
(function () {
  'use strict';

  var SUPABASE_URL = 'https://ykmrystcfwjrrgkglyzr.supabase.co/rest/v1/';
  var SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InlrbXJ5c3RjZndqcnJna2dseXpyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjQ5MzQyODIsImV4cCI6MjA4MDUxMDI4Mn0.T4yy_DqaDxilkAW8NMRbYiXiEyQNWknr04TD0sDC_R8';
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  var DUGME_ID = 'excel-indir-dugme';

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

  function turHaritalari(turler, altTurler) {
    var tur = {}, alt = {};
    (turler || []).forEach(function (t) { tur[t.id] = t.icerik_turu; });
    (altTurler || []).forEach(function (a) { alt[a.id] = a.alt_tur_adi; });
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

  function grupla(dizi, anahtarFn) {
    var g = {};
    dizi.forEach(function (x) {
      var k = anahtarFn(x);
      (g[k] = g[k] || []).push(x);
    });
    return g;
  }

  /* ---------- Sayfa bazlı çalışma kitapları ---------- */
  function sayfaAdi(s) {
    return String(s).replace(/[\\/*?:\[\]]/g, '').substring(0, 31) || 'Sayfa';
  }

  function dosyaAdi(s) {
    var tr = { 'ç': 'c', 'Ç': 'C', 'ğ': 'g', 'Ğ': 'G', 'ı': 'i', 'İ': 'I', 'ö': 'o', 'Ö': 'O', 'ş': 's', 'Ş': 'S', 'ü': 'u', 'Ü': 'U' };
    return String(s).replace(/[çÇğĞıİöÖşŞüÜ]/g, function (c) { return tr[c]; })
      .replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '');
  }

  function bugun() { return new Date().toISOString().slice(0, 10); }

  async function dersKitabi(dersAdi) {
    var kayitlar = await sorgu('icerik_kayitlari?select=*&ders_adi=eq.' + encodeURIComponent(dersAdi) + '&order=sinif,id');
    var turler = await sorgu('icerik_turleri?select=id,icerik_turu');
    var altTurler = await sorgu('icerik_alt_turleri?select=id,alt_tur_adi');
    var h = turHaritalari(turler, altTurler);

    function satir(k, unite) {
      var s = {
        'Ünite': unite,
        'Öğrenme Çıktısı': k.kazanim || '',
        'Süreç Bileşenleri': k.aciklama || ''
      };
      var basliklar = ['1. Öncelik', '2. Öncelik', '3. Öncelik', 'Zenginleştirme', 'Destekleme'];
      var aciklamaAlanlari = ['tur_aciklama_1', 'tur_aciklama_2', 'tur_aciklama_3', 'zenginlestirme_aciklama', 'destekleme_aciklama'];
      basliklar.forEach(function (b, i) {
        var n = i + 1;
        s[b + ' Tür'] = h.tur[k['tur_' + n + '_id']] || '';
        s[b + ' Alt Tür'] = h.alt[k['alt_tur_' + n + '_id']] || '';
        s[b + ' Açıklama'] = k[aciklamaAlanlari[i]] || '';
      });
      s['Diğer'] = k.diger_aciklama || '';
      return s;
    }

    var wb = XLSX.utils.book_new();
    var siniflar = grupla(kayitlar, function (k) { return k.sinif || 'Belirtilmemiş'; });
    Object.keys(siniflar).sort(sinifSirala).forEach(function (sinif) {
      var uniteler = grupla(siniflar[sinif], function (k) {
        return (k.unite || 'Ünite Belirtilmemiş').replace(/·/g, '').trim();
      });
      var satirlar = [];
      Object.keys(uniteler).sort(uniteSirala).forEach(function (unite) {
        var grup = uniteler[unite].slice().sort(function (a, b) { return a.id - b.id; });
        var kazanimlilar = grup.filter(function (k) { return k.kazanim; });
        // Sayfa: ünitede kazanım varsa kazanım satırları; yoksa ünitenin kendisi tek satır
        (kazanimlilar.length ? kazanimlilar : [grup[0]]).forEach(function (k) {
          satirlar.push(satir(k, unite));
        });
      });
      var ws = XLSX.utils.json_to_sheet(satirlar);
      ws['!cols'] = [{ wch: 36 }, { wch: 60 }, { wch: 50 }]
        .concat([1, 2, 3, 4, 5].map(function () { return [{ wch: 18 }, { wch: 24 }, { wch: 40 }]; }).flat())
        .concat([{ wch: 40 }]);
      XLSX.utils.book_append_sheet(wb, ws, sayfaAdi('Sınıf ' + sinif));
    });
    if (!wb.SheetNames.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Bu ders için içerik bulunamadı.']]), 'Bilgi');
    }
    return { wb: wb, ad: 'EBA_' + dosyaAdi(dersAdi) + '_' + bugun() + '.xlsx' };
  }

  async function anaSayfaKitabi() {
    var dersler = await sorgu('ders_ozet_view?select=*&order=gm,ders_adi');
    var satirlar = dersler.map(function (d) {
      return {
        'Genel Müdürlük': d.gm || '',
        'Ders': d.ders_adi || '',
        'Toplam İçerik': d.toplam_icerik || 0,
        'Tamamlanan İçerik': d.tamamlanan_icerik || 0,
        'Tamamlanma (%)': d.tamamlanma_yuzdesi == null ? '' : Number(d.tamamlanma_yuzdesi),
        'Moderatör': d.moderator || '',
        'Katılımcı': d.katilimci || ''
      };
    });
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.json_to_sheet(satirlar);
    ws['!cols'] = [{ wch: 44 }, { wch: 48 }, { wch: 14 }, { wch: 18 }, { wch: 15 }, { wch: 36 }, { wch: 30 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Dersler');
    // Sayfadaki gibi GM başına ayrı sayfa
    var gmler = grupla(satirlar, function (s) { return s['Genel Müdürlük'] || 'Diğer'; });
    Object.keys(gmler).sort(function (a, b) { return a.localeCompare(b, 'tr'); }).forEach(function (gm) {
      var w = XLSX.utils.json_to_sheet(gmler[gm]);
      w['!cols'] = ws['!cols'];
      XLSX.utils.book_append_sheet(wb, w, sayfaAdi(gm));
    });
    return { wb: wb, ad: 'EBA_Ders_Listesi_' + bugun() + '.xlsx' };
  }

  async function ayarlarKitabi() {
    var dersler = await sorgu('ders_ozet_view?select=ders_adi,moderator,gm&order=ders_adi');
    var turler = await sorgu('icerik_turleri?select=*&order=icerik_turu');
    var altTurler = await sorgu('icerik_alt_turleri?select=*&order=alt_tur_adi');
    var h = turHaritalari(turler, altTurler);
    var wb = XLSX.utils.book_new();
    var wsM = XLSX.utils.json_to_sheet(dersler.map(function (d) {
      return { 'Ders Adı': d.ders_adi || '', 'Moderatör': d.moderator || '', 'Genel Müdürlük': d.gm || '' };
    }));
    wsM['!cols'] = [{ wch: 48 }, { wch: 40 }, { wch: 44 }];
    XLSX.utils.book_append_sheet(wb, wsM, 'Ders Moderatörleri');
    var wsT = XLSX.utils.json_to_sheet(altTurler.map(function (a) {
      return { 'İçerik Türü': h.tur[a.icerik_turu_id] || '', 'Alt Tür': a.alt_tur_adi || '', 'Açıklama': a.aciklama || '' };
    }).sort(function (a, b) {
      return a['İçerik Türü'].localeCompare(b['İçerik Türü'], 'tr') || a['Alt Tür'].localeCompare(b['Alt Tür'], 'tr');
    }));
    wsT['!cols'] = [{ wch: 24 }, { wch: 36 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, wsT, 'İçerik Türleri');
    return { wb: wb, ad: 'EBA_Ayarlar_' + bugun() + '.xlsx' };
  }

  function rotaBilgisi() {
    var hash = decodeURIComponent(location.hash || '#/');
    var m = hash.match(/^#\/ders\/(.+)$/);
    if (m) return { tur: 'ders', dersAdi: m[1].replace(/\?.*$/, '') };
    if (/^#\/tur-yonetimi/.test(hash)) return { tur: 'ayarlar' };
    return { tur: 'ana' };
  }

  async function kitapOlustur() {
    var r = rotaBilgisi();
    if (r.tur === 'ders') return dersKitabi(r.dersAdi);
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

  /* ---------- Düğme ---------- */
  var calisiyor = false;
  async function indir(dugme) {
    if (calisiyor) return;
    calisiyor = true;
    var eskiMetin = dugme.innerHTML;
    dugme.disabled = true;
    dugme.innerHTML = ikon() + ' Hazırlanıyor…';
    try {
      await xlsxYukle();
      var k = await kitapOlustur();
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

  function dugmeOlustur() {
    var b = document.createElement('button');
    b.id = DUGME_ID;
    b.type = 'button';
    b.title = 'Bu sayfadaki verileri Excel olarak indir';
    b.className = 'px-4 py-2 bg-green-600 hover:bg-green-700 text-white text-sm font-medium rounded-lg transition-colors shadow-sm';
    b.style.cssText = 'display:inline-flex;align-items:center;gap:6px;white-space:nowrap;background-color:#16a34a;color:#fff;border:0;cursor:pointer;';
    b.innerHTML = ikon() + ' Excel İndir';
    b.addEventListener('mouseenter', function () { b.style.backgroundColor = '#15803d'; });
    b.addEventListener('mouseleave', function () { b.style.backgroundColor = '#16a34a'; });
    b.addEventListener('click', function () { indir(b); });
    return b;
  }

  function yerlestir() {
    var header = document.querySelector('#app header');
    if (!header) return;
    if (header.querySelector('#' + DUGME_ID)) return;
    var eski = document.getElementById(DUGME_ID);
    if (eski) eski.remove();

    var satir = header.querySelector('.flex.items-center');
    if (!satir) return;
    var dugme = dugmeOlustur();
    // Ana sayfa: sağdaki giriş kutusunun önüne; ayarlar: «Ana Sayfa» bağlantısının önüne; ders: satırın sonuna
    var sag = satir.querySelector(':scope > .flex.items-center');
    if (sag) {
      sag.insertBefore(dugme, sag.firstChild);
    } else if (satir.classList.contains('justify-between') && satir.children.length > 1) {
      var sarmal = document.createElement('div');
      sarmal.style.cssText = 'display:flex;align-items:center;gap:8px;';
      var son = satir.lastElementChild;
      satir.insertBefore(sarmal, son);
      sarmal.appendChild(dugme);
      sarmal.appendChild(son);
    } else {
      dugme.style.marginLeft = '8px';
      satir.appendChild(dugme);
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
