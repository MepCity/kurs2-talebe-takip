# GEÇMİŞ — Talebe Takip Sistemi Proje Tarihi

> Bu dosya, bir yapay zekânın (veya yeni bir geliştiricinin) repoyu hızla tanıması için yazıldı.
> Ne yaptık, neden yaptık, nasıl yaptık — kronolojik ve gerekçeli.

## Sistem nedir?

Kur'an kursu öğrenci takip uygulaması. Hocalar telefondan yoklama, Nurlu Kartlar ezberi,
sûre/dua ezberi, Elif-Ba ödevi, namaz vakitleri ve Kur'an sayfası takibi yapar; her
işaretleme anında Google E-Tabloya yazılır ve diğer hocaların ekranına saniyeler içinde düşer.

**Toplam 7 sınıf, 7 site, 2 Google hesabı, 2 Apps Script motoru:**

| | Fındıklı | 6 sınıf (Bağdat, Endülüs, Gazze, Kudüs, Medine, Mekke) |
|---|---|---|
| Site | GitHub Pages (bu repo: `MepCity/kurs2-talebe-takip`) | Vercel: `elmas-{sinif}.vercel.app` (repo: `mephisto4627/TalebeTakip`) |
| Apps Script | Yeni hesap (Google One'lı), güncellenebilir | Eski hesap, **DONMUŞ** (aşağıya bak) |
| E-Tablo | Yeni hesapta (`1vXnTCel...Z8rQ`) | Eski hesapta, 6 ayrı tablo |
| Poll aralığı | 4 sn | 7 sn |

## Mimari — nasıl çalışır?

1. **Site** (`index.html`): tek dosyalık uygulama. Kod, bundler çıktısı olarak 239. satırda
   ~370K karakterlik TEK satırda durur; satır sonları literal `\n` olarak escape'lidir.
   **Düzenleme Edit araçlarıyla değil, python `str.replace` ile yapılır** (önce `count==1`
   doğrula). `public/index.html` birebir kopyadır — her değişiklik İKİSİNE de uygulanır.
2. **Apps Script** (`apps-script/Code.gs`): Sheets okuma/yazma motoru. Web app olarak deploy
   edilir (Execute as: Me, Access: Anyone). Site `/exec` adresine POST atar:
   `{sinif: 'findikli', changes: [{type: 'att', ...}]}`.
3. **Veri akışı:** Tıklama → ekranda anında görünür (optimistic) → `outbox` kuyruğuna girer →
   sunucuya gönderilir → sunucu versiyon sayacını artırır → diğer telefonlar 4 sn'de bir
   sadece sayacı sorar (`readVersion`, ucuz) → sayaç değiştiyse tam veriyi çeker
   (`readAllAttendance`, sunucuda 120 sn önbellekli, versiyon anahtarlı).
4. **Elle Excel düzenlemesi:** `onEdit` tetikleyicisi sayacı artırır → telefonlar ~5 sn'de görür.
5. **Sınıf yönlendirme:** `SINIFLAR` haritası (Code.gs başında) sınıf kodu → E-Tablo ID eşler.
   Site her istekte `sinif` gönderir. Versiyon/önbellek anahtarları sınıf öneklidir
   (`v_findikli`, `att_bagdat_42` gibi); istek sayacı (`req_<PasifikTarihi>`) bilerek ortaktır.

## Kronoloji — ne yaptık, neden?

### 1. Başlangıç (8 Tem ve öncesi)
- `Talebe Takip.html`: ilk sürüm, sadece localStorage. Sonra Apps Script eşitlemesi eklendi.
- Ghost öğrenci sorunu: öğrenci listesinin tek kaynağı Sheets Yoklama sayfası yapıldı.
- 7 sn polling ile eş zamanlılık geldi (commit `85d8750`, `f04ce70`).

### 2. Boş localStorage çökmesi (9 Tem gecesi, `946a3b1`)
Siteyi İLK kez açanlarda "Cannot read properties of undefined (reading 'length')" çöküyordu.
**Neden:** `ensureData()` varsayılan verisinde `sureList` ve `log` alanları yoktu; eski
kullanıcılarda localStorage'dan geliyordu. **Ders:** yeni veri alanı eklerken MUTLAKA
`ensureData()` içinde varsayılanını seed'le.

### 3. Versiyon sayaçlı senkronizasyon (`b269402`)
Yayılım 30-40 sn idi çünkü her poll tüm sayfayı okuyordu. Çözüm: her yazma
`CacheService`/`PropertiesService`'te sayaç artırır; poll önce sayacı sorar (~0.5 sn),
değiştiyse tam çeker. Okumalar sunucuda versiyon-anahtarlı önbelleğe alındı (120 sn TTL).
Silme yayılımı düzeltildi (merge artık boş değerleri de işler). `visibilitychange` ile
sekmeye dönüşte anında kontrol.

### 4. Günlük yoklama kolonu + başlık filtresi (`b7d3630`)
Gece yarısı geçince yeni günün kolonu Sheets'te yoktu; okuma boş L2'yi atlayıp W2'deki
"🔤 A–Z" özet kolonunu tarih sanıyordu. Çözüm: `isDateLabel_` filtresi (sadece "9 Tem"
kalıbı tarih sayılır) + `ensureTodayColumn_` (hafta içi her gün ilk istekte kolonu açar,
İstanbul saati, günde bir kez cache'li).

### 5. UI iyileştirmeleri (`3481cfe`, `59d6f64`, `0f11bee`, `4bfe10b`, `38b8b08`, `79d3fa4`)
- Ezber popup kapanınca kaldığı yere yumuşak dönüş (açılışta zaten yukarı kayıyordu).
- Popup açılışına baloncuk efekti. **KRİTİK KURAL:** bu projede animasyonlarda `transform`/
  `filter` ANCESTOR'larda yasak (iOS Safari `position:fixed`'i bozar — geçmişte saatler
  kaybettirdi). Panel gibi fixed içermeyen yaprak elemanlarda kontrollü kullanılabilir.
- Ayarlar'dan E-Tablo adresi/test/eşitle kartı tamamen kaldırıldı (adres koda gömülü).
- İşlem geçmişine filtreler: öğrenci adı (yazarken öneri çıkar), tarih seçici, kart no —
  tekli/ikili/üçlü kombinasyonlar.

### 6. Hoca listesi ve işlem geçmişi Sheets'e taşındı (`7af5927`)
Eskiden cihaza özeldi (localStorage), tarayıcı temizliğinde kayboluyordu. Artık E-Tabloda
"Hocalar" ve "İşlem Geçmişi" sayfaları var (script yoksa kendisi oluşturur). Her telefon ilk
bağlantıda kendi birikmiş verisini bir defalık yükler (`logBulk`, `hocaSeedDone` bayrağı).
Hoca silme/yeniden adlandırma Ayarlar'a eklendi (`4c3f93d`), cihazlar arası eşitlenir.

### 7. İsim eşleştirme (`150ff18`)
Sayfalar arasında isimler tutarsızdı: Yoklama'da "Abdullah Altun", Nurlu'da "Abdullah".
Birebir eşleşme bulunamazsa **güvenli ön ek eşleşmesi** yapılır — ama satırdaki isim,
yoklamadaki BAŞKA bir öğrencinin tam adıysa eşleştirilmez ("Yiğit" / "Yiğit Hamza" karışmaz;
bulamazsa boş gösterir, asla yanlış çocuğu göstermez). Ayrıca site bir ara mükerrer satırlar
açmıştı ("Abdullah Altun" boş satırı) — elle temizlendi, isimler Yoklama yazımıyla eşitlendi.

### 8. onEdit tetikleyicisi (`e253af5`)
Excel'den elle konan artılar siteye 2-3 dakikada düşüyordu (sayaç artmıyordu). `onEdit`
her elle düzenlemede sayacı artırır → ~5 sn'ye indi.

### 9. Kota yönetimi (`fa972c9`)
Google ücretsiz hesap sınırı ~20.000 istek/gün (Pasifik gece yarısı = TR sabah ~10:00
sıfırlanır). Poll 7 sn yapıldı, sunucu her yanıtta günün istek sayısını döner (`reqToday`),
Ayarlar'ın altında gösterge var, 12.000 üzeri 2×, 16.000 üzeri 4× otomatik yavaşlama.

### 10. Çoklu sınıf altyapısı (`4252c73`, `cd504bd`)
6 yeni sınıf kararı: **tek merkezi Apps Script + sınıf başına site**. `SINIFLAR` haritası,
`sinif` parametresi, sınıf önekli önbellek anahtarları. Sınıflar arası YAPI FARKLARI:
- Yeni düzen (6 sınıf): Nurlu kartlarda **3 madde** (Vecize/İlmihal/Kelime, 60 üzerinden);
  Fındıklı'da 4 (V/Dua-Sûre/İ/K, 80 üzerinden). `duzen: 'eski'|'yeni'` alanı yönetir.
- Elif-Ba: Kudüs/Medine/Mekke'de var (isim B kolonu, satır 3); Fındıklı'da farklı konum.
- Kuran Takip: Bağdat/Endülüs/Gazze + Fındıklı'da var. Yeni sınıflarda öğrenci kartında
  "Kur'an Sayfası" kutusu → günün koluna sayı yazar (`kuranDaily`).
- Namaz: yeni sınıflarda vakit SAYISI yazılır (3 gibi), Fındıklı'da vakit adları.
- Ezber Takip 2: sadece Bağdat/Endülüs — sitede "Sûre Ezberi 2" bölümü (`sure2`).
- Sûre/dua listeleri her sınıfın tablosundan okunur (`sureList`, `sureList2` meta).
- Tarih başlık satırı sınıfa göre 2 veya 4 → `dateHeaderRow_` otomatik bulur.
- "Tel" sayfasına asla dokunulmaz.
- Eksik Nurlu sayfası ilk işaretlemede otomatik oluşturulur, Yoklama'daki öğrenciler
  sırayla hazır yazılır (`3bd41cf`).

### 11. 6 sınıf sitesi + Vercel (`mephisto4627/TalebeTakip` reposu)
`_sablon.html` = kurs2 index.html'in uyarlanmış kopyası (sinif dinamik: `window.SINIF_KODU`).
Sınıf klasörleri şablondan üretilir (fark: `<title>` + SINIF_KODU satırı). Vercel'de aynı
repodan 6 proje (Root Directory = klasör). Varsayılan hoca "Test Hoca" (Fındıklı hocaları
şablondan temizlendi). Şablon değişince 6 dosya yeniden üretilip push edilir.

### 12. Hesap taşınması (9 Tem sabahı — ÖNEMLİ KISIT)
Eski Google hesabının depolaması TAMAMEN doldu → eski Apps Script **düzenlenemez oldu**
(çalışmaya devam ediyor). Fındıklı yeni hesaba taşındı: yeni tablo + yeni script + yeni
`/exec` adresi (`AKfycbzs7szSs...`). **Karar: 6 sınıf eski hesapta kalacak.** Sonuç:
- Arka plan (Code.gs) değişiklikleri artık SADECE Fındıklı'ya deploy edilebilir.
- 6 sınıfın motoru mevcut haliyle sabittir; site tarafı (Vercel) hâlâ güncellenebilir.
- Kota iki hesaba bölündü → Fındıklı poll'u 4 sn'ye hızlandırıldı (`bbb3d25`).

### 13. Tıklama kaybı düzeltmesi (bu commit)
Kullanım sırasında 5-6 tıklamadan biri "gelip siliniyordu". **Neden (yarış durumu):**
`flush()` gönderime başlarken kuyruğu hemen boşaltıyordu; gönderim yoldayken 4 sn'lik
eşitleme tam listeyi çekip ekranı sunucu fotoğrafına göre yeniden kurunca, henüz işlenmemiş
tıklama ne sunucuda ne kuyruktaydı → siliniyordu. **Çözüm:** kuyruk sunucu onaylayana kadar
boşalmaz (`_flushing` bayrağı, onayda `filter` ile düşülür); yeniden kurulumda bekleyenler
her zaman üste yazıldığı için işaret kaybolamaz. Önce Fındıklı'da denenir, sonra 6 sınıfa.

## Bilinen tuzaklar / kurallar (AI için)

1. `index.html` TEK DEV SATIR, escape'li — python replace ile düzenle, iki kopyaya da uygula.
2. Yeni veri alanı → `ensureData()` içinde seed et, yoksa ilk açılış çöker.
3. Animasyonda ancestor'lara transform/filter verme (iOS fixed bozulur).
4. Sheets okumasında başlıkları `isDateLabel_`/özet filtresinden geçir (A–Z, Toplam kolonları).
5. İsimler sayfalar arasında birebir aynı yazılmalı; akıllı eşleştirme yedek, çare değil.
6. Deploy: Fındıklı → yeni hesap script'i, "Manage deployments → New version" (adres sabit
   kalır). "New deployment" YAPMA — adres değişir, siteler kopar.
7. 6 sınıfın Code.gs'i güncellenemez — backend özelliği eklerken bunu hesaba kat.
8. Kurs bitişi: `SON_DERS_GUNU = '20260724'` — 25 Tem'den sonra yeni yoklama kolonu açılmaz
   (frontend'de de aynı tarih kontrolü var).
9. Yoklama tarih bölgesi ~19 kolon; 24 Tem günü 1 kolon taşacak — özet formülleri için
   V'den sonra elle 1 sütun eklenmeli (site etkilenmez).
10. Google kota günü TR saatiyle sabah ~10:00'da yenilenir; sayaç Ayarlar'ın altında.
