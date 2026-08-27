# CLAUDE.md — Proje Bağlamı

> Bu dosya Claude Code'un projeyi devralması için yazılmıştır. Claude Code her
> oturum başında bu dosyayı otomatik okur. Amaç: kod tabanını, kurallarını ve
> "neden böyle" kararlarını tek yerde toplamak.

---

## 1. Proje nedir?

**e-YMM Karşıt İnceleme — KDV Fatura Listesi Bölme Programı** (`exay.py`).

YMM'lerin (Yeminli Mali Müşavir) **karşıt inceleme** sürecinde kullandığı bir
masaüstü araçtır. Girdi olarak GİB'in / muhasebe programının ürettiği bir
**"İndirilecek KDV Listesi"** (yüzlerce fatura satırı) alır; bu listeyi
**satıcı firma (VKN/TCKN) bazında** gruplar, belirli kurallara göre bir alt
küme seçer ve **her seçilen firma için ayrı bir Excel "hazır tutanak"** dosyası
üretir. Tutanaklar GİB sisteminin karşıt inceleme şablonuna birebir uygundur.

Tek dosyalık bir Python + Tkinter GUI uygulamasıdır. PyInstaller ile tek
`.exe`'ye derlenip son kullanıcıya (mali müşavir personeline) dağıtılır.
Kullanıcılar teknik değildir; hata mesajları Türkçe ve anlaşılır olmalıdır.

---

## 2. İş kuralları (EN KRİTİK BÖLÜM — değiştirmeden önce iki kez düşün)

Bir firma (VKN) şu **seçim kurallarıyla** tutanaklanır:

1. **Aşama 1 — Eşik kuralı:**
   - Firmanın **tek bir faturası ≥ 150.000 ₺** ise → seçilir, *veya*
   - Firmanın **toplam faturası ≥ 450.000 ₺** ise → seçilir.
   - (Bu iki eşik GUI'den ayarlanabilir; varsayılan 150.000 / 450.000.)

2. **Aşama 2 — %80 kuralı:**
   - Aşama 1'de seçilenlerin toplamı, **listenin tamamının %80'ini**
     karşılamıyorsa; kalan firmalar **toplam tutara göre büyükten küçüğe**
     eklenir, ta ki kümülatif tutar %80'e ulaşana kadar.

3. **%80 paydası = LİSTENİN TAMAMI.**
   ⚠️ Geçersiz VKN/TCKN'li satırlar için tutanak üretilemez (firma
   belirlenemez), **ama tutarları %80 hesabının paydasında kalır.** Bunları
   paydadan düşmek kuralın ihlalidir — gerçek kapsam %80'in altına düşer.
   Bu, geçmişte yaşanmış ve düzeltilmiş gerçek bir hatadır; koruyun.

4. Seçilen firmalar dosya isimlerinde **büyükten küçüğe (toplam tutar)**
   sıralanır ve **1'den ardışık** numaralandırılır (atlama olmamalı).

---

## 3. Girdi liste tipleri (üçü de otomatik tanınır)

Program farklı kaynaklardan gelen üç format tipini de tek başına tanır.
Başlık satırı sabit değildir; **anahtar kelimeyle otomatik bulunur**
(`ana_listeyi_oku`). Sütun adları tiplere göre değişir; `sutun_bul` esnek
alt-dize eşleşmesiyle bulur.

1. **Eski GİB tipi** — başlık ~3. satırda, seri sütunu adsız ("Unnamed"),
   "KDV'si" (kesme işaretli), "Alınan Mal ve/veya Hizmetin ... Tutarı".
2. **Yeni GİB tipi** — başlık 0. satırda, gerçek "Alış Faturasının Serisi"
   sütunu, **"KDV si"** (kesme işaretsiz!), "Alış Faturasının KDV Hariç Tutarı".
3. **Muhasebe (191 hesabı) dökümü** — GİB değil; sütunlar "Hesap Kodu, Tarih,
   fatura no, vergi kimlik no, Açıklama, **Borç** (=KDV), matrah".
   `_muhasebe_tipini_esle` bunu standart GİB adlarına çevirir.

Dosya biçimi olarak **`.xlsx`, `.xls`, `.csv` ve `.txt`** desteklenir. CSV/TXT
için kodlama (UTF-8 / cp1254) ve ayraç (`;`, sekme, `,`) otomatik saptanır
(`_csv_okuyucu_hazirla`); okuyucu, `read_excel` ile aynı arayüzde (header/skiprows)
çalışır, böylece başlık bulma ve muhasebe eşleme mantığı değişmeden geçerlidir.

**Yeni bir liste tipi eklerken:** genelde sadece (a) `sutun_bul` arama
terimlerini genişletmek veya (b) `_muhasebe_tipini_esle` benzeri bir eşleme
eklemek yeterlidir. Programın geri kalanı standart sütun adlarıyla çalışır.

---

## 4. Çıktı sütunları (GİB tutanak şablonu — `SABLON_SUTUNLAR`)

Sıra sabittir, değiştirmeyin:
`Faturanın Tarihi | Faturanın Serisi | Faturanın Numarası | Faturanın Tutarı (TL) | K.D.V(TL) | Defter Kayıt Tarihi | Yevmiye Numarası | Ödeme Şekli... | Açıklama | Hatalı Satır Açıklama`

Kritik biçimlendirme kuralları (hepsi geçmiş hataların dersleridir):

- **Tutar ve KDV sütunları GERÇEK SAYI olmalı** (metin değil), format `0.00`.
  Metin olursa hedef sistem küsüratı düşürüyor. `para_deger` ile parse edilir.
- **VKN, fatura no, seri metin (`@`) olmalı** — baştaki sıfırlar korunsun.
- **Açıklama sütunu** = listedeki "Alınan Mal ve/veya Hizmetin Cinsi".
- **Seri**: yalnızca gerçek bir seri sütunu varsa yazılır; fatura numarasıyla
  aynıysa boş bırakılır (`seri_sutunu_bul` + satır-içi güvence).

---

## 5. Yan çıktı dosyaları (hepsi "Hazır Tutanaklar" klasörüne)

- `N) DÖNEM_VKN_ÜNVAN.xlsx` — her firma için tutanak (ana çıktı).
- `VKN_LISTESI_DÖNEM.xlsx` — seçilen firmalar; sütunlar: Sıra No, VKN, Ünvan,
  **Örnek Fatura No** (firmadan bir örnek). Dosya sırasıyla birebir paralel.
- `GECERSIZ_SATIRLAR_DÖNEM.xlsx` — atlanan geçersiz VKN'li satırlar + nedeni.
- `OLUSTURULAMAYANLAR_DÖNEM.xlsx` — dosyası üretilemeyen firmalar + hata nedeni.
- `OZET_RAPOR_DÖNEM.xlsx` — tek sayfalık çalışma özeti: liste toplamı, hedef
  tutar, seçilen/oluşturulan firma sayısı, **gerçek kapsam %**, geçersiz satır
  sayısı ve tutarı. Sadece raporlar; iş kuralını yeniden hesaplamaz.
- `ISLEM_GUNLUGU_DÖNEM.txt` — o çalışmanın tüm ekran günlüğü (denetim izi):
  başlık olarak sürüm, zaman, kaynak dosya ve kullanılan kriterler.
- `N) …_.pdf` — **opsiyonel** okunur PDF kopyası (yalnızca `reportlab` kuruluysa ve
  kullanıcı "PDF üret"i işaretlerse; `firma_pdf_olustur`). Resmî yükleme dosyası
  yine Excel'dir; PDF arşiv/imza kopyasıdır.
- `N) …_.doc/.docx` — **opsiyonel** Word tutanağı: kullanıcının hazır şablonları
  **VKN ile eşleştirilip** ("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN" bloğu)
  yalnızca "Karşıt İncelemeye Konu Fatura" tablosu firmanın faturalarıyla
  güncellenir. Diğer her şey sabit kalır. **`.docx` şablonlar `python-docx` ile
  Word GEREKTİRMEDEN** üretilir (çıktı `.docx`); **`.doc` (eski ikili)** için
  yazma adımı Windows + Word (pywin32 COM) gerektirir. Okuma/eşleştirme her iki
  biçim için de saf Python'dur (`.doc`→olefile, `.docx`→python-docx).
- `WORD_ESLESME_DÖNEM.xlsx` — hangi seçili firmanın şablonu var/yok ve Word
  tutanağının üretilip üretilmediği (Word olmadan da çıkarılır).

---

## 6. Kod haritası (`exay.py`, ~880 satır, tek dosya)

| Fonksiyon | Görev |
|---|---|
| `kaynak_yolu` | PyInstaller `.exe` içinde/dışında logo vb. yol çözümü (`_MEIPASS`). |
| `sutun_bul` | Esnek (alt-dize, küçük harf) sütun adı bulucu. Her tipin bel kemiği. |
| `kdv_sutunu_bul` | "KDV'si / KDV si / KDVsi"yi bulur; matrah/toplam/tevkifat KDV'siyle KARIŞMAZ. |
| `seri_sutunu_bul` | Gerçek seri sütununu bulur; numara sütununu seri sanmaz. |
| `ana_listeyi_oku` | Dosyayı okur, başlık satırını otomatik bulur, muhasebe eşlemesini uygular. |
| `_muhasebe_tipini_esle` | 191 hesabı dökümünü standart GİB sütun adlarına çevirir. |
| `para_deger` | **Doğru** sayı ayrıştırıcı ("1.234.567,89" → 1234567.89). Toplamlarda bunu kullan. |
| `para_oku` | ESKİ ayrıştırıcı — binlik ayraçta 0 döner. Yeni kodda KULLANMA. |
| `donem_bul` | Dönemi bulur: ay adı → sayısal ay+yıl → veri tarihleri → bugün. Ay adı eşleşmesi Türkçe karakter duyarsızdır (NISAN = NİSAN). |
| `ozet_rapor_olustur` | Çalışmanın tek sayfalık kapsam özetini üretir (openpyxl Workbook + gerçek kapsam % döndürür). |
| `kriter_dogrula` | Eşik/yüzde girdilerini mantıklı aralıkta mı diye denetler (GUI hatalı girişi engeller). |
| `bulunan_sutunlar` | İşlem öncesi önizleme: kritik alanların hangi başlıklara eşlendiğini döndürür. |
| `kdv_tutarlilik_kontrol` | KDV/matrah oranı makul KDV oranlarından uzaksa yanlış sütun eşleşmesine karşı uyarır (yalnızca uyarı). |
| `mukerrer_fatura_bul` | Aynı (VKN, fatura no) birden çok satırda mı diye bakar (yalnızca uyarı). |
| `_ay_yil` / `donem_disi_tarih_kontrol` | Tarihleri tek-anlamlı ayrıştırır; dönem dışı fatura oranını verir (yanlış dönem dosyası uyarısı). |
| `_csv_okuyucu_hazirla` | CSV/TXT için kodlama+ayraç saptar; read_excel ile aynı arayüzde okuyucu döndürür. |
| `firma_pdf_olustur` / `pdf_destekli` / `_pdf_font_bul` | Opsiyonel PDF kopya (reportlab varsa; Türkçe için Unicode TTF kaydeder). |
| `_doc_metni_oku` | Eski ikili `.doc`'un ana metnini çıkarır (olefile; WordDocument akışı UTF-16LE, 0x07→tab). Yalnızca okuma. |
| `sablon_vkn_metinden` / `_blok_vkn` / `sablon_vkn_oku` | Karşı firmanın (vkn, unvan) bilgisini iki belge tipinden de çıkarır: **karşıt inceleme tutanağı** ("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN") ve **YMM Bilgi İsteme yazısı** ("Hakkında Bilgi İstenilen Mükellef…"). VKN'yi 'V.D.' çevresinden ayıklar (karışık etiket/telefonla karışmaz). |
| `_vkn_metinden_ayikla` | Metinden 10-11 haneli VKN/TCKN (boşlukları temizler, 8-9→zfill, yer tutucu geçersiz) — filtreyle aynı normalize. |
| `sablonlari_indeksle` | Klasördeki `.doc`/`.docx` şablonları VKN→(yol, blok) indeksler. **Çok-firmalı tek `.docx`** (bir dosyada N tutanak) tanınır: her firma bloğu ayrı indekslenir. |
| `_docx_firma_bloklari` / `_docx_blok_belgesi` / `_sablon_kayitlari` | Birleşik `.docx`'i firma bloklarına ayırır (blok başı = "KATMA DEĞER…TUTANAĞI" başlığı), tek bloğu izole eder, dosyadaki tüm (vkn, unvan, blok) kayıtlarını verir. |
| `_docx_govde_ekle` / `firmalar_tek_docx` | Doldurulmuş firma docx'lerini tek dosyada (her firma yeni sayfada) birleştirir. |
| `firma_docx_olustur` / `docx_destekli` | `.docx` şablonu python-docx ile açıp fatura tablosunu doldurur, yeni `.docx` yazar (**Word gerektirmez**). |
| `firma_word_olustur` / `word_destekli` | Eski `.doc` şablonu Word (COM) ile açıp fatura tablosunu günceller (Windows + Word). |
| `firma_word_uret` / `sablon_uretilebilir_mi` | Uzantıya göre doğru üreticiyi seçer (.docx→python-docx, .doc→COM); ön koşulu denetler. `inceleme_dayanagi` geçirir. |
| `_docx_inceleme_dayanagi_yaz` | Tutanaktaki "İNCELEME DAYANAĞI" (sözleşme) değer hücresini günceller — eski şablonun eski yılını otomatik ezer. |
| `_ascii_kucuk` | Türkçe-güvenli küçük harf/ASCII fold (İ→i). Anahtar-kelime eşleşmelerinde `.lower()` yerine bunu kullan. |
| `_docx_metni_oku` | `.docx` metnini (paragraf + tablo hücreleri, sekmeli) çıkarır — VKN okuma için. |
| `_word_fatura_satiri` / `_fatura_tablosu_mu` / `_tr_para_str` | Fatura satırını Word tablo sırasına çevirir; fatura tablosunu başlığından tanır; TR para biçimi. |
| `_gecersizlik_nedeni` | Geçersiz VKN için insan-okur neden metni. |
| `firmalari_filtrele` | **KALP.** VKN normalize, geçerli/geçersiz ayrım, 2 aşamalı %80 seçimi. |
| `guvenli_kaydet` | Windows uzun yol (~260) sorununda dosya adını kısaltarak yeniden kaydeder. |
| `firma_excel_olustur` | Tek firmanın tutanak Excel'ini şablona göre yazar. |
| `dosyalari_isle` | Orkestrasyon: oku → **ön bilgi + doğruluk uyarıları** → filtrele → her firma için üret → yan dosyalar + kalıcı günlük. Opsiyonel `ilerleme_cb`, `cikis_kok`, `pdf_uret`, `sablon_klasor`, **`cikti_turu`** ('excel'/'word'/'ikisi'). Ardışık numara yalnızca üretilen firmalar için. |
| `KDVBolmeApp` | Tkinter GUI (sürükle-bırak **çoklu/toplu**, eşik + **doğrulama**, **çıktı türü seçici**, çıktı klasörü, PDF onayı, Word şablon klasörü, ilerleme çubuğu, log, logo, ayarları hatırlama). |

Akış: `dosyalari_isle` → `ana_listeyi_oku` → `firmalari_filtrele` →
(her firma) `firma_excel_olustur` → `guvenli_kaydet`.

---

## 7. Çalıştırma, bağımlılıklar, derleme

- Python 3, bağımlılıklar: `pandas`, `openpyxl`, `xlrd` (eski `.xls` için),
  `pillow` (logo), `olefile` (`.doc` şablon okuma), `python-docx` (`.docx` şablon
  okuma+yazma). **Opsiyonel:** `reportlab` (PDF kopya), `pywin32` (yalnızca eski
  `.doc` şablonlardan üretim — Windows + Word; `.docx` şablonlar Word'süz üretilir).
  Tkinter standart kütüphanede. Testler için: `pytest`.
- Kullanıcı ayarları (son eşik/yüzde, **çıktı klasörü, PDF tercihi**)
  `~/.exay_ayarlar.json` içinde saklanır (Program Files gibi yazılamayan
  konumlarda sorun çıkmasın diye).
- Sürüm sabiti: `SURUM` (GUI başlığında ve özet/günlükte gösterilir).
- Geliştirme: `python exay.py` (GUI açılır).
- `.exe` derleme: `derle.bat` (PyInstaller `--onefile --windowed`, logoyu
  `logo.ico` olarak gömer). Klasörde `exay.py` + `derle.bat` + `logo.ico`
  yan yana olmalı. Çıktı: `dist/KarsitInceleme.exe`.

---

## 8. Test yöntemi (bu ortamda GUI yok)

**Otomatik test paketi** eklendi: `test_exay.py` + `conftest.py` (pytest).
`conftest.py`, tkinter kurulu değilse başsız çalışabilmek için içe aktarmadan
önce hafif bir `tkinter` stub'ı yerleştirir (CLAUDE.md'nin eski headless
yöntemini otomatikleştirir). Çalıştırma:

```bash
pytest -q        # 43 test: para_deger/tarih, kdv/seri/donem bulma, %80 kuralı,
                 # VKN normalizasyon, üç liste tipi (yeni/eski GİB + muhasebe),
                 # CSV okuma, kriter doğrulama, doğruluk uyarıları (kdv/mükerrer/
                 # dönem-dışı), şablon çıktı, özet, PDF, kalıcı günlük, uçtan uca
```

Gerçek GİB dosyaları depoda olmadığından testler üç liste tipini (yeni GİB,
eski GİB, muhasebe/191) **sentetik** üretip mantığı doğrular. Elde gerçek
dosya varsa manuel doğrulama hâlâ geçerlidir:

```python
import exay
df = exay.ana_listeyi_oku(DOSYA)
sec, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, lambda *a, **k: None)
# GERÇEK kapsam = seçilenlerin tutarı / TÜM listenin tutarı  → %80 olmalı
```

Her değişiklikten sonra **üç liste tipini de** (eski GİB, yeni GİB, muhasebe)
test et. Bilinen gerçek dosyalarda beklenen gerçek kapsamlar:
Nisan %94.2, Ocak %82.2, Muhasebe %82.4.

---

## 9. Bu kod tabanında değişmez ilkeler

1. **İş kurallarını (§2) sessizce değiştirme.** Kural değişikliği kullanıcıya
   danışılmalı; kod içi "iyileştirme" gibi geçiştirilmemeli.
2. **Geçmiş hataların düzeltmelerini geri alma:** %80 paydası = tüm liste;
   tutar/KDV sayısal; VKN metin; seri≠numara; ardışık numaralandırma;
   `para_deger` kullanımı; yer tutucu (tek-rakam) VKN'lerin geçersizliği.
3. **Türkçe kal.** UI metinleri, loglar, hata mesajları, dosya adları Türkçe.
4. **Son kullanıcı teknik değil.** Hatalar sessiz kalmamalı; net Türkçe
   açıklama + yan rapor dosyası üretilmeli.
5. **Tek dosya sadeliği.** Şimdilik `exay.py` tek dosya; bölmeden önce gerçek
   ihtiyaç olduğundan emin ol.

---

## 10. Bilinen sınırlamalar / olası sonraki işler

- Çok derin ağ yollarında dosya adı kısaltma devreye girer (bilgi kaybı değil,
  yalnızca ad kısalır) — `guvenli_kaydet`.
- Geçersiz kimlikli satırlar tutanaklanamaz; kullanıcı kaynak listede
  düzeltirse kapsam iyileşir (program uyarıyor).
- ~~GUI'de ilerleme çubuğu yok~~ → **eklendi** (firma sayısına göre dolar).
- ~~Otomatik test paketi yok~~ → **eklendi** (`pytest`, `test_exay.py`, 43 test).
- ~~İşlem öncesi önizleme/uyarı yok~~ → **eklendi** (ÖN BİLGİ bloğu + KDV
  tutarlılık, mükerrer fatura, dönem-dışı tarih uyarıları — hepsi yalnızca
  uyarır, seçimi/iş kuralını etkilemez).
- ~~İşlem günlüğü kalıcı değil~~ → **eklendi** (`ISLEM_GUNLUGU_DÖNEM.txt`).
- ~~Eşik değeri sınır kontrolü yok~~ → **eklendi** (`kriter_dogrula`).
- ~~Çıktı klasörü seçilemiyor~~ → **eklendi** (`cikis_kok` + GUI seçici).
- ~~Toplu (batch) işleme yok~~ → **eklendi** (çoklu dosya seç / sürükle-bırak).
- ~~CSV girdi yok~~ → **eklendi** (`.csv`/`.txt`, kodlama+ayraç otomatik).
- ~~PDF çıktı yok~~ → **eklendi** (opsiyonel, `reportlab` varsa).
- İlerleme çubuğu adım granülaritesi firma başınadır; tek bir firmanın çok
  büyük olması hâlinde ara ilerleme gösterilmez (yeterince ince).
- Doğruluk kontrolleri (KDV oranı, mükerrer, dönem) **uyarı** niteliğindedir;
  satır silmez / seçimi değiştirmez — kullanıcı kaynakta düzeltir.
- **Çıktı türü seçilebilir:** yalnız Excel / yalnız Word / ikisi. Word modları
  şablon klasörü ister. Şablonda tek satır olsa da firmanın tüm faturaları
  yazılır (veri satırları temizlenip her fatura için satır eklenir).
- **İnceleme Dayanağı (sözleşme):** GUI'den girilirse her Word tutanağının
  "İNCELEME DAYANAĞI" hücresi bununla ezilir — gözden kaçan eski yıl şablonları
  bile güncel sözleşmeyle çıkar. Boşsa şablondaki yazı aynen kalır.
- **Çok-firmalı tek `.docx` şablon:** Bir dosyada birçok firmanın tutanağı
  toplanmışsa (her blok "KATMA DEĞER…TUTANAĞI" başlığıyla), program dosyayı
  bloklara ayırıp her firmayı VKN ile ayrı indeksler; üretirken ilgili firmanın
  bloğunu izole edip fatura tablosunu doldurur. **Birleşik `.doc`** (eski ikili)
  ise indeks aşamasında Word (COM) ile bir kez `.docx`'e çevrilir (`_doc_docx_cevir`)
  ve bloklar oradan okunur; böylece üretim tümüyle test edilmiş `.docx` yolundan
  gider. (Tekli `.doc` COM ile yerinde düzenlenir; birleşik `.doc` dönüştürme adımı
  Windows + Word gerektirir ve kullanıcı makinesinde doğrulanmalıdır.)
- **Word'leri tek dosyada birleştir (`word_tek_dosya`):** opsiyonel; üretilen
  `.docx` tutanaklar tek dosyada (her firma yeni sayfada) toplanır
  (`KARSIT_INCELEME_TUTANAKLAR_DÖNEM.docx`; `firmalar_tek_docx`).
- **Boş/yedek şablon (`bos_sablon`):** eşleşmeyen (şablonu olmayan) firmalar için
  kullanıcı bir boş `.docx` şablonu verir; fatura tablosu doldurulur ve bilinen
  ünvan/VKN NEZDİNDE bloğuna yazılır (`_docx_nezdinde_yaz`); kalan firma bilgisini
  kullanıcı girer. Özet/eşleşme raporunda "Boş şablon oluşturuldu" olarak işaretlenir.
- GUI ayarları artık **sekmeli** (Kriterler / Çıktı / Word Şablon) — büyüyen
  seçenekler kalabalık yapmasın diye. Türkçe casing için `_ascii_kucuk` kullanılır.
- PDF, GİB şablonuyla birebir değil; **okunur/arşiv** kopyasıdır (resmî dosya
  Excel). Türkçe için bir Unicode TTF (DejaVuSans/Arial) gerekir; yoksa
  Helvetica'ya düşer ve bazı Türkçe karakterler bozulabilir.
- **Word tutanak eşleştirme (yeni):** `.doc` şablon **okuma/VKN eşleştirme** saf
  Python'dur ve gerçek 5 şablonda test edildi. Ama **yazma** adımı (fatura
  tablosunu güncelleme) Windows'ta Word'ü COM ile sürer ve bu ortamda
  çalıştırılıp doğrulanamadı — kullanıcı makinesinde test edilip tablonun sütun
  sayısı/başlık satırı sayısına göre ince ayar gerekebilir (`firma_word_olustur`
  loglu yazılmıştır). Miktar sütunu kaynak listeden alınır (`sutun_bul(['miktar'])`);
  liste düzeni netleştikçe eşleme gözden geçirilmeli.
