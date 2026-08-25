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
| `_gecersizlik_nedeni` | Geçersiz VKN için insan-okur neden metni. |
| `firmalari_filtrele` | **KALP.** VKN normalize, geçerli/geçersiz ayrım, 2 aşamalı %80 seçimi. |
| `guvenli_kaydet` | Windows uzun yol (~260) sorununda dosya adını kısaltarak yeniden kaydeder. |
| `firma_excel_olustur` | Tek firmanın tutanak Excel'ini şablona göre yazar. |
| `dosyalari_isle` | Orkestrasyon: oku → filtrele → her firma için üret → yan dosyalar. Opsiyonel `ilerleme_cb(tamamlanan, toplam)` ile GUI ilerleme çubuğunu besler. |
| `KDVBolmeApp` | Tkinter GUI (sürükle-bırak, eşik alanları, log kutusu, logo, **ilerleme çubuğu**, son kriterleri hatırlama). |

Akış: `dosyalari_isle` → `ana_listeyi_oku` → `firmalari_filtrele` →
(her firma) `firma_excel_olustur` → `guvenli_kaydet`.

---

## 7. Çalıştırma, bağımlılıklar, derleme

- Python 3, bağımlılıklar: `pandas`, `openpyxl`, `xlrd` (eski `.xls` için),
  `pillow` (logo). Tkinter standart kütüphanede. Testler için: `pytest`.
- Kullanıcı ayarları (son eşik/yüzde değerleri) `~/.exay_ayarlar.json` içinde
  saklanır (Program Files gibi yazılamayan konumlarda sorun çıkmasın diye).
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
pytest -q        # 25 test: para_deger/tarih, kdv/seri/donem bulma, %80 kuralı,
                 # VKN normalizasyon, muhasebe eşleme, şablon çıktı, özet, uçtan uca
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
- ~~Otomatik test paketi yok~~ → **eklendi** (`pytest`, `test_exay.py`).
- İlerleme çubuğu adım granülaritesi firma başınadır; tek bir firmanın çok
  büyük olması hâlinde ara ilerleme gösterilmez (yeterince ince).
