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
   **Tek istisna — TOPLAM satırı fatura değildir:** listeye eklenmiş VKN'siz
   "GENEL TOPLAM / ARA TOPLAM / NAKLİ YEKÜN" satırı paydaya eklenirse liste
   toplamı ikiye katlanır. `toplam_satirlarini_ayikla` böyle bir satırı ANCAK
   (VKN geçersiz) + (metinde toplam/yekün/total **ya da** tarih, fatura no ve
   ünvan hücrelerinin HEPSİ boş) + (tutarı kendinden önceki faturaların
   toplamına 1 ₺ içinde EŞİT) ise ayıklar ve günlüğe yazar. Bu koşulları
   gevşetmeyin — gevşerse gerçek geçersiz faturalar paydadan düşer.
   **Gerçek vaka (Ağustos 2026):** GİB "İndirilecek KDV listesi yeni formatı"
   dosyasının EN ALTINDA **etiketsiz** bir toplam satırı vardır (yalnız tutar
   sütunları dolu). Eski sürüm bunu geçersiz fatura sayıp payda 163 M yerine
   326 M oldu; %80 tutmadı ve 39 yerine 257 firmanın HEPSİNE tutanak üretti.
   **İthalat istisnası (kullanıcı kararı, Ekim 2026):** "GGB Tescil No'su (Alış
   İthalat İse)" hücresi DOLU satır ithalattır → %80 hesabına HİÇ girmez (ne
   seçilir ne paydada kalır), "geçersiz VKN" de sayılmaz (`ithalat_satirlarini_ayir`;
   `firmalari_filtrele` ve `ozet_rapor_olustur` içinde uygulanır). Ağustos 2026:
   RHEINZINK (yer tutucu VKN 1111111111, 6,1 M) payda dışı → 157,2 M, 39 firma, %95,8.
   GGB sütunu olmayan listelerde (eski GİB, muhasebe) hiçbir satır ithalat sayılmaz.

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

Okuma sağlamlığı (hepsi test edildi): başlıklar **Türkçe büyük harfle** de gelebilir
('TARİH', 'KDV HARİÇ TUTARI', 'AÇIKLAMA', 'BORÇ') — `str.lower()` 'İ'yi 'i̇' yaptığı
için tüm başlık karşılaştırmaları `_ascii_kucuk` ile katlanır. Başlık satırı,
anahtar geçen satırlar arasından **en çok rakamsız metin hücresi** olan satırdır
(`_baslik_satiri_bul`; üstteki 'Vergi Kimlik No: …' unvan satırı ya da açıklaması
anahtar içeren veri satırı başlık sanılmaz). Çok sayfalı Excel'de başlığı tanınan
sayfa okunur, diğer dolu sayfalar günlükte uyarılır. CSV: kodlama tüm dosyadan
(UTF-16/Excel 'Unicode Metin' dahil), ayraç satır tutarlılığından saptanır; başlık
üstü unvan/boş satırlar ve satır sonu ayraçları sorun çıkarmaz.

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
- `01 FİRMA VE MUH. BİLGİLERİ <AY> <YIL>.xlsx` — **ayın firma takip dosyası**
  (kullanıcının her ay elle hazırladığı dosyanın aynısı): `SR | FİRMA | VKN | KDV |
  TÜR | BELGE ID | AÇIKLAMA | SMMM | YMM | TELEFONU | ADRESİ | DURUM`. SR = tutanak
  numarası; KDV = firmanın KDV toplamı; AÇIKLAMA/SMMM/YMM/TELEFONU/ADRESİ geçmiş
  takip dosyalarından TAŞINIR (her alan için en yeni dolu değer); BELGE ID ve DURUM o
  aya özgüdür, boş başlar. Geçmişte olmayan firma SARI. İthalat satıcıları en altta
  `TÜR=İTHALAT`. TÜR sütununda açılır liste (EXCEL/KİT/YMM/İTHALAT). Her modda üretilir.

---

## 6. Kod haritası (`exay.py`, ~880 satır, tek dosya)

| Fonksiyon | Görev |
|---|---|
| `kaynak_yolu` | PyInstaller `.exe` içinde/dışında logo vb. yol çözümü (`_MEIPASS`). |
| `_ascii_kucuk` | Türkçe-güvenli küçük harf/ASCII fold (İ→i). Dosyanın başında (YARDIMCI) tanımlıdır; tüm başlık/anahtar kelime karşılaştırmaları bunu kullanır. |
| `sutun_bul` | Esnek (alt-dize, küçük harf, **Türkçe büyük harf güvenli**) sütun adı bulucu. Her tipin bel kemiği. Arama terimleri artık TEK YERDE, modül düzeyi `ARA_*` sabitlerinde (`ARA_VKN/ARA_TARIH/ARA_FATNO/ARA_MATRAH/ARA_KDVYEDEK/ARA_UNVAN/ARA_CINS/ARA_MIKTAR`) — tüm çağrı yerleri bunları kullanır, yeni başlık varyasyonu tek satırla eklenir. Terimler DAR ve sondaki 'ı'sız ("Tutar"=" Tutarı"); spekülatif terim EKLENMEZ. |
| `kdv_sutunu_bul` | "KDV'si / KDV si / KDVsi"yi bulur; matrah/toplam/tevkifat KDV'siyle KARIŞMAZ. 'tutarı' YASAK DEĞİL (matrah zaten 'hariç' ile dışlanır) → geçerli "KDV Tutarı" adlı sütun da bulunur. |
| `seri_sutunu_bul` | Gerçek seri sütununu bulur; numara sütununu seri sanmaz. |
| `ana_listeyi_oku` | Dosyayı okur, sayfayı/başlık satırını otomatik bulur, muhasebe eşlemesini uygular, TOPLAM satırlarını ayıklar. Ek bilgiyi `df.attrs`'a koyar (`sayfa`, `diger_sayfalar`, `toplam_satirlari`) — `dosyalari_isle` günlüğe yazar. Boş listede net Türkçe `ValueError`. |
| `_baslik_satiri_bul` | Başlık satırı seçimi (anahtar geçen satırlar arasında en çok rakamsız metin hücresi). |
| `toplam_satirlarini_ayikla` | VKN'siz TOPLAM satırını (tutarı üstündekilerin toplamına eşitse) ayıklar — §2.3 istisnası. |
| `_muhasebe_tipini_esle` | 191 hesabı dökümünü standart GİB sütun adlarına çevirir. |
| `para_deger` | **Doğru** sayı ayrıştırıcı ("1.234.567,89" → 1234567.89). Toplamlarda bunu kullan. Ondalıksız binlik ("1.234.567", "1,234,567"), bölünmez boşluk, muhasebe eksisi ("1.234,56-", "(1.234,56)") da tanınır. Tek noktalı "12.345" BELİRSİZDİR (yuvarlanmamış KDV olabilir) → ondalık kabul edilir; değiştirmeyin. |
| `_tarih_coz` / `tarih_fmt` | Tek tarih ayrıştırıcı (ISO, GG.AA.YYYY, GG/AA/YYYY, GG-AA-YYYY, tek haneli gün/ay, saatli); tutanağa GİB biçimi GG.AA.YYYY yazılır. `_ay_yil` de bunu kullanır. |
| `para_oku` | ESKİ ayrıştırıcı — binlik ayraçta 0 döner. Yeni kodda KULLANMA. |
| `donem_bul` | Dönemi bulur: ay adı → sayısal ay+yıl → veri tarihleri → bugün. Ay adı eşleşmesi Türkçe karakter duyarsızdır (NISAN = NİSAN). Yıl yalnızca tek başına duran 20xx'tir (addaki VKN içinden alınmaz). Adda ay var yıl yoksa: yıl verideki o aydan; veri yoksa gelecekteki ay olamayacağından GEÇEN yıl (ARALIK listesi OCAK'ta işlenince). |
| `ozet_rapor_olustur` | Çalışmanın tek sayfalık kapsam özetini üretir (openpyxl Workbook + gerçek kapsam % döndürür). |
| `kriter_tutari_oku` | GUI limit kutusunu okur: '150000', '150.000', '150,000', '150.000,00' → 150000 (eskiden '150,000' → 150 TL). |
| `kriter_dogrula` | Eşik/yüzde girdilerini mantıklı aralıkta mı diye denetler (GUI hatalı girişi engeller). |
| `bulunan_sutunlar` | İşlem öncesi önizleme: kritik alanların hangi başlıklara eşlendiğini döndürür. |
| `kdv_tutarlilik_kontrol` | KDV/matrah oranı makul KDV oranlarından uzaksa yanlış sütun eşleşmesine karşı uyarır (yalnızca uyarı). |
| `mukerrer_fatura_bul` | Aynı (VKN, fatura no) birden çok satırda mı diye bakar (yalnızca uyarı). |
| `bos_fatura_no_kontrol` / `vkn_unvan_tutarsizligi` / `negatif_tutar_kontrol` / `ayristirilamayan_tarih_kontrol` | Veri kalitesi kontrolleri (hepsi **yalnızca uyarı**, `_vkn_std`/`_vkn_gecerli_mi` ile filtrelemeyle aynı VKN'yi görür): boş fatura no, aynı VKN'de farklı ünvan, negatif (iade/düzeltme) tutar, ayrıştırılamayan tarih. |
| `_ay_yil` / `donem_disi_tarih_kontrol` | Tarihleri tek-anlamlı ayrıştırır; dönem dışı fatura oranını verir (yanlış dönem dosyası uyarısı). |
| `_csv_okuyucu_hazirla` | CSV/TXT için kodlama+ayraç saptar; read_excel ile aynı arayüzde okuyucu döndürür. |
| `firma_pdf_olustur` / `pdf_destekli` / `_pdf_font_bul` | Opsiyonel PDF kopya (reportlab varsa; Türkçe için Unicode TTF kaydeder). |
| `_doc_metni_oku` | Eski ikili `.doc`'un ana metnini çıkarır (olefile; WordDocument akışı UTF-16LE, 0x07→tab). Yalnızca okuma. |
| `sablon_vkn_metinden` / `_blok_vkn` / `sablon_vkn_oku` | Karşı firmanın (vkn, unvan) bilgisini iki belge tipinden de çıkarır: **karşıt inceleme tutanağı** ("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN") ve **YMM Bilgi İsteme yazısı** ("Hakkında Bilgi İstenilen Mükellef…"). VKN'yi 3 stratejiyle ayıklar: (1) etiketin yanındaki numara ('V.D. – 6120050961'), (2) 'Vergi Dairesi …Nosu' etiket hücresinden sonraki DEĞER hücresi (etikette 'Hesap' gibi ek kelime olsa da), (3) hücre-bazlı son çare (telefon/faks hücreleri atlanır). `.doc`'ta tüm blok TEK satır olabildiğinden atlama hücre bazlıdır (satır bazlı değil). Karışık etiket/telefonla karışmaz. |
| `_vkn_metinden_ayikla` | Metinden 10-11 haneli VKN/TCKN (boşlukları temizler, 8-9→zfill, yer tutucu geçersiz) — filtreyle aynı normalize. |
| `sablonlari_indeksle` | Klasördeki `.doc`/`.docx` şablonları VKN→(yol, blok) indeksler. **Çok-firmalı tek `.docx`** (bir dosyada N tutanak) tanınır: her firma bloğu ayrı indekslenir. Uzantı harf duyarsız; **`Hazır Tutanaklar*` çıktı klasörleri atlanır** (önceki doldurulmuş tutanaklar şablon sanılmasın); VKN'si okunamayan dosyalar günlükte listelenir. |
| `_docx_firma_bloklari` / `_docx_blok_belgesi` / `_sablon_kayitlari` | Birleşik `.docx`'i firma bloklarına ayırır (blok başı = "KATMA DEĞER…TUTANAĞI" başlığı), tek bloğu izole eder, dosyadaki tüm (vkn, unvan, blok) kayıtlarını verir. |
| `_docx_govde_ekle` / `firmalar_tek_docx` | Doldurulmuş firma docx'lerini tek dosyada (her firma yeni sayfada) birleştirir. Kopyalanan gövdedeki resim/dış bağlantı rId'leri hedef belgeye taşınır (farklı şablonların logosu/imzası bozuk ya da yanlış çıkmasın). |
| `firma_docx_olustur` / `docx_destekli` | `.docx` şablonu python-docx ile açıp fatura tablosunu doldurur, yeni `.docx` yazar (**Word gerektirmez**). |
| `firma_word_olustur` / `word_destekli` | Eski `.doc` şablonu Word (COM) ile açıp fatura tablosunu günceller (Windows + Word). Konumsal doldurur (aşağıya bakın). Artık yalnız YEDEK yol: `.doc` şablonlar indekslenirken `.docx`'e çevrilir. |
| `_doclari_docx_cevir` / `_doc_docx_cevir` / `_com_hazirla` | `.doc` şablonları TEK Word örneğiyle `.docx`'e çevirir (`~/.exay_onbellek`, yol+mtime+boyut anahtarlı önbellek; şablon değişmedikçe yeniden çevrilmez). Böylece `.doc` şablonlar da test edilmiş `.docx` yolundan (devam sayfası, tarih, kalın olmayan satır, tarih sırası) üretilir. Word yoksa `{}` + uyarı. `_com_hazirla`: GUI işi ayrı thread'de çalıştığından her Word çağrısından önce `pythoncom.CoInitialize()` (yoksa aynı oturumda 2. çalıştırmada "CoInitialize has not been called"). |
| `devam_sayfasi_hazirla` / `_devam_aylarini_yaz` | KİT devam sayfası şablonunu (.doc→docx çevrilerek) açar; 'AY / YYYY' başlıklarını dönemin bir önceki ayı + dönem ayı yapar (01.2027 → 'ARALIK / 2026' \| 'OCAK / 2027'). |
| `_docx_bolum_olarak_ekle` / `_normal_bicimini_sabitle` / `_temel_bicim` / `_docx_kopya` | Bir belgeyi YENİ BÖLÜM olarak ekler (dikey KİT + yatay devam ayrı korunur). Stiller `docxcompose` ile (ada göre) taşınır — yoksa yalnız gövde. Kaynağın 'Normal' yazı tipi/boyut/aralığı hedefinkinden farklıysa doğrudan biçime çevrilir (gerçek hata: devam KİT'in Cambria'sıyla taşıp boş sayfa oluşturuyordu). Her firmaya devamın taze kopyası verilir. |
| `_docx_yazi_tarihi_yaz` | YMM Bilgi İsteme yazısında yalnız 'Konu … Bilgi İsteme … TARİH' satırındaki tarihi ÇIKTI GÜNÜ yapar (`??.??.????` yer tutucusu dahil). Sayıyı kullanıcı yazar; İnceleme Dayanağı/fatura tarihleri değişmez; KİT'e dokunmaz. |
| `firma_word_uret` / `sablon_uretilebilir_mi` | Uzantıya göre doğru üreticiyi seçer (.docx→python-docx, .doc→COM); ön koşulu denetler. `inceleme_dayanagi` geçirir. |
| `_docx_inceleme_dayanagi_yaz` | Tutanaktaki "İNCELEME DAYANAĞI" (sözleşme) değer hücresini günceller — eski şablonun eski yılını otomatik ezer. |
| `_docx_metni_oku` | `.docx` metnini (paragraf + tablo hücreleri, sekmeli) çıkarır — VKN okuma için. |
| `_word_fatura_satiri` / `_fatura_tablosu_mu` / `_tr_para_str` | Fatura satırını Word tablo sırasına çevirir (KONUMSAL yedek); fatura tablosunu başlığından tanır (tutanak *ve* YMM yazısı); TR para biçimi. |
| `_fatura_deger_haritasi` / `_fatura_son_sutun_dahil` | Fatura doldurmanın TEK yöntemi: **KONUMSAL**. Sütun sırası tüm gerçek şablonlarda sabit (`Tarih\|No\|Cins\|Miktar\|Matrah\|KDV\|son`); `_fatura_deger_haritasi` rol→değer üretir, satır konumsal yazılır. SON sütun `_fatura_son_sutun_dahil` ile tipe göre: tutanak 'Defter Kayıt' **boş**, YMM 'KDV dahil toplam' **matrah+kdv**. **Hem `.docx` hem `.doc`/COM yolu aynı mantığı** kullanır — başlığa göre rol tahmini KULLANILMAZ (kırılgandı; cins boş kalıyordu). |
| `_gecersizlik_nedeni` | Geçersiz VKN için insan-okur neden metni. |
| `firmalari_filtrele` | **KALP.** VKN normalize, geçerli/geçersiz ayrım, 2 aşamalı %80 seçimi. |
| `_metin_hucre` / `_df_excel_kaydet` | Excel'e METİN yazar: '=' ile başlayan metin (ör. açıklama "=KDV iadesi") formül olup dosyayı bozmasın. Tutanak ve tüm yan raporlar bunu kullanır; geçersiz satır raporu da artık `guvenli_kaydet`'ten geçer. |
| `_bos_klasor_adi` | Var olmayan klasör adı (`_2`, `_3`…). Eski "Hazır Tutanaklar" klasörü (içinde yalnız Word/PDF olsa da) taşınır; taşınamazsa (içindeki dosya açık) yeni çıktılar damgalı ayrı klasöre yazılır. |
| `guvenli_kaydet` / `_dosya_kilitli_mesaji` | Windows uzun yol (~260) sorununda dosya adını kısaltarak yeniden kaydeder. Çıktı dosyası Excel/Word'de AÇIKSA (PermissionError) net Türkçe mesajla yükseltir — farklı adla sessizce kaydetmez. `.docx` karşılığı `_guvenli_docx_kaydet`; yan raporlar da bu yoldan kaydedilir. |
| `firma_excel_olustur` | Tek firmanın tutanak Excel'ini şablona göre yazar. |
| `dosyalari_isle` | Orkestrasyon: oku → **ön bilgi + doğruluk uyarıları** → filtrele → takip bilgisi + belge türü → her firma için üret → yan dosyalar + takip dosyası + kalıcı günlük. Opsiyonel `ilerleme_cb`, `cikis_kok`, `pdf_uret`, `sablon_klasor`, **`cikti_turu`** ('excel'/'word'/'ikisi'/**'firmaya_gore'**), `takip_klasor`, `bos_sablon` (KİT), `bos_ymm_sablon`, `devam_sablon`. Ardışık numara yalnızca üretilen firmalar için. |
| `takip_dosyalarini_bul` / `takip_dosyasi_oku` / `takip_gecmisi_oku` / `takip_bilgisi_esle` | Geçmiş takip dosyalarını ("…FİRMA…BİLGİ….xls[x]") liste klasörü + bir üstü + takip klasöründe (3 düzey, AppData vb. atlanır) bulur; başlık satırını ('FİRMA' + SMMM/YMM/…) kendisi bulur; eskiden yeniye (addaki dönem, `_ad_donemi`) okur; firmayı VKN'den, yoksa ünvandan (`_unvan_anahtari`: Türkçe-katlanmış, A.Ş./LTD./SAN./TİC. atılmış; `_unvan_eslesir`: eşit ya da ≥10 karakterlik baş — kırpılmış ünvanlar) eşler. Gerçek Ağustos 2026: 39 firmadan 38'i eşleşti (kalan 1 takipte yok). |
| `_tur_cikar` / `tur_normalize` | TÜR sütunu olmayan eski takip dosyasında tür çıkarımı: **YMM dolu → YMM; BELGE ID dolu ya da açıklamada 'sistem' → EXCEL; diğer → KİT** (gerçek Ağustos dosyasında kullanıcının 24 Excel / 3 KİT çıktısıyla birebir). Açık TÜR sütunu her zaman önceliklidir. |
| `firma_takip_dosyasi_yaz` / `takip_dosyasi_adi` | Ayın takip dosyasını yazar (bkz. §5). |
| `KDVBolmeApp` | Tkinter GUI (sürükle-bırak **çoklu/toplu**, eşik + **doğrulama**, **çıktı türü seçici** (Excel/Word/İkisi/Firmaya göre), çıktı klasörü, **takip klasörü**, PDF onayı, Word şablon klasörü, boş KİT/YMM şablonu, KİT devam sayfası, ilerleme çubuğu, log, logo, ayarları hatırlama). |

Akış: `dosyalari_isle` → `ana_listeyi_oku` → `firmalari_filtrele` →
(her firma) `firma_excel_olustur` → `guvenli_kaydet`.

---

## 7. Çalıştırma, bağımlılıklar, derleme

- Python 3, bağımlılıklar: `pandas`, `openpyxl`, `xlrd` (eski `.xls` için),
  `pillow` (logo), `olefile` (`.doc` şablon okuma), `python-docx` (`.docx` şablon
  okuma+yazma), `docxcompose` (devam sayfası/tek dosya birleştirmede stil taşıma;
  yoksa birleştirme yine çalışır, yalnız eksik stiller Normal'e düşer). **Opsiyonel:** `reportlab` (PDF kopya), `pywin32` (yalnızca eski
  `.doc` şablonlardan üretim — Windows + Word; `.docx` şablonlar Word'süz üretilir).
  Tkinter standart kütüphanede. Testler için: `pytest`.
- **Sürümler `requirements.txt`'te ARALIKLA sabit** (`>=test edilmiş, <sonraki ana
  sürüm`): `pip install -r requirements.txt` ileride sürpriz bir büyük güncellemeyle
  programı bozmaz, ama yama/küçük güncellemeleri alır. Üst sınırı yükseltmeden önce
  `pytest -q` çalıştır. (Pillow LANCZOS kodu hem eski hem yeni Pillow'u destekler.)
- Kullanıcı ayarları (son eşik/yüzde, **çıktı klasörü, PDF tercihi**)
  `~/.exay_ayarlar.json` içinde saklanır (Program Files gibi yazılamayan
  konumlarda sorun çıkmasın diye).
- Sürüm sabiti: `SURUM` (GUI başlığında ve özet/günlükte gösterilir).
- Geliştirme: `python exay.py` (GUI açılır).
- **GitHub Actions (`.github/workflows/exe.yml`)**: her push'ta Windows'ta testler +
  PyInstaller `.exe` + **öz-test**: `KarsitInceleme.exe --oz-test <klasör>` (`_oz_test`)
  paketlenmiş hâliyle sentetik liste/şablon/devam/takip ile "Firmaya göre" çalıştırır
  (eksik paket verisi derlemede yakalansın). main'de exe "son-surum" yayınına yüklenir.
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
pytest -q        # 181 test: para_deger/tarih, kdv/seri/donem bulma, %80 kuralı,
                 # VKN normalizasyon, üç liste tipi (yeni/eski GİB + muhasebe),
                 # CSV okuma, kriter doğrulama, doğruluk uyarıları (kdv/mükerrer/
                 # dönem-dışı), şablon çıktı, özet, PDF, kalıcı günlük, uçtan uca,
                 # Türkçe büyük harf başlık, TOPLAM satırı, CSV/UTF-16 kenar durumları
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
Nisan %94.2, Ocak %82.2, Muhasebe %82.4, **Ağustos 2026 (yeni GİB formatı,
2 sayfalı .xls, etiketsiz toplam satırlı) → 656 fatura, 39 firma, %95.8**
(ithalat istisnasından önce %92.2 idi).
Bu dosyada program çıktısı kullanıcının Excel tutanaklarıyla ve BARSA Word
tutanağının fatura tablosuyla birebir aynı çıktı. Aynı ayın **YMM 08.2026.doc**
birleşik dosyası (14 Bilgi İsteme yazısı, .docx'e çevrilerek) şablon klasörü
yapıldığında seçilen 39 firmadan 13'ü eşleşti ve 13 yazının fatura tablosu
kullanıcının elle hazırladığıyla birebir aynı çıktı. Gerçek müşteri dosyaları
(VKN/ünvan içerir) DEPOYA KONMAZ; testler bunların düzenini sentetik taklit eder
(`_gib_yeni_bicim_yaz`).

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
- ~~Otomatik test paketi yok~~ → **eklendi** (`pytest`, `test_exay.py`, 181 test).
- ~~İşlem öncesi önizleme/uyarı yok~~ → **eklendi** (ÖN BİLGİ bloğu + KDV
  tutarlılık, mükerrer fatura, dönem-dışı tarih uyarıları — hepsi yalnızca
  uyarır, seçimi/iş kuralını etkilemez).
- ~~İşlem günlüğü kalıcı değil~~ → **eklendi** (`ISLEM_GUNLUGU_DÖNEM.txt`).
- ~~Eşik değeri sınır kontrolü yok~~ → **eklendi** (`kriter_dogrula`).
- ~~Çıktı klasörü seçilemiyor~~ → **eklendi** (`cikis_kok` + GUI seçici).
- ~~Toplu (batch) işleme yok~~ → **eklendi** (çoklu dosya seç / sürükle-bırak).
- ~~CSV girdi yok~~ → **eklendi** (`.csv`/`.txt`, kodlama+ayraç otomatik).
- ~~PDF çıktı yok~~ → **eklendi** (opsiyonel, `reportlab` varsa).
- **Negatif tutar / iade faturası:** Kullanıcının listelerinde hiç yoktur (iade
  faturaları listeye eklenmez). Bu yüzden 2. aşamada sıfır/eksi firma durumu için
  ek kural gerekmedi (soru kapandı, kural değiştirilmedi).
- Toplu işlemde aynı klasördeki her liste bir öncekinin "Hazır Tutanaklar"
  klasörünü zaman damgalı ada taşır (veri kaybı yok; son liste düz adlı klasörde).
- İlerleme çubuğu adım granülaritesi firma başınadır; tek bir firmanın çok
  büyük olması hâlinde ara ilerleme gösterilmez (yeterince ince).
- Doğruluk/veri kalitesi kontrolleri (KDV oranı, mükerrer, dönem, boş fatura no,
  VKN-ünvan tutarsızlığı, negatif tutar, ayrıştırılamayan tarih) **uyarı** niteliğindedir;
  satır silmez / seçimi değiştirmez — kullanıcı kaynakta düzeltir.
- **Çıktı türü seçilebilir:** yalnız Excel / yalnız Word / ikisi / **firmaya göre**.
  Word modları şablon klasörü ister. **Firmaya göre** (`firmaya_gore`): her firmaya
  takip dosyasındaki TÜR'e göre YALNIZ bir belge — EXCEL → Excel tutanağı; KİT → KİT
  (VKN'li KİT şablonu, yoksa boş KİT şablonu) + devam sayfası; YMM → YMM yazısı (YMM
  şablonu, yoksa boş YMM şablonu). Aynı firmanın hem KİT hem YMM şablonu olabilir
  (`_SablonIndeksi.hepsi`), türe uyan seçilir. O türde şablon yoksa resmî **Excel**
  üretilir ve günlükte listelenir; Word üretimi hata verirse de Excel'e düşer. TÜR
  takipte yoksa: firmanın şablon türü, o da yoksa EXCEL. Gerçek Ağustos 2026 (KİT +
  YMM birleşik şablonlarla): 24 Excel + 13 YMM + 2 KİT — kullanıcının elle ürettiği
  dağılımla aynı. Şablonda tek satır olsa da firmanın tüm faturaları
  yazılır (veri satırları temizlenip her fatura için satır eklenir).
- **İnceleme Dayanağı (sözleşme):** GUI'den girilirse her Word tutanağının
  "İNCELEME DAYANAĞI" hücresi bununla ezilir — gözden kaçan eski yıl şablonları
  bile güncel sözleşmeyle çıkar. Boşsa şablondaki yazı aynen kalır.
- **Word dosya adı (Excel'den farklı):** Word tutanakları/yazıları
  `N) [YMM ]İLK ÜÇ KELİME AA-YYYY.docx` biçiminde adlandırılır (`_word_tutanak_adi`):
  firmanın ünvanının ilk üç kelimesi + dönem (`AA-YYYY`, tireli). YMM (Bilgi
  İsteme) yazısında başına `YMM ` gelir (`_sablon_ymm_mi` ile tip tespiti). Boş
  şablon çıktısında sona ` BOŞ` eklenir. **Excel adı değişmedi** (resmî GİB
  dosyası hâlâ `N) DÖNEM_VKN_ÜNVAN.xlsx`). Ardışık `N)` numarası korunur.
- **İki belge tipi (tutanak + YMM yazısı):** Şablonlar iki türdür — (a) *karşıt
  inceleme tutanağı* ("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN", son sütun
  'Defter Kayıt') ve (b) *YMM Bilgi İsteme yazısı* ("Hakkında Bilgi İstenilen
  Mükellef", son sütun 'KDV dahil toplam'). VKN okuma (`sablon_vkn_metinden` +
  `_blok_vkn`) ve fatura doldurma (konumsal + `_fatura_son_sutun_dahil`) ikisini
  de tanır; blok bölme (`_docx_firma_bloklari`, `_metni_bloklara_ayir`) her iki
  başlığı da blok başı sayar, karşı-taraf tablosunu her iki başlıktan bulur.
  **Gerçek YMM yazısında fatura tablosu 6 sütundur** (F.TARİHİ | F. NOSU | MALIN
  CİNSİ | MALIN MİKTARI | MATRAH | KDV; 'KDV dahil' sütunu yok) — 6 ve 7 sütun
  geçerli sayılır (`_FATURA_SUTUN_GECERLI`), yanlış uyarı verilmez. Birleşik
  dosyada yazılar ~30 boş satırla ayrıldığından izole edilen bloğun sonundaki boş
  paragraflar silinir (`_sondaki_bos_paragraflari_sil`; yoksa her çıktıda boş 2.
  sayfa oluşuyordu). Word'süz bilgisayarda birleşik `.doc` "bozuk" değil
  "Word gerekli" diye raporlanır (`_birlesik_doc_mu`).
- **Çok-firmalı tek `.docx` şablon:** Bir dosyada birçok firmanın tutanağı/yazısı
  toplanmışsa (her blok "KATMA DEĞER…TUTANAĞI" ya da "Konu: Bilgi İsteme"
  başlığıyla), program dosyayı bloklara ayırıp her firmayı VKN ile ayrı indeksler;
  üretirken ilgili firmanın bloğunu izole edip fatura tablosunu doldurur. **Birleşik `.doc`** (eski ikili)
  ise indeks aşamasında Word (COM) ile bir kez `.docx`'e çevrilir (`_doc_docx_cevir`)
  ve bloklar oradan okunur; böylece üretim tümüyle test edilmiş `.docx` yolundan
  gider. Artık **tekli `.doc` da** aynı şekilde (önbellekli) `.docx`'e çevrilip
  üretilir; COM ile yerinde düzenleme yalnız çevirme başarısız olursa yedektir.
  (Dönüştürme Windows + Word gerektirir; kullanıcı makinesinde doğrulanmalıdır.)
- **KİT devam sayfası (`devam_sablon`, GUI: "KİT devam sayfası…"):** gerçek KİT 3
  sayfadır; exay 1. sayfayı doldurur, kullanıcı her ay 2 yatay devam sayfasının
  yalnız ay başlıklarını değiştirip her KİT'in arkasına koyuyordu. Şimdi devam
  şablonu bir kez seçilir; her KİT'in arkasına ayrı (yatay) bölüm olarak, ay
  başlıkları döneme göre güncellenerek eklenir. YMM yazısına eklenmez. Gerçek
  dosyalarla (BARSA KİT + devam) LibreOffice'te 3 sayfa (1 dikey + 2 yatay) doğrulandı.
- **KİT Word 2027'de kalkıyor (kullanıcı bilgisi, Ekim 2026):** yalnız Excel tutanak +
  YMM yazısı kalacak. Kuralın çıktı tarihine mi liste dönemine mi göre işleyeceği
  HENÜZ BELLİ DEĞİL → kod DEĞİŞTİRİLMEDİ; o zamana kadar kullanıcı takip dosyasında
  TÜR'ü elle EXCEL yapar. Kesinleşince yapılacak: `dosyalari_isle`'de `turler`
  hesaplanırken TÜR=KİT → EXCEL (günlüğe not) ve `_tur_cikar` varsayılanı EXCEL;
  KİT devam sayfası / boş KİT şablonu seçenekleri kullanılmaz hâle gelir.
- **YMM yazısı tarihi:** 'Konu: Bilgi İsteme' satırındaki tarih çıktının alındığı
  gün olur; Sayı'yı kullanıcı giden evrak defterinden yazar.
- **Word'leri tek dosyada birleştir (`word_tek_dosya`):** opsiyonel; üretilen
  `.docx` tutanaklar tek dosyada (her firma yeni sayfada) toplanır
  (`KARSIT_INCELEME_TUTANAKLAR_DÖNEM.docx`; `firmalar_tek_docx`).
- **Boş/yedek şablon (`bos_sablon`):** eşleşmeyen (şablonu olmayan) firmalar için
  kullanıcı bir boş `.docx` şablonu verir; fatura tablosu doldurulur ve bilinen
  ünvan/VKN NEZDİNDE bloğuna yazılır (`_docx_nezdinde_yaz`); kalan firma bilgisini
  kullanıcı girer. Özet/eşleşme raporunda "Boş şablon oluşturuldu" olarak işaretlenir.
- GUI **"sakin bordo" tek pencere dikey akış** (sekmeler kaldırıldı): üst şerit →
  bırak alanı → yan yana **Kriterler** + **Çıktı** kartları → tam genişlik
  **TUTANAKLARI OLUŞTUR** butonu → **özet metrik kartları** (üretilen/seçilen,
  gerçek kapsam %, tutanak sayısı) + uyarı şeridi → **açılır işlem günlüğü**
  (varsayılan KAPALI, `_log_ac_kapa`). Tek aksan rengi `ANA_RENK="#5E1030"`;
  gövde nötr (`APP`/`KART`/`KENAR`). Kapsam % artık **slider** (`_kapsam_kaydir`,
  `self.yuzde80` StringVar korunur). Çıktı türü **segmented** (`_segment_sec`);
  Word Şablon ayarları Çıktı kartında Word/İkisi seçiliyken açılır (`_word_blok_guncelle`).
  **Akış:** dosya sürükle/seç önce SAKLANIR (`_dosya_sec`), işlem OLUŞTUR ile başlar.
  Metrikler `dosyalari_isle` sonunda `tamam_cb`'ye 4. argüman (özet sözlüğü) ile
  gelir; 3-argümanlı eski geri çağrılar TypeError'a düşüp uyumlu kalır.
  Türkçe casing için `_ascii_kucuk` kullanılır.
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
- **Fatura satırları KALIN yazılmaz** (elle hazırlanan tutanaklarda hiç kalın fatura
  bilgisi yoktur). Gerçek hata (OPUROĞLU GOLD 08-2026.doc): COM yolu tüm veri
  satırlarını silip `Rows.Add()` ile ekliyordu; Word yeni satıra son kalan KALIN
  başlık satırının biçimini kopyaladığı için fatura satırları kalın çıkıyordu
  (başlığı kalın olmayan şablonda çıkmadığından "değişken" görünüyordu). Şimdi
  şablonun İLK veri satırı örnek olarak kalır, yeniler onun biçimini alır ve her
  veri satırına `Range.Font.Bold = False` uygulanır; `.docx` yolunda da veri
  satırı run'ları `bold=False`. COM yolu artık testte **sahte Word nesne
  modeliyle** (`_SahteTablo`: `Rows.Add()` son satırın biçimini kopyalar) sınanır.
- **Word'de faturalar TARİH sırasıyla** dizilir (`_tarihe_gore_sirala`; hem `.docx`
  hem COM yolu): kaynak liste tutara göre sıralı gelse bile (gerçek vaka: OPUROĞLU
  GOLD 08-2026) tablo elle hazırlananlar gibi nizami olur. Sıralama KARARLIDIR (aynı
  tarihli faturalar listedeki sırasını korur — böylece Ağustos 2026'nın 13 YMM
  yazısı ve BARSA tutanağı elle hazırlananlarla hâlâ birebir aynı); tarihi
  okunamayan sona gider. **Excel tutanağının sırası değişmez** (liste sırası).
