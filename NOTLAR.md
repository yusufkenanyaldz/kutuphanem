# NOTLAR.md — Kullanıcıyla çalışma geçmişi, kararlar ve bağlam

> Bu dosya, kullanıcıyla yapılan uzun sohbetin (Eylül–Ekim 2026) kritik bilgilerini
> kalıcı tutmak içindir: sohbet kayıtları ve oturum dosyaları ~30 gün sonra silinir,
> bu dosya depoda kalır. **Yeni bir oturumda önce CLAUDE.md'yi, sonra bu dosyayı oku.**
> CLAUDE.md "kod nasıl çalışır" der; bu dosya "kullanıcı ne istiyor, ne karar verildi,
> ne yarım kaldı" der.
>
> ⚠️ **GİZLİLİK — DEPO HERKESE AÇIK (public).** Bu dosyaya ve depoya müşteri VKN'si,
> firma ünvanı, kişi adı, telefon, e-posta, adres, GGB no, fatura no/tutarı YAZILMAZ.
> Gerçek dosyalar yalnızca kullanıcıdadır; doğrulama gerekirse kullanıcıdan yeniden
> istenir (bkz. §9). Testlerde yalnızca sentetik değerler kullanılır (FIRMA A,
> 1000000001, ALFA ÇELİK…).

---

## 0. Yeni oturumda ilk 5 dakika

1. Dal: `claude/exay-py-development-5a2ofu` (main'e **birleştirilmedi**, PR açılmadı —
   kullanıcı denemesini bekliyoruz). `git log origin/main..HEAD` ile bu dalın işleri görülür.
2. `pip install -r requirements.txt && pytest -q` → 181 test geçmeli.
3. GitHub Actions "Windows .exe" her push'ta: testler + PyInstaller + `--oz-test`.
   Son başarılı çalışma: run 37055503601 (2026-10-02). Exe, çalışmanın "Artifacts"
   bölümünde (`KarsitInceleme`, ~65 MB zip). "Releases → son-surum" yalnızca **main'e**
   push'ta güncellenir.
4. Açık işler: §8. Kullanıcıya sorulmuş ve cevap bekleyen sorular: §8.
5. Kullanıcıyla **Türkçe**, sade, teknik olmayan dille konuş. Önce sonuç, sonra ayrıntı.

---

## 1. Kullanıcı ve çalışma ortamı

- Kullanıcı bir **YMM (Yeminli Mali Müşavir) bürosunda** çalışıyor (Gaziantep).
  İş: **KDV iadesi karşıt incelemesi** — iade talep eden müşteri firmanın
  "İndirilecek KDV Listesi"ndeki satıcı firmalar için tutanak/yazı hazırlamak.
- Teknik değil; programı kendisi ve büro personeli kullanıyor. Hata mesajları Türkçe
  ve anlaşılır olmalı.
- Büroda **Python kurulu olmayan bilgisayarlar var** → `.exe` şart (GitHub Actions
  derliyor, §4).
- Kullanıcı Claude'u telefon/masaüstü uygulamasından kullanıyor; dosyaları ofisten
  (zip olarak) gönderiyor. **Cowork**'ü (masaüstü, klasörlere erişebilen Claude) fatura
  → Word şablonu döngüsü için kurdu (§6).
- GitHub'da güncel kodu "main"de arıyor; dal/PR kavramına aşina değil. Güncel
  `exay.py` gerekirse dosya olarak da gönderilir (SendUserFile).

---

## 2. Kullanıcının aylık iş akışı (gerçek dosyalardan öğrenildi)

1. **Girdi:** GİB'den "İndirilecek KDV listesi yeni formatı" (`.xls`).
   - 2 sayfa: "İndirilecek KDV Listesi" (asıl liste) + "Tutanak Çalışması" (kullanıcının
     çalışma sayfası — okunmaz, günlükte uyarılır).
   - En altta **etiketsiz toplam satırı** (yalnız tutar sütunları dolu) → ayıklanır.
   - "GGB Tescil No'su (Alış İthalat İse)" sütunu: doluysa ithalat (§3).
   - Ağustos 2026: ~656 fatura satırı, ~259 satıcı firma.
2. **Seçim (exay):** tek fatura ≥150.000 ₺ veya toplam ≥450.000 ₺; yetmezse büyükten
   küçüğe %80'e kadar. Ağustos'ta 39 firma.
3. **Her seçilen firmaya TEK belge türü gider** (firmaya göre değişir):
   - **EXCEL** → e-YMM sistemine yüklenen resmî Excel tutanağı ("sistemden gönderin"
     diyen firmalar). Yüklenince takip dosyasında **BELGE ID** dolar.
   - **KİT** (Karşıt İnceleme Tutanağı, Word) → kargo/elden/mail. Gerçek KİT **3 sayfa**:
     1. sayfa tutanak (exay doldurur) + **2 yatay "devam sayfası"** (KDV beyanname
     bilgileri tablosu: "<önceki ay> / YIL | <dönem ayı> / YIL", alt firmalar tablosu,
     ödeme/işçi/oda bilgileri, ortaklar, imza). Kullanıcı eskiden her ay devam
     sayfasında yalnızca ayları değiştirip her KİT'in arkasına koyuyordu → artık exay yapıyor.
   - **YMM yazısı** ("Bilgi İsteme", Word) → satıcı firmanın kendi YMM'si varsa o YMM'ye
     yazılır, kargoyla gider. "Sayı: YMM <sicil>/2026-<sıra>" kullanıcı tarafından
     **giden evrak tablosundan** verilir; tarih = çıktının alındığı gün (exay basar);
     hitap "Sayın <YMM adı> / Yeminli Mali Müşavirlik / <il>" kullanıcı yazar. Fatura
     tablosu **6 sütun**: F.TARİHİ | F. NOSU | MALIN CİNSİ | MALIN MİKTARI | MATRAH | KDV.
4. **Birleşik Word dosyaları:** kullanıcı bir ayın tüm KİT'lerini tek dosyada tutuyor
   ("KİT <AA>.<YYYY>.doc", Ağustos'ta 3 firma), YMM yazılarını da ("YMM <AA>.<YYYY>.doc",
   Ağustos'ta 14 yazı; aralarında ~30 boş satır). Ayrıca "KİT DEVAM SAYFASI.doc".
   Hepsi eski ikili `.doc` → exay Word ile `.docx`'e çevirir (önbellekli).
5. **Takip dosyası** — "01 FİRMA VE MUH. BİLGİLERİ <AY> <YIL>.xls":
   - Düzen: 1. satır boş, A sütunu boş, başlık 2. satırda, en altta KDV toplamı.
   - Sütunlar: SR | FİRMA (sonda boşluklu "FİRMA ") | KDV | BELGE ID | AÇIKLAMA |
     SMMM | YMM | TELEFONU | ADRESİ | DURUM. **VKN sütunu yok** (exay ünvandan eşler).
   - SR = tutanak numarası, sıra KDV/tutar büyükten küçüğe; ithalat satıcısı en altta.
   - AÇIKLAMA örnekleri: e-posta adresi, "FİRMA TUTANAK", "SİSTEMDEN GÖNDERİLECEK",
     "ELDEN YAPILACAK.", "FİRMA BÜNYESİNDE". SMMM sütununa bazen not yazılıyor
     ("firmadan … iletecek").
   - DURUM örnekleri: "<tarih>'DA KARGOYA VERİLDİ.", "<tarih>'DA <personel>'A VERİLDİ.",
     "<tarih>'DA MAİL ATILDI."
   - SMMM/YMM bilgileri **yıl başında** değişebilir, ay içinde değişmez. Kullanıcı her
     ay bunları geçen aydan kopyalıyordu (en büyük zaman kaybı) → artık exay taşıyor.
6. **Excel tutanak adları** (resmî): `N) AA_YYYY_VKN_ÜNVAN.xlsx` — kullanıcının kendi
   çıktılarıyla birebir aynı. **Word adları**: `N) [YMM ]İLK ÜÇ KELİME AA-YYYY.docx`.

---

## 3. Kullanıcı kararları (karar günlüğü — değiştirmeden önce kullanıcıya sor)

| Tarih | Konu | Karar |
|---|---|---|
| Eyl 2026 | Seçim kuralı | **Sabit:** tek ≥150.000, toplam ≥450.000; %80 karşılanmazsa büyükten küçüğe %80'e kadar. |
| Eki 2026 | İthalat | **%80 hesabına katılmaz** (ne seçilir ne paydada kalır, geçersiz VKN sayılmaz). GGB hücresi doluysa ithalat. |
| Eki 2026 | Eksi tutar / iade faturası | Listelerinde **hiç yok** (iade faturaları listeye eklenmez). Soru kapandı; ek kural yok. |
| Eki 2026 | Fatura satırları | Word'de **kalın değil** (elle hazırlananların hiçbirinde kalın yok). |
| Eki 2026 | Word'de sıra | Faturalar **tarih sırasıyla** ("daha nizami olur"). Excel tutanağı liste sırasında kalır. |
| Eki 2026 | YMM yazısı | **Sayı'yı kullanıcı verir** (giden evrak). Program yalnızca **tarihi** = çıktı günü yapar. |
| Eki 2026 | KİT devam sayfası | Her KİT'in arkasına eklenir, yalnız ay başlıkları değişir (önceki ay + dönem ayı). |
| Eki 2026 | Firma başına tür | "YMM'si olanların hepsi Word (YMM yazısı)"; "sistemden gönderin" diyenler Excel; diğerleri KİT. Tek tek VKN yazmak yerine takip dosyasında **TÜR** sütunu (kolay yol). |
| Eki 2026 | Takip dosyası | Geçmiş aydan SMMM/YMM/telefon/adres/açıklama **otomatik taşınsın**. |
| Eki 2026 | .exe | GitHub Actions ile derlensin ("muazzam olur"). |
| Eki 2026 | Cowork şablonları | Ünvan **kısaltılmış** kalsın; **faks yazılmasın**; yalnız firma bilgileri faturadan; fatura tablosu ve sözleşmeyi exay, YMM adı/Sayı'yı kullanıcı yazar. |
| Eki 2026 | **KİT 2027'de kalkıyor** | 2027'de yalnız **Excel + YMM yazısı** kalacak. Uygulama tarihi (çıktı tarihi mi, liste dönemi mi) **HENÜZ BELLİ DEĞİL** → kod değiştirilmedi; o zamana kadar kullanıcı TÜR'ü elle EXCEL yapar. Kesinleşince: `dosyalari_isle`'de TÜR=KİT → EXCEL + günlük notu, `_tur_cikar` varsayılanı EXCEL. |
| Eki 2026 | Gizlilik | Gerçek müşteri dosyaları depoya konmaz (depo public). |

---

## 4. Bu dalda yapılanlar (commit sırasıyla, 2026-09-29 → 2026-10-02)

1. **Büyüteç turu (v5.2):** Türkçe büyük harf başlıklar ('TARİH','KDV HARİÇ TUTARI');
   TOPLAM satırı çift sayımı; `para_deger` (binlik, NBSP, sondaki eksi, parantez);
   `donem_bul` (yıl regex, ARALIK listesi OCAK'ta); CSV kodlama/ayraç; `kriter_tutari_oku`
   ('150,000' → 150 TL hatası); çıktı/klasör sağlamlığı.
2. **Etiketsiz toplam satırı** tüm firmaları seçtiriyordu (gerçek Ağustos listesi:
   payda 163 M yerine 326 M, 39 yerine 257 firma). Düzeltildi (`toplam_satirlarini_ayikla`:
   kimlik hücreleri boş + tutar = üstündekilerin toplamı).
3. **YMM yazıları:** izole edilen blok sonundaki boş paragraflar → boş 2. sayfa
   düzeltildi; 6 sütunlu tablo için yanlış "7 sütun" uyarısı kaldırıldı; Word'süz
   bilgisayarda birleşik `.doc` "Word gerekli" diye raporlanır.
4. **Kalın fatura satırları** (COM yolu): Word, `Rows.Add()` ile kalın başlık
   biçimini kopyalıyordu → ilk veri satırı örnek tutulur + `Font.Bold=False`.
5. **Tarih sırası** (Word): `_tarihe_gore_sirala` (kararlı; okunamayan tarih sona).
6. **GitHub Actions** (`.github/workflows/exe.yml`): windows-latest, Python 3.11,
   testler (`EXAY_TK_STUB=1`), PyInstaller `--onefile --windowed --name KarsitInceleme
   --collect-data docx --collect-all docxcompose --collect-all tkinterdnd2
   --hidden-import win32com.client`, artifact; main'de `son-surum` yayını.
7. **İthalat %80 dışı** (`ithalat_satirlarini_ayir`). Ağustos: kapsam %92,2 → %95,8.
8. **KİT devam sayfası + YMM tarihi + .doc→.docx çevirme + COM CoInitialize**
   (`devam_sayfasi_hazirla`, `_docx_bolum_olarak_ekle`, `_normal_bicimini_sabitle`,
   `_docx_yazi_tarihi_yaz`, `_doclari_docx_cevir`, `_com_hazirla`).
9. **Takip dosyası + "Firmaya göre" modu** (`takip_*`, `tur_normalize`, `_tur_cikar`,
   `firma_takip_dosyasi_yaz`, `_SablonIndeksi.hepsi`; GUI: "Firmaya göre", "Takip
   klasörü…", "Boş YMM şablonu…"; "Boş şablon…" → "Boş KİT şablonu…").
10. **.exe öz-testi** (`--oz-test`, `_oz_test`): CI'da paketlenmiş exe tutanak üretiyor mu.
11. **Gizlilik temizliği:** testlerdeki gerçek firma adı/VKN/kişi adı/telefon/GGB/fatura
    tutarları sentetik değerlerle değiştirildi (davranış aynı, 181 test).

Test sayısı: 96 (oturum başı) → 181.

---

## 5. Gerçek veriyle doğrulama sonuçları (Ağustos 2026 dosyaları; dosyalar depoda YOK)

- Liste: 656 fatura; etiketsiz toplam satırı ayıklandı; 1 ithalat satıcısı (~6,1 M,
  yer tutucu VKN) payda dışı → payda ~157,2 M; **39 firma, gerçek kapsam %95,8**;
  1. aşama yeterli (2. aşama gerekmedi).
- Excel tutanakları kullanıcının elle ürettikleriyle **birebir aynı** (ad + içerik).
- YMM birleşik dosyası şablon yapılınca 13 firma eşleşti; 13 yazının fatura tablosu
  elle hazırlananla birebir. KİT birleşik dosyasından 2 firma eşleşti.
- **Firmaya göre** modu: **24 Excel + 13 YMM + 2 KİT** — kullanıcının o ayki gerçek
  dağılımıyla aynı (kullanıcının Excel klasöründe tam 24 dosya vardı; takipte BELGE ID
  dolu satır sayısı da 24).
- Takip eşleşmesi: seçilen 39 firmadan **38'i** geçmiş takip dosyasından eşleşti
  (kalan 1 firma kullanıcının takibinde hiç yoktu).
- KİT + devam: LibreOffice'te **3 sayfa** (1 dikey + 2 yatay); 3 KİT'lik birleşik
  dosya + devam → 5 sayfa; tek dosyada birleştirme bölümleri korur.
- YMM yazısı: tarih çıktı günü (02.10.2026) basıldı; çok faturalı firmada 2 sayfa (normal).
- **Kullanıcıya bildirilen gözlem (cevap bekleniyor):** Ağustos takip dosyasında
  eşiklerin ALTINDA kalan 3 firma vardı (elle eklenmiş olabilir; program seçmez) ve
  tek faturası 156.500 ₺ olan (kurala göre seçilmesi gereken) 1 firma takipte yoktu.
  Kullanıcı "elle firma ekleme" isterse bu yeni bir özellik olur — sorulmadan kural
  değiştirilmez.

---

## 6. Cowork kurulumu (fatura → Word şablonu döngüsü)

- **Amaç:** Yeni bir firmanın şablonu yoksa kullanıcı faturayı (PDF/foto) Cowork'e
  verir; Claude yalnızca **firma bilgilerini** boş şablona yazar ve "Word Şablonları"
  klasörüne kaydeder. Sonra exay bu şablonu VKN ile eşleyip fatura tablosunu ve
  İnceleme Dayanağı'nı doldurur. YMM adı / Sayı'yı kullanıcı yazar.
- **Neden Cowork:** masaüstü klasörlerine erişebiliyor, döngü işi (çok fatura) için uygun.
- **Klasör düzeni** (bağlı klasör "Karşıt İnceleme"): `Boş Şablonlar\` (BOŞ KİT
  ŞABLONU.docx, BOŞ YMM ŞABLONU.docx — asla değiştirilmez), `Faturalar\KİT\`,
  `Faturalar\YMM\`, `Faturalar\İşlendi\`, `Word Şablonları\`.
- **Boş şablonlar** (kullanıcıda var; depoda YOK çünkü YMM'nin ve müşterinin bilgisi
  içerir). Yapıları:
  - BOŞ KİT: başlık "KATMA DEĞER VERGİSİ İADESİ KARŞIT İNCELEME TUTANAĞI"; 18×2 bilgi
    tablosu (YEMİNLİ MALİ MÜŞAVİRİN: Adı Soyadı, Bağlı Olduğu Oda, Mühür ve Sicil No,
    Vergi Dairesi ve Sicil No, İşyeri Adresi, Telefon/Fax · İADE TALEBİNDE BULUNAN
    FİRMANIN: Ünvanı, Vergi Dairesi/Nosu, Adresi, Telefon/Fax · NEZDİNDE KARŞIT
    İNCELEME YAPILAN FİRMANIN: Ünvanı, Vergi Dairesi/Nosu, Adresi, Telefon/Fax ·
    İNCELEME DAYANAĞI); defter onay tablosu 4×4 (YEVMİYE/ENVANTER/DEFTER-İ KEBİR);
    fatura tablosu 3×7 (2 başlık satırı: FATURANIN Tarihi/Numarası · MALIN
    Cinsi/Miktarı/Tutarı/KDV Tutarı · Defter Kayıt Tarihi/Nosu).
  - BOŞ YMM: "Sayı: YMM <sicil>/2026-????", "Konu: Bilgi İsteme   ??.??.????",
    "Sayın ??", 20×3 bilgi tablosu (YMM + iade talep eden + HAKKINDA BİLGİ İSTENİLEN
    MÜKELLEFİN: Ünvanı, Vergi Dairesi/Nosu, Adresi, Telefon/Fax + İNCELEME DAYANAĞI),
    "İSTENİLEN BİLGİLER" maddeleri, fatura tablosu 2×6, imza bloğu.
- Tam talimat metni: **Ek A** (temizlenmiş). 2027'de KİT kalkınca talimatın KİT
  kısımları gereksizleşir.
- Kullanıcının soruları ve cevaplar: "Cowork masaüstü dosyalarını görebiliyor mu?" →
  evet, bağlı klasör üzerinden. "Tek PDF'te birden çok fatura olursa?" → talimatta
  §1A (sayfa sayfa incele, satıcı VKN'sine göre grupla, her satıcıya tek şablon).

---

## 7. Teknik notlar ve tuzaklar (tekrar yaşanmasın)

- **Devam sayfası birleştirme:** docxcompose tek başına yatay bölümü kaybeder → önce
  hedefin son bölümü kapatılır (sectPr son paragrafın pPr'ına), sonra
  `Composer.append`, sonra gövde sectPr'ı kaynağınkiyle (üst/alt bilgi bağlantısız)
  değiştirilir. Kaynağın Normal yazı tipi farklıysa (gerçek: KİT Cambria, devam Times)
  devam taşar → `_normal_bicimini_sabitle` doğrudan biçime yazar. LibreOffice'in
  çevirdiği `.docx`'te varsayılan stil işareti (`w:default`) yoktur →
  `_varsayilan_paragraf_stili` 'Normal'e düşer. Her firmaya devamın **taze kopyası**.
- **COM:** GUI işi thread'de; `pythoncom.CoInitialize()` her Word çağrısından önce
  (`_com_hazirla`). `.doc` çevirme yalnızca **sahte Word** ile test edildi → gerçek
  Word'de kullanıcı makinesinde doğrulanmalı.
- **Önbellek:** `~/.exay_onbellek` (yol+mtime+boyut anahtarı); yarım dosya silinir.
- **Testler:** tkinter yoksa `conftest.py` stub koyar; Windows CI'da `EXAY_TK_STUB=1`.
  Takip araması liste klasörünün **bir üstünü** de tarar → testlerde liste
  `tmp_path/2026/04 NİSAN/` gibi izole klasöre konmalı.
- **Ortam kurulumu (görsel kontrol için):** `apt-get install -y libreoffice-writer`,
  `pip install pymupdf docxcompose`; `soffice --headless --convert-to pdf …` sonra
  pymupdf ile PNG'ye çevirip bak. `HOME`'u yazılabilir bir klasöre ayarla.
- **GUI bu ortamda görsel test EDİLEMEDİ** (Tk yok). Yeni 4'lü segment ("Excel /
  Word / İkisi / Firmaya göre") ve yeni düğmelerin Windows'ta sığması kullanıcıdan teyit edilmeli.
- **.exe** ~65 MB (pandas + docx + docxcompose/babel). `--windowed` exe'de konsol
  yok; öz-test sonucu dosyaya yazar (`OZ_TEST_TAMAM.txt` / `OZ_TEST_HATA.txt`).

---

## 8. Açık işler ve bekleyen sorular

- [ ] **Kullanıcı denemesi → PR → main.** Kullanıcı yeni sürümü (özellikle "Firmaya
      göre", takip dosyası, devam sayfası) kendi makinesinde deneyecek. Uygunsa main'e
      PR (Releases'te kalıcı exe bağlantısı için). PR'ı kullanıcı istemeden açma.
- [ ] **KİT 2027:** uygulama tarihi kesinleşince §3'teki değişiklik.
- [ ] **Elle firma ekleme?** (§5 gözlem) — kullanıcı cevap vermedi.
- [ ] **Gerçek Word ile doğrulama:** `.doc`→`.docx` çevirme, COM yedek yolu.
- [ ] **GUI yerleşimi** Windows'ta kontrol.
- [ ] **Depo gizliliği:** main'in ve bu dalın eski commit'lerinde/commit mesajlarında
      gerçek firma adları ve bazı VKN'ler var (önceki oturumlardan da). Kullanıcıya
      seçenekler soruldu: depoyu private yapmak (GitHub Pages ücretsiz planda kapanır)
      ve/veya bu dalın geçmişini temizleyip zorla göndermek. Cevaba göre ilerle.

---

## 9. Yeni oturumda gerçek dosyalarla doğrulama

1. Kullanıcıdan örnek klasörü (liste `.xls`, takip `.xls`, KİT/YMM birleşik `.doc`,
   KİT DEVAM SAYFASI.doc, elle hazırlanmış Excel tutanaklar) **zip olarak yeniden iste**;
   oturum karalama klasörüne aç, **depoya koyma**.
2. `.doc` dosyalarını LibreOffice ile `.docx`'e çevir (Word yok).
3. Hızlı kontrol:
   ```python
   import conftest, exay
   df = exay.ana_listeyi_oku(LISTE)
   sec, gec = exay.firmalari_filtrele(df, 150000, 450000, 80, print)
   exay.dosyalari_isle(LISTE, 150000, 450000, 80, print, lambda *a: None,
       sablon_klasor=SABLONLAR, cikti_turu='firmaya_gore',
       devam_sablon=DEVAM_DOCX, inceleme_dayanagi='…')
   ```
   Beklenen (Ağustos 2026): 39 firma, %95,8; 24 Excel + 13 YMM + 2 KİT; takipte 38/39.
4. Word çıktılarını PDF'e çevirip sayfa sayısı/yönünü kontrol et (KİT = 3 sayfa).

---

## Ek A — Cowork TALİMAT.txt (temizlenmiş tam metin)

Kullanıcıya verilen asıl metin budur; yalnızca örnek VKN/telefon sentetik yapıldı.

```
KARŞIT İNCELEME ŞABLON ASİSTANI — TALİMAT

# GÖREV
Bir YMM bürosu için KDV iadesi karşıt inceleme ŞABLONLARI hazırlıyorsun.
Faturadan yalnızca FİRMA BİLGİLERİNİ doldurursun. Fatura tablosunu ve İnceleme
Dayanağı'nı exay programı; YMM yazısının sayı, tarih ve hitap kısımlarını ben
dolduruyorum. Bunlara DOKUNMA.

# KLASÖR DÜZENİ (bağlı klasör: "Karşıt İnceleme")
Boş Şablonlar\       BOŞ KİT ŞABLONU.docx, BOŞ YMM ŞABLONU.docx (ASLA DEĞİŞTİRME)
Faturalar\KİT\       karşıt inceleme tutanağı yapılacak firmaların faturaları
Faturalar\YMM\       YMM bilgi isteme yazısı yapılacak firmaların faturaları
Faturalar\İşlendi\   işlenen faturalar buraya taşınır
Word Şablonları\     ürettiğin şablonlar buraya kaydedilir

# NE ZAMAN NE YAPACAKSIN
- "Faturaları işle" dediğimde: Faturalar\KİT\ ve Faturalar\YMM\ içindeki bütün
  faturaları (PDF, JPG, PNG, XML) işle. KİT klasöründekiler için
  "BOŞ KİT ŞABLONU.docx", YMM klasöründekiler için "BOŞ YMM ŞABLONU.docx" kullan.
- Sohbete doğrudan fatura atıp "KİT" ya da "YMM" yazarsam aynı işi o fatura için
  yap. Kelime yazmadıysam sor.
- Dosyanın bulunduğu klasör (KİT/YMM) o dosyadaki BÜTÜN satıcılar için geçerlidir.
  Mesajımda belirli bir firma ya da sayfa için farklı tür söylersem
  (ör. "TOPLU.pdf 3. sayfadaki firma YMM") onu uygula.

# 1A. ÇOK FATURALI / ÇOK SAYFALI DOSYALAR
Bir PDF'te birden çok fatura olabilir (toplu tarama), bir fatura da birden çok
sayfa sürebilir.
- PDF'i SAYFA SAYFA incele; taranmış PDF'te (metin yoksa) sayfaları görüntüye
  çevirip oku.
- Yeni bir fatura şu işaretlerle başlar: üstte satıcı bloğu ve "FATURA NO" alanı
  yeniden görünür, fatura numarası değişir ya da "Page 1 of N / Sayfa 1/N"
  yeniden başlar.
- Satıcı bloğu olmayan devam sayfası (ör. "Page 2 of 2": yalnız kalemler ya da
  toplamlar) bir önceki faturaya aittir.
- Faturaları SATICININ VKN'sine göre grupla: her farklı satıcı için bir şablon.
  Aynı satıcının faturaları (aynı PDF'te ya da farklı dosyalarda) TEK şablon olur.
- Okuyamadığın ya da emin olamadığın sayfa için tahmin etme; özette dosya adı ve
  sayfa numarasıyla belirt.

# 1. FATURAYI OKU
Yok say: kaşeler ("KAYIT EDİLDİ" vb.), el yazısı notlar, döviz notları, banka
hesapları, ETTN, QR kod. Yalnızca basılı fatura bilgilerini kullan.

SATICI (faturayı düzenleyen, genelde üst-sol blok) ve ALICI ("SAYIN" ya da ikinci
blok) için:
- Ünvanı: BÜYÜK HARF. Şu kısaltmaları uygula: SANAYİ→SAN., TİCARET→TİC.,
  LİMİTED ŞİRKETİ→LTD. ŞTİ., ANONİM ŞİRKETİ→A.Ş., İNŞAAT→İNŞ.
  Ünvanın geri kalanını değiştirme.
- Vergi dairesi adı ve VKN (10 hane) / TCKN (11 hane, şahıs firması).
- Adresi: faturada yazdığı gibi, BÜYÜK HARF. Yalnızca boşlukları düzelt ve ilçe ile
  il arasına " / " koy. Kısaltmaları açma, kelime ekleme.
- Telefon: yalnızca telefon, "0 342 000 00 00" biçiminde. Faks YAZMA.
Okuyamadığın bilgi için "??" yaz. ASLA tahmin etme.

# 2. VKN / TCKN DOĞRULAMA (zorunlu — exay şablonu firmaya VKN ile bağlar)
Kod çalıştırarak kontrol hanesini doğrula (d = rakamlar, soldan 0'dan başlayarak):
- VKN (10 hane): i = 0..8 için t = (d[i] + 9 - i) % 10.
  t ≠ 0 ise p = (t × 2^(9-i)) % 9 ve p = 0 çıkarsa p = 9; t = 0 ise p = 0.
  toplam = p'lerin toplamı. (10 - toplam % 10) % 10 = d[9] olmalı.
- TCKN (11 hane): d[0] ≠ 0;
  ((d0+d2+d4+d6+d8) × 7 - (d1+d3+d5+d7)) % 10 = d[9];
  (d0+d1+...+d9) % 10 = d[10].
Tutmazsa rakamları faturada yeniden oku. Yine tutmazsa O SATICI için şablon
üretme ve özette "⚠️ VKN DOĞRULANAMADI" diye belirt (aynı dosyadaki diğer
satıcılara devam et).

# 3. ŞABLONU DOLDUR
Boş şablonun KOPYASINI python-docx ile aç ve biçimini bozmadan düzenle; sıfırdan
belge oluşturma. Hücreye yazarken hücredeki ilk run'ın metnini değiştir, diğer
run'ları sil (yazı tipi ve hizalama korunur).

a) KARŞI FİRMA bölümü ← SATICI
   KİT'te başlığı "NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN",
   YMM'de "HAKKINDA BİLGİ İSTENİLEN MÜKELLEFİN".
   Ünvanı | Vergi Dairesi/Nosu | Adresi | Telefon/Fax (yalnız telefon)
   Vergi Dairesi/Nosu: VKN 3-3-4, TCKN 3-3-3-2 gruplu:
     KİT: "GAZİKENT / 123 456 7890"
     YMM: "GAZİKENT V.D. / 123 456 7890"

b) "İADE TALEBİNDE BULUNAN FİRMANIN" ← ALICI
   Alıcının VKN'si şablondakiyle AYNIYSA bu bölüme dokunma.
   FARKLIYSA bu bölümü alıcının bilgileriyle (a'daki biçimde, "V.D." yazmadan)
   doldur ve özette belirt.

c) Başka HİÇBİR yere dokunma: YMM bilgileri, İnceleme Dayanağı ("??" kalır),
   defter onay tablosu, fatura tablosu (boş kalır), YMM yazısındaki "??" ve
   "????" yer tutucuları ve "İstenilen Bilgiler" metni olduğu gibi kalır.

# 4. KAYDET VE TOPARLA
- Aynı satıcının birden çok faturası varsa TEK şablon üret.
- Kaydetmeden önce Word Şablonları\ içindeki mevcut şablonların
  "Vergi Dairesi/Nosu" hücrelerine bak. Aynı VKN'li şablon zaten varsa YENİSİNİ
  ÜRETME, özette "zaten var: <dosya adı>" yaz (exay aynı VKN'den iki şablon
  görürse karışır).
- Dosya adı: KİT → "<ünvanın ilk üç kelimesi> KİT.docx"
             YMM → "YMM <ünvanın ilk üç kelimesi>.docx"
  Word Şablonları\ klasörüne kaydet. Aynı adlı dosya varsa üzerine yazmadan
  bana sor.
- Bir dosyadaki BÜTÜN satıcılar için şablon üretildiyse (ya da "zaten var" ise)
  o dosyayı Faturalar\İşlendi\ klasörüne taşı. Dosyada okunamayan bir sayfa ya da
  doğrulanamayan bir VKN varsa dosyayı TAŞIMA, yerinde bırak; özette hangi
  sayfaların kaldığını yaz. Dosyaları bölme ya da birleştirme.
- Boş Şablonlar\ klasöründeki dosyaları ASLA değiştirme.

# 5. ÖZET
İş bitince tek tablo ver; her SATICI için bir satır:
Fatura dosyası | Sayfa(lar) | Fatura sayısı | Tür (KİT/YMM) | Ünvan | VKN |
Doğrulandı mı | Sonuç (üretildi / zaten var / işlenmedi) |
Uyarılar (?? alanlar, alıcı farkı, okunamayan sayfa vb.)
Tablonun altına: işlenen dosya sayısı, İşlendi'ye taşınan ve yerinde bırakılan
dosyalar.
```

> Not: Cowork'ün ürettiği şablonlar "Word Şablonları\" klasörüne gider; exay'da bu
> klasör "Şablon klasörü…" olarak seçilir. Aynı VKN'den iki şablon olmamalı
> ("Firmaya göre" modunda aynı firmanın bir KİT bir YMM şablonu olabilir — türe uyan seçilir).
