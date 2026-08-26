"""exay.py için otomatik test paketi.

CLAUDE.md §8 uyarınca headless çalışır (conftest.py tkinter'ı stub'lar) ve
§2'deki iş kurallarını (2 aşamalı %80 seçimi, %80 paydası = tüm liste, geçersiz
VKN'lerin geçersizliği, ardışık numaralandırma, sayısal tutar/KDV) doğrular.

Gerçek GİB dosyaları bu depoda bulunmadığından, üç liste tipi (yeni GİB, eski
GİB, muhasebe/191) sentetik olarak üretilir ve mantık bunlar üzerinde test edilir.

Çalıştırma:  pytest -q
"""
import openpyxl
import pandas as pd
import pytest

import exay


# ── Sessiz log geri çağrısı ──────────────────────────────────────────────────
def _sessiz(*a, **k):
    pass


# ══════════════════════════════════════════════════════════════════════════
#  Saf yardımcı fonksiyonlar
# ══════════════════════════════════════════════════════════════════════════
@pytest.mark.parametrize("girdi,beklenen", [
    ("1.234.567,89", 1234567.89),   # TR binlik + ondalık
    ("1,234,567.89", 1234567.89),   # EN binlik + ondalık
    ("45927,50", 45927.5),          # TR ondalık
    ("1234", 1234.0),
    (1000, 1000.0),
    (1000.5, 1000.5),
    ("", None),
    (None, None),
    ("abc", None),
    ("1500 ₺", 1500.0),            # para birimi işareti temizlenir
    ("1.234,50 TL", 1234.5),       # TL eki + TR biçim
])
def test_para_deger(girdi, beklenen):
    assert exay.para_deger(girdi) == beklenen


def test_tarih_fmt():
    assert exay.tarih_fmt("2026-04-15") == "15.04.2026"
    assert exay.tarih_fmt("") == ""
    # Tanınmayan biçim olduğu gibi döner
    assert exay.tarih_fmt("15/04/2026") == "15/04/2026"


def test_sayi_fmt():
    assert exay.sayi_fmt(1000) == "1000"      # tam sayı → küsürat gösterilmez
    assert exay.sayi_fmt("1234,5") == "1234,50"  # ondalıklı → 2 hane, virgüllü
    assert exay.sayi_fmt("") == ""


def test_kdv_sutunu_bul():
    # Yeni tip (kesme işaretsiz) doğru seçilir, matrah/toplam KDV'siyle karışmaz
    kols = ["Alış Faturasının KDV Hariç Tutarı", "KDV si", "Toplam İndirilecek KDV"]
    assert exay.kdv_sutunu_bul(kols) == "KDV si"
    # Eski tip (kesme işaretli)
    assert exay.kdv_sutunu_bul(["Matrah", "KDV'si"]) == "KDV'si"
    # Yalnızca yasaklı sütunlar varsa None
    assert exay.kdv_sutunu_bul(["KDV Hariç Tutarı", "Toplam KDV"]) is None


def test_seri_sutunu_bul():
    kols = ["Tarih", "Seri", "No"]
    assert exay.seri_sutunu_bul(kols, "Tarih", "No") == "Seri"
    # Ayrı seri sütunu yok; tarihin sağındaki numara sütunudur → seri yok
    kols2 = ["Tarih", "No"]
    assert exay.seri_sutunu_bul(kols2, "Tarih", "No") is None


def test_donem_bul_turkce_ay():
    # Türkçe karakterli ve ASCII (diakritiksiz) ay adlarının ikisi de tanınmalı
    assert exay.donem_bul("NİSAN_2026_liste") == "04.2026"
    assert exay.donem_bul("NISAN_2026_liste") == "04.2026"   # ASCII yazım
    assert exay.donem_bul("AGUSTOS 2025") == "08.2025"
    # Sayısal ay + yıl
    assert exay.donem_bul("04_2026_kdv") == "04.2026"


def test_gecersizlik_nedeni():
    assert "Boş" in exay._gecersizlik_nedeni("")
    assert "kısa" in exay._gecersizlik_nedeni("123")
    assert "uzun" in exay._gecersizlik_nedeni("123456789012")
    assert "Sayısal değil" in exay._gecersizlik_nedeni("ABC123")
    assert "tutucu" in exay._gecersizlik_nedeni("0000000000")


# ══════════════════════════════════════════════════════════════════════════
#  Sentetik liste üreticileri
# ══════════════════════════════════════════════════════════════════════════
def _yeni_gib_yaz(yol):
    """Yeni GİB tipi liste: başlık 0. satırda, gerçek seri sütunu, 'KDV si'."""
    satirlar = [
        # (tarih, seri, no, matrah, kdv, unvan, vkn)
        ("2026-04-01", "A", "BBK1", 200000, 36000, "FIRMA A", "1000000001"),  # tek≥150k
        ("2026-04-02", "A", "BBK2", 300000, 54000, "FIRMA B", "1000000002"),  # B topl 600k
        ("2026-04-03", "A", "BBK3", 300000, 54000, "FIRMA B", "1000000002"),
        ("2026-04-04", "A", "BBK4", 100000, 18000, "FIRMA C", "1000000003"),  # eşik altı
        ("2026-04-05", "A", "BBK5",  50000,  9000, "FIRMA D", "1000000004"),  # eşik altı
        ("2026-04-06", "A", "BBK6",  40000,  7200, "FIRMA E", "0000000000"),  # geçersiz VKN
    ]
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Serisi",
            "Alış Faturasının Sıra No'su", "Alış Faturasının KDV Hariç Tutarı",
            "KDV si", "Satıcının Adı-Soyadı / Ünvanı",
            "Satıcının Vergi Kimlik Numarası"]
    df = pd.DataFrame(satirlar, columns=kols)
    df.to_excel(yol, index=False)


def _muhasebe_yaz(yol):
    """Muhasebe (191) dökümü: Borç=KDV, Matrah=matrah, Satıcının sütunu yok."""
    satirlar = [
        ("191.01", "2026-01-10", "F1", "1000000001", "FIRMA A", 36000, 200000),
        ("191.01", "2026-01-11", "F2", "1000000002", "FIRMA B",  9000,  50000),
    ]
    kols = ["Hesap Kodu", "Tarih", "Fatura No", "Vergi Kimlik No",
            "Açıklama", "Borç", "Matrah"]
    df = pd.DataFrame(satirlar, columns=kols)
    df.to_excel(yol, index=False)


# ══════════════════════════════════════════════════════════════════════════
#  İş kuralı testleri — firmalari_filtrele (KALP)
# ══════════════════════════════════════════════════════════════════════════
def test_filtrele_kapsam_ve_secim(tmp_path):
    yol = tmp_path / "nisan.xlsx"
    _yeni_gib_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))

    secilen, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)

    # A (tek fatura) ve B (toplam) seçilmeli; C ve D seçilmemeli
    assert "1000000001" in secilen           # FIRMA A
    assert "1000000002" in secilen           # FIRMA B
    assert "1000000003" not in secilen        # FIRMA C
    assert "1000000004" not in secilen        # FIRMA D

    # Geçersiz VKN'li satır (0000000000) geçersizler listesinde olmalı
    assert len(gecersiz) == 1

    # §2/§3: %80 paydası = TÜM liste (geçersiz tutar dahil).
    # Beklenen gerçek kapsam = (200000+600000) / 990000 = %80.8
    tutar_col = exay.sutun_bul(list(df.columns),
                               ['kdv hariç tutarı', 'faturanın tutarı'])
    toplam_liste = df[tutar_col].apply(lambda v: exay.para_deger(v) or 0).sum()
    secilen_tutar = sum(
        grp[tutar_col].apply(lambda v: exay.para_deger(v) or 0).sum()
        for grp, _ in secilen.values())
    kapsam = secilen_tutar / toplam_liste * 100
    assert toplam_liste == pytest.approx(990000)
    assert kapsam == pytest.approx(80.8, abs=0.1)


def test_filtrele_buyukten_kucuge_sirali(tmp_path):
    yol = tmp_path / "nisan.xlsx"
    _yeni_gib_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    secilen, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    # Dosya isimlendirme için büyükten küçüğe sıralı olmalı → B(600k) önce, A(200k) sonra
    anahtarlar = list(secilen.keys())
    assert anahtarlar[0] == "1000000002"   # FIRMA B en büyük
    assert anahtarlar[1] == "1000000001"   # FIRMA A


def test_filtrele_vkn_onde_sifir_tamamlama(tmp_path):
    """8-9 haneli VKN'lerde önde eksik sıfır otomatik tamamlanmalı ve geçerli sayılmalı."""
    satirlar = [
        ("2026-04-01", "A", "N1", 500000, 90000, "KISA VKN FIRMA", "71419747"),  # 8 hane
    ]
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Serisi",
            "Alış Faturasının Sıra No'su", "Alış Faturasının KDV Hariç Tutarı",
            "KDV si", "Satıcının Adı-Soyadı / Ünvanı",
            "Satıcının Vergi Kimlik Numarası"]
    yol = tmp_path / "kisa.xlsx"
    pd.DataFrame(satirlar, columns=kols).to_excel(yol, index=False)
    df = exay.ana_listeyi_oku(str(yol))
    secilen, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert "0071419747" in secilen     # önde sıfır tamamlandı
    assert len(gecersiz) == 0


def test_filtrele_bos_liste_cokmez():
    """Tümü sıfır/boş tutarlı liste bölme hatası vermeden çalışmalı (robustluk)."""
    df = pd.DataFrame({
        "Satıcının Vergi Kimlik Numarası": ["1000000001"],
        "Alış Faturasının KDV Hariç Tutarı": [0],
        "Satıcının Adı-Soyadı / Ünvanı": ["SIFIR FIRMA"],
    })
    secilen, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert isinstance(secilen, dict)   # çökmeden döndü


# ══════════════════════════════════════════════════════════════════════════
#  Muhasebe tipi tanıma
# ══════════════════════════════════════════════════════════════════════════
def test_muhasebe_tipi_eslenir(tmp_path):
    yol = tmp_path / "191.xlsx"
    _muhasebe_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    # Standart GİB adlarına çevrilmiş olmalı
    assert exay.sutun_bul(list(df.columns), ['kdv hariç tutarı']) is not None
    assert exay.kdv_sutunu_bul(list(df.columns)) is not None
    secilen, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert "1000000001" in secilen     # tek fatura 200000 ≥ 150000


# ══════════════════════════════════════════════════════════════════════════
#  Tutanak Excel çıktısı — şablon ve sayısal biçim
# ══════════════════════════════════════════════════════════════════════════
def test_firma_excel_sablon_ve_sayisal(tmp_path):
    yol = tmp_path / "nisan.xlsx"
    _yeni_gib_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    secilen, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    grp = secilen["1000000001"][0]     # FIRMA A
    cikti = tmp_path / "firma_a.xlsx"
    exay.firma_excel_olustur(grp, str(cikti), list(df.columns))

    wb = openpyxl.load_workbook(cikti)
    ws = wb.active
    basliklar = [ws.cell(1, c).value for c in range(1, len(exay.SABLON_SUTUNLAR) + 1)]
    assert basliklar == exay.SABLON_SUTUNLAR
    # 4. sütun tutar, 5. sütun KDV → GERÇEK SAYI olmalı (metin değil)
    assert isinstance(ws.cell(2, 4).value, (int, float))
    assert isinstance(ws.cell(2, 5).value, (int, float))
    assert ws.cell(2, 4).value == pytest.approx(200000)


# ══════════════════════════════════════════════════════════════════════════
#  Özet rapor (yeni özellik)
# ══════════════════════════════════════════════════════════════════════════
def test_ozet_rapor(tmp_path):
    yol = tmp_path / "nisan.xlsx"
    _yeni_gib_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    secilen, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    wb, kapsam = exay.ozet_rapor_olustur(
        df, secilen, gecersiz, 150000, 450000, 80, "04.2026",
        basarili=len(secilen), hatali_sayisi=0)
    assert kapsam == pytest.approx(80.8, abs=0.1)
    ws = wb.active
    metin = "\n".join(str(ws.cell(r, 1).value) for r in range(1, ws.max_row + 1))
    assert "GERÇEK KAPSAM" in metin


# ══════════════════════════════════════════════════════════════════════════
#  Uçtan uca — dosyalari_isle tutanakları + yan dosyaları üretir, ilerleme çağrılır
# ══════════════════════════════════════════════════════════════════════════
def test_dosyalari_isle_uctan_uca(tmp_path):
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)

    ilerleme_kayit = []
    sonuc = {}

    def _tamam(klasor, basarili, hatali):
        sonuc['klasor'] = klasor
        sonuc['basarili'] = basarili
        sonuc['hatali'] = hatali

    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, _tamam,
                        ilerleme_cb=lambda t, top: ilerleme_kayit.append((t, top)))

    assert sonuc['basarili'] == 2         # A ve B
    assert sonuc['hatali'] == 0
    cikis = tmp_path / "Hazır Tutanaklar"
    dosyalar = [p.name for p in cikis.glob("*.xlsx")]
    # 2 firma tutanağı + yan dosyalar (VKN listesi, geçersiz satırlar, özet)
    assert any(n.startswith("1)") for n in dosyalar)
    assert any(n.startswith("2)") for n in dosyalar)
    assert any(n.startswith("VKN_LISTESI") for n in dosyalar)
    assert any(n.startswith("GECERSIZ_SATIRLAR") for n in dosyalar)
    assert any(n.startswith("OZET_RAPOR") for n in dosyalar)
    # Kalıcı işlem günlüğü (.txt) yazılmış olmalı (denetim izi)
    assert any(p.name.startswith("ISLEM_GUNLUGU") for p in cikis.glob("*.txt"))
    # İlerleme geri çağrısı en az bir kez çağrılmış ve sona ulaşmış olmalı
    assert ilerleme_kayit and ilerleme_kayit[-1] == (2, 2)


# ══════════════════════════════════════════════════════════════════════════
#  Eski GİB tipi (başlık alt satırda, adsız seri sütunu, kesme işaretli KDV'si)
# ══════════════════════════════════════════════════════════════════════════
def _eski_gib_yaz(yol):
    """Eski GİB tipi: 2 başlık/altbilgi satırı, sonra gerçek başlık; seri sütunu
    adsız (Unnamed), 'KDV'si' kesme işaretli."""
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["İNDİRİLECEK KDV LİSTESİ"])
    ws.append(["Dönem: 01/2026"])
    ws.append(["Alış Faturasının Tarihi", "", "Alış Faturasının Sıra No'su",
               "Satıcının Adı-Soyadı / Ünvanı", "Satıcının Vergi Kimlik Numarası",
               "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"])
    ws.append(["2026-01-05", "", "F1", "FIRMA X", "1000000010", 500000, 90000])
    ws.append(["2026-01-06", "", "F2", "FIRMA Y", "1000000011", 50000, 9000])
    wb.save(yol)


def test_eski_gib_okuma_ve_secim(tmp_path):
    yol = tmp_path / "OCAK_2026.xlsx"
    _eski_gib_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    # Başlık alt satırdan doğru bulunmalı; kritik sütunlar eşleşmeli
    assert exay.sutun_bul(list(df.columns), ['vergi kimlik']) is not None
    assert exay.kdv_sutunu_bul(list(df.columns)) == "KDV'si"
    assert exay.sutun_bul(list(df.columns), ['kdv hariç tutarı']) is not None
    secilen, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert "1000000010" in secilen        # 500.000 ≥ 150.000 (tek fatura)
    assert "1000000011" not in secilen     # 50.000 eşik altı


# ══════════════════════════════════════════════════════════════════════════
#  CSV girdi (Türkçe kodlama + noktalı virgül ayraç + muhasebe tipi)
# ══════════════════════════════════════════════════════════════════════════
def test_csv_muhasebe_okuma(tmp_path):
    yol = tmp_path / "OCAK_2026.csv"
    icerik = (
        "Hesap Kodu;Tarih;Fatura No;Vergi Kimlik No;Açıklama;Borç;Matrah\n"
        "191.01;2026-01-10;F1;0071419747;FIRMA A;36000,00;200000,00\n"
        "191.01;2026-01-11;F2;1000000002;FIRMA B;9000,00;50000,00\n"
    )
    yol.write_text(icerik, encoding="cp1254")   # TR Windows kodlaması
    df = exay.ana_listeyi_oku(str(yol))
    # Muhasebe eşlemesi uygulanmış olmalı
    assert exay.kdv_sutunu_bul(list(df.columns)) is not None
    secilen, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    # 8 haneli VKN önde sıfır tamamlanarak geçerli sayılmalı ve seçilmeli
    assert "0071419747" in secilen


# ══════════════════════════════════════════════════════════════════════════
#  Kriter doğrulama
# ══════════════════════════════════════════════════════════════════════════
@pytest.mark.parametrize("et,eto,y,ok", [
    ("150000", "450000", "80", True),
    ("0", "450000", "80", False),        # limit 0 olamaz
    ("150000", "-1", "80", False),       # negatif limit
    ("150000", "450000", "0", False),    # yüzde 0 olamaz
    ("150000", "450000", "120", False),  # yüzde 100'den büyük olamaz
    ("abc", "450000", "80", False),      # sayı değil
    ("450000", "150000", "80", True),    # tek>toplam teknik olarak geçerli (engellenmez)
])
def test_kriter_dogrula(et, eto, y, ok):
    assert exay.kriter_dogrula(et, eto, y)[0] is ok


# ══════════════════════════════════════════════════════════════════════════
#  Doğruluk kontrolleri (uyarı üreticiler)
# ══════════════════════════════════════════════════════════════════════════
def test_kdv_tutarlilik_kotu_esleme_uyarir():
    # KDV = matrah (oran %100) → şüpheli, uyarı beklenir
    df = pd.DataFrame({
        "Alış Faturasının KDV Hariç Tutarı": [1000, 2000],
        "KDV si": [1000, 2000],
    })
    assert exay.kdv_tutarlilik_kontrol(df)          # boş değil (uyarı var)


def test_kdv_tutarlilik_iyi_esleme_uyarmaz():
    df = pd.DataFrame({
        "Alış Faturasının KDV Hariç Tutarı": [1000, 2000],
        "KDV si": [180, 360],                        # %18
    })
    assert exay.kdv_tutarlilik_kontrol(df) == []


def test_mukerrer_fatura():
    df = pd.DataFrame({
        "Satıcının Vergi Kimlik Numarası": ["1000000001", "1000000001", "1000000002"],
        "Alış Faturasının Sıra No'su": ["F1", "F1", "F2"],
    })
    m = exay.mukerrer_fatura_bul(df)
    assert m == [("1000000001", "F1", 2)]


def test_ay_yil_bicimleri():
    from datetime import datetime as _dt
    assert exay._ay_yil("2026-04-15") == "04.2026"    # ISO
    assert exay._ay_yil("15.04.2026") == "04.2026"    # gün.ay.yıl
    assert exay._ay_yil(_dt(2026, 4, 15)) == "04.2026"
    assert exay._ay_yil("") is None


def test_donem_disi_tarih():
    df = pd.DataFrame({
        "Alış Faturasının Tarihi": ["2026-04-01", "2026-04-02", "2026-05-15"],
    })
    toplam, disi = exay.donem_disi_tarih_kontrol(df, "04.2026")
    assert toplam == 3 and disi == 1                  # yalnızca Mayıs dönem dışı


# ══════════════════════════════════════════════════════════════════════════
#  Sürükle-bırak yolu ayıklama
# ══════════════════════════════════════════════════════════════════════════
def test_dnd_ayikla():
    ayik = exay.KDVBolmeApp._dnd_ayikla
    assert ayik("/tmp/a.xlsx") == ["/tmp/a.xlsx"]
    assert ayik("{/tmp/bir iki.xlsx} /tmp/c.csv") == ["/tmp/bir iki.xlsx", "/tmp/c.csv"]


# ══════════════════════════════════════════════════════════════════════════
#  PDF çıktı (reportlab kuruluysa)
# ══════════════════════════════════════════════════════════════════════════
@pytest.mark.skipif(not exay.pdf_destekli(), reason="reportlab kurulu değil")
def test_pdf_uretimi(tmp_path):
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    secilen, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    grp = secilen["1000000001"][0]
    pdf = tmp_path / "firma.pdf"
    exay.firma_pdf_olustur(grp, str(pdf), list(df.columns),
                           vkn="1000000001", unvan="ÇĞİÖŞÜ FİRMA", donem="04.2026")
    assert pdf.exists() and pdf.read_bytes()[:5] == b"%PDF-"


def test_dosyalari_isle_pdf_uret(tmp_path):
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        pdf_uret=True)
    cikis = tmp_path / "Hazır Tutanaklar"
    pdfler = list(cikis.glob("*.pdf"))
    if exay.pdf_destekli():
        assert len(pdfler) == 2                       # A ve B için PDF
    else:
        assert len(pdfler) == 0                       # reportlab yoksa sessizce atlanır


# ══════════════════════════════════════════════════════════════════════════
#  Çıktı klasörü yönlendirme (cikis_kok)
# ══════════════════════════════════════════════════════════════════════════
def test_cikis_kok_yonlendirme(tmp_path):
    kaynak = tmp_path / "kaynak"
    kaynak.mkdir()
    yol = kaynak / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)
    hedef = tmp_path / "baska_yer"
    hedef.mkdir()
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        cikis_kok=str(hedef))
    # Çıktı kaynağın yanına DEĞİL, seçilen hedefe yazılmalı
    assert (hedef / "Hazır Tutanaklar").is_dir()
    assert not (kaynak / "Hazır Tutanaklar").exists()


# ══════════════════════════════════════════════════════════════════════════
#  WORD ŞABLON EŞLEŞTİRME (VKN ile) — saf ayrıştırıcılar
# ══════════════════════════════════════════════════════════════════════════
def _tutanak_metni(unvan, vd_hucre):
    """Gerçek tutanak düzenini (sekmeli hücreler) taklit eden sentetik metin."""
    return (
        "KATMA DEĞER VERGİSİ İADESİ KARŞIT İNCELEME TUTANAĞI\n"
        "YEMİNLİ MALİ MÜŞAVİRİN\t\tAdı Soyadı\tSABRİ HAMAMCI\t\t"
        "Vergi Dairesi ve Sicil No\tGAZİKENT V.D. / 464 100 8244\t\t"
        "İADE TALEBİNDE BULUNAN FİRMANIN\t\tÜnvanı\tİNALOĞLU İNŞAAT\t\t"
        "Vergi Dairesi/Nosu\tŞAHİNBEY / 475 056 9431\t\t"
        "NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN\t\t"
        f"Ünvanı\t{unvan}\t\tVergi Dairesi/Nosu\t{vd_hucre}\t\t"
        "Adresi\tBİR ADRES\t\tİNCELEME DAYANAĞI\t31.01.2026 Tarih ve 09 Sayılı\n"
        "Karşıt İncelemeye Konu Fatura ve Benzeri Belgeye ilişkin bilgiler:\n"
    )


def test_vkn_metinden_ayikla():
    assert exay._vkn_metinden_ayikla("493 061 9102") == "4930619102"   # boşluklu
    assert exay._vkn_metinden_ayikla("ASIM GÜNDÜZ V.D. – 30490690382") == "30490690382"
    assert exay._vkn_metinden_ayikla("71419747") == "0071419747"       # 8 hane → zfill
    assert exay._vkn_metinden_ayikla("0000000000") is None             # yer tutucu
    assert exay._vkn_metinden_ayikla("yok") is None


def test_sablon_vkn_metinden():
    # 10 haneli VKN (boşluklu) doğru bloktan alınmalı, ünvan da
    vkn, unvan = exay.sablon_vkn_metinden(
        _tutanak_metni("MESAKO MADEN VE ENERJİ TİC. LTD. ŞTİ.", "ŞAHİNBEY V.D. 6190914983"))
    assert vkn == "6190914983"
    assert "MESAKO" in unvan
    # 11 haneli TCKN + tire ile
    vkn2, unvan2 = exay.sablon_vkn_metinden(
        _tutanak_metni("ONUR FURKAN KARTA", "ASIM GÜNDÜZ V.D. – 30490690382"))
    assert vkn2 == "30490690382"
    assert unvan2 == "ONUR FURKAN KARTA"
    # Blok yoksa (ör. üst yazı) → (None, None)
    assert exay.sablon_vkn_metinden("Sayı: YMM 27103572 ... GAZİANTEP") == (None, None)


def test_sablonlari_indeksle(tmp_path, monkeypatch):
    metinler = {}
    for ad, unvan, vkn in [("a.doc", "FIRMA A", "1234567890"),
                           ("b.doc", "FIRMA B", "9876543210"),
                           ("ustyazi.doc", None, None)]:
        p = tmp_path / ad
        p.write_bytes(b"stub")     # gerçek .doc gerekmez; okuma monkeypatch'li
        if unvan:
            metinler[str(p)] = _tutanak_metni(unvan, f"V.D. {vkn}")
        else:
            metinler[str(p)] = "Sayı: YMM 123 üst yazı"
    monkeypatch.setattr(exay, "_doc_metni_oku", lambda p: metinler[str(p)])
    idx = exay.sablonlari_indeksle(str(tmp_path))
    assert idx["1234567890"].endswith("a.doc")
    assert idx["9876543210"].endswith("b.doc")
    assert len(idx) == 2               # üst yazı eşleşmez


def test_fatura_tablosu_mu():
    assert exay._fatura_tablosu_mu(["FATURANIN", "MALIN", "Tarihi", "Numarası",
                                    "Cinsi", "Miktarı", "Tutarı", "KDV Tutarı"])
    assert not exay._fatura_tablosu_mu(["Defterin Nevi", "Tasdik Makamı"])


def test_word_fatura_satiri_ve_miktar():
    df = pd.DataFrame({
        "Alış Faturasının Tarihi": ["2026-04-01"],
        "Alış Faturasının Sıra No'su": ["BBK1"],
        "Alınan Mal ve/veya Hizmetin Cinsi": ["YEMEK BEDELİ"],
        "Miktarı": ["4.880 ADET"],
        "Alış Faturasının KDV Hariç Tutarı": [683200],
        "KDV si": [68320],
    })
    hucre = exay._word_fatura_satiri(df.iloc[0], list(df.columns))
    # [Tarih, No, Cinsi, Miktar, Tutar, KDV, DefterKayıt(boş)]
    assert hucre[0] == "01.04.2026"
    assert hucre[1] == "BBK1"
    assert hucre[2] == "YEMEK BEDELİ"
    assert hucre[3] == "4.880 ADET"          # miktar sütunu kullanıldı
    assert hucre[4] == "683.200,00"          # TR biçim
    assert hucre[5] == "68.320,00"
    assert hucre[6] == ""                     # defter kayıt boş


def test_tr_para_str():
    assert exay._tr_para_str(683200) == "683.200,00"
    assert exay._tr_para_str("1234567,89") == "1.234.567,89"
    assert exay._tr_para_str("") == ""


def test_word_com_yoksa_hata():
    # Bu ortamda Word/pywin32 yok → firma_word_olustur RuntimeError vermeli
    if exay.word_destekli():
        pytest.skip("Word otomasyonu mevcut; negatif yol test edilemez")
    df = pd.DataFrame({"Alış Faturasının Tarihi": ["2026-04-01"]})
    with pytest.raises(RuntimeError):
        exay.firma_word_olustur("yok.doc", df, "cikti.doc", list(df.columns))


def test_dosyalari_isle_sablon_eslesme_raporu(tmp_path, monkeypatch):
    """Şablon klasörü verildiğinde: eşleşme yapılır ve WORD_ESLESME raporu üretilir
    (Word olmadan da). Eşleşme VKN ile olmalı."""
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)     # FIRMA A=1000000001 (seçilir), FIRMA B=1000000002 (seçilir)
    sablon_kl = tmp_path / "sablonlar"
    sablon_kl.mkdir()
    pa = sablon_kl / "firmaA.doc"; pa.write_bytes(b"stub")
    metin = _tutanak_metni("FIRMA A", "V.D. 1000000001")   # yalnızca A'nın şablonu
    monkeypatch.setattr(exay, "_doc_metni_oku",
                        lambda p: metin if str(p) == str(pa) else "üst yazı")
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sablon_kl))
    rapor = list((tmp_path / "Hazır Tutanaklar").glob("WORD_ESLESME_*.xlsx"))
    assert rapor, "Şablon eşleşme raporu üretilmedi"
    wb = openpyxl.load_workbook(rapor[0]); ws = wb.active
    satirlar = {ws.cell(r, 1).value: ws.cell(r, 3).value for r in range(2, ws.max_row + 1)}
    assert "1000000001" in satirlar             # A eşleşti
    assert "1000000002" in satirlar             # B şablonsuz
    assert satirlar["1000000002"] == "Şablon yok"


def test_gercek_gib_basliklari_word_eslemesi():
    """Gerçek 'Yeni GİB' başlık düzeni (uzun sütun adları, kesme işaretli KDV'si,
    birleşik VKN/TC sütunu, 'Alınan Mal ve/veya Hizmetin Miktarı') doğru eşlenmeli
    ve Word satırı [Tarih,No,Cinsi,Miktar,Tutar,KDV,''] üretmeli."""
    kols = [
        "Alış Faturasının Tarihi", "Alış Faturasının Serisi",
        "Alış Faturasının Sıra No'su", "Satıcının Adı-Soyadı / Ünvanı",
        "Satıcının Vergi Kimlik Numarası / TC Kimlik Numarası",
        "Alınan Mal ve/veya Hizmetin Cinsi", "Alınan Mal ve/veya Hizmetin Miktarı",
        "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si",
        "Toplam İndirilecek KDV Tutarı",
    ]
    df = pd.DataFrame([[
        "2026-07-25", None, "0012026165876119", "TURKCELL ILETİSİM HİZMETLERİ A.S.",
        "8770013456", "İLETİŞİM HİZMET BEDELİ", "5 ADET", 1208.34, 241.67, 241.67,
    ]], columns=kols)
    b = exay.bulunan_sutunlar(df)
    assert b["Matrah"] == "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı"
    assert b["KDV"] == "KDV'si"                          # kesme işaretli, toplam KDV ile karışmaz
    assert exay.sutun_bul(kols, ["miktar"]) == "Alınan Mal ve/veya Hizmetin Miktarı"
    satir = exay._word_fatura_satiri(df.iloc[0], kols)
    assert satir == ["25.07.2026", "0012026165876119", "İLETİŞİM HİZMET BEDELİ",
                     "5 ADET", "1.208,34", "241,67", ""]


# ══════════════════════════════════════════════════════════════════════════
#  GUI kurulum dumanı — _ui() sahte pencereyle çalışır; eksik metot/callback
#  bağlamaları (ör. bind command'ları) burada yakalanır (regresyon güvencesi).
# ══════════════════════════════════════════════════════════════════════════
def test_gui_kurulur_ve_callbackler_tanimli():
    from unittest.mock import MagicMock
    root = MagicMock(name="root")
    app = exay.KDVBolmeApp(root)     # __init__ → _ui() çalışır; eksik metot patlar
    # _ui / _surukle_birak içinde referans verilen tüm callback'ler tanımlı olmalı
    for m in ["_tiklayarak_sec", "_isle", "_isle_coklu", "_batch_worker",
              "_cikis_klasoru_sec", "_sablon_klasoru_sec", "_cikis_ozet",
              "_sablon_ozet", "_kriter_al", "_ilerleme", "_tamam", "_log",
              "_ayar_kaydet", "_birak_guncelle", "_dnd_ayikla"]:
        assert callable(getattr(app, m)), f"Eksik/çağrılamaz metot: {m}"
