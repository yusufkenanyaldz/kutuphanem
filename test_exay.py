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
    # Diğer yaygın biçimler de GİB biçimine (GG.AA.YYYY) çevrilir
    assert exay.tarih_fmt("15/04/2026") == "15.04.2026"
    assert exay.tarih_fmt("5.4.2026") == "05.04.2026"
    assert exay.tarih_fmt("15-04-2026") == "15.04.2026"
    assert exay.tarih_fmt("2026-04-15 00:00:00") == "15.04.2026"
    # Tanınmayan biçim olduğu gibi döner
    assert exay.tarih_fmt("Nisan sonu") == "Nisan sonu"


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
    assert idx["1234567890"][0].endswith("a.doc")
    assert idx["9876543210"][0].endswith("b.doc")
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
def test_tamam_cb_ozet_metrik_ve_geriye_uyum(tmp_path):
    """dosyalari_isle bitince tamam_cb'ye 4. argüman olarak özet metrik sözlüğü
    geçirir; yalnızca 3 argüman kabul eden eski geri çağrılar da kırılmamalı."""
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)                                   # 2 firma seçilir, 1 geçersiz satır
    yakalanan = {}
    def cb4(kl, b, h, ozet=None):
        yakalanan['ozet'] = ozet
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, cb4, lambda *a: None,
                        cikti_turu='excel')
    oz = yakalanan['ozet']
    assert oz['secilen'] == 2 and oz['uretilen'] == 2
    assert isinstance(oz['kapsam'], (int, float)) and oz['kapsam'] > 0
    assert oz['gecersiz'] >= 1                            # placeholder VKN'li satır

    # 3 argümanlı eski imza da çalışmalı (TypeError'a düşüp 3-arg çağrılır)
    calisti = {}
    def cb3(kl, b, h):
        calisti['ok'] = True
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, cb3, lambda *a: None,
                        cikti_turu='excel')
    assert calisti.get('ok')


def test_gui_kurulur_ve_callbackler_tanimli():
    from unittest.mock import MagicMock
    root = MagicMock(name="root")
    app = exay.KDVBolmeApp(root)     # __init__ → _ui() çalışır; eksik metot patlar
    # _ui / _surukle_birak içinde referans verilen tüm callback'ler tanımlı olmalı
    for m in ["_tiklayarak_sec", "_isle", "_isle_coklu", "_batch_worker",
              "_cikis_klasoru_sec", "_sablon_klasoru_sec", "_cikis_ozet",
              "_sablon_ozet", "_kriter_al", "_ilerleme", "_tamam", "_log",
              "_ayar_kaydet", "_birak_guncelle", "_dnd_ayikla",
              "_dosya_sec", "_olustur_tikla", "_segment_sec", "_kapsam_kaydir",
              "_log_ac_kapa", "_metrik_guncelle", "_word_blok_guncelle"]:
        assert callable(getattr(app, m)), f"Eksik/çağrılamaz metot: {m}"


# ══════════════════════════════════════════════════════════════════════════
#  .docx şablon yolu — okuma (VKN) + yazma (fatura tablosu), Word GEREKTİRMEZ
# ══════════════════════════════════════════════════════════════════════════
def _docx_sablon_yaz(yol, unvan, vd_hucre):
    """Gerçek tutanak yapısını taklit eden sentetik .docx şablon üretir:
    2 sütunlu bilgi tablosu (NEZDİNDE bloğu) + 7 sütunlu fatura tablosu."""
    import docx
    d = docx.Document()
    t0 = d.add_table(rows=0, cols=2)
    for etiket, deger in [
        ("YEMİNLİ MALİ MÜŞAVİRİN", "SABRİ HAMAMCI"),
        ("İADE TALEBİNDE BULUNAN FİRMANIN", "İADE TALEBİNDE BULUNAN FİRMANIN"),
        ("Ünvanı", "İNALOĞLU İNŞAAT"),
        ("Vergi Dairesi/Nosu", "ŞAHİNBEY / 475 056 9431"),
        ("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN", "NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN"),
        ("Ünvanı", unvan),
        ("Vergi Dairesi/Nosu", vd_hucre),
        ("İNCELEME DAYANAĞI", "31.01.2026 Tarih ve 09 Sayılı"),
    ]:
        r = t0.add_row().cells
        r[0].text = etiket; r[1].text = deger
    d.add_paragraph("Karşıt İncelemeye Konu Fatura ve Benzeri Belgeye ilişkin bilgiler:")
    t2 = d.add_table(rows=3, cols=7)
    b0 = ["FATURANIN", "FATURANIN", "MALIN", "MALIN", "MALIN", "MALIN", "Defter Kayıt"]
    b1 = ["Tarihi", "Numarası", "Cinsi", "Miktarı", "Tutarı", "KDV Tutarı", "Tarihi/Nosu"]
    for c, v in enumerate(b0): t2.rows[0].cells[c].text = v
    for c, v in enumerate(b1): t2.rows[1].cells[c].text = v
    d.save(yol)


def test_docx_sablon_vkn_oku(tmp_path):
    yol = tmp_path / "ispa.docx"
    _docx_sablon_yaz(yol, "İSPA İNŞ. SAN. PAZ. A.Ş.", "ÜSKÜDAR V.D. – 481 001 7371")
    vkn, unvan = exay.sablon_vkn_oku(str(yol))
    assert vkn == "4810017371"                       # boşluklu VKN düzeltilir
    assert "İSPA" in unvan


def test_sablonlari_indeksle_docx_dahil(tmp_path):
    _docx_sablon_yaz(tmp_path / "a.docx", "FIRMA A", "V.D. 1234567890")
    idx = exay.sablonlari_indeksle(str(tmp_path))
    assert idx["1234567890"][0].endswith("a.docx")   # .docx da indekslenir


def test_firma_docx_olustur_tabloyu_doldurur(tmp_path):
    import docx
    sablon = tmp_path / "sablon.docx"
    _docx_sablon_yaz(sablon, "İSPA İNŞ. SAN. PAZ. A.Ş.", "V.D. 4810017371")
    # iki faturalı bir firma
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
            "Alınan Mal ve/veya Hizmetin Cinsi", "Alınan Mal ve/veya Hizmetin Miktarı",
            "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"]
    firma = pd.DataFrame([
        ["2026-07-22", "S0120260058", "İNŞAAT MALZ.", "6 ADET", 2595795.15, 516799.39],
        ["2026-07-25", "S0120260059", "DEMİR", "10 TON", 100000.00, 20000.00],
    ], columns=kols)
    cikti = tmp_path / "cikti.docx"
    exay.firma_docx_olustur(str(sablon), firma, str(cikti), kols)
    d = docx.Document(str(cikti))
    fatura = next(t for t in d.tables if len(t.columns) == 7)
    # 2 başlık satırı + 2 veri satırı = 4
    assert len(fatura.rows) == 4
    veri = [[c.text.strip() for c in r.cells] for r in fatura.rows[2:]]
    assert veri[0][0] == "22.07.2026"
    assert veri[0][1] == "S0120260058"
    assert veri[0][3] == "6 ADET"                    # miktar
    assert veri[0][4] == "2.595.795,15"              # TR tutar
    assert veri[0][5] == "516.799,39"
    assert veri[1][2] == "DEMİR"
    assert veri[1][3] == "10 TON"


def test_dosyalari_isle_docx_word_uretir(tmp_path):
    """Uçtan uca: .docx şablon eşleşince gerçek Word tutanağı üretilmeli (Word yok)."""
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)     # FIRMA A = 1000000001 (seçilir)
    sk = tmp_path / "sablonlar"; sk.mkdir()
    _docx_sablon_yaz(sk / "firmaA.docx", "FIRMA A", "V.D. 1000000001")
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk))
    uretilen = list((tmp_path / "Hazır Tutanaklar").glob("*.docx"))
    assert any("FIRMA A" in p.name for p in uretilen), "İSPA benzeri Word tutanağı üretilmedi"
    assert not any("1000000001" in p.name for p in uretilen), "Word adında VKN olmamalı"


# ══════════════════════════════════════════════════════════════════════════
#  Çıktı türü seçimi: yalnız Excel / yalnız Word / ikisi
# ══════════════════════════════════════════════════════════════════════════
def _liste_ve_sablon(tmp_path):
    """FIRMA A (1000000001) seçilir; A için bir .docx şablon hazırlanır."""
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)
    sk = tmp_path / "sablonlar"; sk.mkdir()
    _docx_sablon_yaz(sk / "firmaA.docx", "FIRMA A", "V.D. 1000000001")
    return yol, sk


def test_cikti_yalniz_excel(tmp_path):
    yol, sk = _liste_ve_sablon(tmp_path)
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='excel')
    kl = tmp_path / "Hazır Tutanaklar"
    assert list(kl.glob("*.xlsx"))                       # Excel var
    assert not list(kl.glob("*.docx"))                    # Word YOK
    assert not list(kl.glob("WORD_ESLESME_*.xlsx"))       # eşleşme raporu bile yok


def test_cikti_yalniz_word(tmp_path):
    yol, sk = _liste_ve_sablon(tmp_path)
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='word')
    kl = tmp_path / "Hazır Tutanaklar"
    docx_ler = list(kl.glob("*.docx"))
    # Firma tutanağı .docx üretilir; VKN listesi/özet .xlsx yan dosyaları hariç
    # FİRMA tutanağı .xlsx OLMAMALI (numaralı 'N) …_.xlsx')
    firma_xlsx = [p for p in kl.glob("*.xlsx")
                  if p.name[0].isdigit() and ")" in p.name[:4]]
    assert any("FIRMA A" in p.name for p in docx_ler)     # A için Word üretildi (ünvan+dönem adı)
    assert not firma_xlsx                                  # firma Excel tutanağı YOK


def test_cikti_ikisi(tmp_path):
    yol, sk = _liste_ve_sablon(tmp_path)
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='ikisi')
    kl = tmp_path / "Hazır Tutanaklar"
    firma_xlsx = [p for p in kl.glob("*.xlsx")
                  if p.name[0].isdigit() and ")" in p.name[:4]]
    assert firma_xlsx                                      # Excel tutanakları var (A ve B)
    assert any("FIRMA A" in p.name for p in kl.glob("*.docx"))     # A için Word da var


# ══════════════════════════════════════════════════════════════════════════
#  İnceleme Dayanağı (sözleşme) otomatik güncelleme + Türkçe casing güvenliği
# ══════════════════════════════════════════════════════════════════════════
def test_ascii_kucuk_turkce():
    assert exay._ascii_kucuk("İNCELEME DAYANAĞI") == "inceleme dayanagi"
    assert "inceleme dayana" in exay._ascii_kucuk("İNCELEME DAYANAĞI")
    assert exay._ascii_kucuk("ŞUBAT") == "subat"


def test_inceleme_dayanagi_docx_gunceller(tmp_path):
    import docx
    sablon = tmp_path / "s.docx"
    _docx_sablon_yaz(sablon, "İSPA A.Ş.", "V.D. 4810017371")
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
            "Alınan Mal ve/veya Hizmetin Cinsi", "Alınan Mal ve/veya Hizmetin Miktarı",
            "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"]
    firma = pd.DataFrame([["2026-07-22", "F1", "MAL", "6 ADET", 1000.0, 180.0]], columns=kols)
    cikti = tmp_path / "c.docx"
    YENI = "30.06.2026 Tarih ve 27 Sayılı Tam Tasdik Sözleşmesi"
    exay.firma_docx_olustur(str(sablon), firma, str(cikti), kols, inceleme_dayanagi=YENI)
    d = docx.Document(str(cikti))
    bulundu = False
    for tbl in d.tables:
        for row in tbl.rows:
            cs = row.cells
            if cs and 'inceleme dayana' in exay._ascii_kucuk(cs[0].text):
                assert cs[1].text.strip() == YENI          # değer güncellendi
                assert 'inceleme dayana' in exay._ascii_kucuk(cs[0].text)  # etiket sabit
                bulundu = True
    assert bulundu


# ══════════════════════════════════════════════════════════════════════════
#  Çok-firmalı tek .docx: içinden ilgili firmayı VKN ile bulup izole etme
# ══════════════════════════════════════════════════════════════════════════
def _docx_coklu_yaz(yol, firmalar):
    """Tek dosyada birden çok firma tutanağı üretir; her blok
    'KATMA DEĞER ... KARŞIT İNCELEME TUTANAĞI' başlığıyla başlar (gerçek yapı)."""
    import docx
    d = docx.Document()
    for k, (unvan, vd) in enumerate(firmalar):
        if k > 0:
            d.add_page_break()
        d.add_paragraph("KATMA DEĞER VERGİSİ İADESİ KARŞIT İNCELEME TUTANAĞI")
        t0 = d.add_table(rows=0, cols=2)
        for e, v in [("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN", e := "NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN"),
                     ("Ünvanı", unvan), ("Vergi Dairesi/Nosu", vd),
                     ("İNCELEME DAYANAĞI", "31.01.2026 Tarih ve 09 Sayılı")]:
            r = t0.add_row().cells; r[0].text = e if isinstance(e, str) else e; r[1].text = v
        d.add_paragraph("Karşıt İncelemeye Konu Fatura ve Benzeri Belgeye ilişkin bilgiler:")
        t2 = d.add_table(rows=3, cols=7)
        for c, v in enumerate(["FATURANIN", "FATURANIN", "MALIN", "MALIN", "MALIN", "MALIN", "Defter Kayıt"]):
            t2.rows[0].cells[c].text = v
        for c, v in enumerate(["Tarihi", "Numarası", "Cinsi", "Miktarı", "Tutarı", "KDV Tutarı", "Tarihi/Nosu"]):
            t2.rows[1].cells[c].text = v
    d.save(yol)


def test_docx_coklu_firma_kayitlari(tmp_path):
    yol = tmp_path / "hepsi.docx"
    _docx_coklu_yaz(yol, [("FIRMA A", "V.D. 1234567890"),
                          ("FIRMA B", "V.D. 9876543210")])
    kayit = exay._sablon_kayitlari(str(yol))
    vknler = {v: b for v, u, b, y in kayit}
    assert vknler == {"1234567890": 0, "9876543210": 1}   # iki firma, blok 0/1
    idx = exay.sablonlari_indeksle(str(tmp_path))
    assert idx["1234567890"] == (str(yol), 0)
    assert idx["9876543210"] == (str(yol), 1)


def test_docx_blok_izole_ve_doldur(tmp_path):
    import docx
    yol = tmp_path / "hepsi.docx"
    _docx_coklu_yaz(yol, [("FIRMA A", "V.D. 1234567890"),
                          ("FIRMA B", "V.D. 9876543210")])
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
            "Alınan Mal ve/veya Hizmetin Cinsi", "Alınan Mal ve/veya Hizmetin Miktarı",
            "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"]
    firma = pd.DataFrame([["2026-07-02", "B1", "MAL", "5 ADET", 2000.0, 360.0]], columns=kols)
    cikti = tmp_path / "b.docx"
    # blok 1 = FIRMA B izole edilmeli
    exay.firma_docx_olustur(str(yol), firma, str(cikti), kols, blok=1)
    d = docx.Document(str(cikti))
    # Tek firma → sadece B'nin bloğu (1 fatura tablosu) kalmalı
    fat = [t for t in d.tables if len(t.columns) == 7]
    assert len(fat) == 1
    # içindeki firma B olmalı (A değil)
    metin = "\n".join("\t".join(c.text.strip() for c in r.cells)
                      for t in d.tables for r in t.rows)
    vkn, _ = exay.sablon_vkn_metinden(metin)
    assert vkn == "9876543210"
    # fatura satırı dolduruldu
    veri = [c.text.strip() for c in fat[0].rows[2].cells]
    assert veri[1] == "B1" and veri[3] == "5 ADET"


def test_docx_coklu_uctan_uca(tmp_path):
    """Liste + çok-firmalı tek şablon dosyası → her firma ayrı tutanak."""
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)     # FIRMA A=1000000001, FIRMA B=1000000002 seçilir
    sk = tmp_path / "sablonlar"; sk.mkdir()
    _docx_coklu_yaz(sk / "hepsi.docx", [("FIRMA A", "V.D. 1000000001"),
                                        ("FIRMA B", "V.D. 1000000002")])
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='word')
    uretilen = sorted(p.name for p in (tmp_path / "Hazır Tutanaklar").glob("*.docx"))
    assert any("FIRMA A" in n for n in uretilen)      # A tek dosyadan bulundu
    assert any("FIRMA B" in n for n in uretilen)      # B tek dosyadan bulundu


# ══════════════════════════════════════════════════════════════════════════
#  Tek dosyada birleştirme (word_tek_dosya) + boş/yedek şablon (bos_sablon)
# ══════════════════════════════════════════════════════════════════════════
def test_word_tek_dosya_birlestirme(tmp_path):
    import docx
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)   # A=1000000001, B=1000000002 seçilir
    sk = tmp_path / "sablonlar"; sk.mkdir()
    _docx_sablon_yaz(sk / "a.docx", "FIRMA A", "V.D. 1000000001")
    _docx_sablon_yaz(sk / "b.docx", "FIRMA B", "V.D. 1000000002")
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='word', word_tek_dosya=True)
    kl = tmp_path / "Hazır Tutanaklar"
    birlesik = list(kl.glob("KARSIT_INCELEME_TUTANAKLAR_*.docx"))
    assert len(birlesik) == 1                              # tek birleşik dosya
    # ayrı firma .docx'i OLMAMALI
    assert not [p for p in kl.glob("*.docx") if p.name[0].isdigit()]
    d = docx.Document(str(birlesik[0]))
    fat = [t for t in d.tables if len(t.columns) == 7]
    assert len(fat) == 2                                   # A ve B blokları tek dosyada


def test_bos_sablon_eslesmeyen_firma(tmp_path):
    import docx
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)   # A=1000000001, B=1000000002
    sk = tmp_path / "sablonlar"; sk.mkdir()
    _docx_sablon_yaz(sk / "a.docx", "FIRMA A", "V.D. 1000000001")   # yalnızca A'nın şablonu
    bos = tmp_path / "bos.docx"
    _docx_sablon_yaz(bos, "", "")                          # boş NEZDİNDE (ünvan/vd yok)
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='word', bos_sablon=str(bos))
    kl = tmp_path / "Hazır Tutanaklar"
    # B için boş şablondan üretilmiş bir dosya olmalı
    b_dosya = [p for p in kl.glob("*.docx") if "FIRMA B" in p.name]
    assert b_dosya, "Boş şablondan B tutanağı üretilmedi"
    d = docx.Document(str(b_dosya[0]))
    metin = "\n".join("\t".join(c.text.strip() for c in r.cells)
                      for t in d.tables for r in t.rows)
    vkn, unvan = exay.sablon_vkn_metinden(metin)
    assert vkn == "1000000002"                            # bilinen VKN yazıldı
    assert unvan == "FIRMA B"                             # bilinen ünvan yazıldı
    # WORD_ESLESME raporunda "Boş şablon" durumu geçmeli
    rap = list(kl.glob("WORD_ESLESME_*.xlsx"))[0]
    wb = openpyxl.load_workbook(rap); ws = wb.active
    durumlar = {ws.cell(r, 1).value: ws.cell(r, 3).value for r in range(2, ws.max_row + 1)}
    assert "Boş şablon" in str(durumlar.get("1000000002", ""))


# ══════════════════════════════════════════════════════════════════════════
#  Birleşik .doc tanıma (metin bölme) — üretim COM gerektirir, burada tanıma test
# ══════════════════════════════════════════════════════════════════════════
def test_metni_bloklara_ayir():
    tek = _tutanak_metni("FIRMA A", "V.D. 1234567890")
    assert len(exay._metni_bloklara_ayir(tek)) == 1
    ikili = _tutanak_metni("FIRMA A", "V.D. 1234567890") + \
            _tutanak_metni("FIRMA B", "V.D. 9876543210")
    bloklar = exay._metni_bloklara_ayir(ikili)
    assert len(bloklar) == 2
    assert exay.sablon_vkn_metinden(bloklar[0])[0] == "1234567890"
    assert exay.sablon_vkn_metinden(bloklar[1])[0] == "9876543210"


def test_sablon_kayitlari_doc_tekli(monkeypatch, tmp_path):
    p = tmp_path / "tek.doc"; p.write_bytes(b"stub")
    monkeypatch.setattr(exay, "_doc_metni_oku",
                        lambda x: _tutanak_metni("FIRMA A", "V.D. 1234567890"))
    kayit = exay._sablon_kayitlari(str(p))
    assert kayit == [("1234567890", kayit[0][1], None, str(p))]   # tek firma, blok None


# ══════════════════════════════════════════════════════════════════════════
#  İkinci belge tipi: YMM 'Bilgi İsteme' yazısı (Hakkında Bilgi İstenilen Mükellef)
# ══════════════════════════════════════════════════════════════════════════
def test_ymm_yazisi_vkn_eslesme():
    # Etiketler karışık olabilir (Adresi hücresinde V.D./VKN) ve telefon var
    metin = (
        "Sayı : YMM 27103572/2026-363\tGAZİANTEP\n"
        "İade Talebinde Bulunan Firma\t\tUnvanı\tİNALOĞLU İNŞAAT\t\t"
        "Adresi\tŞAHİNBEY / 475 056 9431\t\tTelefon/Fax\t0 342 502 03 15\n"
        "Hakkında Bilgi İstenilen Mükellefin Altı\t\t"
        "Ünvanı\tOYAK ÇİMENTO FABRİKALARI ANONİM ŞİRKETİ\t\t"
        "Adresi\tANKARA KURUMLAR V.D. – 6120050961\t\t"
        "Vergi Dairesi/Hesap Nosu\tÇUKURAMBAR MAH. 1480 SK.\t\t"
        "Telefon/Fax\t0 312 220 0290\n"
        "İNCELEME DAYANAĞI\t31.01.2026 Tarih ve 09 Sayılı\n"
    )
    vkn, unvan = exay.sablon_vkn_metinden(metin)
    assert vkn == "6120050961"          # karşı firma (telefon 03122200290 DEĞİL)
    assert "OYAK" in unvan


def test_blok_vkn_telefon_karistirmaz():
    blok = "Ünvanı\tX A.Ş.\tTelefon/Fax\t0 342 215 10 70\tVergi Dairesi\tŞAHİNBEY V.D. 6190914983"
    assert exay._blok_vkn(blok) == "6190914983"


# ── YMM yazısı: fatura tablosunun son sütunu 'KDV dahil toplam' (tutanakta boş) ──
def _ymm_yazi_docx_yaz(yol, unvan, vd_hucre):
    """Gerçek YMM 'Bilgi İsteme' yazısını taklit eden sentetik .docx: 'Hakkında
    Bilgi İstenilen Mükellef' bloğu + 7 sütunlu, son sütunu 'KDV dahil toplam'
    olan (tek başlık satırlı) fatura tablosu."""
    import docx
    d = docx.Document()
    d.add_paragraph("Sayı : YMM 27103572/2026-363")
    d.add_paragraph("Konu : Bilgi İsteme")
    t0 = d.add_table(rows=0, cols=2)
    for etiket, deger in [
        ("Hakkında Bilgi İstenilen Mükellefin", "Hakkında Bilgi İstenilen Mükellefin"),
        ("Ünvanı", unvan),
        ("Vergi Dairesi/Hesap Nosu", vd_hucre),
        ("Telefon/Fax", "0 312 220 0290"),
        ("İNCELEME DAYANAĞI", "31.01.2026 Tarih ve 09 Sayılı"),
    ]:
        r = t0.add_row().cells
        r[0].text = etiket; r[1].text = deger
    t2 = d.add_table(rows=1, cols=7)
    bas = ["FAT.TARİHİ", "FAT. NOSU", "MALIN CİNSİ", "MALIN MİKTARI",
           "MATRAH", "KDV", "kdv dahİl toplam"]
    for c, v in enumerate(bas):
        t2.rows[0].cells[c].text = v
    d.save(yol)


def _ornek_firma_df():
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
            "Alınan Mal ve/veya Hizmetin Cinsi", "Alınan Mal ve/veya Hizmetin Miktarı",
            "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"]
    firma = pd.DataFrame([
        ["2026-01-08", "TC42026000001608", "HAZIR BETON", "302 M3", 446641.18, 89328.24],
    ], columns=kols)
    return firma, kols


def test_ymm_yazi_fatura_kdv_dahil_toplam(tmp_path):
    import docx
    sablon = tmp_path / "ymm.docx"
    _ymm_yazi_docx_yaz(sablon, "OYAK ÇİMENTO FABRİKALARI A.Ş.", "ANKARA KURUMLAR V.D. – 6120050961")
    firma, kols = _ornek_firma_df()
    cikti = tmp_path / "ymm_cikti.docx"
    exay.firma_docx_olustur(str(sablon), firma, str(cikti), kols)
    d = docx.Document(str(cikti))
    fatura = next(t for t in d.tables if len(t.columns) == 7)
    veri = [c.text.strip() for c in fatura.rows[1].cells]      # tek başlık satırı
    assert veri[0] == "08.01.2026"
    assert veri[1] == "TC42026000001608"                        # fatura no
    assert veri[2] == "HAZIR BETON"                             # CİNS dolu (boş kalmamalı!)
    assert veri[3] == "302 M3"                                  # miktar
    assert veri[4] == "446.641,18"                              # matrah
    assert veri[5] == "89.328,24"                               # kdv
    assert veri[6] == "535.969,42"                              # KDV dahil toplam = matrah+kdv


def test_tutanak_defter_kayit_bos(tmp_path):
    import docx
    sablon = tmp_path / "tut.docx"
    _docx_sablon_yaz(sablon, "İSPA A.Ş.", "V.D. 4810017371")
    firma, kols = _ornek_firma_df()
    cikti = tmp_path / "tut_cikti.docx"
    exay.firma_docx_olustur(str(sablon), firma, str(cikti), kols)
    d = docx.Document(str(cikti))
    fatura = next(t for t in d.tables if len(t.columns) == 7)
    veri = [c.text.strip() for c in fatura.rows[2].cells]      # 2 başlık satırı
    assert veri[0] == "08.01.2026"
    assert veri[1] == "TC42026000001608"                        # fatura no
    assert veri[2] == "HAZIR BETON"                             # CİNS dolu (boş kalmamalı!)
    assert veri[3] == "302 M3"                                  # miktar
    assert veri[4] == "446.641,18" and veri[5] == "89.328,24"
    assert veri[6] == ""                                        # Defter Kayıt boş kalır


def test_fatura_fazla_sutun_proto_sizmaz_ve_uyarir(tmp_path):
    """Beklenenden fazla sütunlu (8) bir fatura tablosunda: proto satırın örnek
    verisi çıktıya SIZMAMALI (fazla sütun boşaltılır) ve kullanıcı uyarılmalı."""
    import docx
    yol = tmp_path / "sekiz.docx"
    d = docx.Document()
    # NEZDİNDE bloğu (indeks/başlık için) + 8 sütunlu fatura tablosu
    t0 = d.add_table(rows=0, cols=2)
    for e, v in [("NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN",
                  "NEZDİNDE KARŞIT İNCELEME YAPILAN FİRMANIN"),
                 ("Ünvanı", "TEST A.Ş."), ("Vergi Dairesi/Nosu", "V.D. 1234567890")]:
        r = t0.add_row().cells; r[0].text = e; r[1].text = v
    t2 = d.add_table(rows=3, cols=8)
    b0 = ["FATURANIN","FATURANIN","MALIN","MALIN","MALIN","MALIN","Defter Kayıt","Ekstra"]
    b1 = ["Tarihi","Numarası","Cinsi","Miktarı","Tutarı","KDV Tutarı","Tarihi/Nosu","Ekstra"]
    proto = ["x","x","x","x","x","x","x","SIZINTI"]              # proto satırın 8. hücresi
    for c, v in enumerate(b0): t2.rows[0].cells[c].text = v
    for c, v in enumerate(b1): t2.rows[1].cells[c].text = v
    for c, v in enumerate(proto): t2.rows[2].cells[c].text = v
    d.save(yol)

    firma, kols = _ornek_firma_df()
    uyarilar = []
    doc, y = exay._firma_docx_hazirla(str(yol), firma, kols,
                                      log_cb=lambda m, t='': uyarilar.append((m, t)))
    fatura = next(t for t in doc.tables if len(t.columns) == 8)
    veri = [c.text.strip() for c in fatura.rows[2].cells]
    assert veri[2] == "HAZIR BETON"                              # cins yine doğru yerde
    assert veri[7] == ""                                         # proto 'SIZINTI' temizlendi
    assert any("sütun" in m.lower() for m, t in uyarilar)        # sütun sayısı uyarısı verildi


def test_ymm_yazi_vkn_indekslenir(tmp_path):
    _ymm_yazi_docx_yaz(tmp_path / "oyak.docx", "OYAK ÇİMENTO A.Ş.",
                       "ANKARA KURUMLAR V.D. – 6120050961")
    idx = exay.sablonlari_indeksle(str(tmp_path))
    assert idx["6120050961"][0].endswith("oyak.docx")          # YMM yazısı da indekslenir


def test_birlesik_ymm_yazi_bloklara_ayrilir(tmp_path):
    """Tek dosyada birden çok YMM yazısı → 'Konu: Bilgi İsteme' başlığından bölünür."""
    import docx
    from copy import deepcopy
    from docx.oxml.ns import qn
    tek = tmp_path / "tek.docx"
    _ymm_yazi_docx_yaz(tek, "OYAK ÇİMENTO A.Ş.", "ANKARA KURUMLAR V.D. – 6120050961")
    d = docx.Document(str(tek))
    body = d.element.body
    els = [e for e in body if e.tag in (qn('w:p'), qn('w:tbl'))]
    sect = body.find(qn('w:sectPr'))
    for e in els:
        yeni = deepcopy(e)
        body.insert(list(body).index(sect), yeni) if sect is not None else body.append(yeni)
    bloklar = exay._docx_firma_bloklari(d)
    assert len(bloklar) == 2
    assert all(b["vkn"] == "6120050961" for b in bloklar)


def test_fatura_son_sutun_dahil():
    # Tutanak son sütunu 'Defter Kayıt' → dahil DEĞİL (boş kalır)
    assert exay._fatura_son_sutun_dahil(
        ["FATURANIN", "MALIN", "Defter Kayıt", "Tarihi", "Numarası",
         "Cinsi", "Miktarı", "Tutarı", "KDV Tutarı", "Tarihi/Nosu"]) is False
    # YMM yazısı son sütunu 'KDV dahil toplam' → dahil (matrah+kdv)
    assert exay._fatura_son_sutun_dahil(
        ["FAT.TARİHİ", "FAT. NOSU", "MALIN CİNSİ", "MALIN MİKTARI",
         "MATRAH", "KDV", "kdv dahİl toplam"]) is True


def test_word_dosya_adi_ilk_uc_kelime_donem():
    # Tutanak: ilk üç kelime + AA-YYYY (VKN yok, numara başta)
    assert exay._word_tutanak_adi(
        1, "İNNOVA MİMARLIK AHŞAP MOB. İNŞ. SAN. TİC. LTD", "07.2026", ".docx"
    ) == "1) İNNOVA MİMARLIK AHŞAP 07-2026.docx"
    # YMM yazısı: başına 'YMM '
    assert exay._word_tutanak_adi(
        1, "OYAK ÇİMENTO FABRİKALARI ANONİM ŞİRKETİ", "07.2026", ".docx", ymm=True
    ) == "1) YMM OYAK ÇİMENTO FABRİKALARI 07-2026.docx"
    # Boş şablon eki
    assert exay._word_tutanak_adi(3, "X A.Ş.", "12.2025", ".docx", ek="BOŞ") \
        == "3) X A.Ş. 12-2025 BOŞ.docx"
    # İki kelimelik ünvan da sorunsuz
    assert exay._ilk_uc_kelime("FIRMA A") == "FIRMA A"


def test_word_adi_ymm_onek_uctan_uca(tmp_path):
    """YMM yazısı şablonu eşleşen firmanın Word adı 'YMM ' ile başlamalı;
    tutanak şablonununki başlamamalı."""
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)                                    # FIRMA A = 1000000001
    sk = tmp_path / "sablonlar"; sk.mkdir()
    _ymm_yazi_docx_yaz(sk / "a_ymm.docx", "FIRMA A", "V.D. – 1000000001")
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        sablon_klasor=str(sk), cikti_turu='word')
    adlar = [p.name for p in (tmp_path / "Hazır Tutanaklar").glob("*.docx")]
    assert any(n.startswith("1) YMM FIRMA A") for n in adlar), adlar


def test_ymm_yazi_sayi_basligi_silinmez(tmp_path):
    """YMM yazısında 'Konu: Bilgi İsteme' satırının ÜSTÜNDEKİ 'Sayı :' başlığı,
    blok izole edilirken düşmemeli (regression: üst kısım siliniyordu)."""
    import docx
    yol = tmp_path / "ymm.docx"
    d = docx.Document()
    d.add_paragraph("Sayı : YMM 27103572/2026-365          GAZİANTEP")
    d.add_paragraph("Konu : Bilgi İsteme                    09.03.2026")
    d.add_paragraph("Sayın, NURULLAH TOSUN")
    t0 = d.add_table(rows=0, cols=2)
    for e, v in [("Hakkında Bilgi İstenilen Mükellefin", "Hakkında Bilgi İstenilen Mükellefin"),
                 ("Ünvanı", "EVYAPAN DEMİR A.Ş."),
                 ("Vergi Dairesi/Nosu", "ŞEHİTKAMİL V.D. – 3830025675")]:
        r = t0.add_row().cells
        r[0].text = e; r[1].text = v
    t2 = d.add_table(rows=1, cols=7)
    for c, v in enumerate(["F.TARİHİ", "F. NOSU", "MALIN CİNSİ", "MALIN MİKTARI",
                           "MATRAH", "KDV", "kdv dahİl toplam"]):
        t2.rows[0].cells[c].text = v
    d.save(yol)

    bloklar = exay._docx_firma_bloklari(docx.Document(str(yol)))
    assert len(bloklar) == 1 and bloklar[0]["ilk"] == 0     # blok belge başından başlar
    doc0 = exay._docx_blok_belgesi(str(yol), 0)
    paras = [p.text for p in doc0.paragraphs]
    assert any(p.strip().startswith("Sayı") for p in paras), "'Sayı :' başlığı silindi"
    assert any("Konu" in p for p in paras)


# ══════════════════════════════════════════════════════════════════════════
#  Çıktı dosyası AÇIK/kilitli iken net Türkçe hata (öneri #1)
# ══════════════════════════════════════════════════════════════════════════
class _KilitliWB:
    """save çağrısı, dosya başka programda açıkmış gibi PermissionError atar."""
    def save(self, yol):
        raise PermissionError(13, "İşlem erişimi reddedildi")


def test_guvenli_kaydet_acik_dosya_net_mesaj(tmp_path):
    with pytest.raises(PermissionError) as ex:
        exay.guvenli_kaydet(_KilitliWB(), str(tmp_path / "1) rapor.xlsx"))
    m = str(ex.value)
    assert "AÇIK" in m and "rapor.xlsx" in m                 # net, dosya adını içerir


def test_guvenli_docx_kaydet_acik_dosya_net_mesaj(tmp_path):
    with pytest.raises(PermissionError) as ex:
        exay._guvenli_docx_kaydet(_KilitliWB(), str(tmp_path / "1) tutanak.docx"))
    assert "AÇIK" in str(ex.value) and "tutanak.docx" in str(ex.value)


def test_guvenli_kaydet_normal_calisir(tmp_path):
    wb = openpyxl.Workbook(); wb.active["A1"] = "x"
    yol = exay.guvenli_kaydet(wb, str(tmp_path / "ok.xlsx"))
    assert yol.endswith("ok.xlsx")                            # kilitsizde normal kaydeder


# ══════════════════════════════════════════════════════════════════════════
#  Veri kalitesi kontrolleri (yalnızca uyarı — seçimi/iş kuralını etkilemez)
# ══════════════════════════════════════════════════════════════════════════
def test_sayi_fmt_binlik_ayrac():
    # Eski naif ayrıştırıcı burada 0/bozuk verirdi; para_deger ile doğru:
    assert exay.sayi_fmt("1.234.567,89") == "1234567,89"   # binlik+ondalık → doğru
    assert exay.sayi_fmt("45927,50") == "45927,50"         # TR ondalık
    assert exay.sayi_fmt(2500) == "2500"                   # tam sayı


def _kalite_df():
    return pd.DataFrame({
        "Satıcının Adı-Soyadı / Ünvanı": ["A LTD", "A A.Ş.", "B LTD", "C LTD"],
        "Satıcının Vergi Kimlik Numarası": ["1000000001", "1000000001", "2000000002", "3"],
        "Alış Faturasının Sıra No'su": ["F1", "", "F3", "F4"],
        "Alış Faturasının Tarihi": ["2026-01-05", "2026-01-06", "boş-tarih", "2026-01-08"],
        "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı": [100.0, -50.0, 200.0, 300.0],
    })


def test_bos_fatura_no_kontrol():
    # Geçerli VKN'li (1000000001) ama No'su boş 1 satır; geçersiz VKN "3" sayılmaz
    assert exay.bos_fatura_no_kontrol(_kalite_df()) == 1


def test_vkn_unvan_tutarsizligi():
    t = exay.vkn_unvan_tutarsizligi(_kalite_df())
    assert len(t) == 1 and t[0][0] == "1000000001"
    assert set(t[0][1]) == {"A LTD", "A A.Ş."}       # aynı VKN, iki farklı ünvan


def test_negatif_tutar_kontrol():
    adet, toplam = exay.negatif_tutar_kontrol(_kalite_df())
    assert adet == 1 and abs(toplam - (-50.0)) < 1e-9


def test_ayristirilamayan_tarih_kontrol():
    # "boş-tarih" ayrıştırılamaz → 1
    assert exay.ayristirilamayan_tarih_kontrol(_kalite_df()) == 1


def test_kalite_kontrolleri_secimi_etkilemez():
    """Kontroller yalnızca uyarır: geçersiz VKN'li satır yine paydada, seçim değişmez."""
    df = _kalite_df()
    sec, gecersiz = exay.firmalari_filtrele(df, 150, 250, 80, _sessiz)
    # "3" geçersiz → gecersiz'e düşer; kalite fonksiyonları df'i değiştirmemeli
    assert len(gecersiz) == 1
    assert exay.negatif_tutar_kontrol(df)[0] == 1    # df hâlâ negatif satırı içeriyor


# ══════════════════════════════════════════════════════════════════════════
#  Sütun bulma sağlamlığı: başlık varyasyonları + sondaki 'ı' tuzağı
# ══════════════════════════════════════════════════════════════════════════
def test_sutun_bul_tutar_sondaki_i_tuzagi():
    # "KDV Hariç Tutar" (sonda ı YOK) da bulunmalı — eski 'tutarı' terimi kaçırırdı
    assert exay.sutun_bul(["KDV Hariç Tutar"], exay.ARA_MATRAH) == "KDV Hariç Tutar"
    assert exay.sutun_bul(["KDV Hariç Tutarı"], exay.ARA_MATRAH) == "KDV Hariç Tutarı"
    # KDV (tutarı) sütununu matrah SANMAMALI
    assert exay.sutun_bul(["KDV Tutarı"], exay.ARA_MATRAH) is None


def test_sutun_bul_vkn_sinonimleri():
    for baslik in ["Satıcının Vergi Kimlik Numarası", "TCKN", "Vergi No",
                   "Vergi Numarası", "VKN/TCKN"]:
        assert exay.sutun_bul([baslik], exay.ARA_VKN) == baslik, baslik


def test_sutun_bul_fatura_no_ve_unvan_varyasyon():
    assert exay.sutun_bul(["Fatura Numarası"], exay.ARA_FATNO) == "Fatura Numarası"
    assert exay.sutun_bul(["Ünvan"], exay.ARA_UNVAN) == "Ünvan"
    assert exay.sutun_bul(["Unvanı"], exay.ARA_UNVAN) == "Unvanı"


def test_varyant_basliklarla_uctan_uca(tmp_path):
    """Farklı ama geçerli başlıklarla (Vergi No / KDV Hariç Tutar / Fatura Numarası)
    filtreleme sorunsuz çalışmalı."""
    df = pd.DataFrame({
        "Satıcı Unvanı": ["A LTD", "B LTD"],
        "Vergi No": ["1000000001", "2000000002"],
        "Fatura Numarası": ["F1", "F2"],
        "Fatura Tarihi": ["2026-01-05", "2026-01-06"],
        "KDV Hariç Tutar": [500000.0, 100000.0],
        "KDV Tutarı": [100000.0, 20000.0],
    })
    sec, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert "1000000001" in sec and len(gecersiz) == 0     # A tek faturası ≥150K
    # Excel çıktısında matrah/kdv doğru sütunlardan gelmeli (KDV≠matrah)
    kols = list(df.columns)
    assert exay.sutun_bul(kols, exay.ARA_MATRAH) == "KDV Hariç Tutar"
    assert exay.kdv_sutunu_bul(kols) == "KDV Tutarı"


def test_ymm_vergi_dairesi_hesap_nosu_ve_tek_satir_doc():
    """YMM .doc: karşı firma bloğu tek satırda (tab'lı), etiket 'Vergi Dairesi/
    Hesap Nosu', ve blokta 'Telefon/Fax' de var. VKN değer hücresinden gelmeli,
    telefon VKN sanılmamalı (gerçek TIRSAN .doc bu düzendeydi)."""
    metin = (
        "Konu : Bilgi İsteme\n"
        "Hakkında Bilgi İstenilen Mükellefin Altı\t\t"
        "Ünvanı\tTIRSAN TREYLER SAN. VE TİC. A.Ş.\t\t"
        "Vergi Dairesi/Hesap Nosu\tALİ FUAT CEBESOY / 844 005 7150\t\t"
        "Adresi\tADLİYE MAH. 1520 NOLU SOK. NO:3 ARİFİYE / SAKARYA\t\t"
        "Telefon/Fax\t0 264 295 30 00\n"
        "İNCELEME DAYANAĞI\t31.01.2023 Tarih ve 11 Sayılı\n"
    )
    vkn, unvan = exay.sablon_vkn_metinden(metin)
    assert vkn == "8440057150"                 # değer hücresindeki VKN (telefon değil)
    assert "TIRSAN" in unvan


def test_blok_vkn_tek_satir_telefon_atlanir():
    # Her şey tek satırda; telefon hücresi VKN sanılmamalı, V.D. değeri seçilmeli
    blok = ("Ünvanı\tX A.Ş.\tVergi Dairesi/Hesap Nosu\tKADIKÖY / 493 061 9102\t"
            "Telefon/Fax\t0 216 000 00 00")
    assert exay._blok_vkn(blok) == "4930619102"


# ══════════════════════════════════════════════════════════════════════════
#  Büyüteç turu — gerçek dosyalarda karşılaşılabilecek kenar durumlar
# ══════════════════════════════════════════════════════════════════════════
_BUYUK_GIB = ["ALIŞ FATURASININ TARİHİ", "ALIŞ FATURASININ SERİSİ",
              "ALIŞ FATURASININ SIRA NO'SU", "SATICININ ADI-SOYADI / ÜNVANI",
              "SATICININ VERGİ KİMLİK NUMARASI", "ALINAN MAL VE/VEYA HİZMETİN CİNSİ",
              "ALIŞ FATURASININ KDV HARİÇ TUTARI", "KDV'Sİ"]


def test_sutun_bul_turkce_buyuk_harf():
    """'TARİH'.lower() → 'tari̇h' (noktalı i) olduğu için BÜYÜK HARFLİ başlıklar
    bulunamıyor, 'KDV HARİÇ TUTARI' yasağı atlatıp KDV sanılıyordu."""
    K = _BUYUK_GIB
    assert exay.sutun_bul(K, exay.ARA_TARIH) == "ALIŞ FATURASININ TARİHİ"
    assert exay.sutun_bul(K, exay.ARA_VKN) == "SATICININ VERGİ KİMLİK NUMARASI"
    assert exay.sutun_bul(K, exay.ARA_FATNO) == "ALIŞ FATURASININ SIRA NO'SU"
    assert exay.sutun_bul(K, exay.ARA_MATRAH) == "ALIŞ FATURASININ KDV HARİÇ TUTARI"
    assert exay.sutun_bul(K, exay.ARA_UNVAN) == "SATICININ ADI-SOYADI / ÜNVANI"
    assert exay.sutun_bul(K, exay.ARA_CINS) == "ALINAN MAL VE/VEYA HİZMETİN CİNSİ"
    assert exay.kdv_sutunu_bul(K) == "KDV'Sİ"                      # matrah DEĞİL
    assert exay.kdv_sutunu_bul(["KDV HARİÇ TUTARI", "TOPLAM KDV"]) is None
    assert exay.seri_sutunu_bul(K, K[0], K[2]) == "ALIŞ FATURASININ SERİSİ"


def test_buyuk_harf_gib_uctan_uca(tmp_path):
    """Büyük harfli başlıklı liste okunur, başlık bulunur, KDV sütununa KDV yazılır."""
    yol = tmp_path / "NISAN_2026.xlsx"
    pd.DataFrame([
        ["2026-04-01", "A", "F1", "FIRMA A", "1000000001", "mal", 200000, 40000],
        ["2026-04-02", "A", "F2", "FIRMA B", "1000000002", "mal", 10000, 2000],
    ], columns=_BUYUK_GIB).to_excel(yol, index=False)
    df = exay.ana_listeyi_oku(str(yol))
    sec, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert "1000000001" in sec
    cikti = tmp_path / "a.xlsx"
    exay.firma_excel_olustur(sec["1000000001"][0], str(cikti), list(df.columns))
    ws = openpyxl.load_workbook(cikti).active
    assert ws.cell(2, 4).value == pytest.approx(200000)             # matrah
    assert ws.cell(2, 5).value == pytest.approx(40000)              # KDV (matrah değil)


def test_muhasebe_buyuk_harf_basliklar():
    df = pd.DataFrame({"HESAP KODU": ["191.01"], "TARİH": ["2026-01-10"],
                       "FATURA NO": ["F1"], "VERGİ KİMLİK NO": ["1000000001"],
                       "AÇIKLAMA": ["FIRMA A"], "BORÇ": [36000], "MATRAH": [200000]})
    d2 = exay._muhasebe_tipini_esle(df)
    assert "Alış Faturasının Tarihi" in d2.columns
    assert "Satıcının Adı-Soyadı / Ünvanı" in d2.columns             # AÇIKLAMA → ünvan
    assert "KDV'si" in d2.columns                                   # BORÇ → KDV


@pytest.mark.parametrize("girdi,beklenen", [
    ("1.234.567", 1234567.0),        # TR binlik, ondalıksız (eskiden None → 0)
    ("1,234,567", 1234567.0),        # EN binlik, ondalıksız
    ("1\xa0234,56", 1234.56),        # bölünmez boşluk binlik ayracı
    ("1.234,56-", -1234.56),         # muhasebe: sondaki eksi
    ("(1.234,56)", -1234.56),        # muhasebe: parantezli eksi
    ("-1.234,56", -1234.56),
    ("12.345", 12.35),               # tek nokta belirsiz → ondalık (dokunulmaz)
    ("1.23.4", None),                # bozuk gruplama → tahmin edilmez
    ("nan", None),                   # metin 'nan' toplamı NaN yapmasın
    ("-", None),
])
def test_para_deger_kenar_durumlar(girdi, beklenen):
    assert exay.para_deger(girdi) == beklenen


class _SahteZaman(exay.datetime):
    """donem_bul'un 'bugün'ünü sabitlemek için (Ocak 2027)."""
    @classmethod
    def now(cls, tz=None):
        return cls(2027, 1, 20)


def test_donem_bul_yil_sinirli_ve_gecen_yil(monkeypatch):
    # Dosya adındaki VKN'nin içinden yıl (2001) kapılmamalı
    assert exay.donem_bul("KDV 3920012345 NISAN 2026") == "04.2026"
    monkeypatch.setattr(exay, "datetime", _SahteZaman)
    # ARALIK listesi OCAK 2027'de işleniyor → 12.2026 (12.2027 değil)
    assert exay.donem_bul("ARALIK") == "12.2026"
    assert exay.donem_bul("OCAK") == "01.2027"
    # Veri varsa yıl verideki o aydan alınır
    df = pd.DataFrame({"Alış Faturasının Tarihi": ["15.12.2025", "20.12.2025"]})
    assert exay.donem_bul("ARALIK_liste", df) == "12.2025"


def _toplamli_liste_yaz(yol, toplam_tutar):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["İNDİRİLECEK KDV LİSTESİ"])
    ws.append([])
    ws.append(["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
               "Satıcının Adı-Soyadı / Ünvanı", "Satıcının Vergi Kimlik Numarası",
               "Alış Faturasının KDV Hariç Tutarı", "KDV si"])
    ws.append(["01.04.2026", "A1", "FİRMA A", "1234567890", 200000, 40000])
    ws.append(["02.04.2026", "B1", "FİRMA B", "2234567890", 50000, 10000])
    ws.append(["03.04.2026", "C1", "FİRMA C", "3234567890", 30000, 6000])
    ws.append(["04.04.2026", "X1", "VKN'Sİ YOK", "", 20000, 4000])  # gerçek geçersiz fatura
    ws.append([None, None, "GENEL TOPLAM", None, toplam_tutar, 60000])
    wb.save(yol)


def test_toplam_satiri_payda_cift_sayilmaz(tmp_path):
    """VKN'siz 'GENEL TOPLAM' satırı fatura değildir: paydaya eklenirse liste
    toplamı ikiye katlanıyor, %80 hiç tutmuyordu."""
    yol = tmp_path / "t.xlsx"
    _toplamli_liste_yaz(yol, 300000)                  # = üstteki 4 faturanın toplamı
    df = exay.ana_listeyi_oku(str(yol))
    assert len(df) == 4
    assert df.attrs["toplam_satirlari"] == [("GENEL TOPLAM", 300000.0)]
    sec, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    # Geçersiz VKN'li GERÇEK fatura (20.000) paydada KALIR (§2 kuralı korunur)
    assert len(gecersiz) == 1
    wb, kapsam = exay.ozet_rapor_olustur(df, sec, gecersiz, 150000, 450000, 80,
                                         "04.2026", len(sec), 0)
    assert kapsam == pytest.approx(250000 / 300000 * 100, abs=0.1)   # A+B, payda 300K


def test_toplam_satiri_tutar_tutmazsa_kalir(tmp_path):
    """Tutarı üstteki faturaların toplamına eşit değilse satır AYIKLANMAZ
    (yanlışlıkla fatura silinmesin — geçersiz satır olarak paydada kalır)."""
    yol = tmp_path / "t.xlsx"
    _toplamli_liste_yaz(yol, 123456)
    df = exay.ana_listeyi_oku(str(yol))
    assert len(df) == 5 and df.attrs["toplam_satirlari"] == []


def test_csv_unvan_satiri_bos_satir_ve_gec_kodlama(tmp_path):
    """Başlık üstü unvan satırı + boş satırlar + 8 KB'tan sonra gelen cp1254
    karakteri + satır sonu ayraçları + ';' dosyada ondalık virgüller."""
    yol = tmp_path / "NISAN_2026.csv"
    satirlar = ["İNDİRİLECEK KDV LİSTESİ", "", "",
                "Alış Faturasının Tarihi;Alış Faturasının Sıra No'su;Satıcının Adı-Soyadı / Ünvanı;"
                "Satıcının Vergi Kimlik Numarası;Alış Faturasının KDV Hariç Tutarı;KDV si"]
    for i in range(300):   # ~9 KB ASCII veri
        satirlar.append(f"01.04.2026;F{i};FIRMA {i};{1000000100 + i};1.000,00;200,00;")
    satirlar.append("02.04.2026;SON;ÇAĞRI ŞİRKETİ;1000009999;500.000,00;100.000,00;")
    yol.write_bytes(("\r\n".join(satirlar) + "\r\n").encode("cp1254"))
    df = exay.ana_listeyi_oku(str(yol))
    assert len(df) == 301
    assert "Satıcının Vergi Kimlik Numarası" in df.columns        # sütunlar kaymadı
    assert df.iloc[-1]["Satıcının Adı-Soyadı / Ünvanı"] == "ÇAĞRI ŞİRKETİ"
    sec, _ = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    assert "1000009999" in sec                                     # 500.000,00 doğru okundu


def test_excel_unicode_metin_utf16_tab(tmp_path):
    """Excel 'Unicode Metin (*.txt)' kaydı UTF-16 + sekme ayraçlıdır."""
    yol = tmp_path / "NISAN_2026.txt"
    icerik = ("Alış Faturasının Tarihi\tSatıcının Adı-Soyadı / Ünvanı\t"
              "Satıcının Vergi Kimlik Numarası\tAlış Faturasının KDV Hariç Tutarı\tKDV si\r\n"
              "01.04.2026\tŞİŞECAM\t1000000001\t200000,00\t40000,00\r\n")
    yol.write_bytes(icerik.encode("utf-16"))
    df = exay.ana_listeyi_oku(str(yol))
    assert df.iloc[0]["Satıcının Adı-Soyadı / Ünvanı"] == "ŞİŞECAM"
    assert exay.para_deger(df.iloc[0]["KDV si"]) == 40000.0


def test_baslik_ustu_vergi_kimlik_satiri_baslik_sanilmaz(tmp_path):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Mükellefin Vergi Kimlik No: 9999999990"])          # tek anahtar
    ws.append(["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
               "Satıcının Adı-Soyadı / Ünvanı", "Satıcının Vergi Kimlik Numarası",
               "Alış Faturasının KDV Hariç Tutarı", "KDV si"])
    ws.append(["01.04.2026", "A1", "FIRMA A", "1234567890", 200000, 40000])
    yol = tmp_path / "t.xlsx"; wb.save(yol)
    df = exay.ana_listeyi_oku(str(yol))
    assert "Satıcının Vergi Kimlik Numarası" in df.columns and len(df) == 1


def test_veri_satiri_aciklamasi_baslik_sanilmaz(tmp_path):
    """Muhasebe başlığında tek anahtar var ('vergi kimlik'); açıklaması iki anahtar
    içeren bir VERİ satırı başlığın önüne geçmemeli."""
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Hesap Kodu", "Tarih", "Fatura No", "Vergi Kimlik No", "Açıklama", "Borç", "Matrah"])
    ws.append(["191.01", "2026-01-10", "F1", "1000000001",
               "ALIŞ FATURASI KDV HARİÇ BEDEL", 36000, 200000])
    yol = tmp_path / "m.xlsx"; wb.save(yol)
    df = exay.ana_listeyi_oku(str(yol))
    assert len(df) == 1 and exay.kdv_sutunu_bul(list(df.columns)) == "KDV'si"


def test_coklu_sayfa_liste_sayfasi_secilir(tmp_path):
    yol = tmp_path / "t.xlsx"
    with pd.ExcelWriter(yol) as w:
        pd.DataFrame({"Bilgi": ["Kapak sayfası"]}).to_excel(w, sheet_name="Kapak", index=False)
        pd.DataFrame({"Alış Faturasının Tarihi": ["01.04.2026"],
                      "Satıcının Vergi Kimlik Numarası": ["1234567890"],
                      "Alış Faturasının KDV Hariç Tutarı": [200000]}).to_excel(
            w, sheet_name="Liste", index=False)
    df = exay.ana_listeyi_oku(str(yol))
    assert df.attrs["sayfa"] == "Liste" and len(df) == 1
    assert df.attrs["diger_sayfalar"] == ["Kapak"]


def test_bos_liste_ve_tutarsiz_liste_net_hata(tmp_path):
    yol = tmp_path / "bos.csv"; yol.write_text("", encoding="utf-8")
    with pytest.raises(ValueError, match="boş"):
        exay.ana_listeyi_oku(str(yol))
    yol2 = tmp_path / "baslik.xlsx"
    pd.DataFrame(columns=["Alış Faturasının Tarihi", "Satıcının Vergi Kimlik Numarası"]
                 ).to_excel(yol2, index=False)
    with pytest.raises(ValueError):
        exay.ana_listeyi_oku(str(yol2))
    df = pd.DataFrame({"Satıcının Vergi Kimlik Numarası": ["1000000001"],
                       "Açıklama": ["x"]})
    with pytest.raises(ValueError, match="Tutar"):
        exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)


def test_eski_cikti_yalniz_word_olsa_da_tasinir(tmp_path):
    yol = tmp_path / "NISAN_2026.xlsx"
    _yeni_gib_yaz(yol)
    eski = tmp_path / "Hazır Tutanaklar"; eski.mkdir()
    (eski / "1) ESKI FIRMA 04-2026.docx").write_bytes(b"eski")
    exay.dosyalari_isle(str(yol), 150000, 450000, 80, _sessiz, lambda *a: None,
                        cikti_turu='excel')
    yedekler = [p for p in tmp_path.iterdir() if p.name.startswith("Hazır Tutanaklar_")]
    assert len(yedekler) == 1 and (yedekler[0] / "1) ESKI FIRMA 04-2026.docx").exists()
    assert not (eski / "1) ESKI FIRMA 04-2026.docx").exists()      # yeni klasör temiz


def test_bos_klasor_adi_cakismaz(tmp_path):
    (tmp_path / "K").mkdir(); (tmp_path / "K_2").mkdir()
    assert exay._bos_klasor_adi(tmp_path / "K").name == "K_3"
    assert exay._bos_klasor_adi(tmp_path / "Y").name == "Y"


def test_sablon_taramasi_cikti_klasorunu_atlar_ve_vknsizi_bildirir(tmp_path):
    _docx_sablon_yaz(tmp_path / "gercek.docx", "GERÇEK A.Ş.", "KADIKÖY / 1234567890")
    cikti = tmp_path / "Hazır Tutanaklar_20260101_101010"; cikti.mkdir()
    _docx_sablon_yaz(cikti / "1) DOLU.docx", "DOLU A.Ş.", "KADIKÖY / 5555555550")
    import docx
    docx.Document().save(tmp_path / "ustyazi.DOCX")                 # VKN'siz, büyük uzantı
    loglar = []
    idx = exay.sablonlari_indeksle(str(tmp_path), lambda m, t='': loglar.append(m))
    assert set(idx) == {"1234567890"}                               # çıktı klasörü atlandı
    metin = "\n".join(loglar)
    assert "ustyazi.DOCX" in metin and "Hazır Tutanaklar" in metin


def test_excel_esittir_ile_baslayan_metin_formul_olmaz(tmp_path):
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su",
            "Alınan Mal ve/veya Hizmetin Cinsi", "Alış Faturasının KDV Hariç Tutarı", "KDV si"]
    df = pd.DataFrame([["2026-04-01", "F1", "=KDV iadesi", 1000, 200]], columns=kols)
    yol = tmp_path / "f.xlsx"
    exay.firma_excel_olustur(df, str(yol), kols)
    h = openpyxl.load_workbook(yol).active.cell(2, 9)
    assert h.value == "=KDV iadesi" and h.data_type == "s"


def test_dosya_adi_temizle_satir_sonu():
    assert exay.dosya_adi_temizle("ABC\nLTD  ŞTİ") == "ABC LTD ŞTİ"
    assert exay.dosya_adi_temizle('A/B:C') == "A_B_C"


def _resimli_docx(yol, renk):
    import docx
    from PIL import Image
    png = yol.with_suffix(".png")
    Image.new("RGB", (8, 8), renk).save(png)
    d = docx.Document()
    d.add_paragraph(f"Belge {renk}")
    d.add_picture(str(png))
    d.save(yol)
    return docx.Document(yol)


def test_tek_docx_birlestirme_resim_iliskileri(tmp_path):
    """Farklı şablonlardan gelen resimli gövdeler birleşince rId'ler hedefte
    doğru resmi göstermeli (eskiden kaynak rId'si kalıyor, belge bozuluyordu)."""
    import docx
    from docx.oxml.ns import qn
    d1 = _resimli_docx(tmp_path / "a.docx", "red")
    d2 = _resimli_docx(tmp_path / "b.docx", "blue")
    yol = exay.firmalar_tek_docx([d1, d2], str(tmp_path / "birlesik.docx"))
    m = docx.Document(yol)
    blips = m.element.body.findall('.//' + qn('a:blip'))
    rids = [b.get(qn('r:embed')) for b in blips]
    assert len(rids) == 2 and len(set(rids)) == 2
    bloblar = [m.part.related_parts[r].blob for r in rids]
    assert bloblar[0] != bloblar[1]                                 # iki ayrı resim


def test_mukerrer_fatura_vkn_normalize():
    df = pd.DataFrame({"Satıcının Vergi Kimlik Numarası": ["71419747", "0071419747"],
                       "Alış Faturasının Sıra No'su": ["F1", "F1"]})
    assert exay.mukerrer_fatura_bul(df) == [("0071419747", "F1", 2)]


def test_gui_firma_yok_hata_sayilmaz():
    from unittest.mock import MagicMock
    root = MagicMock(name="root")
    root.after.side_effect = lambda _ms, f: f()                     # hemen çalıştır
    app = exay.KDVBolmeApp(root)
    app.birak_yazi = MagicMock(); app.durum_lbl = MagicMock()
    app._tamam(None, 0, 0, {})
    assert "firma yok" in app.birak_yazi.config.call_args.kwargs["text"]
    app._tamam(None, 0, 1, {})
    assert "Hata" in app.birak_yazi.config.call_args.kwargs["text"]


def _gib_yeni_bicim_yaz(yol, toplam_satiri=True, toplam=None):
    """GİB 'İndirilecek KDV listesi yeni formatı' düzeni (gerçek dosyadan): üstte 3 boş
    satır, A sütunu boş, 'Sıra No' sütunu, 16 sütun ve EN ALTTA etiketsiz toplam satırı
    (yalnızca tutar sütunlarında toplam; sıra no/tarih/ünvan/VKN boş)."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "İndirilecek KDV Listesi"
    bas = ["Sıra No", "Alış Faturasının Tarihi", "Alış Faturasının Serisi",
           "Alış Faturasının Sıra No'su", "Satıcının Adı-Soyadı / Ünvanı",
           "Satıcının Vergi Kimlik Numarası / TC Kimlik Numarası",
           "Alınan Mal ve/veya Hizmetin Cinsi", "Alınan Mal ve/veya Hizmetin Miktarı",
           "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si",
           "Tevkifata Tabi Olmayan Ve Bu Dönemde İndirilen Kdv Tutarı",
           "2 Nolu Beyannamede Ödenen Kdv Tutarı", "Toplam İndirilecek KDV Tutarı"]
    for c, b in enumerate(bas, 2):
        ws.cell(4, c, value=b)
    satirlar = [
        ("ATAKAŞ ÇELİK", "0950303776", 1350916.0),     # tek fatura ≥150K
        ("AY PROFİL",    "1060068320", 650007.32),
        ("KÜÇÜK A",      "1000000001", 40000.0),
        ("KÜÇÜK B",      "1000000002", 30000.0),
        ("KÜÇÜK C",      "1000000003", 20000.0),
        ("RHEINZINK GMBH", "1111111111", 60000.0),      # yabancı → geçersiz VKN, paydada KALIR
    ]
    for i, (u, v, t) in enumerate(satirlar, 1):
        ws.append([None, i, exay.datetime(2026, 8, i), None, f"F{i}", u, v, "MAL", "1 AD",
                   t, round(t * 0.2, 2), round(t * 0.2, 2), 0, round(t * 0.2, 2)])
    if toplam_satiri:
        top = sum(t for *_, t in satirlar) if toplam is None else toplam
        ws.append([None] * 9 + [top, round(top * 0.2, 2), None, None, round(top * 0.2, 2)])
    wb.save(yol)
    return sum(t for *_, t in satirlar)


def test_etiketsiz_toplam_satiri_tum_firmalari_sectirmez(tmp_path):
    """GERÇEK HATA (Ağustos 2026 listesi): GİB yeni biçimli listenin en altındaki
    etiketsiz toplam satırı geçersiz VKN'li fatura sanılıyor, liste toplamı ikiye
    katlanıyor, %80 tutmuyor ve 257 firmanın HEPSİNE tutanak çıkıyordu."""
    yol = tmp_path / "İndirilecek KDV listesi yeni formatı 08.xlsx"
    gercek_toplam = _gib_yeni_bicim_yaz(yol)
    df = exay.ana_listeyi_oku(str(yol))
    assert len(df) == 6
    assert df.attrs["toplam_satirlari"] == [("TOPLAM (etiketsiz)", gercek_toplam)]
    sec, gecersiz = exay.firmalari_filtrele(df, 150000, 450000, 80, _sessiz)
    # Payda = gerçek liste toplamı (RHEINZINK dahil, §2.3); 2 firma %80'i karşılar
    assert set(sec) == {"0950303776", "1060068320"}
    assert len(gecersiz) == 1                                  # yalnız RHEINZINK


def test_etiketsiz_satir_tutar_tutmazsa_kalir(tmp_path):
    """Kimliksiz satırın tutarı üstteki faturaların toplamına eşit değilse
    ayıklanmaz (gerçek bir eksik kayıt olabilir — paydada kalır)."""
    yol = tmp_path / "l.xlsx"
    _gib_yeni_bicim_yaz(yol, toplam=999.0)
    df = exay.ana_listeyi_oku(str(yol))
    assert len(df) == 7 and df.attrs["toplam_satirlari"] == []


@pytest.mark.parametrize("girdi,beklenen", [
    ("150000", 150000.0), ("150.000", 150000.0), ("150,000", 150000.0),
    ("150.000,00", 150000.0), ("150 000 ₺", 150000.0), ("1.500.000", 1500000.0),
    ("1500,50", 1500.5),
])
def test_kriter_tutari_oku(girdi, beklenen):
    assert exay.kriter_tutari_oku(girdi) == beklenen


def test_kriter_tutari_oku_hatali():
    with pytest.raises(ValueError):
        exay.kriter_tutari_oku("yüz elli bin")


def _ymm_birlesik_bosluklu_yaz(yol, firmalar, bos_satir=30):
    """Gerçek 'YMM 08.2026' düzeni: her mektup 'Sayı:'/'Konu: Bilgi İsteme' ile başlar,
    fatura tablosu 6 sütundur (F.TARİHİ…KDV) ve mektuplar ~30 BOŞ paragrafla ayrılır."""
    import docx
    d = docx.Document()
    for unvan, vd in firmalar:
        d.add_paragraph("Sayı: YMM 27103572/2026-2209          GAZİANTEP")
        d.add_paragraph("Konu: Bilgi İsteme          30.09.2026")
        t = d.add_table(rows=0, cols=2)
        for a, b in [("İADE TALEBİNDE BULUNAN FİRMANIN", "İADE TALEBİNDE BULUNAN FİRMANIN"),
                     ("Ünvanı", "ESKA METAL SAN. TİC. A.Ş."),
                     ("Vergi Dairesi/Nosu", "ŞEHİTKAMİL / 380 119 8516"),
                     ("HAKKINDA BİLGİ İSTENİLEN MÜKELLEFİN", "HAKKINDA BİLGİ İSTENİLEN MÜKELLEFİN"),
                     ("Ünvanı", unvan), ("Vergi Dairesi/Nosu", vd),
                     ("İNCELEME DAYANAĞI", "04.03.2026 Tarih ve 46 Sayılı")]:
            r = t.add_row().cells; r[0].text = a; r[1].text = b
        f = d.add_table(rows=2, cols=6)
        for c, v in enumerate(["F.TARİHİ", "F. NOSU", "MALIN CİNSİ", "MALIN MİKTARI", "MATRAH", "KDV"]):
            f.rows[0].cells[c].text = v
        d.add_paragraph("SABRİ HAMAMCI")
        for _ in range(bos_satir):
            d.add_paragraph("")
    d.save(yol)


def test_ymm_birlesik_blok_sondaki_bos_satirlar_silinir_ve_6_sutun_uyarmaz(tmp_path):
    """GERÇEK (YMM 08.2026): mektuplar 30 boş satırla ayrılıyor → izole edilen her
    yazıda boş 2. sayfa çıkıyordu; ayrıca 6 sütunlu (gerçek) YMM tablosu için her
    seferinde yanlış 'beklenen 7 sütun' uyarısı veriliyordu."""
    import docx
    from docx.oxml.ns import qn
    yol = tmp_path / "YMM 08.2026.docx"
    _ymm_birlesik_bosluklu_yaz(yol, [("ATAKAŞ ÇELİK SAN. VE TİC. A.Ş.", "DÖRTYOL V.D. / 095 030 3776"),
                                     ("TEZCAN GALVANİZLİ A.Ş.", "BEYKOZ V.D. / 841 005 2600")])
    kay = exay._sablon_kayitlari(str(yol))
    assert [k[0] for k in kay] == ["0950303776", "8410052600"]
    d = exay._docx_blok_belgesi(str(yol), 0)
    from docx.text.paragraph import Paragraph
    son = [el for el in d.element.body if el.tag in (qn('w:p'), qn('w:tbl'))][-1]
    assert Paragraph(son, d).text.strip() == "SABRİ HAMAMCI"         # sonda boş satır yok
    uyarilar = []
    out = tmp_path / "o.docx"
    firma, kols = _ornek_firma_df()
    exay.firma_docx_olustur(str(yol), firma, str(out), kols,
                            log_cb=lambda m, t='': uyarilar.append(m), blok=0)
    assert not [u for u in uyarilar if "sütun" in u]                  # 6 sütun geçerli
    tablo = docx.Document(out).tables[-1]
    assert len(tablo.columns) == 6 and len(tablo.rows) > 1


def test_birlesik_doc_word_yoksa_net_mesaj(tmp_path, monkeypatch):
    """Birleşik .doc'u bölmek Word ister; Word yoksa 'bozuk/şifreli' değil
    'Word gerekli' denmeli."""
    p = tmp_path / "YMM 08.2026.doc"; p.write_bytes(b"stub")
    iki = ("Sayı: YMM 1\nKonu: Bilgi İsteme\nHakkında Bilgi İstenilen Mükellefin\tÜnvanı\tA A.Ş.\t"
           "Vergi Dairesi/Nosu\tX V.D. / 123 456 7890\n"
           "Sayı: YMM 2\nKonu: Bilgi İsteme\nHakkında Bilgi İstenilen Mükellefin\tÜnvanı\tB A.Ş.\t"
           "Vergi Dairesi/Nosu\tY V.D. / 223 456 7890\n")
    monkeypatch.setattr(exay, "_doc_metni_oku", lambda path: iki)
    monkeypatch.setattr(exay, "word_destekli", lambda: False)
    loglar = []
    exay.sablonlari_indeksle(str(tmp_path), lambda m, t='': loglar.append(m))
    metin = "\n".join(loglar)
    assert "Word kurulu olmalı" in metin and "bozuk" not in metin


# ── Sahte Word (COM) — .doc yolu bu ortamda Word olmadan test edilebilsin ──
class _SahteFont:
    def __init__(self, hucreler): self._h = hucreler
    @property
    def Bold(self): return all(h.kalin for h in self._h)
    @Bold.setter
    def Bold(self, v):
        for h in self._h: h.kalin = bool(v)

class _SahteAralik:
    def __init__(self, hucreler): self._h = hucreler
    @property
    def Text(self): return self._h[0].metin
    @Text.setter
    def Text(self, v): self._h[0].metin = v
    @property
    def Font(self): return _SahteFont(self._h)

class _SahteHucre:
    def __init__(self, metin, kalin): self.metin, self.kalin = metin, kalin
    @property
    def Range(self): return _SahteAralik([self])

class _SahteHucreler(list):
    @property
    def Count(self): return len(self)
    def __call__(self, i): return self[i - 1]

class _SahteSatir:
    def __init__(self, tablo, metinler, kalin):
        self._t = tablo; self._h = [_SahteHucre(m, kalin) for m in metinler]
    @property
    def Cells(self): return _SahteHucreler(self._h)
    @property
    def Range(self):
        a = _SahteAralik(self._h); return a
    def Delete(self): self._t.satirlar.remove(self)

class _SahteSatirlar:
    def __init__(self, tablo): self._t = tablo
    @property
    def Count(self): return len(self._t.satirlar)
    def __call__(self, i): return self._t.satirlar[i - 1]
    def Add(self):
        # Gerçek Word gibi: yeni satır SON satırın biçimini (kalınlık) kopyalar
        son = self._t.satirlar[-1]
        yeni = _SahteSatir(self._t, [''] * len(son._h), son._h[0].kalin)
        self._t.satirlar.append(yeni); return yeni

class _SahteTablo:
    def __init__(self, satirlar):
        self.satirlar = []
        for metinler, kalin in satirlar:
            self.satirlar.append(_SahteSatir(self, metinler, kalin))
    @property
    def Rows(self): return _SahteSatirlar(self)

def _sahte_word_kur(monkeypatch, tablo):
    import sys, types
    from unittest.mock import MagicMock
    belge = MagicMock(); belge.Tables = [tablo]
    word = MagicMock(); word.Documents.Open.return_value = belge
    paket = types.ModuleType("win32com"); istemci = types.ModuleType("win32com.client")
    istemci.DispatchEx = lambda ad: word; istemci.constants = MagicMock()
    paket.client = istemci
    monkeypatch.setitem(sys.modules, "win32com", paket)
    monkeypatch.setitem(sys.modules, "win32com.client", istemci)
    return belge

_KIT_BASLIK = [(["FATURANIN", "FATURANIN", "MALIN", "MALIN", "MALIN", "MALIN", "Defter Kayıt"], True),
               (["Tarihi", "Numarası", "Cinsi", "Miktarı", "Tutarı", "KDV Tutarı", "Tarihi/Nosu"], True)]

def _uc_fatura():
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su", "Alınan Mal ve/veya Hizmetin Cinsi",
            "Alınan Mal ve/veya Hizmetin Miktarı", "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"]
    df = pd.DataFrame([["2026-08-03", "CGT1", "JÜT İPLİK", "3255,30 Kg", 198573.3, 19857.33],
                       ["2026-08-05", "CGT2", "JÜT İPLİK", "3461 Kg", 211121.0, 21112.1],
                       ["2026-08-08", "CGT3", "JÜT İPLİK", "3425,80 Kg", 208973.8, 20897.38]], columns=kols)
    return df, kols


@pytest.mark.parametrize("ornek_satir", [True, False])
def test_word_com_fatura_satirlari_kalin_degil(tmp_path, monkeypatch, ornek_satir):
    """GERÇEK HATA (OPUROĞLU GOLD 08-2026.doc): COM yolu tüm veri satırlarını silip
    Rows.Add() ile ekliyordu; Word yeni satıra son kalan KALIN başlık satırının
    biçimini kopyaladığından fatura bilgileri kalın çıkıyordu. Elle hazırlanan
    tutanaklarda fatura satırı hiçbir zaman kalın değildir."""
    satirlar = list(_KIT_BASLIK)
    if ornek_satir:   # şablonda eski firmanın verisiyle normal (kalın olmayan) bir satır
        satirlar.append((["31.03.2026", "CEF51", "Bobin İplik", "1 Adet", "2.811.358,00", "562.271,60", ""], False))
    tablo = _SahteTablo(satirlar)
    _sahte_word_kur(monkeypatch, tablo)
    df, kols = _uc_fatura()
    exay.firma_word_olustur(str(tmp_path / "sablon.doc"), df, str(tmp_path / "cikti.doc"), kols)
    veri = tablo.satirlar[2:]
    assert len(veri) == 3                                            # eski veri gitti, 3 fatura
    assert [s._h[1].metin for s in veri] == ["CGT1", "CGT2", "CGT3"]
    assert not any(h.kalin for s in veri for h in s._h)              # hiçbiri kalın değil
    assert all(h.kalin for s in tablo.satirlar[:2] for h in s._h)    # başlıklar kalın kaldı
    assert "Bobin" not in " ".join(h.metin for s in veri for h in s._h)


def test_docx_fatura_satirlari_kalin_degil(tmp_path):
    """.docx yolu: şablonun örnek veri satırı kalın olsa bile fatura bilgisi
    kalın yazılmaz; firma bilgileri (başka tablo) biçimini korur."""
    import docx
    yol = tmp_path / "s.docx"
    _docx_sablon_yaz(yol, "ÖRNEK A.Ş.", "KADIKÖY / 1234567890")
    d = docx.Document(yol)
    t = d.tables[-1]
    for c in t.rows[2].cells:                                        # örnek satırı kalın yap
        c.paragraphs[0].add_run("eski").bold = True
    d.save(yol)
    df, kols = _uc_fatura()
    out = tmp_path / "o.docx"
    exay.firma_docx_olustur(str(yol), df, str(out), kols)
    t2 = docx.Document(out).tables[-1]
    runs = [r for row in t2.rows[2:] for c in row.cells for p in c.paragraphs for r in p.runs if r.text]
    assert runs and not any(r.bold for r in runs)


def _karisik_sirali_faturalar():
    """Liste tutara göre sıralı gelmiş (gerçek vaka: OPUROĞLU GOLD 08-2026)."""
    kols = ["Alış Faturasının Tarihi", "Alış Faturasının Sıra No'su", "Alınan Mal ve/veya Hizmetin Cinsi",
            "Alınan Mal ve/veya Hizmetin Miktarı", "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı", "KDV'si"]
    df = pd.DataFrame([["18.08.2026", "CGT610", "JÜT", "1 KG", 225004.6, 22500.46],
                       ["05.08.2026", "CGT575", "JÜT", "1 KG", 211121.0, 21112.1],
                       ["okunamayan", "CGT999", "JÜT", "1 KG", 200000.0, 20000.0],
                       ["18.08.2026", "CGT609", "JÜT", "1 KG", 88572.0, 8857.2],     # aynı tarih: liste sırası
                       [exay.datetime(2026, 8, 3), "CGT565", "JÜT", "1 KG", 198573.3, 19857.33]],
                      columns=kols)
    return df, kols

_BEKLENEN_SIRA = ["CGT565", "CGT575", "CGT610", "CGT609", "CGT999"]


def test_word_faturalar_tarih_sirasinda_docx(tmp_path):
    import docx
    yol = tmp_path / "s.docx"
    _docx_sablon_yaz(yol, "ÖRNEK A.Ş.", "KADIKÖY / 1234567890")
    df, kols = _karisik_sirali_faturalar()
    out = tmp_path / "o.docx"
    exay.firma_docx_olustur(str(yol), df, str(out), kols)
    t = docx.Document(out).tables[-1]
    assert [r.cells[1].text for r in t.rows[2:]] == _BEKLENEN_SIRA


def test_word_faturalar_tarih_sirasinda_com(tmp_path, monkeypatch):
    tablo = _SahteTablo(list(_KIT_BASLIK))
    _sahte_word_kur(monkeypatch, tablo)
    df, kols = _karisik_sirali_faturalar()
    exay.firma_word_olustur(str(tmp_path / "s.doc"), df, str(tmp_path / "o.doc"), kols)
    assert [s._h[1].metin for s in tablo.satirlar[2:]] == _BEKLENEN_SIRA


def test_excel_tutanak_sirasi_degismez(tmp_path):
    """Sıralama yalnız Word içindir; GİB'e yüklenen Excel listedeki sırayı korur."""
    df, kols = _karisik_sirali_faturalar()
    out = tmp_path / "e.xlsx"
    exay.firma_excel_olustur(df, str(out), kols)
    ws = openpyxl.load_workbook(out).active
    assert [ws.cell(r, 3).value for r in range(2, 7)] == list(df[kols[1]])
