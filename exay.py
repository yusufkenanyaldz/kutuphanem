"""
e-YMM Karşıt İnceleme — KDV Fatura Listesi Bölme Programı v5.0
Kapsam kriterleri:
  1) Tek fatura >= esik_tek  VEYA  toplam >= esik_toplam
  2) Seçilenler listenin %80'ini karşılamıyorsa kalan firmalar
     büyükten küçüğe eklenir (tüm faturalarıyla birlikte)
Seri numarası: tarihten sonraki ayrı sütundan alınır, yoksa boş
Fatura numarası: olduğu gibi (harfler dahil)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading, os, re, sys, json
from datetime import datetime
from pathlib import Path

# Kullanıcının son kullandığı eşik/yüzde değerleri burada saklanır ki program
# her açılışta aynı kriterlerle gelsin. Program Files gibi yazılamayan yerlere
# kurulmuş .exe'de sorun olmaması için kullanıcı ana klasörüne yazılır.
AYAR_YOLU = Path.home() / '.exay_ayarlar.json'


def kaynak_yolu(rel_yol):
    """Hem normal çalışmada hem de PyInstaller ile paketlenmiş .exe içinde
    (logo gibi) yardımcı dosyaların doğru yolunu döndürür."""
    if getattr(sys, 'frozen', False):          # .exe olarak paketlenmişse
        taban = Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent))
    else:                                       # normal .py çalışması
        taban = Path(__file__).parent
    return taban / rel_yol

try:
    import pandas as pd
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable,'-m','pip','install',
                           'pandas','openpyxl','xlrd','--quiet'])
    import pandas as pd, openpyxl

SABLON_SUTUNLAR = [
    "Faturanın Tarihi","Faturanın Serisi","Faturanın Numarası",
    "Faturanın Tutarı (TL)","K.D.V(TL)","Defter Kayıt Tarihi",
    "Yevmiye Numarası","Ödeme Şekli ve Ödemeye İlişkin Bilgiler",
    "Açıklama","Hatalı Satır Açıklama",
]

# ══════════════════════════════════════════
#  YARDIMCI
# ══════════════════════════════════════════
def sutun_bul(kolonlar, aranacak):
    for col in kolonlar:
        cs = str(col).lower().strip()
        for a in aranacak:
            if a.lower() in cs: return col
    return None

def kdv_sutunu_bul(kolonlar):
    """Faturanın KDV'si sütununu bulur. 'KDV'si', 'KDV si', 'KDVsi', 'KDV' gibi
    biçimlerin hepsini yakalar; 'KDV Hariç Tutarı' (matrah), 'Toplam İndirilecek KDV',
    tevkifat / 2 nolu beyanname KDV sütunlarıyla KARIŞMAZ.
    Böylece hem eski (KDV'si) hem yeni (KDV si) liste tipinde doğru sütun seçilir."""
    yasak = ('hariç', 'haric', 'toplam', 'tevkifat', '2 nolu', 'ödenen', 'odenen',
             'indirilecek', 'indirilen', 'dönem', 'donem', 'matrah', 'tutarı', 'tutari')
    # 1) Dar eşleşme: sadeleştirince 'kdvsi' veya 'kdv' olan sütun (asıl KDV'si sütunu)
    for col in kolonlar:
        cs = str(col).lower().strip()
        if 'kdv' not in cs or any(y in cs for y in yasak):
            continue
        temiz = cs.replace("'", '').replace(' ', '').replace('’', '')
        if temiz in ('kdvsi', 'kdv'):
            return col
    # 2) Gevşek son çare: 'kdv' içeren ama yasaklı kelime içermeyen ilk sütun
    for col in kolonlar:
        cs = str(col).lower().strip()
        if 'kdv' in cs and not any(y in cs for y in yasak):
            return col
    return None

def seri_sutunu_bul(kolonlar, tarih_col, faturano_col=None):
    """
    Fatura serisi sütununu bul:
      1) Adında 'seri' geçen bir sütun varsa doğrudan onu kullan.
      2) Yoksa tarih sütununun hemen sağındaki sütunu aday al — ancak bu sütun
         fatura NUMARASI sütununun ta kendisiyse seri sütunu yok demektir; boş bırak.
    Böylece tarihin sağında ayrı bir seri sütunu olmayan listelerde
    fatura numarası yanlışlıkla 'seri' alanına yazılmaz.
    """
    # 1) Adında "seri" geçen sütun (en güvenilir)
    for col in kolonlar:
        if 'seri' in str(col).lower():
            return col
    # 2) Tarihin sağındaki sütun — ama numara sütunu değilse
    if tarih_col is None: return None
    try:
        idx = list(kolonlar).index(tarih_col)
        if idx + 1 < len(kolonlar):
            aday = kolonlar[idx + 1]
            if faturano_col is not None and aday == faturano_col:
                return None   # sağdaki sütun aslında numara → seri yok
            return aday
    except: pass
    return None

def ana_listeyi_oku(dosya):
    ext = Path(dosya).suffix.lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    raw = pd.read_excel(dosya, engine=engine, header=None)

    # Başlık satırını bul — birden fazla anahtar kelimeyle dene
    ANAHTAR = ['alış faturası', 'satıcının', 'vergi kimlik',
               'fatura tarihi', 'faturanın tarihi', 'kdv hariç']
    baslik = None
    for i, row in raw.iterrows():
        v = ' '.join(str(x) for x in row if pd.notna(x)).lower()
        if any(a in v for a in ANAHTAR):
            baslik = i; break

    if baslik is None:
        # Son çare: en fazla dolu hücre içeren satırı başlık say
        dolu = raw.apply(lambda r: r.notna().sum(), axis=1)
        baslik = int(dolu.idxmax())

    df = pd.read_excel(dosya, engine=engine, skiprows=baslik, header=0)
    df = df.dropna(how='all').reset_index(drop=True)

    # Başlık satırı tekrar veri olarak geldiyse at
    if df.iloc[0].astype(str).str.contains(
            'Alış Faturası|Satıcın|Vergi Kimlik|Fatura Tarihi', na=False).any():
        df = df.iloc[1:].reset_index(drop=True)

    # Sütun adlarını normalize et (baştaki/sondaki boşluk, satır sonu)
    df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]

    # Muhasebe (191 hesabı) dökümü tipiyse sütunları standart GİB adlarına çevir
    df = _muhasebe_tipini_esle(df)
    return df


def _muhasebe_tipini_esle(df):
    """Muhasebe programından alınan 191 hesabı dökümünü tanır ve sütunlarını
    standart GİB listesi adlarına çevirir; böylece programın geri kalanı
    hiçbir değişiklik gerekmeden çalışır.
    Tanıma: 'Borç' + 'vergi kimlik' sütunları var, 'Satıcının...' sütunu yok.
    Eşleme: Tarih→Alış Faturasının Tarihi, fatura no→Sıra No'su,
            Açıklama→Satıcı Ünvanı, Borç→KDV'si,
            matrah / KDV'nin sağındaki adsız sayısal sütun→KDV Hariç Tutar."""
    kolonlar = list(df.columns)
    kl = [str(c).lower().strip() for c in kolonlar]
    var = lambda k: any(k in c for c in kl)
    if not (var('borç') or var('borc')):   return df   # muhasebe tipi değil
    if not var('vergi kimlik'):            return df
    if var('satıcının') or var('saticinin'): return df # zaten GİB tipi

    esle = {}
    for orij, cl in zip(kolonlar, kl):
        if   cl == 'tarih':                 esle[orij] = 'Alış Faturasının Tarihi'
        elif cl.startswith('fatura no'):    esle[orij] = "Alış Faturasının Sıra No'su"
        elif cl in ('açıklama', 'aciklama'):esle[orij] = 'Satıcının Adı-Soyadı / Ünvanı'
        elif cl in ('borç', 'borc'):        esle[orij] = "KDV'si"
        elif 'matrah' in cl:                esle[orij] = 'Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı'
    df = df.rename(columns=esle)

    # Matrah sütunu adsızsa: KDV'nin hemen sağındaki, çoğunluğu sayısal sütun matrahtır
    MATRAH = 'Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı'
    if MATRAH not in df.columns and "KDV'si" in df.columns:
        cols = list(df.columns)
        i = cols.index("KDV'si")
        if i + 1 < len(cols):
            aday = cols[i + 1]
            vals = pd.to_numeric(df[aday].apply(para_deger), errors='coerce')
            if len(vals) and vals.notna().mean() > 0.6:
                df = df.rename(columns={aday: MATRAH})
    return df

def tarih_fmt(val):
    if pd.isna(val) or str(val).strip() == '': return ''
    if isinstance(val, datetime): return val.strftime('%d.%m.%Y')
    s = str(val).strip()
    try: return datetime.strptime(s[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
    except: return s

def sayi_fmt(val):
    if pd.isna(val) or str(val).strip() == '': return ''
    try:
        f = float(str(val).replace(',', '.'))
        return str(int(f)) if f == int(f) else f'{f:.2f}'.replace('.', ',')
    except: return str(val).strip()

def para_deger(val):
    """Kaynaktaki tutarı gerçek sayıya (float) çevirir; küsürat korunur.
    float, '45927,50', '1.234.567,89', '1,234,567.89' gibi girişleri destekler.
    Boş/geçersizse None döner."""
    if pd.isna(val) or str(val).strip() == '':
        return None
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).strip().replace(' ', '').replace('₺', '').replace('TL', '')
    if ',' in s and '.' in s:
        # Hem nokta hem virgül varsa: en sağdaki ondalık ayracıdır
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')   # 1.234,56 -> 1234.56
        else:
            s = s.replace(',', '')                     # 1,234.56 -> 1234.56
    elif ',' in s:
        s = s.replace(',', '.')                        # 1234,56 -> 1234.56
    try:
        return round(float(s), 2)
    except:
        return None

def para_oku(val):
    try: return float(str(val).replace(',', '.').replace(' ', ''))
    except: return 0.0

def dosya_adi_temizle(m): return re.sub(r'[\\/*?:"<>|]', '_', str(m)).strip()[:45]

def donem_bul(stem, df=None):
    """Dönemi (AA.YYYY) tespit eder, öncelik sırasıyla:
       1) Dosya adındaki Türkçe ay adı (+ varsa yıl)
       2) Dosya adındaki sayısal ay + yıl (ör: 04_2026)
       3) Listedeki fatura tarihlerinin en sık görülen ay/yılı
       4) Bugünün ay/yılı (son çare)"""
    # Türkçe karakterleri ASCII'ye indir: kullanıcılar dosya adında ay adını
    # çoğu zaman diakritiksiz yazar (NISAN, SUBAT, AGUSTOS, EYLUL, ARALIK...).
    # Böyle yazımların da doğru aya eşlenmesi için karşılaştırmayı sadeleştir.
    def _ascii_fold(s):
        tr = {'İ':'I','I':'I','Ş':'S','Ğ':'G','Ü':'U','Ö':'O','Ç':'C'}
        return ''.join(tr.get(c, c) for c in s.upper())
    fn = _ascii_fold(stem)
    aylar = {'OCAK':'01','ŞUBAT':'02','MART':'03','NİSAN':'04',
              'MAYIS':'05','HAZİRAN':'06','TEMMUZ':'07','AĞUSTOS':'08',
              'EYLÜL':'09','EKİM':'10','KASIM':'11','ARALIK':'12'}
    yil_m = re.search(r'(20\d{2})', fn)
    for ad, no in aylar.items():
        if _ascii_fold(ad) in fn:
            return f"{no}.{yil_m.group(1) if yil_m else datetime.now().year}"
    # Sayısal ay (tek başına 01-12; uzun sayı dizilerinin içi sayılmaz)
    ay_m = re.search(r'(?<!\d)(0[1-9]|1[0-2])(?!\d)', fn)
    if ay_m and yil_m:
        return f"{ay_m.group(1)}.{yil_m.group(1)}"
    # Veri tarihlerinden: en sık görülen ay/yıl
    if df is not None:
        tarih_col = sutun_bul(list(df.columns),
                              ['alış faturasının tarihi', 'fatura tarihi', 'tarih'])
        if tarih_col is not None:
            try:
                t = pd.to_datetime(df[tarih_col], errors='coerce', dayfirst=True).dropna()
                if len(t):
                    mod = t.dt.strftime('%m.%Y').mode()
                    if len(mod):
                        return mod.iloc[0]
            except Exception:
                pass
    if ay_m:
        return f"{ay_m.group(1)}.{datetime.now().year}"
    return datetime.now().strftime('%m.%Y')

def _gecersizlik_nedeni(x) -> str:
    s = str(x).strip().replace('.0', '').replace(' ', '')
    if s == '' or s.lower() in ('nan', 'none', 'nat'):
        return 'Boş/tanımsız'
    if not s.isdigit():
        return f'Sayısal değil: "{s[:20]}"'
    if len(s) < 10:
        return f'Çok kısa ({len(s)} hane, min 10)'
    if len(s) > 11:
        return f'Çok uzun ({len(s)} hane, max 11)'
    if s == s[0] * len(s):
        return f'Yer tutucu/sahte kimlik ({s})'
    return 'Geçersiz'

# ══════════════════════════════════════════
#  FİLTRELEME
# ══════════════════════════════════════════
def firmalari_filtrele(df, esik_tek, esik_toplam, yuzde80, log_cb):
    vkn_col   = sutun_bul(list(df.columns), ['vergi kimlik', 'vkn', 'tc kimlik'])
    tutar_col = sutun_bul(list(df.columns), ['kdv hariç tutarı', 'faturanın tutarı'])
    unvan_col = sutun_bul(list(df.columns), ['satıcının adı', 'ünvanı', 'unvan'])

    if not vkn_col:
        raise ValueError("VKN sütunu bulunamadı.")

    # Geçerli / geçersiz satır ayrımı
    def normalize_vkn(x):
        """VKN/TCKN normalize et: önde sıfır eksikse tamamla"""
        s = str(x).strip().replace('.0', '').replace(' ', '')
        if not s.isdigit(): return s
        # 8-9 hane → 10 haneye (VKN önde sıfır eksik)
        if len(s) in (8, 9):
            s = s.zfill(10)
        return s

    def gecerli_kimlik(x):
        s = normalize_vkn(x)
        if not (s.isdigit() and len(s) in (10, 11)):
            return False
        # Yer tutucu / sahte kimlikler: 0000000000, 1111111111, 2222222222 ...
        # (muhasebe dökümlerinde VKN'si girilmemiş satırlar için kullanılır;
        #  ayrı satıcıları tek sahte firmada birleştirmemek için geçersiz sayılır)
        if s == s[0] * len(s):
            return False
        return True

    # Önce normalize et, sonra geçerli/geçersiz ayır
    df = df.copy()
    df[vkn_col] = df[vkn_col].apply(normalize_vkn)

    mask_gecerli  = df[vkn_col].apply(gecerli_kimlik)
    df_t          = df[mask_gecerli].copy()
    df_gecersiz   = df[~mask_gecerli].copy()   # atlanacak satırlar

    df_t[vkn_col] = df_t[vkn_col].apply(
        lambda x: str(x).strip().replace('.0', '').replace(' ', ''))

    # Kaç VKN düzeltildi bilgisini logla
    duzeltilen = (df[vkn_col].apply(lambda x: len(str(x).strip().replace('.0','').replace(' ',''))).isin([8,9]) & mask_gecerli).sum()
    if duzeltilen > 0:
        log_cb(f"  🔧 {duzeltilen} satırda önde sıfır eksikti, otomatik tamamlandı (ör: 71419747 → 0071419747)", "warn")
    # Sağlam ayrıştırıcı: "1.234.567,89" gibi binlik ayraçlı tutarlar 0 olmasın
    df_t['_tutar'] = (df_t[tutar_col].apply(lambda v: para_deger(v) or 0.0)
                      if tutar_col else 0.0)

    # Geçersiz satır sayısını logla
    if len(df_gecersiz) > 0:
        log_cb(f"  ⚠️  {len(df_gecersiz)} geçersiz satır atlandı "
               f"(hatalı/eksik VKN-TC kimlik)", "warn")

    # ── %80 hedefi LİSTENİN TAMAMI üzerinden hesaplanır ──
    # Geçersiz VKN'li satırlar için tutanak üretilemez, ancak bu satırların
    # tutarları da listenin bir parçasıdır ve paydadan DÜŞÜLEMEZ. Aksi halde
    # gerçek kapsam %80'in altına düşer (yasal kural listenin bütünü içindir).
    TOPLAM_GECERLI = df_t['_tutar'].sum()
    TOPLAM_GECERSIZ = (df_gecersiz[tutar_col].apply(para_deger).fillna(0).sum()
                       if (tutar_col and len(df_gecersiz)) else 0.0)
    TOPLAM_LISTE = TOPLAM_GECERLI + TOPLAM_GECERSIZ
    HEDEF        = TOPLAM_LISTE * (yuzde80 / 100.0)
    # Boş liste / tümü sıfır tutar durumunda %oran hesaplarında bölme hatası olmasın
    pct = lambda x: (x / TOPLAM_LISTE * 100.0) if TOPLAM_LISTE else 0.0

    if TOPLAM_LISTE <= 0:
        log_cb("  ⚠️  Listede toplanabilir tutar bulunamadı "
               "(tutar sütunu boş ya da tüm satırlar geçersiz). "
               "Kaynak listeyi kontrol edin.", "warn")

    log_cb(f"  📊 Liste toplam: {TOPLAM_LISTE:,.2f} ₺  |  "
           f"%{yuzde80:.0f} hedef: {HEDEF:,.2f} ₺", "info")
    if TOPLAM_GECERSIZ > 0:
        log_cb(f"     (bunun {TOPLAM_GECERSIZ:,.2f} ₺'si geçersiz kimlikli satırlarda — "
               f"tutanak üretilemez ama %{yuzde80:.0f} hesabına dahildir)", "warn")

    # Firma bazında özet
    firma_ozet = {}
    for vkn, grp in df_t.groupby(vkn_col):
        uv = grp[unvan_col].dropna().unique() if unvan_col else []
        firma_ozet[vkn] = {
            'grp':    grp.drop(columns=['_tutar']),
            'toplam': grp['_tutar'].sum(),
            'max':    grp['_tutar'].max(),
            'unvan':  str(uv[0]) if len(uv) else vkn,
        }

    # ── 1. Aşama: 150K / 450K kriteri ──
    secilen   = {}
    kalan     = []

    for vkn, ozet in firma_ozet.items():
        if ozet['max'] >= esik_tek:
            secilen[vkn] = (ozet['grp'], f"tek fatura {ozet['max']:,.0f} ₺")
        elif ozet['toplam'] >= esik_toplam:
            secilen[vkn] = (ozet['grp'], f"toplam {ozet['toplam']:,.0f} ₺")
        else:
            kalan.append(vkn)

    secilen_tutar = sum(firma_ozet[v]['toplam'] for v in secilen)
    log_cb(f"  1️⃣  Kriter (tek≥{esik_tek:,.0f} / toplam≥{esik_toplam:,.0f}): "
           f"{len(secilen)} firma → {secilen_tutar:,.2f} ₺ "
           f"(%{pct(secilen_tutar):.1f})", "ok")

    # ── 2. Aşama: %80 tamamlama ──
    if secilen_tutar < HEDEF:
        kalan_sirali = sorted(kalan,
                              key=lambda v: firma_ozet[v]['toplam'],
                              reverse=True)
        ekstra = 0
        log_cb(f"  2️⃣  %{yuzde80:.0f} için ek firma ekleniyor:", "warn")
        for vkn in kalan_sirali:
            if secilen_tutar >= HEDEF:
                break
            ozet = firma_ozet[vkn]
            secilen[vkn] = (ozet['grp'],
                            f"%{yuzde80:.0f} tamamlama ({ozet['toplam']:,.0f} ₺)")
            secilen_tutar += ozet['toplam']
            ekstra += 1
            log_cb(f"     + {ozet['unvan'][:40]}  → "
                   f"kümülatif %{pct(secilen_tutar):.1f}", "warn")
        log_cb(f"  {ekstra} ek firma eklendi.", "warn")
        if secilen_tutar < HEDEF:
            log_cb(f"  ⛔ DİKKAT: Tüm uygun firmalar seçildiği halde %{yuzde80:.0f} "
                   f"hedefine ULAŞILAMADI (%{pct(secilen_tutar):.1f}). "
                   f"Sebep: geçersiz kimlikli satırlar ({TOPLAM_GECERSIZ:,.0f} ₺). "
                   f"Bu satırların VKN/TC bilgilerini listede düzeltip tekrar çalıştırın.", "err")
    else:
        log_cb(f"  ✅ 1. aşama zaten %{pct(secilen_tutar):.1f} — "
               f"ek firma gerekmedi.", "ok")

    log_cb(f"  📋 Toplam kapsam: {len(secilen)} firma  "
           f"(%{pct(secilen_tutar):.1f})", "ok")
    # Büyükten küçüğe sırala (dosya isimlendirme için)
    secilen_sirali = dict(
        sorted(secilen.items(),
               key=lambda kv: firma_ozet[kv[0]]['toplam'],
               reverse=True)
    )
    return secilen_sirali, df_gecersiz

# ══════════════════════════════════════════
#  EXCEL ÜRETME — Sistemin şablonu birebir
# ══════════════════════════════════════════
def guvenli_kaydet(wb, tam_yol):
    """Excel'i kaydeder. Windows'ta yol çok uzunsa (~260 karakter MAX_PATH sınırı)
    veya dosya adı geçersizse, adı kısaltarak yeniden dener; böylece uzun ünvanlı
    firmalarda dosya oluşturulamayıp numara atlaması yaşanmaz.
    Kaydedilen gerçek yolu döndürür."""
    yol    = Path(tam_yol)
    klasor = yol.parent
    stem   = yol.stem              # "N) 07_2026_VKN_UNVAN" (baştaki numara benzersizdir)
    # 1) Olduğu gibi dene
    try:
        wb.save(str(yol)); return str(yol)
    except Exception:
        pass
    # 2) Adı kısaltarak dene — baştaki "N) dönem_VKN" korunur, uzun ünvan kısalır
    for uzunluk in (80, 60, 40, 25, 15):
        try:
            kisa = stem[:uzunluk].rstrip(' ._-')
            yeni = klasor / f"{kisa}.xlsx"
            wb.save(str(yeni)); return str(yeni)
        except Exception:
            continue
    # 3) Windows uzun-yol ön eki ile son bir deneme
    if os.name == 'nt':
        try:
            uzun_pref = '\\\\?\\' + os.path.abspath(str(yol))
            wb.save(uzun_pref); return str(yol)
        except Exception:
            pass
    # Hiçbiri olmadıysa hatayı çağırana bildir
    wb.save(str(yol))
    return str(yol)


def firma_excel_olustur(firma_df, cikis_dosya, tum_kolonlar):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Veriler"

    for ci, b in enumerate(SABLON_SUTUNLAR, 1):
        c = ws.cell(1, ci, value=b); c.number_format = '@'

    tarih_col   = sutun_bul(tum_kolonlar, ["alış faturasının tarihi"])
    # Fatura no: olduğu gibi (BBK2026... → fatura no)
    faturano_col= sutun_bul(tum_kolonlar, ["alış faturasının sıra no"])
    # Seri: yalnızca gerçek bir seri sütunu varsa; numara sütununu seri sanma
    seri_col    = seri_sutunu_bul(tum_kolonlar, tarih_col, faturano_col)
    tutar_col   = sutun_bul(tum_kolonlar, ["kdv hariç tutarı", "faturanın tutarı"])
    kdv_col     = kdv_sutunu_bul(tum_kolonlar)
    if not kdv_col: kdv_col = sutun_bul(tum_kolonlar, ["toplam indirilecek kdv"])
    defter_col  = sutun_bul(tum_kolonlar, ["defter kayıt tarihi"])
    yevmiye_col = sutun_bul(tum_kolonlar, ["yevmiye numarası"])
    odeme_col   = sutun_bul(tum_kolonlar, ["ödeme şekli"])
    # Açıklama: listedeki "Alınan Mal ve/veya Hizmetin Cinsi" sütununu kullan,
    # yoksa varsa bir "Açıklama" sütununa düş
    aciklama_col= sutun_bul(tum_kolonlar, ["alınan mal ve/veya hizmetin cinsi", "hizmetin cinsi", "cinsi"])
    if not aciklama_col:
        aciklama_col = sutun_bul(tum_kolonlar, ["açıklama"])

    for ri, (_, row) in enumerate(firma_df.iterrows()):
        er = ri + 2
        def yaz(ci, val, tip=''):
            # Para modu: gerçek SAYI olarak yaz (küsürat korunur, metin değil)
            if tip == 'para':
                f = para_deger(val)
                if f is None:
                    ws.cell(er, ci, value=None)
                else:
                    c = ws.cell(er, ci, value=f)
                    c.number_format = '0.00'
                return
            if   tip == 't': val = tarih_fmt(val)
            elif tip == 'n': val = sayi_fmt(val)
            else: val = '' if pd.isna(val) else str(val).strip()
            if val is None or val in ('None', 'nan'): val = ''
            c = ws.cell(er, ci, value=val); c.number_format = '@'

        # 1: Tarih
        yaz(1, row.get(tarih_col, None) if tarih_col else None, 't')

        # 2: Faturanın Serisi — ayrı seri sütunu varsa ve doluysa al, yoksa boş
        seri_val = ''
        if seri_col:
            sv = row.get(seri_col, None)
            if pd.notna(sv) and str(sv).strip() not in ('', 'None', 'nan'):
                seri_val = str(sv).strip()
        # Güvence: seri, fatura numarasıyla birebir aynıysa bu bir seri değildir → boş bırak
        if seri_val and faturano_col:
            fno_ayni = row.get(faturano_col, None)
            if pd.notna(fno_ayni) and str(fno_ayni).strip() == seri_val:
                seri_val = ''
        yaz(2, seri_val if seri_val else '')

        # 3: Faturanın Numarası — komple fatura no (BBK2026000001422)
        fno = row.get(faturano_col, None) if faturano_col else None
        yaz(3, '' if pd.isna(fno) else str(fno).strip())

        # 4-5: Tutar (matrah) ve KDV — gerçek SAYI olarak (küsürat korunur)
        yaz(4, row.get(tutar_col,    None) if tutar_col    else None, 'para')
        yaz(5, row.get(kdv_col,      None) if kdv_col      else None, 'para')

        # 6-10
        yaz(6, row.get(defter_col,   None) if defter_col   else None, 't')
        yaz(7, row.get(yevmiye_col,  None) if yevmiye_col  else None)
        yaz(8, row.get(odeme_col,    None) if odeme_col    else None)
        yaz(9, row.get(aciklama_col, None) if aciklama_col else None)
        yaz(10, '')

    return guvenli_kaydet(wb, cikis_dosya)

# ══════════════════════════════════════════
#  ÖZET RAPOR
# ══════════════════════════════════════════
def ozet_rapor_olustur(df, secilen, df_gecersiz, esik_tek, esik_toplam,
                       yuzde80, donem, basarili, hatali_sayisi):
    """Bir çalışmanın tek sayfalık özetini (kapsam, hedef, gerçek %, geçersiz
    tutar vb.) hesaplayıp döndürür. Sayılar tutanaklarla aynı kaynaktan
    (para_deger ile) üretilir; iş kurallarını yeniden hesaplamaz, sadece
    zaten seçilmiş sonucu raporlar. Bir openpyxl Workbook döndürür."""
    tutar_col = sutun_bul(list(df.columns), ['kdv hariç tutarı', 'faturanın tutarı'])

    def _topla(frame):
        if frame is None or tutar_col is None or len(frame) == 0:
            return 0.0
        return float(frame[tutar_col].apply(lambda v: para_deger(v) or 0.0).sum())

    toplam_liste  = _topla(df)
    secilen_tutar = sum(_topla(grp) for grp, _neden in secilen.values())
    gecersiz_tutar= _topla(df_gecersiz)
    hedef_tutar   = toplam_liste * (yuzde80 / 100.0)
    kapsam_pct    = (secilen_tutar / toplam_liste * 100.0) if toplam_liste else 0.0

    satirlar = [
        ("Dönem",                         donem),
        ("Tek fatura limiti (₺)",         esik_tek),
        ("Toplam fatura limiti (₺)",      esik_toplam),
        ("Hedef kapsam yüzdesi (%)",      yuzde80),
        ("Liste toplam tutarı (₺)",       round(toplam_liste, 2)),
        (f"Hedef tutar (%{yuzde80:.0f}) (₺)", round(hedef_tutar, 2)),
        ("Seçilen firma sayısı",          len(secilen)),
        ("Tutanağı oluşturulan firma",    basarili),
        ("Oluşturulamayan firma",         hatali_sayisi),
        ("Seçilenlerin toplam tutarı (₺)", round(secilen_tutar, 2)),
        ("GERÇEK KAPSAM (%)",             round(kapsam_pct, 1)),
        ("Geçersiz kimlikli satır sayısı", 0 if df_gecersiz is None else len(df_gecersiz)),
        ("Geçersiz satırların tutarı (₺)", round(gecersiz_tutar, 2)),
    ]

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Özet"
    kalin = openpyxl.styles.Font(bold=True)
    baslik = ws.cell(1, 1, value="KARŞIT İNCELEME — ÇALIŞMA ÖZETİ"); baslik.font = kalin
    ws.cell(2, 1, value="Ölçüt").font = kalin
    ws.cell(2, 2, value="Değer").font = kalin
    for r, (etiket, deger) in enumerate(satirlar, start=3):
        ws.cell(r, 1, value=etiket)
        h = ws.cell(r, 2, value=deger)
        if isinstance(deger, float):
            h.number_format = '#,##0.00'
    # Gerçek kapsam satırını vurgula (hedefin altındaysa dikkat çeksin)
    kapsam_satir = 3 + [e for e, _ in satirlar].index("GERÇEK KAPSAM (%)")
    ws.cell(kapsam_satir, 1).font = kalin
    ws.cell(kapsam_satir, 2).font = kalin
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 24
    return wb, kapsam_pct

# ══════════════════════════════════════════
#  ANA İŞLEM
# ══════════════════════════════════════════
def dosyalari_isle(kaynak, esik_tek, esik_toplam, yuzde80, log_cb, tamam_cb,
                   ilerleme_cb=None):
    # ilerleme_cb(tamamlanan, toplam): GUI ilerleme çubuğunu günceller.
    # None geçilirse (ör. başsız test) hiçbir şey yapmaz.
    def _ilerle(t, top):
        if ilerleme_cb:
            try: ilerleme_cb(t, top)
            except Exception: pass
    try:
        _ilerle(0, 1)
        log_cb("📂 Dosya okunuyor...", "info")
        df    = ana_listeyi_oku(kaynak)
        donem = donem_bul(Path(kaynak).stem, df)
        log_cb(f"📅 Dönem: {donem}", "info")
        log_cb("🔍 Firmalar filtreleniyor...", "info")

        secilen, df_gecersiz = firmalari_filtrele(df, esik_tek, esik_toplam, yuzde80, log_cb)

        if not secilen:
            log_cb("⚠️  Hiçbir firma kriterleri karşılamıyor.", "warn")
            tamam_cb(None, 0, 0); return

        cikis_kl = Path(kaynak).parent / "Hazır Tutanaklar"
        cikis_kl.mkdir(exist_ok=True)

        # Klasörde eski Excel dosyası varsa uyar
        eski_dosyalar = list(cikis_kl.glob("*.xlsx"))
        if eski_dosyalar:
            log_cb(f"⚠️  Klasörde {len(eski_dosyalar)} eski dosya var.", "warn")
            log_cb(f"   Üzerine yazmamak için klasör yeniden adlandırılıyor...", "warn")
            zaman_damgasi = datetime.now().strftime("%Y%m%d_%H%M%S")
            yedek_kl = cikis_kl.parent / f"Hazır Tutanaklar_{zaman_damgasi}"
            cikis_kl.rename(yedek_kl)
            log_cb(f"   Eski klasör: {yedek_kl.name}", "warn")
            cikis_kl = Path(kaynak).parent / "Hazır Tutanaklar"
            cikis_kl.mkdir(exist_ok=True)
            log_cb(f"   Yeni klasör oluşturuldu.", "ok")

        log_cb(f"📁 Klasör: {cikis_kl}", "info")
        log_cb(f"{'─'*50}", "info")

        unvan_col = sutun_bul(list(df.columns), ['satıcının adı', 'ünvanı'])
        fno_col   = sutun_bul(list(df.columns), ['alış faturasının sıra no'])
        basarili  = 0; hatali = []
        vkn_sirali = []   # tutanağı oluşturulan firmaların VKN'leri (dosya sırasıyla)
        sira_no = 0       # yalnızca başarılı dosyalar için ARDIŞIK numara (atlama olmaz)

        toplam_firma = len(secilen)
        _ilerle(0, toplam_firma)
        for islenen, (vkn, (grp, neden)) in enumerate(secilen.items(), start=1):
            uv    = grp[unvan_col].dropna().unique() if unvan_col else []
            unvan = str(uv[0]) if len(uv) else ''
            # Firmadan örnek bir fatura numarası (birden fazla olsa da biri yeterli)
            ornek_fno = ''
            if fno_col:
                fdiz = grp[fno_col].dropna()
                if len(fdiz):
                    ornek_fno = str(fdiz.iloc[0]).strip()
            sira_no += 1
            try:
                temiz = dosya_adi_temizle(unvan) if unvan else vkn
                ad    = f"{sira_no}) {donem.replace('.','_')}_{vkn}_{temiz}.xlsx"
                dosya = str(cikis_kl / ad)
                firma_excel_olustur(grp, dosya, list(df.columns))
                basarili += 1
                vkn_sirali.append((sira_no, vkn, unvan, ornek_fno))
                log_cb(f"  [{sira_no:3}/{len(secilen)}] ✔ {vkn}  {unvan[:35]}", "ok")
            except Exception as e:
                sira_no -= 1   # başarısız → numarayı geri al ki sonraki dosyada atlama olmasın
                hatali.append((vkn, unvan, str(e)))
                log_cb(f"  ✘ HATA  {vkn}  {unvan[:30]}: {e}", "err")
            finally:
                _ilerle(islenen, toplam_firma)

        log_cb(f"\n{'═'*50}", "info")
        log_cb(f"✅ {basarili}/{len(secilen)} firma tamamlandı."
               + (f"  ⚠️ {len(hatali)} firma oluşturulamadı!" if hatali else ""), "ok")

        # ── Oluşturulamayan firmaları belirgin şekilde raporla (sessizce kaybolmasın) ──
        if hatali:
            log_cb(f"{'─'*50}", "err")
            log_cb(f"⚠️  AŞAĞIDAKİ {len(hatali)} FİRMANIN TUTANAĞI OLUŞTURULAMADI:", "err")
            for v, uv, sebep in hatali:
                log_cb(f"   • {v:15} {uv[:35]:35} → {sebep}", "err")
            try:
                import openpyxl as _opx
                hyol = str(cikis_kl / f"OLUSTURULAMAYANLAR_{donem.replace('.','_')}.xlsx")
                wbh = _opx.Workbook(); wsh = wbh.active; wsh.title = "Oluşturulamayanlar"
                for c, b in enumerate(["Vergi/TC No", "Ünvan", "Hata Nedeni"], 1):
                    hc = wsh.cell(1, c, value=b); hc.font = _opx.styles.Font(bold=True)
                for r, (v, uv, sebep) in enumerate(hatali, 2):
                    wsh.cell(r, 1, value=str(v)).number_format = '@'
                    wsh.cell(r, 2, value=str(uv))
                    wsh.cell(r, 3, value=str(sebep))
                wsh.column_dimensions['A'].width = 16
                wsh.column_dimensions['B'].width = 45
                wsh.column_dimensions['C'].width = 60
                wbh.save(hyol)
                log_cb(f"📄 Detay: OLUSTURULAMAYANLAR_{donem.replace('.','_')}.xlsx", "err")
            except Exception:
                pass

        # ── Tutanağı oluşturulan firmaların VKN listesi (dosya sırasıyla, büyükten küçüğe) ──
        if vkn_sirali:
            try:
                vkn_yolu = str(cikis_kl / f"VKN_LISTESI_{donem.replace('.','_')}.xlsx")
                wb_v = openpyxl.Workbook(); ws_v = wb_v.active; ws_v.title = "VKN Listesi"
                basliklar = ["Sıra No", "Vergi Kimlik / TC Kimlik No", "Ünvan", "Örnek Fatura No"]
                for c, bsl in enumerate(basliklar, start=1):
                    hc = ws_v.cell(1, c, value=bsl); hc.font = openpyxl.styles.Font(bold=True)
                for r, (sira, v, uv, fno) in enumerate(vkn_sirali, start=2):
                    ws_v.cell(r, 1, value=sira).number_format = '@'
                    ws_v.cell(r, 2, value=str(v)).number_format = '@'   # baştaki sıfırlar korunur
                    ws_v.cell(r, 3, value=str(uv) if uv else '')
                    ws_v.cell(r, 4, value=str(fno) if fno else '').number_format = '@'
                # Sütun genişlikleri
                ws_v.column_dimensions['A'].width = 8
                ws_v.column_dimensions['B'].width = 26
                ws_v.column_dimensions['C'].width = 45
                ws_v.column_dimensions['D'].width = 22
                wb_v.save(vkn_yolu)
                log_cb(f"📇 VKN listesi: VKN_LISTESI_{donem.replace('.','_')}.xlsx "
                       f"({len(vkn_sirali)} firma, dosya sırasıyla)", "ok")
            except Exception as e:
                log_cb(f"⚠️ VKN listesi yazılamadı: {e}", "warn")

        # ── Geçersiz satır raporu ──
        if df_gecersiz is not None and len(df_gecersiz) > 0:
            try:
                rapor_yolu = str(cikis_kl / f"GECERSIZ_SATIRLAR_{donem.replace('.','_')}.xlsx")
                # Geçersiz satırların hangi VKN/TC içerdiğini de göster
                vkn_col_r = sutun_bul(list(df.columns), ['vergi kimlik','vkn','tc kimlik'])
                unvan_col_r = sutun_bul(list(df.columns), ['satıcının adı','ünvanı'])
                rapor_df = df_gecersiz.copy()
                rapor_df.insert(0, 'Sorun', rapor_df[vkn_col_r].apply(
                    lambda x: _gecersizlik_nedeni(x)) if vkn_col_r else 'Bilinmiyor')
                rapor_df.to_excel(rapor_yolu, index=False)
                log_cb(f"\n📋 Geçersiz satır raporu: GECERSIZ_SATIRLAR_{donem.replace('.','_')}.xlsx", "warn")
                log_cb(f"   {len(df_gecersiz)} satır — VKN/TC kimlik hatalı veya eksik:", "warn")
                # Günlüğe de yaz
                if vkn_col_r:
                    for _, row in df_gecersiz.iterrows():
                        vkn_val = str(row.get(vkn_col_r,'')).strip()
                        unvan_val = str(row.get(unvan_col_r,''))[:40].strip() if unvan_col_r else ''
                        neden = _gecersizlik_nedeni(vkn_val)
                        log_cb(f"   • {vkn_val:15}  {unvan_val:40}  → {neden}", "warn")
            except Exception as e:
                log_cb(f"⚠️  Rapor oluşturulamadı: {e}", "err")

        # ── Çalışma özeti (tek sayfalık kapsam raporu) ──
        try:
            ozet_wb, kapsam_pct = ozet_rapor_olustur(
                df, secilen, df_gecersiz, esik_tek, esik_toplam,
                yuzde80, donem, basarili, len(hatali))
            ozet_yolu = str(cikis_kl / f"OZET_RAPOR_{donem.replace('.','_')}.xlsx")
            ozet_wb.save(ozet_yolu)
            log_cb(f"📊 Özet rapor: OZET_RAPOR_{donem.replace('.','_')}.xlsx "
                   f"(gerçek kapsam %{kapsam_pct:.1f})", "ok")
        except Exception as e:
            log_cb(f"⚠️  Özet rapor yazılamadı: {e}", "warn")

        log_cb(f"📁 {cikis_kl}", "ok")
        tamam_cb(str(cikis_kl), basarili, len(hatali))

    except Exception as e:
        log_cb(f"\n❌ {e}", "err")
        tamam_cb(None, 0, 1)

# ══════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════
APP="#F5F6FA"; KART="#FFFFFF"; KOYU="#1A1A2E"
KIRMIZI="#E63946"; GRI="#6B7280"; TURUNCU="#F4A261"
F_ANA=('Segoe UI',10); F_BLK=('Segoe UI',10,'bold')
F_KUC=('Segoe UI',9);  F_BAS=('Segoe UI',13,'bold')
F_CON=('Consolas',9)

class KDVBolmeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("YMM Karşıt İnceleme Asistanı")
        self.root.geometry("780x620")
        self.root.minsize(780, 580)
        self.root.configure(bg=APP)
        # Pencere başlık çubuğu ikonu (logo.ico varsa)
        try:
            ico = kaynak_yolu('logo.ico')
            if ico.exists():
                self.root.iconbitmap(str(ico))
        except Exception:
            pass
        self._isleniyor = False
        self._logo_img  = None

        # Son kullanılan kriterleri yükle (yoksa varsayılanlar)
        ayar = self._ayar_yukle()
        self.esik_tek    = tk.StringVar(value=ayar.get("esik_tek",    "150000"))
        self.esik_toplam = tk.StringVar(value=ayar.get("esik_toplam", "450000"))
        self.yuzde80     = tk.StringVar(value=ayar.get("yuzde80",     "80"))

        self._ui()
        self._surukle_birak()

    @staticmethod
    def _ayar_yukle():
        """Kayıtlı kriterleri okur; dosya yok/bozuksa boş sözlük döner."""
        try:
            with open(AYAR_YOLU, 'r', encoding='utf-8') as f:
                d = json.load(f)
            return d if isinstance(d, dict) else {}
        except Exception:
            return {}

    def _ayar_kaydet(self):
        """Geçerli kriterleri diske yazar; hata olursa sessizce geçilir
        (ayar kaydedememek işlemi engellememeli)."""
        try:
            with open(AYAR_YOLU, 'w', encoding='utf-8') as f:
                json.dump({
                    "esik_tek":    self.esik_tek.get(),
                    "esik_toplam": self.esik_toplam.get(),
                    "yuzde80":     self.yuzde80.get(),
                }, f, ensure_ascii=False)
        except Exception:
            pass

    def _ui(self):
        # ── Sol panel ──
        sol = tk.Frame(self.root, bg=APP, width=310)
        sol.pack(side='left', fill='y', padx=(18,8), pady=18)
        sol.pack_propagate(False)
        self.sol = sol

        self._logo_yukle()

        tk.Label(sol, text="Karşıt İnceleme Hazırlayıcı",
                 font=F_BAS, bg=APP, fg=KOYU,
                 wraplength=270, justify='center').pack(pady=(6,2))
        tk.Label(sol, text="Excel dosyasını aşağıya bırakın",
                 font=F_KUC, bg=APP, fg=GRI).pack(pady=(0,10))

        # ── Eşik kutusu ──
        ef = tk.LabelFrame(sol, text="  Kapsam Kriterleri  ",
                           font=('Segoe UI',9,'bold'),
                           bg=APP, fg=KOYU, relief='solid', bd=1,
                           padx=10, pady=8)
        ef.pack(fill='x', pady=(0,10))

        self._esik_satir(ef, "Tek fatura limiti (₺):", self.esik_tek, "ör: 150000")
        tk.Frame(ef, bg='#E2E8F0', height=1).pack(fill='x', pady=5)
        self._esik_satir(ef, "Toplam fatura limiti (₺):", self.esik_toplam, "ör: 450000")
        tk.Frame(ef, bg='#E2E8F0', height=1).pack(fill='x', pady=5)
        self._esik_satir(ef, "Kapsam yüzdesi (%):", self.yuzde80, "ör: 80")

        tk.Label(ef, text="Tek fatura ≥ limit  VEYA  toplam ≥ limit\n"
                          "Yetmezse büyükten küçüğe ekleyerek\n"
                          "belirtilen % karşılanana kadar devam eder.",
                 font=('Segoe UI',8), bg=APP, fg=GRI, justify='left').pack(
                 fill='x', pady=(6,0))

        # ── Sürükle-bırak ──
        self.birak = tk.Frame(sol, bg=KART, relief='solid', bd=1, cursor='hand2')
        self.birak.pack(fill='both', expand=True, pady=(0,8))

        self.birak_ikon = tk.Label(self.birak, text="🗂",
                                   font=('Segoe UI',34), bg=KART)
        self.birak_ikon.pack(expand=True, pady=(20,4))

        self.birak_yazi = tk.Label(self.birak,
                                   text="Dosyayı Buraya Sürükleyin",
                                   font=F_BLK, bg=KART, fg=GRI)
        self.birak_yazi.pack(expand=True, pady=(0,4))

        self.birak_alt = tk.Label(self.birak, text="ya da tıklayın",
                                  font=F_KUC, bg=KART, fg='#9CA3AF')
        self.birak_alt.pack(pady=(0,14))

        for w in (self.birak, self.birak_ikon, self.birak_yazi, self.birak_alt):
            w.bind('<Button-1>', self._tiklayarak_sec)

        self.durum_lbl = tk.Label(sol, text="Hazır",
                                  font=F_KUC, bg=APP, fg=GRI)
        self.durum_lbl.pack()

        # ── İlerleme çubuğu (işlem sırasında firma sayısına göre dolar) ──
        try:
            stil = ttk.Style()
            stil.configure("Ymm.Horizontal.TProgressbar",
                           troughcolor='#E2E8F0', background='#68D391')
            self.ilerleme_var = tk.DoubleVar(value=0)
            self.ilerleme_bar = ttk.Progressbar(
                sol, style="Ymm.Horizontal.TProgressbar",
                orient='horizontal', mode='determinate',
                variable=self.ilerleme_var, maximum=100)
            self.ilerleme_bar.pack(fill='x', pady=(6, 0))
            self.ilerleme_yazi = tk.Label(sol, text="", font=('Segoe UI', 8),
                                          bg=APP, fg=GRI)
            self.ilerleme_yazi.pack()
        except Exception:
            self.ilerleme_var = None
            self.ilerleme_bar = None
            self.ilerleme_yazi = None

        # ── Sağ panel (log) ──
        sag = tk.Frame(self.root, bg=APP)
        sag.pack(side='right', fill='both', expand=True, padx=(0,18), pady=18)

        tk.Label(sag, text="İşlem Günlüğü",
                 font=F_BLK, bg=APP, fg=KOYU, anchor='w').pack(fill='x')

        lf = tk.Frame(sag, bg='#0F0F1A')
        lf.pack(fill='both', expand=True, pady=(4,0))

        self.log = tk.Text(lf, font=F_CON, bg='#0F0F1A', fg='#A0AEC0',
                           relief='flat', bd=8, wrap='word', state='disabled')
        sb = ttk.Scrollbar(lf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.log.pack(fill='both', expand=True)

        self.log.tag_config('ok',   foreground='#68D391')
        self.log.tag_config('err',  foreground='#FC8181')
        self.log.tag_config('info', foreground='#90CDF4')
        self.log.tag_config('warn', foreground='#F6AD55')

        self._log("e-YMM Karşıt İnceleme Asistanı hazır.\n"
                  "Kriterleri güncelleyin,\n"
                  "KDV listesini sol tarafa sürükleyin.", "info")

    def _esik_satir(self, parent, etiket, var, ipucu):
        tk.Label(parent, text=etiket, font=F_KUC, bg=APP,
                 fg=KOYU, anchor='w').pack(fill='x')
        f = tk.Frame(parent, bg=APP); f.pack(fill='x', pady=(2,0))
        ttk.Entry(f, textvariable=var, font=F_ANA, width=14).pack(side='left', padx=(0,6))
        tk.Label(f, text=ipucu, font=('Segoe UI',8),
                 bg=APP, fg='#9CA3AF').pack(side='left')

    def _logo_yukle(self):
        for yol in [kaynak_yolu('logo.png'),
                    kaynak_yolu('logo.jpg'),
                    kaynak_yolu('logo.ico')]:
            if yol.exists():
                try:
                    from PIL import Image, ImageTk
                    img = Image.open(str(yol)).convert('RGBA')
                    img.thumbnail((100,100), Image.LANCZOS)
                    bg = Image.new('RGBA', img.size, (245,246,250,255))
                    bg.paste(img, mask=img.split()[3])
                    self._logo_img = ImageTk.PhotoImage(bg)
                    tk.Label(self.sol, image=self._logo_img, bg=APP).pack(pady=(0,4))
                    return
                except: pass
        c = tk.Canvas(self.sol, width=90, height=90, bg=APP, highlightthickness=0)
        c.pack(pady=(0,4))
        c.create_oval(5,5,85,85, fill=KIRMIZI, outline='')
        c.create_text(45,45, text='YMM', fill='white', font=('Segoe UI',16,'bold'))

    def _surukle_birak(self):
        try:
            import tkinterdnd2
            self.root.drop_target_register(tkinterdnd2.DND_FILES)
            self.root.dnd_bind('<<Drop>>', lambda e: self._isle(
                e.data.strip().strip('{}')))
        except: pass

    def _tiklayarak_sec(self, e=None):
        d = filedialog.askopenfilename(
            title="Ana KDV Listesini Seçin",
            filetypes=[("Excel","*.xls *.xlsx"),("Tümü","*.*")])
        if d: self._isle(d)

    def _isle(self, dosya):
        if self._isleniyor:
            self._log("⚠️  İşlem devam ediyor...", "warn"); return
        if Path(dosya).suffix.lower() not in ('.xls','.xlsx'):
            self._log("❌ Geçersiz format.", "err"); return
        try:
            esik_tek    = float(self.esik_tek.get().replace('.','').replace(',','.'))
            esik_toplam = float(self.esik_toplam.get().replace('.','').replace(',','.'))
            yuzde80     = float(self.yuzde80.get().replace(',','.'))
        except:
            messagebox.showerror("Hata", "Kriter değerleri geçersiz.\nSadece sayı girin.")
            return

        # Kullanılan kriterleri bir sonraki açılış için sakla
        self._ayar_kaydet()

        self._isleniyor = True
        self._birak_guncelle("⏳ İşleniyor...", TURUNCU)
        self.durum_lbl.config(text=Path(dosya).name, fg=KOYU)
        self._ilerleme(0, 0)   # çubuğu sıfırla
        self._log(f"\n{'═'*50}", "info")
        self._log(f"📂 {Path(dosya).name}", "info")

        threading.Thread(
            target=dosyalari_isle,
            args=(dosya, esik_tek, esik_toplam, yuzde80, self._log, self._tamam,
                  self._ilerleme),
            daemon=True
        ).start()

    def _log(self, msg, tip=''):
        def _():
            self.log.config(state='normal')
            self.log.insert('end', msg+'\n', tip or '')
            self.log.see('end')
            self.log.config(state='disabled')
        self.root.after(0, _)

    def _ilerleme(self, tamamlanan, toplam):
        """İş parçacığından çağrılır; ilerleme çubuğunu ana thread'de günceller."""
        if self.ilerleme_bar is None:
            return
        yuzde = (tamamlanan / toplam * 100.0) if toplam else 0.0
        def _():
            try:
                self.ilerleme_var.set(yuzde)
                if toplam and tamamlanan < toplam:
                    self.ilerleme_yazi.config(text=f"{tamamlanan}/{toplam} firma")
                elif toplam:
                    self.ilerleme_yazi.config(text=f"{toplam}/{toplam} firma ✓")
                else:
                    self.ilerleme_yazi.config(text="")
            except Exception:
                pass
        self.root.after(0, _)

    def _tamam(self, klasor, basarili, hatali):
        self._isleniyor = False
        def _():
            if klasor and basarili > 0:
                self._birak_guncelle(f"✅ {basarili} dosya hazır", "#68D391")
                self.durum_lbl.config(text=f"{basarili} Excel oluşturuldu", fg='#276749')
                try: os.startfile(klasor)
                except: pass
            else:
                self._birak_guncelle("❌ Hata oluştu", "#FC8181")
                self.durum_lbl.config(text="Hata — log'u inceleyin", fg='#C53030')
            def _sifirla():
                self._birak_guncelle("Yeni dosya için buraya bırakın", GRI)
                if self.ilerleme_bar is not None:
                    try:
                        self.ilerleme_var.set(0)
                        self.ilerleme_yazi.config(text="")
                    except Exception:
                        pass
            self.root.after(4000, _sifirla)
        self.root.after(0, _)

    def _birak_guncelle(self, yazi, renk):
        self.birak_yazi.config(text=yazi, fg=renk)

# ══════════════════════════════════════════
if __name__ == '__main__':
    try:
        import tkinterdnd2; root = tkinterdnd2.Tk()
    except: root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except: pass
    KDVBolmeApp(root)
    root.mainloop()
