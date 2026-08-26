"""Test altyapısı — bu ortamda (ve CI'da) GUI/tkinter bulunmayabilir.

CLAUDE.md §8'de tarif edildiği gibi, exay.py bir masaüstü GUI uygulamasıdır ve
modül düzeyinde `import tkinter` yapar. Başsız (headless) testlerde tkinter
kurulu olmayabileceğinden, exay içe aktarılmadan ÖNCE hafif bir tkinter stub'ı
sys.modules'e yerleştiriyoruz. Böylece iş mantığı fonksiyonları gerçek bir
ekran/GUI olmadan doğrudan çağrılıp test edilebilir.

Gerçek bir tkinter kuruluysa ona dokunmayız (stub yalnızca eksikse devreye girer).
"""
import sys
from unittest.mock import MagicMock


def _tkinter_stub_kur():
    try:
        import tkinter  # noqa: F401  (gerçek tkinter varsa stub'a gerek yok)
        return
    except Exception:
        pass

    # tkinter'ı MagicMock ile taklit et: her widget/çağrı çocuk MagicMock döndürür.
    # Böylece iş mantığı test edilebildiği gibi, KDVBolmeApp gerçek bir ekran
    # olmadan sahte bir pencereyle KURULABİLİR de — _ui içindeki eksik metot/
    # callback bağlamaları (ör. bind command'ları) testte hemen yakalanır.
    tk = MagicMock(name='tkinter')
    for alt in ('ttk', 'messagebox', 'filedialog'):
        m = MagicMock(name='tkinter.' + alt)
        setattr(tk, alt, m)
        sys.modules['tkinter.' + alt] = m
    sys.modules['tkinter'] = tk


_tkinter_stub_kur()
