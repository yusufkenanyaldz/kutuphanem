# Sabri Hamamcı — Yeminli Mali Müşavir · Web Sitesi

Tek dosyalık, bağımsız çalışan bir tanıtım sitesi. Derleme aracı, npm paketi veya
sunucu tarafı kod gerektirmez — `index.html` dosyasını bir tarayıcıda açmanız yeterlidir.

```
ymm/
├── index.html      ← sitenin tamamı (HTML + CSS + JS)
├── README.md       ← bu dosya
└── assets/         ← (opsiyonel) fotoğraf, logo vb. buraya
```

---

## 1. Önce şunları doldurun

`index.html` dosyasını açın, en alttaki `<script>` bloğunun başında yer alan
**`SITE`** nesnesini düzenleyin. Buradaki değerler sayfanın her yerine
(üst bar, menü, iletişim bölümü, footer, WhatsApp butonu) otomatik yansır:

```js
const SITE = {
  telefon:        "+90 (212) 000 00 00",   // ekranda görünen telefon
  telefonRaw:     "+902120000000",         // tıklanınca aranan numara
  whatsapp:       "905000000000",          // ülke kodu + numara, + ve boşluk YOK
  whatsappMesaj:  "Merhaba, bilgi almak istiyorum.",
  eposta:         "info@sabrihamamci.com",
  adres:          "... Mah. ... Cad. No: 1 Kat: 3, Şişli / İstanbul",
  haritaEmbed:    "",                      // aşağıdaki 4. adıma bakın
  formEndpoint:   "",                      // aşağıdaki 3. adıma bakın
  oranGuncelleme: "1 Ocak 2026"
};
```

> `whatsapp` alanını boş bırakırsanız sağ alttaki WhatsApp butonu otomatik olarak kaldırılır.

---

## 2. İçerik güncelleme

Sık değişen içerikler HTML'in içine dağılmış değil; script bloğundaki
listelerde toplanmıştır. Yeni satır eklemek için listeye bir satır yazmanız yeterli.

| Liste | İçerik | Not |
|---|---|---|
| `TAKVIM` | Vergi takvimi tablosu | `["Beyanname", "Dönem", "Son gün"]` |
| `ORANLAR` | Oran ve tutarlar tablosu | `["Konu", "Oran", "Açıklama"]` |
| `BAGLANTILAR` | Faydalı bağlantı kartları | `["Ad", "Alan adı", "URL", "ikon"]` |
| `DUYURULAR` | Mevzuat / duyuru kartları | `["Etiket", "YYYY-AA-GG", "Başlık", "Özet", "URL"]` |
| `SSS` | Sıkça sorulan sorular | `["Soru", "Cevap"]` |

İkon adları: `bank`, `screen`, `shield`, `doc`, `users`.

Duyurularda URL alanını boş bırakırsanız kartın altındaki bağlantı otomatik olarak
"Konu hakkında bilgi alın" şeklinde iletişim bölümüne yönlenir.

### ⚠️ Kontrol edilmesi gereken yerler

Aşağıdaki alanlar **örnek/temsili** değerlerle dolduruldu; yayına almadan önce
gerçek bilgilerle değiştirin. Dosya içinde `GÜNCELLENECEK` yazarak arayabilirsiniz.

- **Hero altındaki 4 rakam** (25+ yıl deneyim, 150+ müşteri, 500+ rapor, %100) —
  `index.html` içinde `<div class="stats">` bloğu.
- **`ORANLAR` tablosu** — oranlar mevzuat değişikliğiyle güncellenir; yayına almadan
  önce yürürlükteki değerlerle karşılaştırın ve `SITE.oranGuncelleme` tarihini yazın.
- **`TAKVIM` tablosu** — sirküler kaynaklı süre uzatmalarında güncelleyin.
- **`DUYURULAR`** — şu an 3 adet örnek duyuru var; kendi sirkülerlerinizle değiştirin.
- **Oda üyeliği / sicil numarası** — "Hakkımızda" bölümündeki `Oda Üyeliği` kartı.
- **KVKK Aydınlatma Metni** — footer'daki açılır bölüm; kurumunuza göre
  bir hukuk danışmanı gözetiminde nihai hâline getirin.

---

## 3. İletişim formunu çalışır hâle getirme

Site statik olduğu için formun bir gönderim servisine bağlanması gerekir.

**Seçenek A — Hiçbir şey yapmayın (varsayılan).**
`formEndpoint` boşken form, ziyaretçinin e-posta programını doldurulmuş bir taslakla
açar. Kurulum gerektirmez, ama ziyaretçinin cihazında e-posta uygulaması olmalıdır.

**Seçenek B — Formspree (önerilen, ücretsiz başlangıç paketi var).**

1. <https://formspree.io> üzerinde ücretsiz hesap açın.
2. Yeni bir form oluşturup e-posta adresinizi doğrulayın.
3. Size verilen adresi (`https://formspree.io/f/xxxxxxx`) `formEndpoint` alanına yazın.

Aynı biçimde Web3Forms, Getform veya Basin de kullanılabilir — hepsi `FormData`
gövdesiyle POST kabul eder, kod değişikliği gerekmez.

Formda görünmez bir "bal küpü" (honeypot) alanı bulunur; botların doldurduğu
gönderimler sessizce elenir.

---

## 4. Haritayı ekleme

1. Google Haritalar'da ofis adresini bulun.
2. **Paylaş → Harita yerleştir → HTML'yi kopyala**.
3. Kopyaladığınız kodun içindeki `src="..."` değerini `SITE.haritaEmbed` alanına yapıştırın.

Alan boş bırakılırsa yerine bilgilendirici bir yer tutucu görünür.

---

## 5. Fotoğraf ekleme

1. Portre fotoğrafını `ymm/assets/sabri-hamamci.jpg` olarak kaydedin
   (önerilen: 4:5 oran, en az 800×1000 px, 300 KB altı).
2. `index.html` içinde "FOTOĞRAF" yorumunu bulun, `<img>` satırının başındaki ve
   sonundaki yorum işaretlerini kaldırın, altındaki `<div class="portrait-ph">`
   bloğunu silin.

---

## 6. Yayına alma

**GitHub Pages:** Depo ayarlarından *Settings → Pages* bölümünde branch'i seçip
kaydedin. Site `https://<kullanıcı>.github.io/kutuphanem/ymm/` adresinde yayına girer.

**Netlify / Vercel / Cloudflare Pages:** Klasörü sürükleyip bırakmanız yeterli;
derleme komutu yok, yayın dizini `ymm`.

**Klasik hosting:** `index.html` (ve varsa `assets/`) dosyalarını FTP ile
`public_html` altına kopyalayın.

Alan adı bağladıktan sonra `index.html` içindeki `<link rel="canonical">` ve
`og:` meta etiketlerindeki adresleri gerçek alan adıyla güncelleyin.

---

## Teknik notlar

- Tek dosya, harici JS kütüphanesi yok. Tek dış bağımlılık Google Fonts'tur
  (Lora + Inter); internet erişimi olmadığında sistem yazı tiplerine düşer.
- Mobil öncelikli duyarlı tasarım; 680 px ve 900 px kırılma noktaları.
- Erişilebilirlik: klavye ile tam gezinme, `aria` etiketleri, görünür odak halkaları,
  `prefers-reduced-motion` desteği, "İçeriğe geç" bağlantısı.
- Yazdırma stili tanımlıdır — sayfa Ctrl+P ile temiz biçimde çıktı alınabilir.
- Bölümlerin sırası değiştirilebilir; menüdeki `href="#..."` değerleri ile
  `<section id="...">` değerlerinin eşleşmesi yeterlidir.
