# 🏊 🚴 🏃 IRONMAN Yolculuğu

Sıfırdan Full Ironman'e giden yolu takip eden, kendi bilgisayarınızda çalışan web uygulaması.
Excel dosyasındaki seviye sistemi, haftalık özetler ve hedef takibinin tamamı Python'a taşındı;
veriler **SQLite** veritabanında saklanıyor.

---

## Kurulum

### macOS / Linux
```bash
cd ironman-tracker
./run.sh
```

### Windows
`run.bat` dosyasına çift tıklayın.

İlk çalıştırmada sanal ortam kurulur ve bağımlılıklar (`Flask`, `openpyxl`) indirilir —
bir dakika sürer. Sonraki açılışlar anında olur.

Tarayıcıda açın: **http://127.0.0.1:5000**

### Elle kurulum
```bash
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

### Telefondan erişim
```bash
./run.sh --host 0.0.0.0
```
Uygulama, aynı Wi-Fi ağındaki telefondan da açılabilir; başlangıçta yazdırılan adresi kullanın.
Antrenman sonrası kaydı telefondan girmek için pratiktir.

> Uygulama internete açık değildir ve kimlik doğrulaması yoktur — sadece kendi ağınızda kullanın.

---

## Sayfalar

| Sayfa | Ne yapar |
|---|---|
| **Dashboard** | Genel seviye, branş seviyeleri, toplam istatistikler, 4 grafik, son antrenmanlar |
| **Antrenmanlar** | Tüm kayıtlar; filtre, arama, sayfalama, ekleme/düzenleme/silme |
| **Seviyeler** | Her branş için 7 basamaklı seviye merdiveni, tamamlanma tarihleriyle |
| **Branş İlerleme** | Branş bazlı detaylı analiz, pace/hız trendi, seviye bileşenleri |
| **Haftalık** | Hafta hafta hacim ve süre, önceki haftaya göre % değişim |
| **Brick** | Bisiklet→koşu antrenmanları ve bisiklet sonrası pace kaybı |
| **Hedefler** | 14 kilometre taşı + kendi hedeflerinizi ekleme |
| **İçe Aktar** | Strava / Garmin / kendi CSV'nizden toplu yükleme |
| **Ayarlar** | Seviye eşikleri, yedekler, dışa aktarma |

---

## Seviye motoru

Her branş için **iki bağımsız seviye** hesaplanır:

- **Mesafe Seviyesi (rozet)** — en uzun *tek* antrenmanınıza bakar. Bir kez kazanıldı mı asla düşmez.
- **Form Seviyesi (süreklilik)** — son 30 gündeki *toplam* hacminize bakar. Ara verirseniz düşer.

```
Mevcut seviye  = min(mesafe seviyesi, form seviyesi)
Genel seviye   = min(yüzme, bisiklet, koşu)        ← en zayıf halka
Ironman %      = ortalama( (seviye + seviye içi ilerleme) / 6 )
```

Yani bir kez uzun mesafe yapmak seviye atlatmaz; düzenli antrenman da gerekir.
Eşikler **Ayarlar → Seviye Eşikleri**'nden değiştirilebilir (soldan sağa azalmamalıdır).

| Sv. | Ad | 🏊 | 🚴 | 🏃 |
|---|---|---|---|---|
| 0 | Başlangıç | 50–100 m | 10 km | 1–2 km |
| 1 | Temel Dayanıklılık | 500 m | 30 km | 5 km |
| 2 | Sprint Triathlon | 750 m | 20 km yarış · 30 km taban | 5 km |
| 3 | Olympic Triathlon | 1.500 m | 40 km | 10 km |
| 4 | 70.3 / Half Ironman | 1.900 m | 90 km | 21,1 km |
| 5 | Ironman Hazırlığı | 3.000–4.000 m | 120–180 km | 25–35 km |
| 6 | **IRONMAN** 🏆 | 3.800 m | 180 km | 42,2 km |

---

## Veri girişi

- **Mesafe her branşta kilometre**: yüzme 750 m → `0,75` · 3.800 m → `3,8`
- **Süre her zaman dakika**: 1 saat 12 dk → `72` · 3 saat 30 dk → `210`
- Virgül de nokta da kabul edilir (`8,5` = `8.5`)
- Pace otomatik hesaplanır: yüzmede **dk/100 m**, koşuda **dk/km**, bisiklette **km/s**

### Brick nasıl girilir?
Bisiklet ve koşuyu **iki ayrı kayıt** olarak aynı tarihe girin, ikisinde de “Brick” kutusunu
işaretleyin. BRICK sayfası bunları otomatik eşleştirir ve bisiklet sonrası koşu pace'inizi,
aynı dönemin normal koşu pace'iyle karşılaştırır.

---

## CSV içe aktarma

**Strava:** Ayarlar → Hesabım → *Hesabınızın bir kopyasını indirin* → arşivdeki `activities.csv`
**Garmin Connect:** Aktiviteler → Tüm Aktiviteler → *Dışa Aktar (CSV)*

Kolon adları, birimler (m / km / mil, saniye / dakika / `hh:mm:ss`) ve branş isimleri
(Run, Koşu, Ride, Bisiklet, Swim, Yüzme…) otomatik algılanır. Yükleme öncesi önizleme gösterilir.
Aynı dosyayı iki kez yüklemek güvenlidir — mükerrer kayıt eklenmez.

---

## Excel dışa aktarma

**Ayarlar → Excel (.xlsx)** veya Dashboard'daki düğme, 9 sayfalık ve **formülleri canlı** bir
çalışma kitabı üretir: DASHBOARD, SEVİYELER, ANTRENMAN KAYIT, HAFTALIK TAKİP, BRANŞ İLERLEME,
BRICK, HEDEFLER, NASIL KULLANILIR, AYARLAR. Grafikler, koşullu biçimlendirme ve açılır listeler
dahildir; dosya uygulamadan bağımsız çalışır.

> Excel dosyasında yaptığınız değişiklikler uygulamaya **geri yazılmaz**.
> Veri kaynağınız uygulama olsun; Excel'i rapor ve yedek olarak kullanın.

---

## Yedekleme

- Açılışta günde bir kez otomatik yedek alınır (`data/backups/`), son **10** yedek saklanır
- **Ayarlar → Şimdi yedek al** ile elle yedek alabilirsiniz
- Her yedek indirilebilir ve tek tıkla geri yüklenebilir (geri yükleme öncesi mevcut halin
  yedeği de alınır)
- Tüm veriniz tek bir dosyadadır: `data/ironman.db` — kopyalayıp taşıyabilirsiniz

---

## Proje yapısı

```
ironman-tracker/
├── app.py                  Flask rotaları ve giriş noktası
├── selftest.py             80+ testlik doğrulama paketi
├── requirements.txt
├── run.sh / run.bat        kurulum + çalıştırma
├── ironman/
│   ├── config.py           branşlar, seviye eşikleri, varsayılan hedefler
│   ├── db.py               SQLite şeması ve CRUD
│   ├── levels.py           seviye motoru
│   ├── stats.py            toplamalar (branş, haftalık, brick, hedef)
│   ├── charts.py           bağımlılıksız SVG grafik üreteci
│   ├── importer.py         Strava/Garmin CSV okuyucu
│   ├── backup.py           yedekleme
│   └── excel_export.py     formüllü .xlsx üreteci
├── templates/              11 Jinja şablonu
├── static/style.css        tasarım sistemi (açık + koyu tema)
└── data/                   ironman.db · backups/ · exports/   (git'e girmez)
```

---

## Test

```bash
python selftest.py
```

Seviye motorunun çıktılarını Excel formüllerinin sonuçlarıyla karşılaştırır, tüm rotaları,
form akışlarını, CSV içe aktarmayı, yedeklemeyi ve Excel çıktısını doğrular.
Geçici bir veritabanı kullanır — gerçek verinize dokunmaz.

---

## Notlar

- Python 3.10+ gerekir
- Tek kullanıcılıdır, kimlik doğrulaması yoktur
- Tarayıcı teması otomatik algılanır; sağ üstteki ◐ düğmesiyle değiştirilebilir
- Grafikler sunucuda SVG olarak üretilir — internet bağlantısı gerekmez

---

## Hızlı deneme

Uygulamayı dolu haliyle görmek için (gerçek verinize dokunmaz):

```bash
python demo_data.py --run
```

22 haftalık örnek antrenman üretir ve `data/demo/` klasöründeki ayrı bir veritabanıyla başlatır.
Silmek için o klasörü kaldırmanız yeterlidir.
