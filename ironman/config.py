# -*- coding: utf-8 -*-
"""Sabitler: branşlar, seviye eşikleri, renkler, yollar."""
from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = Path(os.environ.get("IRONMAN_DATA_DIR", BASE_DIR / "data"))
DB_PATH = DATA_DIR / "ironman.db"
BACKUP_DIR = DATA_DIR / "backups"
EXPORT_DIR = DATA_DIR / "exports"
BACKUP_KEEP = 10

APP_NAME = "IRONMAN Yolculuğu"
APP_TAGLINE = "Sıfırdan Full Ironman'e — Antrenman & Seviye Takip Sistemi"

# --------------------------------------------------------------------------
# Branşlar
# --------------------------------------------------------------------------
SPORTS = ("swim", "bike", "run")

SPORT_META = {
    "swim": {
        "key": "swim", "label": "Yüzme", "icon": "🏊", "color": "#0284c7",
        "soft": "#e0f2fe", "race_km": 3.8,
        # pace birimi: dk/100 m
        "pace_unit": "/100m", "pace_divisor": 0.1,
        "decimals": 2,
    },
    "bike": {
        "key": "bike", "label": "Bisiklet", "icon": "🚴", "color": "#d97706",
        "soft": "#fef3c7", "race_km": 180.0,
        "pace_unit": None, "pace_divisor": None,   # bisiklette hız kullanılır
        "decimals": 1,
    },
    "run": {
        "key": "run", "label": "Koşu", "icon": "🏃", "color": "#059669",
        "soft": "#d1fae5", "race_km": 42.2,
        "pace_unit": "/km", "pace_divisor": 1.0,
        "decimals": 1,
    },
}

SPORT_ORDER = ["swim", "bike", "run"]
ACCENT = "#e11d48"

LEVEL_NAMES = [
    "Başlangıç",
    "Temel Dayanıklılık",
    "Sprint Triathlon",
    "Olympic Triathlon",
    "70.3 / Half Ironman",
    "Ironman Hazırlığı",
    "IRONMAN",
]
MAX_LEVEL = 6

# --------------------------------------------------------------------------
# Seviye eşikleri  (Excel'deki AYARLAR sayfasının birebir karşılığı)
#   distance_km   : tek antrenmanda ulaşılması gereken mesafe  → "Mesafe Seviyesi"
#   volume_30d_km : son 30 gündeki toplam hacim                → "Form Seviyesi"
# Eşikler azalmayan sırada olmalıdır.
# --------------------------------------------------------------------------
DEFAULT_THRESHOLDS = {
    "swim": {
        "distance_km":   [0.05, 0.50, 0.75, 1.50, 1.90, 3.00, 3.80],
        "volume_30d_km": [0.20, 2.00, 4.00, 8.00, 12.0, 20.0, 28.0],
    },
    "bike": {
        "distance_km":   [10, 30, 30, 40, 90, 120, 180],
        "volume_30d_km": [20, 60, 100, 160, 280, 400, 500],
    },
    "run": {
        "distance_km":   [1, 5, 5, 10, 21.1, 25, 42.2],
        "volume_30d_km": [4, 20, 30, 50, 80, 110, 140],
    },
}

# Seviye tablosunun açıklayıcı metinleri (SEVİYELER sayfası)
LEVEL_DETAIL = {
    "swim": [
        ("50–100 m", "—", "Suda rahatlık + temel nefes kontrolü; 50–100 m kesintisiz.",
         "Havuzda haftada 2 seans, teknik odaklı başla."),
        ("500 m", "—", "500 m kesintisiz; nefes ritmi (3'te bir) oturmuş.",
         "Kulaç sayısını (SWOLF) takip et."),
        ("750 m  (Sprint yarış)", "≤ 20 dk", "750 m kesintisiz; sprint triatlon yüzme ayağı mesafesi.",
         "İlk açık su denemeni burada yap."),
        ("1.500 m  (Olympic)", "≤ 40 dk", "1.500 m kesintisiz; açık suda en az 1 antrenman.",
         "Sighting (baş kaldırma) tekniğini çalış."),
        ("1.900 m  (70.3)", "≤ 50 dk", "1.900 m kesintisiz; wetsuit ile açık su tecrübesi.",
         "Toplu start provası yap."),
        ("3.000–4.000 m", "—", "3.000 m+ antrenman seti; haftada 2–3 yüzme sürekliliği.",
         "Uzun set + teknik seti dengesi kur."),
        ("3.800 m  (IRONMAN)", "≤ 1s 50 dk", "3.800 m kesintisiz, açık suda, yarış temposunda.",
         "Yarış öncesi en az 2 kez tam mesafe."),
    ],
    "bike": [
        ("10 km", "—", "10 km rahat sürüş; temel bisiklet hakimiyeti ve fren/vites kullanımı.",
         "Kask + temel bakım bilgisi şart."),
        ("30 km", "—", "30 km kesintisiz; kadans 85–95 aralığında tutulabiliyor.",
         "Haftada 1 uzun sürüş ekle."),
        ("20 km yarış · 30 km taban", "≤ 45 dk", "Sprint yarış ayağı 20 km; antrenman tabanı 30 km korunuyor.",
         "Yarış temposunu (FTP %85–90) dene."),
        ("40 km  (Olympic)", "≤ 1s 30 dk", "40 km tempo sürüşü; aero pozisyonda rahatlık.",
         "Bidon/jel alma pratiği yap."),
        ("90 km  (70.3)", "≤ 3s 30 dk", "90 km kesintisiz; beslenme stratejisi denendi.",
         "Saatte 60–80 g karbonhidrat hedefle."),
        ("120–180 km", "—", "120 km+ uzun sürüş; haftada 1 uzun bisiklet sürekliliği.",
         "Uzun sürüş sonrası brick koşusu ekle."),
        ("180 km  (IRONMAN)", "≤ 7s", "180 km yarış temposu + hemen ardından koşabilme.",
         "Yarış öncesi en az 2 kez 160 km+."),
    ],
    "run": [
        ("1–2 km  (koş/yürü)", "—", "Koş/yürü kombinasyonu ile toplam 1–2 km.",
         "Haftada 3 seans, 1'/2' koş-yürü ile başla."),
        ("5 km", "—", "5 km kesintisiz koşu.",
         "Haftalık hacmi %10'dan fazla artırma."),
        ("5 km  (Sprint yarış)", "≤ 35 dk", "Sprint yarış ayağı; bisiklet sonrası koşabilme (brick).",
         "Haftada 1 brick koşusu ekle."),
        ("10 km  (Olympic)", "≤ 60 dk", "10 km kesintisiz; tempo ve interval çalışması var.",
         "Uzun koşuyu 12–14 km'ye taşı."),
        ("21,1 km  (70.3)", "≤ 2s 15 dk", "Yarı maraton mesafesi kesintisiz.",
         "Koşu ayakkabısı rotasyonu kur."),
        ("25–35 km", "—", "Uzun koşu 25 km+; brick koşuları düzenli hale geldi.",
         "Yağ yakım (düşük tempo) koşuları ekle."),
        ("42,2 km  (IRONMAN)", "≤ 5s", "Maraton mesafesi; 180 km bisiklet sonrası tamamlanabilir.",
         "Yarışta ilk 10 km'yi kontrollü koş."),
    ],
}

WORKOUT_KINDS = [
    "Teknik", "Dayanıklılık", "Interval", "Tempo", "Uzun",
    "Toparlanma", "Brick", "Açık Su", "Kuvvet", "Yarış",
]

# --------------------------------------------------------------------------
# Varsayılan hedefler (HEDEFLER sayfası)  — (ad, branş, tip, hedef, tarih, not)
# tip: "distance" → target_value km ;  "event" → (swim, bike, run) km
# --------------------------------------------------------------------------
DEFAULT_GOALS = [
    ("5 km kesintisiz koşu", "run", "distance", 5.0, None, "2026-10-31",
     "En uzun tek koşu ≥ 5 km olduğunda tamamlanır."),
    ("50 km bisiklet", "bike", "distance", 50.0, None, "2026-11-30",
     "En uzun tek sürüş ≥ 50 km olduğunda tamamlanır."),
    ("1.000 m kesintisiz yüzme", "swim", "distance", 1.0, None, "2026-12-31",
     "En uzun tek yüzme ≥ 1,00 km olduğunda tamamlanır."),
    ("10 km koşu", "run", "distance", 10.0, None, "2027-03-31",
     "En uzun tek koşu ≥ 10 km olduğunda tamamlanır."),
    ("Sprint Triatlon — 0,75 / 20 / 5 km", "tri", "event", None, (0.75, 20, 5), "2027-05-31",
     "Üç ayağın hedef mesafesi de tamamlandığında yarışa hazırsınız."),
    ("Olympic Triatlon — 1,5 / 40 / 10 km", "tri", "event", None, (1.5, 40, 10), "2027-09-30",
     "Üç ayağın hedef mesafesi de tamamlandığında yarışa hazırsınız."),
    ("1.900 m yüzme", "swim", "distance", 1.9, None, "2027-12-31",
     "70.3 yüzme ayağı mesafesi."),
    ("21,1 km koşu (yarı maraton)", "run", "distance", 21.1, None, "2028-03-31",
     "70.3 koşu ayağı mesafesi."),
    ("90 km bisiklet", "bike", "distance", 90.0, None, "2028-04-30",
     "70.3 bisiklet ayağı mesafesi."),
    ("70.3 Half Ironman — 1,9 / 90 / 21,1 km", "tri", "event", None, (1.9, 90, 21.1), "2028-09-30",
     "Üç ayağın hedef mesafesi de tamamlandığında yarışa hazırsınız."),
    ("3.800 m yüzme", "swim", "distance", 3.8, None, "2029-03-31",
     "IRONMAN yüzme ayağı mesafesi."),
    ("42,2 km koşu (maraton)", "run", "distance", 42.2, None, "2029-05-31",
     "IRONMAN koşu ayağı mesafesi."),
    ("180 km bisiklet", "bike", "distance", 180.0, None, "2029-06-30",
     "IRONMAN bisiklet ayağı mesafesi."),
    ("🏆 FULL IRONMAN — 3,8 / 180 / 42,2 km", "tri", "event", None, (3.8, 180, 42.2), "2029-09-30",
     "NİHAİ HEDEF. Üç ayağın da tamamlanması + yarış kaydı."),
]

DEFAULT_SETTINGS = {
    "athlete_name": "Birkan",
    "start_date": "2026-08-24",
    "weekly_hour_target": "6",
    "auto_backup": "1",
}
