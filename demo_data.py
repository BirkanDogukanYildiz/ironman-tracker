# -*- coding: utf-8 -*-
"""Uygulamayı dolu haliyle görmek için örnek veri üretir.

Gerçek veritabanınıza DOKUNMAZ — ayrı bir klasöre (data/demo) yazar:

    python demo_data.py          # örnek veriyi üret
    python demo_data.py --run    # üret ve demo veriyle uygulamayı başlat

Demo klasörünü silmek yeterlidir; kalıcı hiçbir etkisi yoktur.
"""
from __future__ import annotations

import datetime as dt
import os
import random
import subprocess
import sys
from pathlib import Path

DEMO_DIR = Path(__file__).resolve().parent / "data" / "demo"
os.environ["IRONMAN_DATA_DIR"] = str(DEMO_DIR)

from ironman import db  # noqa: E402


def build(weeks: int = 22, seed: int = 7) -> int:
    random.seed(seed)
    conn = db.connect()
    db.init_db(conn)
    if db.count_workouts(conn):
        n = db.count_workouts(conn)
        conn.close()
        return n

    today = dt.date.today()
    start = today - dt.timedelta(weeks=weeks)
    day = start
    while day <= today:
        if day.weekday() in (0, 2, 4, 5):
            prog = min(1.0, ((day - start).days / 7) / (weeks - 2))
            sport = random.choices(["run", "bike", "swim"], [0.38, 0.34, 0.28])[0]
            if sport == "swim":
                km = round(random.uniform(0.4, 0.8) + prog * 1.2, 2)
                mn = round(km * 10 * random.uniform(2.0, 2.6))
                kind = random.choice(["Teknik", "Dayanıklılık", "Açık Su"])
            elif sport == "bike":
                km = round(random.uniform(18, 30) + prog * 60, 1)
                mn = round(km / random.uniform(23, 28) * 60)
                kind = random.choice(["Dayanıklılık", "Tempo", "Uzun"])
            else:
                km = round(random.uniform(4, 7) + prog * 9, 1)
                mn = round(km * random.uniform(5.6, 6.6))
                kind = random.choice(["Dayanıklılık", "Interval", "Uzun", "Toparlanma"])
            db.add_workout(conn, {
                "date": day.isoformat(), "sport": sport, "kind": kind,
                "distance_km": km, "duration_min": mn,
                "avg_hr": random.randint(122, 162), "rpe": random.randint(3, 8),
                "elevation_m": random.randint(0, 500), "calories": int(mn * 9),
                "is_brick": 0, "is_openwater": 1 if kind == "Açık Su" else 0,
                "notes": "", "source": "demo", "external_id": None})
        day += dt.timedelta(days=1)

    for offset in (7, 21, 35, 49):
        d = (today - dt.timedelta(days=offset)).isoformat()
        bike_km = round(random.uniform(60, 95), 1)
        run_km = round(random.uniform(4, 9), 1)
        db.add_workout(conn, {
            "date": d, "sport": "bike", "kind": "Uzun", "distance_km": bike_km,
            "duration_min": round(bike_km / 26 * 60), "avg_hr": 138, "rpe": 7,
            "elevation_m": 600, "calories": 1800, "is_brick": 1, "is_openwater": 0,
            "notes": "brick sürüşü", "source": "demo", "external_id": None})
        db.add_workout(conn, {
            "date": d, "sport": "run", "kind": "Brick", "distance_km": run_km,
            "duration_min": round(run_km * 6.4), "avg_hr": 156, "rpe": 8,
            "elevation_m": 40, "calories": int(run_km * 62), "is_brick": 1,
            "is_openwater": 0, "notes": "bacaklar ağırdı", "source": "demo",
            "external_id": None})

    n = db.count_workouts(conn)
    conn.close()
    return n


if __name__ == "__main__":
    total = build()
    print(f"✓ {total} örnek antrenman üretildi → {DEMO_DIR}")
    if "--run" in sys.argv:
        print("→ Demo veriyle başlatılıyor: http://127.0.0.1:5000\n")
        subprocess.run([sys.executable, "app.py"], env={**os.environ,
                                                        "IRONMAN_DATA_DIR": str(DEMO_DIR)})
    else:
        print(f"Başlatmak için:\n  IRONMAN_DATA_DIR={DEMO_DIR} python app.py")
        print("  (veya:  python demo_data.py --run)")
        print(f"\nSilmek için bu klasörü kaldırın: {DEMO_DIR}")
