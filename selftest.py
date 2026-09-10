# -*- coding: utf-8 -*-
"""Kendi kendini test eder: seviye motoru + tüm rotalar + CSV/Excel akışı."""
from __future__ import annotations

import datetime as dt
import io
import os
import shutil
import sys
import tempfile
from pathlib import Path

TMP = Path(tempfile.mkdtemp(prefix="ironman-test-"))
os.environ["IRONMAN_DATA_DIR"] = str(TMP)

import app as flask_app                                    # noqa: E402
from ironman import backup, body, db                       # noqa: E402
from ironman.importer import read_csv                      # noqa: E402
from ironman.stats import (brick_sessions, build_snapshot,  # noqa: E402
                           evaluate_goals, weekly_rows)

TODAY = dt.date(2026, 9, 8)
FAILS: list[str] = []


def check(name, got, want, tol=1e-6):
    ok = (abs(got - want) <= tol) if isinstance(want, (int, float)) and \
         isinstance(got, (int, float)) else got == want
    print(("  ✓ " if ok else "  ✗ ") + f"{name}: {got}" + ("" if ok else f"  (beklenen {want})"))
    if not ok:
        FAILS.append(name)


SAMPLE = [
    ("2026-08-24", "run",  "Dayanıklılık", 5.0,  32, 148, 5,  40,  350, 0, 0),
    ("2026-08-25", "swim", "Teknik",       0.60, 30, 125, 4,   0,  220, 0, 0),
    ("2026-08-27", "bike", "Dayanıklılık", 25.0, 62, 132, 5, 220,  600, 0, 0),
    ("2026-08-29", "run",  "Uzun",         8.0,  54, 145, 6,  60,  560, 0, 0),
    ("2026-08-31", "swim", "Dayanıklılık", 0.80, 34, 130, 5,   0,  280, 0, 0),
    ("2026-09-01", "run",  "Interval",     6.0,  33, 158, 8,  25,  420, 0, 0),
    ("2026-09-03", "bike", "Tempo",        32.0, 72, 141, 6, 310,  760, 0, 0),
    ("2026-09-05", "bike", "Uzun",         45.0, 108, 135, 6, 480, 1050, 1, 0),
    ("2026-09-05", "run",  "Brick",        3.0,  19, 155, 7,  15,  210, 1, 0),
    ("2026-09-06", "swim", "Açık Su",      1.00, 40, 133, 6,   0,  340, 0, 1),
    ("2026-09-07", "run",  "Toparlanma",   4.0,  26, 128, 3,  20,  280, 0, 0),
    ("2026-09-08", "swim", "Teknik",       0.70, 32, 124, 4,   0,  240, 0, 0),
]


def seed(conn):
    for row in SAMPLE:
        d, sport, kind, km, mn, hr, rpe, elev, kcal, brick, ow = row
        db.add_workout(conn, {
            "date": d, "sport": sport, "kind": kind, "distance_km": km,
            "duration_min": mn, "avg_hr": hr, "rpe": rpe, "elevation_m": elev,
            "calories": kcal, "is_brick": brick, "is_openwater": ow,
            "notes": "test", "source": "manual", "external_id": None})


def main():
    print("\n=== 1. Veritabanı ===")
    conn = db.connect()
    db.init_db(conn)
    seed(conn)
    workouts = db.list_workouts(conn)
    thresholds = db.get_thresholds(conn)
    check("kayıt sayısı", len(workouts), 12)
    check("hedef sayısı", len(db.list_goals(conn)), 14)

    print("\n=== 2. Seviye motoru (Excel sonuçlarıyla karşılaştırma) ===")
    snap = build_snapshot(workouts, thresholds, today=TODAY)
    sw, bk, rn = snap.sports["swim"], snap.sports["bike"], snap.sports["run"]

    check("yüzme maks (km)", sw.max_km, 1.0)
    check("bisiklet maks (km)", bk.max_km, 45.0)
    check("koşu maks (km)", rn.max_km, 8.0)
    check("yüzme 30g hacim", round(sw.km_30d, 2), 3.10)
    check("bisiklet 30g hacim", round(bk.km_30d, 2), 102.0)
    check("koşu 30g hacim", round(rn.km_30d, 2), 26.0)

    check("yüzme mesafe seviyesi", sw.level.distance_level, 2)
    check("bisiklet mesafe seviyesi", bk.level.distance_level, 3)
    check("koşu mesafe seviyesi", rn.level.distance_level, 2)
    check("yüzme form seviyesi", sw.level.volume_level, 1)
    check("bisiklet form seviyesi", bk.level.volume_level, 2)
    check("koşu form seviyesi", rn.level.volume_level, 1)
    check("yüzme mevcut seviye", sw.level.current_level, 1)
    check("bisiklet mevcut seviye", bk.level.current_level, 2)
    check("koşu mevcut seviye", rn.level.current_level, 1)

    check("yüzme seviye içi %", round(sw.level.intra_pct, 4), 0.3333, 1e-3)
    check("bisiklet seviye içi %", round(bk.level.intra_pct, 4), 0.1000, 1e-3)
    check("koşu seviye içi %", round(rn.level.intra_pct, 4), 0.6000, 1e-3)
    check("yüzme ironman %", round(sw.level.ironman_pct, 4), 0.3889, 1e-3)
    check("bisiklet ironman %", round(bk.level.ironman_pct, 4), 0.5167, 1e-3)
    check("koşu ironman %", round(rn.level.ironman_pct, 4), 0.4333, 1e-3)

    check("genel form seviyesi", snap.overall["form_level"], 1)
    check("genel rozet seviyesi", snap.overall["badge_level"], 2)
    check("genel ironman %", round(snap.overall["ironman_pct"], 4), 0.4463, 1e-3)
    check("ham mesafe oranı", round(snap.race_ratio, 4), 0.2343, 1e-3)

    check("toplam süre (saat)", round(snap.total_hours, 4), 9.0333, 1e-3)
    check("toplam mesafe", round(snap.total_km, 2), 131.10)
    check("yüzme ort. pace (dk/100m)", round(sw.avg_pace, 3), 4.387, 1e-2)
    check("koşu ort. pace (dk/km)", round(rn.avg_pace, 3), 6.308, 1e-2)
    check("bisiklet ort. hız", round(bk.avg_speed, 2), 25.29, 1e-2)
    check("son 7 gün bisiklet", round(bk.km_7d, 1), 77.0)

    print("\n=== 3. Haftalık özet ===")
    weeks = {w.start.isoformat(): w for w in weekly_rows(workouts, weeks=12, today=TODAY)}
    w1, w2, w3 = weeks["2026-08-24"], weeks["2026-08-31"], weeks["2026-09-07"]
    check("H1 toplam km", round(w1.total_km, 1), 38.6)
    check("H1 süre (saat)", round(w1.hours, 3), 2.967, 1e-2)
    check("H1 antrenman", w1.count, 4)
    check("H2 toplam km", round(w2.total_km, 1), 87.8)
    check("H2 brick sayısı", w2.bricks, 1)
    check("H2 uzun bisiklet", round(w2.long_bike, 1), 45.0)
    check("H3 süre Δ", round(w3.delta_min, 3), -0.810, 1e-2)

    print("\n=== 4. Brick ===")
    bricks = brick_sessions(workouts)
    check("brick oturumu", len(bricks), 1)
    b = bricks[0]
    check("brick bisiklet km", b.bike_km, 45.0)
    check("brick koşu km", b.run_km, 3.0)
    check("brick pace (dk/km)", round(b.brick_pace, 3), 6.333, 1e-2)
    check("normal pace (dk/km)", round(b.normal_pace, 3), 6.263, 1e-2)
    check("pace farkı (sn)", round(b.pace_delta_sec, 1), 4.2, 0.3)

    print("\n=== 5. Hedefler ===")
    statuses = evaluate_goals(db.list_goals(conn), workouts, snap.sports, today=TODAY)
    by_name = {s.goal["name"]: s for s in statuses}
    check("5 km koşu tamam", by_name["5 km kesintisiz koşu"].state, "done")
    check("5 km koşu tarihi", str(by_name["5 km kesintisiz koşu"].completed_on), "2026-08-24")
    check("50 km bisiklet %", round(by_name["50 km bisiklet"].progress, 2), 0.90)
    check("1000 m yüzme tamam", by_name["1.000 m kesintisiz yüzme"].state, "done")
    check("Sprint tri durumu",
          by_name["Sprint Triatlon — 0,75 / 20 / 5 km"].state, "ready")
    check("Olympic tri ayak",
          by_name["Olympic Triatlon — 1,5 / 40 / 10 km"].current, 1)
    nxt = next((s for s in statuses if s.state != "done"), None)
    check("sonraki hedef", nxt.goal["name"], "50 km bisiklet")

    print("\n=== 5b. Vücut ölçümleri (Excel ile karşılaştırma) ===")
    mrows = db.list_measurements(conn)
    check("tohumlanan ölçüm", len(mrows), 5)
    series = body.build_series(mrows, height_cm=174.0, sex="male")
    # «kg-bel-boyun-yağ%» sayfasındaki beklenen değerler
    EXPECTED = {
        "2026-06-12": (33.22, 71.45, 35.55, 1.159),
        "2026-06-19": (31.00, 72.97, 32.78, 1.193),
        "2026-06-26": (30.42, 73.65, 32.20, 1.204),
        "2026-07-03": (29.84, 75.07, 31.93, 1.238),
        "2026-07-18": (29.26, 76.05, 31.45, 1.255),
    }
    for m in series:
        fat, lean, fatkg, vr = EXPECTED[m.day.isoformat()]
        d = m.day.strftime("%d.%m")
        check(f"{d} yağ oranı %", round(m.body_fat * 100, 2), fat, 0.02)
        check(f"{d} yağsız kitle", round(m.lean_mass_kg, 2), lean, 0.03)
        check(f"{d} yağ kitlesi", round(m.fat_mass_kg, 2), fatkg, 0.03)
        check(f"{d} omuz/bel", round(m.v_ratio, 3), vr, 0.002)
    summary = body.summarize(series, today=dt.date(2026, 7, 20))
    check("toplam bel değişimi", round(summary.total_waist, 1), -7.0)
    check("toplam yağsız kitle değişimi", round(summary.total_lean, 2), 4.60, 0.03)
    check("boy yoksa yağ oranı hesaplanmaz",
          body.build_series(mrows, height_cm=None)[0].body_fat, None)
    check("bel<boyun ise hesaplanmaz",
          body.navy_body_fat(40, 41, 174), None)
    check("kadın formülü kalçasız çalışmaz",
          body.navy_body_fat(80, 33, 165, "female"), None)
    check("kadın formülü kalçayla çalışır",
          round(body.navy_body_fat(80, 33, 165, "female", 100) * 100, 1), 31.9, 0.2)
    check("BMI", round(body.bmi(107.5, 174), 2), 35.51, 0.02)
    # aynı tarihe ikinci giriş üzerine yazar
    db.add_measurement(conn, {"date": "2026-07-18", "weight_kg": 106.0,
                              "waist_cm": 105.0, "neck_cm": 41.0, "shoulder_cm": 133.0})
    check("aynı tarih güncellendi", db.count_measurements(conn), 5)
    check("değer güncellendi",
          db.get_measurement_by_date(conn, "2026-07-18")["weight_kg"], 106.0)
    db.update_measurement(conn, db.get_measurement_by_date(conn, "2026-07-18")["id"],
                          {"weight_kg": 107.5, "waist_cm": 106.0})

    print("\n=== 6. CSV içe aktarma ===")
    strava = (
        "Activity ID,Activity Date,Activity Name,Activity Type,Elapsed Time,Distance,"
        "Average Heart Rate,Elevation Gain,Calories\n"
        "9001,\"Sep 10, 2026, 6:12:00 AM\",Morning Run,Run,2400,10000,152,85,700\n"
        "9002,\"Sep 11, 2026, 7:00:00 AM\",Pool,Pool Swim,2100,1500,128,0,420\n"
        "9003,\"Sep 12, 2026, 8:00:00 AM\",Long ride,Ride,10800,62000,138,640,1500\n"
        "9004,bozuk tarih,Bilinmeyen,Yoga,1200,0,,,\n")
    prev = read_csv(strava.encode())
    check("CSV toplam", prev.total, 4)
    check("CSV geçerli", prev.valid, 3)
    ok_rows = [r.data for r in prev.rows if r.ok]
    check("CSV koşu km", ok_rows[0]["distance_km"], 10.0)
    check("CSV koşu dk", ok_rows[0]["duration_min"], 40.0)
    check("CSV yüzme km", ok_rows[1]["distance_km"], 1.5)
    check("CSV bisiklet km", ok_rows[2]["distance_km"], 62.0)
    added, skipped = db.bulk_add_workouts(conn, ok_rows)
    check("CSV eklenen", added, 3)
    added2, skipped2 = db.bulk_add_workouts(conn, ok_rows)
    check("mükerrer engellendi", skipped2, 3)

    garmin = ("Aktivite Türü;Tarih;Mesafe;Süre;Ortalama Nabız\n"
              "Koşu;13.09.2026;12,5;01:05:30;149\n"
              "Bisiklet;14.09.2026;70,2;02:48:00;136\n")
    prev2 = read_csv(garmin.encode())
    check("Garmin geçerli", prev2.valid, 2)
    g = [r.data for r in prev2.rows if r.ok]
    check("Garmin koşu km", g[0]["distance_km"], 12.5)
    check("Garmin koşu dk", round(g[0]["duration_min"], 1), 65.5)
    check("Garmin bisiklet dk", round(g[1]["duration_min"], 1), 168.0)

    print("\n=== 7. Yedekleme ===")
    p = backup.create_backup(tag="test")
    check("yedek oluştu", bool(p and p.exists()), True)
    check("yedek listesi", len(backup.list_backups()) >= 1, True)
    conn.close()

    print("\n=== 8. Excel dışa aktarma ===")
    from ironman.excel_export import build_workbook
    conn = db.connect()
    out = TMP / "export.xlsx"
    build_workbook(out, workouts=db.list_workouts(conn, order="ASC"),
                   goals=db.list_goals(conn), thresholds=db.get_thresholds(conn),
                   settings=db.get_settings(conn),
                   measurements=db.list_measurements(conn, order="ASC"))
    check("xlsx oluştu", out.exists(), True)
    check("xlsx boyutu > 40 KB", out.stat().st_size > 40_000, True)
    import zipfile
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        check("xlsx grafik sayısı", len([n for n in names if "charts/chart" in n]), 4)
        check("xlsx sayfa sayısı", len([n for n in names if n.startswith("xl/worksheets/sheet")]), 10)
        check("fullCalcOnLoad", 'fullCalcOnLoad="1"' in z.read("xl/workbook.xml").decode(), True)
    conn.close()

    print("\n=== 9. Web rotaları ===")
    flask_app.bootstrap()
    client = flask_app.app.test_client()
    routes = ["/", "/antrenmanlar", "/antrenmanlar/yeni", "/seviyeler", "/branslar",
              "/haftalik", "/haftalik?hafta=52", "/brick", "/hedefler", "/ice-aktar",
              "/ayarlar", "/saglik", "/antrenmanlar?brans=run&brick=1",
              "/olcumler", "/telefon", "/disa-aktar/csv", "/disa-aktar/excel",
              "/disa-aktar/olcumler.csv"]
    for r in routes:
        resp = client.get(r)
        check(f"GET {r}", resp.status_code, 200)
    check("GET /yok-boyle-sayfa", client.get("/yok-boyle-sayfa").status_code, 404)

    print("\n=== 10. Form akışları ===")
    resp = client.post("/antrenmanlar/yeni", data={
        "date": "2026-09-09", "sport": "run", "kind": "Tempo",
        "distance_km": "12,5", "duration_min": "70", "rpe": "7"},
        follow_redirects=True)
    check("POST yeni antrenman", resp.status_code, 200)
    conn = db.connect()
    row = db.list_workouts(conn, date_from="2026-09-09", date_to="2026-09-09")
    check("virgüllü mesafe kaydedildi", row[0]["distance_km"], 12.5)
    wid = row[0]["id"]
    conn.close()

    resp = client.post(f"/antrenmanlar/{wid}/duzenle", data={
        "date": "2026-09-09", "sport": "run", "kind": "Uzun",
        "distance_km": "14", "duration_min": "80"}, follow_redirects=True)
    check("POST düzenle", resp.status_code, 200)
    conn = db.connect()
    check("güncellendi", db.get_workout(conn, wid)["distance_km"], 14.0)
    conn.close()

    resp = client.post("/antrenmanlar/yeni", data={
        "date": "", "sport": "run", "distance_km": "", "duration_min": ""})
    check("boş form reddedildi (200 + hata)", resp.status_code, 200)
    check("hata mesajı gösterildi", "Geçerli bir tarih" in resp.get_data(as_text=True), True)

    resp = client.post(f"/antrenmanlar/{wid}/sil", follow_redirects=True)
    check("POST sil", resp.status_code, 200)
    conn = db.connect()
    check("silindi", db.get_workout(conn, wid), None)
    conn.close()

    data = {"file": (io.BytesIO(strava.encode()), "activities.csv")}
    resp = client.post("/ice-aktar", data=data, content_type="multipart/form-data")
    check("CSV önizleme", resp.status_code, 200)
    check("önizlemede 3 geçerli", "3 geçerli" in resp.get_data(as_text=True), True)

    resp = client.post("/olcumler/yeni", data={
        "date": "2026-08-01", "weight_kg": "106,2", "waist_cm": "104",
        "neck_cm": "41", "shoulder_cm": "134"}, follow_redirects=True)
    check("POST yeni ölçüm", resp.status_code, 200)
    conn = db.connect()
    row = db.get_measurement_by_date(conn, "2026-08-01")
    check("virgüllü kilo kaydedildi", row["weight_kg"], 106.2)
    mid = row["id"]
    conn.close()
    resp = client.post(f"/olcumler/{mid}/duzenle", data={
        "date": "2026-08-01", "weight_kg": "105", "waist_cm": "103",
        "neck_cm": "41", "shoulder_cm": "134"}, follow_redirects=True)
    check("POST ölçüm düzenle", resp.status_code, 200)
    conn = db.connect()
    check("ölçüm güncellendi", db.get_measurement(conn, mid)["weight_kg"], 105.0)
    conn.close()
    resp = client.post("/olcumler/yeni", data={
        "date": "2026-08-05", "waist_cm": "40", "neck_cm": "41"}, follow_redirects=True)
    check("bel<boyun reddedildi", "boyun çevresinden büyük" in resp.get_data(as_text=True), True)
    resp = client.post("/olcumler/yeni", data={"date": "2026-08-06"}, follow_redirects=True)
    check("boş ölçüm reddedildi", "En az bir ölçüm" in resp.get_data(as_text=True), True)
    resp = client.post("/olcumler/yeni", data={
        "date": "2026-08-07", "weight_kg": "900"}, follow_redirects=True)
    check("saçma kilo reddedildi", "mantıklı aralıkta değil" in resp.get_data(as_text=True), True)
    resp = client.post(f"/olcumler/{mid}/sil", follow_redirects=True)
    check("POST ölçüm sil", resp.status_code, 200)
    conn = db.connect()
    check("ölçüm silindi", db.get_measurement(conn, mid), None)
    conn.close()

    resp = client.post("/ayarlar/kaydet", data={
        "athlete_name": "Birkan", "start_date": "2026-08-24",
        "weekly_hour_target": "7", "auto_backup": "1", "height_cm": "174",
        "sex": "male", "birth_date": ""}, follow_redirects=True)
    check("ayar kaydet", resp.status_code, 200)
    conn = db.connect()
    check("boy ayarı kaydedildi", db.get_settings(conn)["height_cm"], "174")
    conn.close()

    bad = {f"{s}_{t}_{i}": v for s in ("swim", "bike", "run")
           for t in ("d", "v") for i, v in enumerate([9, 8, 7, 6, 5, 4, 3])}
    resp = client.post("/ayarlar/esikler", data=bad, follow_redirects=True)
    check("azalan eşik reddedildi", "azalan sırada olamaz" in resp.get_data(as_text=True), True)

    resp = client.post("/ayarlar/temizle", data={"confirm": "yanlış"}, follow_redirects=True)
    check("yanlış onay silmedi", "eşleşmedi" in resp.get_data(as_text=True), True)

    print("\n" + "=" * 58)
    if FAILS:
        print(f"  ✗ {len(FAILS)} TEST BAŞARISIZ:")
        for name in FAILS:
            print(f"      · {name}")
    else:
        print("  ✓ TÜM TESTLER GEÇTİ")
    print("=" * 58 + "\n")
    shutil.rmtree(TMP, ignore_errors=True)
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
