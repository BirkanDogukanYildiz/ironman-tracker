# -*- coding: utf-8 -*-
"""IRONMAN Antrenman & Seviye Takip Sistemi — Flask uygulaması.

Çalıştırmak için:   python app.py
Telefondan erişim:  python app.py --host 0.0.0.0
"""
from __future__ import annotations

import argparse
import csv
import io
import os
import secrets
import socket
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

from flask import (Flask, abort, flash, g, jsonify, redirect, render_template,
                   request, send_file, session, url_for)

from ironman import __version__, backup, body, charts, db, network
from ironman.config import (APP_NAME, APP_TAGLINE, DATA_DIR, EXPORT_DIR, LEVEL_NAMES,
                            MAX_LEVEL, SPORT_META, SPORT_ORDER, WORKOUT_KINDS)
from ironman.excel_export import build_workbook
from ironman.importer import read_csv
from ironman.stats import (BrickSession, brick_sessions, build_snapshot, evaluate_goals,
                           fmt_hours, fmt_pace, monday_of, next_goal, pace_minutes,
                           parse_date, speed_kmh, weekly_rows)

TMP_DIR = DATA_DIR / "tmp"
app = Flask(__name__)
app.jinja_env.trim_blocks = True
app.jinja_env.lstrip_blocks = True


# --------------------------------------------------------------------------
# Bağlantı yaşam döngüsü
# --------------------------------------------------------------------------
def get_db():
    if "db" not in g:
        g.db = db.connect()
    return g.db


@app.teardown_appcontext
def close_db(exc):
    conn = g.pop("db", None)
    if conn is not None:
        conn.close()


def bootstrap() -> None:
    for d in (DATA_DIR, TMP_DIR, EXPORT_DIR):
        Path(d).mkdir(parents=True, exist_ok=True)
    conn = db.connect()
    try:
        db.init_db(conn)
        settings = db.get_settings(conn)
        key = settings.get("secret_key")
        if not key:
            key = secrets.token_hex(32)
            db.set_setting(conn, "secret_key", key)
        app.secret_key = key
        backup.maybe_auto_backup(settings.get("auto_backup", "1") == "1")
    finally:
        conn.close()


# --------------------------------------------------------------------------
# Jinja yardımcıları
# --------------------------------------------------------------------------
@app.template_filter("tr")
def _tr(value, decimals=1):
    if value is None:
        return "—"
    return charts.tr_num(float(value), decimals)


@app.template_filter("pct")
def _pct(value, decimals=1):
    if value is None:
        return "—"
    return "%" + charts.tr_num(float(value) * 100, decimals)


@app.template_filter("signed_pct")
def _signed_pct(value, decimals=1):
    if value is None:
        return "—"
    sign = "+" if value >= 0 else "−"
    return f"{sign}%{charts.tr_num(abs(float(value)) * 100, decimals)}"


@app.template_filter("gun")
def _gun(value, short=False):
    d = parse_date(value)
    if not d:
        return "—"
    return f"{d.day:02d}.{d.month:02d}" if short else f"{d.day:02d}.{d.month:02d}.{d.year}"


@app.template_filter("pace")
def _pace(value):
    return fmt_pace(value)


@app.template_filter("saat")
def _saat(minutes):
    return fmt_hours(minutes or 0)


@app.context_processor
def inject_globals():
    return {
        "APP_NAME": APP_NAME,
        "APP_TAGLINE": APP_TAGLINE,
        "VERSION": __version__,
        "SPORT_META": SPORT_META,
        "SPORT_ORDER": SPORT_ORDER,
        "LEVEL_NAMES": LEVEL_NAMES,
        "MAX_LEVEL": MAX_LEVEL,
        "KINDS": WORKOUT_KINDS,
        "today": date.today(),
        "charts": charts,
    }


# --------------------------------------------------------------------------
# Ortak veri
# --------------------------------------------------------------------------
def load_context():
    conn = get_db()
    workouts = db.list_workouts(conn)
    thresholds = db.get_thresholds(conn)
    snap = build_snapshot(workouts, thresholds)
    return conn, workouts, thresholds, snap


def _f(name, cast=float, default=None):
    raw = (request.form.get(name) or "").strip().replace(",", ".")
    if raw == "":
        return default
    try:
        return cast(float(raw)) if cast is int else cast(raw)
    except (TypeError, ValueError):
        return default


def _form_to_workout() -> dict:
    day = (request.form.get("date") or "").strip()
    return {
        "date": day,
        "sport": request.form.get("sport"),
        "kind": (request.form.get("kind") or "").strip() or None,
        "distance_km": _f("distance_km", float, 0.0) or 0.0,
        "duration_min": _f("duration_min", float, 0.0) or 0.0,
        "avg_hr": _f("avg_hr", int),
        "rpe": _f("rpe", int),
        "elevation_m": _f("elevation_m", int),
        "calories": _f("calories", int),
        "is_brick": 1 if request.form.get("is_brick") else 0,
        "is_openwater": 1 if request.form.get("is_openwater") else 0,
        "notes": (request.form.get("notes") or "").strip() or None,
        "source": "manual",
        "external_id": None,
    }


def _validate(data: dict) -> list[str]:
    errors = []
    if not parse_date(data["date"]):
        errors.append("Geçerli bir tarih girin.")
    if data["sport"] not in SPORT_ORDER:
        errors.append("Branş seçin.")
    if (data["distance_km"] or 0) <= 0 and (data["duration_min"] or 0) <= 0:
        errors.append("Mesafe veya süreden en az biri girilmeli.")
    if (data["distance_km"] or 0) < 0 or (data["duration_min"] or 0) < 0:
        errors.append("Negatif değer girilemez.")
    if data["rpe"] is not None and not (1 <= data["rpe"] <= 10):
        errors.append("RPE 1–10 arasında olmalı.")
    return errors


# ==========================================================================
# DASHBOARD
# ==========================================================================
@app.route("/")
def dashboard():
    conn, workouts, thresholds, snap = load_context()
    goals = db.list_goals(conn)
    statuses = evaluate_goals(goals, workouts, snap.sports)
    nxt = next_goal(statuses)
    weeks = weekly_rows(workouts, weeks=12)
    settings = db.get_settings(conn)

    week_target = float(settings.get("weekly_hour_target") or 0)
    ctx = {
        "snap": snap,
        "weeks": weeks,
        "next_goal": nxt,
        "goal_statuses": statuses,
        "done_goals": sum(1 for s in statuses if s.is_done),
        "settings": settings,
        "chart_hours": charts.line_chart(
            [w.label for w in weeks], [w.hours for w in weeks], "#e11d48",
            unit="saat", target=week_target or None),
        "chart_stack": charts.stacked_chart(
            [w.label for w in weeks],
            [{"name": SPORT_META[s]["label"], "color": SPORT_META[s]["color"],
              "values": [getattr(w, f"{s}_km") for w in weeks]} for s in SPORT_ORDER]),
        "chart_progress": charts.hbar_chart([
            {"label": f'{SPORT_META[s]["icon"]} {SPORT_META[s]["label"]}',
             "value": snap.sports[s].level.ironman_pct,
             "color": SPORT_META[s]["color"],
             "note": (f"Seviye {snap.sports[s].level.current_level} · "
                      f"maks {snap.sports[s].max_display} / "
                      f"{charts.tr_num(SPORT_META[s]['race_km'], SPORT_META[s]['decimals'])} km")}
            for s in SPORT_ORDER]),
        "chart_split": charts.donut_chart([
            {"label": SPORT_META[s]["label"], "value": snap.sports[s].total_hours,
             "color": SPORT_META[s]["color"]} for s in SPORT_ORDER]),
        "recent": db.list_workouts(conn, limit=8),
    }
    bseries = body.build_series(db.list_measurements(conn), **_body_context(conn))
    ctx["body"] = body.summarize(bseries)
    ctx["body_spark"] = (charts.sparkline([m.weight_kg or 0 for m in bseries], "#e11d48")
                         if len(bseries) > 1 else None)
    return render_template("dashboard.html", **ctx)


# ==========================================================================
# ANTRENMANLAR
# ==========================================================================
@app.route("/antrenmanlar")
def workouts_list():
    conn = get_db()
    page = max(1, request.args.get("sayfa", 1, type=int))
    per = 50
    filters = {
        "sport": request.args.get("brans") or None,
        "date_from": request.args.get("baslangic") or None,
        "date_to": request.args.get("bitis") or None,
        "brick_only": request.args.get("brick") == "1",
        "search": request.args.get("q") or None,
    }
    total = db.count_workouts(conn, **filters)
    rows = db.list_workouts(conn, limit=per, offset=(page - 1) * per, **filters)
    for w in rows:
        w["_pace"] = pace_minutes(w["sport"], w["distance_km"], w["duration_min"])
        w["_speed"] = speed_kmh(w["distance_km"], w["duration_min"])
    agg = {
        "count": total,
        "km": sum(float(w["distance_km"] or 0) for w in db.list_workouts(conn, **filters)),
        "min": sum(float(w["duration_min"] or 0) for w in db.list_workouts(conn, **filters)),
    }
    return render_template("workouts.html", rows=rows, total=total, page=page,
                           pages=max(1, (total + per - 1) // per), filters=filters, agg=agg)


@app.route("/antrenmanlar/yeni", methods=["GET", "POST"])
def workout_new():
    if request.method == "POST":
        data = _form_to_workout()
        errors = _validate(data)
        if errors:
            for e in errors:
                flash(e, "error")
            return render_template("workout_form.html", w=data, mode="new")
        data["date"] = parse_date(data["date"]).isoformat()
        db.add_workout(get_db(), data)
        flash("Antrenman kaydedildi. 💪", "ok")
        if request.form.get("save_and_new"):
            return redirect(url_for("workout_new", date=data["date"], sport=data["sport"]))
        return redirect(url_for("workouts_list"))

    prefill = {
        "date": request.args.get("date") or date.today().isoformat(),
        "sport": request.args.get("sport") or "run",
        "kind": "", "distance_km": "", "duration_min": "", "avg_hr": "", "rpe": "",
        "elevation_m": "", "calories": "", "is_brick": 0, "is_openwater": 0, "notes": "",
    }
    return render_template("workout_form.html", w=prefill, mode="new")


@app.route("/antrenmanlar/<int:wid>/duzenle", methods=["GET", "POST"])
def workout_edit(wid):
    conn = get_db()
    row = db.get_workout(conn, wid)
    if not row:
        abort(404)
    if request.method == "POST":
        data = _form_to_workout()
        errors = _validate(data)
        if errors:
            for e in errors:
                flash(e, "error")
            data["id"] = wid
            return render_template("workout_form.html", w=data, mode="edit")
        data["date"] = parse_date(data["date"]).isoformat()
        data.pop("external_id", None)
        data.pop("source", None)
        db.update_workout(conn, wid, data)
        flash("Antrenman güncellendi.", "ok")
        return redirect(url_for("workouts_list"))
    return render_template("workout_form.html", w=row, mode="edit")


@app.route("/antrenmanlar/<int:wid>/sil", methods=["POST"])
def workout_delete(wid):
    db.delete_workout(get_db(), wid)
    flash("Antrenman silindi.", "ok")
    return redirect(request.form.get("next") or url_for("workouts_list"))


# ==========================================================================
# SEVİYELER · BRANŞLAR · HAFTALIK · BRICK
# ==========================================================================
@app.route("/seviyeler")
def levels_page():
    conn, workouts, thresholds, snap = load_context()
    from ironman.config import LEVEL_DETAIL
    from ironman.levels import level_from

    table = {}
    for sport in SPORT_ORDER:
        st = snap.sports[sport]
        rows = []
        for lvl in range(7):
            crit = thresholds[sport]["distance_km"][lvl]
            reached = st.max_km >= crit
            first = None
            if reached:
                days = [parse_date(w["date"]) for w in workouts
                        if w["sport"] == sport and float(w["distance_km"] or 0) >= crit
                        and parse_date(w["date"])]
                first = min(days) if days else None
            label, target_time, criterion, tip = LEVEL_DETAIL[sport][lvl]
            rows.append({
                "level": lvl, "name": LEVEL_NAMES[lvl], "label": label,
                "criterion_km": crit, "target_time": target_time,
                "criterion": criterion, "tip": tip, "current": st.max_km,
                "progress": min(1.0, st.max_km / crit) if crit else 0,
                "done": reached, "first_date": first,
                "volume_km": thresholds[sport]["volume_30d_km"][lvl],
                "volume_done": st.km_30d >= thresholds[sport]["volume_30d_km"][lvl],
            })
        table[sport] = rows
    return render_template("levels.html", snap=snap, table=table)


@app.route("/branslar")
def progress_page():
    conn, workouts, thresholds, snap = load_context()
    weeks = weekly_rows(workouts, weeks=26)
    spark = {s: charts.sparkline([getattr(w, f"{s}_km") for w in weeks],
                                 SPORT_META[s]["color"]) for s in SPORT_ORDER}
    trend = {}
    for s in SPORT_ORDER:
        rows = [w for w in workouts if w["sport"] == s and float(w["distance_km"] or 0) > 0]
        rows = sorted(rows, key=lambda w: str(w["date"]))[-30:]
        if SPORT_META[s]["pace_unit"]:
            values = [pace_minutes(s, w["distance_km"], w["duration_min"]) or 0 for w in rows]
            unit = f"dk{SPORT_META[s]['pace_unit']}"
        else:
            values = [speed_kmh(w["distance_km"], w["duration_min"]) or 0 for w in rows]
            unit = "km/s"
        labels = [f"{parse_date(w['date']).day:02d}.{parse_date(w['date']).month:02d}"
                  for w in rows if parse_date(w["date"])]
        trend[s] = {"chart": charts.line_chart(labels, values, SPORT_META[s]["color"],
                                               unit=unit, decimals=2, height=200),
                    "unit": unit}
    return render_template("progress.html", snap=snap, spark=spark, trend=trend,
                           thresholds=thresholds)


@app.route("/haftalik")
def weekly_page():
    conn, workouts, thresholds, snap = load_context()
    n = request.args.get("hafta", 26, type=int)
    n = max(4, min(260, n))
    settings = db.get_settings(conn)
    rows = weekly_rows(workouts, weeks=n, start_date=parse_date(settings.get("start_date")))
    shown = list(reversed(rows))
    target = float(settings.get("weekly_hour_target") or 0)
    chart = charts.stacked_chart(
        [w.label for w in rows],
        [{"name": SPORT_META[s]["label"], "color": SPORT_META[s]["color"],
          "values": [getattr(w, f"{s}_km") for w in rows]} for s in SPORT_ORDER])
    chart_h = charts.bar_chart([w.label for w in rows], [w.hours for w in rows],
                               "#1e293b", unit="saat")
    return render_template("weekly.html", rows=shown, chart=chart, chart_hours=chart_h,
                           n=n, target=target, snap=snap, this_monday=monday_of(date.today()))


@app.route("/brick")
def brick_page():
    conn, workouts, thresholds, snap = load_context()
    sessions = brick_sessions(workouts)
    labels = [f"{s.day.day:02d}.{s.day.month:02d}" for s in reversed(sessions)][-15:]
    values = [(s.pace_delta_sec or 0) for s in reversed(sessions)][-15:]
    chart = (charts.bar_chart(labels, values, "#e11d48", unit="sn/km", decimals=0,
                              show_values=True, height=240, neg_color="#16a34a")
             if sessions else charts.empty_state(height=240))
    return render_template("brick.html", sessions=sessions, chart=chart, snap=snap)


# ==========================================================================
# HEDEFLER
# ==========================================================================
@app.route("/hedefler")
def goals_page():
    conn, workouts, thresholds, snap = load_context()
    statuses = evaluate_goals(db.list_goals(conn), workouts, snap.sports)
    return render_template("goals.html", statuses=statuses, snap=snap)


@app.route("/hedefler/<int:gid>/guncelle", methods=["POST"])
def goal_update(gid):
    conn = get_db()
    payload = {}
    td = (request.form.get("target_date") or "").strip()
    payload["target_date"] = parse_date(td).isoformat() if parse_date(td) else None
    cd = (request.form.get("completed_date") or "").strip()
    payload["completed_date"] = parse_date(cd).isoformat() if parse_date(cd) else None
    db.update_goal(conn, gid, payload)
    flash("Hedef güncellendi.", "ok")
    return redirect(url_for("goals_page"))


@app.route("/hedefler/yeni", methods=["POST"])
def goal_new():
    conn = get_db()
    name = (request.form.get("name") or "").strip()
    sport = request.form.get("sport")
    target = _f("target_value", float)
    td = parse_date(request.form.get("target_date") or "")
    if not name or sport not in SPORT_ORDER or not target:
        flash("Hedef adı, branş ve hedef mesafe zorunlu.", "error")
        return redirect(url_for("goals_page"))
    db.add_goal(conn, {"name": name, "sport": sport, "goal_type": "distance",
                       "target_value": target,
                       "target_date": td.isoformat() if td else None,
                       "note": (request.form.get("note") or "").strip()})
    flash("Hedef eklendi.", "ok")
    return redirect(url_for("goals_page"))


@app.route("/hedefler/<int:gid>/sil", methods=["POST"])
def goal_delete(gid):
    db.delete_goal(get_db(), gid)
    flash("Hedef silindi.", "ok")
    return redirect(url_for("goals_page"))


# ==========================================================================
# İÇE AKTARMA
# ==========================================================================
@app.route("/ice-aktar", methods=["GET", "POST"])
def import_page():
    if request.method == "GET":
        return render_template("import.html", preview=None)

    token = request.form.get("token")
    default_kind = (request.form.get("default_kind") or "").strip()
    mark_brick = bool(request.form.get("mark_brick"))

    if token:  # onay adımı
        path = TMP_DIR / f"{Path(token).name}.csv"
        if not path.exists():
            flash("Yükleme süresi doldu, dosyayı tekrar seçin.", "error")
            return redirect(url_for("import_page"))
        preview = read_csv(path.read_bytes(), default_kind=default_kind, mark_brick=mark_brick)
        rows = [r.data for r in preview.rows if r.ok]
        added, skipped = db.bulk_add_workouts(get_db(), rows)
        path.unlink(missing_ok=True)
        flash(f"{added} antrenman eklendi · {skipped} mükerrer atlandı · "
              f"{preview.skipped} satır okunamadı.", "ok")
        return redirect(url_for("workouts_list"))

    upload = request.files.get("file")
    if not upload or not upload.filename:
        flash("Bir CSV dosyası seçin.", "error")
        return redirect(url_for("import_page"))
    raw = upload.read()
    if len(raw) > 20 * 1024 * 1024:
        flash("Dosya çok büyük (en fazla 20 MB).", "error")
        return redirect(url_for("import_page"))

    preview = read_csv(raw, default_kind=default_kind, mark_brick=mark_brick)
    token = uuid.uuid4().hex
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    (TMP_DIR / f"{token}.csv").write_bytes(raw)
    for old in TMP_DIR.glob("*.csv"):
        if old.stat().st_mtime < (datetime.now() - timedelta(hours=6)).timestamp():
            old.unlink(missing_ok=True)
    return render_template("import.html", preview=preview, token=token,
                           filename=upload.filename, default_kind=default_kind,
                           mark_brick=mark_brick)


# ==========================================================================
# DIŞA AKTARMA
# ==========================================================================
@app.route("/disa-aktar/excel")
def export_excel():
    conn = get_db()
    path = EXPORT_DIR / f"IRONMAN_Takip_{datetime.now():%Y%m%d_%H%M}.xlsx"
    build_workbook(path,
                   workouts=db.list_workouts(conn, order="ASC"),
                   goals=db.list_goals(conn),
                   thresholds=db.get_thresholds(conn),
                   settings=db.get_settings(conn),
                   measurements=db.list_measurements(conn, order="ASC"))
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/disa-aktar/csv")
def export_csv():
    conn = get_db()
    rows = db.list_workouts(conn, order="ASC")
    buf = io.StringIO()
    writer = csv.writer(buf, delimiter=";")
    writer.writerow(["Tarih", "Branş", "Tür", "Mesafe (km)", "Süre (dk)", "Nabız",
                     "RPE", "Yükselti (m)", "Kalori", "Brick", "Açık Su", "Notlar"])
    for w in rows:
        writer.writerow([
            w["date"], SPORT_META[w["sport"]]["label"], w["kind"] or "",
            str(w["distance_km"]).replace(".", ","), str(w["duration_min"]).replace(".", ","),
            w["avg_hr"] or "", w["rpe"] or "", w["elevation_m"] or "", w["calories"] or "",
            "Evet" if w["is_brick"] else "Hayır",
            "Evet" if w["is_openwater"] else "Hayır", w["notes"] or ""])
    data = io.BytesIO(("﻿" + buf.getvalue()).encode("utf-8"))
    return send_file(data, mimetype="text/csv", as_attachment=True,
                     download_name=f"antrenmanlar_{date.today():%Y%m%d}.csv")


@app.route("/disa-aktar/db")
def export_db():
    path = backup.create_backup(tag="indirme")
    if not path:
        flash("Veritabanı bulunamadı.", "error")
        return redirect(url_for("settings_page"))
    return send_file(path, as_attachment=True, download_name=path.name)


# ==========================================================================
# AYARLAR
# ==========================================================================
@app.route("/ayarlar")
def settings_page():
    conn = get_db()
    return render_template("settings.html",
                           settings=db.get_settings(conn),
                           thresholds=db.get_thresholds(conn),
                           backups=backup.list_backups(),
                           info=db.db_stats(conn),
                           measurement_count=db.count_measurements(conn))


@app.route("/ayarlar/kaydet", methods=["POST"])
def settings_save():
    conn = get_db()
    for key in ("athlete_name", "weekly_hour_target", "height_cm"):
        if key in request.form:
            db.set_setting(conn, key, (request.form.get(key) or "").strip().replace(",", "."))
    if request.form.get("sex") in ("male", "female"):
        db.set_setting(conn, "sex", request.form["sex"])
    bd = parse_date(request.form.get("birth_date") or "")
    db.set_setting(conn, "birth_date", bd.isoformat() if bd else "")
    sd = parse_date(request.form.get("start_date") or "")
    if sd:
        db.set_setting(conn, "start_date", sd.isoformat())
    db.set_setting(conn, "auto_backup", "1" if request.form.get("auto_backup") else "0")
    flash("Ayarlar kaydedildi.", "ok")
    return redirect(url_for("settings_page"))


@app.route("/ayarlar/esikler", methods=["POST"])
def thresholds_save():
    conn = get_db()
    data = {}
    for sport in SPORT_ORDER:
        d, v = [], []
        for lvl in range(7):
            d.append(_f(f"{sport}_d_{lvl}", float, 0.0) or 0.0)
            v.append(_f(f"{sport}_v_{lvl}", float, 0.0) or 0.0)
        if any(d[i] > d[i + 1] for i in range(6)) or any(v[i] > v[i + 1] for i in range(6)):
            flash(f"{SPORT_META[sport]['label']}: eşikler azalan sırada olamaz.", "error")
            return redirect(url_for("settings_page"))
        data[sport] = {"distance_km": d, "volume_30d_km": v}
    db.save_thresholds(conn, data)
    flash("Seviye eşikleri güncellendi.", "ok")
    return redirect(url_for("settings_page"))


@app.route("/ayarlar/esikler/sifirla", methods=["POST"])
def thresholds_reset():
    db.reset_thresholds(get_db())
    flash("Eşikler varsayılana döndürüldü.", "ok")
    return redirect(url_for("settings_page"))


@app.route("/ayarlar/yedek", methods=["POST"])
def backup_now():
    path = backup.create_backup(tag="elle")
    flash(f"Yedek alındı: {path.name}" if path else "Yedek alınamadı.",
          "ok" if path else "error")
    return redirect(url_for("settings_page"))


@app.route("/ayarlar/yedek/<name>/indir")
def backup_download(name):
    path = backup.BACKUP_DIR / Path(name).name
    if not path.exists():
        abort(404)
    return send_file(path, as_attachment=True, download_name=path.name)


@app.route("/ayarlar/yedek/<name>/geri-yukle", methods=["POST"])
def backup_restore(name):
    ok = backup.restore_backup(name)
    flash("Yedek geri yüklendi." if ok else "Yedek bulunamadı.", "ok" if ok else "error")
    return redirect(url_for("settings_page"))


@app.route("/ayarlar/temizle", methods=["POST"])
def wipe():
    if (request.form.get("confirm") or "").strip().upper() != "SİL":
        flash("Onay metni eşleşmedi, hiçbir şey silinmedi.", "error")
        return redirect(url_for("settings_page"))
    backup.create_backup(tag="silme-oncesi")
    n = db.wipe_workouts(get_db())
    flash(f"{n} antrenman silindi (silme öncesi yedek alındı).", "ok")
    return redirect(url_for("settings_page"))


# ==========================================================================
# VÜCUT ÖLÇÜMLERİ
# ==========================================================================
def _body_context(conn):
    s = db.get_settings(conn)
    height = None
    try:
        height = float(s.get("height_cm") or 0) or None
    except ValueError:
        height = None
    return {"height_cm": height, "sex": s.get("sex") or "male",
            "birth_date": s.get("birth_date") or None}


def _form_to_measurement() -> dict:
    day = (request.form.get("date") or "").strip()
    return {
        "date": day,
        "weight_kg": _f("weight_kg", float),
        "waist_cm": _f("waist_cm", float),
        "neck_cm": _f("neck_cm", float),
        "shoulder_cm": _f("shoulder_cm", float),
        "hip_cm": _f("hip_cm", float),
        "height_cm": _f("height_cm", float),
        "body_fat_pct": _f("body_fat_pct", float),
        "notes": (request.form.get("notes") or "").strip() or None,
    }


def _validate_measurement(data: dict) -> list[str]:
    errors = []
    if not parse_date(data["date"]):
        errors.append("Geçerli bir tarih girin.")
    numeric = ("weight_kg", "waist_cm", "neck_cm", "shoulder_cm", "hip_cm", "height_cm")
    if not any(data.get(k) for k in numeric) and not data.get("body_fat_pct"):
        errors.append("En az bir ölçüm girin (kilo, bel, boyun, omuz…).")
    for key, label, lo, hi in (("weight_kg", "Kilo", 20, 400),
                               ("waist_cm", "Bel", 30, 250),
                               ("neck_cm", "Boyun", 20, 90),
                               ("shoulder_cm", "Omuz", 50, 250),
                               ("hip_cm", "Kalça", 40, 250),
                               ("height_cm", "Boy", 100, 250),
                               ("body_fat_pct", "Yağ oranı", 1, 70)):
        v = data.get(key)
        if v is not None and not (lo <= v <= hi):
            errors.append(f"{label} değeri mantıklı aralıkta değil ({lo}–{hi}).")
    if data.get("waist_cm") and data.get("neck_cm") and data["waist_cm"] <= data["neck_cm"]:
        errors.append("Bel çevresi boyun çevresinden büyük olmalı (yağ oranı formülü için).")
    return errors


@app.route("/olcumler")
def body_page():
    conn = get_db()
    ctx = _body_context(conn)
    rows = db.list_measurements(conn)
    series = body.build_series(rows, **ctx)
    summary = body.summarize(series)

    labels = [f"{m.day.day:02d}.{m.day.month:02d}" for m in series]
    weights = [m.weight_kg or 0 for m in series]
    waists = [m.waist_cm or 0 for m in series]
    fats = [(m.body_fat or 0) * 100 for m in series]

    charts_ = {}
    if series:
        charts_["weight"] = charts.line_chart(labels, weights, "#e11d48", unit="kg",
                                              decimals=1, tight=True, height=230)
        charts_["waist"] = charts.line_chart(labels, waists, "#0284c7", unit="cm",
                                             decimals=1, tight=True, height=230)
        charts_["fat"] = charts.line_chart(labels, fats, "#d97706", unit="%",
                                           decimals=1, tight=True, height=230)
        charts_["comp"] = charts.stacked_chart(labels, [
            {"name": "Yağsız kitle", "color": "#059669",
             "values": [m.lean_mass_kg or 0 for m in series]},
            {"name": "Yağ kitlesi", "color": "#d97706",
             "values": [m.fat_mass_kg or 0 for m in series]},
        ], unit="kg", height=230)

    return render_template("body.html", series=list(reversed(series)), summary=summary,
                           charts_=charts_, settings=db.get_settings(conn),
                           today=date.today().isoformat())


@app.route("/olcumler/yeni", methods=["POST"])
def measurement_new():
    conn = get_db()
    data = _form_to_measurement()
    errors = _validate_measurement(data)
    if errors:
        for e in errors:
            flash(e, "error")
        return redirect(url_for("body_page"))
    day = parse_date(data["date"]).isoformat()
    data["date"] = day
    existed = db.get_measurement_by_date(conn, day)
    db.add_measurement(conn, data)
    flash("Ölçüm güncellendi." if existed else "Ölçüm kaydedildi. 📏", "ok")
    return redirect(url_for("body_page"))


@app.route("/olcumler/<int:mid>/duzenle", methods=["GET", "POST"])
def measurement_edit(mid):
    conn = get_db()
    row = db.get_measurement(conn, mid)
    if not row:
        abort(404)
    if request.method == "POST":
        data = _form_to_measurement()
        errors = _validate_measurement(data)
        if errors:
            for e in errors:
                flash(e, "error")
            data["id"] = mid
            return render_template("body_form.html", m=data)
        clash = db.get_measurement_by_date(conn, parse_date(data["date"]).isoformat())
        if clash and clash["id"] != mid:
            flash("O tarihte zaten bir ölçüm var. Önce onu silin veya düzenleyin.", "error")
            data["id"] = mid
            return render_template("body_form.html", m=data)
        data["date"] = parse_date(data["date"]).isoformat()
        db.update_measurement(conn, mid, data)
        flash("Ölçüm güncellendi.", "ok")
        return redirect(url_for("body_page"))
    return render_template("body_form.html", m=row)


@app.route("/olcumler/<int:mid>/sil", methods=["POST"])
def measurement_delete(mid):
    db.delete_measurement(get_db(), mid)
    flash("Ölçüm silindi.", "ok")
    return redirect(url_for("body_page"))


@app.route("/disa-aktar/olcumler.csv")
def export_measurements_csv():
    conn = get_db()
    ctx = _body_context(conn)
    series = body.build_series(db.list_measurements(conn), **ctx)
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";")
    w.writerow(["Tarih", "Kilo (kg)", "Bel (cm)", "Boyun (cm)", "Omuz (cm)", "Boy (cm)",
                "Yağ oranı (%)", "Yağsız kitle (kg)", "Yağ kitlesi (kg)",
                "Omuz/Bel", "BMI", "Not"])

    def d(v, n=2):
        return "" if v is None else f"{v:.{n}f}".replace(".", ",")

    for m in series:
        w.writerow([m.day.isoformat(), d(m.weight_kg), d(m.waist_cm, 1), d(m.neck_cm, 1),
                    d(m.shoulder_cm, 1), d(m.height_cm, 1),
                    d(m.body_fat * 100 if m.body_fat else None), d(m.lean_mass_kg),
                    d(m.fat_mass_kg), d(m.v_ratio, 3), d(m.bmi), m.notes or ""])
    data = io.BytesIO(("﻿" + buf.getvalue()).encode("utf-8"))
    return send_file(data, mimetype="text/csv", as_attachment=True,
                     download_name=f"olcumler_{date.today():%Y%m%d}.csv")


# ==========================================================================
# TELEFONDAN BAĞLANMA
# ==========================================================================
@app.route("/telefon")
def phone_page():
    port = request.host.split(":")[-1] if ":" in request.host else "5000"
    addresses = network.lan_addresses()
    listening_all = app.config.get("LISTEN_HOST") == "0.0.0.0"
    cards = []
    for a in addresses:
        url = f"http://{a.ip}:{port}"
        cards.append({"ip": a.ip, "url": url, "likely": a.likely, "note": a.note,
                      "qr": network.qr_svg(url)})
    return render_template("phone.html", cards=cards, port=port,
                           listening_all=listening_all,
                           qr_available=bool(cards and cards[0]["qr"]))


# --------------------------------------------------------------------------
@app.route("/saglik")
def health():
    return jsonify({"status": "ok", "version": __version__})


@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code=404,
                           message="Aradığınız sayfa bulunamadı."), 404


@app.errorhandler(500)
def server_error(e):
    return render_template("error.html", code=500,
                           message="Beklenmeyen bir hata oluştu."), 500


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="IRONMAN Antrenman & Seviye Takip Sistemi")
    ap.add_argument("--host", default="127.0.0.1",
                    help="0.0.0.0 verilirse aynı ağdaki telefondan da açılır")
    ap.add_argument("--port", type=int, default=5000)
    ap.add_argument("--telefon", action="store_true",
                    help="--host 0.0.0.0 ile aynı: telefondan erişime aç")
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args()

    host = "0.0.0.0" if args.telefon else args.host
    app.config["LISTEN_HOST"] = host
    bootstrap()

    line = "─" * 54
    print(f"\n  🏊 🚴 🏃  {APP_NAME} v{__version__}")
    print(f"  {line}")
    print(f"  Bu bilgisayarda :  http://127.0.0.1:{args.port}")
    if host == "0.0.0.0":
        addrs = network.lan_addresses()
        if addrs:
            print(f"  Telefondan      :  http://{addrs[0].ip}:{args.port}")
            for extra in addrs[1:]:
                tag = f"   ({extra.note})" if extra.note else ""
                print(f"     ya da        :  http://{extra.ip}:{args.port}{tag}")
        else:
            print("  Telefondan      :  ağ adresi bulunamadı — Wi-Fi bağlı mı?")
        print(f"  QR kod          :  http://127.0.0.1:{args.port}/telefon")
        print(f"  {line}")
        print("  Windows ilk açılışta güvenlik duvarı soracak →")
        print("  «Özel ağlar» kutusunu işaretleyip ERİŞİME İZİN VER deyin.")
    else:
        print(f"  Telefondan açmak için:  --telefon parametresiyle başlatın")
    print(f"  {line}")
    print(f"  Veritabanı      :  {db.DB_PATH}")
    print(f"  Durdurmak için  :  Ctrl+C\n")
    app.run(host=host, port=args.port, debug=args.debug)


bootstrap()

if __name__ == "__main__":
    main()
