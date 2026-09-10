# -*- coding: utf-8 -*-
"""SQLite bağlantısı, şema, migration ve tohum verisi."""
from __future__ import annotations

import json
import sqlite3
from datetime import datetime, date
from pathlib import Path
from typing import Any, Iterable, Optional

from .config import (
    DB_PATH, DATA_DIR, BACKUP_DIR, EXPORT_DIR,
    DEFAULT_THRESHOLDS, DEFAULT_GOALS, DEFAULT_SETTINGS, INITIAL_MEASUREMENTS,
    LEVEL_NAMES, LEVEL_DETAIL, SPORT_ORDER,
)

SCHEMA_VERSION = 2

SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS workouts (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    date          TEXT    NOT NULL,                 -- ISO  YYYY-MM-DD
    sport         TEXT    NOT NULL CHECK (sport IN ('swim','bike','run')),
    kind          TEXT,
    distance_km   REAL    NOT NULL DEFAULT 0 CHECK (distance_km >= 0),
    duration_min  REAL    NOT NULL DEFAULT 0 CHECK (duration_min >= 0),
    avg_hr        INTEGER,
    rpe           INTEGER CHECK (rpe IS NULL OR (rpe BETWEEN 1 AND 10)),
    elevation_m   INTEGER,
    calories      INTEGER,
    is_brick      INTEGER NOT NULL DEFAULT 0,
    is_openwater  INTEGER NOT NULL DEFAULT 0,
    notes         TEXT,
    source        TEXT    NOT NULL DEFAULT 'manual',
    external_id   TEXT,
    created_at    TEXT    NOT NULL,
    updated_at    TEXT    NOT NULL
);

CREATE INDEX IF NOT EXISTS ix_workouts_date  ON workouts(date DESC);
CREATE INDEX IF NOT EXISTS ix_workouts_sport ON workouts(sport, date);
CREATE INDEX IF NOT EXISTS ix_workouts_brick ON workouts(is_brick, date);
CREATE UNIQUE INDEX IF NOT EXISTS ux_workouts_external
    ON workouts(external_id) WHERE external_id IS NOT NULL;

CREATE TABLE IF NOT EXISTS goals (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    position       INTEGER NOT NULL,
    name           TEXT    NOT NULL,
    sport          TEXT    NOT NULL,               -- swim|bike|run|tri
    goal_type      TEXT    NOT NULL,               -- distance|event
    target_value   REAL,
    target_swim    REAL,
    target_bike    REAL,
    target_run     REAL,
    target_date    TEXT,
    completed_date TEXT,
    note           TEXT,
    archived       INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS thresholds (
    sport          TEXT    NOT NULL,
    level          INTEGER NOT NULL,
    distance_km    REAL    NOT NULL,
    volume_30d_km  REAL    NOT NULL,
    PRIMARY KEY (sport, level)
);

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS measurements (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    date          TEXT    NOT NULL UNIQUE,       -- günde tek ölçüm
    weight_kg     REAL,
    waist_cm      REAL,
    neck_cm       REAL,
    shoulder_cm   REAL,
    hip_cm        REAL,
    height_cm     REAL,                          -- boş ise ayarlardaki boy
    body_fat_pct  REAL,                          -- elle girilirse formülü ezer
    notes         TEXT,
    created_at    TEXT NOT NULL,
    updated_at    TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS ix_measurements_date ON measurements(date DESC);
"""


# --------------------------------------------------------------------------
# Bağlantı
# --------------------------------------------------------------------------
def _dict_factory(cursor, row):
    return {col[0]: row[idx] for idx, col in enumerate(cursor.description)}


def connect(path: Path | str = DB_PATH) -> sqlite3.Connection:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path), detect_types=0)
    conn.row_factory = _dict_factory
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


# --------------------------------------------------------------------------
# Kurulum
# --------------------------------------------------------------------------
def init_db(conn: sqlite3.Connection) -> None:
    for directory in (DATA_DIR, BACKUP_DIR, EXPORT_DIR):
        Path(directory).mkdir(parents=True, exist_ok=True)
    conn.executescript(SCHEMA)
    _seed(conn)
    conn.execute(
        "INSERT INTO meta(key, value) VALUES('schema_version', ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (str(SCHEMA_VERSION),),
    )
    conn.commit()


def _seed(conn: sqlite3.Connection) -> None:
    # eşikler
    if conn.execute("SELECT COUNT(*) c FROM thresholds").fetchone()["c"] == 0:
        rows = []
        for sport in SPORT_ORDER:
            d = DEFAULT_THRESHOLDS[sport]
            for lvl in range(7):
                rows.append((sport, lvl, float(d["distance_km"][lvl]),
                             float(d["volume_30d_km"][lvl])))
        conn.executemany(
            "INSERT INTO thresholds(sport, level, distance_km, volume_30d_km) VALUES (?,?,?,?)",
            rows,
        )

    # hedefler
    if conn.execute("SELECT COUNT(*) c FROM goals").fetchone()["c"] == 0:
        rows = []
        for i, (name, sport, gtype, target, legs, tdate, note) in enumerate(DEFAULT_GOALS):
            s, b, r = (legs if legs else (None, None, None))
            rows.append((i + 1, name, sport, gtype, target, s, b, r, tdate, None, note))
        conn.executemany(
            "INSERT INTO goals(position,name,sport,goal_type,target_value,"
            "target_swim,target_bike,target_run,target_date,completed_date,note) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            rows,
        )

    # ayarlar
    for key, value in DEFAULT_SETTINGS.items():
        conn.execute(
            "INSERT INTO settings(key, value) VALUES(?,?) ON CONFLICT(key) DO NOTHING",
            (key, value),
        )

    # ilk vücut ölçümleri (Excel'den) — yalnızca tablo tamamen boşken
    already = conn.execute(
        "SELECT value FROM meta WHERE key = 'measurements_seeded'").fetchone()
    empty = conn.execute("SELECT COUNT(*) c FROM measurements").fetchone()["c"] == 0
    if empty and not already and INITIAL_MEASUREMENTS:
        ts = now_iso()
        conn.executemany(
            "INSERT OR IGNORE INTO measurements(date, weight_kg, waist_cm, neck_cm,"
            " shoulder_cm, notes, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
            [(d, kg, bel, boyun, omuz, note, ts, ts)
             for d, kg, bel, boyun, omuz, note in INITIAL_MEASUREMENTS],
        )
        conn.execute("INSERT OR REPLACE INTO meta(key, value) VALUES('measurements_seeded','1')")


# --------------------------------------------------------------------------
# Ayarlar
# --------------------------------------------------------------------------
def get_settings(conn) -> dict[str, str]:
    out = dict(DEFAULT_SETTINGS)
    for row in conn.execute("SELECT key, value FROM settings"):
        out[row["key"]] = row["value"]
    return out


def set_setting(conn, key: str, value: str) -> None:
    conn.execute(
        "INSERT INTO settings(key, value) VALUES(?,?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, str(value)),
    )
    conn.commit()


# --------------------------------------------------------------------------
# Eşikler
# --------------------------------------------------------------------------
def get_thresholds(conn) -> dict[str, dict[str, list[float]]]:
    out = {s: {"distance_km": [0.0] * 7, "volume_30d_km": [0.0] * 7} for s in SPORT_ORDER}
    for row in conn.execute("SELECT * FROM thresholds ORDER BY sport, level"):
        if row["sport"] in out and 0 <= row["level"] <= 6:
            out[row["sport"]]["distance_km"][row["level"]] = float(row["distance_km"])
            out[row["sport"]]["volume_30d_km"][row["level"]] = float(row["volume_30d_km"])
    return out


def save_thresholds(conn, data: dict[str, dict[str, list[float]]]) -> None:
    for sport, d in data.items():
        for lvl in range(7):
            conn.execute(
                "INSERT INTO thresholds(sport, level, distance_km, volume_30d_km) "
                "VALUES (?,?,?,?) ON CONFLICT(sport, level) DO UPDATE SET "
                "distance_km = excluded.distance_km, volume_30d_km = excluded.volume_30d_km",
                (sport, lvl, float(d["distance_km"][lvl]), float(d["volume_30d_km"][lvl])),
            )
    conn.commit()


def reset_thresholds(conn) -> None:
    conn.execute("DELETE FROM thresholds")
    _seed(conn)
    conn.commit()


# --------------------------------------------------------------------------
# Antrenmanlar
# --------------------------------------------------------------------------
WORKOUT_FIELDS = (
    "date", "sport", "kind", "distance_km", "duration_min", "avg_hr", "rpe",
    "elevation_m", "calories", "is_brick", "is_openwater", "notes",
    "source", "external_id",
)


def list_workouts(conn, *, sport: str | None = None, date_from: str | None = None,
                  date_to: str | None = None, brick_only: bool = False,
                  search: str | None = None, limit: int | None = None,
                  offset: int = 0, order: str = "DESC") -> list[dict]:
    where, params = [], []
    if sport in SPORT_ORDER:
        where.append("sport = ?"); params.append(sport)
    if date_from:
        where.append("date >= ?"); params.append(date_from)
    if date_to:
        where.append("date <= ?"); params.append(date_to)
    if brick_only:
        where.append("is_brick = 1")
    if search:
        where.append("(IFNULL(notes,'') LIKE ? OR IFNULL(kind,'') LIKE ?)")
        params += [f"%{search}%", f"%{search}%"]
    sql = "SELECT * FROM workouts"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += f" ORDER BY date {order}, id {order}"
    if limit is not None:
        sql += " LIMIT ? OFFSET ?"; params += [limit, offset]
    return list(conn.execute(sql, params))


def count_workouts(conn, **kw) -> int:
    kw.pop("limit", None); kw.pop("offset", None); kw.pop("order", None)
    where, params = [], []
    if kw.get("sport") in SPORT_ORDER:
        where.append("sport = ?"); params.append(kw["sport"])
    if kw.get("date_from"):
        where.append("date >= ?"); params.append(kw["date_from"])
    if kw.get("date_to"):
        where.append("date <= ?"); params.append(kw["date_to"])
    if kw.get("brick_only"):
        where.append("is_brick = 1")
    if kw.get("search"):
        where.append("(IFNULL(notes,'') LIKE ? OR IFNULL(kind,'') LIKE ?)")
        params += [f"%{kw['search']}%", f"%{kw['search']}%"]
    sql = "SELECT COUNT(*) c FROM workouts"
    if where:
        sql += " WHERE " + " AND ".join(where)
    return conn.execute(sql, params).fetchone()["c"]


def get_workout(conn, wid: int) -> Optional[dict]:
    return conn.execute("SELECT * FROM workouts WHERE id = ?", (wid,)).fetchone()


def add_workout(conn, data: dict) -> int:
    payload = {k: data.get(k) for k in WORKOUT_FIELDS}
    payload["source"] = payload.get("source") or "manual"
    payload["is_brick"] = int(bool(payload.get("is_brick")))
    payload["is_openwater"] = int(bool(payload.get("is_openwater")))
    ts = now_iso()
    cols = ", ".join(WORKOUT_FIELDS) + ", created_at, updated_at"
    marks = ", ".join("?" * (len(WORKOUT_FIELDS) + 2))
    cur = conn.execute(
        f"INSERT INTO workouts({cols}) VALUES ({marks})",
        [payload[k] for k in WORKOUT_FIELDS] + [ts, ts],
    )
    conn.commit()
    return cur.lastrowid


def update_workout(conn, wid: int, data: dict) -> None:
    payload = {k: data.get(k) for k in WORKOUT_FIELDS if k in data}
    if "is_brick" in payload:
        payload["is_brick"] = int(bool(payload["is_brick"]))
    if "is_openwater" in payload:
        payload["is_openwater"] = int(bool(payload["is_openwater"]))
    if not payload:
        return
    sets = ", ".join(f"{k} = ?" for k in payload) + ", updated_at = ?"
    conn.execute(
        f"UPDATE workouts SET {sets} WHERE id = ?",
        list(payload.values()) + [now_iso(), wid],
    )
    conn.commit()


def delete_workout(conn, wid: int) -> None:
    conn.execute("DELETE FROM workouts WHERE id = ?", (wid,))
    conn.commit()


def bulk_add_workouts(conn, rows: Iterable[dict]) -> tuple[int, int]:
    """Toplu ekleme. (eklenen, atlanan-mükerrer) döner."""
    added = skipped = 0
    ts = now_iso()
    cols = ", ".join(WORKOUT_FIELDS) + ", created_at, updated_at"
    marks = ", ".join("?" * (len(WORKOUT_FIELDS) + 2))
    for data in rows:
        payload = {k: data.get(k) for k in WORKOUT_FIELDS}
        payload["source"] = payload.get("source") or "csv"
        payload["is_brick"] = int(bool(payload.get("is_brick")))
        payload["is_openwater"] = int(bool(payload.get("is_openwater")))
        try:
            conn.execute(
                f"INSERT INTO workouts({cols}) VALUES ({marks})",
                [payload[k] for k in WORKOUT_FIELDS] + [ts, ts],
            )
            added += 1
        except sqlite3.IntegrityError:
            skipped += 1
    conn.commit()
    return added, skipped


def brick_dates(conn) -> list[str]:
    rows = conn.execute(
        "SELECT DISTINCT date FROM workouts WHERE is_brick = 1 ORDER BY date DESC"
    )
    return [r["date"] for r in rows]


# --------------------------------------------------------------------------
# Hedefler
# --------------------------------------------------------------------------
def list_goals(conn, include_archived: bool = False) -> list[dict]:
    sql = "SELECT * FROM goals"
    if not include_archived:
        sql += " WHERE archived = 0"
    sql += " ORDER BY position, id"
    return list(conn.execute(sql))


def get_goal(conn, gid: int) -> Optional[dict]:
    return conn.execute("SELECT * FROM goals WHERE id = ?", (gid,)).fetchone()


def update_goal(conn, gid: int, data: dict) -> None:
    allowed = ("name", "sport", "goal_type", "target_value", "target_swim",
               "target_bike", "target_run", "target_date", "completed_date",
               "note", "position", "archived")
    payload = {k: data[k] for k in allowed if k in data}
    if not payload:
        return
    sets = ", ".join(f"{k} = ?" for k in payload)
    conn.execute(f"UPDATE goals SET {sets} WHERE id = ?", list(payload.values()) + [gid])
    conn.commit()


def add_goal(conn, data: dict) -> int:
    pos = conn.execute("SELECT IFNULL(MAX(position),0)+1 p FROM goals").fetchone()["p"]
    cur = conn.execute(
        "INSERT INTO goals(position,name,sport,goal_type,target_value,target_swim,"
        "target_bike,target_run,target_date,completed_date,note) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (data.get("position", pos), data["name"], data["sport"], data["goal_type"],
         data.get("target_value"), data.get("target_swim"), data.get("target_bike"),
         data.get("target_run"), data.get("target_date"), data.get("completed_date"),
         data.get("note")),
    )
    conn.commit()
    return cur.lastrowid


def delete_goal(conn, gid: int) -> None:
    conn.execute("DELETE FROM goals WHERE id = ?", (gid,))
    conn.commit()


# --------------------------------------------------------------------------
# Vücut ölçümleri
# --------------------------------------------------------------------------
MEASUREMENT_FIELDS = ("date", "weight_kg", "waist_cm", "neck_cm", "shoulder_cm",
                      "hip_cm", "height_cm", "body_fat_pct", "notes")


def list_measurements(conn, order: str = "DESC", limit: int | None = None) -> list[dict]:
    sql = f"SELECT * FROM measurements ORDER BY date {order}"
    if limit:
        sql += f" LIMIT {int(limit)}"
    return list(conn.execute(sql))


def get_measurement(conn, mid: int) -> Optional[dict]:
    return conn.execute("SELECT * FROM measurements WHERE id = ?", (mid,)).fetchone()


def get_measurement_by_date(conn, day: str) -> Optional[dict]:
    return conn.execute("SELECT * FROM measurements WHERE date = ?", (day,)).fetchone()


def add_measurement(conn, data: dict) -> int:
    """Aynı tarihe ikinci ölçüm girilirse mevcut kayıt güncellenir."""
    existing = get_measurement_by_date(conn, data.get("date"))
    if existing:
        update_measurement(conn, existing["id"], data)
        return existing["id"]
    ts = now_iso()
    cols = ", ".join(MEASUREMENT_FIELDS) + ", created_at, updated_at"
    marks = ", ".join("?" * (len(MEASUREMENT_FIELDS) + 2))
    cur = conn.execute(
        f"INSERT INTO measurements({cols}) VALUES ({marks})",
        [data.get(k) for k in MEASUREMENT_FIELDS] + [ts, ts],
    )
    conn.commit()
    return cur.lastrowid


def update_measurement(conn, mid: int, data: dict) -> None:
    payload = {k: data[k] for k in MEASUREMENT_FIELDS if k in data}
    if not payload:
        return
    sets = ", ".join(f"{k} = ?" for k in payload) + ", updated_at = ?"
    conn.execute(f"UPDATE measurements SET {sets} WHERE id = ?",
                 list(payload.values()) + [now_iso(), mid])
    conn.commit()


def delete_measurement(conn, mid: int) -> None:
    conn.execute("DELETE FROM measurements WHERE id = ?", (mid,))
    conn.commit()


def count_measurements(conn) -> int:
    return conn.execute("SELECT COUNT(*) c FROM measurements").fetchone()["c"]


# --------------------------------------------------------------------------
# Bakım
# --------------------------------------------------------------------------
def wipe_workouts(conn) -> int:
    n = conn.execute("SELECT COUNT(*) c FROM workouts").fetchone()["c"]
    conn.execute("DELETE FROM workouts")
    conn.commit()
    return n


def db_stats(conn) -> dict:
    row = conn.execute(
        "SELECT COUNT(*) n, MIN(date) first, MAX(date) last FROM workouts"
    ).fetchone()
    size = Path(DB_PATH).stat().st_size if Path(DB_PATH).exists() else 0
    return {"workouts": row["n"], "first": row["first"], "last": row["last"],
            "size_kb": round(size / 1024, 1)}
