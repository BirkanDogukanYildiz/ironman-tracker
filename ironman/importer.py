# -*- coding: utf-8 -*-
"""Strava / Garmin Connect (ve genel) CSV içe aktarma.

Kolon adlarını otomatik tanır, birimleri (m/km, sn/dk/hh:mm:ss) tespit eder,
mükerrer kayıtları external_id ile engeller.
"""
from __future__ import annotations

import csv
import hashlib
import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional

from .config import SPORT_ORDER

# --------------------------------------------------------------------------
# Kolon takma adları  (küçük harfe indirgenmiş, aksan duyarsız karşılaştırma)
# --------------------------------------------------------------------------
ALIASES = {
    "date": ["activity date", "date", "tarih", "start time", "başlangıç zamanı",
             "baslangic zamani", "activity_date", "datum", "fecha", "workout date"],
    "sport": ["activity type", "type", "sport", "activity name type", "aktivite türü",
              "aktivite turu", "aktivite tipi", "spor", "activity_type", "sport type"],
    "name": ["activity name", "name", "title", "aktivite adı", "aktivite adi", "başlık"],
    "distance": ["distance", "mesafe", "distance (km)", "distance (m)", "mesafe (km)",
                 "mesafe (m)", "total distance", "toplam mesafe", "distancia"],
    "duration": ["moving time", "elapsed time", "duration", "time", "süre", "sure",
                 "hareket süresi", "hareket suresi", "geçen süre", "gecen sure",
                 "total time", "toplam süre", "toplam sure", "tiempo"],
    "hr": ["average heart rate", "avg hr", "heart rate", "ortalama nabız",
           "ortalama nabiz", "avg heart rate", "ort. nabız", "nabız", "nabiz"],
    "elevation": ["elevation gain", "total ascent", "yükseklik kazancı",
                  "yukseklik kazanci", "toplam tırmanış", "toplam tirmanis",
                  "elevation", "yükselti", "yukselti", "ascent"],
    "calories": ["calories", "kalori", "kcal", "calorias"],
    "id": ["activity id", "aktivite id", "id", "activity_id"],
    "notes": ["activity description", "description", "notes", "not", "notlar", "açıklama"],
}

SPORT_PATTERNS = [
    ("swim", ["swim", "yüzme", "yuzme", "pool", "havuz", "open water", "açık su", "acik su",
              "natación", "schwimmen"]),
    ("bike", ["ride", "bike", "cycl", "bisiklet", "biking", "virtual ride", "e-bike",
              "handcycle", "gravel", "mountain bike", "spinning", "indoor cycling",
              "ciclismo", "radfahren"]),
    ("run", ["run", "koşu", "kosu", "jog", "trail", "treadmill", "koşu bandı",
             "walk", "yürüyüş", "yuruyus", "hike", "carrera", "laufen"]),
]

TR_MAP = str.maketrans("ıİşŞğĞüÜöÖçÇ", "iisSgGuUoOcC")


def _norm(text: str) -> str:
    return (text or "").strip().translate(TR_MAP).lower()


def detect_sport(value: str) -> Optional[str]:
    v = _norm(value)
    if not v:
        return None
    for sport, pats in SPORT_PATTERNS:
        for p in pats:
            if _norm(p) in v:
                return sport
    return None


def parse_number(value) -> Optional[float]:
    """'12,5' · '1.234,5' · '1,234.5' · '45 km' hepsini sayıya çevirir."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"[^\d,.\-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None
    if "," in text and "." in text:
        # son görülen ayraç ondalıktır
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_duration(value, header: str = "") -> Optional[float]:
    """Dakika cinsinden süre döner. hh:mm:ss, saniye veya dakika kabul eder."""
    if value is None or str(value).strip() == "":
        return None
    text = str(value).strip()
    if ":" in text:
        parts = [parse_number(p) or 0 for p in text.split(":")]
        if len(parts) == 3:
            return parts[0] * 60 + parts[1] + parts[2] / 60
        if len(parts) == 2:
            return parts[0] + parts[1] / 60
    num = parse_number(text)
    if num is None:
        return None
    h = _norm(header)
    if any(k in h for k in ("(s)", "sec", "saniye", "(sn)", "seconds")):
        return num / 60
    if any(k in h for k in ("hour", "saat", "(h)")):
        return num * 60
    if any(k in h for k in ("min", "dakika", "(dk)", "(m)")):
        return num
    # Strava varsayılanı saniyedir; 10 saatten uzun "dakika" mantıksızdır
    return num / 60 if num > 600 else num


def parse_date_cell(value) -> Optional[date]:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("T", " ")
    candidates = [text, text.split(" ")[0], " ".join(text.split(" ")[:2])]
    formats = ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d",
               "%b %d, %Y", "%d %b %Y", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M",
               "%d/%m/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%b %d, %Y, %I:%M:%S %p")
    for cand in candidates:
        for fmt in formats:
            try:
                return datetime.strptime(cand.strip(), fmt).date()
            except ValueError:
                continue
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", text)
    if m:
        try:
            return date(int(m[1]), int(m[2]), int(m[3]))
        except ValueError:
            pass
    return None


def normalize_distance(raw: Optional[float], header: str, sport: Optional[str]) -> float:
    """Kilometreye çevirir. Başlıkta birim yoksa büyüklükten tahmin eder."""
    if raw is None:
        return 0.0
    h = _norm(header)
    if "(m)" in h or "metre" in h or "meters" in h or h.endswith(" m"):
        return raw / 1000
    if "(km)" in h or "kilometre" in h:
        return raw
    if "mile" in h or "mil" in h:
        return raw * 1.609344
    # birim belirtilmemiş: Strava activities.csv metre verir
    if sport == "swim":
        return raw / 1000 if raw > 20 else raw
    return raw / 1000 if raw > 400 else raw


# --------------------------------------------------------------------------
@dataclass
class ImportRow:
    ok: bool
    data: dict = field(default_factory=dict)
    reason: str = ""
    raw: dict = field(default_factory=dict)


@dataclass
class ImportPreview:
    headers: list[str]
    mapping: dict[str, str]
    rows: list[ImportRow]
    total: int = 0
    valid: int = 0
    skipped: int = 0

    @property
    def sample(self) -> list[ImportRow]:
        return self.rows[:12]


def _build_mapping(headers: list[str]) -> dict[str, str]:
    normalized = {h: _norm(h) for h in headers}
    mapping: dict[str, str] = {}
    for field_name, aliases in ALIASES.items():
        norm_aliases = [_norm(a) for a in aliases]
        best = None
        for h, nh in normalized.items():
            if nh in norm_aliases:
                best = h
                break
        if best is None:
            for h, nh in normalized.items():
                if any(a and a in nh for a in norm_aliases):
                    best = h
                    break
        if best:
            mapping[field_name] = best
    return mapping


def _sniff(text: str) -> str:
    head = text[:4096]
    try:
        return csv.Sniffer().sniff(head, delimiters=",;\t|").delimiter
    except csv.Error:
        counts = {d: head.count(d) for d in ",;\t|"}
        return max(counts, key=counts.get) or ","


def read_csv(content: bytes | str, *, default_kind: str = "",
             mark_brick: bool = False) -> ImportPreview:
    if isinstance(content, bytes):
        for enc in ("utf-8-sig", "utf-8", "cp1254", "latin-1"):
            try:
                text = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = content.decode("utf-8", errors="replace")
    else:
        text = content

    delim = _sniff(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    headers = [h for h in (reader.fieldnames or []) if h is not None]
    mapping = _build_mapping(headers)

    rows: list[ImportRow] = []
    for raw in reader:
        rows.append(_parse_row(raw, mapping, default_kind, mark_brick))

    preview = ImportPreview(headers=headers, mapping=mapping, rows=rows)
    preview.total = len(rows)
    preview.valid = sum(1 for r in rows if r.ok)
    preview.skipped = preview.total - preview.valid
    return preview


def _parse_row(raw: dict, mapping: dict, default_kind: str, mark_brick: bool) -> ImportRow:
    def cell(key):
        col = mapping.get(key)
        return raw.get(col) if col else None

    day = parse_date_cell(cell("date"))
    if not day:
        return ImportRow(False, reason="Tarih okunamadı", raw=raw)

    sport_raw = cell("sport") or cell("name") or ""
    sport = detect_sport(str(sport_raw))
    if not sport:
        return ImportRow(False, reason=f"Branş tanınmadı: “{str(sport_raw)[:32]}”", raw=raw)

    dist_header = mapping.get("distance", "")
    distance = normalize_distance(parse_number(cell("distance")), dist_header, sport)
    duration = parse_duration(cell("duration"), mapping.get("duration", "")) or 0.0
    if distance <= 0 and duration <= 0:
        return ImportRow(False, reason="Mesafe ve süre boş", raw=raw)

    hr = parse_number(cell("hr"))
    elev = parse_number(cell("elevation"))
    kcal = parse_number(cell("calories"))
    ext = cell("id")
    name = (cell("name") or "").strip()
    note = (cell("notes") or "").strip()

    if ext:
        external = f"csv:{str(ext).strip()}"
    else:
        digest = hashlib.sha1(
            f"{day.isoformat()}|{sport}|{round(distance, 3)}|{round(duration, 1)}".encode()
        ).hexdigest()[:16]
        external = f"csv:{digest}"

    data = {
        "date": day.isoformat(),
        "sport": sport,
        "kind": default_kind or ("Açık Su" if "open water" in _norm(str(sport_raw)) else ""),
        "distance_km": round(distance, 3),
        "duration_min": round(duration, 1),
        "avg_hr": int(hr) if hr else None,
        "rpe": None,
        "elevation_m": int(elev) if elev else None,
        "calories": int(kcal) if kcal else None,
        "is_brick": 1 if mark_brick else 0,
        "is_openwater": 1 if (sport == "swim" and "open water" in _norm(str(sport_raw))) else 0,
        "notes": " · ".join(x for x in (name, note) if x)[:400],
        "source": "csv",
        "external_id": external,
    }
    return ImportRow(True, data=data, raw=raw)
