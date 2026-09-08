# -*- coding: utf-8 -*-
"""Tüm toplamalar: branş istatistikleri, haftalık özet, brick eşleştirme, hedef durumu."""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Iterable, Optional

from .config import SPORT_META, SPORT_ORDER, MAX_LEVEL
from .levels import LevelState, overall as overall_levels


# --------------------------------------------------------------------------
# Yardımcılar
# --------------------------------------------------------------------------
def parse_date(value) -> Optional[date]:
    if value is None or value == "":
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    text = str(value).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def monday_of(d: date) -> date:
    return d - timedelta(days=d.weekday())


def fmt_pace(minutes: Optional[float]) -> str:
    """Ondalık dakikayı  m:ss  metnine çevirir."""
    if not minutes or minutes <= 0 or minutes != minutes:
        return "—"
    total = int(round(minutes * 60))
    return f"{total // 60}:{total % 60:02d}"


def fmt_hours(minutes: float) -> str:
    if not minutes:
        return "0 dk"
    h, m = divmod(int(round(minutes)), 60)
    return f"{h}s {m:02d}dk" if h else f"{m} dk"


def pace_minutes(sport: str, distance_km: float, duration_min: float) -> Optional[float]:
    """Yüzme → dk/100 m · Koşu → dk/km · Bisiklet → yok (hız kullanılır)."""
    meta = SPORT_META[sport]
    if not meta["pace_divisor"] or not distance_km or not duration_min:
        return None
    return duration_min / (distance_km / meta["pace_divisor"])


def speed_kmh(distance_km: float, duration_min: float) -> Optional[float]:
    if not distance_km or not duration_min:
        return None
    return distance_km / (duration_min / 60.0)


# --------------------------------------------------------------------------
# Branş istatistikleri
# --------------------------------------------------------------------------
@dataclass
class SportStats:
    sport: str
    label: str = ""
    icon: str = ""
    color: str = ""
    count: int = 0
    total_km: float = 0.0
    total_min: float = 0.0
    max_km: float = 0.0
    max_date: Optional[date] = None
    avg_km: float = 0.0
    avg_pace: Optional[float] = None
    avg_speed: Optional[float] = None
    km_7d: float = 0.0
    km_30d: float = 0.0
    km_90d: float = 0.0
    count_30d: int = 0
    min_30d: float = 0.0
    race_km: float = 0.0
    race_ratio: float = 0.0
    level: Optional[LevelState] = None

    @property
    def total_hours(self) -> float:
        return self.total_min / 60.0

    @property
    def pace_text(self) -> str:
        if SPORT_META[self.sport]["pace_unit"] is None:
            return "—"
        return f"{fmt_pace(self.avg_pace)} {SPORT_META[self.sport]['pace_unit']}"

    @property
    def max_display(self) -> str:
        if self.sport == "swim":
            return f"{self.max_km * 1000:,.0f} m".replace(",", ".")
        return f"{self.max_km:,.1f} km".replace(",", " ").replace(".", ",")


def _round_km(sport: str, km: float) -> float:
    return round(km, SPORT_META[sport]["decimals"])


def compute_sport_stats(workouts: list[dict], thresholds: dict,
                        today: Optional[date] = None) -> dict[str, SportStats]:
    today = today or date.today()
    d7, d30, d90 = today - timedelta(days=6), today - timedelta(days=29), today - timedelta(days=89)

    out: dict[str, SportStats] = {}
    for sport in SPORT_ORDER:
        meta = SPORT_META[sport]
        st = SportStats(sport=sport, label=meta["label"], icon=meta["icon"],
                        color=meta["color"], race_km=meta["race_km"])
        rows = [w for w in workouts if w["sport"] == sport]
        for w in rows:
            wd = parse_date(w["date"])
            km = float(w["distance_km"] or 0)
            mn = float(w["duration_min"] or 0)
            st.count += 1
            st.total_km += km
            st.total_min += mn
            if km > st.max_km:
                st.max_km, st.max_date = km, wd
            if wd:
                if wd >= d7 and wd <= today:
                    st.km_7d += km
                if d30 <= wd <= today:
                    st.km_30d += km
                    st.count_30d += 1
                    st.min_30d += mn
                if d90 <= wd <= today:
                    st.km_90d += km

        st.avg_km = st.total_km / st.count if st.count else 0.0
        st.avg_pace = pace_minutes(sport, st.total_km, st.total_min)
        st.avg_speed = speed_kmh(st.total_km, st.total_min)
        st.race_ratio = min(1.0, st.max_km / st.race_km) if st.race_km else 0.0
        st.level = LevelState(
            sport=sport,
            max_km=st.max_km,
            volume_30d_km=st.km_30d,
            thresholds_distance=thresholds[sport]["distance_km"],
            thresholds_volume=thresholds[sport]["volume_30d_km"],
        )
        out[sport] = st
    return out


# --------------------------------------------------------------------------
# Genel özet
# --------------------------------------------------------------------------
@dataclass
class Snapshot:
    today: date
    sports: dict[str, SportStats]
    overall: dict
    total_workouts: int = 0
    total_km: float = 0.0
    total_min: float = 0.0
    week_count: int = 0
    week_min: float = 0.0
    week_km: float = 0.0
    last30_min: float = 0.0
    prev30_min: float = 0.0
    last30_km: float = 0.0
    delta30: Optional[float] = None
    brick_count: int = 0
    streak_weeks: int = 0
    last_workout: Optional[dict] = None
    race_ratio: float = 0.0

    @property
    def total_hours(self) -> float:
        return self.total_min / 60.0


def build_snapshot(workouts: list[dict], thresholds: dict,
                   today: Optional[date] = None) -> Snapshot:
    today = today or date.today()
    sports = compute_sport_stats(workouts, thresholds, today)
    ov = overall_levels({s: st.level for s, st in sports.items()})

    week_start = monday_of(today)
    d30, p30a, p30b = (today - timedelta(days=29),
                       today - timedelta(days=59), today - timedelta(days=30))

    snap = Snapshot(today=today, sports=sports, overall=ov)
    snap.total_workouts = len(workouts)
    for w in workouts:
        wd = parse_date(w["date"])
        km, mn = float(w["distance_km"] or 0), float(w["duration_min"] or 0)
        snap.total_km += km
        snap.total_min += mn
        if w["is_brick"]:
            snap.brick_count += 1
        if wd:
            if wd >= week_start and wd <= today:
                snap.week_count += 1
                snap.week_min += mn
                snap.week_km += km
            if d30 <= wd <= today:
                snap.last30_min += mn
                snap.last30_km += km
            if p30a <= wd <= p30b:
                snap.prev30_min += mn

    snap.delta30 = (snap.last30_min / snap.prev30_min - 1) if snap.prev30_min else None
    snap.race_ratio = sum(s.race_ratio for s in sports.values()) / len(sports) if sports else 0.0
    dated = [w for w in workouts if parse_date(w["date"])]
    snap.last_workout = max(dated, key=lambda w: (parse_date(w["date"]), w["id"])) if dated else None

    # kesintisiz antrenman yapılan hafta serisi
    weeks = {monday_of(parse_date(w["date"])) for w in dated}
    cur, streak = week_start, 0
    while cur in weeks:
        streak += 1
        cur -= timedelta(days=7)
    snap.streak_weeks = streak
    return snap


# --------------------------------------------------------------------------
# Haftalık özet
# --------------------------------------------------------------------------
@dataclass
class WeekRow:
    start: date
    index: int
    swim_km: float = 0.0
    bike_km: float = 0.0
    run_km: float = 0.0
    total_km: float = 0.0
    total_min: float = 0.0
    count: int = 0
    long_run: float = 0.0
    long_bike: float = 0.0
    long_swim: float = 0.0
    bricks: int = 0
    delta_min: Optional[float] = None
    delta_km: Optional[float] = None

    @property
    def end(self) -> date:
        return self.start + timedelta(days=6)

    @property
    def hours(self) -> float:
        return self.total_min / 60.0

    @property
    def label(self) -> str:
        return f"{self.start.day:02d}.{self.start.month:02d}"


def weekly_rows(workouts: list[dict], weeks: int = 26,
                today: Optional[date] = None,
                start_date: Optional[date] = None) -> list[WeekRow]:
    """Son `weeks` haftayı eskiden yeniye döner (veri olmayan haftalar da dahil)."""
    today = today or date.today()
    this_monday = monday_of(today)
    first = this_monday - timedelta(days=7 * (weeks - 1))
    if start_date:
        first = max(first, monday_of(start_date))

    buckets: dict[date, WeekRow] = {}
    cur, i = first, 1
    while cur <= this_monday:
        buckets[cur] = WeekRow(start=cur, index=i)
        cur += timedelta(days=7)
        i += 1

    brick_days: dict[date, set[date]] = defaultdict(set)
    for w in workouts:
        wd = parse_date(w["date"])
        if not wd:
            continue
        m = monday_of(wd)
        row = buckets.get(m)
        if row is None:
            continue
        km, mn = float(w["distance_km"] or 0), float(w["duration_min"] or 0)
        row.count += 1
        row.total_min += mn
        row.total_km += km
        if w["sport"] == "swim":
            row.swim_km += km
            row.long_swim = max(row.long_swim, km)
        elif w["sport"] == "bike":
            row.bike_km += km
            row.long_bike = max(row.long_bike, km)
        else:
            row.run_km += km
            row.long_run = max(row.long_run, km)
        if w["is_brick"]:
            brick_days[m].add(wd)

    for m, days in brick_days.items():
        if m in buckets:
            buckets[m].bricks = len(days)

    rows = [buckets[k] for k in sorted(buckets)]
    for prev, cur_row in zip(rows, rows[1:]):
        if prev.total_min:
            cur_row.delta_min = cur_row.total_min / prev.total_min - 1
        if prev.total_km:
            cur_row.delta_km = cur_row.total_km / prev.total_km - 1
    return rows


# --------------------------------------------------------------------------
# Brick oturumları
# --------------------------------------------------------------------------
@dataclass
class BrickSession:
    day: date
    bike_km: float = 0.0
    bike_min: float = 0.0
    run_km: float = 0.0
    run_min: float = 0.0
    swim_km: float = 0.0
    swim_min: float = 0.0
    rpe: Optional[float] = None
    notes: str = ""
    brick_pace: Optional[float] = None
    normal_pace: Optional[float] = None
    pace_delta_sec: Optional[float] = None
    bike_speed: Optional[float] = None

    @property
    def total_min(self) -> float:
        return self.bike_min + self.run_min + self.swim_min


def brick_sessions(workouts: list[dict], lookback_days: int = 60) -> list[BrickSession]:
    """Brick işaretli kayıtları güne göre gruplar; bisiklet sonrası koşu pace'ini
    aynı dönemin normal koşu pace'i ile karşılaştırır."""
    by_day: dict[date, list[dict]] = defaultdict(list)
    for w in workouts:
        wd = parse_date(w["date"])
        if wd and w["is_brick"]:
            by_day[wd].append(w)

    normal_runs = [(parse_date(w["date"]), float(w["distance_km"] or 0),
                    float(w["duration_min"] or 0))
                   for w in workouts
                   if w["sport"] == "run" and not w["is_brick"] and parse_date(w["date"])]

    out: list[BrickSession] = []
    for day in sorted(by_day, reverse=True):
        rows = by_day[day]
        s = BrickSession(day=day)
        rpes, notes = [], []
        for w in rows:
            km, mn = float(w["distance_km"] or 0), float(w["duration_min"] or 0)
            if w["sport"] == "bike":
                s.bike_km += km; s.bike_min += mn
            elif w["sport"] == "run":
                s.run_km += km; s.run_min += mn
            else:
                s.swim_km += km; s.swim_min += mn
            if w["rpe"]:
                rpes.append(float(w["rpe"]))
            if w["notes"]:
                notes.append(w["notes"])
        s.rpe = sum(rpes) / len(rpes) if rpes else None
        s.notes = " · ".join(notes)
        s.brick_pace = pace_minutes("run", s.run_km, s.run_min)
        s.bike_speed = speed_kmh(s.bike_km, s.bike_min)

        window_start = day - timedelta(days=lookback_days - 1)
        km_sum = sum(km for d, km, _ in normal_runs if window_start <= d <= day)
        mn_sum = sum(mn for d, _, mn in normal_runs if window_start <= d <= day)
        s.normal_pace = pace_minutes("run", km_sum, mn_sum)
        if s.brick_pace and s.normal_pace:
            s.pace_delta_sec = (s.brick_pace - s.normal_pace) * 60
        out.append(s)
    return out


# --------------------------------------------------------------------------
# Hedefler
# --------------------------------------------------------------------------
@dataclass
class GoalStatus:
    goal: dict
    current: float = 0.0
    target: float = 0.0
    progress: float = 0.0
    unit: str = "km"
    state: str = "todo"          # todo | active | ready | done
    state_label: str = "Başlanmadı"
    completed_on: Optional[date] = None
    overdue: bool = False
    legs: list[dict] = field(default_factory=list)

    @property
    def is_done(self) -> bool:
        return self.state == "done"


def _first_date_reaching(workouts: list[dict], sport: str, target_km: float) -> Optional[date]:
    days = [parse_date(w["date"]) for w in workouts
            if w["sport"] == sport and float(w["distance_km"] or 0) >= target_km
            and parse_date(w["date"])]
    return min(days) if days else None


def evaluate_goals(goals: list[dict], workouts: list[dict],
                   sports: dict[str, SportStats],
                   today: Optional[date] = None) -> list[GoalStatus]:
    today = today or date.today()
    out: list[GoalStatus] = []
    for g in goals:
        gs = GoalStatus(goal=g)
        if g["goal_type"] == "distance":
            sport = g["sport"]
            gs.target = float(g["target_value"] or 0)
            gs.current = sports[sport].max_km if sport in sports else 0.0
            gs.progress = min(1.0, gs.current / gs.target) if gs.target else 0.0
            if gs.progress >= 1:
                gs.state, gs.state_label = "done", "Tamamlandı"
                gs.completed_on = (parse_date(g["completed_date"])
                                   or _first_date_reaching(workouts, sport, gs.target))
            elif gs.progress > 0:
                gs.state, gs.state_label = "active", "Devam ediyor"
        else:
            legs = [("swim", g["target_swim"]), ("bike", g["target_bike"]), ("run", g["target_run"])]
            met = 0
            for sport, target in legs:
                target = float(target or 0)
                cur = sports[sport].max_km if sport in sports else 0.0
                ok = target > 0 and cur >= target
                met += 1 if ok else 0
                gs.legs.append({
                    "sport": sport, "target": target, "current": cur, "ok": ok,
                    "pct": min(1.0, cur / target) if target else 0.0,
                    "label": SPORT_META[sport]["label"], "icon": SPORT_META[sport]["icon"],
                    "color": SPORT_META[sport]["color"],
                })
            gs.current, gs.target, gs.unit = met, 3, "ayak"
            gs.progress = sum(l["pct"] for l in gs.legs) / 3
            gs.completed_on = parse_date(g["completed_date"])
            if gs.completed_on:
                gs.state, gs.state_label = "done", "Tamamlandı"
            elif met == 3:
                gs.state, gs.state_label = "ready", "Yarışa hazır"
            elif gs.progress > 0:
                gs.state, gs.state_label = "active", "Devam ediyor"

        td = parse_date(g["target_date"])
        gs.overdue = bool(td and td < today and gs.state != "done")
        out.append(gs)
    return out


def next_goal(statuses: list[GoalStatus]) -> Optional[GoalStatus]:
    for gs in statuses:
        if gs.state != "done":
            return gs
    return None
