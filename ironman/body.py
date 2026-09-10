# -*- coding: utf-8 -*-
"""Vücut ölçümleri ve türetilen değerler.

Excel'deki «kg-bel-boyun-yağ%» sayfasının hesap mantığı birebir korundu:

    YAĞ ORAN      = US Navy formülü (bel, boyun, boy)
    YAĞSIZ KİTLE  = kilo × (1 − yağ oranı)
    YAĞ KİTLESİ   = kilo × yağ oranı
    VÜCUT KOMP.   = omuz ÷ bel      (V-oranı)
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from .stats import parse_date


def navy_body_fat(waist_cm: Optional[float], neck_cm: Optional[float],
                  height_cm: Optional[float], sex: str = "male",
                  hip_cm: Optional[float] = None) -> Optional[float]:
    """US Navy vücut yağ oranı (0–1 arası oran döner)."""
    if not waist_cm or not neck_cm or not height_cm:
        return None
    try:
        if sex == "female":
            if not hip_cm:
                return None
            girth = waist_cm + hip_cm - neck_cm
            if girth <= 0:
                return None
            denom = (1.29579 - 0.35004 * math.log10(girth)
                     + 0.22100 * math.log10(height_cm))
        else:
            girth = waist_cm - neck_cm
            if girth <= 0:
                return None
            denom = (1.0324 - 0.19077 * math.log10(girth)
                     + 0.15456 * math.log10(height_cm))
        if denom <= 0:
            return None
        pct = 495.0 / denom - 450.0
    except (ValueError, ZeroDivisionError):
        return None
    if not (0 < pct < 80):
        return None
    return pct / 100.0


def bmi(weight_kg: Optional[float], height_cm: Optional[float]) -> Optional[float]:
    if not weight_kg or not height_cm:
        return None
    m = height_cm / 100.0
    return weight_kg / (m * m)


def age_on(birth_date, when: Optional[date] = None) -> Optional[int]:
    b = parse_date(birth_date)
    if not b:
        return None
    when = when or date.today()
    return when.year - b.year - ((when.month, when.day) < (b.month, b.day))


@dataclass
class Measurement:
    """Bir ölçüm satırı + türetilen tüm değerler."""
    row: dict
    height_cm: Optional[float] = None
    sex: str = "male"

    day: Optional[date] = None
    weight_kg: Optional[float] = None
    waist_cm: Optional[float] = None
    neck_cm: Optional[float] = None
    shoulder_cm: Optional[float] = None
    hip_cm: Optional[float] = None
    body_fat: Optional[float] = None          # 0–1 oran
    lean_mass_kg: Optional[float] = None
    fat_mass_kg: Optional[float] = None
    v_ratio: Optional[float] = None           # omuz ÷ bel
    bmi: Optional[float] = None
    age: Optional[int] = None
    notes: str = ""

    # bir önceki ölçüme göre değişimler
    d_weight: Optional[float] = None
    d_waist: Optional[float] = None
    d_fat: Optional[float] = None
    d_lean: Optional[float] = None

    def __post_init__(self) -> None:
        r = self.row
        self.day = parse_date(r.get("date"))
        self.weight_kg = _f(r.get("weight_kg"))
        self.waist_cm = _f(r.get("waist_cm"))
        self.neck_cm = _f(r.get("neck_cm"))
        self.shoulder_cm = _f(r.get("shoulder_cm"))
        self.hip_cm = _f(r.get("hip_cm"))
        self.notes = r.get("notes") or ""

        h = _f(r.get("height_cm")) or self.height_cm
        self.height_cm = h
        manual = _f(r.get("body_fat_pct"))
        self.body_fat = (manual / 100.0 if manual
                         else navy_body_fat(self.waist_cm, self.neck_cm, h,
                                            self.sex, self.hip_cm))
        if self.weight_kg and self.body_fat is not None:
            self.fat_mass_kg = self.weight_kg * self.body_fat
            self.lean_mass_kg = self.weight_kg - self.fat_mass_kg
        if self.shoulder_cm and self.waist_cm:
            self.v_ratio = self.shoulder_cm / self.waist_cm
        self.bmi = bmi(self.weight_kg, h)


def _f(value) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _delta(new: Optional[float], old: Optional[float]) -> Optional[float]:
    if new is None or old is None:
        return None
    return new - old


def build_series(rows: list[dict], *, height_cm: Optional[float] = None,
                 sex: str = "male", birth_date=None) -> list[Measurement]:
    """Kayıtları eskiden yeniye işler, ardışık farkları hesaplar."""
    ordered = sorted([r for r in rows if parse_date(r.get("date"))],
                     key=lambda r: parse_date(r["date"]))
    out: list[Measurement] = []
    for r in ordered:
        m = Measurement(row=r, height_cm=height_cm, sex=sex)
        m.age = age_on(birth_date, m.day) if birth_date else None
        if out:
            prev = out[-1]
            m.d_weight = _delta(m.weight_kg, prev.weight_kg)
            m.d_waist = _delta(m.waist_cm, prev.waist_cm)
            m.d_fat = _delta(m.body_fat, prev.body_fat)
            m.d_lean = _delta(m.lean_mass_kg, prev.lean_mass_kg)
        out.append(m)
    return out


@dataclass
class BodySummary:
    latest: Optional[Measurement] = None
    first: Optional[Measurement] = None
    count: int = 0
    span_days: int = 0
    total_weight: Optional[float] = None
    total_waist: Optional[float] = None
    total_fat: Optional[float] = None
    total_lean: Optional[float] = None
    weight_30d: Optional[float] = None
    waist_30d: Optional[float] = None

    @property
    def has_data(self) -> bool:
        return self.latest is not None


def summarize(series: list[Measurement], today: Optional[date] = None) -> BodySummary:
    s = BodySummary()
    if not series:
        return s
    today = today or date.today()
    s.count = len(series)
    s.first, s.latest = series[0], series[-1]
    if s.first.day and s.latest.day:
        s.span_days = (s.latest.day - s.first.day).days
    s.total_weight = _delta(s.latest.weight_kg, s.first.weight_kg)
    s.total_waist = _delta(s.latest.waist_cm, s.first.waist_cm)
    s.total_fat = _delta(s.latest.body_fat, s.first.body_fat)
    s.total_lean = _delta(s.latest.lean_mass_kg, s.first.lean_mass_kg)

    older = [m for m in series if m.day and (today - m.day).days >= 30]
    ref = older[-1] if older else s.first
    if ref is not s.latest:
        s.weight_30d = _delta(s.latest.weight_kg, ref.weight_kg)
        s.waist_30d = _delta(s.latest.waist_cm, ref.waist_cm)
    return s
