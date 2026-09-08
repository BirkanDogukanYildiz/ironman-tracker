# -*- coding: utf-8 -*-
"""Seviye motoru — Excel'deki BRANŞ İLERLEME formüllerinin birebir Python karşılığı.

Her branş için iki bağımsız seviye hesaplanır:

  Mesafe Seviyesi (rozet)  : en uzun TEK antrenmana bakar, asla düşmez.
  Form Seviyesi (süreklilik): son 30 günün TOPLAM hacmine bakar, ara verilince düşer.

  Mevcut seviye  = min(mesafe, form)
  Genel seviye   = üç branşın mevcut seviyelerinin en küçüğü (en zayıf halka)
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .config import LEVEL_NAMES, MAX_LEVEL


def level_from(value: float, thresholds: list[float]) -> int:
    """Karşılanan eşik sayısı − 1.  Eşit eşikler (ör. bisiklet L1=L2=30) güvenli çalışır."""
    met = sum(1 for t in thresholds if value >= t)
    return max(0, met - 1)


def intra_level_progress(value: float, thresholds: list[float], level: int) -> float:
    """Mevcut seviye ile bir sonraki seviye arasındaki ilerleme (0–1)."""
    if level >= MAX_LEVEL:
        return 1.0
    lo, hi = thresholds[level], thresholds[level + 1]
    if hi <= lo:
        return 1.0
    return min(1.0, max(0.0, (value - lo) / (hi - lo)))


def level_name(level: int) -> str:
    return LEVEL_NAMES[max(0, min(MAX_LEVEL, int(level)))]


def next_target(thresholds: list[float], level: int) -> float:
    if level >= MAX_LEVEL:
        return thresholds[MAX_LEVEL]
    return thresholds[level + 1]


@dataclass
class LevelState:
    """Tek bir branşın seviye durumu."""
    sport: str
    max_km: float
    volume_30d_km: float
    thresholds_distance: list[float]
    thresholds_volume: list[float]

    distance_level: int = 0
    volume_level: int = 0
    current_level: int = 0
    current_name: str = ""
    badge_name: str = ""
    next_level: int = 0
    next_level_name: str = ""
    next_target_km: float = 0.0
    remaining_km: float = 0.0
    intra_pct: float = 0.0
    ironman_pct: float = 0.0
    is_max: bool = False

    def __post_init__(self) -> None:
        self.distance_level = level_from(self.max_km, self.thresholds_distance)
        self.volume_level = level_from(self.volume_30d_km, self.thresholds_volume)
        self.current_level = min(self.distance_level, self.volume_level)
        self.current_name = level_name(self.current_level)
        self.badge_name = level_name(self.distance_level)
        self.is_max = self.distance_level >= MAX_LEVEL
        self.next_level = min(MAX_LEVEL, self.distance_level + 1)
        self.next_level_name = level_name(self.next_level)
        self.next_target_km = next_target(self.thresholds_distance, self.distance_level)
        self.remaining_km = max(0.0, self.next_target_km - self.max_km)
        self.intra_pct = intra_level_progress(
            self.max_km, self.thresholds_distance, self.distance_level)
        self.ironman_pct = min(1.0, (self.distance_level + self.intra_pct) / MAX_LEVEL)


def overall(states: dict[str, LevelState]) -> dict:
    """Üç branştan genel durumu türetir."""
    if not states:
        return {"form_level": 0, "form_name": LEVEL_NAMES[0], "badge_level": 0,
                "badge_name": LEVEL_NAMES[0], "ironman_pct": 0.0,
                "weakest": None, "strongest": None}
    form = min(s.current_level for s in states.values())
    badge = min(s.distance_level for s in states.values())
    pct = sum(s.ironman_pct for s in states.values()) / len(states)
    weakest = min(states.values(), key=lambda s: (s.current_level, s.ironman_pct))
    strongest = max(states.values(), key=lambda s: (s.current_level, s.ironman_pct))
    return {
        "form_level": form,
        "form_name": level_name(form),
        "badge_level": badge,
        "badge_name": level_name(badge),
        "ironman_pct": pct,
        "weakest": weakest.sport,
        "strongest": strongest.sport,
    }
