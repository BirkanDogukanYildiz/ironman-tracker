# -*- coding: utf-8 -*-
"""Bağımlılıksız SVG grafik üreteci.

Renkler dışındaki her şey CSS sınıflarıyla boyanır (.c-grid / .c-axis / .c-val),
böylece grafikler açık ve koyu temada otomatik doğru görünür.
"""
from __future__ import annotations

import math
from html import escape
from typing import Sequence

from markupsafe import Markup


# --------------------------------------------------------------------------
def tr_num(value: float, decimals: int = 1) -> str:
    """Türkçe sayı biçimi: 1.234,5"""
    if value is None:
        return "—"
    text = f"{value:,.{decimals}f}"
    return text.replace(",", " ").replace(".", ",").replace(" ", ".")


def _nice_max(value: float) -> float:
    if value <= 0:
        return 1.0
    exp = math.floor(math.log10(value))
    base = 10 ** exp
    for mult in (1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10):
        if value <= base * mult:
            return base * mult
    return base * 10


def _svg(width: int, height: int, body: str, extra_class: str = "") -> Markup:
    return Markup(
        f'<svg class="chart {extra_class}" viewBox="0 0 {width} {height}" '
        f'role="img" preserveAspectRatio="xMidYMid meet">{body}</svg>'
    )


def _y_axis(x0: int, y0: int, y1: int, x1: int, vmax: float,
            ticks: int = 4, fmt=lambda v: tr_num(v, 0)) -> str:
    out = []
    for i in range(ticks + 1):
        value = vmax * i / ticks
        y = y1 - (y1 - y0) * i / ticks
        out.append(f'<line class="c-grid" x1="{x0}" y1="{y:.1f}" x2="{x1}" y2="{y:.1f}"/>')
        out.append(f'<text class="c-axis" x="{x0 - 8}" y="{y + 4:.1f}" '
                   f'text-anchor="end">{escape(fmt(value))}</text>')
    return "".join(out)


# --------------------------------------------------------------------------
def bar_chart(labels: Sequence[str], values: Sequence[float], color: str,
              *, unit: str = "", decimals: int = 1, height: int = 250,
              width: int = 700, show_values: bool = False,
              neg_color: str | None = None) -> Markup:
    """Dikey sütun grafiği. Negatif değerler sıfır çizgisinin altına çizilir."""
    if not labels:
        return empty_state()
    L, R, T, B = 56, 14, 20, 34
    x0, x1, y0, y1 = L, width - R, T, height - B
    vals = [float(v or 0) for v in values]
    hi = _nice_max(max(vals + [0])) if max(vals + [0]) > 0 else 0.0
    lo = -_nice_max(-min(vals + [0])) if min(vals + [0]) < 0 else 0.0
    span = (hi - lo) or 1.0
    n = len(labels)
    slot = (x1 - x0) / n
    bw = min(46, slot * 0.62)

    def ypx(v: float) -> float:
        return y1 - (y1 - y0) * ((v - lo) / span)

    fmt = lambda v: tr_num(v, decimals if span < 10 else 0)
    body = []
    ticks = 4
    for i in range(ticks + 1):
        value = lo + span * i / ticks
        y = ypx(value)
        body.append(f'<line class="c-grid" x1="{x0}" y1="{y:.1f}" x2="{x1}" y2="{y:.1f}"/>')
        body.append(f'<text class="c-axis" x="{x0 - 8}" y="{y + 4:.1f}" '
                    f'text-anchor="end">{escape(fmt(value))}</text>')

    zero_y = ypx(0)
    for i, (lab, val) in enumerate(zip(labels, vals)):
        cx = x0 + slot * (i + 0.5)
        x = cx - bw / 2
        y_val = ypx(val)
        top, h = min(zero_y, y_val), abs(zero_y - y_val)
        fillc = neg_color if (val < 0 and neg_color) else color
        body.append(
            f'<rect class="c-bar" x="{x:.1f}" y="{top:.1f}" width="{bw:.1f}" '
            f'height="{max(h, 1.0):.1f}" rx="3" fill="{fillc}">'
            f'<title>{escape(str(lab))}: {tr_num(val, decimals)} {escape(unit)}</title></rect>'
        )
        if show_values and val != 0:
            if val > 0:
                ty, cls = top - 6, "c-val"
            else:
                # negatif sütunun etiketini sıfır çizgisinin hemen altına, çubuğun
                # içine koy — aksi halde x ekseni etiketleriyle çakışır
                ty, cls = zero_y + 15, ("c-val c-val--on" if h > 22 else "c-val")
                if h <= 22:
                    ty = zero_y + h + 14
            body.append(f'<text class="{cls}" x="{cx:.1f}" y="{ty:.1f}" '
                        f'text-anchor="middle">{tr_num(val, decimals)}</text>')
        if n <= 20 or i % max(1, n // 12) == 0:
            body.append(f'<text class="c-axis" x="{cx:.1f}" y="{y1 + 18:.1f}" '
                        f'text-anchor="middle">{escape(str(lab))}</text>')
    body.append(f'<line class="c-base" x1="{x0}" y1="{zero_y:.1f}" x2="{x1}" y2="{zero_y:.1f}"/>')
    return _svg(width, height, "".join(body))


def line_chart(labels: Sequence[str], values: Sequence[float], color: str,
               *, unit: str = "", decimals: int = 1, height: int = 250,
               width: int = 700, target: float | None = None) -> Markup:
    if not labels:
        return empty_state()
    L, R, T, B = 52, 14, 18, 34
    x0, x1, y0, y1 = L, width - R, T, height - B
    vmax = _nice_max(max(list(values) + ([target] if target else []) + [0]))
    n = len(labels)
    step = (x1 - x0) / max(1, n - 1) if n > 1 else 0
    pts = []
    for i, val in enumerate(values):
        val = float(val or 0)
        x = x0 + step * i if n > 1 else (x0 + x1) / 2
        y = y1 - (y1 - y0) * (val / vmax if vmax else 0)
        pts.append((x, y, val))

    body = [_y_axis(x0, y0, y1, x1, vmax, fmt=lambda v: tr_num(v, decimals if vmax < 10 else 0))]
    if target:
        ty = y1 - (y1 - y0) * (target / vmax)
        body.append(f'<line class="c-target" x1="{x0}" y1="{ty:.1f}" x2="{x1}" y2="{ty:.1f}"/>')
        body.append(f'<text class="c-axis" x="{x1}" y="{ty - 6:.1f}" text-anchor="end">'
                    f'hedef {tr_num(target, 0)}</text>')

    area = f'M {pts[0][0]:.1f} {y1} ' + " ".join(f'L {x:.1f} {y:.1f}' for x, y, _ in pts) + \
           f' L {pts[-1][0]:.1f} {y1} Z'
    body.append(f'<path d="{area}" fill="{color}" opacity="0.10"/>')
    path = "M " + " L ".join(f'{x:.1f} {y:.1f}' for x, y, _ in pts)
    body.append(f'<path d="{path}" fill="none" stroke="{color}" stroke-width="2.5" '
                f'stroke-linejoin="round" stroke-linecap="round"/>')
    for i, (x, y, val) in enumerate(pts):
        body.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.6" fill="{color}" '
                    f'class="c-dot"><title>{escape(str(labels[i]))}: '
                    f'{tr_num(val, decimals)} {escape(unit)}</title></circle>')
        if n <= 20 or i % max(1, n // 12) == 0:
            body.append(f'<text class="c-axis" x="{x:.1f}" y="{y1 + 18:.1f}" '
                        f'text-anchor="middle">{escape(str(labels[i]))}</text>')
    body.append(f'<line class="c-base" x1="{x0}" y1="{y1}" x2="{x1}" y2="{y1}"/>')
    return _svg(width, height, "".join(body))


def stacked_chart(labels: Sequence[str], series: Sequence[dict],
                  *, unit: str = "km", decimals: int = 1,
                  height: int = 250, width: int = 700) -> Markup:
    """series: [{'name':..., 'color':..., 'values':[...]}, ...]"""
    if not labels or not series:
        return empty_state()
    L, R, T, B = 52, 14, 18, 34
    x0, x1, y0, y1 = L, width - R, T, height - B
    totals = [sum(float(s["values"][i] or 0) for s in series) for i in range(len(labels))]
    vmax = _nice_max(max(totals + [0]))
    n = len(labels)
    slot = (x1 - x0) / n
    bw = min(46, slot * 0.62)

    body = [_y_axis(x0, y0, y1, x1, vmax, fmt=lambda v: tr_num(v, 0))]
    for i, lab in enumerate(labels):
        cx = x0 + slot * (i + 0.5)
        x = cx - bw / 2
        acc = 0.0
        for s in series:
            val = float(s["values"][i] or 0)
            if val <= 0:
                continue
            h = (y1 - y0) * (val / vmax)
            y = y1 - (y1 - y0) * ((acc + val) / vmax)
            body.append(
                f'<rect x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{h:.1f}" '
                f'fill="{s["color"]}" class="c-bar">'
                f'<title>{escape(str(lab))} · {escape(s["name"])}: '
                f'{tr_num(val, decimals)} {escape(unit)}</title></rect>'
            )
            acc += val
        if n <= 20 or i % max(1, n // 12) == 0:
            body.append(f'<text class="c-axis" x="{cx:.1f}" y="{y1 + 18:.1f}" '
                        f'text-anchor="middle">{escape(str(lab))}</text>')
    body.append(f'<line class="c-base" x1="{x0}" y1="{y1}" x2="{x1}" y2="{y1}"/>')
    return _svg(width, height, "".join(body))


def donut_chart(items: Sequence[dict], *, unit: str = "saat", decimals: int = 1,
                size: int = 240) -> Markup:
    """items: [{'label':..., 'value':..., 'color':...}]"""
    total = sum(float(i["value"] or 0) for i in items)
    if total <= 0:
        return empty_state(height=size)
    cx = cy = size / 2
    r_out, r_in = size / 2 - 8, size / 2 - 46
    body, angle = [], -math.pi / 2

    for it in items:
        val = float(it["value"] or 0)
        if val <= 0:
            continue
        sweep = 2 * math.pi * (val / total)
        end = angle + sweep
        large = 1 if sweep > math.pi else 0
        x1o, y1o = cx + r_out * math.cos(angle), cy + r_out * math.sin(angle)
        x2o, y2o = cx + r_out * math.cos(end), cy + r_out * math.sin(end)
        x1i, y1i = cx + r_in * math.cos(end), cy + r_in * math.sin(end)
        x2i, y2i = cx + r_in * math.cos(angle), cy + r_in * math.sin(angle)
        d = (f"M {x1o:.2f} {y1o:.2f} A {r_out} {r_out} 0 {large} 1 {x2o:.2f} {y2o:.2f} "
             f"L {x1i:.2f} {y1i:.2f} A {r_in} {r_in} 0 {large} 0 {x2i:.2f} {y2i:.2f} Z")
        pct = val / total * 100
        body.append(f'<path d="{d}" fill="{it["color"]}" class="c-slice">'
                    f'<title>{escape(it["label"])}: {tr_num(val, decimals)} {escape(unit)} '
                    f'(%{tr_num(pct, 0)})</title></path>')
        mid = angle + sweep / 2
        if pct >= 7:
            lx, ly = cx + (r_out + r_in) / 2 * math.cos(mid), cy + (r_out + r_in) / 2 * math.sin(mid)
            body.append(f'<text class="c-slice-val" x="{lx:.1f}" y="{ly + 4:.1f}" '
                        f'text-anchor="middle">%{tr_num(pct, 0)}</text>')
        angle = end

    body.append(f'<text class="c-donut-total" x="{cx}" y="{cy - 2}" text-anchor="middle">'
                f'{tr_num(total, decimals)}</text>')
    body.append(f'<text class="c-donut-unit" x="{cx}" y="{cy + 16}" text-anchor="middle">'
                f'{escape(unit)}</text>')
    return _svg(size, size, "".join(body), "chart--donut")


def hbar_chart(items: Sequence[dict], *, height_per: int = 62, width: int = 700,
               pct: bool = True) -> Markup:
    """items: [{'label':..., 'value': 0..1, 'color':..., 'note':...}]"""
    if not items:
        return empty_state()
    L, R, T = 140, 66, 14
    height = T * 2 + height_per * len(items)
    body = []
    for i, it in enumerate(items):
        y = T + height_per * i + height_per / 2
        val = max(0.0, min(1.0, float(it["value"] or 0)))
        track_w = width - L - R
        body.append(f'<text class="c-axis c-axis--lg" x="{L - 14}" y="{y + 5}" '
                    f'text-anchor="end">{escape(it["label"])}</text>')
        body.append(f'<rect class="c-track" x="{L}" y="{y - 11}" width="{track_w}" '
                    f'height="22" rx="11"/>')
        body.append(f'<rect x="{L}" y="{y - 11}" width="{track_w * val:.1f}" height="22" '
                    f'rx="11" fill="{it["color"]}"><title>{escape(it["label"])}: '
                    f'%{tr_num(val * 100, 1)}</title></rect>')
        label = f"%{tr_num(val * 100, 1)}" if pct else it.get("note", "")
        body.append(f'<text class="c-val" x="{L + track_w + 12}" y="{y + 5}">{escape(label)}</text>')
        if it.get("note"):
            body.append(f'<text class="c-axis" x="{L}" y="{y + 26}">{escape(it["note"])}</text>')
    return _svg(width, height, "".join(body))


def sparkline(values: Sequence[float], color: str, *, width: int = 120,
              height: int = 32) -> Markup:
    vals = [float(v or 0) for v in values]
    if not vals or max(vals) <= 0:
        return Markup(f'<svg class="spark" viewBox="0 0 {width} {height}"></svg>')
    vmax = max(vals)
    step = width / max(1, len(vals) - 1)
    pts = [(i * step, height - 3 - (height - 8) * (v / vmax)) for i, v in enumerate(vals)]
    path = "M " + " L ".join(f"{x:.1f} {y:.1f}" for x, y in pts)
    area = f"M 0 {height} " + " ".join(f"L {x:.1f} {y:.1f}" for x, y in pts) + f" L {width} {height} Z"
    return Markup(
        f'<svg class="spark" viewBox="0 0 {width} {height}" preserveAspectRatio="none">'
        f'<path d="{area}" fill="{color}" opacity="0.14"/>'
        f'<path d="{path}" fill="none" stroke="{color}" stroke-width="1.8" '
        f'stroke-linejoin="round" stroke-linecap="round"/></svg>'
    )


def empty_state(height: int = 250, width: int = 700,
                text: str = "Henüz veri yok") -> Markup:
    return _svg(width, height,
                f'<text class="c-empty" x="{width/2}" y="{height/2}" '
                f'text-anchor="middle">{escape(text)}</text>')
