# domnut.py
# Donut chart (SVG) with transparent background, outside labels (name on top, % below),
# perfectly centered using tspans (same x), no leader lines, thicker ring.
# Produces ONE chart per investor from InvestorDataTest.xlsx ("Investor Data - LongForm" sheet).
# Requires: pip install svgwrite pandas openpyxl  (optional: cairosvg if rsvg-convert not available)

import math
import os
import re
import shutil
from typing import Dict, Iterable, List, Tuple
import pandas as pd
import svgwrite
from pathlib import Path
try:
    import tomllib  # py3.11+
except Exception:
    import tomli as tomllib  # type: ignore

# --- anchor everything to this file's folder ---
HERE = Path(__file__).parent

def _load_cfg(path: str | None = None) -> dict:
    cfg_path = HERE / "Configs" / "config.toml"
    if cfg_path.exists():
        with open(cfg_path, "rb") as f:
            return tomllib.load(f)
    return {}

_CFG = _load_cfg()

# Read defaults from config, but keep sane fallbacks
EXCEL_PATH = str(HERE / _CFG.get("paths", {}).get("excel", "InvestorDataTest.xlsx"))
CHART_DIR  = str(HERE / _CFG.get("paths", {}).get("donut_dir", "charts"))

# Use the same font names your PDF code registers (names only; svgwrite will embed as text)
FONT_BOLD_NAME = _CFG.get("fonts", {}).get("bold_name",  "HKGrotesk-Bold")
FONT_MED_NAME  = _CFG.get("fonts", {}).get("medium_name","HKGrotesk-Medium")

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def _config_name_maps(cfg: dict):
    """
    Build maps so we can turn a fund key (or various aliases) into a display name.
    Expected config shape (aligned with the PDF builder):
      [funds] = [{key, name, abbr, url}], fund_key_map = { "Buligo Industrial Fund": "BuligoIF", ... }
    We prefer 'abbr' for chart labels if present.
    """
    funds = cfg.get("funds", []) or []
    key_to_name = {}
    name_to_key = {}
    for f in funds:
        k = str(f.get("key", "")).strip()
        nm = str(f.get("abbr", "") or f.get("name", "")).strip()
        if k and nm:
            key_to_name[k] = nm
            name_to_key[_norm(nm)] = k
    for nm, k in (cfg.get("fund_key_map", {}) or {}).items():
        if k:
            name_to_key[_norm(str(nm))] = str(k)
    return name_to_key, key_to_name

def _display_name_for_fund(raw_label: str, cfg: dict) -> str:
    """
    If raw_label is a known KEY, return its display name (abbr if provided).
    If raw_label matches a known display name/alias, normalize to the canonical display name.
    Otherwise, return raw_label.
    """
    s = str(raw_label or "").strip()
    if not s:
        return s
    name_to_key, key_to_name = _config_name_maps(cfg)
    if s in key_to_name:
        return key_to_name[s]
    k = name_to_key.get(_norm(s))
    if k and k in key_to_name:
        return key_to_name[k]
    return s

# =======================
# Geometry helpers
# =======================
def _arc_path(cx, cy, r, a0_deg, a1_deg):
    a0 = math.radians(a0_deg)
    a1 = math.radians(a1_deg)
    x0 = cx + r * math.cos(a0); y0 = cy + r * math.sin(a0)
    x1 = cx + r * math.cos(a1); y1 = cy + r * math.sin(a1)
    large = 1 if (a1_deg - a0_deg) % 360 > 180 else 0
    sweep = 1
    return x0, y0, x1, y1, large, sweep

def _donut_slice_path(cx, cy, r_outer, r_inner, a0, a1):
    x0, y0, x1, y1, large, sweep = _arc_path(cx, cy, r_outer, a0, a1)
    xi0, yi0, xi1, yi1, _, _       = _arc_path(cx, cy, r_inner, a1, a0)
    return " ".join([
        f"M {x0:.3f},{y0:.3f}",
        f"A {r_outer:.3f},{r_outer:.3f} 0 {large} {sweep} {x1:.3f},{y1:.3f}",
        f"L {xi0:.3f},{yi0:.3f}",
        f"A {r_inner:.3f},{r_inner:.3f} 0 {large} {1 - sweep} {xi1:.3f},{yi1:.3f}",
        "Z",
    ])

def _donut_full_ring_path(cx, cy, r_outer, r_inner, start_deg=-90.0):
    """
    Build a proper path for a full 360° donut ring using two 180° arcs
    for the outer edge and two 180° arcs (reverse) for the inner edge.
    This avoids the SVG-arc 360° degeneracy.
    """
    a0 = math.radians(start_deg)
    a180 = math.radians(start_deg + 180.0)

    # outer: from a0 -> a0+180 -> a0+360
    x0 = cx + r_outer * math.cos(a0)
    y0 = cy + r_outer * math.sin(a0)
    x180 = cx + r_outer * math.cos(a180)
    y180 = cy + r_outer * math.sin(a180)

    # inner (reverse): from a0+360 -> a0+180 -> a0
    xi0 = cx + r_inner * math.cos(a0)
    yi0 = cy + r_inner * math.sin(a0)
    xi180 = cx + r_inner * math.cos(a180)
    yi180 = cy + r_inner * math.sin(a180)

    d = [
        f"M {x0:.3f},{y0:.3f}",
        # outer two half-arcs (large-arc=0, sweep=1)
        f"A {r_outer:.3f},{r_outer:.3f} 0 0 1 {x180:.3f},{y180:.3f}",
        f"A {r_outer:.3f},{r_outer:.3f} 0 0 1 {x0:.3f},{y0:.3f}",
        # line to inner rim
        f"L {xi0:.3f},{yi0:.3f}",
        # inner two half-arcs in reverse (sweep=0 to go CW)
        f"A {r_inner:.3f},{r_inner:.3f} 0 0 0 {xi180:.3f},{yi180:.3f}",
        f"A {r_inner:.3f},{r_inner:.3f} 0 0 0 {xi0:.3f},{yi0:.3f}",
        "Z",
    ]
    return " ".join(d)


# =======================
# Text helpers + collision avoidance
# =======================
def _add_two_line_label_centered(
    dwg,
    text_top,
    text_bottom,
    x,
    y,
    fill,
    font_family,
    size_top,
    size_bottom,
    line_gap=4,
):
    # Compute vertical positions so the block is centered on (x, y)
    block_height = size_top + line_gap + size_bottom
    y_top_baseline = y - block_height / 2 + size_top
    y_bottom_baseline = y_top_baseline + line_gap + size_bottom

    text_node = dwg.text(
        "",
        insert=(x, 0),
        text_anchor="middle",
        fill=fill,
        font_family=font_family,
    )
    t1 = dwg.tspan(text_top,   x=[x], y=[y_top_baseline],    font_size=size_top)
    t2 = dwg.tspan(text_bottom, x=[x], y=[y_bottom_baseline], font_size=size_bottom)
    text_node.add(t1); text_node.add(t2)
    dwg.add(text_node)

def _estimate_text_block(top: str, bottom: str, fs_top: float, fs_bottom: float, line_gap: float) -> tuple[float, float]:
    """Rough width/height in px for the 2-line label block (sans-serif heuristic)."""
    char_px = 0.55
    w_top = max(1.0, len(str(top))    * fs_top    * char_px)
    w_bot = max(1.0, len(str(bottom)) * fs_bottom * char_px)
    width = max(w_top, w_bot)
    height = fs_top + line_gap + fs_bottom
    return width, height

def _fits(x: float, y: float, w: float, h: float,
          width: float, height_: float, margin: float,
          avoid_rects: list[tuple] | None = None,
          avoid_pad: float = 6.0) -> bool:
    # canvas margins
    if not (x - w/2 >= margin and x + w/2 <= width - margin and
            y - h/2 >= margin and y + h/2 <= height_ - margin):
        return False

    # avoid rectangles (each rect can be (x,y,w,h) or (x,y,w,h,pad))
    if avoid_rects:
        lx, ly = x - w/2, y - h/2
        for rect in avoid_rects:
            if len(rect) == 5:
                rx, ry, rw, rh, pad = rect
            else:
                rx, ry, rw, rh = rect
                pad = avoid_pad
            ax, ay = rx - pad, ry - pad
            aw, ah = rw + 2*pad, rh + 2*pad
            # overlap test
            if not (lx + w <= ax or ax + aw <= lx or ly + h <= ay or ay + ah <= ly):
                return False
    return True



def _adjust_angle_to_fit(
    angle_deg: float,
    cx: float, cy: float,
    outer_radius: float,
    base_offset: float,
    block_w: float, block_h: float,
    width: float, height_: float,
    margin: float,
    avoid_rects: List[Tuple[float,float,float,float]] | None = None,
    avoid_pad: float = 6.0,
    max_iter: int = 120,   # kept for signature
) -> tuple[float, float]:
    """
    Minimal rotation only when overlapping an avoid box,
    plus right-side page-cutoff correction (restores old behavior).
    """
    a = angle_deg % 360.0
    off = float(base_offset)
    R = outer_radius
    MIN_OFF = 24.0
    MAX_OFF = base_offset + 80.0

    def xy(angle, off_):
        r = math.radians(angle)
        return (
            cx + (R + off_) * math.cos(r),
            cy + (R + off_) * math.sin(r),
            r,
        )

    def overlaps_avoid_at(x, y):
        if not avoid_rects:
            return False
        lx, ly = x - block_w/2, y - block_h/2
        for rect in avoid_rects:
            if len(rect) == 5:
                rx, ry, rw, rh, pad = rect
            else:
                rx, ry, rw, rh = rect
                pad = avoid_pad
            ax, ay = rx - pad, ry - pad
            aw, ah = rw + 2*pad, rh + 2*pad
            if not (lx + block_w <= ax or ax + aw <= lx or ly + block_h <= ay or ay + ah <= ly):
                return True
        return False

    # ---- 1) centered position
    x0, y0, rad0 = xy(a, off)
    if not overlaps_avoid_at(x0, y0):
        c, s = math.cos(rad0), math.sin(rad0)

        # --- RIGHT-SIDE page-cutoff rotation (old behavior) ---
        if c > 0 and (x0 + block_w/2) > (width - margin):
            step = 1.0
            for _ in range(40):  # up to ±20°
                a += step if s > 0 else -step
                xt, yt, _ = xy(a, off)
                if (xt + block_w/2) <= (width - margin):
                    break
            x0, y0, rad0 = xy(a, off)
            c, s = math.cos(rad0), math.sin(rad0)

        # --- analytic minimal outward radius (same as your current good code) ---
        deltas = []
        if (x0 + block_w/2) > (width - margin) and c > 0:
            target = (width - margin - block_w/2 - cx) / c
            deltas.append(max(0.0, target - (R + off)))
        if (x0 - block_w/2) < margin and c < 0:
            target = (margin + block_w/2 - cx) / c
            deltas.append(max(0.0, target - (R + off)))
        if (y0 + block_h/2) > (height_ - margin) and s > 0:
            target = (height_ - margin - block_h/2 - cy) / s
            deltas.append(max(0.0, target - (R + off)))
        if (y0 - block_h/2) < margin and s < 0:
            target = (margin + block_h/2 - cy) / s
            deltas.append(max(0.0, target - (R + off)))
        if deltas:
            dr = max(deltas)
            if dr > 0:
                off = min(MAX_OFF, max(MIN_OFF, off + dr))
        return a, off

    # ---- 2) overlap with avoid box → minimal rotation to clear, then margin tidy-up ----
    step = 0.5
    max_search = 30.0
    best_a = a
    found = False
    for delta in (d*step for d in range(1, int(max_search/step) + 1)):
        for sign in (+1, -1):
            ta = (a + sign*delta) % 360.0
            xt, yt, _ = xy(ta, off)
            if not overlaps_avoid_at(xt, yt):
                best_a = ta
                found = True
                break
        if found:
            break

    xb, yb, rb = xy(best_a, off)
    cb, sb = math.cos(rb), math.sin(rb)
    deltas = []
    if (xb + block_w/2) > (width - margin) and cb > 0:
        target = (width - margin - block_w/2 - cx) / cb
        deltas.append(max(0.0, target - (R + off)))
    if (xb - block_w/2) < margin and cb < 0:
        target = (margin + block_w/2 - cx) / cb
        deltas.append(max(0.0, target - (R + off)))
    if (yb + block_h/2) > (height_ - margin) and sb > 0:
        target = (height_ - margin - block_h/2 - cy) / sb
        deltas.append(max(0.0, target - (R + off)))
    if (yb - block_h/2) < margin and sb < 0:
        target = (margin + block_h/2 - cy) / sb
        deltas.append(max(0.0, target - (R + off)))
    if deltas:
        dr = max(deltas)
        if dr > 0:
            off = min(MAX_OFF, max(MIN_OFF, off + dr))

    return best_a, off





# =======================
# Renderers
# =======================
def donut_svg(
    svg_path,
    data: Dict[str, float],
    colors: Dict[str, str],
    width=900,
    height=600,
    outer_radius=230,
    inner_radius=110,                # smaller => thicker ring
    start_deg=-90.0,                 # 12 o’clock start
    text_color="#ffffff",
    font_family=FONT_MED_NAME,
    name_font_size=20,
    pct_font_size=18,
    label_offset=24,                 # base radial offset for labels
    margin: int = 36,                # keep labels off the edges
):
    assert inner_radius < outer_radius
    cx, cy = width / 2, height / 2

    dwg = svgwrite.Drawing(svg_path, size=(width, height), profile="full")
    dwg.attribs["viewBox"] = f"0 0 {width} {height}"

    values = [float(v) for v in data.values()]
    total = sum(values) or 1.0

    # slices
    angle = float(start_deg)
    for label, value in data.items():
        frac = float(value) / total
        sweep = 360.0 * frac
        if sweep >= 359.999:
            d = _donut_full_ring_path(cx, cy, outer_radius, inner_radius, angle)
        else:
            d = _donut_slice_path(cx, cy, outer_radius, inner_radius, angle, angle + sweep)
        dwg.add(dwg.path(d=d, fill=colors.get(label, "#cccccc"), stroke="none"))
        angle += sweep

    # labels (adaptive)
    angle = float(start_deg)
    for label, value in data.items():
        frac = float(value) / total
        sweep = 360.0 * frac
        mid_deg = (angle + sweep / 2.0) % 360.0

        pct_txt = f"{round(frac * 100)}%"
        block_w, block_h = _estimate_text_block(str(label), pct_txt, name_font_size, pct_font_size, line_gap=4)

        # gentle dynamic offset that keeps labels from hugging the left/right
        rad_mid = math.radians(mid_deg)
        est_name_width = max(1.0, len(str(label)) * (name_font_size * 0.55))
        horiz_factor = abs(math.cos(rad_mid))
        base_offset = float(label_offset) + (0.30 * est_name_width * horiz_factor + 2)

        adj_deg, adj_offset = _adjust_angle_to_fit(
            angle_deg=mid_deg,
            cx=cx, cy=cy,
            outer_radius=outer_radius,
            base_offset=base_offset,
            block_w=block_w, block_h=block_h,
            width=width, height_=height,
            margin=margin,
        )

        rad = math.radians(adj_deg)
        x_lab = cx + (outer_radius + adj_offset) * math.cos(rad)
        y_lab = cy + (outer_radius + adj_offset) * math.sin(rad)

        _add_two_line_label_centered(
            dwg,
            text_top=str(label),
            text_bottom=pct_txt,
            x=x_lab,
            y=y_lab,
            fill=text_color,
            font_family=font_family,
            size_top=name_font_size,
            size_bottom=pct_font_size,
            line_gap=4,
        )
        angle += sweep

    dwg.save()

def donut_svg_square(
    svg_path: str,
    data: Dict[str, float],
    colors: Dict[str, str],
    size: int = 900,                 # square canvas
    ring_ratio: float = 0.478,       # inner = ring_ratio * outer  (≈110/230)
    margin: int = 60,                # padding to keep labels off edges
    start_deg: float = -90.0,
    text_color: str = "#ffffff",
    font_family: str = FONT_MED_NAME,
    name_font_size: int = 20,
    pct_font_size: int = 18,
    label_offset_px: int | None = 35,  # base offset before adaptive tweaks
    debug_avoid: bool = False,
):
    width = height = int(size)
    cx, cy = width / 2, height / 2
    base_label_offset = label_offset_px if label_offset_px is not None else max(28, int(0.07 * size))

    # ring sizes
    outer_radius = max(40, 230 * (size / 900.0))
    inner_radius = max(10, outer_radius * ring_ratio)

    # keep ring inside square after labels/margins
    max_outer = (size / 2) - (base_label_offset + margin)
    outer_radius = min(outer_radius, max_outer)
    inner_radius = max(10, outer_radius * ring_ratio)

    dwg = svgwrite.Drawing(svg_path, size=(width, height), profile="full")
    dwg.attribs["viewBox"] = f"0 0 {width} {height}"

    avoid_rects = [
    (-220, 420, 300, 60),        # box 1 (uses default avoid_pad)
    (-270, 430, 330, 180),   # box 2 with custom pad (optional 5th value)
    (-170, 135, 390, 60),
    ]
    avoid_pad = 10.0

    # Draw every box (visible red) and its padded avoid area (pink)
    if debug_avoid:
        for rect in avoid_rects:
            if len(rect) == 5:
                rx, ry, rw, rh, pad = rect
            else:
                rx, ry, rw, rh = rect
                pad = avoid_pad
            dwg.add(dwg.rect(insert=(rx, ry), size=(rw, rh),
                            fill="none", stroke="#ff5c5c", stroke_dasharray="6,4"))
            dwg.add(dwg.rect(insert=(rx - pad, ry - pad),
                            size=(rw + 2*pad, rh + 2*pad),
                            fill="#ffb3b3", opacity=0.25, stroke="none"))


    values = [float(v) for v in data.values()]
    total = sum(values) or 1.0

    # slices
    angle = float(start_deg)
    for label, value in data.items():
        frac = float(value) / total
        sweep = 360.0 * frac
        if sweep >= 359.999:
            d = _donut_full_ring_path(cx, cy, outer_radius, inner_radius, angle)
        else:
            d = _donut_slice_path(cx, cy, outer_radius, inner_radius, angle, angle + sweep)
        dwg.add(dwg.path(d=d, fill=colors.get(label, "#cccccc"), stroke="none"))
        angle += sweep

    # labels (adaptive)
    angle = float(start_deg)
    for label, value in data.items():
        frac = float(value) / total
        sweep = 360.0 * frac
        mid_deg = (angle + sweep / 2.0) % 360.0

        pct_txt = f"{round(frac * 100)}%"
        block_w, block_h = _estimate_text_block(str(label), pct_txt, name_font_size, pct_font_size, line_gap=4)

        rad_mid = math.radians(mid_deg)
        est_name_width = max(1.0, len(str(label)) * (name_font_size * 0.55))
        horiz_factor = abs(math.cos(rad_mid))
        base_offset = float(base_label_offset) + (0.30 * est_name_width * horiz_factor + 2)

        adj_deg, adj_offset = _adjust_angle_to_fit(
            angle_deg=mid_deg,
            cx=cx, cy=cy,
            outer_radius=outer_radius,
            base_offset=base_offset,
            block_w=block_w, block_h=block_h,
            width=width, height_=height,
            margin=margin,
            avoid_rects=avoid_rects,
            avoid_pad=avoid_pad,
        )

        rad = math.radians(adj_deg)
        x_lab = cx + (outer_radius + adj_offset) * math.cos(rad)
        y_lab = cy + (outer_radius + adj_offset) * math.sin(rad)

        _add_two_line_label_centered(
            dwg,
            text_top=str(label),
            text_bottom=pct_txt,
            x=x_lab,
            y=y_lab,
            fill=text_color,
            font_family=font_family,
            size_top=name_font_size,
            size_bottom=pct_font_size,
            line_gap=4,
        )
        angle += sweep

    dwg.save()

# =======================
# Colors + IO
# =======================
def _stable_color_cycle(n: int) -> Iterable[str]:
    base = [
        "#41b8d5",  # blue
        "#a9d6b9",  # green
        "#2d8bba",  # dark blue
        "#505870",  # navy
        "#7d87a4",  # dark grey
        "#c4dce8",  # grey
        "#ffffff",  # white
        "#2E8BC0",  # deep sky blue
    ]
    if n <= len(base):
        return base[:n]
    out = list(base)
    def tweak(hx, k):
        r = int(hx[1:3], 16); g = int(hx[3:5], 16); b = int(hx[5:7], 16)
        r = max(0, min(255, int(r * (0.9 + 0.02 * k))))
        g = max(0, min(255, int(g * (0.9 + 0.02 * k))))
        b = max(0, min(255, int(b * (0.9 + 0.02 * k))))
        return f"#{r:02x}{g:02x}{b:02x}"
    k = 1
    while len(out) < n:
        for c in base:
            out.append(tweak(c, k))
            if len(out) >= n:
                break
        k += 1
    return out

def _fund_color_map(fund_labels: Iterable[str]) -> Dict[str, str]:
    labels = list(fund_labels)
    colors = list(_stable_color_cycle(len(labels)))
    return {lab: col for lab, col in zip(labels, colors)}

def _safe_slug(text: str) -> str:
    s = "".join(ch if ch.isalnum() else "_" for ch in str(text).strip())
    while "__" in s:
        s = s.replace("__", "_")
    return s.strip("_") or "chart"

def _to_png(svg_path: str, png_path: str):
    """Convert SVG → PNG. Prefer rsvg-convert; fallback to cairosvg if installed."""
    rsvg = shutil.which("rsvg-convert")
    if rsvg:
        os.system(f'"{rsvg}" "{svg_path}" -a -f png -o "{png_path}"')
        return
    try:
        import cairosvg  # type: ignore
        cairosvg.svg2png(url=svg_path, write_to=png_path)
    except Exception as exc:
        print(f"[warn] PNG not created for {svg_path}: {exc}")

# =======================
# Excel → multiple charts
# =======================
def generate_investor_donuts_from_excel(
    excel_path: str = EXCEL_PATH,
    out_dir: str = CHART_DIR,
    sheet_name: str = "Investor Data - LongForm",
    investor_col: str = "Investor Name",
    fund_col: str = "Fund Name",
    amount_col: str = "Amount ($)",
    *,
    square: bool = True,
    square_size: int = 900,
):
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df = df[[investor_col, fund_col, amount_col]].copy().dropna(subset=[investor_col, fund_col, amount_col])
    df[amount_col] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0)

    grouped = (
        df.groupby([investor_col, fund_col], dropna=False)[amount_col]
          .sum()
          .reset_index()
    )

    # Build color map from display names (so colors are consistent even if raw labels vary)
    all_funds = (
        grouped[fund_col]
        .dropna()
        .astype(str)
        .map(lambda v: _display_name_for_fund(v, _CFG))
        .unique()
        .tolist()
    )
    global_colors = _fund_color_map(all_funds)

    emitted: List[Tuple[str, str, str]] = []

    for investor, sub in grouped.groupby(investor_col):
        pairs = [
            (_display_name_for_fund(lbl, _CFG), float(val))
            for lbl, val in sub[[fund_col, amount_col]].values.tolist()
            if float(val) > 0
        ]
        if not pairs:
            continue

        pairs.sort(key=lambda x: x[1], reverse=True)
        data = {lbl: val for lbl, val in pairs}
        if len(data) == 1:
            # Single-fund donut: force the first palette color
            only_label = next(iter(data.keys()))
            colors = {only_label: _stable_color_cycle(1)[0]}   # or "#C7DCE9" if you prefer a fixed hex
        else:
            # Multi-fund donut: keep global, fund-stable colors
            colors = {lbl: global_colors.get(lbl, "#cccccc") for lbl in data.keys()}

        base = _safe_slug(str(investor))
        svg_path = os.path.join(out_dir, f"{base}.svg")
        png_path = os.path.join(out_dir, f"{base}.png")

        if square:
            donut_svg_square(
                svg_path=svg_path,
                data=data,
                colors=colors,
                size=square_size,
                ring_ratio=0.478,
                margin=90,
                start_deg=-90.0,
                text_color="#ffffff",
                name_font_size=24,
                pct_font_size=22,
                label_offset_px=35,  # modest base; adaptive logic will nudge per-label
            )
        else:
            donut_svg(
                svg_path=svg_path,
                data=data,
                colors=colors,
                width=900,
                height=600,
                outer_radius=230,
                inner_radius=110,
                start_deg=-90.0,
                text_color="#ffffff",
                font_family=FONT_MED_NAME,
                name_font_size=24,
                pct_font_size=22,
                label_offset=24,
                margin=36,
            )

        _to_png(svg_path, png_path)
        emitted.append((str(investor), svg_path, png_path))

    return emitted

if __name__ == "__main__":
    results = generate_investor_donuts_from_excel(EXCEL_PATH, CHART_DIR, square=True, square_size=900)
    print(f"Emitted {len(results)} charts to '{CHART_DIR}'")
    for inv, svg, png in results:
        print(f"- {inv}: {os.path.basename(svg)}  |  {os.path.basename(png)}")
