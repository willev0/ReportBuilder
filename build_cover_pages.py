# build_cover_pages.py
# Generate investor cover pages by stamping text onto CoverPageTemplate.pdf
# - Uses cover_name/name from config [[funds]]
# - Wraps fund names within adjustable bounds with adjustable line spacings

import os, io
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import tomllib  # stdlib (3.11+) to read config.toml (Python 3.11+)
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from pypdf import PdfReader, PdfWriter

# ========= load config (static brand + paths only) =========
def _load_cfg() -> dict:
    here = Path(__file__).parent
    cfg_path = here / "Configs" / "config.toml"
    if cfg_path.exists():
        with open(cfg_path, "rb") as f:
            return tomllib.load(f)
    return {}

_cfg = _load_cfg()

# ========= CONFIG (from config.toml, with sane defaults) =========
HERE = Path(__file__).parent

EXCEL_PATH   = str(HERE / _cfg.get("paths", {}).get("excel", "InvestorDataTest.xlsx"))
TEMPLATE_PDF = str(HERE / _cfg.get("paths", {}).get("cover_template_pdf", "Configs/CoverPageTemplate.pdf"))
OUTPUT_DIR   = str(HERE / _cfg.get("paths", {}).get("out_cover", "CoverPages"))

FONT_BOLD_NAME = _cfg.get("fonts", {}).get("bold_name", "HKGrotesk-Bold")
FONT_MED_NAME  = _cfg.get("fonts", {}).get("medium_name", "HKGrotesk-Medium")
FONT_BOLD_PATH = str(HERE / _cfg.get("fonts", {}).get("bold_path", "Configs/hk-grotesk.bold.ttf"))
FONT_MED_PATH  = str(HERE / _cfg.get("fonts", {}).get("medium_path", "Configs/hk-grotesk.medium.ttf"))

ACCENT_GRN = HexColor(_cfg.get("brand", {}).get("accent", "#A9D6B9"))

# ========= FUNDS LAYOUT (tweak here) =========
# Left X position for the funds list
FUNDS_LEFT_X: float = 270.0
# Right boundary for wrapping (text will wrap before exceeding this X)
FUNDS_RIGHT_X: float = 540.0  # adjust to change the wrap width
# Font settings for funds list
FUNDS_FONT_NAME: str = FONT_MED_NAME
FUNDS_FONT_SIZE: int = 22
# Line spacing within a single (wrapped) fund name
LINE_SPACING_INTRA: float = 26.0  # distance between wrapped lines of the SAME fund
# Spacing between different fund names (slightly larger)
LINE_SPACING_INTER: float = 36.0  # distance between last line of fund A and first line of fund B
# Starting Y for the funds block
FUNDS_START_Y_OFFSET: float = 350.0  # measured down from top of page
# Auto-shrink bounds
MIN_FUNDS_FONT_SIZE: int = 14      # won't shrink below this
FUNDS_BOTTOM_Y_GUARD: float = 180  # keep the funds block ABOVE this Y (protects "PREPARED FOR"/name)


# ========= FUND DISPLAY NAME RESOLUTION =========
def _norm(s: Optional[str]) -> Optional[str]:
    """Normalize keys/names for robust matching."""
    if s is None:
        return None
    s = str(s).strip()
    return " ".join(s.split())  # collapse internal whitespace

def _build_fund_display_map(cfg: dict) -> Dict[str, str]:
    """
    Build a lookup so that any of these inputs:
      - fund key (e.g., 10FSSAC3)
      - fund name from config (e.g., "10 Federal SSAC Fund 3")
      - optional legacy aliases from cfg['fund_name_map'] (if present)
    all map to the display string chosen as:
      cover_name (if non-empty) else name (else key).
    """
    mapping: Dict[str, str] = {}

    for f in cfg.get("funds", []) or []:
        key = _norm(f.get("key"))
        name = _norm(f.get("name"))
        cover_name_raw = f.get("cover_name")
        cover_name = _norm(cover_name_raw) if cover_name_raw is not None else None

        display = cover_name or name or key or ""
        if not display:
            continue  # skip malformed entries

        # Map both key and name to the chosen display string (exact + lower fallback)
        for alias in (key, name):
            if alias:
                mapping[alias] = display
                mapping[alias.lower()] = display

    # Optional legacy support: allow user to provide extra aliases
    legacy_map = cfg.get("fund_name_map", {}) or {}
    for k, v in legacy_map.items():
        kn = _norm(k)
        vn = _norm(v) or v
        if kn and vn:
            mapping[kn] = vn
            mapping[kn.lower()] = vn

    return mapping

_FUND_DISPLAY_MAP = _build_fund_display_map(_cfg)

def _pretty_fund(code_or_name: str) -> str:
    """
    Given whatever the Excel “Fund Name” column has (or a key),
    return the display name from config:
      cover_name -> name -> key -> original string (fallback).
    """
    if code_or_name is None:
        return ""
    raw = str(code_or_name)
    n = _norm(raw)
    return (
        _FUND_DISPLAY_MAP.get(n)
        or _FUND_DISPLAY_MAP.get(n.lower())
        or raw
    )

# ========= FONTS =========
def register_fonts():
    try:
        pdfmetrics.registerFont(TTFont(FONT_BOLD_NAME, FONT_BOLD_PATH))
    except Exception:
        pass
    try:
        pdfmetrics.registerFont(TTFont(FONT_MED_NAME, FONT_MED_PATH))
    except Exception:
        pass

# ========= QUARTER HELPERS =========
def most_recent_quarter(df_fund: pd.DataFrame):
    try:
        q = (
            df_fund["Quarter"].dropna().astype(str)
            .str.extract(r"(?P<year>\d{4})-Q(?P<q>[1-4])")
            .dropna()
        )
        q["year"] = q["year"].astype(int)
        q["q"]    = q["q"].astype(int)
        q = q.sort_values(["year", "q"])
        if q.empty:
            return "Q2 2025", "JUNE 30, 2025"
        row = q.iloc[-1]
        q_label = f"Q{row['q']} {row['year']}"
        if row['q'] == 1:
            date_str = f"MARCH 31, {row['year']}"
        elif row['q'] == 2:
            date_str = f"JUNE 30, {row['year']}"
        elif row['q'] == 3:
            date_str = f"SEPTEMBER 30, {row['year']}"
        else:
            date_str = f"DECEMBER 31, {row['year']}"
        return q_label, date_str
    except Exception:
        return "QX 20XX", "Date N/A"

# ========= TEXT WRAPPING UTILITIES =========
def _string_width(text: str, font_name: str, font_size: float) -> float:
    return pdfmetrics.stringWidth(text, font_name, font_size)

def _wrap_text_by_words(
    text: str,
    font_name: str,
    font_size: float,
    max_width: float
) -> List[str]:
    """
    Greedy word-wrapping using measured widths.
    Falls back to character wrapping if a single token exceeds max_width.
    """
    if not text:
        return [""]

    words = text.split()
    lines: List[str] = []
    current: List[str] = []

    def line_width(parts: List[str]) -> float:
        s = " ".join(parts) if parts else ""
        return _string_width(s, font_name, font_size)

    for w in words:
        # If the word itself is too long, split by characters
        if _string_width(w, font_name, font_size) > max_width:
            # flush current line first
            if current:
                lines.append(" ".join(current))
                current = []
            # split long token
            chunk = ""
            for ch in w:
                if _string_width(chunk + ch, font_name, font_size) <= max_width:
                    chunk += ch
                else:
                    if chunk:
                        lines.append(chunk)
                    chunk = ch
            if chunk:
                lines.append(chunk)
            continue

        trial = (current + [w]) if current else [w]
        if line_width(trial) <= max_width:
            current = trial
        else:
            # push current and start a new line with w
            lines.append(" ".join(current) if current else "")
            current = [w]

    if current:
        lines.append(" ".join(current))

    # Ensure at least one line
    if not lines:
        lines = [""]

    return lines


def _measure_funds_block_height(
    funds: List[str],
    font_name: str,
    font_size: float,
    max_width: float,
    intra_spacing: float,
    inter_spacing: float,
) -> float:
    """
    Simulate _draw_wrapped_funds to compute how much vertical space the funds block will consume.
    Matches the same decrement logic used during drawing:
      total drop = sum((len(lines)-1)*intra_spacing + inter_spacing) over all funds
    """
    total_drop = 0.0
    for f in funds or []:
        display = _pretty_fund(str(f))
        lines = _wrap_text_by_words(display, font_name, font_size, max_width)
        total_drop += max(0, (len(lines) - 1)) * intra_spacing
        total_drop += inter_spacing
    return total_drop


def _draw_wrapped_funds(
    c: canvas.Canvas,
    funds: List[str],
    left_x: float,
    right_x: float,
    start_y: float,
    font_name: str,
    font_size: float,
    intra_spacing: float,
    inter_spacing: float,
) -> float:
    """
    Draw a list of fund names, wrapping each to (right_x - left_x).
    Returns the final y after drawing.
    """
    max_width = max(0.0, right_x - left_x)
    c.setFont(font_name, font_size)

    y = start_y
    for f in funds:
        display = _pretty_fund(str(f))
        lines = _wrap_text_by_words(display, font_name, font_size, max_width)

        # draw wrapped lines for this fund
        for i, line in enumerate(lines):
            c.drawString(left_x, y, line)
            # intra-line spacing except after last line of this fund
            if i < len(lines) - 1:
                y -= intra_spacing

        # after the last line of this fund, apply inter-fund spacing
        y -= inter_spacing

    return y

# ========= OVERLAY =========
def paint_overlay(investor: str, funds: List[str], quarter_label: str, report_date: str) -> bytes:
    width, height = letter
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)

    # Title block
    c.setFillColor(colors.white)
    c.setFont(FONT_MED_NAME, 46)
    c.drawString(270, height - 200, "QUARTERLY")
    c.drawString(270, height - 250, "REPORT")

    # Quarter
    c.setFont(FONT_MED_NAME, 46)
    c.drawString(270, height - 300, quarter_label)

    # Date (top right)
    c.setFont(FONT_MED_NAME, 18)
    c.drawRightString(width - 80, height - 90, report_date)

    # Funds list (wrapped) — autosize to fit above the guard line
    c.setFillColor(colors.white)
    funds_start_y = height - FUNDS_START_Y_OFFSET
    max_width = max(0.0, FUNDS_RIGHT_X - FUNDS_LEFT_X)

    # Start from configured size; shrink until it fits or we hit the minimum.
    font_size = float(FUNDS_FONT_SIZE)
    while font_size > MIN_FUNDS_FONT_SIZE:
        # scale spacings with font size so leading looks right when shrinking
        scale = font_size / float(FUNDS_FONT_SIZE)
        intra = LINE_SPACING_INTRA * scale
        inter = LINE_SPACING_INTER * scale

        needed = _measure_funds_block_height(
            funds=funds,
            font_name=FUNDS_FONT_NAME,
            font_size=font_size,
            max_width=max_width,
            intra_spacing=intra,
            inter_spacing=inter,
        )
        # Will the block end above the guard?
        if (funds_start_y - needed) >= FUNDS_BOTTOM_Y_GUARD:
            break
        font_size -= 1.0  # shrink and try again

    # Draw with the chosen size and scaled spacings
    scale = font_size / float(FUNDS_FONT_SIZE)
    intra = LINE_SPACING_INTRA * scale
    inter = LINE_SPACING_INTER * scale
    _draw_wrapped_funds(
        c=c,
        funds=funds,
        left_x=FUNDS_LEFT_X,
        right_x=FUNDS_RIGHT_X,
        start_y=funds_start_y,
        font_name=FUNDS_FONT_NAME,
        font_size=font_size,
        intra_spacing=intra,
        inter_spacing=inter,
    )


    # Prepared for
    c.setFillColor(ACCENT_GRN)
    c.setFont(FONT_BOLD_NAME, 20)
    c.drawString(270, 150, "PREPARED FOR")

    c.setFillColor(colors.white)
    c.setFont(FONT_MED_NAME, 18)
    c.drawString(270, 120, investor)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

def merge_overlay(template_pdf: str, overlay_bytes: bytes, out_path: Path):
    base_reader = PdfReader(template_pdf)
    base_page = base_reader.pages[0]
    overlay_reader = PdfReader(io.BytesIO(overlay_bytes))
    overlay_page = overlay_reader.pages[0]
    base_page.merge_page(overlay_page)
    writer = PdfWriter()
    writer.add_page(base_page)
    with open(out_path, "wb") as f:
        writer.write(f)

# ========= MAIN =========
def main():
    register_fonts()
    outdir = Path(OUTPUT_DIR)
    outdir.mkdir(exist_ok=True)

    df_long = pd.read_excel(EXCEL_PATH, sheet_name="Investor Data - LongForm")
    df_fund = pd.read_excel(EXCEL_PATH, sheet_name="FundData")
    q_label, q_date = most_recent_quarter(df_fund)

    groups = (
        df_long.groupby("Investor Name")["Fund Name"]
               .apply(lambda s: sorted(pd.unique(s)))
               .to_dict()
    )

    for investor, funds in groups.items():
        safe = "".join(ch for ch in investor if ch.isalnum() or ch in (" ", "_", "-", ".")).strip().replace(" ", "_")
        out_path = outdir / f"{safe}_cover.pdf"
        overlay_bytes = paint_overlay(investor, funds, q_label, q_date)
        merge_overlay(TEMPLATE_PDF, overlay_bytes, out_path)
        print(f"[OK] {investor} -> {out_path}")

if __name__ == "__main__":
    main()
