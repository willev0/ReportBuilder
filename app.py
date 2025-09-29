#!/usr/bin/env python3
"""
Streamlit UI for the new, simplified report generation flow.
Now there are ONLY two builder scripts, both driven by config.toml:
  1) build_cover_pages.py       → outputs to [paths.out_cover]
  2) BuildInvestorPage.py       → outputs to [paths.out_consolidated] (per‑investor pages)

This app lets you upload the Excel, run those two builders, and merge a final
per‑investor PDF (Cover → InvestorPage). It also creates an all‑in‑one
InvestorReports.pdf and a ZIP of all finals.
"""
from __future__ import annotations
import io
import os
import shutil
import traceback
import importlib.util
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from pypdf import PdfReader, PdfWriter
import streamlit as st

# ----------------------------
# Project layout & constants
# ----------------------------
HERE = Path(__file__).resolve().parent

BUILD_COVER_PAGES   = HERE / "build_cover_pages.py"
BUILD_INVESTOR_PAGE = HERE / "BuildInvestorPage.py"

# Defaults mirror config.toml keys; you can override from the sidebar
DEFAULT_COVER_DIR        = "CoverPages"
DEFAULT_INVESTOR_DIR     = "2ConsolidatedStatementPage"  # BuildInvestorPage.py default OUTPUT_DIR
FINAL_DIR_NAME           = "FinalInvestorReports"

# ----------------------------
# Helpers
# ----------------------------

def log(msg: str) -> None:
    st.session_state.setdefault("log", [])
    st.session_state.log.append(msg)


def reset_log() -> None:
    st.session_state["log"] = []


def load_module_by_path(module_path: Path, modname: str):
    spec = importlib.util.spec_from_file_location(modname, str(module_path))
    if not spec or not spec.loader:
        raise RuntimeError(f"Could not load {modname} from {module_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # type: ignore
    return mod


def safe_slug(text: str) -> str:
    s = "".join(ch if ch.isalnum() else "_" for ch in str(text).strip())
    while "__" in s:
        s = s.replace("__", "_")
    return s.strip("_") or "chart"


def filename_for_cover(investor: str) -> str:
    safe = "".join(ch for ch in str(investor) if ch.isalnum() or ch in (" ", "_", "-"))\
            .strip().replace(" ", "_")
    return f"{safe}_cover.pdf"

# ----------------------------
# Data containers
# ----------------------------

@dataclass
class BuilderOutputs:
    root: Path
    cover_dir: Path
    investor_dir: Path
    final_dir: Path

# ----------------------------
# Core steps
# ----------------------------

def prepare_output_dirs(root: Path, clean: bool, cover_dir_name: str, investor_dir_name: str) -> BuilderOutputs:
    cover_dir = root / cover_dir_name
    investor_dir = root / investor_dir_name
    final_dir = root / FINAL_DIR_NAME

    for p in (cover_dir, investor_dir, final_dir):
        if clean and p.exists():
            shutil.rmtree(p, ignore_errors=True)
        p.mkdir(parents=True, exist_ok=True)
    return BuilderOutputs(root, cover_dir, investor_dir, final_dir)


def run_builder_module(module_path: Path, modname: str, excel_path: Path, output_dir: Path) -> None:
    """Load a builder by path and run it, setting EXCEL_PATH/OUTPUT_DIR if present.
    Accepts multiple entry points to be compatible with older/newer builder scripts.
    Preferred order: generate_investor_reports → generate_consolidated_pages → main.
    """
    mod = load_module_by_path(module_path, modname)
    # Set common globals if they exist (both builders respect these)
    if hasattr(mod, "EXCEL_PATH"):
        setattr(mod, "EXCEL_PATH", str(excel_path))
    if hasattr(mod, "OUTPUT_DIR"):
        setattr(mod, "OUTPUT_DIR", str(output_dir))

    # Try known entry points
    try:
        if hasattr(mod, "generate_investor_reports"):
            mod.generate_investor_reports(str(excel_path), str(output_dir))  # type: ignore
        elif hasattr(mod, "generate_consolidated_pages"):
            mod.generate_consolidated_pages(str(excel_path), str(output_dir))  # type: ignore
        elif hasattr(mod, "main"):
            mod.main()  # type: ignore
        else:
            tried = ["generate_investor_reports(excel, out)", "generate_consolidated_pages(excel, out)", "main()"]
            raise RuntimeError("Module is missing an entry point. Tried: " + ", ".join(tried))
    except Exception as e:
        raise RuntimeError(f"Failed running {module_path.name}: {e}\n{traceback.format_exc()}")


def merge_final_reports(excel_path: Path, outs: BuilderOutputs, sheet_name: str, investor_col: str) -> Tuple[int, List[Path]]:
    """Merge per‑investor PDFs in order: Cover → InvestorPage."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    if investor_col not in df.columns:
        raise ValueError(
            f"Column '{investor_col}' not found in sheet '{sheet_name}'. Columns present: {list(df.columns)}"
        )
    investors = sorted(df[investor_col].dropna().astype(str).unique().tolist())

    finals: List[Path] = []
    built = 0
    for inv in investors:
        parts: List[Path] = []

        cover_pdf = outs.cover_dir / filename_for_cover(inv)
        if cover_pdf.exists():
            parts.append(cover_pdf)
        else:
            log(f"⚠️ Missing cover for '{inv}': {cover_pdf.name}")

        # BuildInvestorPage.py writes <slug>.pdf in its OUTPUT_DIR
        investor_pdf = outs.investor_dir / f"{safe_slug(inv)}.pdf"
        if investor_pdf.exists():
            parts.append(investor_pdf)
        else:
            log(f"⚠️ Missing investor page for '{inv}': {investor_pdf.name}")

        if not parts:
            continue

        writer = PdfWriter()
        for p in parts:
            try:
                reader = PdfReader(str(p))
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                log(f"❌ Error reading {p.name}: {e}")

        final_name = f"{safe_slug(inv)}.pdf"
        final_path = outs.final_dir / final_name
        with open(final_path, "wb") as f:
            writer.write(f)
        finals.append(final_path)
        built += 1
        log(f"✅ Merged {final_name}")

    return built, finals


def zip_files(paths: List[Path], zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, arcname=p.name)


def merge_all_into_one(paths: List[Path], out_pdf: Path) -> None:
    writer = PdfWriter()
    for p in paths:
        try:
            r = PdfReader(str(p))
            for page in r.pages:
                writer.add_page(page)
        except Exception as e:
            log(f"❌ Skipping {p.name} while building master PDF: {e}")
    with open(out_pdf, "wb") as f:
        writer.write(f)

# ----------------------------
# Streamlit UI
# ----------------------------

def main() -> None:
    st.set_page_config(page_title="Investor Report Builder (2‑step)", page_icon="📄", layout="centered")
    if "log" not in st.session_state:
        reset_log()

    st.title("Investor Report Builder")
    st.caption("Upload the Excel, run the two builders, and merge a per‑investor report.")

    # Sidebar options
    with st.sidebar:
        st.header("Options")
        sheet_name = st.text_input("Excel sheet name", value="Investor Data - LongForm")
        investor_col = st.text_input("Investor name column", value="Investor Name")
        cover_dir_name = st.text_input("Cover output folder", value=DEFAULT_COVER_DIR)
        investor_dir_name = st.text_input("Investor pages output folder", value=DEFAULT_INVESTOR_DIR)
        clean_run = st.checkbox("Clean previous outputs first", value=True)
        st.caption("If folders contain old PDFs, enable cleanup to avoid stale merges.")

    uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], accept_multiple_files=False)

    # Detect available builders (only two now)
    detect_cols = st.columns(2)
    with detect_cols[0]:
        have_cover = BUILD_COVER_PAGES.exists()
        run_cover = st.toggle("Run cover pages", value=have_cover, disabled=not have_cover)
    with detect_cols[1]:
        have_investor = BUILD_INVESTOR_PAGE.exists()
        run_investor = st.toggle("Run investor pages", value=have_investor, disabled=not have_investor)

    st.divider()
    run = st.button("Run builders and merge reports", type="primary", use_container_width=True, disabled=uploaded is None)

    st.write("### Log")
    log_area = st.empty()
    log_area.code("\n".join(st.session_state.log) or "Ready.", language="text")

    if not run:
        return

    reset_log()
    if not uploaded:
        st.error("Please upload an Excel file first.")
        return

    # Save uploaded Excel to a temp run directory under ./runs/YYYYMMDD_HHMMSS
    run_root = HERE / "runs" / pd.Timestamp.now(tz=None).strftime("%Y%m%d_%H%M%S")
    run_root.mkdir(parents=True, exist_ok=True)
    excel_path = run_root / "source.xlsx"
    with open(excel_path, "wb") as f:
        f.write(uploaded.read())

    outs = prepare_output_dirs(run_root, clean_run, cover_dir_name, investor_dir_name)

    # Execute selected builders, logging progress live
    try:
        if run_cover:
            log("▶️ Running cover pages …")
            run_builder_module(BUILD_COVER_PAGES, "build_cover_pages", excel_path, outs.cover_dir)
            log("✅ Cover pages done.")
        else:
            log("⏩ Skipping cover pages.")
        log_area.code("\n".join(st.session_state.log), language="text")

        if run_investor:
            log("▶️ Running investor pages …")
            run_builder_module(BUILD_INVESTOR_PAGE, "BuildInvestorPage", excel_path, outs.investor_dir)
            log("✅ Investor pages done.")
        else:
            log("⏩ Skipping investor pages.")
        log_area.code("\n".join(st.session_state.log), language="text")

        # Merge
        log("Merging per‑investor PDFs …")
        built, finals = merge_final_reports(excel_path, outs, sheet_name=sheet_name, investor_col=investor_col)
        log(f"Built {built} final report(s) → {outs.final_dir}")
        log_area.code("\n".join(st.session_state.log), language="text")

    except Exception as e:
        st.error("Run failed. See log below.")
        log(f"❌ Fatal error: {e}\n{traceback.format_exc()}")
        log_area.code("\n".join(st.session_state.log), language="text")
        return

    # Download options
    if outs.final_dir.exists():
        final_paths = sorted(outs.final_dir.glob("*.pdf"))
        if final_paths:
            # Build unified InvestorReports.pdf
            master_pdf = outs.final_dir / "InvestorReports.pdf"
            merge_all_into_one(final_paths, master_pdf)

            zip_path = outs.root / "FinalInvestorReports.zip"
            zip_files(final_paths, zip_path)
            st.success(f"Done. Built {len(final_paths)} full report(s).")
            st.download_button(
                "Download ALL Final Reports (ZIP)",
                data=open(zip_path, "rb").read(),
                file_name=zip_path.name,
                mime="application/zip",
                use_container_width=True,
            )
            st.download_button(
                "Download InvestorReports.pdf (all‑in‑one)",
                data=open(master_pdf, "rb").read(),
                file_name=master_pdf.name,
                mime="application/pdf",
                use_container_width=True,
            )
            st.write("Or download individual PDFs:")
            for p in final_paths:
                st.download_button(
                    f"Download {p.name}",
                    data=open(p, "rb").read(),
                    file_name=p.name,
                    mime="application/pdf",
                )

    # Persist log view
    log_area.code("\n".join(st.session_state.log), language="text")


if __name__ == "__main__":
    main()
