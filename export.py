from __future__ import annotations

import io
from datetime import date
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
import pandas as pd

# ── Petrol palette  (Analysis sheet) ─────────────────────────────────────────
PT_TITLE  = "0A1F28"   # very dark petrol  – title bar
PT_HDR    = "1B5566"   # petrol            – header row
PT_ALT    = "EBF5F7"   # ice blue          – alternating rows
PT_SUM    = "A8CDD6"   # medium petrol     – sum/key columns
PT_PAY    = "FFF8E1"   # warm cream        – pay columns (Gross salary, Hourly rate)

# ── Orange palette  (Breakdown sheet) ────────────────────────────────────────
OR_TITLE  = "9E3515"   # dark burnt orange – title bar
OR_HDR    = "E8622A"   # M3 orange         – header row
OR_ALT    = "FEF2EC"   # pale peach        – alternating rows
OR_OT     = "FAD4BC"   # soft peach        – overtime columns
OR_GROSS  = "1B5566"   # petrol accent     – total gross pay column

RED_FONT  = "C00000"   # red for negative overtime

SUM_COLS  = {"Total travel time", "Total working time", "Target hours", "Overtime"}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _hdr_border(color: str) -> Border:
    s = Side(style="medium", color=color)
    return Border(bottom=s)

def _thin_border() -> Border:
    return Border(
        bottom=Side(style="thin", color="D8D8D8"),
        right =Side(style="thin", color="E8E8E8"),
    )


# ── Analysis sheet  (Petrol) ─────────────────────────────────────────────────

def _fill_analysis_sheet(
    ws,
    df: pd.DataFrame,
    meta: dict[str, Any],
    holiday_label: str,
) -> None:
    """8-category detail sheet — petrol colour scheme."""
    d_from: date | None = meta.get("date_from")
    d_to:   date | None = meta.get("date_to")
    soll: float         = meta.get("soll_hours", 0.0)

    period_str = (
        f"{d_from.strftime('%d.%m.%Y')} – {d_to.strftime('%d.%m.%Y')}"
        if d_from and d_to else "–"
    )
    title_text = (
        f"M3 Croatia  ·  Time & Travel Analysis  ·  "
        f"{period_str}  ·  Target {soll:.1f} h  ·  Holidays: {holiday_label}"
    )

    n_cols = len(df.columns)

    # ── Title row ──────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value=title_text)
    tc.font      = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    tc.fill      = PatternFill("solid", fgColor=PT_TITLE)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # ── Header row ─────────────────────────────────────────────────────────────
    hdr_fill  = PatternFill("solid", fgColor=PT_HDR)
    hdr_font  = Font(bold=True, color="FFFFFF", size=9, name="Calibri")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    hdr_bdr   = _hdr_border("0D3D4D")
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=2, column=c_idx, value=col_name)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = hdr_bdr
    ws.row_dimensions[2].height = 46

    # ── Data rows ──────────────────────────────────────────────────────────────
    alt_fill = PatternFill("solid", fgColor=PT_ALT)
    thin_bdr = _thin_border()
    num_fmt  = '#,##0.00" h"'
    eur_fmt  = '#,##0.00" €"'
    eur4_fmt = '#,##0.0000" €"'

    for r_idx, row_data in enumerate(df.itertuples(index=False), start=3):
        use_alt = (r_idx % 2 == 0)
        ws.row_dimensions[r_idx].height = 18
        for c_idx, (col_name, val) in enumerate(zip(df.columns, row_data), start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_bdr
            cell.fill   = alt_fill if use_alt else PatternFill()

            is_neg = isinstance(val, (int, float)) and val < 0

            if col_name == "Employee":
                cell.font      = Font(bold=True, size=10, color=PT_TITLE, name="Calibri")
                cell.alignment = Alignment(vertical="center", indent=1)
            elif col_name == "Hourly rate (€)":
                cell.number_format = eur4_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10, name="Calibri")
            elif col_name in {"Gross salary (€)"}:
                cell.number_format = eur_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10, name="Calibri")
            elif col_name == "Overtime" and is_neg:
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10, name="Calibri", color=RED_FONT)
            else:
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10, name="Calibri")

    # ── Column widths ──────────────────────────────────────────────────────────
    for c_idx, col_name in enumerate(df.columns, start=1):
        if col_name == "Employee":
            ws.column_dimensions[get_column_letter(c_idx)].width = 26
        elif col_name in PAY_COLS | {"Hourly rate (€)"}:
            ws.column_dimensions[get_column_letter(c_idx)].width = 18
        else:
            ws.column_dimensions[get_column_letter(c_idx)].width = 15

    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"
    ws.sheet_view.showGridLines = False


# ── Breakdown sheet  (Orange) ─────────────────────────────────────────────────

def _breakdown_sheet_content(
    ws,
    headers: list[str],
    rows_iter,          # iterable of (r_idx, values_list, aw_row_or_none)
    meta: dict[str, Any],
    holiday_label: str,
    has_pay_cols: bool,
) -> None:
    """Shared rendering for both full-report and standalone Breakdown sheets."""
    d_from: date | None = meta.get("date_from")
    d_to:   date | None = meta.get("date_to")

    period_str = (
        f"{d_from.strftime('%d.%m.%Y')} – {d_to.strftime('%d.%m.%Y')}"
        if d_from and d_to else "–"
    )
    title_text = (
        f"M3 Croatia  ·  Breakdown Summary  ·  "
        f"{period_str}  ·  Holidays: {holiday_label}"
    )

    n_cols = len(headers)

    # ── Title row ──────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value=title_text)
    tc.font      = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    tc.fill      = PatternFill("solid", fgColor=OR_TITLE)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # ── Header row ─────────────────────────────────────────────────────────────
    hdr_fill  = PatternFill("solid", fgColor=OR_HDR)
    hdr_font  = Font(bold=True, color="FFFFFF", size=9, name="Calibri")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    hdr_bdr   = _hdr_border("A83F18")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c_idx, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = hdr_align; cell.border = hdr_bdr
    ws.row_dimensions[2].height = 46

    # ── Data rows ──────────────────────────────────────────────────────────────
    alt_fill = PatternFill("solid", fgColor=OR_ALT)
    thin_bdr = _thin_border()
    num_fmt  = '#,##0.00" h"'
    eur_fmt  = '#,##0.00" €"'

    OT_COLS = {"Overtime total", "Overtime travel (110%)", "Overtime work (110%)"}

    for r_idx, values in rows_iter:
        use_alt = (r_idx % 2 == 0)
        ws.row_dimensions[r_idx].height = 18
        for c_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_bdr
            cell.fill   = alt_fill if use_alt else PatternFill()
            h = headers[c_idx - 1]

            is_neg = isinstance(val, (int, float)) and val < 0

            if h == "Employee":
                cell.font      = Font(bold=True, size=10, color=OR_TITLE, name="Calibri")
                cell.alignment = Alignment(vertical="center", indent=1)
            elif h == "Total gross pay (€)":
                cell.number_format = eur_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(bold=True, size=10, name="Calibri")
            elif h in OT_COLS and is_neg:
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10, name="Calibri", color=RED_FONT)
            else:
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10, name="Calibri")

    # ── Column widths ──────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 26
    for c_idx, h in enumerate(headers[1:], start=2):
        ws.column_dimensions[get_column_letter(c_idx)].width = (
            22 if h == "Total gross pay (€)" else 20
        )

    ws.freeze_panes = "B3"
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"
    ws.sheet_view.showGridLines = False


def _build_analyse_sheet(
    wb: openpyxl.Workbook,
    df: pd.DataFrame,
    soll: float = 0.0,
    meta: dict[str, Any] | None = None,
    holiday_label: str = "",
) -> None:
    """Breakdown sheet with cross-sheet formulas referencing the Analysis sheet."""
    ws = wb.create_sheet("Breakdown")

    _col_map      = {name: get_column_letter(i + 1) for i, name in enumerate(df.columns)}
    _gross_letter = _col_map.get("Gross salary (€)")
    _hourly_letter= _col_map.get("Hourly rate (€)")
    _has_pay_cols = bool(_gross_letter and _hourly_letter)

    headers = [
        "Employee",
        "Travel + Working time",
        "Target hours",
        "Overtime total",
        "Overtime travel (110%)",
        "Overtime work (110%)",
        "Nights & Sun/holidays (75%)",
        "Nights (25%)",
        "Sun/holidays (50%)",
    ]
    if _has_pay_cols:
        headers.append("Total gross pay (€)")

    def _rows():
        for r_idx, (_, row) in enumerate(df.iterrows(), start=3):
            aw_row   = r_idx + 1   # Analysis sheet: title row 1, header row 2 → data from row 3
            emp_soll = row["Target hours"] if "Target hours" in df.columns else soll

            values = [
                row["Employee"],
                f"=Analysis!F{aw_row}+Analysis!K{aw_row}",
                emp_soll,
                f"=B{r_idx}-C{r_idx}",
                f"=IF(D{r_idx}<=Analysis!F{aw_row},D{r_idx},Analysis!F{aw_row})",
                f"=D{r_idx}-E{r_idx}",
                f"=Analysis!G{aw_row}+Analysis!B{aw_row}",
                f"=Analysis!H{aw_row}+Analysis!C{aw_row}",
                f"=Analysis!I{aw_row}+Analysis!D{aw_row}",
            ]
            if _has_pay_cols:
                values.append(
                    f"=Analysis!{_gross_letter}{aw_row}"
                    f"+((E{r_idx}*1.1)+(F{r_idx}*1.1)+(G{r_idx}*0.75)"
                    f"+(H{r_idx}*0.25)+(I{r_idx}*0.5))"
                    f"*Analysis!{_hourly_letter}{aw_row}"
                )
            yield r_idx, values

    _breakdown_sheet_content(ws, headers, _rows(), meta or {}, holiday_label, _has_pay_cols)


def _fill_breakdown_standalone_sheet(
    ws,
    df_analyse: pd.DataFrame,
    meta: dict[str, Any],
    holiday_label: str,
) -> None:
    """Standalone Breakdown sheet with precomputed values (no cross-sheet formulas)."""
    headers = list(df_analyse.columns)
    _has_pay = "Total gross pay (€)" in headers

    def _rows():
        for r_idx, (_, row) in enumerate(df_analyse.iterrows(), start=3):
            yield r_idx, [row[h] for h in headers]

    _breakdown_sheet_content(ws, headers, _rows(), meta, holiday_label, _has_pay)


# ── Public builders ───────────────────────────────────────────────────────────

def build_breakdown_standalone_bytes(
    df_analyse: pd.DataFrame,
    meta: dict[str, Any],
    holiday_label: str,
) -> bytes:
    """Standalone Breakdown sheet — for tax office."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Breakdown"
    _fill_breakdown_standalone_sheet(ws, df_analyse, meta, holiday_label)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_excel_bytes(
    df: pd.DataFrame,
    meta: dict[str, Any],
    holiday_label: str,
) -> bytes:
    """Full report: Analysis (petrol) + Breakdown (orange), opens on Breakdown."""
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Analysis"
    _fill_analysis_sheet(ws, df, meta, holiday_label)

    _build_analyse_sheet(wb, df, meta.get("soll_hours", 0.0), meta, holiday_label)

    wb.active = wb["Breakdown"]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
