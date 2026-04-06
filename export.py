from __future__ import annotations

import io
from datetime import date
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# ── Farben ────────────────────────────────────────────────────────────────────
M3_BLUE  = "1F4E79"   # dunkles Blau für Header
M3_LIGHT = "D6E4F0"   # helles Blau für alternierende Zeilen
GOLD     = "F4B942"   # Gold für Summen-/Überstunden-Spalten
RED_FONT = "C00000"   # Rot für negative Überstunden

SUM_COLS = {"Total travel time", "Total working time", "Target hours", "Overtime"}


def build_excel_bytes(
    df: pd.DataFrame,
    meta: dict[str, Any],
    holiday_label: str,
) -> bytes:
    """
    Erstellt eine formatierte Excel-Datei aus dem Auswertungs-DataFrame
    und gibt sie als Bytes zurück (für st.download_button).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Auswertung"

    d_from: date | None = meta.get("date_from")
    d_to:   date | None = meta.get("date_to")
    soll: float         = meta.get("soll_hours", 0.0)

    # ── Titelzeile ────────────────────────────────────────────────────────────
    title_parts = ["M3 Croatia Time & Travel Report"]
    if d_from and d_to:
        title_parts.append(
            f"{d_from.strftime('%d.%m.%Y')} – {d_to.strftime('%d.%m.%Y')}"
        )
    title_parts.append(f"Soll: {soll:.1f} h")
    title_parts.append(f"Feiertage: {holiday_label}")
    title_text = "   |   ".join(title_parts)

    n_cols = len(df.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    tc = ws.cell(row=1, column=1, value=title_text)
    tc.font      = Font(bold=True, color="FFFFFF", size=12)
    tc.fill      = PatternFill("solid", fgColor=M3_BLUE)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── Kopfzeile ─────────────────────────────────────────────────────────────
    hdr_fill  = PatternFill("solid", fgColor=M3_BLUE)
    hdr_font  = Font(bold=True, color="FFFFFF", size=10)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=2, column=c_idx, value=col_name)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = hdr_align
    ws.row_dimensions[2].height = 44

    # ── Datenzeilen ───────────────────────────────────────────────────────────
    alt_fill  = PatternFill("solid", fgColor=M3_LIGHT)
    sum_fill  = PatternFill("solid", fgColor=GOLD)
    thin_side = Side(style="thin", color="CCCCCC")
    thin_bdr  = Border(
        bottom=thin_side,
        right=Side(style="thin", color="E0E0E0"),
    )
    num_fmt = '#,##0.00" h"'

    for r_idx, row_data in enumerate(df.itertuples(index=False), start=3):
        use_alt = (r_idx % 2 == 0)
        for c_idx, (col_name, val) in enumerate(zip(df.columns, row_data), start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_bdr

            if col_name == "Employee":
                cell.font      = Font(bold=True, size=10)
                cell.alignment = Alignment(vertical="center")
                if use_alt:
                    cell.fill = alt_fill

            elif col_name in SUM_COLS:
                cell.fill          = sum_fill
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                is_neg = isinstance(val, (int, float)) and val < 0
                cell.font = Font(
                    bold=True,
                    size=10,
                    color=(RED_FONT if (col_name == "Overtime" and is_neg) else "000000"),
                )

            else:
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10)
                if use_alt:
                    cell.fill = alt_fill

    # ── Spaltenbreiten ────────────────────────────────────────────────────────
    for c_idx, col_name in enumerate(df.columns, start=1):
        col_letter = get_column_letter(c_idx)
        if col_name == "Employee":
            ws.column_dimensions[col_letter].width = 24
        else:
            ws.column_dimensions[col_letter].width = 15

    # ── Fixierte Kopfzeile ────────────────────────────────────────────────────
    ws.freeze_panes = "B3"

    # ── Autofilter ────────────────────────────────────────────────────────────
    ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"

    # ── Analyse-Sheet ─────────────────────────────────────────────────────────
    _build_analyse_sheet(wb, df, soll)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _build_analyse_sheet(wb: openpyxl.Workbook, df: pd.DataFrame, soll: float) -> None:
    """
    Zweites Sheet 'Analyse' mit aggregierten Werten:
    Gesamtstunden, Mehrarbeit (aufgeteilt in Reise/Arbeit),
    und kombinierte Nacht-/Feiertags-Stunden.
    """
    GREEN      = "375623"   # dunkelgrün Header
    GREEN_LIGHT = "E2EFDA"  # hellgrün alternierend
    ORANGE     = "FCE4D6"   # orange für Mehrarbeit-Spalten

    ws = wb.create_sheet("Analyse")

    # Englische Spaltennamen (müssen mit app.py übereinstimmen)
    COL_REISE_NIGHT_SH  = "Travel: nights & Sun/holidays"
    COL_REISE_NIGHT     = "Travel: nights"
    COL_REISE_SH        = "Travel: Sun/holidays"
    COL_ARBEIT_NIGHT_SH = "Work: nights & Sun/holidays"
    COL_ARBEIT_NIGHT    = "Work: nights"
    COL_ARBEIT_SH       = "Work: Sun/holidays"
    COL_SUM_REISE       = "Total travel time"
    COL_SUM_ARBEIT      = "Total working time"

    headers = [
        "Employee",
        "Travel + Working time",
        "Target hours",
        "Overtime total",
        "Overtime travel",
        "Overtime work",
        "Nights & Sun/holidays",
        "Nights",
        "Sun/holidays",
    ]

    # Styling
    hdr_fill  = PatternFill("solid", fgColor=GREEN)
    hdr_font  = Font(bold=True, color="FFFFFF", size=10)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alt_fill  = PatternFill("solid", fgColor=GREEN_LIGHT)
    ora_fill  = PatternFill("solid", fgColor=ORANGE)
    thin_side = Side(style="thin", color="CCCCCC")
    thin_bdr  = Border(bottom=thin_side, right=Side(style="thin", color="E0E0E0"))
    num_fmt   = '#,##0.00" h"'

    # Kopfzeile
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = hdr_align
    ws.row_dimensions[1].height = 44

    # Auswertungs-Sheet: Titel in Zeile 1, Header in Zeile 2, Daten ab Zeile 3.
    # Analyse-Sheet:     Header in Zeile 1, Daten ab Zeile 2.
    # → Analyse-Zeile r  entspricht Auswertungs-Zeile r+1
    #
    # Auswertungs-Spalten (nach build_excel_bytes):
    #   A=Mitarbeiter, B=RT_NIGHT_SH, C=RT_NIGHT, D=RT_SH, E=RT_OTHER,
    #   F=Summe Reisezeit, G=WK_NIGHT_SH, H=WK_NIGHT, I=WK_SH, J=WK_OTHER,
    #   K=Summe Arbeitszeit, L=Soll-Stunden, M=Überstunden

    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        use_alt   = (r_idx % 2 == 0)
        aw_row    = r_idx + 1   # entsprechende Zeile im Auswertungs-Sheet

        # Formeln – Excel speichert intern englische Namen; Semikolon für IF→WENN
        f_gesamt      = f"=Auswertung!F{aw_row}+Auswertung!K{aw_row}"
        f_mehrarbeit  = f"=B{r_idx}-C{r_idx}"
        f_meh_reise   = f"=IF(D{r_idx}<=Auswertung!F{aw_row},D{r_idx},Auswertung!F{aw_row})"
        f_meh_arbeit  = f"=D{r_idx}-E{r_idx}"
        f_night_sh    = f"=Auswertung!G{aw_row}+Auswertung!B{aw_row}"
        f_night       = f"=Auswertung!H{aw_row}+Auswertung!C{aw_row}"
        f_sun_hol     = f"=Auswertung!I{aw_row}+Auswertung!D{aw_row}"

        values = [
            row["Employee"],  # A – Wert, kein Formel nötig
            f_gesamt,            # B
            soll,                # C – fixer Wert (Werktage × 8h)
            f_mehrarbeit,        # D
            f_meh_reise,         # E
            f_meh_arbeit,        # F
            f_night_sh,          # G
            f_night,             # H
            f_sun_hol,           # I
        ]

        for c_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_bdr

            if c_idx == 1:  # Employee
                cell.font      = Font(bold=True, size=10)
                cell.alignment = Alignment(vertical="center")
                if use_alt:
                    cell.fill = alt_fill

            elif headers[c_idx - 1] in ("Overtime total", "Overtime travel", "Overtime work"):
                cell.fill          = ora_fill
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                is_neg = isinstance(val, (int, float)) and val < 0
                cell.font = Font(bold=True, size=10, color=(RED_FONT if is_neg else "000000"))

            else:
                cell.number_format = num_fmt
                cell.alignment     = Alignment(horizontal="right", vertical="center")
                cell.font          = Font(size=10)
                if use_alt:
                    cell.fill = alt_fill

    # Spaltenbreiten
    ws.column_dimensions["A"].width = 24
    for c_idx in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 20

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
