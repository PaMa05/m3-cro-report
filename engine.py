from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time, timedelta, date
from pathlib import Path
from typing import List

CATEGORIES = [
    "Reisezeit nachts an Sonn- oder Feiertagen",
    "Reisezeit nachts",
    "Reisezeit an Sonn- oder Feiertagen",
    "sonstige Reisezeit",
    "Arbeitszeit nachts an Sonn- oder Feiertagen",
    "Arbeitszeit nachts",
    "Arbeitszeit an Sonn- oder Feiertagen",
    "sonstige Arbeitszeit",
]

CAT = {
    "RT_NIGHT_SUNHOL": CATEGORIES[0],  # Reisezeit nachts an Sonn- oder Feiertagen
    "RT_NIGHT": CATEGORIES[1],
    "RT_SUNHOL": CATEGORIES[2],
    "RT_OTHER": CATEGORIES[3],
    "WK_NIGHT_SUNHOL": CATEGORIES[4],
    "WK_NIGHT": CATEGORIES[5],
    "WK_SUNHOL": CATEGORIES[6],
    "WK_OTHER": CATEGORIES[7],
}


@dataclass(frozen=True)
class Segment:
    employee: str
    start: datetime
    end: datetime
    is_travel: bool


def _pause_to_timedelta(x) -> timedelta:
    import pandas as pd

    if x is None or pd.isna(x):
        return timedelta(0)

    if isinstance(x, (int, float)) and not pd.isna(x):
        v = float(x)
        if v >= 1.0:
            return timedelta(minutes=v)
        return timedelta(days=v)

    try:
        td = pd.to_timedelta(x, errors="coerce")
        if pd.isna(td):
            s = str(x).strip().replace(",", ".")
            try:
                return timedelta(minutes=float(s))
            except Exception:
                return timedelta(0)
        return td.to_pytimedelta()
    except Exception:
        return timedelta(0)


def _parse_excel(path: Path) -> List[Segment]:
    import pandas as pd

    df = pd.read_excel(path, engine="openpyxl")

    required = [
        "Vorname (bürgerlich)",
        "Nachname (bürgerlich)",
        "Startzeit der Anwesenheit",
        "Endzeit der Anwesenheit",
        "Anwesenheitsprojekt",
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Spalten fehlen:\n- " + "\n- ".join(missing))

    emp = (
        df["Vorname (bürgerlich)"].astype(str).str.strip()
        + " "
        + df["Nachname (bürgerlich)"].astype(str).str.strip()
    )

    start = pd.to_datetime(df["Startzeit der Anwesenheit"], errors="coerce")
    end = pd.to_datetime(df["Endzeit der Anwesenheit"], errors="coerce")

    project = df["Anwesenheitsprojekt"].astype(str).str.strip().str.lower()
    is_travel = project.eq("travel time")

    pause_col = "Erfasste Pausen zur Anwesenheit"
    pauses_raw = df[pause_col] if pause_col in df.columns else None

    segs: List[Segment] = []
    for idx, (e, s, t, tr) in enumerate(zip(emp, start, end, is_travel)):
        if pd.isna(s) or pd.isna(t):
            continue

        if pauses_raw is not None:
            pause_td = _pause_to_timedelta(pauses_raw.iloc[idx])
            if pause_td.total_seconds() > 0:
                continue

        sdt = s.to_pydatetime()
        edt = t.to_pydatetime()

        if edt <= sdt:
            continue

        segs.append(Segment(str(e), sdt, edt, bool(tr)))

    return segs


def _overlap(a0, a1, b0, b1) -> float:
    lo = max(a0, b0)
    hi = min(a1, b1)
    if hi <= lo:
        return 0.0
    return (hi - lo).total_seconds() / 3600.0


def _night_key(dt: datetime) -> date:
    if dt.time() >= time(22, 0):
        return dt.date()
    if dt.time() < time(6, 0):
        return dt.date() - timedelta(days=1)
    return dt.date()


def _is_sun_or_holiday(d: date, holiday_set: set[date]) -> bool:
    return d.weekday() == 6 or d in holiday_set


def _build_holidays(d0: date, d1: date, holiday_mode: str) -> set[date]:
    import holidays as holidays_lib

    years = list(range(d0.year, d1.year + 1))
    if not years:
        return set()

    if holiday_mode == "DE-NW":
        cal = holidays_lib.Germany(prov="NW", years=years)
    elif holiday_mode == "HR":
        cal = holidays_lib.Croatia(years=years)
    else:
        raise ValueError("Unbekannter Feiertagsmodus")

    return {d for d in cal.keys() if d0 <= d <= d1}


def _calc_soll_hours(d0: date, d1: date, holiday_set: set[date]) -> float:
    workdays = 0
    cur = d0
    while cur <= d1:
        is_weekday = cur.weekday() < 5
        is_holiday = cur in holiday_set
        if is_weekday and not is_holiday:
            workdays += 1
        cur += timedelta(days=1)
    return float(workdays) * 8.0


def _read_extra_cols(path: Path) -> tuple[dict, list[str]]:
    """
    Reads optional columns 'Effektive Arbeitsstunden' and 'Bruttolohn' from the
    Excel file.  Returns:
      extra  – dict  employee → {"eff_arb": float | None, "bruttolohn": float | None}
      warns  – list of warning strings for inconsistent values per employee
    """
    import pandas as pd

    df = pd.read_excel(path, engine="openpyxl")

    # Build employee key (same logic as _parse_excel)
    if "Vorname (bürgerlich)" not in df.columns or "Nachname (bürgerlich)" not in df.columns:
        return {}, []

    emp_series = (
        df["Vorname (bürgerlich)"].astype(str).str.strip()
        + " "
        + df["Nachname (bürgerlich)"].astype(str).str.strip()
    )

    extra: dict[str, dict] = {}
    warns: list[str] = []

    for col_name, agg_fn, key in [
        ("Effektive Arbeitsstunden", "max", "eff_arb"),
        ("Bruttolohn",              "min", "bruttolohn"),
    ]:
        if col_name not in df.columns:
            continue

        tmp = pd.DataFrame({"emp": emp_series, "val": pd.to_numeric(df[col_name], errors="coerce")})
        tmp = tmp.dropna(subset=["val"])
        if tmp.empty:
            continue

        for emp, grp in tmp.groupby("emp"):
            unique_vals = grp["val"].unique()
            if len(unique_vals) > 1:
                warns.append(
                    f"⚠️ '{col_name}' has multiple values for {emp} "
                    f"({', '.join(str(v) for v in sorted(unique_vals))}) — "
                    f"using {'highest' if agg_fn == 'max' else 'lowest'} value."
                )
            chosen = float(grp["val"].max() if agg_fn == "max" else grp["val"].min())
            extra.setdefault(str(emp), {})[key] = chosen

    return extra, warns


def evaluate_excel(path: Path, holiday_mode: str = "DE-NW"):
    segs = _parse_excel(path)
    if not segs:
        raise ValueError("Keine verwertbaren Zeilen gefunden.")

    min_dt = min(s.start for s in segs)
    max_dt = max(s.end for s in segs)

    holiday_set = _build_holidays(min_dt.date(), max_dt.date(), holiday_mode)

    night_hours: dict[tuple[str, date], float] = {}

    for s in segs:
        d = s.start.date() - timedelta(days=1)
        while d <= s.end.date():
            n0 = datetime.combine(d, time(22, 0))
            n1 = datetime.combine(d + timedelta(days=1), time(6, 0))
            h = _overlap(s.start, s.end, n0, n1)
            if h > 0:
                key = (s.employee, d)
                night_hours[key] = night_hours.get(key, 0.0) + h
            d += timedelta(days=1)

    night_valid = {k: (h >= 2.0) for k, h in night_hours.items()}

    slice_len = timedelta(minutes=15)

    travel = []
    work = []

    for s in segs:
        t = s.start
        while t < s.end:
            t2 = min(s.end, t + slice_len)
            nk = _night_key(t)
            is_night = (
                ((t.time() >= time(22)) or (t.time() < time(6)))
                and night_valid.get((s.employee, nk), False)
            )
            is_sunhol = _is_sun_or_holiday(t.date(), holiday_set)
            rec = (s.employee, t, t2, is_night, is_sunhol)
            (travel if s.is_travel else work).append(rec)
            t = t2

    out: dict[str, dict[str, float]] = {}

    def add(emp: str, cat: str, h: float):
        out.setdefault(emp, {})
        out[emp][cat] = out[emp].get(cat, 0.0) + h

    for emp, a, b, n, sh in travel:
        h = (b - a).total_seconds() / 3600.0
        if n and sh:
            add(emp, CAT["RT_NIGHT_SUNHOL"], h)
        elif n:
            add(emp, CAT["RT_NIGHT"], h)
        elif sh:
            add(emp, CAT["RT_SUNHOL"], h)
        else:
            add(emp, CAT["RT_OTHER"], h)

    for emp, a, b, n, sh in work:
        h = (b - a).total_seconds() / 3600.0
        if n and sh:
            add(emp, CAT["WK_NIGHT_SUNHOL"], h)
        elif n:
            add(emp, CAT["WK_NIGHT"], h)
        elif sh:
            add(emp, CAT["WK_SUNHOL"], h)
        else:
            add(emp, CAT["WK_OTHER"], h)

    for emp in out:
        for cat in CATEGORIES:
            out[emp].setdefault(cat, 0.0)

    # ── Extra columns (optional) ──────────────────────────────────────────────
    extra, extra_warns = _read_extra_cols(path)

    warnings_list: list[str] = extra_warns

    for emp in out:
        ed = extra.get(emp, {})
        eff_arb   = ed.get("eff_arb")
        bruttolohn = ed.get("bruttolohn")

        out[emp]["__eff_arb__"]    = eff_arb    # float or None
        out[emp]["__bruttolohn__"] = bruttolohn  # float or None
        if eff_arb and eff_arb > 0 and bruttolohn is not None:
            out[emp]["__stundenlohn__"] = round(bruttolohn / eff_arb, 4)
        else:
            out[emp]["__stundenlohn__"] = None

    soll = _calc_soll_hours(min_dt.date(), max_dt.date(), holiday_set)

    meta = {
        "date_from": min_dt.date(),
        "date_to": max_dt.date(),
        "soll_hours": soll,
    }

    warning_text = "\n".join(warnings_list) if warnings_list else None
    return out, warning_text, meta
