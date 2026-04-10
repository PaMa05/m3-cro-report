"""
M3 Croatia Time & Travel Report – Streamlit Web-App
====================================================
Starten mit:   streamlit run app.py
"""

from __future__ import annotations

import tempfile
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ── Engine-interne Namen (Deutsch) ────────────────────────────────────────────
_REISEZEIT_DE = [
    "Reisezeit nachts an Sonn- oder Feiertagen",
    "Reisezeit nachts",
    "Reisezeit an Sonn- oder Feiertagen",
    "sonstige Reisezeit",
]
_ARBEITSZEIT_DE = [
    "Arbeitszeit nachts an Sonn- oder Feiertagen",
    "Arbeitszeit nachts",
    "Arbeitszeit an Sonn- oder Feiertagen",
    "sonstige Arbeitszeit",
]

# ── Englische Spaltennamen (Anzeige + Export) ─────────────────────────────────
REISEZEIT_CATS = [
    "Travel: nights & Sun/holidays",
    "Travel: nights",
    "Travel: Sun/holidays",
    "Travel: other",
]
ARBEITSZEIT_CATS = [
    "Work: nights & Sun/holidays",
    "Work: nights",
    "Work: Sun/holidays",
    "Work: other",
]
SUM_REISE  = "Total travel time"
SUM_ARBEIT = "Total working time"
SOLL_COL   = "Target hours"
UEBER_COL  = "Overtime"

HOLIDAY_OPTIONS = {
    "Croatia":          "HR",
    "NRW (Germany)":    "DE-NW",
}

# ── Seite einrichten ──────────────────────────────────────────────────────────
icon_path = Path(__file__).parent / "assets" / "icon.png"

st.set_page_config(
    page_title="CRO Report",
    page_icon=str(icon_path) if icon_path.exists() else "⏱️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Brand CSS ─────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
        /* ── Layout ── */
        .block-container { padding-top: 0 !important; max-width: 1400px; }
        [data-testid="stAppViewContainer"] { background: #FFFFFF; }

        /* ── Hero-Header ── */
        .m3-header {
            background: linear-gradient(135deg, #0D2B35 0%, #1B5566 100%);
            border-radius: 0 0 16px 16px;
            padding: 28px 36px 24px 36px;
            margin-bottom: 28px;
            display: flex;
            align-items: center;
            gap: 20px;
        }
        .m3-header img {
            height: 110px;
            width: auto;
            object-fit: contain;
            filter: invert(1) brightness(2);
        }
        .m3-header-text h1 {
            color: #FFFFFF;
            font-size: 1.6rem;
            font-weight: 700;
            margin: 0;
            letter-spacing: -0.3px;
        }
        .m3-header-text p {
            color: #E8622A;
            font-size: 0.85rem;
            font-weight: 600;
            margin: 2px 0 0 0;
            letter-spacing: 1.5px;
            text-transform: uppercase;
        }

        /* ── Metric cards ── */
        [data-testid="stMetric"] {
            background: #F2F6F7;
            border-left: 4px solid #E8622A;
            border-radius: 10px;
            padding: 14px 18px !important;
        }
        [data-testid="stMetricLabel"] { color: #444444 !important; font-weight: 600; }
        [data-testid="stMetricValue"] { color: #0D2B35 !important; font-size: 1.25rem !important; }

        /* ── Tabs ── */
        [data-baseweb="tab-list"] { border-bottom: 2px solid #E8E8E8; gap: 4px; }
        [data-baseweb="tab"] {
            font-weight: 600;
            color: #333333 !important;
            border-radius: 8px 8px 0 0 !important;
            padding: 10px 20px !important;
        }
        [aria-selected="true"][data-baseweb="tab"] {
            background: #FFF4F0 !important;
            color: #E8622A !important;
            border-bottom: 3px solid #E8622A !important;
        }

        /* ── Download button ── */
        [data-testid="stDownloadButton"] > button {
            background: #E8622A !important;
            color: #FFFFFF !important;
            border: none !important;
            border-radius: 8px !important;
            font-weight: 700 !important;
            padding: 10px 24px !important;
            font-size: 0.95rem !important;
            box-shadow: 0 2px 8px rgba(232,98,42,0.3);
            transition: all 0.2s;
        }
        [data-testid="stDownloadButton"] > button:hover {
            background: #C94F1E !important;
            box-shadow: 0 4px 14px rgba(232,98,42,0.4);
            transform: translateY(-1px);
        }

        /* ── File uploader ── */
        [data-testid="stFileUploader"] {
            border: 2px dashed #1B5566 !important;
            border-radius: 10px;
            padding: 8px;
            background: #F2F6F7;
        }

        /* ── Selectbox ── */
        [data-testid="stSelectbox"] label { color: #333333; font-weight: 600; }

        /* ── Table ── */
        [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; box-shadow: 0 1px 6px rgba(0,0,0,0.08); }

        /* ── Divider ── */
        hr { border-color: #E8622A !important; opacity: 0.25; }

        /* ── Info box ── */
        [data-testid="stAlert"] { border-radius: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Header ────────────────────────────────────────────────────────────────────
import base64

def _img_to_b64(path: Path) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

logo_b64 = _img_to_b64(icon_path) if icon_path.exists() else ""
logo_html = f'<img src="data:image/png;base64,{logo_b64}" />' if logo_b64 else ""

st.markdown(
    f"""
    <div class="m3-header">
        {logo_html}
        <div class="m3-header-text">
            <h1>Croatia Time &amp; Travel Report</h1>
            <p>Time &amp; Travel Analysis</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Password protection ───────────────────────────────────────────────────────
def _check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    st.markdown(
        """
        <div style="max-width:400px;margin:80px auto;text-align:center;">
            <p style="font-size:2.5rem;margin-bottom:4px;">🔒</p>
            <p style="color:#1B5566;font-weight:700;font-size:1.1rem;margin-bottom:24px;">
                Enter password to access the report
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col = st.columns([1, 2, 1])[1]
    with col:
        pw = st.text_input("Password", type="password", label_visibility="collapsed",
                           placeholder="Password …")
        if pw:
            if pw == "LieseDoesntWork":
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Wrong password.")
    st.stop()

_check_password()

# ── Help Section ─────────────────────────────────────────────────────────────
with st.expander("ℹ️  How to use this tool", expanded=False):
    st.markdown("""
**What this tool does**

Analyzes employee attendance data exported from your HR system and classifies all hours
into 8 categories — split by activity type (work vs. travel) and time conditions
(night hours, Sundays, and public holidays). Results are shown in two views and can be
downloaded as a formatted Excel report.

---

**Input requirements**

- File format: `.xlsx` or `.xlsm`
- The file must contain the following columns (exact names as exported from the HR system):
  `Vorname (bürgerlich)`, `Nachname (bürgerlich)`, `Startzeit der Anwesenheit`,
  `Endzeit der Anwesenheit`, `Anwesenheitsprojekt`
- Rows with a recorded break (`Erfasste Pausen zur Anwesenheit`) are automatically excluded
- Rows where end time ≤ start time are skipped

---

**Classification rules**

| Condition | Rule |
|---|---|
| **Travel vs. Work** | Project = `travel time` → Travel; everything else → Working time |
| **Night hours** | 22:00 – 06:00; only applied if the employee has ≥ 2 h in that night window |
| **Night attribution** | A night starting 22:00 on day X is attributed to day X |
| **Sun/holidays** | Sundays always count; public holidays selectable (Croatia or NRW) |
| **Calculation** | Hours are sliced into 15-minute intervals for precise classification |
| **Target hours** | Working days in the period (Mon–Fri, excl. holidays) × 8 h |

The 8 output categories are the combinations of Travel/Work × Night+Sun/holiday / Night only / Sun/holiday only / Other.

---

**How to use**

1. Select the applicable public holiday calendar (Croatia or NRW Germany)
2. Upload your Excel export file
3. Review results in the **Breakdown** tab (all 8 categories per employee) and the **Analysis** tab (totals, overtime split)
4. Click **Download result as Excel** to get a formatted two-sheet report
""")

# ── Controls ──────────────────────────────────────────────────────────────────
col_upload, col_holiday = st.columns([3, 1])

with col_upload:
    uploaded = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xlsm"],
        help="File must contain columns: first name, last name, start time, end time, project.",
    )

with col_holiday:
    holiday_label = st.selectbox(
        "Public holidays",
        list(HOLIDAY_OPTIONS.keys()),
        index=0,
    )
    holiday_mode = HOLIDAY_OPTIONS[holiday_label]

# ── Platzhalter wenn keine Datei ──────────────────────────────────────────────
if uploaded is None:
    st.info("👆 Please upload an Excel file to start the evaluation.")
    st.stop()


# ── Berechnung (gecacht nach Dateiinhalt + Feiertagsmodus) ────────────────────
@st.cache_data(show_spinner=False)
def run_eval(file_bytes: bytes, holiday_mode: str):
    """Evaluates the Excel; result is cached as long as file + mode are the same."""
    from engine import evaluate_excel

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = Path(tmp.name)

    return evaluate_excel(tmp_path, holiday_mode=holiday_mode)


file_bytes = uploaded.read()

with st.spinner("Calculating times …"):
    try:
        result, warning, meta = run_eval(file_bytes, holiday_mode)
    except Exception as e:
        st.error(f"**Error during evaluation:** {e}")
        st.stop()

if warning:
    st.warning(warning)

# ── Ergebnis-DataFrame aufbauen ───────────────────────────────────────────────
soll      = meta.get("soll_hours", 0.0)
employees = sorted(result.keys())

rows = []
for emp in employees:
    d = result[emp]
    row: dict = {"Employee": emp}

    for de, en in zip(_REISEZEIT_DE, REISEZEIT_CATS):
        row[en] = round(d.get(de, 0.0), 2)
    row[SUM_REISE] = round(sum(d.get(de, 0.0) for de in _REISEZEIT_DE), 2)

    for de, en in zip(_ARBEITSZEIT_DE, ARBEITSZEIT_CATS):
        row[en] = round(d.get(de, 0.0), 2)
    row[SUM_ARBEIT] = round(sum(d.get(de, 0.0) for de in _ARBEITSZEIT_DE), 2)

    # Use per-employee Effektive Arbeitsstunden if available, else fall back to period soll
    eff_arb = d.get("__eff_arb__")
    emp_soll = round(eff_arb, 2) if eff_arb else round(soll, 2)
    row[SOLL_COL]  = emp_soll
    row[UEBER_COL] = round(row[SUM_ARBEIT] - emp_soll, 2)

    # Optional pay columns
    bruttolohn  = d.get("__bruttolohn__")
    stundenlohn = d.get("__stundenlohn__")
    row["Bruttolohn (€)"]  = round(bruttolohn,  2) if bruttolohn  is not None else None
    row["Stundenlohn (€)"] = round(stundenlohn, 4) if stundenlohn is not None else None

    rows.append(row)

df = pd.DataFrame(rows)

# Detect whether pay columns have any data (drop them if all None)
_has_pay = df["Bruttolohn (€)"].notna().any()
if not _has_pay:
    df = df.drop(columns=["Bruttolohn (€)", "Stundenlohn (€)"])

# ── Kennzahlen-Zeile ──────────────────────────────────────────────────────────
d_from = meta.get("date_from")
d_to   = meta.get("date_to")

m1, m2, m3, m4 = st.columns(4)

with m1:
    zeitraum = (
        f"{d_from.strftime('%d.%m.%Y')} – {d_to.strftime('%d.%m.%Y')}"
        if d_from and d_to else "–"
    )
    st.metric("Period", zeitraum)

with m2:
    st.metric("Employees", len(employees))

with m3:
    st.metric("Target hours (period)", f"{soll:.1f} h")

with m4:
    avg_ueber = df[UEBER_COL].mean() if not df.empty else 0.0
    st.metric("Avg. overtime", f"{avg_ueber:+.1f} h")

st.divider()

# ── Analyse-DataFrame aufbauen ────────────────────────────────────────────────
gesamt     = (df[SUM_REISE] + df[SUM_ARBEIT]).round(2)
mehrarbeit = (gesamt - df[SOLL_COL]).round(2)   # per-employee soll
meh_reise  = mehrarbeit.combine(df[SUM_REISE], min).round(2)
meh_arbeit = (mehrarbeit - meh_reise).round(2)

_analyse_data: dict = {
    "Employee":                    df["Employee"],
    "Travel + Working time":       gesamt,
    "Target hours":                df[SOLL_COL],
    "Overtime total":              mehrarbeit,
    "Overtime travel (110%)":      meh_reise,
    "Overtime work (110%)":        meh_arbeit,
    "Nights & Sun/holidays (75%)": (df[REISEZEIT_CATS[0]] + df[ARBEITSZEIT_CATS[0]]).round(2),
    "Nights (25%)":                (df[REISEZEIT_CATS[1]] + df[ARBEITSZEIT_CATS[1]]).round(2),
    "Sun/holidays (50%)":          (df[REISEZEIT_CATS[2]] + df[ARBEITSZEIT_CATS[2]]).round(2),
}
if _has_pay:
    _analyse_data["Bruttolohn (€)"]  = df["Bruttolohn (€)"]
    _analyse_data["Stundenlohn (€)"] = df["Stundenlohn (€)"]

df_analyse = pd.DataFrame(_analyse_data)

# ── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 Analysis", "📋 Breakdown"])

with tab1:
    num_cols = [c for c in df.columns if c != "Employee"]
    column_config = {
        col: st.column_config.NumberColumn(label=col, format="%.2f h", help=col)
        for col in num_cols
    }
    column_config[UEBER_COL] = st.column_config.NumberColumn(
        label=UEBER_COL,
        format="%.2f h",
        help="Total working time − Target hours",
    )
    if _has_pay:
        column_config["Bruttolohn (€)"]  = st.column_config.NumberColumn(
            label="Bruttolohn (€)", format="%.2f €"
        )
        column_config["Stundenlohn (€)"] = st.column_config.NumberColumn(
            label="Stundenlohn (€)", format="%.4f €"
        )
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config=column_config,
        height=min(600, 60 + len(df) * 36),
    )

with tab2:
    analyse_num_cols = [c for c in df_analyse.columns if c != "Employee"]
    analyse_col_config = {
        col: st.column_config.NumberColumn(label=col, format="%.2f h")
        for col in analyse_num_cols
    }
    for col in ("Overtime total", "Overtime travel (110%)", "Overtime work (110%)"):
        analyse_col_config[col] = st.column_config.NumberColumn(
            label=col,
            format="%.2f h",
            help="Positive = overtime, Negative = undertime",
        )
    if _has_pay:
        analyse_col_config["Bruttolohn (€)"]  = st.column_config.NumberColumn(
            label="Bruttolohn (€)", format="%.2f €"
        )
        analyse_col_config["Stundenlohn (€)"] = st.column_config.NumberColumn(
            label="Stundenlohn (€)", format="%.4f €"
        )
    st.dataframe(
        df_analyse,
        use_container_width=True,
        hide_index=True,
        column_config=analyse_col_config,
        height=min(600, 60 + len(df_analyse) * 36),
    )

# ── Export ────────────────────────────────────────────────────────────────────
st.divider()

@st.cache_data(show_spinner=False)
def make_excel(df_json: str, meta_serializable: dict, holiday_label: str) -> bytes:
    """Full report (Analysis + Breakdown sheets) — cached."""
    import io
    from export import build_excel_bytes
    df_export = pd.read_json(io.StringIO(df_json), orient="split")
    return build_excel_bytes(df_export, meta_serializable, holiday_label)


@st.cache_data(show_spinner=False)
def make_excel_tax(df_json: str, meta_serializable: dict, holiday_label: str) -> bytes:
    """Analysis-only sheet for tax office (Steuerbüro) — cached."""
    import io
    from export import build_breakdown_only_bytes
    df_export = pd.read_json(io.StringIO(df_json), orient="split")
    return build_breakdown_only_bytes(df_export, meta_serializable, holiday_label)


meta_for_cache = {
    "date_from":  d_from,
    "date_to":    d_to,
    "soll_hours": soll,
}

df_json = df.to_json(orient="split")

excel_bytes     = make_excel(df_json, meta_for_cache, holiday_label)
excel_tax_bytes = make_excel_tax(df_json, meta_for_cache, holiday_label)

timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")

dl_col1, dl_col2, _ = st.columns([2, 2, 3])

with dl_col1:
    st.download_button(
        label="⬇️  Full report (Analysis + Breakdown)",
        data=excel_bytes,
        file_name=f"report_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with dl_col2:
    st.download_button(
        label="⬇️  Tax office (Analysis only)",
        data=excel_tax_bytes,
        file_name=f"report_tax_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Email ─────────────────────────────────────────────────────────────────────
st.divider()

import urllib.parse

_FIXED_TO  = "franka.grandes@vlahov.com"
_FIXED_CC  = ["a.privara@m3connect.hr", "Hr@m3connect.de"]

period_str = (
    f"{d_from.strftime('%d.%m.%Y')} – {d_to.strftime('%d.%m.%Y')}"
    if d_from and d_to else "–"
)
email_subject = f"M3 Croatia Time & Travel Report – {period_str}"
email_body    = (
    f"Dear Franka,\n\n"
    f"please find attached the Croatia Time & Travel Analysis for the period {period_str}.\n\n"
    f"Best regards,\nM3 HR"
)

def _mailto_link(to: str, cc: list[str], subject: str, body: str) -> str:
    params = urllib.parse.urlencode(
        {"cc": ",".join(cc), "subject": subject, "body": body},
        quote_via=urllib.parse.quote,
    )
    return f"mailto:{urllib.parse.quote(to)}?{params}"

with st.expander("📧  Open in Outlook / Mail", expanded=False):
    st.caption(
        "Clicking the button opens your default mail client (Outlook) with "
        "recipients, subject and body pre-filled. "
        "Please attach the downloaded Excel file manually before sending."
    )

    # ── Custom address ────────────────────────────────────────────────────────
    st.markdown("**Send to a custom address**")
    custom_addr = st.text_input(
        "Email address", placeholder="you@example.com",
        label_visibility="collapsed", key="custom_email_addr",
    )
    if custom_addr:
        custom_link = _mailto_link(
            to=custom_addr, cc=[],
            subject=f"[TEST] {email_subject}",
            body=email_body,
        )
        st.markdown(
            f'<a href="{custom_link}" target="_blank">'
            f'<button style="background:#1B5566;color:#fff;border:none;border-radius:8px;'
            f'padding:9px 22px;font-weight:700;cursor:pointer;font-size:0.93rem;">'
            f'📨 Open draft for {custom_addr}</button></a>',
            unsafe_allow_html=True,
        )

    st.divider()

    # ── Fixed recipients ──────────────────────────────────────────────────────
    st.markdown(
        f"**Send to Steuerbüro**  \n"
        f"To: `{_FIXED_TO}`  •  CC: `{', '.join(_FIXED_CC)}`"
    )
    fixed_link = _mailto_link(
        to=_FIXED_TO, cc=_FIXED_CC,
        subject=email_subject,
        body=email_body,
    )
    st.markdown(
        f'<a href="{fixed_link}" target="_blank">'
        f'<button style="background:#E8622A;color:#fff;border:none;border-radius:8px;'
        f'padding:9px 22px;font-weight:700;cursor:pointer;font-size:0.93rem;">'
        f'📨 Open Outlook draft for Steuerbüro</button></a>',
        unsafe_allow_html=True,
    )
