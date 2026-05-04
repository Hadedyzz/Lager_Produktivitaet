from datetime import datetime
from typing import List, Dict
# app.py (top of file) — add these lines BEFORE other imports that use matplotlib
import matplotlib
matplotlib.use("Agg")  # must come before importing matplotlib.pyplot anywhere
matplotlib.rcParams["font.family"] = "DejaVu Sans"   # available on Linux, supports umlauts
matplotlib.rcParams["figure.dpi"] = 96  # cloud-friendly default

import streamlit as st

# Your modules
from data_loader import load_excel
from aggregation import aggregate_weekly, aggregate_daily
from weekly_plots import plot_weekly_charts
from daily_plots import plot_daily_charts
from download_utils import render_download_section
from helpers import make_context_key
from config import SAEGEN_TARGET


# =========================
# Page settings
# =========================
st.set_page_config(page_title="Rollenbewegung Dashboard", layout="wide")
st.title("📊 Rollenbewegung Dashboard")


def _reset_selection_state_if_needed(context_key: str):
    prev = st.session_state.get("analysis_context_key")
    if prev != context_key:
        st.session_state["analysis_context_key"] = context_key
        for k in list(st.session_state.keys()):
            if k.startswith("download_") or k.startswith("dl_"):
                del st.session_state[k]


def _fmt_number(value, decimals=0):
    if value is None:
        return "-"
    if decimals == 0:
        return f"{value:,.0f}".replace(",", ".")
    return f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _render_weekly_kpi_header(weekly_data: dict):
    total_rolls_per_day = weekly_data.get("total_rolls_per_day")
    workers_per_day = weekly_data.get("workers_per_day")
    saegen_by_day_shift = weekly_data.get("saegen_by_day_shift")

    total_rolls = float(total_rolls_per_day.sum()) if total_rolls_per_day is not None else 0
    active_days = int((total_rolls_per_day > 0).sum()) if total_rolls_per_day is not None else 0
    avg_rolls_day = total_rolls / active_days if active_days else 0

    total_workers = float(workers_per_day.sum()) if workers_per_day is not None else 0
    avg_workers_day = float(workers_per_day.mean()) if workers_per_day is not None and len(workers_per_day) else 0
    rolls_per_ma = total_rolls / total_workers if total_workers else 0

    if saegen_by_day_shift is not None and not saegen_by_day_shift.empty:
        saegen_totals = saegen_by_day_shift.sum(axis=1)
        valid_saegen = saegen_totals[saegen_totals > 0]
        avg_saegen = float(valid_saegen.mean()) if not valid_saegen.empty else 0
    else:
        avg_saegen = 0

    st.markdown("### Executive KPIs")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rollen gesamt", _fmt_number(total_rolls), delta=f"{_fmt_number(avg_rolls_day)} / aktiver Tag")
    c2.metric("Ø Sägen / Tag", _fmt_number(avg_saegen), delta=f"{avg_saegen - SAEGEN_TARGET:+.0f} vs Ziel")
    c3.metric("Rollen pro MA", _fmt_number(rolls_per_ma, 1), delta=f"{_fmt_number(total_workers, 1)} MA gesamt")
    c4.metric("Ø MA / Tag", _fmt_number(avg_workers_day, 1), delta=f"{active_days} aktive Tage")


def _render_daily_kpi_header(daily_data: dict):
    rolls_pivot = daily_data.get("rolls_pivot")
    hours_pivot = daily_data.get("hours_pivot")
    shift_task_merged = daily_data.get("shift_task_merged")

    total_rolls = float(rolls_pivot.values.sum()) if rolls_pivot is not None and not rolls_pivot.empty else 0
    total_hours = float(hours_pivot.values.sum()) if hours_pivot is not None and not hours_pivot.empty else 0
    calculated_ma = total_hours / 7.5 if total_hours else 0

    if shift_task_merged is not None and not shift_task_merged.empty:
        available_ma = float(
            shift_task_merged.loc[
                shift_task_merged["Metric"].str.lower().str.strip() == "vorhandene ma",
                "Value",
            ].sum()
        )
    else:
        available_ma = 0

    ma_delta = available_ma - calculated_ma
    productivity = total_rolls / available_ma if available_ma else 0

    st.markdown("### Executive KPIs")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rollen gesamt", _fmt_number(total_rolls), delta=f"{_fmt_number(productivity, 1)} Rollen / MA")
    c2.metric("Arbeitsstunden", _fmt_number(total_hours, 1), delta=f"{_fmt_number(calculated_ma, 1)} berechnete MA")
    c3.metric("Vorhandene MA", _fmt_number(available_ma, 1), delta=f"{ma_delta:+.1f} vs Bedarf")
    c4.metric("Differenz MA", _fmt_number(ma_delta, 1), delta="Überdeckung" if ma_delta >= 0 else "Unterdeckung")


# =========================
# Sidebar controls
# =========================
with st.sidebar:
    st.subheader("🔧 Steuerung")

    uploaded_file = st.file_uploader("Excel hochladen", type=["xlsx"])
    mode = st.selectbox("Analysemodus wählen", ["Woche", "Tag"], index=0)

    if mode == "Woche":
        date_input = st.date_input("Datum innerhalb der Woche auswählen", value=datetime.today())
        week_number = date_input.isocalendar()[1]
        zip_name = f"rollenbewegung_KW{week_number}.zip"
    else:
        date_input = st.date_input("Tag auswählen", value=datetime.today())
        zip_name = f"rollenbewegung_{date_input.strftime('%Y-%m-%d')}.zip"

    start_analysis = st.button("Analyse starten")



# =========================
# Main analysis
# =========================
figs_meta: List[Dict] = []

if start_analysis:
    if not uploaded_file:
        st.warning("Bitte laden Sie zuerst eine Excel-Datei hoch.")
        st.stop()

    context_key = make_context_key(uploaded_file, mode, datetime.combine(date_input, datetime.min.time()))
    _reset_selection_state_if_needed(context_key)

    # Load data
    df_raw, summary_long, angaben_df, minutes_col = load_excel(uploaded_file)
    if summary_long is None or summary_long.empty:
        st.error("⚠️ Die hochgeladene Datei enthält keine gültigen Daten.")
        st.stop()

    # Aggregate & plot
    if mode == "Woche":
        week_number = date_input.isocalendar()[1]
        st.subheader(f"📈 Wochenanalyse – KW {week_number}")
        weekly_data = aggregate_weekly(summary_long, date_input)
        if not weekly_data:
            st.info("Keine Daten für die gewählte Woche.")
            st.stop()
        _render_weekly_kpi_header(weekly_data)
        figs_meta = plot_weekly_charts(weekly_data, date_input)
    else:
        date_str = date_input.strftime("%d.%m.%Y")
        st.subheader(f"📆 Tagesanalyse – {date_str}")
        daily_data = aggregate_daily(summary_long, angaben_df, date_input, minutes_col)
        if not daily_data:
            st.info("Keine Daten für den gewählten Tag.")
            st.stop()
        _render_daily_kpi_header(daily_data)
        figs_meta = plot_daily_charts(daily_data, date_input)

    if not figs_meta:
        st.info("Keine Diagramme für die gewählte Auswahl.")
        st.stop()

    # Main download section
    st.markdown("### ⬇️ Download")
    render_download_section(figs_meta=figs_meta, context_key=context_key, zip_name=zip_name)

    # Charts
    st.markdown("---")
    for item in figs_meta:
        st.pyplot(item["fig"], clear_figure=False)

