import io
from datetime import datetime

import matplotlib

matplotlib.use("Agg")
matplotlib.rcParams["font.family"] = "DejaVu Sans"
matplotlib.rcParams["figure.dpi"] = 96

import pandas as pd
import streamlit as st

from aggregation import (
    aggregate_daily,
    aggregate_longterm,
    aggregate_monthly,
    aggregate_weekly,
)
from config import SAEGEN_TARGET, SAVEFIG_DPI
from daily_plots import plot_daily_charts
from data_loader import load_excel
from download_utils import render_download_section
from helpers import make_context_key
from kpi_header import render_kpi_header
from longterm_insights import generate_longterm_insights
from longterm_plots import plot_longterm_charts
from monthly_plots import plot_monthly_charts
from weekly_plots import plot_weekly_charts


st.set_page_config(page_title="Rollenbewegung Dashboard", layout="wide")
st.title("📊 Rollenbewegung Dashboard")


def _fmt_number(value, decimals=0):
    if value is None:
        return "-"
    if decimals == 0:
        return f"{float(value):,.0f}".replace(",", ".")
    return f"{float(value):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fig_to_png_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=SAVEFIG_DPI, bbox_inches="tight")
    buf.seek(0)
    return buf


def _render_chart_outputs(figs_meta, context_key, zip_name):
    st.markdown("### Download")
    render_download_section(figs_meta=figs_meta, context_key=context_key, zip_name=zip_name)
    st.markdown("---")
    for item in figs_meta:
        try:
            st.pyplot(item["fig"], clear_figure=False)
            st.download_button(
                "PNG herunterladen",
                data=_fig_to_png_bytes(item["fig"]),
                file_name=item.get("filename", "diagramm.png"),
                mime="image/png",
                key=f"single_{context_key}_{item.get('filename', item['title'])}",
            )
        except Exception as e:
            st.warning(f"Dieses Diagramm konnte nicht angezeigt werden: {item.get('title', 'Diagramm')} ({e})")


def _daily_totals(daily_data):
    rolls_pivot = daily_data.get("rolls_pivot")
    hours_pivot = daily_data.get("hours_pivot")
    merged = daily_data.get("shift_task_merged")
    total_rolls = float(rolls_pivot.values.sum()) if rolls_pivot is not None and not rolls_pivot.empty else 0
    total_hours = float(hours_pivot.values.sum()) if hours_pivot is not None and not hours_pivot.empty else 0
    calculated_ma = total_hours / 7.5 if total_hours else 0
    available_ma = 0.0
    if merged is not None and not merged.empty:
        available_ma = float(merged.loc[merged["Metric"].str.lower().str.strip() == "vorhandene ma", "Value"].sum())
    delta_ma = available_ma - calculated_ma
    return {
        "total_rolls": total_rolls,
        "total_hours": total_hours,
        "calculated_ma": calculated_ma,
        "available_ma": available_ma,
        "delta_ma": delta_ma,
    }


def _delta_vs_previous(current, previous, decimals=0):
    if previous is None:
        return None
    change = current - previous
    return f"{change:+.{decimals}f} vs Vortag"


def _daily_kpis(daily_data, previous_daily_data=None):
    totals = _daily_totals(daily_data)
    previous = _daily_totals(previous_daily_data) if previous_daily_data else None
    total_rolls = totals["total_rolls"]
    total_hours = totals["total_hours"]
    calculated_ma = totals["calculated_ma"]
    available_ma = totals["available_ma"]
    delta_ma = totals["delta_ma"]
    return [
        {"label": "Total Rollen", "value": _fmt_number(total_rolls), "delta": _delta_vs_previous(total_rolls, previous["total_rolls"] if previous else None, 0)},
        {"label": "Σ Stunden", "value": _fmt_number(total_hours, 1), "delta": _delta_vs_previous(total_hours, previous["total_hours"] if previous else None, 1)},
        {"label": "Verfügbare MA", "value": _fmt_number(available_ma, 1), "delta": _delta_vs_previous(available_ma, previous["available_ma"] if previous else None, 1)},
        {"label": "Berechnete MA", "value": _fmt_number(calculated_ma, 1), "delta": _delta_vs_previous(calculated_ma, previous["calculated_ma"] if previous else None, 1)},
        {"label": "Δ MA", "value": _fmt_number(delta_ma, 1), "delta": _delta_vs_previous(delta_ma, previous["delta_ma"] if previous else None, 1)},
    ]


def _weekly_kpis(weekly_data, previous_week_data=None):
    total_rolls = float(weekly_data["total_rolls_per_day"].sum())
    workers = float(weekly_data["workers_per_day"].sum())
    rolls_per_ma = total_rolls / workers if workers else 0
    saegen_daily = weekly_data["saegen_by_day_shift"].sum(axis=1)
    active_saegen = saegen_daily[saegen_daily > 0]
    attainment = float((active_saegen >= SAEGEN_TARGET).mean()) if len(active_saegen) else 0
    gap = weekly_data.get("differenz_ma_per_day", pd.Series(dtype=float))
    gap = gap[gap != 0]
    volatility = float(weekly_data["total_rolls_per_day"].std(ddof=0) / weekly_data["total_rolls_per_day"].mean()) if weekly_data["total_rolls_per_day"].mean() else 0
    prev_delta = None
    if previous_week_data:
        prev_total = float(previous_week_data["total_rolls_per_day"].sum())
        prev_workers = float(previous_week_data["workers_per_day"].sum())
        prev_rpm = prev_total / prev_workers if prev_workers else 0
        prev_delta = f"{rolls_per_ma - prev_rpm:+.1f} vs Vorwoche"
    return [
        {"label": "Σ Rollen", "value": _fmt_number(total_rolls), "delta": None},
        {"label": "Ø Rollen/MA", "value": _fmt_number(rolls_per_ma, 1), "delta": prev_delta},
        {"label": "Ziel-Erreichung Sägen", "value": f"{attainment * 100:.0f}%", "delta": f"Ziel {SAEGEN_TARGET}"},
        {"label": "Ø Δ MA", "value": _fmt_number(gap.mean() if len(gap) else 0, 1), "delta": None},
        {"label": "Volatilität", "value": _fmt_number(volatility, 2), "delta": None},
    ]


def _monthly_kpis(monthly_data, previous_month_data=None):
    total_rolls = monthly_data.get("total_rolls", 0)
    workdays = monthly_data.get("workdays", 0)
    attainment = monthly_data["saegen"]["attainment_rate"] if monthly_data.get("saegen") else 0
    gap_mean = monthly_data["staffing"]["mean"] if monthly_data.get("staffing") else 0
    anomaly_count = len(monthly_data.get("anomalies", []))
    delta = None
    if previous_month_data:
        delta = f"{total_rolls - previous_month_data.get('total_rolls', 0):+.0f} vs Vormonat"
    return [
        {"label": "Σ Rollen", "value": _fmt_number(total_rolls), "delta": delta},
        {"label": "Arbeitstage", "value": _fmt_number(workdays), "delta": None},
        {"label": "Sägen-Attainment", "value": f"{attainment * 100:.0f}%", "delta": None},
        {"label": "Ø Δ MA", "value": _fmt_number(gap_mean, 1), "delta": None},
        {"label": "Anomalie-Tage", "value": _fmt_number(anomaly_count), "delta": None},
    ]


def _longterm_kpis(longterm_data):
    weeks = longterm_data.get("weeks", [])
    rpm = longterm_data.get("rolls_per_ma", pd.Series(dtype=float))
    saegen = longterm_data.get("saegen_week")
    volatility = longterm_data.get("volatility", 0)
    start = rpm.head(4).mean() if len(rpm) else 0
    end = rpm.tail(4).mean() if len(rpm) else 0
    trend = ((end - start) / start * 100) if start else 0
    saegen_start = saegen["mean"].head(4).mean() if saegen is not None and not saegen.empty else 0
    saegen_end = saegen["mean"].tail(4).mean() if saegen is not None and not saegen.empty else 0
    return [
        {"label": "Anzahl Wochen", "value": _fmt_number(len(weeks)), "delta": None},
        {"label": "Trend Rollen/MA", "value": f"{trend:+.1f}%", "delta": "letzte 4 vs erste 4 Wochen"},
        {"label": "Trend Sägen-Mittel", "value": _fmt_number(saegen_end, 1), "delta": f"{saegen_end - saegen_start:+.1f}"},
        {"label": "Volatilität-Trend", "value": _fmt_number(volatility, 2), "delta": None},
        {"label": "Auslastungs-Drift", "value": _fmt_number(end - start, 1), "delta": None},
    ]


def _render_empty(message):
    st.markdown(
        f"""
        <div style="text-align:center; padding:2.5rem 1rem; border:1px solid #e6e6e6; border-radius:8px;">
            <div style="font-size:2rem; margin-bottom:0.5rem;">ℹ️</div>
            <div style="font-size:1.05rem;">{message}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


with st.sidebar:
    st.subheader("Steuerung")
    uploaded_file = st.file_uploader("Excel hochladen", type=["xlsx"])

if not uploaded_file:
    _render_empty("Bitte laden Sie eine Excel-Datei hoch, um die Analyse zu starten.")
    st.stop()

with st.spinner("Daten werden analysiert..."):
    df_raw, summary_long, angaben_df, minutes_col = load_excel(uploaded_file)

if summary_long is None or summary_long.empty:
    st.error("Die hochgeladene Datei enthält keine gültigen Daten.")
    st.stop()

summary_long["Datum"] = pd.to_datetime(summary_long["Datum"], errors="coerce")
with st.sidebar:
    shift_options = [shift for shift in ["Früh", "Spät", "Nacht"] if shift in set(summary_long["Schicht"].dropna().astype(str))]
    selected_shifts = st.multiselect("Schichten", shift_options, default=shift_options)
    team_options = sorted(summary_long["Team"].dropna().astype(str).unique().tolist()) if "Team" in summary_long else []
    selected_teams = st.multiselect("Teams", team_options, default=team_options)

if selected_shifts:
    summary_long = summary_long[summary_long["Schicht"].astype(str).isin(selected_shifts)].copy()
if selected_teams:
    summary_long = summary_long[summary_long["Team"].astype(str).isin(selected_teams)].copy()
if summary_long.empty:
    _render_empty("Für die gewählten Filter sind keine Daten vorhanden.")
    st.stop()

roll_metric_mask = ~summary_long["Metric"].str.lower().str.strip().isin(
    ["vorhandene ma", "benötigte ma", "differenz ma", "sonstiges / aufräumarbeiten (in std)"]
)
active_roll_dates = (
    summary_long.loc[roll_metric_mask]
    .groupby(summary_long["Datum"].dt.normalize(), observed=False)["Value"]
    .sum()
    .loc[lambda s: s > 0]
)
date_source = active_roll_dates.index if len(active_roll_dates) else summary_long["Datum"].dropna().dt.normalize().unique()
available_dates = sorted(date_source)
if not available_dates:
    _render_empty("Nach dem Laden wurden keine gültigen Datumswerte gefunden.")
    st.stop()

min_date = pd.Timestamp(available_dates[0]).date()
max_date = pd.Timestamp(available_dates[-1]).date()
available_months = sorted(pd.Series(available_dates).dt.to_period("M").astype(str).unique())

tab_day, tab_week, tab_month, tab_long = st.tabs(["Tag", "Woche", "Monat", "Langzeit"])

with tab_day:
    date_input = st.date_input("Tag auswählen", value=max_date, min_value=min_date, max_value=max_date, key="day_picker")
    context_key = make_context_key(uploaded_file, "Tag", datetime.combine(date_input, datetime.min.time()))
    previous_dates = [pd.Timestamp(day).date() for day in available_dates if pd.Timestamp(day).date() < date_input]
    previous_day = previous_dates[-1] if previous_dates else None
    with st.spinner("Daten werden analysiert..."):
        daily_data = aggregate_daily(summary_long, angaben_df, pd.Timestamp(date_input), minutes_col)
        previous_daily_data = (
            aggregate_daily(summary_long, angaben_df, pd.Timestamp(previous_day), minutes_col)
            if previous_day else None
        )
    if not daily_data:
        _render_empty("Keine Daten für den gewählten Tag.")
    else:
        render_kpi_header(_daily_kpis(daily_data, previous_daily_data))
        try:
            figs_meta = plot_daily_charts(daily_data, pd.Timestamp(date_input))
        except Exception as e:
            st.error(f"Die Tagesdiagramme konnten nicht gerendert werden: {e}")
            figs_meta = []
        if figs_meta:
            _render_chart_outputs(figs_meta, context_key, f"rollenbewegung_{date_input.strftime('%Y-%m-%d')}.zip")

with tab_week:
    week_date = st.date_input("Datum innerhalb der Woche auswählen", value=max_date, min_value=min_date, max_value=max_date, key="week_picker")
    context_key = make_context_key(uploaded_file, "Woche", datetime.combine(week_date, datetime.min.time()))
    with st.spinner("Daten werden analysiert..."):
        weekly_data = aggregate_weekly(summary_long, pd.Timestamp(week_date))
        previous_week_data = aggregate_weekly(summary_long, pd.Timestamp(week_date) - pd.Timedelta(days=7))
    if not weekly_data:
        _render_empty("Keine Daten für die gewählte Woche.")
    else:
        render_kpi_header(_weekly_kpis(weekly_data, previous_week_data))
        try:
            figs_meta = plot_weekly_charts(weekly_data, pd.Timestamp(week_date))
        except Exception as e:
            st.error(f"Die Wochendiagramme konnten nicht gerendert werden: {e}")
            figs_meta = []
        if figs_meta:
            kw = pd.Timestamp(week_date).isocalendar().week
            _render_chart_outputs(figs_meta, context_key, f"rollenbewegung_KW{kw}.zip")

with tab_month:
    selected_month = st.selectbox("Monat auswählen", available_months, index=len(available_months) - 1)
    period = pd.Period(selected_month)
    context_key = make_context_key(uploaded_file, "Monat", datetime(period.year, period.month, 1))
    previous_period = period - 1
    with st.spinner("Daten werden analysiert..."):
        monthly_data = aggregate_monthly(summary_long, period.year, period.month, angaben_df, minutes_col)
        previous_month_data = aggregate_monthly(summary_long, previous_period.year, previous_period.month, angaben_df, minutes_col)
    if not monthly_data:
        _render_empty("Keine Daten für den gewählten Monat.")
    else:
        render_kpi_header(_monthly_kpis(monthly_data, previous_month_data))
        try:
            figs_meta = plot_monthly_charts(monthly_data, previous_month_data)
        except Exception as e:
            st.error(f"Die Monatsdiagramme konnten nicht gerendert werden: {e}")
            figs_meta = []
        anomalies = monthly_data.get("anomalies")
        st.markdown("### Anomalie-Tage")
        if anomalies is None or anomalies.empty:
            st.info("Keine Anomalien gefunden.")
        else:
            st.dataframe(anomalies, use_container_width=True)
        if figs_meta:
            _render_chart_outputs(figs_meta, context_key, f"rollenbewegung_{period.year}-{period.month:02d}.zip")

with tab_long:
    context_key = make_context_key(uploaded_file, "Langzeit", datetime(1900, 1, 1))
    with st.spinner("Daten werden analysiert..."):
        longterm_data = aggregate_longterm(
            summary_long,
            angaben_df,
            minutes_col,
            as_of_date=pd.Timestamp.today().normalize(),
        )
    if not longterm_data:
        _render_empty("Keine Langzeitdaten verfügbar.")
    else:
        render_kpi_header(_longterm_kpis(longterm_data))
        with st.expander("📊 Automatische Erkenntnisse", expanded=True):
            for insight in generate_longterm_insights(longterm_data):
                st.write(insight)
        try:
            figs_meta = plot_longterm_charts(longterm_data)
        except Exception as e:
            st.error(f"Die Langzeitdiagramme konnten nicht gerendert werden: {e}")
            figs_meta = []
        if figs_meta:
            _render_chart_outputs(figs_meta, context_key, "rollenbewegung_langzeit.zip")
