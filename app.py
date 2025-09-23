# app.py
import hashlib
from datetime import datetime
from typing import List, Dict

import streamlit as st

# Your modules
from data_loader import load_excel
from aggregation import aggregate_weekly, aggregate_daily
from weekly_plots import plot_weekly_charts
from daily_plots import plot_daily_charts
from download_utils import render_download_section


# =========================
# Page settings
# =========================
st.set_page_config(page_title="Rollenbewegung Dashboard", layout="wide")
st.title("📊 Rollenbewegung Dashboard")


# =========================
# Helpers
# =========================
def _file_hash(uploaded_file) -> str:
    """Hash file content efficiently in chunks."""
    if not uploaded_file:
        return ""
    pos = uploaded_file.tell()
    sha1 = hashlib.sha1()
    for chunk in iter(lambda: uploaded_file.read(8192), b""):
        sha1.update(chunk)
    uploaded_file.seek(pos)
    return sha1.hexdigest()[:10]


def _make_context_key(file_hash: str, mode: str, date_value: datetime) -> str:
    date_str = date_value.strftime("%Y-%m-%d") if isinstance(date_value, datetime) else str(date_value)
    raw = f"{file_hash}|{mode}|{date_str}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]


def _reset_selection_state_if_needed(context_key: str):
    prev = st.session_state.get("analysis_context_key")
    if prev != context_key:
        st.session_state["analysis_context_key"] = context_key
        for k in list(st.session_state.keys()):
            if k.startswith("download_") or k.startswith("dl_"):
                del st.session_state[k]


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

    file_hash = _file_hash(uploaded_file)
    context_key = _make_context_key(file_hash, mode, datetime.combine(date_input, datetime.min.time()))
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
        figs_meta = plot_weekly_charts(weekly_data, date_input)
    else:
        date_str = date_input.strftime("%d.%m.%Y")
        st.subheader(f"📆 Tagesanalyse – {date_str}")
        daily_data = aggregate_daily(summary_long, angaben_df, date_input, minutes_col)
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

