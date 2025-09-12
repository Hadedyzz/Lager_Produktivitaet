import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime

# --- new: fixed shift order and color mapping ---
shift_color_map = {"Früh": "#1f77b4", "Spät": "#ff7f0e", "Nacht": "#2ca02c"}
shift_order = ["Früh", "Spät", "Nacht"]

def shift_colors_for_cols(cols):
    return [shift_color_map.get(c, "#cccccc") for c in cols]

def format_day_month(dates):
    """Return short day.month labels for x-ticks."""
    idx = pd.to_datetime(dates)
    return [d.strftime("%d.%m") if not pd.isna(d) else "" for d in idx]

# =====================
# 1. LOAD AND PREPARE DATA
# =====================
@st.cache_data
def load_excel(file):
    """Load Excel, parse September sheet into tidy format, and load Angaben sheet"""
    # try both capitalizations (September / september)
    try:
        raw = pd.read_excel(file, sheet_name="September", header=None)
    except Exception:
        raw = pd.read_excel(file, sheet_name="september", header=None)

    angaben_df = pd.read_excel(file, sheet_name="Angaben")

    # --- parse into tidy structure ---
    dates = raw.iloc[0, 1:].tolist()
    records = []
    i, n = 1, len(raw)

    while i < n:
        # skip empty rows
        while i < n and (pd.isna(raw.iloc[i, 0]) or str(raw.iloc[i, 0]).strip() == ''):
            i += 1
        if i >= n:
            break
        block_start = i
        i += 1
        # find end of block
        while i < n and not (pd.isna(raw.iloc[i, 0]) or str(raw.iloc[i, 0]).strip() == ''):
            i += 1
        block_end = i

        block = raw.iloc[block_start:block_end].reset_index(drop=True)
        if block.empty:
            continue

        teams    = block.iloc[0, 1:].tolist()
        schichten = block.iloc[1, 1:].tolist()

        for kpi_row in range(2, block.shape[0]):
            kpi_name = block.iloc[kpi_row, 0]
            if pd.isna(kpi_name) or str(kpi_name).strip() == '':
                continue
            for col in range(1, block.shape[1]):
                datum   = dates[col-1]
                team    = teams[col-1]
                schicht = schichten[col-1]
                value   = block.iloc[kpi_row, col]
                if pd.isna(datum) or str(datum).strip() == '':
                    continue
                records.append({
                    "Datum": datum,
                    "Team": team,
                    "Schicht": schicht,
                    "Metric": kpi_name,
                    "Value": value
                })
        i += 1

    df_long = pd.DataFrame(records)
    df_long["Value"] = pd.to_numeric(df_long["Value"], errors="coerce")
    df_long["Datum"] = pd.to_datetime(df_long["Datum"], errors="coerce", dayfirst=True)

    return df_long, angaben_df

# =====================
# 2. WEEKLY PLOTS
# =====================
def plot_weekly_charts(df, target_date):
    """Weekly charts: df must include Datum, Schicht, Metric, Value"""
    figs = []
    # prepare date range for the week
    target_date = pd.to_datetime(target_date)
    start_of_week = target_date - pd.Timedelta(days=target_date.weekday())
    end_of_week = start_of_week + pd.Timedelta(days=6)
    all_dates = pd.date_range(start=start_of_week, end=end_of_week)

    # normalize shifts
    df = df.copy()
    df["Schicht"] = df["Schicht"].astype(str).str.strip().str.title()

    # ------------------------------------------------------------
    # 1. Line chart: total rolls per day
    # ------------------------------------------------------------
    rolls_df = df[df["Metric"].str.contains("rollen", case=False, na=False)].copy()
    rolls_df["Value"] = pd.to_numeric(rolls_df["Value"], errors="coerce")
    rolls_series = rolls_df.groupby("Datum")["Value"].sum().reindex(all_dates).fillna(0)

    fig, ax = plt.subplots(figsize=(10,5))
    rolls_series.plot(kind="line", ax=ax, marker="o")
    ax.set_title("Gesamtanzahl der Rollenbewegungen pro Tag")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Anzahl Rollenbewegungen")
    ax.set_xticklabels(format_day_month(rolls_series.index), rotation=45)
    target = 70
    avg_rolls = rolls_series.mean()
    ax.axhline(target, color="red", linestyle="--", linewidth=2, label="Ziel 70")
    ax.axhline(avg_rolls, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
    ax.legend()
    figs.append(fig)

    # ------------------------------------------------------------
    # 2. Stacked column chart for "sägen" rolls per day by shift
    # ------------------------------------------------------------
    saegen_df = df[df["Metric"].str.contains("sägen", case=False, na=False)].copy()
    saegen_df["Value"] = pd.to_numeric(saegen_df["Value"], errors="coerce")
    saegen_by_day_shift = saegen_df.groupby(["Datum", "Schicht"])["Value"].sum().unstack("Schicht").fillna(0)
    # enforce fixed shift order and drop missing columns
    saegen_by_day_shift = saegen_by_day_shift.reindex(columns=shift_order).fillna(0)
    saegen_by_day_shift = saegen_by_day_shift.reindex(all_dates).fillna(0)

    fig, ax = plt.subplots(figsize=(12,6))
    cols = shift_colors_for_cols(saegen_by_day_shift.columns)
    saegen_by_day_shift.plot(kind="bar", stacked=True, ax=ax, color=cols)
    ax.set_title("Anzahl gesägte Rollen pro Tag")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Anzahl gesägte Rollen")
    ax.set_xticklabels(format_day_month(saegen_by_day_shift.index), rotation=45)
    target = 70
    avg_saegen = saegen_by_day_shift.sum(axis=1).mean()
    ax.axhline(target, color="red", linestyle="--", linewidth=2, label="Ziel 70")
    ax.axhline(avg_saegen, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
    # color legend text to match shift
    leg = ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    if leg:
        texts = leg.get_texts()
        # first two texts are the horizontal lines labels if present -> find shift labels by name
        for t in texts:
            txt = t.get_text()
            if txt in saegen_by_day_shift.columns:
                t.set_color(shift_color_map.get(txt, "#000"))
    figs.append(fig)

    # ------------------------------------------------------------
    # 4. Stacked bar graph of available workers per day by shift
    # ------------------------------------------------------------
    workers = df[df["Metric"] == "Vorhandene MA"].copy()
    workers_by_day_shift = workers.groupby(["Datum", "Schicht"])["Value"].sum().unstack("Schicht").fillna(0)
    workers_by_day_shift = workers_by_day_shift.reindex(columns=shift_order).fillna(0)
    workers_by_day_shift = workers_by_day_shift.reindex(all_dates).fillna(0)

    fig, ax = plt.subplots(figsize=(10,5))
    cols = shift_colors_for_cols(workers_by_day_shift.columns)
    workers_by_day_shift.plot(kind="bar", stacked=True, ax=ax, color=cols)
    ax.set_title("Anzahl MA pro Tag")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Anzahl MA")
    ax.set_xticklabels(format_day_month(workers_by_day_shift.index), rotation=45)
    avg_workers = workers_by_day_shift.sum(axis=1).mean()
    ax.axhline(avg_workers, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
    leg = ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    if leg:
        for t in leg.get_texts():
            txt = t.get_text()
            if txt in workers_by_day_shift.columns:
                t.set_color(shift_color_map.get(txt, "#000"))
    figs.append(fig)

    # ------------------------------------------------------------
    # 5. Stacked column chart: total rolls per day by shift
    # ------------------------------------------------------------
    total_rolls_per_day_shift = rolls_df.groupby(["Datum", "Schicht"])["Value"].sum().unstack("Schicht").fillna(0)
    total_rolls_per_day_shift = total_rolls_per_day_shift.reindex(columns=shift_order).fillna(0)
    total_rolls_per_day_shift = total_rolls_per_day_shift.reindex(all_dates).fillna(0)

    fig, ax = plt.subplots(figsize=(12,6))
    cols = shift_colors_for_cols(total_rolls_per_day_shift.columns)
    total_rolls_per_day_shift.plot(kind="bar", stacked=True, ax=ax, color=cols)
    ax.set_title("Gesamte Anzahl an bewegten Rollen pro Tag nach Schicht")
    ax.set_xlabel("Datum")
    ax.set_ylabel("Gesamtrollen")
    ax.set_xticklabels(format_day_month(total_rolls_per_day_shift.index), rotation=45)
    avg_shift = total_rolls_per_day_shift.sum(axis=1).mean()
    ax.axhline(avg_shift, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
    leg = ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    if leg:
        for t in leg.get_texts():
            txt = t.get_text()
            if txt in total_rolls_per_day_shift.columns:
                t.set_color(shift_color_map.get(txt, "#000"))
    figs.append(fig)

    return figs

# =====================
# 3. DAILY PLOTS
# =====================
def plot_daily_charts(df, angaben_df, target_day):
    figs = []
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    day_data = df[(df["Datum"].dt.date == target_day) & (df["Metric"].str.contains(r"\(Stück\)$"))].copy()

    # Join with Angaben
    angaben_df["Task"] = angaben_df["Task"].str.strip().str.lower()
    day_data["Task"] = day_data["Metric"].str.replace(r"\s*\(Stück\)$", "", regex=True).str.strip().str.lower()
    shift_task = day_data.groupby(["Schicht", "Task"])["Value"].sum().reset_index()
    shift_task = pd.merge(shift_task, angaben_df, on="Task", how="inner")
    shift_task["Hours"] = (shift_task["Value"] * shift_task["Minutes/roll"]) / 60
    shift_task["FTE"] = shift_task["Hours"] / 7.5
    shift_task = shift_task[shift_task["Hours"] > 0]

    shift_color_map = {"Früh": "#1f77b4", "Spät": "#ff7f0e", "Nacht": "#2ca02c"}

    # Pareto hours
    stacked_df = shift_task.pivot_table(index="Task", columns="Schicht", values="Hours", aggfunc="sum", fill_value=0)
    stacked_df = stacked_df.reindex(columns=shift_order).fillna(0)
    task_order = stacked_df.sum(axis=1).sort_values(ascending=False).index
    stacked_df = stacked_df.reindex(task_order)
    fig, ax = plt.subplots(figsize=(12,6))
    cols = shift_colors_for_cols(stacked_df.columns)
    stacked_df.plot(kind="bar", stacked=True, ax=ax, color=cols)
    ax.set_title(f"Arbeitsstunden pro Aufgabe und Schicht am {target_day.strftime('%d.%m.%Y')}")
    ax.set_xlabel("Aufgabe")
    ax.set_ylabel("Stunden")
    ax.set_xticklabels([t.capitalize() for t in task_order], rotation=60, ha="right")
    leg = ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    if leg:
        for t in leg.get_texts():
            txt = t.get_text()
            if txt in stacked_df.columns:
                t.set_color(shift_color_map.get(txt, "#000"))
    figs.append(fig)

    # Pareto rolls
    stacked_rolls = shift_task.pivot_table(index="Task", columns="Schicht", values="Value", aggfunc="sum", fill_value=0)
    stacked_rolls = stacked_rolls.reindex(columns=shift_order).fillna(0)
    task_order_rolls = stacked_rolls.sum(axis=1).sort_values(ascending=False).index
    stacked_rolls = stacked_rolls.reindex(task_order_rolls)
    fig2, ax2 = plt.subplots(figsize=(12,6))
    cols2 = shift_colors_for_cols(stacked_rolls.columns)
    stacked_rolls.plot(kind="bar", stacked=True, ax=ax2, color=cols2)
    ax2.set_title(f"Rollen pro Aufgabe und Schicht am {target_day.strftime('%d.%m.%Y')}")
    ax2.set_xlabel("Aufgabe")
    ax2.set_ylabel("Rollen")
    ax2.set_xticklabels([t.capitalize() for t in task_order_rolls], rotation=60, ha="right")
    leg2 = ax2.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    if leg2:
        for t in leg2.get_texts():
            txt = t.get_text()
            if txt in stacked_rolls.columns:
                t.set_color(shift_color_map.get(txt, "#000"))
    figs.append(fig2)

    return figs

# =====================
# 4. STREAMLIT APP
# =====================
st.set_page_config(page_title="Rollenbewegung Dashboard", layout="wide")
st.title("📊 Rollenbewegung Dashboard")

uploaded_file = st.file_uploader("Excel hochladen", type=["xlsx"])
if uploaded_file:
    df_raw, angaben_df = load_excel(uploaded_file)

    if not {"Datum", "Schicht", "Metric", "Value"}.issubset(df_raw.columns):
        st.error("⚠️ Deine Excel-Datei muss Spalten enthalten: Datum, Schicht, Metric, Value")
    else:
        mode = st.selectbox("Analysemodus wählen", ["Woche", "Tag"])

        if mode == "Woche":
            week_date = st.date_input("Datum innerhalb der Woche auswählen", datetime.today())
            if st.button("Wochen-Analyse starten"):
                figs = plot_weekly_charts(df_raw, week_date)
                for fig in figs:
                    st.pyplot(fig)

        elif mode == "Tag":
            day_date = st.date_input("Tag auswählen", datetime.today())
            if st.button("Tages-Analyse starten"):
                figs = plot_daily_charts(df_raw, angaben_df, day_date)
                for fig in figs:
                    st.pyplot(fig)

