# aggregation.py
import pandas as pd
import streamlit as st

SHIFT_ORDER = ["Früh", "Spät", "Nacht"]
MA_KPIS = ["vorhandene ma", "benötigte ma", "differenz ma"]
SONSTIGES = "sonstiges / aufräumarbeiten (in std)"  # loader's Sonstiges

# Define Hauptaufgaben (main tasks) in lowercase to match loader normalization
MAIN_TASKS = ["absetzen", "richten", "verladen", "zusammenfahren"]
TASK_ORDER = ["Absetzen", "Richten", "Verladen", "Zusammenfahren", "Sonstige"]


def _reindex_shift(df: pd.DataFrame, weekdays: pd.DatetimeIndex) -> pd.DataFrame:
    """Helper to reindex by shift and weekdays consistently."""
    return (
        df.reindex(columns=SHIFT_ORDER).fillna(0)
          .reindex(weekdays).fillna(0)
    )


@st.cache_data(show_spinner=False)
def aggregate_weekly(summary_long: pd.DataFrame, target_date: pd.Timestamp):
    """
    Aggregate summary_long for the given week.

    Weekly rules:
    - Exclude MA KPIs (vorhandene/benötigte/differenz MA).
    - Exclude loader's "Sonstiges / Aufräumarbeiten (in Std)" (hours-only).
    - Group rolls into Hauptaufgaben (Absetzen, Richten, Verladen, Zusammenfahren)
      and bucket everything else as "Sonstige".
    """
    target_date = pd.to_datetime(target_date)
    start_of_week = target_date - pd.Timedelta(days=target_date.weekday())  # Monday
    end_of_week = start_of_week + pd.Timedelta(days=6)
    all_dates = pd.date_range(start=start_of_week, end=end_of_week)
    weekdays = all_dates[all_dates.weekday < 5]  # Mon–Fri

    df = summary_long.copy()
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)

    df = df[(df["Datum"] >= start_of_week) & (df["Datum"] <= end_of_week)]
    df = df[df["Datum"].isin(weekdays)]
    if df.empty:
        return {}

    # ---------------- Specialized ----------------
    # Sägen metrics (keep original rolls)
    saegen = df[df["Metric"].str.contains("sägen", case=False, na=False)]
    saegen_by_day_shift = (
        saegen.groupby(["Datum", "Schicht"], observed=False)["Value"].sum().unstack("Schicht")
    )
    saegen_by_day_shift = _reindex_shift(saegen_by_day_shift, weekdays)

    # Hauptaufgaben vs Sonstige
    def classify_task(metric: str) -> str:
        if not isinstance(metric, str):
            return "Sonstige"
        m = metric.strip().lower()
        if m.startswith("absetzen"):
            return "Absetzen"
        elif m.startswith("richten"):
            return "Richten"
        elif m.startswith("verladen"):
            return "Verladen"
        elif m.startswith("zusammenfahren"):
            return "Zusammenfahren"
        else:
            return "Sonstige"

    # Exclude MA KPIs and loader's Sonstiges (hours-only)
    rolls = df[
        (~df["Metric"].isin(MA_KPIS)) &
        (df["Metric"].str.lower().str.strip() != SONSTIGES)
    ].copy()

    rolls["TaskGroup"] = rolls["Metric"].apply(classify_task)

    total_rolls_by_group = (
        rolls.groupby(["Datum", "TaskGroup"], observed=False)["Value"].sum().unstack("TaskGroup")
        .reindex(columns=TASK_ORDER).fillna(0)
        .reindex(weekdays).fillna(0)
    )

    # Rolls per shift (all tasks, excluding MA KPIs and loader Sonstiges)
    total_shift = (
        rolls.groupby(["Datum", "Schicht"], observed=False)["Value"].sum().unstack("Schicht")
    )
    total_shift = _reindex_shift(total_shift, weekdays)
    
    # ---------------- MA per shift ----------------
    ma_by_shift = (
        df[df["Metric"].str.lower().str.strip() == "vorhandene ma"]
        .groupby(["Datum", "Schicht"], observed=False)["Value"]
        .sum()
        .unstack("Schicht")
    )
    ma_by_shift = _reindex_shift(ma_by_shift, weekdays)

    # Workers per day (only "vorhandene ma")
    workers_per_day = (
        df[df["Metric"].str.lower().str.strip() == "vorhandene ma"]
        .groupby("Datum", observed=False)["Value"].sum()
        .reindex(weekdays).fillna(0)
    )
    # Workers per shift (only "vorhandene ma")
    workers_per_shift = (
        df[df["Metric"].str.lower().str.strip() == "vorhandene ma"]
        .groupby(["Datum", "Schicht"], observed=False)["Value"].sum()
        .unstack("Schicht")  # wide format with shifts as columns
        .reindex(index=weekdays, columns=SHIFT_ORDER)  # enforce order
        .fillna(0)
    )
    # Total rolls per day (excluding MA KPIs + loader Sonstiges)
    total_rolls_per_day = (
        rolls.groupby("Datum", observed=False)["Value"].sum()
        .reindex(weekdays).fillna(0)
    )

    return {
        "dates": weekdays,
        "kw": target_date.isocalendar()[1],
        "saegen_by_day_shift": saegen_by_day_shift,
        "total_rolls_by_group": total_rolls_by_group,
        "total_shift": total_shift,
        "workers_per_shift": workers_per_shift,
        "workers_per_day": workers_per_day,
        "total_rolls_per_day": total_rolls_per_day,
        "ma_by_shift": ma_by_shift,  # <--- NEW
    }



@st.cache_data(show_spinner=False)
def aggregate_daily(summary_long: pd.DataFrame, angaben_df: pd.DataFrame,
                    target_day: pd.Timestamp, minutes_col: str):
    """
    Aggregate summary_long and angaben_df for the given day,
    returning only what daily plots need.

    Daily rules:
    - Exclude MA KPIs from task pivots.
    - Keep loader's Sonstiges (hours-only).
    - Drop tasks where Value == 0 across all shifts (except keep Sonstiges).
    """
    target_day = pd.to_datetime(target_day).normalize()

    df = summary_long.copy()
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)
    df = df[df["Datum"].dt.normalize() == target_day]
    if df.empty:
        return {}

    # ---------------- Merge with Angaben ----------------
    shift_task_rolls = df.groupby(["Schicht", "Metric"], observed=False)["Value"].sum().reset_index()

    angaben = angaben_df.copy()
    shift_task_merged = pd.merge(
        shift_task_rolls,
        angaben,
        left_on="Metric",
        right_on="Task",
        how="left"
    )

    # Task-level Hours & FTE
    if minutes_col and minutes_col in shift_task_merged.columns:
        shift_task_merged["Hours"] = (
            shift_task_merged["Value"] * shift_task_merged[minutes_col]
        ) / 60
        shift_task_merged["FTE"] = shift_task_merged["Hours"] / 7.5
    else:
        shift_task_merged["Hours"] = 0.0
        shift_task_merged["FTE"] = 0.0

    # Add loader's Sonstiges directly to Hours (ignore its Value)
    mask_sonst = shift_task_merged["Metric"].str.lower().str.strip() == SONSTIGES
    shift_task_merged.loc[mask_sonst, "Hours"] = shift_task_merged.loc[mask_sonst, "Value"]
    shift_task_merged.loc[mask_sonst, "Value"] = 0  # remove rolls for loader Sonstiges

    # Pretty labels
    if "Pretty" not in shift_task_merged.columns:
        shift_task_merged["Pretty"] = shift_task_merged["Metric"].str.title()

    # Normalize loader Sonstiges name
    shift_task_merged.loc[
        shift_task_merged["Metric"].str.lower().str.strip() == SONSTIGES,
        "Pretty"
    ] = "Sonstiges"

    # ---------------- Filter tasks ----------------
    tasks_df = shift_task_merged[
        ~shift_task_merged["Metric"].str.lower().str.strip().isin(MA_KPIS)
    ].copy()

    # Drop tasks with Value==0 for all shifts (except loader Sonstiges, now in Hours)
    nonzero_tasks = (
        tasks_df.groupby("Pretty")[["Value", "Hours"]].sum()
        .loc[lambda s: (s["Value"] != 0) | (s["Hours"] != 0)]
    )
    task_order = nonzero_tasks.sort_values("Value", ascending=False).index

    tasks_df = tasks_df[tasks_df["Pretty"].isin(task_order)]

    # ---------------- Pivots ----------------
    hours_pivot = (
        tasks_df.pivot_table(
            index="Pretty",
            columns="Schicht",
            values="Hours",
            aggfunc="sum",
            fill_value=0,
            observed=False,
        )
        .reindex(task_order)
        .reindex(columns=SHIFT_ORDER)
        .fillna(0)
    )

    rolls_pivot = (
        tasks_df.pivot_table(
            index="Pretty",
            columns="Schicht",
            values="Value",
            aggfunc="sum",
            fill_value=0,
            observed=False,
        )
        .reindex(task_order)
        .reindex(columns=SHIFT_ORDER)
        .fillna(0)
    )

    return {
        "shift_task_merged": shift_task_merged,  # includes MA KPIs
        "hours_pivot": hours_pivot,              # with loader Sonstiges
        "rolls_pivot": rolls_pivot,              # no loader Sonstiges
    }
