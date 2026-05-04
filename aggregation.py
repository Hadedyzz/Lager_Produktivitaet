import logging

import pandas as pd
import streamlit as st
from config import SAEGEN_TARGET
from notifications import add_message

SHIFT_ORDER = ["Früh", "Spät", "Nacht"]
MA_KPIS = ["vorhandene ma", "benötigte ma", "differenz ma"]
SONSTIGES = "sonstiges / aufräumarbeiten (in std)"  # loader's Sonstiges
logger = logging.getLogger(__name__)

# Define Hauptaufgaben (main tasks) in lowercase to match loader normalization
MAIN_TASKS = ["absetzen", "richten", "verladen", "zusammenfahren"]
TASK_ORDER = ["Absetzen", "Richten", "Verladen", "Zusammenfahren", "Sonstige"]


def filter_summary(summary_long: pd.DataFrame, start_date=None, end_date=None, metrics=None, shifts=None) -> pd.DataFrame:
    """Return a filtered, normalized copy of summary_long."""
    df = summary_long.copy()
    if df.empty:
        return df
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)
    df["Metric"] = df["Metric"].astype(str).str.strip().str.lower()
    if start_date is not None:
        df = df[df["Datum"] >= pd.to_datetime(start_date)]
    if end_date is not None:
        df = df[df["Datum"] <= pd.to_datetime(end_date)]
    if metrics is not None:
        metric_set = {str(metric).strip().lower() for metric in metrics}
        df = df[df["Metric"].isin(metric_set)]
    if shifts is not None:
        shift_set = set(shifts)
        df = df[df["Schicht"].isin(shift_set)]
    return df.copy()


def _roll_rows(df: pd.DataFrame) -> pd.DataFrame:
    return df[
        (~df["Metric"].isin(MA_KPIS))
        & (df["Metric"].str.lower().str.strip() != SONSTIGES)
    ].copy()


def compute_rolls_per_ma(df: pd.DataFrame, group_by=None):
    """Compute rolls per available MA for a dataframe slice."""
    if df.empty:
        return pd.Series(dtype=float) if group_by else 0.0
    group_cols = [] if group_by is None else ([group_by] if isinstance(group_by, str) else list(group_by))
    rolls = _roll_rows(df)
    ma = df[df["Metric"].str.lower().str.strip() == "vorhandene ma"]
    if not group_cols:
        workers = ma["Value"].sum()
        return float(rolls["Value"].sum() / workers) if workers else 0.0
    roll_sum = rolls.groupby(group_cols, observed=False)["Value"].sum()
    ma_sum = ma.groupby(group_cols, observed=False)["Value"].sum()
    return roll_sum.div(ma_sum.replace(0, float("nan"))).fillna(0).astype(float)


def compute_saegen_attainment(df: pd.DataFrame, target=SAEGEN_TARGET):
    """Return target attainment statistics for daily Sägen totals."""
    if df.empty:
        return {
            "hit_count": 0,
            "total_count": 0,
            "attainment_rate": 0.0,
            "worst_day": None,
            "best_day": None,
            "by_shift": pd.Series(dtype=float),
            "daily": pd.Series(dtype=float),
        }
    saegen = df[df["Metric"].str.contains("sägen", case=False, na=False)]
    daily = saegen.groupby("Datum", observed=False)["Value"].sum()
    daily = daily[daily > 0]
    by_shift = saegen.groupby("Schicht", observed=False)["Value"].sum()
    hit_count = int((daily >= target).sum())
    total_count = int(len(daily))
    return {
        "hit_count": hit_count,
        "total_count": total_count,
        "attainment_rate": float(hit_count / total_count) if total_count else 0.0,
        "worst_day": daily.idxmin() if total_count else None,
        "best_day": daily.idxmax() if total_count else None,
        "by_shift": by_shift,
        "daily": daily,
    }


def compute_staffing_gap_stats(df: pd.DataFrame):
    """Return distribution statistics for daily Differenz MA."""
    gap = df[df["Metric"].str.lower().str.strip() == "differenz ma"].groupby("Datum", observed=False)["Value"].sum()
    gap = gap[gap != 0]
    if gap.empty:
        return {"mean": 0.0, "std": 0.0, "min": 0.0, "max": 0.0, "pct_outside_band": 0.0, "daily": gap}
    return {
        "mean": float(gap.mean()),
        "std": float(gap.std(ddof=0)),
        "min": float(gap.min()),
        "max": float(gap.max()),
        "pct_outside_band": float((gap.abs() > 2).mean()),
        "daily": gap,
    }


def compute_task_time_allocation(df: pd.DataFrame, angaben_df: pd.DataFrame, minutes_col: str):
    """Compute hours and percent of total hours by task."""
    tasks = _roll_rows(df).groupby("Metric", observed=False)["Value"].sum().reset_index()
    if tasks.empty or minutes_col not in angaben_df.columns:
        return pd.DataFrame(columns=["Metric", "Hours", "Percent"])
    angaben = angaben_df.copy()
    angaben["Task"] = angaben["Task"].astype(str).str.strip().str.lower()
    merged = tasks.merge(angaben, left_on="Metric", right_on="Task", how="left")
    merged[minutes_col] = pd.to_numeric(merged[minutes_col], errors="coerce").fillna(0)
    merged["Hours"] = merged["Value"] * merged[minutes_col] / 60
    sonst = df[df["Metric"].str.lower().str.strip() == SONSTIGES].groupby("Metric", observed=False)["Value"].sum()
    if not sonst.empty:
        merged = pd.concat(
            [merged[["Metric", "Hours"]], pd.DataFrame({"Metric": ["Sonstiges"], "Hours": [float(sonst.sum())]})],
            ignore_index=True,
        )
    else:
        merged = merged[["Metric", "Hours"]]
    total = merged["Hours"].sum()
    merged["Percent"] = merged["Hours"] / total if total else 0
    return merged.sort_values("Hours", ascending=False)


def compute_volatility(series):
    """Return coefficient of variation for a numeric series."""
    s = pd.to_numeric(pd.Series(series), errors="coerce").dropna()
    mean = s.mean()
    return float(s.std(ddof=0) / mean) if mean else 0.0


def detect_anomalies(df, metric, method="iqr"):
    """Flag outlier days for a metric using IQR or z-score."""
    if isinstance(df, pd.Series):
        series = pd.to_numeric(df, errors="coerce").dropna()
    else:
        filtered = df[df["Metric"].str.lower().str.strip() == str(metric).lower()]
        series = filtered.groupby("Datum", observed=False)["Value"].sum()
    if series.empty:
        return pd.Series(dtype=bool)
    if method == "zscore":
        std = series.std(ddof=0)
        if std == 0:
            return pd.Series(False, index=series.index)
        return ((series - series.mean()).abs() / std) > 2
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series(False, index=series.index)
    return (series < q1 - 1.5 * iqr) | (series > q3 + 1.5 * iqr)


def classify_task_group(metric: str) -> str:
    """Map metric names to leadership task groups."""
    if not isinstance(metric, str):
        return "Sonstige"
    m = metric.strip().lower()
    if m.startswith("absetzen"):
        return "Absetzen"
    if m.startswith("richten"):
        return "Richten"
    if m.startswith("verladen"):
        return "Verladen"
    if m.startswith("zusammenfahren"):
        return "Zusammenfahren"
    return "Sonstige"


def _task_time_share_by_period(df: pd.DataFrame, angaben_df: pd.DataFrame,
                               minutes_col: str, period_col: str,
                               period_index: pd.DatetimeIndex) -> pd.DataFrame:
    """Return percent time allocation by task group for a weekly/monthly period."""
    if angaben_df is None or not minutes_col or minutes_col not in angaben_df.columns:
        return pd.DataFrame(index=period_index, columns=TASK_ORDER).fillna(0)

    task_values = _roll_rows(df).groupby([period_col, "Metric"], observed=False)["Value"].sum().reset_index()
    if task_values.empty:
        return pd.DataFrame(index=period_index, columns=TASK_ORDER).fillna(0)

    angaben = angaben_df.copy()
    angaben["Task"] = angaben["Task"].astype(str).str.strip().str.lower()
    task_hours = task_values.merge(angaben[["Task", minutes_col]], left_on="Metric", right_on="Task", how="left")
    task_hours[minutes_col] = pd.to_numeric(task_hours[minutes_col], errors="coerce").fillna(0)
    task_hours["Hours"] = task_hours["Value"] * task_hours[minutes_col] / 60
    task_hours["TaskGroup"] = task_hours["Metric"].apply(classify_task_group)
    task_hours = task_hours[[period_col, "TaskGroup", "Hours"]]

    sonstiges = (
        df[df["Metric"].eq(SONSTIGES)]
        .groupby(period_col, observed=False)["Value"]
        .sum()
        .reset_index(name="Hours")
    )
    if not sonstiges.empty:
        sonstiges["TaskGroup"] = "Sonstige"
        task_hours = pd.concat([task_hours, sonstiges[[period_col, "TaskGroup", "Hours"]]], ignore_index=True)

    hours = (
        task_hours.groupby([period_col, "TaskGroup"], observed=False)["Hours"].sum()
        .unstack("TaskGroup")
        .reindex(index=period_index, columns=TASK_ORDER)
        .fillna(0)
    )
    return hours.div(hours.sum(axis=1).replace(0, float("nan")), axis=0).mul(100).fillna(0).astype(float)


def _latest_full_week_start(df: pd.DataFrame, as_of_date=None):
    """Return the last eligible week start, excluding future and current partial weeks."""
    rolls = _roll_rows(df)
    active_rolls = (
        rolls.groupby("Datum", observed=False)["Value"]
        .sum()
        .loc[lambda s: s > 0]
    )
    if active_rolls.empty:
        return None

    latest_active_date = pd.Timestamp(active_rolls.index.max()).normalize()
    latest_active_week = latest_active_date - pd.Timedelta(days=latest_active_date.weekday())

    as_of = pd.Timestamp.today().normalize() if as_of_date is None else pd.Timestamp(as_of_date).normalize()
    current_week_start = as_of - pd.Timedelta(days=as_of.weekday())
    latest_completed_workweek = (
        current_week_start
        if as_of.weekday() >= 4
        else current_week_start - pd.Timedelta(days=7)
    )

    return min(latest_active_week, latest_completed_workweek)


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

    df = filter_summary(summary_long, start_of_week, end_of_week)
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

    # Exclude MA KPIs and loader's Sonstiges (hours-only)
    rolls = _roll_rows(df)

    rolls["TaskGroup"] = rolls["Metric"].apply(classify_task_group)

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
    differenz_ma_per_day = (
        df[df["Metric"].str.lower().str.strip() == "differenz ma"]
        .groupby("Datum", observed=False)["Value"].sum()
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
        "differenz_ma_per_day": differenz_ma_per_day,
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

    df = filter_summary(summary_long, target_day, target_day + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1))
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

    unmatched = shift_task_merged[
        shift_task_merged["Task"].isna()
        & ~shift_task_merged["Metric"].str.lower().str.strip().isin(MA_KPIS + [SONSTIGES])
        & (shift_task_merged["Value"] != 0)
    ]["Metric"].dropna().drop_duplicates().head(12).tolist()
    if unmatched:
        msg = (
            "Einige Aufgaben aus den Monatstabellen konnten nicht mit 'Angaben' gematcht werden. "
            "Die Stundenberechnung für diese Aufgaben ist 0. Bitte Schreibweise, Leerzeichen und Encoding/Umlaute prüfen: "
            + ", ".join(unmatched)
        )
        logger.warning(msg)
        add_message("warning", msg)

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


@st.cache_data(show_spinner=False)
def aggregate_monthly(summary_long: pd.DataFrame, target_year: int, target_month: int,
                      angaben_df: pd.DataFrame = None, minutes_col: str = None):
    """Aggregate all month-level KPI structures needed by the monthly view."""
    start = pd.Timestamp(year=int(target_year), month=int(target_month), day=1)
    end = start + pd.offsets.MonthEnd(0)
    df = filter_summary(summary_long, start, end)
    if df.empty:
        return {}
    workdays = pd.bdate_range(start, end)
    rolls = _roll_rows(df)
    total_rolls_per_day = rolls.groupby("Datum", observed=False)["Value"].sum().reindex(workdays).fillna(0)
    workers_per_day = (
        df[df["Metric"].eq("vorhandene ma")]
        .groupby("Datum", observed=False)["Value"].sum()
        .reindex(workdays).fillna(0)
    )
    needed_workers_per_day = (
        df[df["Metric"].eq("benötigte ma")]
        .groupby("Datum", observed=False)["Value"].sum()
        .reindex(workdays).fillna(0)
    )
    ma_by_day_shift = (
        df[df["Metric"].eq("vorhandene ma")]
        .groupby(["Datum", "Schicht"], observed=False)["Value"].sum()
        .unstack("Schicht")
        .reindex(index=workdays, columns=SHIFT_ORDER)
        .fillna(0)
    )
    if total_rolls_per_day.sum() == 0 and workers_per_day.sum() == 0:
        return {}
    rolls_per_shift = rolls.groupby("Schicht", observed=False)["Value"].sum()
    ma_per_shift = df[df["Metric"].eq("vorhandene ma")].groupby("Schicht", observed=False)["Value"].sum()
    rolls_per_ma_by_shift = (
        rolls_per_shift.div(ma_per_shift.replace(0, float("nan")))
        .fillna(0)
        .reindex(SHIFT_ORDER)
        .fillna(0)
        .astype(float)
    )
    saegen = compute_saegen_attainment(df)
    staffing = compute_staffing_gap_stats(df)
    task_hours = (
        compute_task_time_allocation(df, angaben_df, minutes_col)
        if angaben_df is not None and minutes_col
        else pd.DataFrame(columns=["Metric", "Hours", "Percent"])
    )
    rolls_anomaly = detect_anomalies(total_rolls_per_day, "total_rolls")
    rpm = total_rolls_per_day.div(workers_per_day.replace(0, float("nan"))).fillna(0).astype(float)
    rpm_anomaly = detect_anomalies(rpm, "rolls_per_ma")
    anomaly_rows = []
    for day, flagged in rolls_anomaly.items():
        if flagged:
            anomaly_rows.append({
                "Datum": day,
                "Wochentag": day.strftime("%A"),
                "Schicht": "Alle",
                "KPI": "Rollen gesamt",
                "Wert": float(total_rolls_per_day.loc[day]),
                "Abweichung vom Median": float(total_rolls_per_day.loc[day] - total_rolls_per_day.median()),
            })
    for day, flagged in rpm_anomaly.items():
        if flagged:
            anomaly_rows.append({
                "Datum": day,
                "Wochentag": day.strftime("%A"),
                "Schicht": "Alle",
                "KPI": "Rollen pro MA",
                "Wert": float(rpm.loc[day]),
                "Abweichung vom Median": float(rpm.loc[day] - rpm.median()),
            })
    anomalies = pd.DataFrame(anomaly_rows)
    return {
        "start": start,
        "end": end,
        "dates": workdays,
        "df": df,
        "total_rolls_per_day": total_rolls_per_day,
        "workers_per_day": workers_per_day,
        "needed_workers_per_day": needed_workers_per_day,
        "ma_by_day_shift": ma_by_day_shift,
        "rolls_per_ma_per_day": rpm,
        "rolls_per_ma_by_shift": rolls_per_ma_by_shift,
        "saegen": saegen,
        "staffing": staffing,
        "task_hours": task_hours,
        "anomalies": anomalies,
        "total_rolls": float(total_rolls_per_day.sum()),
        "workdays": int((total_rolls_per_day > 0).sum()),
        "volatility": compute_volatility(rpm[rpm > 0]),
    }


@st.cache_data(show_spinner=False)
def aggregate_longterm(summary_long: pd.DataFrame, angaben_df: pd.DataFrame = None,
                       minutes_col: str = None, as_of_date=None):
    """Aggregate weekly trend structures across the full available history."""
    df = filter_summary(summary_long)
    if df.empty:
        return {}
    df["WeekStart"] = df["Datum"] - pd.to_timedelta(df["Datum"].dt.weekday, unit="D")
    cutoff_week_start = _latest_full_week_start(df, as_of_date)
    if cutoff_week_start is None:
        return {}
    df = df[df["WeekStart"] <= cutoff_week_start].copy()
    if df.empty:
        return {}
    week_index = pd.date_range(df["WeekStart"].min(), df["WeekStart"].max(), freq="W-MON")
    df["MonthStart"] = df["Datum"].dt.to_period("M").dt.to_timestamp()
    month_index = pd.date_range(df["MonthStart"].min(), df["MonthStart"].max(), freq="MS")
    rolls = _roll_rows(df)
    total_rolls_week = (
        rolls.groupby("WeekStart", observed=False)["Value"]
        .sum()
        .reindex(week_index)
        .fillna(0)
        .sort_index()
    )
    workers_week = (
        df[df["Metric"].eq("vorhandene ma")]
        .groupby("WeekStart", observed=False)["Value"]
        .sum()
        .reindex(week_index)
        .fillna(0)
        .sort_index()
    )
    needed_workers_week = (
        df[df["Metric"].eq("benötigte ma")]
        .groupby("WeekStart", observed=False)["Value"]
        .sum()
        .reindex(week_index)
        .fillna(0)
        .sort_index()
    )
    rolls_per_ma = total_rolls_week.div(workers_week.replace(0, float("nan"))).fillna(0).astype(float)
    verladen_week = (
        df[df["Metric"].eq("verladen")]
        .groupby("WeekStart", observed=False)["Value"]
        .sum()
        .reindex(week_index)
        .fillna(0)
        .sort_index()
    )
    verladen_month = (
        df[df["Metric"].eq("verladen")]
        .groupby("MonthStart", observed=False)["Value"]
        .sum()
        .reindex(month_index)
        .fillna(0)
        .sort_index()
    )
    saegen_daily = (
        df[df["Metric"].str.contains("sägen", case=False, na=False)]
        .groupby(["WeekStart", "Datum"], observed=False)["Value"]
        .sum()
        .loc[lambda s: s > 0]
    )
    saegen_week = (
        saegen_daily.groupby("WeekStart", observed=False)
        .agg(["mean", "min", "max", "sum"])
        .reindex(week_index)
        .fillna(0)
    )
    task = rolls.copy()
    task["TaskGroup"] = task["Metric"].apply(classify_task_group)
    task_trends = (
        task.groupby(["WeekStart", "TaskGroup"], observed=False)["Value"].sum()
        .unstack("TaskGroup")
        .reindex(index=week_index, columns=TASK_ORDER)
        .fillna(0)
    )
    base = task_trends.replace(0, float("nan")).head(4).mean().replace(0, float("nan"))
    task_indexed = task_trends.div(base).mul(100).fillna(0).astype(float)
    task_time_share_week = _task_time_share_by_period(df, angaben_df, minutes_col, "WeekStart", week_index)
    task_time_share_month = _task_time_share_by_period(df, angaben_df, minutes_col, "MonthStart", month_index)
    return {
        "df": df,
        "weeks": total_rolls_week.index,
        "cutoff_week_start": cutoff_week_start,
        "total_rolls_week": total_rolls_week,
        "workers_week": workers_week,
        "needed_workers_week": needed_workers_week,
        "rolls_per_ma": rolls_per_ma,
        "verladen_week": verladen_week,
        "verladen_month": verladen_month,
        "saegen_week": saegen_week,
        "task_trends": task_trends,
        "task_indexed": task_indexed,
        "task_time_share_week": task_time_share_week,
        "task_time_share_month": task_time_share_month,
    }
