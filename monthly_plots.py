import calendar

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.patches import Rectangle

from config import DPI, SAEGEN_TARGET, SHIFT_COLORS
from helpers import sanitize_filename


def _previous_month_percent_lookup(previous_month_data: dict | None):
    if not previous_month_data:
        return {}
    previous = previous_month_data.get("task_hours")
    if previous is None or previous.empty:
        return {}
    return dict(zip(previous["Metric"], previous["Percent"]))


def plot_monthly_charts(monthly_data: dict, previous_month_data: dict | None = None):
    """Create monthly leadership charts from aggregate_monthly output."""
    figs_meta = []
    if not monthly_data:
        return figs_meta

    start = monthly_data["start"]
    year = int(start.year)
    month = int(start.month)
    label = start.strftime("%Y-%m")
    previous_percent = _previous_month_percent_lookup(previous_month_data)

    daily_saegen = monthly_data["saegen"].get("daily", pd.Series(dtype=float))
    if not daily_saegen.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        ax.set_title(f"Sägen-Zielerreichung im Monat {label}")
        ax.set_xlim(0, 7)
        ax.set_ylim(0, 6)
        ax.axis("off")
        month_calendar = calendar.Calendar(firstweekday=0).monthdatescalendar(year, month)
        for row_idx, week in enumerate(month_calendar):
            for col_idx, day in enumerate(week):
                if day.month != month or day.weekday() >= 5:
                    color = "#f2f2f2"
                    text = ""
                else:
                    ts = pd.Timestamp(day)
                    value = float(daily_saegen.get(ts, 0))
                    if value >= SAEGEN_TARGET:
                        color = "#2ca02c"
                    elif value >= SAEGEN_TARGET * 0.8:
                        color = "#ffbf00"
                    else:
                        color = "#d62728"
                    text = f"{day.day}\n{int(value)}"
                y = 5 - row_idx
                ax.add_patch(Rectangle((col_idx, y), 0.95, 0.85, facecolor=color, edgecolor="white"))
                ax.text(col_idx + 0.48, y + 0.42, text, ha="center", va="center", fontsize=10, fontweight="bold")
        hit = monthly_data["saegen"]["hit_count"]
        total = monthly_data["saegen"]["total_count"]
        rate = monthly_data["saegen"]["attainment_rate"] * 100
        ax.text(0, -0.25, f"{hit} von {total} Arbeitstagen haben das Ziel erreicht ({rate:.0f}%).", fontsize=11)
        title = f"Monatsanalyse - Sägen Zielerreichung ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    task_hours = monthly_data.get("task_hours")
    if task_hours is not None and not task_hours.empty and task_hours["Hours"].sum() > 0:
        top = task_hours.head(8).copy()
        if len(task_hours) > 8:
            other_hours = task_hours.iloc[8:]["Hours"].sum()
            top = pd.concat([top, pd.DataFrame({"Metric": ["Sonstige"], "Hours": [other_hours]})], ignore_index=True)
        fig, ax = plt.subplots(figsize=(8, 6), dpi=DPI)
        ax.pie(top["Hours"], labels=top["Metric"].str.title(), autopct="%1.0f%%", startangle=90, pctdistance=0.8)
        ax.add_artist(plt.Circle((0, 0), 0.55, fc="white"))
        ax.set_title(f"Zeitanteile nach Aufgabe ({label})")
        if previous_percent:
            lines = []
            for _, row in task_hours.head(5).iterrows():
                metric = row["Metric"]
                delta_pp = (row["Percent"] - previous_percent.get(metric, 0)) * 100
                lines.append(f"{metric.title()}: {delta_pp:+.1f} pp")
            ax.text(
                1.15,
                0.5,
                "vs. Vormonat\n" + "\n".join(lines),
                transform=ax.transAxes,
                va="center",
                fontsize=9,
                bbox=dict(boxstyle="round,pad=0.35", fc="white", ec="#888888", alpha=0.9),
            )
        title = f"Monatsanalyse - Zeitanteile ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    staffing = monthly_data.get("staffing", {})
    gap = staffing.get("daily", pd.Series(dtype=float))
    if not gap.empty:
        fig, ax = plt.subplots(figsize=(10, 5), dpi=DPI)
        ax.hist(gap.values, bins=min(10, max(3, len(gap))), color="#1f77b4", edgecolor="white")
        ax.axvline(0, color="black", linewidth=2, label="Ziel")
        ax.axvline(-2, color="red", linestyle="--", label="-2 MA")
        ax.axvline(2, color="green", linestyle="--", label="+2 MA")
        in_band = (gap.abs() <= 2).mean() * 100
        ax.set_title(f"Verteilung Δ MA ({label}) - {in_band:.0f}% im akzeptablen Bereich")
        ax.set_xlabel("Differenz MA")
        ax.set_ylabel("Arbeitstage")
        ax.legend()
        title = f"Monatsanalyse - Differenz MA ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    rpm_shift = monthly_data.get("rolls_per_ma_by_shift")
    if rpm_shift is not None and not rpm_shift.empty:
        rpm_shift = rpm_shift.sort_values(ascending=True)
        fig, ax = plt.subplots(figsize=(9, 5), dpi=DPI)
        colors = [SHIFT_COLORS.get(shift, "#cccccc") for shift in rpm_shift.index]
        ax.barh(range(len(rpm_shift)), rpm_shift.values, color=colors)
        ax.set_yticks(range(len(rpm_shift)))
        ax.set_yticklabels(rpm_shift.index)
        ax.set_xlim(0, max(rpm_shift.max() * 1.15, 1))
        ax.set_title(f"Rollen pro MA nach Schicht ({label})")
        ax.set_xlabel("Rollen pro MA")
        for i, value in enumerate(rpm_shift.values):
            ax.text(value, i, f" {value:.1f}", va="center", fontweight="bold")
        title = f"Monatsanalyse - Rollen pro MA nach Schicht ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    return figs_meta
