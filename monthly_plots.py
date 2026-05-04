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


def _day_axis(index):
    dates = pd.to_datetime(index)
    x = list(range(len(dates)))
    labels = [day.strftime("%d.%m") for day in dates]
    return x, labels


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

    rpm = monthly_data.get("rolls_per_ma_per_day", pd.Series(dtype=float))
    if rpm is not None and not rpm.empty:
        active_rpm = rpm[rpm > 0]
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x, labels = _day_axis(rpm.index)
        ax.plot(x, rpm.values, marker="o", linewidth=2.5, color="#1f77b4", label="Rollen/MA")
        if not active_rpm.empty:
            avg = float(active_rpm.mean())
            ax.axhline(avg, color="#222222", linestyle="--", linewidth=1.8, label=f"Monatsmittel {avg:.1f}")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45)
        ax.set_xlim(-0.5, len(rpm) - 0.5)
        ax.set_ylim(0, max(float(rpm.max()) * 1.15, 1))
        ax.set_title(f"Produktivität - Rollen pro MA ({label})")
        ax.set_ylabel("Rollen pro MA")
        ax.legend()
        title = f"Monatsanalyse - Produktivität Rollen pro MA ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

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

    needed_ma = monthly_data.get("needed_workers_per_day", pd.Series(dtype=float))
    available_ma = monthly_data.get("workers_per_day", pd.Series(dtype=float))
    if needed_ma is not None and available_ma is not None and not available_ma.empty:
        needed_ma = needed_ma.reindex(available_ma.index).fillna(0)
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x, labels = _day_axis(available_ma.index)
        available_values = available_ma.astype(float).values
        needed_values = needed_ma.astype(float).values
        ax.plot(x, needed_values, marker="o", linewidth=2.5, color="#d62728", label="Gerechnete MA")
        ax.plot(x, available_values, marker="o", linewidth=2.5, color="#2ca02c", label="Vorhandene MA")
        ax.fill_between(
            x,
            needed_values,
            available_values,
            where=available_values >= needed_values,
            interpolate=True,
            color="#2ca02c",
            alpha=0.18,
            label="Reserve",
        )
        ax.fill_between(
            x,
            needed_values,
            available_values,
            where=available_values < needed_values,
            interpolate=True,
            color="#d62728",
            alpha=0.20,
            label="Unterdeckung",
        )
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45)
        ax.set_xlim(-0.5, len(available_ma) - 0.5)
        ax.set_ylim(0, max(float(available_ma.max()), float(needed_ma.max()), 1) * 1.15)
        ax.set_title(f"Gerechnete MA vs. vorhandene MA ({label})")
        ax.set_ylabel("MA")
        ax.legend(ncol=2)
        title = f"Monatsanalyse - Gerechnete vs vorhandene MA ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    ma_by_day_shift = monthly_data.get("ma_by_day_shift")
    if ma_by_day_shift is not None and not ma_by_day_shift.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x, labels = _day_axis(ma_by_day_shift.index)
        bottom = pd.Series(0.0, index=ma_by_day_shift.index)
        for shift in ma_by_day_shift.columns:
            values = ma_by_day_shift[shift].astype(float)
            ax.bar(
                x,
                values.values,
                bottom=bottom.values,
                label=shift,
                color=SHIFT_COLORS.get(shift, "#cccccc"),
            )
            bottom = bottom.add(values, fill_value=0)
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45)
        ax.set_xlim(-0.5, len(ma_by_day_shift.index) - 0.5)
        ax.set_ylim(0, max(float(bottom.max()) * 1.15, 1))
        ax.set_title(f"MA pro Tag nach Schicht ({label})")
        ax.set_ylabel("Vorhandene MA")
        ax.legend(ncol=3)
        title = f"Monatsanalyse - MA pro Tag ({label})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    return figs_meta
