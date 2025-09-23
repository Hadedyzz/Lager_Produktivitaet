# daily_plots.py
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
import pandas as pd

# Silence pandas warnings and info messages
pd.options.mode.chained_assignment = None
pd.set_option("mode.chained_assignment", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.set_option("display.width", None)
pd.set_option("display.expand_frame_repr", False)

from helpers import sanitize_filename
from config import (
    SHIFT_COLORS,
    DPI,
    LEGEND_LOC,
    LEGEND_NCOL,
)


def plot_daily_charts(daily_data: dict, target_day):
    """
    Create daily charts from pre-aggregated data.

    Args:
        daily_data: dict returned by aggregate_daily()
        target_day: datetime for labeling

    Returns:
        List of dicts {title, filename, fig}
    """
    figs_meta = []
    if not daily_data:
        return figs_meta

    # Pull pre-aggregated pivots
    hours_pivot = daily_data.get("hours_pivot")
    rolls_pivot = daily_data.get("rolls_pivot")
    shift_task_merged = daily_data.get("shift_task_merged")

    if hours_pivot is None or rolls_pivot is None or shift_task_merged is None:
        return figs_meta

    # ---------------------------------------
    # 1) Arbeitsstunden pro Aufgabe und Schicht
    # ---------------------------------------
    if not hours_pivot.empty:
        # Order by total hours (descending)
        order_hours = hours_pivot.sum(axis=1).sort_values(ascending=False).index
        stacked_df = hours_pivot.loc[order_hours]

        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)

        # Totals
        shift_fte_totals = stacked_df.sum(axis=0) / 7.5
        total_fte = round(shift_fte_totals.sum(), 1)
        total_hours = stacked_df.sum().sum()

        # Plot
        shift_colors = [SHIFT_COLORS.get(shift, "#cccccc") for shift in stacked_df.columns]
        stacked_df.plot(kind="bar", stacked=True, ax=ax, color=shift_colors)

        # Set ymax
        max_value = stacked_df.sum(axis=1).max()
        ax.set_ylim(top=1.1 * max_value)

        # Annotate totals and FTEs inside bars
        for i, task in enumerate(stacked_df.index):
            total_h = stacked_df.loc[task].sum()
            ax.text(
                i, total_h + stacked_df.values.max() * 0.02,
                f"{total_h:.1f}h",
                ha="center", va="bottom", fontsize=9
            )
            bottom = 0
            for shift in stacked_df.columns:
                h = stacked_df.loc[task, shift]
                if h > 0:
                    fte = h / 7.5
                    ax.text(
                        i, bottom + h / 2, f"{fte:.1f} MA",
                        ha="center", va="center", fontsize=9,
                        bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="black", alpha=0.8)
                    )
                    bottom += h

        # Info box with Gesamt + per-shift MA
        box_lines = [
            f"Gesamtstunden: {total_hours:.1f}",
            f"Gesamt MA (7,5h): {total_fte:.1f}",
            "------------------------",
            "Gerechnete / Vorhandene MA",
        ]

        for shift in stacked_df.columns:
            benoetigt = shift_fte_totals[shift]
            vorhanden = (
                shift_task_merged.loc[
                    (shift_task_merged["Schicht"] == shift)
                    & (shift_task_merged["Metric"].str.lower().str.strip() == "vorhandene ma"),
                    "Value",
                ].sum()
            )
            vorhanden_str = f"{int(vorhanden)}" if vorhanden == int(vorhanden) else f"{vorhanden:.1f}"
            box_lines.append(f"{shift}: {benoetigt:.1f} / {vorhanden_str}")

        box_text = "\n".join(box_lines)
        props = dict(boxstyle="round", facecolor="white", alpha=0.8)
        ax.text(
            0.98, 0.98, box_text,
            transform=ax.transAxes,
            fontsize=11,
            verticalalignment="top",
            horizontalalignment="right",
            bbox=props,
        )

        ax.tick_params(axis="x", labelrotation=60, labelsize=9)
        ax.set_title(f"Arbeitsstunden pro Aufgabe und Schicht am {target_day.strftime('%d.%m.%Y')}")
        ax.set_xlabel("Aufgabe")
        ax.set_ylabel("Stunden")
        ax.legend(loc=LEGEND_LOC, ncol=LEGEND_NCOL, frameon=False)

        plt.tight_layout()
        title = f"Tagesanalyse – Arbeitsstunden (am {target_day.strftime('%d-%m-%Y')})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    # ---------------------------------------
    # 2) Rollen pro Aufgabe und Schicht
    # ---------------------------------------
    if not rolls_pivot.empty:
        # Order by total rolls (descending)
        order_rolls = rolls_pivot.sum(axis=1).sort_values(ascending=False).index
        stacked_rolls_df = rolls_pivot.loc[order_rolls]

        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)

        shift_colors_rolls = [SHIFT_COLORS.get(shift, "#cccccc") for shift in stacked_rolls_df.columns]
        stacked_rolls_df.plot(kind="bar", stacked=True, ax=ax, color=shift_colors_rolls)

        # Set ymax
        max_value = stacked_rolls_df.sum(axis=1).max()
        ax.set_ylim(top=1.1 * max_value)

        # Annotate totals and counts
        for i, task in enumerate(stacked_rolls_df.index):
            total_rolls = stacked_rolls_df.loc[task].sum()
            ax.text(
                i, total_rolls + stacked_rolls_df.values.max() * 0.02,
                f"{int(total_rolls)}", ha="center", va="bottom", fontsize=9
            )
            bottom = 0
            for shift in stacked_rolls_df.columns:
                rolls = stacked_rolls_df.loc[task, shift]
                if rolls > 0:
                    ax.text(
                        i, bottom + rolls / 2, f"{int(rolls)}",
                        ha="center", va="center", fontsize=9,
                        bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="black", alpha=0.8)
                    )
                    bottom += rolls

        ax.tick_params(axis="x", labelrotation=60, labelsize=9)
        total_rolls_all = int(stacked_rolls_df.values.sum())
        ax.set_title(f"Rollen pro Aufgabe und Schicht am {target_day.strftime('%d.%m.%Y')} (Total: {total_rolls_all})")
        ax.set_xlabel("Aufgabe")
        ax.set_ylabel("Rollen")

        legend_patches = [
            Patch(facecolor=SHIFT_COLORS.get(shift, "#cccccc"), edgecolor="black", label=shift)
            for shift in stacked_rolls_df.columns
        ]
        ax.legend(handles=legend_patches, loc=LEGEND_LOC, ncol=LEGEND_NCOL, frameon=False)

        plt.tight_layout()
        title = f"Tagesanalyse – Rollen (am {target_day.strftime('%d-%m-%Y')})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    return figs_meta
