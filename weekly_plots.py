import matplotlib.pyplot as plt
from helpers import format_day_month, sanitize_filename
from config import (
    SHIFT_ORDER,
    SHIFT_COLORS,
    DPI,
    SHOW_AVERAGE_LINE,
    AVERAGE_LINE_STYLE,
    AVERAGE_LABEL_FMT,
    SAEGEN_TARGET,
    LEGEND_LOC,
    LEGEND_NCOL,
)

# Fixed order for Hauptaufgaben + Sonstiges
TASK_ORDER = ["Absetzen", "Richten", "Verladen", "Zusammenfahren", "Sonstige"]


def _add_average_line(ax, values, x_positions, ymax, label_prefix="Durchschnitt"):
    """Helper to add average line + label if enabled in config."""
    if not SHOW_AVERAGE_LINE:
        return

    valid = values[values > 0]
    if valid.empty:
        return

    avg = valid.mean()
    ax.axhline(avg, **AVERAGE_LINE_STYLE)
    ax.text(
        x_positions[-1],
        avg + ymax * 0.02,
        AVERAGE_LABEL_FMT.format(avg),
        color=AVERAGE_LINE_STYLE.get("color", "blue"),
        fontsize=12,
        fontweight="bold",
        ha="left",
        va="bottom",
    )


def plot_weekly_charts(weekly_data: dict, target_date):
    """
    Create weekly charts from pre-aggregated data.

    Args:
        weekly_data: dict returned by aggregate_weekly()
        target_date: datetime for labeling (inside the week)

    Returns:
        List of dicts {title, filename, fig}
    """
    figs_meta = []
    dates = weekly_data["dates"]
    kw = weekly_data["kw"]   # <-- calendar week

    # ------------------------
    # 1) Sägen per day by shift
    # ------------------------
    saegen_by_day_shift = weekly_data["saegen_by_day_shift"]
    if not saegen_by_day_shift.empty:
        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)
        saegen_by_day_shift.plot(
            kind="bar",
            stacked=True,
            ax=ax,
            color=[SHIFT_COLORS[s] for s in SHIFT_ORDER],
        )
        ax.set_title(f"Anzahl gesägte Rollen pro Tag (KW {kw})")
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(format_day_month(dates), rotation=45)

        ymax = max(saegen_by_day_shift.sum(axis=1).max(), SAEGEN_TARGET) * 1.2
        ax.set_ylim(0, ymax)

       # Ziel line + annotation
        ax.axhline(SAEGEN_TARGET, color="red", linestyle="--", linewidth=2)

        # place annotation in the middle of the x-axis, just above the line
        x_pos = 2.5
        ax.text(
            x_pos, SAEGEN_TARGET + ymax * 0.02,
            f"Ziel {SAEGEN_TARGET}",
            color="red",
            fontsize=10,
            fontweight="bold",
            ha="center", va="bottom",
            bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.7, ec="red")
        )

        # Average line + annotation
        values = saegen_by_day_shift.sum(axis=1)
        valid = values[values > 0]
        if not valid.empty:
            avg = valid.mean()
            ax.axhline(avg, color="blue", linestyle="--", linewidth=2)

            ax.text(
                x_pos - 1,  # shift half a column to the right to avoid overlap with Ziel
                avg + ymax * 0.02,
                f"Ø {int(round(avg))}",
                color="blue",
                fontsize=10,
                fontweight="bold",
                ha="center", va="bottom",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.7, ec="blue")
            )

        # Legend (only shifts, placed bottom center)
        handles, labels = ax.get_legend_handles_labels()
        shift_handles = [h for h, l in zip(handles, labels) if l in SHIFT_ORDER]
        shift_labels = [l for l in labels if l in SHIFT_ORDER]
        
        # Legend inside, upper center
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 0.98),
            ncol=5,
            frameon=True,
            fontsize=9
        )

        # Numbers inside bars
        for i, day in enumerate(saegen_by_day_shift.index):
            bottom = 0
            total = 0
            for shift in SHIFT_ORDER:
                value = saegen_by_day_shift.loc[day, shift]
                total += value
                if value > 0:
                    ax.text(
                        i,
                        bottom + value / 2,
                        f"{int(round(value))}",
                        ha="center",
                        va="center",
                        color="white",
                        fontsize=10,
                        fontweight="bold",
                    )
                bottom += value
            if total > 0:
                ax.text(
                    i,
                    min(total + ymax * 0.03, ymax * 0.97),
                    f"{int(round(total))}",
                    ha="center",
                    va="bottom",
                    color="black",
                    fontsize=12,
                    fontweight="bold",
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="black")
                )

        plt.tight_layout()
        title = f"Wochenanalyse – Sägen (KW {kw})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})
    
    # ------------------------
    # 2) Total rolls per day by task group (Hauptaufgaben + Sonstige)
    # ------------------------
    total_rolls_by_group = weekly_data["total_rolls_by_group"]
    if not total_rolls_by_group.empty:
        # enforce fixed order
        total_rolls_by_group = total_rolls_by_group.reindex(columns=TASK_ORDER, fill_value=0)

        # fixed colors (same order as TASK_ORDER)
        TASK_COLORS = {
            "Absetzen": "#1f77b4",       # blue
            "Richten": "#ff7f0e",        # orange
            "Verladen": "#2ca02c",       # green
            "Zusammenfahren": "#8c564b", # brown
            "Sonstige": "#888888",       # grey
        }

        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)
        total_rolls_by_group.plot(
            kind="bar",
            stacked=True,
            ax=ax,
            color=[TASK_COLORS[t] for t in TASK_ORDER],
        )

        ax.set_title(f"Gesamte Anzahl bewegter Rollen nach Aufgaben (KW {kw})")
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(format_day_month(dates), rotation=45)

        # Scale a bit lower
        ymax = max(total_rolls_by_group.sum(axis=1).max(), 1) * 1.2
        ax.set_ylim(0, ymax)

        # Legend inside, upper center
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 0.98),
            ncol=5,
            frameon=True,
            fontsize=9
        )

        # Average line
        values = total_rolls_by_group.sum(axis=1)
        valid = values[values > 0]
        if not valid.empty:
            avg = valid.mean()
            ax.axhline(avg, color="blue", linestyle="--", linewidth=2)

            # Annotation in box inside upper center
            ax.text(
                x_pos - 1,  # shift half a column to the right to avoid overlap with Ziel
                avg + ymax * 0.02,
                f"Ø {int(round(avg))}",
                color="blue",
                fontsize=10,
                fontweight="bold",
                ha="center", va="bottom",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.7, ec="blue")
            )

        # numbers inside bars
        for i, day in enumerate(dates):
            bottom = 0
            for task in TASK_ORDER:
                value = total_rolls_by_group.loc[day, task]
                if value > 0:
                    ax.text(
                        i,
                        bottom + value / 2,
                        f"{int(round(value))}",
                        ha="center",
                        va="center",
                        color="white",
                        fontsize=10,
                        fontweight="bold",
                    )
                bottom += value
            total = total_rolls_by_group.loc[day].sum()
            ax.text(
                i,
                min(total + ymax * 0.03, ymax * 0.97),
                f"{int(round(total))}",
                ha="center",
                va="bottom",
                color="black",
                fontsize=12,
                fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="black")
            )
    # --- NEW: small box to the right with MA available ---
            ma_val = weekly_data["workers_per_day"].loc[day] if day in weekly_data["workers_per_day"].index else 0
            ma_str = f"MA: {int(ma_val)}" if ma_val == int(ma_val) else f"MA: {ma_val:.1f}"

            ax.text(
                i + 0.27,   # shift right of the bar
                bottom / 2, # place vertically in middle of bar
                ma_str,
                ha="left",
                va="center",
                fontsize=8,
                color="black",
                bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="gray", alpha=0.8)
            )

        plt.tight_layout()
        title = f"Wochenanalyse – Rollen pro Tag (KW {kw})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    # ------------------------
    # 3) Total rolls per day by shift
    # ------------------------
    total_shift = weekly_data["total_shift"]
    if not total_shift.empty:
        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)
        total_shift.plot(
            kind="bar",
            stacked=True,
            ax=ax,
            color=[SHIFT_COLORS[s] for s in SHIFT_ORDER],
        )

        ax.set_title(f"Gesamte Anzahl bewegter Rollen nach Schicht (KW {kw})")
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(format_day_month(dates), rotation=45)

        # Scale a bit lower
        ymax = max(total_shift.sum(axis=1).max(), 1) * 1.2
        ax.set_ylim(0, ymax)

        # Legend inside, upper center
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 0.98),
            ncol=3,
            frameon=True,
            fontsize=9
        )

        # Average line + annotation
        values = total_shift.sum(axis=1)
        valid = values[values > 0]
        if not valid.empty:
            avg = valid.mean()
            ax.axhline(avg, color="blue", linestyle="--", linewidth=2)

            x_pos = 1.5  # between 2nd and 3rd bar
            ax.text(
                x_pos,
                avg + ymax * 0.02,
                f"Ø {int(round(avg))}",
                color="blue",
                fontsize=10,
                fontweight="bold",
                ha="center", va="bottom",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="blue")
            )

        # Numbers inside bars + MA box per shift
        for i, day in enumerate(dates):
            bottom = 0
            for shift in SHIFT_ORDER:
                value = total_shift.loc[day, shift]
                if value > 0:
                    # Rolls number inside bar
                    ax.text(
                        i,
                        bottom + value / 2,
                        f"{int(round(value))}",
                        ha="center",
                        va="center",
                        color="white",
                        fontsize=10,
                        fontweight="bold",
                    )

                    # MA box next to this shift segment
                    if "workers_per_shift" in weekly_data:
                        try:
                            ma_val = weekly_data["workers_per_shift"].loc[(day, shift)]
                            ma_str = f"MA: {int(ma_val)}" if ma_val == int(ma_val) else f"MA: {ma_val:.1f}"

                            ax.text(
                                i + 0.3,  # small horizontal offset to the right of bar
                                bottom + value / 2,  # align with block center
                                ma_str,
                                ha="left", va="center",
                                fontsize=8,
                                color="black",
                                bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="gray")
                            )
                        except KeyError:
                            pass  # if no MA data available

                bottom += value

            # Total rolls on top of the stack
            total = total_shift.loc[day].sum()
            ax.text(
                i,
                min(total + ymax * 0.03, ymax * 0.97),
                f"{int(round(total))}",
                ha="center",
                va="bottom",
                color="black",
                fontsize=12,
                fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="black")
            )

        plt.tight_layout()
        title = f"Wochenanalyse – Rollen nach Schicht (KW {kw})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

        
  # ------------------------
    # 4) Total rolls per MA
    # ------------------------
    workers_per_day = weekly_data["workers_per_day"]
    total_rolls_per_day = weekly_data["total_rolls_per_day"]
    if (workers_per_day > 0).any() and (total_rolls_per_day > 0).any():
        norm = (total_rolls_per_day / workers_per_day).fillna(0)

        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)
        ax.bar(range(len(norm.index)), norm.values, color="#1f77b4", width=0.5)

        ax.set_title(f"Gesamtrollen pro MA (KW {kw})")
        ax.set_xticks(range(len(norm.index)))
        ax.set_xticklabels(format_day_month(norm.index), rotation=45)

        # Scale consistent with others
        ymax = max(norm.values) * 1.2 if len(norm) > 0 else 1
        ax.set_ylim(0, ymax)

        # Average line + annotation
        values = norm
        valid = values[values > 0]
        if not valid.empty:
            avg = valid.mean()
            ax.axhline(avg, color="blue", linestyle="--", linewidth=2)

            x_pos = 1.5  # between 2nd and 3rd bar
            ax.text(
                x_pos,
                avg + ymax * 0.02,
                f"Ø {int(round(avg))}",
                color="blue",
                fontsize=10,
                fontweight="bold",
                ha="center", va="bottom",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.7, ec="blue")
            )

        # Numbers above bars
        for i, v in enumerate(norm.values):
            ax.text(
                i,
                v + ymax * 0.03,
                f"{int(round(v))}",
                ha="center",
                va="bottom",
                color="black",
                fontsize=12,
                fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="black")
            )

        plt.tight_layout()
        title = f"Wochenanalyse – Rollen pro MA (KW {kw})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})
    
 # ------------------------
    # 5) Anzahl MA pro Schicht
    # ------------------------
    ma_by_shift = weekly_data["ma_by_shift"]
    if not ma_by_shift.empty:
        fig, ax = plt.subplots(figsize=(12, 6), dpi=DPI)
        ma_by_shift.plot(
            kind="bar",
            stacked=True,
            ax=ax,
            color=[SHIFT_COLORS[s] for s in SHIFT_ORDER],
        )

        ax.set_title(f"Anzahl MA pro Schicht (KW {kw})")
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(format_day_month(dates), rotation=45)

        # Scale similar to others
        ymax = max(ma_by_shift.sum(axis=1).max(), 1) * 1.2
        ax.set_ylim(0, ymax)

        # Average line + annotation
        values = ma_by_shift.sum(axis=1)
        valid = values[values > 0]
        if not valid.empty:
            avg = valid.mean()
            ax.axhline(avg, color="blue", linestyle="--", linewidth=2)

            x_pos = 1.5  # place between 2nd and 3rd bar
            ax.text(
                x_pos,
                avg + ymax * 0.02,
                f"Ø {int(round(avg))}",
                color="blue",
                fontsize=10,
                fontweight="bold",
                ha="center", va="bottom",
                bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.7, ec="blue")
            )

        # Numbers inside bars + totals
        for i, day in enumerate(dates):
            bottom = 0
            total = 0
            for shift in SHIFT_ORDER:
                value = ma_by_shift.loc[day, shift]
                total += value
                if value > 0:
                    ax.text(
                        i,
                        bottom + value / 2,
                        f"{int(round(value))}",
                        ha="center",
                        va="center",
                        color="white",
                        fontsize=9,
                        fontweight="bold",
                    )
                bottom += value
            if total > 0:
                ax.text(
                    i,
                    min(total + ymax * 0.03, ymax * 0.97),
                    f"{int(round(total))}",
                    ha="center",
                    va="bottom",
                    color="black",
                    fontsize=11,
                    fontweight="bold",
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="black")
                )

        # Legend inside top center
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, 0.98),
            ncol=5,
            frameon=True,
            fontsize=9
        )

        plt.tight_layout()
        title = f"Wochenanalyse – MA nach Schicht (KW {kw})"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})


    return figs_meta
