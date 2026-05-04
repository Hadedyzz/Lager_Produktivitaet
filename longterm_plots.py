import matplotlib.pyplot as plt

from config import DPI, SAEGEN_TARGET, SHIFT_COLORS
from helpers import sanitize_filename


def _week_labels(index):
    return [f"KW {int(day.isocalendar().week)}" for day in index]


def plot_longterm_charts(longterm_data: dict):
    """Create long-term leadership charts from aggregate_longterm output."""
    figs_meta = []
    if not longterm_data:
        return figs_meta

    rpm = longterm_data["rolls_per_ma"]
    if not rpm.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = range(len(rpm))
        ax.plot(x, rpm.values, marker="o", label="Rollen/MA")
        ax.plot(x, rpm.rolling(4, min_periods=1).mean().values, linewidth=3, label="4-Wochen-Mittel")
        ax.set_xticks(x)
        ax.set_xticklabels(_week_labels(rpm.index), rotation=45)
        ax.set_xlim(-0.5, len(rpm) - 0.5)
        ax.set_title("Produktivitätstrend")
        ax.set_ylabel("Rollen pro MA")
        if len(rpm):
            best = rpm.idxmax()
            worst = rpm.idxmin()
            ax.annotate("Beste", (list(rpm.index).index(best), rpm.loc[best]), textcoords="offset points", xytext=(0, 10), ha="center")
            ax.annotate("Schlechteste", (list(rpm.index).index(worst), rpm.loc[worst]), textcoords="offset points", xytext=(0, -15), ha="center")
        ax.legend()
        title = "Langzeit - Produktivitätstrend"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    saegen = longterm_data["saegen_week"]
    if not saegen.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = range(len(saegen))
        ax.plot(x, saegen["mean"].values, marker="o", label="Mittelwert")
        ax.fill_between(x, saegen["min"].values, saegen["max"].values, alpha=0.2, label="Min/Max")
        ax.axhline(SAEGEN_TARGET, color="red", linestyle="--", label=f"Ziel {SAEGEN_TARGET}")
        ax.set_xticks(x)
        ax.set_xticklabels(_week_labels(saegen.index), rotation=45)
        ax.set_xlim(-0.5, len(saegen) - 0.5)
        ax.set_title("Sägen-Performance über Zeit")
        ax.set_ylabel("Rollen")
        ax.legend()
        title = "Langzeit - Sägen Performance"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    staffing = longterm_data["staffing_by_shift"]
    if not staffing.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = range(len(staffing))
        ax.stackplot(x, [staffing[col].values for col in staffing.columns], labels=staffing.columns, colors=[SHIFT_COLORS.get(s) for s in staffing.columns])
        ax.set_xticks(x)
        ax.set_xticklabels(_week_labels(staffing.index), rotation=45)
        ax.set_xlim(-0.5, len(staffing) - 0.5)
        ax.set_title("Personalbedarf nach Schicht")
        ax.set_ylabel("Benötigte MA")
        ax.legend(loc="upper left")
        title = "Langzeit - Personalbedarf"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    task_indexed = longterm_data["task_indexed"]
    if not task_indexed.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = range(len(task_indexed))
        for col in task_indexed.columns:
            ax.plot(x, task_indexed[col].values, marker="o", label=col)
        ax.axhline(100, color="black", linewidth=1)
        ax.set_xticks(x)
        ax.set_xticklabels(_week_labels(task_indexed.index), rotation=45)
        ax.set_xlim(-0.5, len(task_indexed) - 0.5)
        ax.set_title("Aufgabenvolumen indexiert (Start = 100)")
        ax.set_ylabel("Index")
        ax.legend(ncol=3)
        title = "Langzeit - Aufgabenvolumen"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    shift_share = longterm_data["shift_share"]
    if not shift_share.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = range(len(shift_share))
        ax.stackplot(x, [shift_share[col].values * 100 for col in shift_share.columns], labels=shift_share.columns, colors=[SHIFT_COLORS.get(s) for s in shift_share.columns])
        ax.set_xticks(x)
        ax.set_xticklabels(_week_labels(shift_share.index), rotation=45)
        ax.set_xlim(-0.5, len(shift_share) - 0.5)
        ax.set_ylim(0, 100)
        ax.set_title("Schichtstruktur MA-Anteil")
        ax.set_ylabel("Anteil %")
        ax.legend(loc="upper left")
        title = "Langzeit - Schichtstruktur"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    heatmap = longterm_data["volatility_heatmap"]
    if not heatmap.empty:
        fig, ax = plt.subplots(figsize=(10, 6), dpi=DPI)
        image = ax.imshow(heatmap.values, aspect="auto", cmap="YlOrRd")
        ax.set_yticks(range(len(heatmap.index)))
        ax.set_yticklabels(_week_labels(heatmap.index))
        ax.set_xticks(range(len(heatmap.columns)))
        ax.set_xticklabels(heatmap.columns, rotation=45)
        ax.set_title("Volatilität Rollen/MA nach Wochentag")
        fig.colorbar(image, ax=ax, label="Variationskoeffizient")
        title = "Langzeit - Volatilität Heatmap"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    return figs_meta
