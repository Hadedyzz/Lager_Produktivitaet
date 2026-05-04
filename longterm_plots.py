import matplotlib.pyplot as plt

from config import DPI, SAEGEN_TARGET
from helpers import sanitize_filename


TASK_COLORS = {
    "Absetzen": "#1f77b4",
    "Richten": "#ff7f0e",
    "Verladen": "#2ca02c",
    "Zusammenfahren": "#9467bd",
    "Sonstige": "#7f7f7f",
}


def _week_labels(index):
    return [f"KW {int(day.isocalendar().week)}" for day in index]


def _month_labels(index):
    return [day.strftime("%Y-%m") for day in index]


def _add_line_axis_labels(ax, x, labels):
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45)
    ax.set_xlim(-0.5, len(labels) - 0.5)


def _plot_task_time_share(df, labels, title):
    fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
    x = list(range(len(df)))
    columns = [col for col in df.columns if df[col].sum() > 0]
    ax.stackplot(
        x,
        [df[col].values for col in columns],
        labels=columns,
        colors=[TASK_COLORS.get(col, "#cccccc") for col in columns],
    )
    _add_line_axis_labels(ax, x, labels)
    ax.set_ylim(0, 100)
    ax.set_title(title)
    ax.set_ylabel("Zeitanteil %")
    ax.legend(loc="upper left", ncol=3)
    return fig


def plot_longterm_charts(longterm_data: dict):
    """Create long-term leadership charts from aggregate_longterm output."""
    figs_meta = []
    if not longterm_data:
        return figs_meta

    rpm = longterm_data["rolls_per_ma"]
    if not rpm.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = list(range(len(rpm)))
        ax.plot(x, rpm.values, marker="o", label="Rollen/MA")
        ax.plot(x, rpm.rolling(4, min_periods=1).mean().values, linewidth=3, label="4-Wochen-Mittel")
        _add_line_axis_labels(ax, x, _week_labels(rpm.index))
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
        x = list(range(len(saegen)))
        ax.plot(x, saegen["mean"].values, marker="o", label="Mittelwert")
        ax.fill_between(x, saegen["min"].values, saegen["max"].values, alpha=0.2, label="Min/Max")
        ax.axhline(SAEGEN_TARGET, color="red", linestyle="--", label=f"Ziel {SAEGEN_TARGET}")
        _add_line_axis_labels(ax, x, _week_labels(saegen.index))
        ax.set_title("Sägen-Performance über Zeit")
        ax.set_ylabel("Rollen")
        ax.legend()
        title = "Langzeit - Sägen Performance"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    verladen_week = longterm_data.get("verladen_week")
    if verladen_week is not None and not verladen_week.empty:
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = list(range(len(verladen_week)))
        ax.bar(x, verladen_week.values, color="#2ca02c")
        _add_line_axis_labels(ax, x, _week_labels(verladen_week.index))
        ax.set_title("Verladen Rollen pro Woche")
        ax.set_ylabel("Rollen")
        title = "Langzeit - Verladen Rollen pro Woche"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    verladen_month = longterm_data.get("verladen_month")
    if verladen_month is not None and not verladen_month.empty:
        fig, ax = plt.subplots(figsize=(10, 5), dpi=DPI)
        x = list(range(len(verladen_month)))
        ax.bar(x, verladen_month.values, color="#2ca02c")
        _add_line_axis_labels(ax, x, _month_labels(verladen_month.index))
        ax.set_title("Verladen Rollen pro Monat")
        ax.set_ylabel("Rollen")
        title = "Langzeit - Verladen Rollen pro Monat"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    needed_ma = longterm_data.get("needed_workers_week")
    available_ma = longterm_data.get("workers_week")
    if needed_ma is not None and available_ma is not None and not available_ma.empty:
        needed_ma = needed_ma.reindex(available_ma.index).fillna(0)
        fig, ax = plt.subplots(figsize=(12, 5), dpi=DPI)
        x = list(range(len(available_ma)))
        available_values = available_ma.astype(float).values
        needed_values = needed_ma.astype(float).values
        ax.plot(x, needed_values, marker="o", linewidth=2.5, color="#d62728", label="Gerechnete MA")
        ax.plot(x, available_values, marker="o", linewidth=2.5, color="#2ca02c", label="Vorhandene MA")
        ax.fill_between(x, needed_values, available_values, where=available_values >= needed_values, interpolate=True, color="#2ca02c", alpha=0.18, label="Reserve")
        ax.fill_between(x, needed_values, available_values, where=available_values < needed_values, interpolate=True, color="#d62728", alpha=0.20, label="Unterdeckung")
        _add_line_axis_labels(ax, x, _week_labels(available_ma.index))
        ax.set_ylim(0, max(float(available_ma.max()), float(needed_ma.max()), 1) * 1.15)
        ax.set_title("Gerechnete MA vs. vorhandene MA pro Woche")
        ax.set_ylabel("MA")
        ax.legend(ncol=2)
        title = "Langzeit - Gerechnete vs vorhandene MA"
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    task_time_week = longterm_data.get("task_time_share_week")
    if task_time_week is not None and not task_time_week.empty:
        title = "Langzeit - Zeitanteile Aufgaben pro Woche"
        fig = _plot_task_time_share(task_time_week, _week_labels(task_time_week.index), title)
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    task_time_month = longterm_data.get("task_time_share_month")
    if task_time_month is not None and not task_time_month.empty:
        title = "Langzeit - Zeitanteile Aufgaben pro Monat"
        fig = _plot_task_time_share(task_time_month, _month_labels(task_time_month.index), title)
        figs_meta.append({"title": title, "filename": sanitize_filename(title), "fig": fig})

    return figs_meta
