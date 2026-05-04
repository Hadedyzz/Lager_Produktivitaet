from config import SAEGEN_TARGET


def _pct_change(start, end):
    return ((end - start) / start * 100) if start else 0


def _consecutive_weeks_above_target(series, target):
    count = 0
    for value in reversed(series.tolist()):
        if value > target:
            count += 1
        else:
            break
    return count


def generate_longterm_insights(longterm_data: dict) -> list[str]:
    """Generate German plain-text observations from long-term data."""
    if not longterm_data:
        return []

    insights = []
    rpm = longterm_data.get("rolls_per_ma")
    if rpm is not None and len(rpm) >= 2:
        first = rpm[rpm > 0].head(1)
        last = rpm[rpm > 0].tail(1)
        if not first.empty and not last.empty:
            change = _pct_change(float(first.iloc[0]), float(last.iloc[0]))
            week = last.index[-1].isocalendar().week
            insights.append(f"Produktivität ist bis KW{week} um {change:+.1f}% verändert.")

    saegen = longterm_data.get("saegen_week")
    if saegen is not None and not saegen.empty:
        saegen_mean = saegen["mean"]
        above_count = int((saegen_mean > SAEGEN_TARGET).sum())
        insights.append(f"Sägen-Mittelwert übersteigt das Ziel in {above_count} Wochen.")
        consecutive = _consecutive_weeks_above_target(saegen_mean, SAEGEN_TARGET)
        if consecutive > 8:
            insights.append(
                f"Aktuelles Ziel ist seit {consecutive} Wochen unter dem Durchschnitt - Neukalibrierung empfohlen?"
            )

    task_indexed = longterm_data.get("task_indexed")
    if task_indexed is not None and "Verladen" in task_indexed and not task_indexed.empty:
        verladen = task_indexed["Verladen"].replace(0, float("nan")).dropna()
        if not verladen.empty:
            start = float(verladen.head(1).iloc[0])
            end = float(verladen.tail(1).iloc[0])
            insights.append(f"Verladen-Volumen ist {_pct_change(start, end):+.1f}% gegenüber dem Startniveau verändert.")

    return [insight for insight in insights if insight]
