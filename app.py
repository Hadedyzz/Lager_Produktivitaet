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
    try:
        # Read Angaben sheet (unchanged)
        angaben_df = pd.read_excel(file, sheet_name="Angaben", decimal=',')
        if angaben_df.empty or angaben_df.isna().all().all():
            angaben_df = pd.read_excel(file, sheet_name="Angaben", decimal='.')

        # Attempt to read multiple month sheets: Juli, August, September, Oktober
        months = ["Juli", "August", "September", "Oktober"]
        records = []
        for month in months:
            # Try common casings for sheet names
            for sheet_name in (month, month.lower(), month.capitalize()):
                try:
                    raw = pd.read_excel(file, sheet_name=sheet_name, header=None, decimal=',')
                except Exception:
                    raw = None
                # Fallback to dot decimal if read produced empty frame
                if (raw is None) or raw.empty or raw.isna().all().all():
                    try:
                        raw = pd.read_excel(file, sheet_name=sheet_name, header=None, decimal='.')
                    except Exception:
                        raw = None
                if raw is None or raw.empty or raw.isna().all().all():
                    continue

                # Parse this raw sheet into records (same logic as before)
                try:
                    dates = raw.iloc[0, 1:].tolist()
                except Exception:
                    # Malformed sheet, skip
                    continue

                i, n = 1, len(raw)
                while i < n:
                    # Skip empty rows
                    while i < n and (pd.isna(raw.iloc[i, 0]) or str(raw.iloc[i, 0]).strip() == ''):
                        i += 1
                    if i >= n:
                        break
                    block_start = i
                    i += 1
                    # Find end of block
                    while i < n and not (pd.isna(raw.iloc[i, 0]) or str(raw.iloc[i, 0]).strip() == ''):
                        i += 1
                    block_end = i

                    block = raw.iloc[block_start:block_end].reset_index(drop=True)
                    if block.empty:
                        continue

                    teams = block.iloc[0, 1:].tolist()
                    schichten = block.iloc[1, 1:].tolist()

                    for kpi_row in range(2, block.shape[0]):
                        kpi_name = block.iloc[kpi_row, 0]
                        if pd.isna(kpi_name) or str(kpi_name).strip() == '':
                            continue
                        for col in range(1, block.shape[1]):
                            datum = dates[col-1]
                            team = teams[col-1]
                            schicht = schichten[col-1]
                            value = block.iloc[kpi_row, col]
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
                # Stop trying other casings for this month once a valid sheet parsed
                break

        # If no records found, return empty frames (preserve prior error behavior)
        if not records:
            st.error("⚠️ Keine Daten in den Blättern Juli/August/September/Oktober gefunden.")
            return pd.DataFrame(), pd.DataFrame(), angaben_df

        # Build df_long exactly as before
        df_long = pd.DataFrame(records)
        df_long["Value"] = pd.to_numeric(df_long["Value"], errors="coerce")
        df_long["Datum"] = pd.to_datetime(df_long["Datum"], errors="coerce", dayfirst=True)

        # --- Compute per-shift summary (same logic as summary scripts) ---
        # Pivot to wide
        df_wide = (
            df_long
            .pivot_table(
                index=["Datum", "Team", "Schicht"],
                columns="Metric",
                values="Value",
                aggfunc='first'
            )
            .reset_index()
        )
        df_wide.columns = [str(col) for col in df_wide.columns]

        # Determine angaben minutes column (flexible)
        angaben_col = None
        for c in angaben_df.columns:
            if "min" in str(c).lower() or "minute" in str(c).lower() or "vorgabe" in str(c).lower():
                angaben_col = c
                break
        if angaben_col is None:
            angaben_col = "Minutes/roll" if "Minutes/roll" in angaben_df.columns else angaben_df.columns[1]

        angaben_dict = dict(zip(angaben_df['Task'], angaben_df[angaben_col]))

        def calc_shift_summary(row):
            def get_value(col):
                val = row.get(col, 0)
                return val if pd.notnull(val) else 0

            summary = {
                'Sägen': get_value('Auftragsrollen gesägt') + get_value('Abfallrollen gesägt'),
                'rauslegen': get_value('rausgelegte Rollen') + get_value('Säge raus'),
                'zusammenfahren': get_value('zusammengefahrene Rollen'),
                'verladen': get_value('verladene Rollen'),
                'cutten': get_value('Cut Rollen'),
                'absetzen': get_value('eingelagerte Rollen Produktion'),
                'absetzen 2': get_value('Rollen umgelagert Absetzer') + get_value('Säge eingelagert'),
                'kontrolle DMG/Retouren': get_value('Damaged bearbeitet') + get_value('Retouren bearbeitet'),
                'packen Paletten liegend': get_value('Rollen auf Palette liegend (Rollenanzahl)'),
                'packen Paletten stehend': get_value('Rollen auf Palette stehend (Rollenanzahl)'),
                'Souscouche abladen': get_value('Souscouche abgeladen (Rollen)'),
                'Serbien abladen': get_value('entladen Serbien'),
                'Serbien abladen Tautliner': get_value('entladen Serbien Tautliner'),
                'Serbien einlagern ': get_value('Serbien Rollen eingelagert'),
                'sonstiges / Aufräumarbeiten (in Std)': get_value('dafür gebraucht (Stunden)'),
                'vorhandene MA': get_value('Anzahl MA')
            }

            kpi_to_angaben = {
                'Sägen': 'sägen',
                'rauslegen': 'rauslegen',
                'zusammenfahren': 'zusammenfahren',
                'verladen': 'verladen',
                'cutten': 'cutten',
                'absetzen': 'absetzen',
                'absetzen 2': 'absetzen 2',
                'kontrolle DMG/Retouren': 'kontrolle DMG/Retouren',
                'packen Paletten liegend': 'packen Paletten liegend',
                'packen Paletten stehend': 'packen Paletten stehend',
                'Souscouche abladen': 'Souscouche abladen',
                'Serbien abladen': 'Serbien abladen',
                'Serbien abladen Tautliner': 'Serbien abladen Tautliner',
                'Serbien einlagern ': 'Serbien einlagern '
            }

            total_minutes = sum(
                summary[kpi] * angaben_dict.get(task, 0)
                for kpi, task in kpi_to_angaben.items()
            )
            total_hours = total_minutes / 60 + summary.get('sonstiges / Aufräumarbeiten (in Std)', 0)
            summary['benötigte MA'] = round(total_hours / 7.5, 1) if total_hours > 0 else 0
            summary['Differenz MA'] = round(summary['vorhandene MA'] - summary['benötigte MA'], 1)
            return pd.Series(summary)

        summary_df = df_wide.apply(calc_shift_summary, axis=1)
        summary_df[['Datum', 'Team', 'Schicht']] = df_wide[['Datum', 'Team', 'Schicht']]

        # Rename to nice names
        nice_names = {
            "Sägen": "Sägen (Stück)",
            "richten": "Richten (Stück)",
            "rauslegen": "Rauslegen (Stück)",
            "zusammenfahren": "Zusammenfahren (Stück)",
            "verladen": "Verladen (Stück)",
            "cutten": "Cutten (Stück)",
            "absetzen": "Absetzen (Stück)",
            "absetzen 2": "Absetzen 2 (Stück)",
            "kontrolle DMG/Retouren": "Kontrolle DMG/Retouren (Stück)",
            "packen Paletten liegend": "Packen Paletten liegend (Stück)",
            "packen Paletten stehend": "Packen Paletten stehend (Stück)",
            "Souscouche abladen": "Souscouche abladen (Stück)",
            "Serbien abladen": "Serbien abladen (Stück)",
            "Serbien abladen Tautliner": "Serbien abladen Tautliner (Stück)",
            "Serbien einlagern ": "Serbien einlagern (Stück)",
            "sonstiges / Aufräumarbeiten (in Std)": "Sonstiges / Aufräumarbeiten (Std)",
            "vorhandene MA": "Vorhandene MA",
            "benötigte MA": "Benötigte MA",
            "Differenz MA": "Differenz MA"
        }
        summary_df = summary_df.rename(columns=nice_names)

        # Re-order columns
        main_cols = ["Datum", "Schicht", "Team", "Vorhandene MA", "Benötigte MA", "Differenz MA"]
        rest_cols = [c for c in summary_df.columns if c not in main_cols]
        ordered_cols = main_cols + rest_cols
        summary_df = summary_df.loc[:, ordered_cols]

        # Normalize Schicht and order
        summary_df["Schicht"] = summary_df["Schicht"].astype(str).str.strip().str.title()
        summary_df["Schicht"] = pd.Categorical(summary_df["Schicht"], categories=shift_order, ordered=True)
        summary_df = summary_df.sort_values(["Datum", "Schicht", "Team"]).reset_index(drop=True)

        # Melt tidy for downstream plots (same shape as shift_summary_kwXX_tidy.csv)
        id_cols = ["Datum", "Schicht", "Team"]
        value_cols = [c for c in summary_df.columns if c not in id_cols]
        summary_long = (
            summary_df
            .melt(id_vars=id_cols, value_vars=value_cols, var_name="Metric", value_name="Value")
            .sort_values(id_cols + ["Metric"])
        )

        return df_long, summary_long, angaben_df
    except Exception as e:
        st.error(f"⚠️ Fehler beim Laden der Excel-Datei: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# =====================
# 2. WEEKLY PLOTS
# =====================
def plot_weekly_charts(df, target_date):
    """Weekly charts: df must include Datum, Schicht, Metric, Value"""
    try:
        figs = []
        # week range
        target_date = pd.to_datetime(target_date)
        start_of_week = target_date - pd.Timedelta(days=target_date.weekday())
        end_of_week = start_of_week + pd.Timedelta(days=6)
        all_dates = pd.date_range(start=start_of_week, end=end_of_week)

        # Filter to Monday to Friday only
        weekdays = all_dates[all_dates.weekday < 5]

        # normalize + filter
        df = df.copy()
        df["Schicht"] = df["Schicht"].astype(str).str.strip().str.title()
        df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
        df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)
        df = df[(df["Datum"] >= start_of_week) & (df["Datum"] <= end_of_week)]
        df = df[df["Datum"].isin(weekdays)]

        # Update all_dates to weekdays only
        all_dates = weekdays

        # ========== 1) "Sägen" per day by shift ==========
        saegen_rolls = df[df["Metric"].str.contains("sägen", case=False, na=False)].copy()
        if not saegen_rolls.empty:
            saegen_by_day_shift = saegen_rolls.groupby(["Datum", "Schicht"])["Value"].sum().unstack("Schicht")
            saegen_by_day_shift = saegen_by_day_shift.reindex(columns=shift_order).fillna(0)
            saegen_by_day_shift = saegen_by_day_shift.reindex(all_dates).fillna(0)
            plot_days_saegen = saegen_by_day_shift.index

            fig, ax = plt.subplots(figsize=(12, 6))
            saegen_by_day_shift.plot(kind="bar", stacked=True, ax=ax, color=shift_colors_for_cols(shift_order))
            ax.set_title("Anzahl gesägte Rollen pro Tag")
            ax.set_xlabel("Datum")
            ax.set_ylabel("Anzahl gesägte Rollen")
            ax.set_xticks(range(len(plot_days_saegen)))
            ax.set_xticklabels(format_day_month(plot_days_saegen), rotation=45)
            target = 70
            ymax = max(saegen_by_day_shift.sum(axis=1).max(), target) * 1.15
            ax.set_ylim(0, ymax)
            ax.axhline(target, color="red", linestyle="--", linewidth=2, label="Ziel 70")

            valid_saegen_days = saegen_by_day_shift.sum(axis=1) > 0
            if valid_saegen_days.any():
                avg_saegen = saegen_by_day_shift.sum(axis=1)[valid_saegen_days].mean()
                ax.axhline(avg_saegen, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
                ax.text(len(plot_days_saegen) - 0.5, avg_saegen + ymax * 0.02, f"Durchschnitt: {int(round(avg_saegen))}",
                        color="blue", fontsize=12, fontweight="bold", ha="left", va="bottom")
            ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1), borderaxespad=0)

            # numbers on stacks
            for i, day in enumerate(saegen_by_day_shift.index):
                bottom = 0
                total = 0
                for shift in saegen_by_day_shift.columns:
                    value = saegen_by_day_shift.loc[day, shift]
                    total += value
                    if value > 0:
                        ax.text(i, bottom + value/2, f"{int(round(value))}",
                                ha="center", va="center", color="white", fontsize=10, fontweight="bold")
                    bottom += value
                if total > 0:
                    ax.text(i, min(total + ymax*0.03, ymax*0.97), f"{int(round(total))}",
                            ha="center", va="bottom", color="black", fontsize=12, fontweight="bold")
            ax.text(len(plot_days_saegen)-0.5, target + ymax*0.04, "Ziel: 70",
                    ha="left", va="bottom", color="red", fontsize=13, fontweight="bold")
            plt.tight_layout()
            figs.append(fig)

        # ========== 2) Total rolls per day by task group ==========
        main_tasks = ["Absetzen (Stück)", "Richten (Stück)", "Verladen (Stück)", "Zusammenfahren (Stück)"]
        task_order = ["Absetzen", "Richten", "Verladen", "Zusammenfahren", "Sonstige"]
        task_colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#8c564b", "#888888"]

        def classify_task(metric):
            return metric.replace(" (Stück)", "") if metric in main_tasks else "Sonstige"

        rolls = df[df["Metric"].str.contains(r"\(Stück\)$", regex=True)].copy()
        if not rolls.empty:
            rolls["TaskGroup"] = rolls["Metric"].apply(classify_task)
            total_rolls_by_group = (rolls.groupby(["Datum", "TaskGroup"])["Value"]
                                    .sum().unstack("TaskGroup"))
            total_rolls_by_group = total_rolls_by_group.reindex(columns=task_order).fillna(0)
            total_rolls_by_group = total_rolls_by_group.reindex(all_dates).fillna(0)
            plot_days_group = total_rolls_by_group.index

            workers_per_day = df[df["Metric"] == "Vorhandene MA"].groupby("Datum")["Value"].sum()
            workers_per_day = workers_per_day.reindex(all_dates).fillna(0)

            fig, ax = plt.subplots(figsize=(12, 6))
            total_rolls_by_group.plot(kind="bar", stacked=True, ax=ax, color=task_colors)
            ax.set_title("Gesamte Anzahl an bewegten Rollen pro Tag")
            ax.set_xlabel("Datum")
            ax.set_ylabel("Gesamtrollen")
            ax.set_xticks(range(len(plot_days_group)))
            ax.set_xticklabels(format_day_month(plot_days_group), rotation=45)
            ymax_group = max(total_rolls_by_group.sum(axis=1).max(), 1) * 1.15
            ax.set_ylim(0, ymax_group)
            ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1), borderaxespad=0)

            valid_group_days = total_rolls_by_group.sum(axis=1) > 0
            if valid_group_days.any():
                avg_group = total_rolls_by_group.sum(axis=1)[valid_group_days].mean()
                ax.axhline(avg_group, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
                ax.text(len(plot_days_group) - 0.5, avg_group + ymax_group * 0.02, f"Durchschnitt: {int(round(avg_group))}",
                        color="blue", fontsize=12, fontweight="bold", ha="left", va="bottom")
            # annotations
            for i, day in enumerate(plot_days_group):
                bottom = 0
                for task in task_order:
                    value = total_rolls_by_group.loc[day, task]
                    if value > 0:
                        ax.text(i, bottom + value/2, f"{int(round(value))}",
                                ha="center", va="center", color="white", fontsize=10, fontweight="bold")
                    bottom += value
                total = total_rolls_by_group.loc[day].sum()
                ax.text(i, min(bottom + ymax_group*0.03, ymax_group*0.97), f"{int(round(total))}",
                        ha="center", va="bottom", color="black", fontsize=12, fontweight="bold")
                # MA box
                ma = workers_per_day.get(day, 0)
                ma_str = f"{ma:.1f}" if not float(ma).is_integer() else f"{int(ma)}"
                ax.text(i + 0.35, bottom / 2, f"MA: {ma_str}",
                        ha="left", va="center", color="black", fontsize=8,
                        bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="gray", alpha=0.7))
            plt.tight_layout()
            figs.append(fig)

        # ========== 3) Gesamtrollen pro Tag nach Schicht ==========
        if not rolls.empty:
            total_shift = rolls.groupby(["Datum", "Schicht"])["Value"].sum().unstack("Schicht")
            total_shift = total_shift.reindex(columns=shift_order).fillna(0)
            total_shift = total_shift.reindex(all_dates).fillna(0)

            ymax_shift = max(total_shift.sum(axis=1).max(), 1) * 1.15

            fig, ax = plt.subplots(figsize=(12, 6))
            total_shift.plot(kind="bar", stacked=True, ax=ax, color=shift_colors_for_cols(shift_order))
            ax.set_title("Gesamte Anzahl an bewegten Rollen pro Tag nach Schicht")
            ax.set_xlabel("Datum")
            ax.set_ylabel("Gesamtrollen")
            ax.set_xticks(range(len(all_dates)))
            ax.set_xticklabels(format_day_month(all_dates), rotation=45)

            # >>> average only over days with any rolls recorded
            day_totals_rolls = total_shift.sum(axis=1)
            valid_roll_days = day_totals_rolls > 0
            avg_total = day_totals_rolls[valid_roll_days].mean() if valid_roll_days.any() else 0.0

            ax.axhline(avg_total, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
            ax.text(
                len(all_dates) - 0.5,
                avg_total + ymax_shift * 0.02,
                f"Durchschnitt: {int(round(avg_total))}",
                color="blue", fontsize=12, fontweight="bold", ha="left", va="bottom"
            )

            # numbers on stacks + totals
            for i, day in enumerate(all_dates):
                bottom = 0
                for shift in shift_order:
                    value = total_shift.loc[day, shift] if shift in total_shift.columns else 0
                    if value > 0:
                        ax.text(i, bottom + value/2, f"{int(round(value))}",
                                ha="center", va="center", color="white", fontsize=10, fontweight="bold")
                    bottom += value
                total = total_shift.loc[day].sum()
                ax.text(i, min(total + ymax_shift * 0.03, ymax_shift * 0.97),
                        f"{int(round(total))}", ha="center", va="bottom",
                        color="black", fontsize=12, fontweight="bold")

            plt.tight_layout()
            figs.append(fig)


        # ========== 4) Gesamtrollen pro MA ==========
        workers_per_day = df[df["Metric"] == "Vorhandene MA"].groupby("Datum")["Value"].sum().reindex(all_dates).fillna(0)
        total_rolls_per_day = (df[df["Metric"].str.contains(r"\(Stück\)$", regex=True)]
                               .groupby("Datum")["Value"].sum().reindex(all_dates).fillna(0))
        if (workers_per_day > 0).any() and (total_rolls_per_day > 0).any():
            norm = (total_rolls_per_day / workers_per_day).fillna(0)
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.bar(range(len(norm.index)), norm.values, color="#1f77b4")
            ax.set_title("Gesamtrollen pro MA")
            ax.set_xlabel("Datum")
            ax.set_ylabel("Rollen pro MA")
            ax.set_xticks(range(len(norm.index)))
            ax.set_xticklabels(format_day_month(norm.index), rotation=45)

            # >>> FIX: define ymax_norm
            ymax_norm = max(norm.values) * 1.15 if len(norm) > 0 else 1
            ax.set_ylim(0, ymax_norm)

            avg_norm = norm[norm > 0].mean() if (norm > 0).any() else 0
            ax.axhline(avg_norm, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
            ax.text(len(norm.index) - 0.5, avg_norm + ymax_norm * 0.02,
                    f"Durchschnitt: {int(round(avg_norm))}",
                    color="blue", fontsize=12, fontweight="bold", ha="left", va="bottom")

            for i, v in enumerate(norm.values):
                ax.text(i, v + ymax_norm * 0.03, f"{int(round(v))}",
                        ha="center", va="bottom", color="black", fontsize=12, fontweight="bold")

            plt.tight_layout()
            figs.append(fig)

        # ========== 5) Anzahl MA pro Tag ==========
        workers = df[df["Metric"] == "Vorhandene MA"].copy()
        if not workers.empty:
            w = workers.groupby(["Datum", "Schicht"])["Value"].sum().unstack("Schicht").fillna(0)
            w = w.reindex(columns=shift_order).fillna(0)
            w = w.reindex(all_dates).fillna(0)

            ymax_workers = max(w.sum(axis=1).max(), 1) * 1.15

            fig, ax = plt.subplots(figsize=(10, 5))
            w.plot(kind="bar", stacked=True, ax=ax, color=shift_colors_for_cols(shift_order))
            ax.set_title("Anzahl MA pro Tag")
            ax.set_xlabel("Datum")
            ax.set_ylabel("Anzahl MA")
            ax.set_xticks(range(len(all_dates)))
            ax.set_xticklabels(format_day_month(all_dates), rotation=45)

            # >>> average only over days with any workers recorded
            day_totals_workers = w.sum(axis=1)
            valid_workers_days = day_totals_workers > 0
            avg_workers = day_totals_workers[valid_workers_days].mean() if valid_workers_days.any() else 0.0

            ax.axhline(avg_workers, color="blue", linestyle="--", linewidth=2, label="Durchschnitt")
            ax.text(
                len(all_dates) - 0.5,
                avg_workers + ymax_workers * 0.02,
                f"Durchschnitt: {avg_workers:.1f}",
                color="blue", fontsize=12, fontweight="bold", ha="left", va="bottom"
            )

            # numbers on stacks + totals
            for i, day in enumerate(all_dates):
                bottom = 0
                for shift in shift_order:
                    value = w.loc[day, shift] if shift in w.columns else 0
                    if value > 0:
                        val_str = f"{value:.1f}" if not float(value).is_integer() else f"{int(value)}"
                        ax.text(i, bottom + value/2, val_str, ha="center", va="center",
                                color="white", fontsize=10, fontweight="bold")
                    bottom += value
                total = w.loc[day].sum()
                total_str = f"{total:.1f}" if not float(total).is_integer() else f"{int(total)}"
                ax.text(i, min(total + ymax_workers * 0.03, ymax_workers * 0.97),
                        total_str, ha="center", va="bottom", color="black", fontsize=12, fontweight="bold")

            plt.tight_layout()
            figs.append(fig)


        return figs
    except Exception as e:
        st.error(f"⚠️ Fehler beim Erstellen der Wochenanalyse: {e}")
        return []

# =====================
# 3. DAILY PLOTS
# =====================
def plot_daily_charts(df, angaben_df, target_day):
    try:
        figs = []
        df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
        df["Value"] = pd.to_numeric(df["Value"], errors="coerce").fillna(0)
        day_data = df[(df["Datum"].dt.date == target_day) & (df["Metric"].str.contains(r"\(Stück\)$"))].copy()

        # Join with Angaben
        angaben_df["Task"] = angaben_df["Task"].str.strip().str.lower()
        day_data["Task"] = day_data["Metric"].str.replace(r"\s*\(Stück\)$", "", regex=True).str.strip().str.lower()
        shift_task = day_data.groupby(["Schicht", "Task"])["Value"].sum().reset_index()
        shift_task = pd.merge(shift_task, angaben_df, on="Task", how="inner")
        shift_task["Hours"] = (shift_task["Value"] * shift_task["Minutes/roll"]) / 60
        shift_task["FTE"] = shift_task["Hours"] / 7.5
        shift_task = shift_task[shift_task["Hours"] > 0]

        # Get task order for consistent x-axis
        task_order = shift_task.groupby("Task")["Hours"].sum().sort_values(ascending=False).index.tolist()
        shifts = shift_task["Schicht"].unique()

        # --- Color mapping for shifts ---
        shift_colors = [shift_color_map.get(shift, "#cccccc") for shift in shifts]

        # Grouped bar plot: Hours per Task by Shift
        if len(shifts) > 0:
            bar_width = 0.8 / len(shifts)
            x = range(len(task_order))

            fig, ax = plt.subplots(figsize=(12, 6))
            for idx, shift in enumerate(shifts):
                shift_data = shift_task[shift_task["Schicht"] == shift]
                hours = [shift_data[shift_data["Task"] == t]["Hours"].sum() if t in shift_data["Task"].values else 0 for t in task_order]
                ftes = [shift_data[shift_data["Task"] == t]["FTE"].sum() if t in shift_data["Task"].values else 0 for t in task_order]
                ax.bar(
                    [i + idx * bar_width for i in x],
                    hours,
                    width=bar_width,
                    color=shift_color_map.get(shift, "#cccccc"),
                    label=f"{shift} (MA: {sum(ftes):.1f})"
                )
            ax.set_xticks([i + bar_width * (len(shifts) - 1) / 2 for i in x])
            ax.set_xticklabels([t.capitalize() for t in task_order], rotation=60, ha="right", fontsize=10)
            ax.set_title(f"Arbeitsstunden pro Aufgabe und Schicht am {target_day.strftime('%d.%m.%Y')}")
            ax.set_xlabel("Aufgabe")
            ax.set_ylabel("Stunden")
            ax.legend(title="Schicht (MA gesamt)", loc="upper left", bbox_to_anchor=(1.02, 1))
            plt.tight_layout()
            figs.append(fig)

        return figs
    except Exception as e:
        st.error(f"⚠️ Fehler beim Erstellen der Tagesanalyse: {e}")
        return []

# =====================
# 4. STREAMLIT APP
# =====================
st.set_page_config(page_title="Rollenbewegung Dashboard", layout="wide")
st.title("📊 Rollenbewegung Dashboard")

uploaded_file = st.file_uploader("Excel hochladen", type=["xlsx"])
if uploaded_file:
    try:
        df_raw, df_summary, angaben_df = load_excel(uploaded_file)

        if df_raw.empty or df_summary.empty or angaben_df.empty:
            st.warning("⚠️ Die hochgeladene Datei enthält keine gültigen Daten.")
        else:
            st.session_state["df_raw"] = df_raw
            st.session_state["df_summary"] = df_summary
            st.session_state["angaben_df"] = angaben_df

            mode = st.selectbox("Analysemodus wählen", ["Woche", "Tag"])

            if mode == "Woche":
                week_date = st.date_input("Datum innerhalb der Woche auswählen", datetime.today())
                if st.button("Wochen-Analyse starten"):
                    figs = plot_weekly_charts(df_summary, week_date)
                    for fig in figs:
                        st.pyplot(fig)

            elif mode == "Tag":
                day_date = st.date_input("Tag auswählen", datetime.today())
                if st.button("Tages-Analyse starten"):
                    figs = plot_daily_charts(df_summary, angaben_df, day_date)
                    for fig in figs:
                        st.pyplot(fig)
    except Exception as e:
        st.error(f"⚠️ Fehler in der Streamlit-App: {e}")
