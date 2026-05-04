import logging
import unicodedata

import pandas as pd
import streamlit as st

from config import ALL_GERMAN_MONTHS, MONTH_NUMBER, SHIFT_ORDER

logger = logging.getLogger(__name__)

def safe_get(df, col):
    return df[col].fillna(0) if col in df else pd.Series(0, index=df.index)


def normalize_text(value) -> str:
    """Normalize workbook labels before matching."""
    if pd.isna(value):
        return ""
    return unicodedata.normalize("NFC", str(value)).strip().lower()


def _empty_result(angaben_df=None, minutes_col=None):
    return (
        pd.DataFrame(),
        pd.DataFrame(),
        angaben_df if angaben_df is not None else pd.DataFrame(),
        minutes_col,
    )


def _read_sheet_with_decimal_retry(source, sheet_name, **kwargs):
    df = pd.read_excel(source, sheet_name=sheet_name, decimal=",", **kwargs)
    if df.empty or df.isna().all().all():
        df = pd.read_excel(source, sheet_name=sheet_name, decimal=".", **kwargs)
    return df


def _find_column_case_insensitive(columns, expected_name):
    for col in columns:
        if normalize_text(col) == normalize_text(expected_name):
            return col
    return None


def _has_month_block_structure(raw: pd.DataFrame) -> bool:
    if raw.empty or raw.shape[0] < 3 or raw.shape[1] < 2:
        return False
    has_date = pd.notna(_parse_excel_date(raw.iloc[0, 1]))
    first_labels = [normalize_text(v) for v in raw.iloc[1:4, 0].tolist()]
    return has_date and any("team" in label for label in first_labels) and any("schicht" in label for label in first_labels)


def validate_workbook(xls):
    """Validate workbook structure and return readable month sheet names."""
    if "Angaben" not in xls.sheet_names:
        st.error("Die Excel-Datei enthält kein Tabellenblatt 'Angaben'. Bitte laden Sie die richtige Vorlage hoch.")
        return None

    month_sheets = [sheet_name for sheet_name in xls.sheet_names if sheet_name in ALL_GERMAN_MONTHS]
    if not month_sheets:
        st.error(
            "Keine deutschen Monatstabellen gefunden. Gefundene Blätter: "
            + ", ".join(xls.sheet_names)
        )
        return None

    try:
        angaben_preview = _read_sheet_with_decimal_retry(xls, "Angaben")
    except Exception as e:
        st.error(f"Die Tabelle 'Angaben' konnte nicht gelesen werden. Details: {e}")
        return None

    task_col = _find_column_case_insensitive(angaben_preview.columns, "Task")
    if task_col is None:
        st.error("In der Tabelle 'Angaben' fehlt die Spalte 'Task'. Bitte prüfen Sie die Excel-Vorlage.")
        return None

    minutes_col = None
    for col in angaben_preview.columns:
        if any(token in normalize_text(col) for token in ["min", "minute", "vorgabe"]):
            minutes_col = col
            break
    if minutes_col is None and len(angaben_preview.columns) <= 1:
        st.error("In der Tabelle 'Angaben' fehlt eine Minuten-/Vorgabe-Spalte.")
        return None

    has_structured_month = False
    for month in month_sheets:
        try:
            raw = _read_sheet_with_decimal_retry(xls, month, header=None)
        except Exception:
            continue
        if _has_month_block_structure(raw):
            has_structured_month = True
            break

    if not has_structured_month:
        st.error(
            "Keine Monatstabelle hat die erwartete Struktur mit Datum, Team, Schicht und KPI-Blöcken."
        )
        return None

    return month_sheets


def _parse_excel_date(value):
    if pd.isna(value) or str(value).strip() == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce", dayfirst=True)


def _validate_angaben(angaben_df):
    if angaben_df.empty or angaben_df.isna().all().all():
        st.error("Die Tabelle 'Angaben' ist leer. Bitte füllen Sie die Aufgaben und Vorgabe-Minuten aus.")
        return None, None

    task_col = _find_column_case_insensitive(angaben_df.columns, "Task")
    if task_col is None:
        st.error("In der Tabelle 'Angaben' fehlt die Spalte 'Task'. Bitte prüfen Sie die Excel-Vorlage.")
        return None, None
    if task_col != "Task":
        angaben_df = angaben_df.rename(columns={task_col: "Task"})

    minutes_col = None
    for c in angaben_df.columns:
        if any(x in str(c).lower() for x in ["min", "minute", "vorgabe"]):
            minutes_col = c
            break
    if minutes_col is None and len(angaben_df.columns) > 1:
        minutes_col = angaben_df.columns[1]
        st.warning(
            f"Keine eindeutige Minuten-Spalte in 'Angaben' gefunden. "
            f"Die App verwendet '{minutes_col}' als Vorgabe-Minuten."
        )

    if minutes_col is None:
        st.error("In der Tabelle 'Angaben' fehlt eine Minuten-/Vorgabe-Spalte.")
        return None, None

    angaben_df["Task"] = angaben_df["Task"].apply(normalize_text)
    angaben_df = angaben_df[angaben_df["Task"].ne("") & angaben_df["Task"].ne("nan")].copy()
    if angaben_df.empty:
        st.error("Die Spalte 'Task' in 'Angaben' enthält keine gültigen Aufgaben.")
        return None, None

    angaben_df[minutes_col] = pd.to_numeric(angaben_df[minutes_col], errors="coerce")
    missing_minutes = angaben_df[minutes_col].isna()
    if missing_minutes.any():
        missing_tasks = ", ".join(angaben_df.loc[missing_minutes, "Task"].head(8).tolist())
        st.warning(
            "Einige Aufgaben in 'Angaben' haben keine numerische Vorgabe-Minuten. "
            f"Diese werden mit 0 Minuten gerechnet: {missing_tasks}"
        )
        angaben_df[minutes_col] = angaben_df[minutes_col].fillna(0)

    return angaben_df, minutes_col


def _warn_metric_matching_issues(df_wide, metric_sources):
    missing = []
    for derived_metric, source_metrics in metric_sources.items():
        missing_sources = [metric for metric in source_metrics if metric not in df_wide.columns]
        if len(missing_sources) == len(source_metrics):
            missing.append(f"{derived_metric}: {', '.join(missing_sources)}")

    if not missing:
        return

    msg = (
        "Einige erwartete KPI-Namen wurden in den Monatstabellen nicht gefunden. "
        "Das kann an abweichender Schreibweise oder Encoding/Umlaut-Problemen liegen. "
        "Betroffene Zuordnungen: " + " | ".join(missing[:8])
    )
    logger.warning(msg)
    st.warning(msg)


def _warn_relevant_metric_matching_issues(df_wide, metric_sources, angaben_tasks):
    relevant_tasks = set(angaben_tasks)
    missing = []
    for derived_metric, source_metrics in metric_sources.items():
        if derived_metric not in relevant_tasks:
            continue
        missing_sources = [metric for metric in source_metrics if metric not in df_wide.columns]
        if len(missing_sources) == len(source_metrics):
            missing.append(f"{derived_metric}: {', '.join(missing_sources)}")

    if not missing:
        return

    msg = (
        "Einige Aufgaben aus 'Angaben' konnten in den Monatstabellen nicht gefunden werden. "
        "Bitte Schreibweise, Leerzeichen und Encoding/Umlaute prüfen: "
        + " | ".join(missing[:8])
    )
    logger.warning(msg)
    st.warning(msg)


@st.cache_data(show_spinner=False)
def load_excel(file):
    """
    Load and preprocess the Excel file into tidy DataFrames.
    Returns:
        df_long      : long-format records from all month sheets
        summary_long : tidy shift-level summary
        angaben_df   : Angaben sheet (normalized)
        minutes_col  : identified Minuten column
    """
    try:
        try:
            xls = pd.ExcelFile(file)
        except Exception as e:
            st.error(f"Die Excel-Datei konnte nicht geöffnet werden. Bitte prüfen Sie, ob es eine gültige .xlsx-Datei ist. Details: {e}")
            return _empty_result()

        month_sheets = validate_workbook(xls)
        if month_sheets is None:
            return _empty_result()

        # ---------------- Load Angaben sheet ----------------
        try:
            angaben_df = _read_sheet_with_decimal_retry(xls, "Angaben")
        except Exception as e:
            st.error(f"Die Tabelle 'Angaben' konnte nicht gelesen werden. Details: {e}")
            return _empty_result()

        angaben_df, minutes_col = _validate_angaben(angaben_df)
        if angaben_df is None or minutes_col is None:
            return _empty_result()

        # Dictionary: task → minutes
        angaben_dict = dict(zip(
            angaben_df["Task"],
            angaben_df[minutes_col]
        ))

        # ---------------- Load all month sheets ----------------
        records = []
        empty_months = []
        skipped_date_columns = []
        invalid_date_headers = []

        for month in month_sheets:
            try:
                raw = _read_sheet_with_decimal_retry(xls, month, header=None)
            except Exception as e:
                st.warning(f"Das Tabellenblatt '{month}' konnte nicht gelesen werden und wird übersprungen. Details: {e}")
                continue

            if raw.empty or raw.isna().all().all():
                empty_months.append(month)
                continue

            # Extract column headers (dates)
            try:
                dates = raw.iloc[0, 1:].tolist()
                parsed_dates = [_parse_excel_date(date_value) for date_value in dates]
            except Exception:
                continue

            i, n = 1, len(raw)
            while i < n:
                while i < n and (pd.isna(raw.iloc[i, 0]) or str(raw.iloc[i, 0]).strip() == ""):
                    i += 1
                if i >= n:
                    break

                block_start = i
                i += 1
                while i < n and not (pd.isna(raw.iloc[i, 0]) or str(raw.iloc[i, 0]).strip() == ""):
                    i += 1
                block_end = i

                block = raw.iloc[block_start:block_end].reset_index(drop=True)
                if block.empty:
                    continue
                if block.shape[0] < 3:
                    st.warning(
                        f"Ein Datenblock in '{month}' hat zu wenige Zeilen "
                        "für Team, Schicht und KPI-Werte und wird übersprungen."
                    )
                    continue

                teams = block.iloc[0, 1:].tolist()
                schichten = block.iloc[1, 1:].tolist()

                for kpi_row in range(2, block.shape[0]):
                    kpi_name = normalize_text(block.iloc[kpi_row, 0])
                    if not kpi_name:
                        continue
                    for col in range(1, block.shape[1]):
                        datum = dates[col - 1]
                        parsed_date = parsed_dates[col - 1] if col - 1 < len(parsed_dates) else pd.NaT
                        team = teams[col - 1]
                        schicht = schichten[col - 1]
                        value = block.iloc[kpi_row, col]
                        if pd.isna(datum) or str(datum).strip() == "":
                            continue
                        if pd.isna(parsed_date):
                            invalid_date_headers.append((month, col, datum))
                            continue
                        if parsed_date.month != MONTH_NUMBER[month]:
                            skipped_date_columns.append((month, col, parsed_date.strftime("%d.%m.%Y")))
                            continue
                        records.append(
                            {
                                "Datum": parsed_date,
                                "Team": team,
                                "Schicht": schicht,
                                "Metric": kpi_name,
                                "Value": value,
                            }
                        )
                i += 1

        if empty_months and len(empty_months) < len(month_sheets):
            st.info("Leere Monatstabellen wurden übersprungen: " + ", ".join(empty_months))

        if invalid_date_headers:
            unique_invalid = list(dict.fromkeys(
                f"{month} Spalte {col}: {datum}" for month, col, datum in invalid_date_headers
            ))
            st.warning(
                "Einige Datumsspalten konnten nicht gelesen werden und wurden übersprungen: "
                + " | ".join(unique_invalid[:8])
            )

        if skipped_date_columns:
            unique_skipped = list(dict.fromkeys(
                f"{month} Spalte {col}: {date_value}" for month, col, date_value in skipped_date_columns
            ))
            st.info(
                "Datumsspalten außerhalb des jeweiligen Monats wurden übersprungen: "
                + " | ".join(unique_skipped[:8])
            )

        if not records:
            st.error("Keine verwertbaren Daten in den Monatstabellen gefunden. Bitte prüfen Sie Datumszeile, Teamzeile, Schichtzeile und KPI-Blöcke.")
            return _empty_result(angaben_df, minutes_col)

        # ---------------- Build df_long ----------------
        df_long = pd.DataFrame(records)
        df_long["Value"] = pd.to_numeric(df_long["Value"], errors="coerce").fillna(0)
        df_long["Datum"] = pd.to_datetime(df_long["Datum"], errors="coerce", dayfirst=True)
        invalid_dates = df_long["Datum"].isna().sum()
        if invalid_dates:
            st.info(f"{invalid_dates} Datensätze ohne gültiges Datum wurden ignoriert.")
            df_long = df_long[df_long["Datum"].notna()].copy()
        if df_long.empty:
            st.error("Nach der Datumsprüfung sind keine gültigen Monatstabellen-Daten übrig.")
            return _empty_result(angaben_df, minutes_col)

        # Normalize Metric names
        df_long["Metric"] = df_long["Metric"].apply(normalize_text)

        duplicate_keys = (
            df_long.groupby(["Datum", "Team", "Schicht", "Metric"], observed=False)
            .size()
            .loc[lambda s: s > 1]
        )
        if not duplicate_keys.empty:
            sample = []
            for datum, team, schicht, metric in duplicate_keys.head(5).index:
                sample.append(f"{datum.strftime('%d.%m.%Y')} / {team} / {schicht} / {metric}")
            st.warning(
                "Mehrere Werte mit gleichem Datum, Team, Schicht und KPI wurden gefunden und summiert. "
                "Bitte prüfen Sie doppelte Datumsspalten in Excel: " + " | ".join(sample)
            )

        # ---------------- Pivot to wide ----------------
        df_wide = (
            df_long.pivot_table(
                index=["Datum", "Team", "Schicht"],
                columns="Metric",
                values="Value",
                aggfunc="sum",
            )
            .reset_index()
        )

        # ---------------- Vectorized shift summary ----------------
        summary_df = pd.DataFrame(index=df_wide.index)
        metric_sources = {
            "sägen": ["auftragsrollen gesägt", "abfallrollen gesägt"],
            "rauslegen": ["rausgelegte rollen", "säge raus"],
            "richten": ["gerichtete rollen", "säge gerichtet"],
            "zusammenfahren": ["zusammengefahrene rollen"],
            "verladen": ["verladene rollen"],
            "cutten": ["cut rollen"],
            "absetzen": ["eingelagerte rollen produktion"],
            "absetzen 2": ["rollen umgelagert absetzer", "säge eingelagert"],
            "kontrolle dmg/retouren": ["damaged bearbeitet", "retouren bearbeitet"],
            "packen paletten liegend": ["rollen auf palette liegend (rollenanzahl)"],
            "packen paletten stehend": ["rollen auf palette stehend (rollenanzahl)"],
            "souscouche abladen": ["souscouche abgeladen(rollen)", "souscouche abgeladen (rollen)"],
            "serbien abladen": ["entladen serbien"],
            "serbien abladen tautliner": ["entladen serbien tautliner"],
            "serbien einlagern": ["serbien rollen eingelagert"],
            "sonstiges / aufräumarbeiten (in std)": ["dafür gebraucht (stunden)"],
            "vorhandene ma": ["anzahl ma"],
        }
        _warn_relevant_metric_matching_issues(df_wide, metric_sources, angaben_dict.keys())

        known_sources = {source for sources in metric_sources.values() for source in sources}
        ignored_metrics = sorted(
            metric for metric in df_wide.columns
            if metric not in known_sources and metric not in ["Datum", "Team", "Schicht"]
        )
        if ignored_metrics:
            logger.info("Unknown KPI metrics skipped: %s", ignored_metrics)
            with st.expander("⚠️ Unbekannte KPIs übersprungen"):
                st.write(
                    "Diese KPI-Zeilen wurden eingelesen, aber keiner berechneten Kennzahl zugeordnet:"
                )
                st.write(", ".join(ignored_metrics))

        for derived_metric, source_metrics in metric_sources.items():
            summary_df[derived_metric] = sum(safe_get(df_wide, source) for source in source_metrics)

        # ---------------- Compute benötigte & differenz MA ----------------
        workload_minutes = pd.Series(0, index=summary_df.index)
        for kpi in angaben_dict.keys():
            if kpi in summary_df.columns:
                workload_minutes += summary_df[kpi] * angaben_dict[kpi]
        total_hours = workload_minutes / 60 + summary_df["sonstiges / aufräumarbeiten (in std)"]
        summary_df["benötigte ma"] = (total_hours / 7.5).round(1)
        summary_df["differenz ma"] = (summary_df["vorhandene ma"] - summary_df["benötigte ma"]).round(1)

        # ---------------- Add identifiers ----------------
        summary_df[["Datum", "Team", "Schicht"]] = df_wide[["Datum", "Team", "Schicht"]]

        # ---------------- Melt tidy for plotting ----------------
        id_cols = ["Datum", "Schicht", "Team"]
        summary_long = summary_df.melt(id_vars=id_cols, var_name="Metric", value_name="Value")

        # Normalize Schicht
        summary_long["Schicht"] = (
            summary_long["Schicht"].astype(str).str.strip().str.title()
        )
        summary_long["Schicht"] = pd.Categorical(
            summary_long["Schicht"], categories=SHIFT_ORDER, ordered=True
        )

        return df_long, summary_long, angaben_df, minutes_col

    except Exception as e:
        st.error(f"Fehler beim Laden der Excel-Datei: {e}")
        return _empty_result()
