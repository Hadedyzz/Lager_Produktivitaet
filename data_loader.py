# data_loader.py
import pandas as pd
import streamlit as st

SHIFT_ORDER = ["Früh", "Spät", "Nacht"]

def safe_get(df, col):
    return df[col].fillna(0) if col in df else pd.Series(0, index=df.index)


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
        # ---------------- Load Angaben sheet ----------------
        angaben_df = pd.read_excel(file, sheet_name="Angaben", decimal=",")
        if angaben_df.empty or angaben_df.isna().all().all():
            angaben_df = pd.read_excel(file, sheet_name="Angaben", decimal=".")

        # Identify Minuten column once
        minutes_col = None
        for c in angaben_df.columns:
            if any(x in str(c).lower() for x in ["min", "minute", "vorgabe"]):
                minutes_col = c
                break
        if minutes_col is None and len(angaben_df.columns) > 1:
            minutes_col = angaben_df.columns[1]

        # Normalize task names in Angaben
        angaben_df["Task"] = (
            angaben_df["Task"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # Dictionary: task → minutes
        angaben_dict = dict(zip(
            angaben_df["Task"],
            angaben_df[minutes_col]
        ))

        # ---------------- Load all month sheets ----------------
        xls = pd.ExcelFile(file)
        months = ["Juli", "August", "September", "Oktober"]
        records = []

        for month in months:
            if month not in xls.sheet_names:
                continue
            raw = pd.read_excel(xls, sheet_name=month, header=None, decimal=",")
            if raw.empty or raw.isna().all().all():
                raw = pd.read_excel(xls, sheet_name=month, header=None, decimal=".")

            if raw.empty or raw.isna().all().all():
                continue

            # Extract column headers (dates)
            try:
                dates = raw.iloc[0, 1:].tolist()
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

                teams = block.iloc[0, 1:].tolist()
                schichten = block.iloc[1, 1:].tolist()

                for kpi_row in range(2, block.shape[0]):
                    kpi_name = str(block.iloc[kpi_row, 0]).strip()
                    if not kpi_name:
                        continue
                    for col in range(1, block.shape[1]):
                        datum = dates[col - 1]
                        team = teams[col - 1]
                        schicht = schichten[col - 1]
                        value = block.iloc[kpi_row, col]
                        if pd.isna(datum) or str(datum).strip() == "":
                            continue
                        records.append(
                            {
                                "Datum": datum,
                                "Team": team,
                                "Schicht": schicht,
                                "Metric": kpi_name.strip().lower(),  # normalize
                                "Value": value,
                            }
                        )
                i += 1

        if not records:
            st.error("⚠️ Keine Daten in den Blättern Juli/August/September/Oktober gefunden.")
            return pd.DataFrame(), pd.DataFrame(), angaben_df, minutes_col

        # ---------------- Build df_long ----------------
        df_long = pd.DataFrame(records)
        df_long["Value"] = pd.to_numeric(df_long["Value"], errors="coerce").fillna(0)
        df_long["Datum"] = pd.to_datetime(df_long["Datum"], errors="coerce", dayfirst=True)

        # Normalize Metric names
        df_long["Metric"] = (
            df_long["Metric"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # ---------------- Pivot to wide ----------------
        df_wide = (
            df_long.pivot_table(
                index=["Datum", "Team", "Schicht"],
                columns="Metric",
                values="Value",
                aggfunc="first",
            )
            .reset_index()
        )

        # ---------------- Vectorized shift summary ----------------
        summary_df = pd.DataFrame(index=df_wide.index)

        summary_df["sägen"] = safe_get(df_wide, "auftragsrollen gesägt") + safe_get(df_wide, "abfallrollen gesägt")
        summary_df["rauslegen"] = safe_get(df_wide, "rausgelegte rollen") + safe_get(df_wide, "säge raus")
        summary_df["richten"] = safe_get(df_wide, "gerichtete rollen") + safe_get(df_wide, "säge gerichtet")
        summary_df["zusammenfahren"] = safe_get(df_wide, "zusammengefahrene rollen")
        summary_df["verladen"] = safe_get(df_wide, "verladene rollen")
        summary_df["cutten"] = safe_get(df_wide, "cut rollen")
        summary_df["absetzen"] = safe_get(df_wide, "eingelagerte rollen produktion")
        summary_df["absetzen 2"] = safe_get(df_wide, "rollen umgelagert absetzer") + safe_get(df_wide, "säge eingelagert")
        summary_df["kontrolle dmg/retouren"] = safe_get(df_wide, "damaged bearbeitet") + safe_get(df_wide, "retouren bearbeitet")
        summary_df["packen paletten liegend"] = safe_get(df_wide, "rollen auf palette liegend (rollenanzahl)")
        summary_df["packen paletten stehend"] = safe_get(df_wide, "rollen auf palette stehend (rollenanzahl)")
        summary_df["souscouche abladen"] = safe_get(df_wide, "souscouche abgeladen (rollen)")
        summary_df["serbien abladen"] = safe_get(df_wide, "entladen serbien")
        summary_df["serbien abladen tautliner"] = safe_get(df_wide, "entladen serbien tautliner")
        summary_df["serbien einlagern"] = safe_get(df_wide, "serbien rollen eingelagert")
        summary_df["sonstiges / aufräumarbeiten (in std)"] = safe_get(df_wide, "dafür gebraucht (stunden)")
        summary_df["vorhandene ma"] = safe_get(df_wide, "anzahl ma")

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
        st.error(f"⚠️ Fehler beim Laden der Excel-Datei: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None
