import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

# Funktion für Fensteranteil-Rundung
def runde_fensteranteil(anteil):
    stufen = [10, 20, 30, 40, 50]
    return min(stufen, key=lambda x: abs(x - anteil))

# Funktion für Sheet-Auswahl basierend auf Baujahr
def finde_passendes_sheet(gebäude_datum):
    stufen = {
        datetime(1, 1, 1): "vor 1900",
        datetime(1900, 1, 1): "ab 01.01.1900",
        datetime(1945, 1, 1): "ab 01.01.1945",
        datetime(1960, 1, 1): "ab 01.01.1960",
        datetime(1976, 11, 15): "ab 15.11.1976",
        datetime(1993, 10, 1): "ab 01.10.1993",
        datetime(2001, 10, 26): "ab 26.10.2001"
    }
    passende_stufen = {k: v for k, v in stufen.items() if k <= gebäude_datum}
    if not passende_stufen:
        return None
    bestes_datum = max(passende_stufen.keys())
    return passende_stufen[bestes_datum]

# --- Streamlit App ---
st.title("Gebäude-Sanierung: Optimale Maßnahmen finden")

st.header("Gib die Gebäudedaten ein:")

baujahr_input = st.date_input("Baujahr des Gebäudes (Datum eingeben)", format="DD.MM.YYYY")
neues_gebäude_VA = st.number_input("V/A Verhältnis", min_value=0.0, format="%.3f")
neues_gebäude_AWBF = st.number_input("AW/BF Verhältnis", min_value=0.0, format="%.3f")
neuer_fensteranteil_input = st.number_input("Fensteranteil in % (beliebige Eingabe erlaubt)", min_value=0.0, max_value=100.0, format="%.1f")

# Optionale Eingabe der bebaute Fläche
grundfläche_input = st.number_input("Bebaute Fläche (optional, in m²)", min_value=0.0, format="%.2f")

if st.button("Analyse starten"):
    # Baujahr (Datum) in datetime.datetime umwandeln
    baujahr_input_dt = datetime.combine(baujahr_input, datetime.min.time())
    blattname = finde_passendes_sheet(baujahr_input_dt)
    if blattname is None:
        st.error("Für dieses Baujahr ist kein passendes Blatt vorhanden.")
    else:
        # Excel-Daten für das richtige Blatt einlesen
        file_path = "Simulationen.xlsx"  # Excel-Datei muss im gleichen Verzeichnis liegen
        excel_data = pd.read_excel(file_path, sheet_name=blattname, header=None)

        # Extrahieren
        va_werte = excel_data.iloc[0, 2:].values
        awbf_werte = excel_data.iloc[1, 2:].values
        maßnahmen_beschreibung = excel_data.iloc[2:, 0].values
        fensteranteile = excel_data.iloc[2:, 1].values
        einsparungen = excel_data.iloc[2:, 2:].values

        # Datenpool aufbauen
        datenpool_liste = []
        for gebäude_index in range(len(va_werte)):
            va = va_werte[gebäude_index]
            awbf = awbf_werte[gebäude_index]

            for maßnahmen_index in range(len(maßnahmen_beschreibung)):
                beschreibung = maßnahmen_beschreibung[maßnahmen_index]
                fensteranteil = fensteranteile[maßnahmen_index]
                einsparung = einsparungen[maßnahmen_index, gebäude_index]

                datenpool_liste.append({
                    "V_A": va,
                    "AW_BF": awbf,
                    "Maßnahme": beschreibung,
                    "Fensteranteil": fensteranteil,
                    "Einsparung_%": einsparung
                })

        datenpool_df = pd.DataFrame(datenpool_liste)

        # Normierung vorbereiten
        VA_min, VA_max = datenpool_df["V_A"].min(), datenpool_df["V_A"].max()
        AWBF_min, AWBF_max = datenpool_df["AW_BF"].min(), datenpool_df["AW_BF"].max()

        datenpool_df["V_A_normiert"] = (datenpool_df["V_A"] - VA_min) / (VA_max - VA_min)
        datenpool_df["AW_BF_normiert"] = (datenpool_df["AW_BF"] - AWBF_min) / (AWBF_max - AWBF_min)

        # Fensteranteil runden
        neuer_fensteranteil = runde_fensteranteil(neuer_fensteranteil_input)
        st.info(f"Fensteranteil {neuer_fensteranteil_input}% wurde gerundet auf {neuer_fensteranteil}% für die Analyse.")

        # Nur passende Datensätze auswählen
        df_passend = datenpool_df[datenpool_df["Fensteranteil"] == neuer_fensteranteil]

        if df_passend.empty:
            st.error(f"Keine Simulationen mit {neuer_fensteranteil}% Fensteranteil vorhanden!")
        else:
            # Distanz berechnen
            df_passend = df_passend.copy()
            df_passend["Distanz"] = np.sqrt(
                (df_passend["V_A_normiert"] - ((neues_gebäude_VA - VA_min) / (VA_max - VA_min)))**2 +
                (df_passend["AW_BF_normiert"] - ((neues_gebäude_AWBF - AWBF_min) / (AWBF_max - AWBF_min)))**2
            )

            # Beste Simulation bestimmen
            bester_va = df_passend.loc[df_passend["Distanz"].idxmin(), "V_A"]
            bester_awbf = df_passend.loc[df_passend["Distanz"].idxmin(), "AW_BF"]

            # Alle Maßnahmen zu diesem Gebäude und Fensteranteil filtern
            alle_massnahmen = datenpool_df[
                (datenpool_df["V_A"] == bester_va) &
                (datenpool_df["AW_BF"] == bester_awbf) &
                (datenpool_df["Fensteranteil"] == neuer_fensteranteil)
            ]

            # Maßnahmen gruppieren
            def klassifiziere_massnahme(text):
                if "oberste Decke" in text or "Dachaufbau" in text:
                    return "Oberste Decke"
                elif "unterste Decke" in text:
                    return "Unterste Decke"
                elif "Fenstertausch" in text:
                    return "Fenster"
                elif "AW Sanierung" in text or "Fassade" in text:
                    return "AW"
                else:
                    return "Sonstige"

            alle_massnahmen["Gruppe"] = alle_massnahmen["Maßnahme"].apply(klassifiziere_massnahme)

            # Durchschnittliche Einsparung pro Gruppe
            gruppen_ergebnis = alle_massnahmen.groupby("Gruppe")["Einsparung_%"].mean().sort_values(ascending=False)

            # Ausgabe
            st.success("Beste Varianten:")
            for gruppe in gruppen_ergebnis.head(2).index:
                st.markdown(f"### {gruppe}")
                details = alle_massnahmen[alle_massnahmen["Gruppe"] == gruppe]
                for idx, row in details.iterrows():
                    st.markdown(f"- Maßnahme: **{row['Maßnahme']}**, Energieeinsparung: **{row['Einsparung_%']:.2f}%**")
