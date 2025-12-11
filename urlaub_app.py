import streamlit as st
import pandas as pd
from datetime import datetime, date
import plotly.express as px

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


# -----------------------------------------------------------
# CONFIG
# -----------------------------------------------------------

st.set_page_config(page_title="Urlaubsplaner 2026", layout="wide")
st.title("🏖 Urlaubsplaner 2026 – Google Sheets API Version")

st.markdown("""
Diese App:

- liest **direkt aus Google Sheets** (keine Uploads notwendig)
- erkennt Urlaubstage anhand von **„u“** in Zellen
- zählt alle Urlaubstage im Jahr **2026**
- zeigt Kontingent, genommene Tage und Resturlaub pro Person
""")


# -----------------------------------------------------------
# GOOGLE SHEETS API SETUP
# -----------------------------------------------------------

# Die JSON-Credentials müssen in deinem GitHub Repo liegen:
SERVICE_ACCOUNT_FILE = "service_account.json"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

credentials = Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=SCOPES
)

service = build("sheets", "v4", credentials=credentials)

# -----------------------------------------------------------
# GOOGLE SHEET EINSTELLUNGEN
# -----------------------------------------------------------

SPREADSHEET_ID = "1Bm1kGFe_Pokr0zNiP8IBW-is2vDlGbf1oCvxZQEQhQs"
RANGE = "Sheet1"   # Wenn dein Sheet anders heißt → anpassen


# -----------------------------------------------------------
# HILFSFUNKTIONEN
# -----------------------------------------------------------

def extract_dates_row(values):
    """
    Findet die Zeile mit den meisten Datumswerten.
    Liefert: (row_index, dict{col_index -> date})
    """
    best_idx = None
    best_count = -1
    best_map = {}

    for i, row in enumerate(values):
        count = 0
        date_map = {}
        for j, cell in enumerate(row[1:], start=1):  # Spalte 0 = Name
            d = None
            if isinstance(cell, str):
                for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
                    try:
                        d = datetime.strptime(cell.strip(), fmt).date()
                        break
                    except:
                        pass
            elif isinstance(cell, datetime):
                d = cell.date()

            if d is not None:
                count += 1
                date_map[j] = d

        if count > best_count:
            best_count = count
            best_idx = i
            best_map = date_map

    return best_idx, best_map


def normalize_name(s):
    """
    Entfernt ':' und Groß/Kleinschreibung.
    Beispiel: "Mareike:" → "mareike"
    """
    return s.split(":")[0].strip().lower()


# -----------------------------------------------------------
# PERSONEN & URLAUBSKONTINGENT
# -----------------------------------------------------------

st.subheader("1️⃣ Personen & Kontingent")

personen_input = st.text_input(
    "Mitarbeitende (Komma-getrennt)",
    "Sonja, Mareike, Sophia, Ruta, Xenia, Anna"
)

personen = [p.strip() for p in personen_input.split(",") if p.strip()]
personen_norm = [normalize_name(p) for p in personen]

default_kontingent = st.number_input(
    "Standard-Urlaubstage 2026 pro Person",
    min_value=1, max_value=60, value=30
)


# -----------------------------------------------------------
# AUSWERTUNG STARTEN
# -----------------------------------------------------------

st.subheader("2️⃣ Auswertung starten")

if st.button("🚀 Urlaub 2026 auswerten"):

    try:
        # SHEET LADEN
        sheet = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE
        ).execute()

        values = sheet.get("values", [])

        if not values:
            st.error("Keine Daten im Google Sheet gefunden.")
            st.stop()

        # DATUMSZEILE FINDEN
        date_row_idx, date_map = extract_dates_row(values)

        if date_row_idx is None or len(date_map) == 0:
            st.error("Konnte keine Datumszeile finden.")
            st.stop()

        st.success(f"Datumszeile erkannt: Zeile {date_row_idx + 1} im Sheet.")

        # PERSONENZEILEN FINDEN
        person_rows = {}
        for i, row in enumerate(values):
            if len(row) == 0:
                continue
            cell = row[0]
            if isinstance(cell, str):
                nm = normalize_name(cell)
                if nm in personen_norm:
                    original = personen[personen_norm.index(nm)]
                    person_rows[original] = i

        if not person_rows:
            st.error("Keine Personen im Sheet gefunden.")
            st.stop()

        st.write("Gefundene Personenzeilen:")
        st.json(person_rows)

        # URLAUB ZÄHLEN
        urlaub_genommen = {p: 0 for p in personen}

        for person, row_idx in person_rows.items():
            row = values[row_idx]
            for col_idx, d in date_map.items():
                if d.year == 2026:
                    if col_idx < len(row):
                        cell = row[col_idx]
                        if isinstance(cell, str) and cell.strip().lower() == "u":
                            urlaub_genommen[person] += 1

        # RESULTAT BAUEN
        rows = []
        for p in personen:
            genommen = urlaub_genommen[p]
            kont = default_kontingent
            rest = kont - genommen
            rows.append({
                "Person": p,
                "Kontingent_2026": kont,
                "Urlaub_2026_genommen": genommen,
                "Resturlaub_2026": rest
            })

        df = pd.DataFrame(rows)

        st.subheader("📊 Ergebnis – Urlaub 2026")
        st.dataframe(df, use_container_width=True)

        # BALKENDIAGRAMM
        fig = px.bar(
            df,
            x="Person",
            y="Resturlaub_2026",
            title="Resturlaub 2026 pro Person",
            text="Resturlaub_2026"
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

        # EXPORT
        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "CSV herunterladen",
            csv,
            "Urlaub_2026.csv",
            "text/csv"
        )

    except Exception as e:
        st.error(f"Fehler beim Lesen des Google Sheets: {e}")
