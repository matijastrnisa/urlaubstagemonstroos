import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import json
import re

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


# -----------------------------------------------------------
# CONFIG
# -----------------------------------------------------------
st.set_page_config(page_title="Urlaubsplaner 2026", layout="wide")
st.title("🏖 Urlaubsplaner 2026 – Google Sheets API Version (Secrets)")

st.markdown("""
Diese App:
- liest direkt aus Google Sheets (API)
- erkennt Urlaubstage anhand von **„u“** in Zellen
- zählt alle Urlaubstage im Jahr **2026**
- zeigt Kontingent, genommene Tage und Resturlaub pro Person
""")


# -----------------------------------------------------------
# GOOGLE SHEETS API SETUP (via Streamlit Secrets)
# -----------------------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

if "GCP_SERVICE_ACCOUNT_JSON" not in st.secrets:
    st.error("❌ Streamlit Secret 'GCP_SERVICE_ACCOUNT_JSON' fehlt. Manage app → Settings → Secrets.")
    st.stop()

try:
    sa_info = json.loads(st.secrets["GCP_SERVICE_ACCOUNT_JSON"])
    credentials = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
except Exception as e:
    st.error(f"❌ Konnte Service Account JSON nicht laden. Prüfe Secrets-Format. Details: {e}")
    st.stop()

service = build("sheets", "v4", credentials=credentials, cache_discovery=False)


# -----------------------------------------------------------
# SHEET SETTINGS
# -----------------------------------------------------------
SPREADSHEET_ID = "1Bm1kGFe_Pokr0zNiP8IBW-is2vDlGbf1oCvxZQEQhQs"

st.subheader("0️⃣ Sheet-Einstellungen")
sheet_name = st.text_input("Sheet Tab Name (z.B. 'Sheet1')", value="Sheet1")
RANGE = f"{sheet_name}"


# -----------------------------------------------------------
# HILFSFUNKTIONEN
# -----------------------------------------------------------
def extract_dates_row(values):
    best_idx = None
    best_count = -1
    best_map = {}

    for i, row in enumerate(values):
        count = 0
        date_map = {}
        for j, cell in enumerate(row[1:], start=1):  # col0 = name
            d = None
            if isinstance(cell, str):
                for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
                    try:
                        d = datetime.strptime(cell.strip(), fmt).date()
                        break
                    except Exception:
                        pass
            if d is not None:
                count += 1
                date_map[j] = d

        if count > best_count:
            best_count = count
            best_idx = i
            best_map = date_map

    return best_idx, best_map


def normalize_name(s):
    return s.split(":")[0].strip().lower()


# -----------------------------------------------------------
# PERSONEN & KONTINGENT
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

st.subheader("2️⃣ Auswertung starten")

if st.button("🚀 Urlaub 2026 auswerten"):
    try:
        sheet = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE
        ).execute()

        values = sheet.get("values", [])
        if not values:
            st.error("Keine Daten im Sheet gefunden. Prüfe Tab-Name.")
            st.stop()

        date_row_idx, date_map = extract_dates_row(values)
        if date_row_idx is None or len(date_map) == 0:
            st.error("Konnte keine Datumszeile erkennen. (Sind Datumswerte als Text vorhanden?)")
            st.stop()

        st.success(f"Datumszeile erkannt (0-basiert): {date_row_idx}")

        # Personenzeilen finden
        person_rows = {}
        for i, row in enumerate(values):
            if not row:
                continue
            cell0 = row[0]
            if isinstance(cell0, str):
                nm = normalize_name(cell0)
                if nm in personen_norm:
                    original = personen[personen_norm.index(nm)]
                    person_rows[original] = i

        if not person_rows:
            st.error("Keine Personen gefunden. Prüfe, ob Namen in Spalte A stehen.")
            st.stop()

        # Urlaub zählen
        urlaub_genommen = {p: 0 for p in personen}

        for person, row_idx in person_rows.items():
            row = values[row_idx]
            for col_idx, d in date_map.items():
                if d.year == 2026:
                    if col_idx < len(row):
                        cell = row[col_idx]
                        if isinstance(cell, str) and cell.strip().lower() == "u":
                            urlaub_genommen[person] += 1

        rows = []
        for p in personen:
            genommen = urlaub_genommen.get(p, 0)
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

        fig = px.bar(
            df,
            x="Person",
            y="Resturlaub_2026",
            title="Resturlaub 2026 pro Person",
            text="Resturlaub_2026"
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

        st.download_button(
            "CSV herunterladen",
            df.to_csv(index=False).encode("utf-8"),
            "Urlaub_2026.csv",
            "text/csv"
        )

    except Exception as e:
        st.error(f"Fehler beim Lesen des Google Sheets: {e}")
