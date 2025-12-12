import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from datetime import datetime
import plotly.express as px

# ----------------------------------------
# PAGE SETUP
# ----------------------------------------
st.set_page_config(page_title="Urlaubsplaner 2026 – Google Sheets API", layout="wide")
st.title("🏖️ Urlaubsplaner 2026 – Google Sheets API Version (Secrets)")

st.markdown("""
Diese App:
- liest **direkt aus Google Sheets (API)**
- erkennt Urlaubstage anhand von **„u“** in Zellen
- zählt **alle Urlaubstage im Jahr 2026**
- zeigt **Kontingent, genommene Tage & Resturlaub** pro Person
""")

# ----------------------------------------
# GOOGLE API AUTH (STREAMLIT SECRETS)
# ----------------------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

if "gcp_service_account" not in st.secrets:
    st.error("❌ Service Account fehlt in Streamlit Secrets.")
    st.stop()

try:
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )
    service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
except Exception as e:
    st.error(f"❌ Konnte Service Account nicht laden: {e}")
    st.stop()

# ----------------------------------------
# INPUTS
# ----------------------------------------
st.header("1️⃣ Personen & Kontingent")

names_input = st.text_input(
    "Mitarbeitende (Komma-getrennt)",
    "Sonja, Mareike, Sophia, Ruta, Xenia, Anna"
)
personen = [n.strip() for n in names_input.split(",") if n.strip()]

urlaubstage_pro_person = st.number_input(
    "Standard-Urlaubstage 2026 pro Person",
    min_value=0,
    max_value=60,
    value=30
)

st.header("2️⃣ Google Sheet Daten")

spreadsheet_id = st.text_input(
    "Google Sheet ID",
    value="1Bm1kGFe_Pokr0zNiP8IBW-is2vDlGbf1oCvxZQEQhQs"
)

sheet_name = st.text_input(
    "Sheet-Tab-Name",
    value="Projektkalender 24/25"
)

# ----------------------------------------
# HELPERS
# ----------------------------------------
def normalize_name(s: str) -> str:
    return str(s).split(":")[0].strip().lower()

def pad_rows(values):
    """Macht alle Zeilen gleich lang, indem sie mit "" aufgefüllt werden."""
    max_len = max(len(r) for r in values) if values else 0
    padded = []
    for r in values:
        r2 = list(r) + [""] * (max_len - len(r))
        padded.append(r2)
    return padded

def extract_dates_row(values):
    """
    Findet die Zeile mit den meisten Datumswerten (ab Spalte 1).
    Unterstützt Formate: YYYY-MM-DD, DD.MM.YYYY, DD/MM/YYYY
    """
    best_idx = None
    best_count = -1
    best_map = {}

    for i, row in enumerate(values):
        count = 0
        date_map = {}

        for j, cell in enumerate(row[1:], start=1):
            d = None
            if isinstance(cell, str):
                c = cell.strip()
                for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
                    try:
                        d = datetime.strptime(c, fmt).date()
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

def find_person_rows(values, personen):
    """
    Sucht Personen in Spalte 0. Match ist case-insensitive.
    """
    target = {normalize_name(p): p for p in personen}
    found = {}
    for i, row in enumerate(values):
        if not row:
            continue
        nm = normalize_name(row[0])
        if nm in target:
            found[target[nm]] = i
    return found

# ----------------------------------------
# MAIN
# ----------------------------------------
st.header("3️⃣ Auswertung starten")

if st.button("🚀 Urlaub 2026 auswerten"):
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name
        ).execute()

        values = result.get("values", [])
        if not values:
            st.error("❌ Sheet ist leer oder nicht lesbar.")
            st.stop()

        # 1) Zeilen auf gleiche Länge bringen (wichtig!)
        values = pad_rows(values)

        # 2) Datumszeile automatisch finden
        date_row_idx, date_map = extract_dates_row(values)
        if date_row_idx is None or len(date_map) == 0:
            st.error("❌ Konnte keine Datumszeile finden. Sind Datumswerte als Text vorhanden?")
            st.stop()

        st.success(f"✅ Datumszeile erkannt (0-basiert): {date_row_idx}")

        # 3) Personenzeilen finden
        person_rows = find_person_rows(values, personen)
        if not person_rows:
            st.error("❌ Keine Personen gefunden. Prüfe, ob Namen in Spalte A stehen.")
            st.stop()

        st.write("Gefundene Personenzeilen:")
        st.json(person_rows)

        # 4) Urlaub zählen (nur 2026, nur 'u')
        urlaub_genommen = {p: 0 for p in personen}

        for person, row_idx in person_rows.items():
            row = values[row_idx]
            for col_idx, d in date_map.items():
                if d.year == 2026:
                    cell = str(row[col_idx]).strip().lower()
                    if cell == "u":
                        urlaub_genommen[person] += 1

        # 5) Ergebnis
        rows = []
        for p in personen:
            genommen = urlaub_genommen.get(p, 0)
            rows.append({
                "Person": p,
                "Kontingent_2026": urlaubstage_pro_person,
                "Urlaub_2026_genommen": genommen,
                "Resturlaub_2026": urlaubstage_pro_person - genommen
            })

        df = pd.DataFrame(rows)

        st.subheader("📊 Ergebnis – Urlaub 2026")
        st.dataframe(df, use_container_width=True)

        fig = px.bar(df, x="Person", y="Resturlaub_2026", text="Resturlaub_2026",
                     title="Resturlaub 2026 pro Person")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

        st.download_button(
            "⬇️ Ergebnis als CSV herunterladen",
            df.to_csv(index=False).encode("utf-8"),
            "Urlaub_2026.csv",
            "text/csv"
        )

        # Optional: Debug-Vorschau (kannst du später rausnehmen)
        with st.expander("Debug: Rohdaten-Vorschau (erste 15 Zeilen)"):
            preview_df = pd.DataFrame(values[:15])
            st.dataframe(preview_df, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Fehler beim Lesen des Google Sheets: {e}")
