import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from datetime import datetime, date
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

Wichtig:
- Diese Version sucht **keine Datumszeile**.
- Stattdessen gibst du das **Startdatum der ersten Tages-Spalte** an (Spalte B).
""")

# ----------------------------------------
# GOOGLE API AUTH (STREAMLIT SECRETS)
# ----------------------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

if "gcp_service_account" not in st.secrets:
    st.error("❌ Service Account fehlt in Streamlit Secrets. (Manage app → Settings → Secrets)")
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

st.header("2️⃣b Datum-Logik (wichtig)")
start_date = st.date_input(
    "Startdatum der ersten Tages-Spalte (Spalte B)",
    value=date(2025, 1, 1)
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

def find_person_rows(values, personen):
    """Sucht Personen in Spalte 0. Match ist case-insensitive."""
    target = {normalize_name(p): p for p in personen}
    found = {}
    for i, row in enumerate(values):
        if not row:
            continue
        nm = normalize_name(row[0])
        if nm in target:
            found[target[nm]] = i
    return found

def date_for_column(col_idx: int) -> date:
    """
    col_idx 1 -> start_date (Spalte B)
    col_idx 2 -> start_date + 1 Tag (Spalte C)
    usw.
    """
    return (pd.Timestamp(start_date) + pd.Timedelta(days=col_idx - 1)).date()

# ----------------------------------------
# MAIN
# ----------------------------------------
st.header("3️⃣ Auswertung starten")

if st.button("🚀 Urlaub 2026 auswerten"):
    try:
        # 1) Sheet lesen
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name
        ).execute()

        values = result.get("values", [])
        if not values:
            st.error("❌ Sheet ist leer oder nicht lesbar.")
            st.stop()

        # 2) Zeilen auf gleiche Länge bringen
        values = pad_rows(values)

        # 3) Personenzeilen finden
        person_rows = find_person_rows(values, personen)
        if not person_rows:
            st.error("❌ Keine Personen gefunden. Prüfe, ob Namen in Spalte A stehen.")
            st.stop()

        st.success("✅ Personen gefunden.")
        with st.expander("Debug: Gefundene Personenzeilen"):
            st.json(person_rows)

        # 4) Urlaub zählen (nur 2026, nur 'u')
        urlaub_genommen = {p: 0 for p in personen}

        for person, row_idx in person_rows.items():
            row = values[row_idx]

            # Spalte 0 = Name, ab Spalte 1 = Tage
            for col_idx in range(1, len(row)):
                d = date_for_column(col_idx)

                if d.year == 2026:
                    cell = str(row[col_idx]).strip().lower()
                    if cell == "u":
                        urlaub_genommen[person] += 1

        # 5) Ergebnis-DataFrame
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

        fig = px.bar(
            df,
            x="Person",
            y="Resturlaub_2026",
            text="Resturlaub_2026",
            title="Resturlaub 2026 pro Person"
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

        st.download_button(
            "⬇️ Ergebnis als CSV herunterladen",
            df.to_csv(index=False).encode("utf-8"),
            "Urlaub_2026.csv",
            "text/csv"
        )

        with st.expander("Debug: Rohdaten-Vorschau (erste 10 Zeilen / 30 Spalten)"):
            preview = pd.DataFrame(values[:10])
            st.dataframe(preview.iloc[:, :30], use_container_width=True)

    except Exception as e:
        st.error(f"❌ Fehler beim Lesen des Google Sheets: {e}")
