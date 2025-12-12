import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from datetime import datetime
import calendar

# ----------------------------------------
# PAGE SETUP
# ----------------------------------------
st.set_page_config(
    page_title="Urlaubsplaner 2026 – Google Sheets API",
    layout="centered"
)

st.title("🏖️ Urlaubsplaner 2026 – Google Sheets API Version (Secrets)")

st.markdown("""
Diese App:
- liest **direkt aus Google Sheets (API)**
- erkennt Urlaubstage anhand von **„u“ in Zellen**
- zählt **alle Urlaubstage im Jahr 2026**
- zeigt **Kontingent, genommene Tage & Resturlaub pro Person**
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

    service = build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False
    )
except Exception as e:
    st.error(f"❌ Konnte Service Account nicht laden: {e}")
    st.stop()

# ----------------------------------------
# USER INPUT
# ----------------------------------------
st.header("1️⃣ Personen & Kontingent")

names_input = st.text_input(
    "Mitarbeitende (Komma-getrennt)",
    "Sonja, Mareike, Sophia, Ruta, Xenia, Anna"
)

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
    value="MP"
)

# ----------------------------------------
# HELPERS
# ----------------------------------------
def get_2026_dates():
    dates = []
    for month in range(1, 13):
        for day in range(1, calendar.monthrange(2026, month)[1] + 1):
            dates.append(datetime(2026, month, day))
    return dates


def read_sheet():
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet_name
    ).execute()
    return result.get("values", [])


# ----------------------------------------
# MAIN LOGIC
# ----------------------------------------
st.header("3️⃣ Auswertung starten")

if st.button("🚀 Urlaub 2026 auswerten"):
    try:
        raw_data = read_sheet()

        if not raw_data:
            st.error("❌ Sheet ist leer oder nicht lesbar.")
            st.stop()

        header = raw_data[0]
        rows = raw_data[1:]

        df = pd.DataFrame(rows, columns=header)
        df.fillna("", inplace=True)

        personen = [n.strip() for n in names_input.split(",") if n.strip()]
        date_columns = df.columns[1:]  # erste Spalte = Name

        results = []

        for person in personen:
            if person not in df.iloc[:, 0].values:
                results.append({
                    "Person": person,
                    "Genommene Urlaubstage": 0,
                    "Urlaubskontingent": urlaubstage_pro_person,
                    "Resturlaub": urlaubstage_pro_person
                })
                continue

            row = df[df.iloc[:, 0] == person].iloc[0]
            genommen = 0

            for col in date_columns:
                cell_value = str(row[col]).strip().lower()
                if cell_value == "u":
                    genommen += 1

            results.append({
                "Person": person,
                "Genommene Urlaubstage": genommen,
                "Urlaubskontingent": urlaubstage_pro_person,
                "Resturlaub": urlaubstage_pro_person - genommen
            })

        result_df = pd.DataFrame(results)

        st.success("✅ Auswertung erfolgreich")
        st.dataframe(result_df, use_container_width=True)

        # ----------------------------------------
        # DOWNLOAD
        # ----------------------------------------
        st.download_button(
            "⬇️ Ergebnis als Excel herunterladen",
            data=result_df.to_excel(index=False, engine="openpyxl"),
            file_name="Urlaubsplan_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Fehler beim Lesen des Google Sheets: {e}")
