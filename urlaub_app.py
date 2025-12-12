import streamlit as st
import pandas as pd
from datetime import date, timedelta
import json

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


# ===============================
# CONFIG
# ===============================
YEAR = 2026
START_DATE = date(2026, 1, 1)   # Spalte ABK
START_COLUMN_INDEX = 0          # ABK = erste Tages-Spalte


# ===============================
# GOOGLE AUTH (SECRETS)
# ===============================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

def load_credentials_from_secrets():
    try:
        info = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    except Exception as e:
        st.error(f"Konnte Service Account JSON nicht laden: {e}")
        st.stop()


# ===============================
# GOOGLE SHEETS READ
# ===============================
def read_google_sheet(sheet_id, tab_name):
    creds = load_credentials_from_secrets()
    service = build("sheets", "v4", credentials=creds)

    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=tab_name
    ).execute()

    values = result.get("values", [])
    if not values:
        st.error("Google Sheet ist leer.")
        st.stop()

    df = pd.DataFrame(values)
    return df


# ===============================
# EASTER + BERLIN HOLIDAYS
# ===============================
def easter_sunday(year):
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def berlin_holidays_2026():
    y = 2026
    easter = easter_sunday(y)

    return {
        date(y, 1, 1),    # Neujahr
        date(y, 3, 8),    # Frauentag
        easter - timedelta(days=2),   # Karfreitag
        easter + timedelta(days=1),   # Ostermontag
        date(y, 5, 1),    # Tag der Arbeit
        easter + timedelta(days=39),  # Christi Himmelfahrt
        easter + timedelta(days=50),  # Pfingstmontag
        date(y, 10, 3),   # Tag der Deutschen Einheit
        date(y, 10, 31),  # Reformationstag
        date(y, 12, 25),  # 1. Weihnachtstag
        date(y, 12, 26),  # 2. Weihnachtstag
    }


BERLIN_HOLIDAYS = berlin_holidays_2026()


# ===============================
# STREAMLIT UI
# ===============================
st.set_page_config(page_title="Urlaubsplaner 2026", layout="wide")

st.title("🏖️ Urlaubsplaner 2026 – Google Sheets API Version (Secrets)")

st.markdown("""
Diese App:
- liest direkt aus Google Sheets (API)
- erkennt Urlaubstage anhand von **„u“**
- zählt **keine Wochenenden**
- zählt **keine Berliner Feiertage**
- zeigt Resturlaub pro Person
""")

# -------------------------------
# INPUTS
# -------------------------------
st.header("1️⃣ Personen & Kontingent")

persons_input = st.text_input(
    "Mitarbeitende (Komma-getrennt)",
    "Sonja, Mareike, Sophia, Ruta, Xenia, Anna"
)

annual_vacation = st.number_input(
    "Standard-Urlaubstage 2026 pro Person",
    min_value=0,
    max_value=60,
    value=30
)

persons = [p.strip() for p in persons_input.split(",") if p.strip()]

st.header("2️⃣ Google Sheet Daten")

sheet_id = st.text_input(
    "Google Sheet ID",
    "1Bm1kGFe_Pokr0zNiP8IBW-is2vDlGbf1oCvxZQEQhQs"
)

tab_name = st.text_input(
    "Sheet-Tab-Name",
    "Projektkalender 24/25"
)


# ===============================
# MAIN LOGIC
# ===============================
if st.button("🚀 Urlaub 2026 auswerten"):

    df = read_google_sheet(sheet_id, tab_name)

    # Erste Spalte = Namen
    name_col = df.iloc[:, 0]

    # Tages-Spalten
    day_cols = df.iloc[:, 1:]

    vacation_count = {p: 0 for p in persons}

    for row_idx, name in name_col.items():
        if name not in persons:
            continue

        for col_offset, cell in enumerate(day_cols.iloc[row_idx]):
            if str(cell).lower() != "u":
                continue

            current_date = START_DATE + timedelta(days=col_offset)

            # Filter
            if current_date.year != YEAR:
                continue
            if current_date.weekday() >= 5:  # Samstag / Sonntag
                continue
            if current_date in BERLIN_HOLIDAYS:
                continue

            vacation_count[name] += 1

    # ---------------------------
    # RESULT TABLE
    # ---------------------------
    result = []
    for p in persons:
        taken = vacation_count.get(p, 0)
        result.append({
            "Person": p,
            "Kontingent": annual_vacation,
            "Genommen": taken,
            "Rest": annual_vacation - taken
        })

    result_df = pd.DataFrame(result)

    st.header("📊 Ergebnis")
    st.dataframe(result_df, use_container_width=True)
