import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from datetime import date
import plotly.express as px

# ----------------------------------------
# PAGE SETUP
# ----------------------------------------
st.set_page_config(page_title="Urlaubsplaner 2026 – Google Sheets API", layout="wide")
st.title("🏖️ Urlaubsplaner 2026 – Google Sheets API (Secrets)")

st.markdown("""
Diese App:
- liest **direkt aus Google Sheets (API)**
- erkennt Urlaubstage anhand von **„u“** in Zellen
- zählt **Urlaubstage im Jahr 2026**
- zählt **keine Wochenenden (Sa/So)**
- zählt **keine Berliner Feiertage 2026**
- zeigt **Kontingent, genommene Tage & Resturlaub**
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

st.header("3️⃣ Spalten-/Datum-Referenz (entscheidend)")
st.caption("Du hast gesagt: **01.01.2026 ist Spalte ABK**. Das nutzen wir als Referenz.")

ref_col_letters = st.text_input("Referenz-Spalte (z.B. ABK)", value="ABK")
ref_date = st.date_input("Referenz-Datum für diese Spalte", value=date(2026, 1, 1))

# ----------------------------------------
# BERLIN HOLIDAYS 2026
# ----------------------------------------
def easter_sunday(year: int) -> date:
    # Gregorian Easter (Anonymous algorithm)
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

def berlin_holidays_2026() -> set:
    y = 2026
    easter = easter_sunday(y)
    good_friday = easter + pd.Timedelta(days=-2)
    easter_monday = easter + pd.Timedelta(days=1)
    ascension = easter + pd.Timedelta(days=39)
    whit_monday = easter + pd.Timedelta(days=50)

    # Berlin-specific: Frauentag (Mar 8) + Reformationstag (Oct 31)
    return {
        date(y, 1, 1),   # Neujahr
        date(y, 3, 8),   # Frauentag (Berlin)
        good_friday.date(),     # Karfreitag
        easter_monday.date(),   # Ostermontag
        date(y, 5, 1),   # Tag der Arbeit
        ascension.date(),       # Christi Himmelfahrt
        whit_monday.date(),     # Pfingstmontag
        date(y, 10, 3),  # Tag der Deutschen Einheit
        date(y, 10, 31), # Reformationstag (Berlin)
        date(y, 12, 25), # 1. Weihnachtstag
        date(y, 12, 26), # 2. Weihnachtstag
    }

BERLIN_HOLIDAYS_2026 = berlin_holidays_2026()

def is_workday_berlin_2026(d: date) -> bool:
    # Mo=0 .. So=6
    if d.weekday() >= 5:
        return False
    if d in BERLIN_HOLIDAYS_2026:
        return False
    return True

# ----------------------------------------
# HELPERS
# ----------------------------------------
def normalize_name(s: str) -> str:
    return str(s).split(":")[0].strip().lower()

def pad_rows(values):
    max_len = max(len(r) for r in values) if values else 0
    return [list(r) + [""] * (max_len - len(r)) for r in values]

def find_person_rows(values, personen):
    target = {normalize_name(p): p for p in personen}
    found = {}
    for i, row in enumerate(values):
        if not row:
            continue
        nm = normalize_name(row[0])
        if nm in target:
            found[target[nm]] = i
    return found

def col_to_index(col_letters: str) -> int:
    """
    Excel/Sheets Column letters -> 1-based index:
    A=1, B=2, ..., Z=26, AA=27, ...
    """
    col_letters = col_letters.strip().upper()
    n = 0
    for ch in col_letters:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Ungültige Spalte: {col_letters}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n

def date_for_column(col_idx_1based: int, ref_col_idx_1based: int, ref_date_val: date) -> date:
    """
    Mappt jede Spalte auf ein Datum, basierend auf Referenz.
    Annahme: jede Spalte entspricht +1 Kalendertag.
    """
    delta_days = col_idx_1based - ref_col_idx_1based
    return (pd.Timestamp(ref_date_val) + pd.Timedelta(days=delta_days)).date()

# ----------------------------------------
# MAIN
# ----------------------------------------
st.header("4️⃣ Auswertung starten")

if st.button("🚀 Urlaub 2026 auswerten"):
    try:
        # sheet lesen
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=sheet_name
        ).execute()

        values = result.get("values", [])
        if not values:
            st.error("❌ Sheet ist leer oder nicht lesbar.")
            st.stop()

        values = pad_rows(values)

        # personen finden
        person_rows = find_person_rows(values, personen)
        if not person_rows:
            st.error("❌ Keine Personen gefunden. Prüfe, ob Namen in Spalte A stehen.")
            st.stop()

        ref_col_idx = col_to_index(ref_col_letters)  # 1-based

        # Urlaub zählen
        urlaub_genommen = {p: 0 for p in personen}
        ignored_weekend_or_holiday = {p: 0 for p in personen}  # optional debug

        for person, row_idx in person_rows.items():
            row = values[row_idx]

            # Achtung:
            # values() liefert 0-based Listenindex:
            # list index 0 = Spalte A (Name)
            # list index 1 = Spalte B
            # -> col_idx_1based = list_index + 1
            for list_idx in range(1, len(row)):  # ab Spalte B
                col_idx_1based = list_idx + 1
                d = date_for_column(col_idx_1based, ref_col_idx, ref_date)

                if d.year != 2026:
                    continue

                cell = str(row[list_idx]).strip().lower()
                if cell == "u":
                    if is_workday_berlin_2026(d):
                        urlaub_genommen[person] += 1
                    else:
                        ignored_weekend_or_holiday[person] += 1

        # Ergebnis
        rows = []
        for p in personen:
            genommen = urlaub_genommen.get(p, 0)
            rows.append({
                "Person": p,
                "Kontingent_2026": urlaubstage_pro_person,
                "Urlaub_2026_genommen": genommen,
                "Resturlaub_2026": urlaubstage_pro_person - genommen,
                "u_ignoriert_(Wochenende/Feiertag)": ignored_weekend_or_holiday.get(p, 0)
            })

        df = pd.DataFrame(rows)

        st.subheader("📊 Ergebnis – Urlaub 2026 (ohne Wochenenden/Feiertage)")
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

        with st.expander("Debug: Referenz-Check (erste 10 Tage um den 01.01.2026)"):
            # Zeige, ob die Spalten-Map sinnvoll ist
            ref = ref_col_idx
            debug = []
            for delta in range(-3, 7):
                col_num = ref + delta
                d = date_for_column(col_num, ref_col_idx, ref_date)
                debug.append({
                    "Spalte (Index 1-based)": col_num,
                    "Datum": d.isoformat(),
                    "Ist Arbeitstag?": is_workday_berlin_2026(d) if d.year == 2026 else None
                })
            st.dataframe(pd.DataFrame(debug), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Fehler: {e}")
