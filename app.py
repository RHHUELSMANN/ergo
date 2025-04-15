import streamlit as st
import pandas as pd
from datetime import date, timedelta

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("SmallLogoBW_png.png", width=250)
st.title("🦾 Der Ergo Chuck – Gruppierte Tarife (fix für KV Einmal)")

EUROPA_CODES = ["PMI", "FRA", "BER", "VIE", "ZRH", "LIS", "CDG", "AMS", "BCN", "ROM"]
WELT_CODES = ["PUJ", "BKK", "JFK", "LAX", "DXB", "CUN", "MEX", "CPT", "SIN", "HND"]

def ermittle_zielgebiet(code):
    code = code.upper()
    if code in EUROPA_CODES: return "Europa"
    if code in WELT_CODES: return "Welt"
    return "Unbekannt"

def ermittle_altersgruppe(alter):
    if alter <= 40: return "bis 40 Jahre"
    elif 41 <= alter <= 64: return "41–64 Jahre"
    return "ab 65 Jahre"

def ermittle_personengruppe(alter_liste):
    if len(alter_liste) == 1: return "Einzelperson"
    elif len(alter_liste) == 2: return "Paar"
    return "Familie"

def berechne_reisetage(von, bis):
    return max(1, (bis - von).days + 1)

def fmt(p, c, preis):
    p = float(p)
    betrag = p * preis if p < 1.0 else p
    return f"{betrag:.2f}".replace(".", ",") + f" € ({c})"

def first_hit(df):
    col = [c for c in df.columns if c.strip().lower() == "reisepreis bis"]
    if col:
        return df.sort_values(col[0]).iloc[0]
    return df.iloc[0]

# Eingabe
ziel = st.text_input("🌍 Reiseziel (IATA-Code)")
preis = st.number_input("💶 Reisepreis (€)", min_value=0.0)
alter_text = st.text_input("👥 Alter der Reisenden (z. B. 45 48)")
von = st.date_input("📅 Reise von", value=date.today(), format="DD.MM.YYYY")
bis = st.date_input("📅 Reise bis", value=date.today() + timedelta(days=7), format="DD.MM.YYYY")

if st.button("✅ Gruppierte Tarife anzeigen"):
    try:
        alter_liste = [int(a) for a in alter_text.strip().split()]
        max_alter = max(alter_liste)
        altersgruppe = ermittle_altersgruppe(max_alter)
        personengruppe = ermittle_personengruppe(alter_liste)
        zielgebiet = ermittle_zielgebiet(ziel)
        reisetage = berechne_reisetage(von, bis)

        blatt = {
            "rrv_ew_mit": "rrv-ev-mit", "rrv_ew_ohne": "rrv-ev-ohne",
            "rrv_jw_mit": "rrv-jv-mit", "rrv_jw_ohne": "rrv-jv-ohne",
            "rrv_jw_spf_mit": "rrv-jv-spf-mit", "rrv_jw_spf_ohne": "rrv-jv-spf-ohne",
            "kv_ew_mit": "kv-ev-mit", "kv_ew_ohne": "kv-ev-ohne",
            "kv_jw_mit": "kv-jv-mit", "kv_jw_ohne": "kv-jv-ohne",
            "rus_ew_mit": "rus-ev-mit", "rus_ew_ohne": "rus-ev-ohne",
            "rus_jw_mit": "rus-jv-mit", "rus_jw_ohne": "rus-jv-ohne"
        }

        excel = pd.ExcelFile("ergo.xlsx")
        t = {key: excel.parse(name) for key, name in blatt.items()}

        def tarif(d, *args): return fmt(*args, preis) if not d.empty else "–"
        def first(d, cond, keys): f = d[cond]; return tarif(f, *first_hit(f)[keys]) if not f.empty else "–"

        def kv_einmal(name):
            d = t[name]
            f = d[
                (d["Zielgebiet"].str.lower().str.strip() == zielgebiet.lower()) &
                (d["Altersgruppe"].str.strip() == altersgruppe) &
                (d["Personengruppe"].str.lower().str.strip() == personengruppe.lower())
            ]
            if not f.empty:
                tagespreis = float(f.iloc[0]["Tagesprämie"])
                return fmt(reisetage * tagespreis, f.iloc[0]["Tarifcode"], preis)
            return "–"

        rows = [
            ["Reisekranken", "Einmal",
             kv_einmal("kv_ew_mit"),
             kv_einmal("kv_ew_ohne")],
        ]

        df = pd.DataFrame(rows, columns=["Produktgruppe", "Tarif", "mit SB", "ohne SB"])
        st.subheader("📊 Krankenversicherung (Einmal)")
        st.table(df)

    except Exception as e:
        st.error(f"❌ Fehler: {e}")


