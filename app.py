import streamlit as st
import pandas as pd
from datetime import date, timedelta

# Seiten-Setup
st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("SmallLogoBW_png.png", width=250)
st.title("🦾 Der Ergo Chuck – Tarifrechner")

# Zielgebiet-Logik
EUROPA_CODES = ["PMI", "FRA", "BER", "VIE", "ZRH", "LIS", "CDG", "AMS", "BCN", "ROM"]
WELT_CODES = ["PUJ", "BKK", "JFK", "LAX", "DXB", "CUN", "MEX", "CPT", "SIN", "HND"]

def ermittle_zielgebiet(code):
    code = code.upper()
    if code in EUROPA_CODES:
        return "Europa"
    elif code in WELT_CODES:
        return "Welt"
    return "Unbekannt"

def ermittle_altersgruppe(alter):
    if alter <= 40:
        return "bis 40 Jahre"
    elif 41 <= alter <= 64:
        return "41–64 Jahre"
    return "ab 65 Jahre"

def ermittle_personengruppe(alter_liste):
    if len(alter_liste) == 1:
        return "Einzelperson"
    elif len(alter_liste) == 2:
        return "Paar"
    return "Familie"

def berechne_reisetage(von, bis):
    return max(1, (bis - von).days + 1)

def first_hit(df): return df.sort_values("Reisepreis bis").iloc[0]

def fmt(p, c, preis):
    p = float(p)
    betrag = p * preis if p < 1.0 else p
    return f"{betrag:.2f}".replace(".", ",") + f" € ({c})"

# Eingabefelder
name = st.text_input("👤 Kundenname")
ziel = st.text_input("🌍 Reiseziel (IATA-Code)")
preis = st.number_input("💶 Reisepreis (€)", min_value=0.0, step=10.0)
alter_text = st.text_input("👥 Alter der Reisenden (z. B. 45 48)")
von = st.date_input("📅 Reise von", value=date.today(), format="DD.MM.YYYY")
bis = st.date_input("📅 Reise bis", value=date.today() + timedelta(days=7), format="DD.MM.YYYY")

if st.button("✅ Tarife berechnen"):
    try:
        alter_liste = [int(a) for a in alter_text.strip().split()]
        max_alter = max(alter_liste)
        altersgruppe = ermittle_altersgruppe(max_alter)
        personengruppe = ermittle_personengruppe(alter_liste)
        zielgebiet = ermittle_zielgebiet(ziel)
        reisetage = berechne_reisetage(von, bis)

        # Excel laden
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

        def rrv_ew_mit():
            d = t["rrv_ew_mit"]
            f = d[d["Reisepreis bis"] >= preis]
            return fmt(*first_hit(f)[["Prämie", "Tarifcode"]], preis) if not f.empty else "–"

        def rrv_ew_ohne():
            d = t["rrv_ew_ohne"]
            f = d[(d["Reisepreis bis"] >= preis) & (d["Altersgruppe"].str.strip() == altersgruppe)]
            return fmt(*first_hit(f)[["Prämie", "Tarifcode"]], preis) if not f.empty else "–"

        def rrv_jahres(name):
            d = t[name]
            f = d[
                (d["Reisepreis bis"] >= preis) &
                (d["Altersgruppe"].str.strip() == altersgruppe) &
                (d["Personengruppe"].str.strip().str.lower() == personengruppe.lower())
            ]
            return fmt(*first_hit(f)[["Prämie", "Tarifcode"]], preis) if not f.empty else "–"

        def kv_einmal(name):
            d = t[name]
            f = d[
                (d["Zielgebiet"].str.strip().str.lower() == zielgebiet.lower()) &
                (d["Altersgruppe"].str.strip() == altersgruppe) &
                (d["Personengruppe"].str.strip().str.lower() == personengruppe.lower())
            ]
            return fmt(round(reisetage * float(f.iloc[0]["Tagesprämie"]), 2), f.iloc[0]["Tarifcode"], preis) if not f.empty else "–"

        def kv_jahres(name):
            d = t[name]
            f = d[
                (d["Altersgruppe"].str.strip() == altersgruppe) &
                (d["Personengruppe"].str.strip().str.lower() == personengruppe.lower())
            ]
            return fmt(*f.iloc[0][["Prämie", "Tarifcode"]], preis) if not f.empty else "–"

        def rus_ew(name, mit_sb):
            d = t[name]
            cond = (
                (d["Reisepreis bis"] >= preis) &
                (d["Zielgebiet"].str.strip().str.lower() == zielgebiet.lower())
            )
            if not mit_sb:
                cond &= d["Altersgruppe"].str.strip() == altersgruppe
            f = d[cond]
            return fmt(*first_hit(f)[["Prämie", "Tarifcode"]], preis) if not f.empty else "–"

        def rus_jahres(name):
            d = t[name]
            f = d[
                (d["Reisepreis bis"] >= preis) &
                (d["Altersgruppe"].str.strip() == altersgruppe) &
                (d["Personengruppe"].str.strip().str.lower() == personengruppe.lower())
            ]
            return fmt(*first_hit(f)[["Prämie", "Tarifcode"]], preis) if not f.empty else "–"

        rows = [
            ("Reiserücktritt", rrv_ew_mit(), rrv_ew_ohne()),
            ("Jahres", rrv_jahres("rrv_jw_mit"), rrv_jahres("rrv_jw_ohne")),
            ("Jahres – Sparfuchs", rrv_jahres("rrv_jw_spf_mit"), rrv_jahres("rrv_jw_spf_ohne")),
            ("Reisekranken", kv_einmal("kv_ew_mit"), kv_einmal("kv_ew_ohne")),
            ("Jahresversicherung", kv_jahres("kv_jw_mit"), kv_jahres("kv_jw_ohne")),
            ("RundumSorglos", rus_ew("rus_ew_mit", True), rus_ew("rus_ew_ohne", False)),
            ("Jahres", rus_jahres("rus_jw_mit"), rus_jahres("rus_jw_ohne")),
        ]
        df = pd.DataFrame(rows, columns=["Produkt", "mit SB", "ohne SB"])
        st.subheader("📊 Ergebnisse")
        st.table(df)

    except Exception as e:
        st.error(f"❌ Fehler bei der Berechnung: {e}")

