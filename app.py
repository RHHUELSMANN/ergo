
import streamlit as st
st.set_page_config(page_title="Ergo Berechnung & Angebot", page_icon="🧾")
import pandas as pd
from datetime import datetime
from word_export import export_word_dokument

@st.cache_resource
def load_excel():
    return pd.ExcelFile("ergo.xlsx")

excel_data = load_excel()

def berechne_praemie(p, preis=None, tage=None):
    p = float(p)
    if tage is not None:
        return round(p * tage, 2)
    return round(p, 2)

def berechne_tage(von, bis):
    default_year = "2025"
    if len(von) <= 4:
        von = von + default_year
    if len(bis) <= 4:
        bis = bis + default_year
    for fmt in ("%d%m%Y", "%d.%m.%Y", "%d%m", "%d.%m."):
        try:
            d1 = datetime.strptime(von, fmt)
            d2 = datetime.strptime(bis, fmt)
            return (d2 - d1).days + 1
        except:
            continue
    return 0

def altersgruppe(alter):
    return "bis 40 Jahre" if alter <= 40 else "41–64 Jahre" if alter <= 64 else "ab 65 Jahre"

def personengruppe(alter_liste):
    return "Einzelperson" if len(alter_liste) == 1 else "Paar" if len(alter_liste) == 2 else "Familie"

def first_hit(df):
    return df.sort_values("Reisepreis bis").iloc[0]

def fmt(p, reisepreis=None, tarifcode=None):
    p = float(p)
    betrag = p * reisepreis if reisepreis is not None and p < 1 else p
    text = f"{betrag:,.2f}".replace(".", "X").replace(",", ".").replace("X", ",")
    return f"{text} ({tarifcode})" if tarifcode else text

def tarif(df, bedingungen):
    for k, v in bedingungen.items():
        if k in df.columns:
            df = df[df[k].astype(str).str.strip().str.lower() == v.lower()]
    return df.iloc[0] if not df.empty else None

blatt = {key: excel_data.parse(key) for key in excel_data.sheet_names}

st.image("logo.png", width=200)
st.markdown("<h2 style='font-size:22pt;'>Reiseversicherung // Berechnung und Angebot</h2>", unsafe_allow_html=True)
st.write("")

reise_von = st.text_input("Reisedatum von")
reise_bis = st.text_input("Reisedatum bis")
raw_reisepreis = st.text_input("Reisepreis").replace(",", ".")
alter_input = st.text_input("Alter")

zielgebiet = st.selectbox("Zielgebiet", ["Europa", "Welt"])
name = st.text_input("Name (für Word-Angebot)", key="word_export_name")

def parse_geburtsdatum(text):
    heute = datetime.today()
    for fmt in ("%d%m%Y", "%d.%m.%Y", "%d%m%y", "%d.%m.%y"):
        try:
            geb = datetime.strptime(text.strip(), fmt)
            if geb > heute:
                geb = geb.replace(year=geb.year - 100)
            alter = heute.year - geb.year - ((heute.month, heute.day) < (geb.month, geb.day))
            return str(alter)
        except:
            continue
    return "-"

col1, col2, col3, col4 = st.columns(4)
with col1:
    geb1 = st.text_input("Geb. 1", key="geb1_input")
    st.markdown(f"<div style='font-size: 13px;'>Alter: {parse_geburtsdatum(geb1)}</div>", unsafe_allow_html=True)
with col2:
    geb2 = st.text_input("Geb. 2", key="geb2_input")
    st.markdown(f"<div style='font-size: 13px;'>Alter: {parse_geburtsdatum(geb2)}</div>", unsafe_allow_html=True)
with col3:
    geb3 = st.text_input("Geb. 3", key="geb3_input")
    st.markdown(f"<div style='font-size: 13px;'>Alter: {parse_geburtsdatum(geb3)}</div>", unsafe_allow_html=True)
with col4:
    geb4 = st.text_input("Geb. 4", key="geb4_input")
    st.markdown(f"<div style='font-size: 13px;'>Alter: {parse_geburtsdatum(geb4)}</div>", unsafe_allow_html=True)

try:
    reisepreis = float(raw_reisepreis)
except:
    reisepreis = None

if st.button("Prämie berechnen"):
    if reisepreis and alter_input and reise_von and reise_bis:
        alter_liste = [int(a) for a in alter_input.split()]
        max_alter = max(alter_liste)
        a_gruppe = altersgruppe(max_alter)
        p_gruppe = personengruppe(alter_liste)
        tage = berechne_tage(reise_von, reise_bis)

        kv_ev_mit_tarif = tarif(blatt["kv-ev-mit"], {"Zielgebiet": zielgebiet, "Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        kv_ev_ohne_tarif = tarif(blatt["kv-ev-ohne"], {"Zielgebiet": zielgebiet, "Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})

        kranken_mit = berechne_praemie(kv_ev_mit_tarif["Prämie"], tage=tage)
        kranken_ohne = berechne_praemie(kv_ev_ohne_tarif["Prämie"], tage=tage)

        rrv_ev_mit = first_hit(blatt["rrv-ev-mit"][blatt["rrv-ev-mit"]["Reisepreis bis"] >= reisepreis])["Prämie"]
        rrv_ev_ohne_df = blatt["rrv-ev-ohne"][(blatt["rrv-ev-ohne"]["Altersgruppe"] == a_gruppe) & (blatt["rrv-ev-ohne"]["Reisepreis bis"] >= reisepreis)]
        rrv_ev_ohne = first_hit(rrv_ev_ohne_df)["Prämie"]

        rus_ev_mit = first_hit(blatt["rus-ev-mit"][(blatt["rus-ev-mit"]["Zielgebiet"] == zielgebiet) & (blatt["rus-ev-mit"]["Reisepreis bis"] >= reisepreis)])["Prämie"]
        rus_ev_ohne = first_hit(blatt["rus-ev-ohne"][(blatt["rus-ev-ohne"]["Zielgebiet"] == zielgebiet) & (blatt["rus-ev-ohne"]["Altersgruppe"] == a_gruppe) & (blatt["rus-ev-ohne"]["Reisepreis bis"] >= reisepreis)])["Prämie"]

        kv_jv_mit = tarif(blatt["kv-jv-mit"], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        kv_jv_ohne = tarif(blatt["kv-jv-ohne"], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})

        rrv_jv_mit = tarif(blatt["rrv-jv-mit"][blatt["rrv-jv-mit"]["Reisepreis bis"] >= reisepreis], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        rrv_jv_ohne = tarif(blatt["rrv-jv-ohne"][blatt["rrv-jv-ohne"]["Reisepreis bis"] >= reisepreis], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        spf_mit = tarif(blatt["rrv-jv-spf-mit"][blatt["rrv-jv-spf-mit"]["Reisepreis bis"] >= reisepreis], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        spf_ohne = tarif(blatt["rrv-jv-spf-ohne"][blatt["rrv-jv-spf-ohne"]["Reisepreis bis"] >= reisepreis], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        rus_jv_mit = tarif(blatt["rus-jv-mit"][blatt["rus-jv-mit"]["Reisepreis bis"] >= reisepreis], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})
        rus_jv_ohne = tarif(blatt["rus-jv-ohne"][blatt["rus-jv-ohne"]["Reisepreis bis"] >= reisepreis], {"Altersgruppe": a_gruppe, "Personengruppe": p_gruppe})

        rows = [
            ("Reiserücktritt Einmal", (rrv_ev_mit, "RNM104"), (rrv_ev_ohne, "RNX104")),
            ("Reiserücktritt Jahres", (rrv_jv_mit["Prämie"], rrv_jv_mit["Tarifcode"]), (rrv_jv_ohne["Prämie"], rrv_jv_ohne["Tarifcode"])),
            ("Reiserücktritt Jahres- Sparfuchs", (spf_mit["Prämie"], spf_mit["Tarifcode"]), (spf_ohne["Prämie"], spf_ohne["Tarifcode"])),
            ("Kranken Einmal", (kranken_mit, kv_ev_mit_tarif["Tarifcode"]), (kranken_ohne, kv_ev_ohne_tarif["Tarifcode"])),
            ("Kranken Jahres", (berechne_praemie(kv_jv_mit["Prämie"]), kv_jv_mit["Tarifcode"]), (berechne_praemie(kv_jv_ohne["Prämie"]), kv_jv_ohne["Tarifcode"])),
            ("RundumSorglos Einmal", (rus_ev_mit, "PNM104"), (rus_ev_ohne, "PNX104")),
            ("RundumSorglos Jahres", (rus_jv_mit["Prämie"], rus_jv_mit["Tarifcode"]), (rus_jv_ohne["Prämie"], rus_jv_ohne["Tarifcode"]))
        ]
        df = pd.DataFrame(rows, columns=["Produkt", "mit SB", "ohne SB"])
        df["mit SB"] = df["mit SB"].apply(lambda x: fmt(x[0], reisepreis, tarifcode=x[1]))
        df["ohne SB"] = df["ohne SB"].apply(lambda x: fmt(x[0], reisepreis, tarifcode=x[1]))
        st.session_state["ergebnis_df"] = df
        st.session_state["word_data"] = {
            "name": name, "reisepreis": reisepreis, "reise_von": reise_von, "reise_bis": reise_bis,
            "zielgebiet": zielgebiet, "alter_liste": alter_liste, "max_alter": max_alter,
            "rrv_ev_mit": rrv_ev_mit, "rrv_ev_ohne": rrv_ev_ohne,
            "rrv_jv_mit": rrv_jv_mit["Prämie"], "rrv_jv_ohne": rrv_jv_ohne["Prämie"],
            "spf_mit": spf_mit["Prämie"], "spf_ohne": spf_ohne["Prämie"],
            "kranken_mit": kranken_mit, "kranken_ohne": kranken_ohne,
            "kv_jv_mit": kv_jv_mit["Prämie"], "kv_jv_ohne": kv_jv_ohne["Prämie"],
            "rus_ev_mit": rus_ev_mit, "rus_ev_ohne": rus_ev_ohne,
            "rus_jv_mit": rus_jv_mit["Prämie"], "rus_jv_ohne": rus_jv_ohne["Prämie"]
        }

if "ergebnis_df" in st.session_state:
    st.table(st.session_state["ergebnis_df"])

if "word_data" in st.session_state:
    if st.button("Angebot als Word-Datei erstellen"):
        export_word_dokument(**st.session_state["word_data"])
