import os
import re
import textwrap
import difflib

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from openai import OpenAI
from word_styling import export_doc
from datetime import datetime, date

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def parse_datum(text):
    try:
        if "." in text:
            return datetime.strptime(text.strip(), "%d.%m.%Y").date()
        elif len(text.strip()) == 4:
            return date.today().replace(day=int(text[:2]), month=int(text[2:]))
        elif len(text.strip()) == 8:
            return datetime.strptime(text.strip(), "%d%m%Y").date()
    except:
        return None

def parse_geburtstag(text):
    try:
        text = text.strip().replace(" ", "")
        if len(text) == 6:
            return datetime.strptime(text, "%d%m%y").date()
        elif len(text) == 8:
            return datetime.strptime(text, "%d%m%Y").date()
        elif "." in text and len(text.split(".")[-1]) == 2:
            return datetime.strptime(text, "%d.%m.%y").date()
        elif "." in text:
            return datetime.strptime(text, "%d.%m.%Y").date()
    except:
        return None

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

def fmt(p, c, preis):
    p = float(p)
    betrag = p * preis if p < 1.0 else p
    return f"{betrag:.2f}".replace(".", ",") + f" € ({c})"

def first_hit(df):
    col = [c for c in df.columns if c.strip().lower() == "reisepreis bis"]
    if col:
        return df.sort_values(col[0]).iloc[0]
    return df.iloc[0]

def bereinige_export_daten(daten):
    export_daten = {}
    for key, value in daten.items():
        if isinstance(value, str) and "€" in value and "(" in value and ")" in value:
            betrag = value.split("€")[0].strip()
            export_daten[key] = betrag + " €"
        else:
            export_daten[key] = value
    return export_daten

# Formular
with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    alter_text = st.text_input("Alter (z. B. 45 48)")

    col1, col2, col3, col4 = st.columns(4)
    gb_eingaben = []
    with col1:
        gb_eingaben.append(st.text_input("", key="gb1", label_visibility="collapsed"))
    with col2:
        gb_eingaben.append(st.text_input("", key="gb2", label_visibility="collapsed"))
    with col3:
        gb_eingaben.append(st.text_input("", key="gb3", label_visibility="collapsed"))
    with col4:
        gb_eingaben.append(st.text_input("", key="gb4", label_visibility="collapsed"))

    heute = date.today()
    geburts_alter = []
    for geb in gb_eingaben:
        gebdat = parse_geburtstag(geb)
        if gebdat:
            alter = heute.year - gebdat.year - ((heute.month, heute.day) < (gebdat.month, gebdat.day))
            geburts_alter.append(alter)

    if geburts_alter:
        st.markdown(
            f"<small>👥 Berechnete Alter: {', '.join(str(a) for a in geburts_alter)}</small>",
            unsafe_allow_html=True
        )
        alter_text = " ".join(str(a) for a in geburts_alter)

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    submit = st.form_submit_button("Tarife anzeigen")

if submit:
    try:
        von = parse_datum(von_raw)
        bis = parse_datum(bis_raw)
        if not von or not bis:
            st.error("❌ Bitte gültige Datumswerte eingeben.")
            st.stop()

        alter_liste = [int(a) for a in alter_text.strip().split()]
        max_alter = max(alter_liste)
        altersgruppe = ermittle_altersgruppe(max_alter)
        personengruppe = ermittle_personengruppe(alter_liste)
        reisetage = berechne_reisetage(von, bis)

        excel = pd.ExcelFile("ergo.xlsx")
        blatt = {
            "rrv_ew_mit": "rrv-ev-mit", "rrv_ew_ohne": "rrv-ev-ohne",
            "rrv_jw_mit": "rrv-jv-mit", "rrv_jw_ohne": "rrv-jv-ohne",
            "rrv_jw_spf_mit": "rrv-jv-spf-mit", "rrv_jw_spf_ohne": "rrv-jv-spf-ohne",
            "kv_ew_mit": "kv-ev-mit", "kv_ew_ohne": "kv-ev-ohne",
            "kv_jw_mit": "kv-jv-mit", "kv_jw_ohne": "kv-jv-ohne",
            "rus_ew_mit": "rus-ev-mit", "rus_ew_ohne": "rus-ev-ohne",
            "rus_jw_mit": "rus-jv-mit", "rus_jw_ohne": "rus-jv-ohne"
        }
        t = {key: excel.parse(name) for key, name in blatt.items()}

        def tarif(d, *args): return fmt(*args, preis) if not d.empty else "–"
        def first(d, cond, keys): f = d[cond]; return tarif(f, *first_hit(f)[keys]) if not f.empty else "–"

        def kv_einmal(name):
            d = t[name]
            f = d[(d["Zielgebiet"].str.strip().str.lower() == zielgebiet.lower()) &
                  (d["Altersgruppe"].str.strip() == altersgruppe) &
                  (d["Personengruppe"].str.strip().str.lower() == personengruppe.lower())]
            if not f.empty:
                tagespreis = float(f.iloc[0]["Tagesprämie"])
                return fmt(round(reisetage * tagespreis, 2), f.iloc[0]["Tarifcode"], preis)
            return "–"

        daten = {
            "Kundenname": name,
            "Reisedatum": f"{von.strftime('%d.%m.%Y')} – {bis.strftime('%d.%m.%Y')}",
            "Reisepreis": f"{preis:,.2f} €".replace(".", ","),
            "Anzahl": str(len(alter_liste)),
            "Alter": ", ".join(str(a) for a in alter_liste),
            "Reiseziel": zielgebiet,
            "Reiseruecktritt_Einmal_mit_SB": first(t["rrv_ew_mit"], t["rrv_ew_mit"]["Reisepreis bis"] >= preis, ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Einmal_ohne_SB": first(t["rrv_ew_ohne"], (t["rrv_ew_ohne"]["Reisepreis bis"] >= preis) & (t["rrv_ew_ohne"]["Altersgruppe"].str.strip() == altersgruppe), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_mit_SB": first(t["rrv_jw_mit"], (t["rrv_jw_mit"]["Reisepreis bis"] >= preis) & (t["rrv_jw_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_ohne_SB": first(t["rrv_jw_ohne"], (t["rrv_jw_ohne"]["Reisepreis bis"] >= preis) & (t["rrv_jw_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_Sparfuchs_mit_SB": first(t["rrv_jw_spf_mit"], (t["rrv_jw_spf_mit"]["Reisepreis bis"] >= preis) & (t["rrv_jw_spf_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_spf_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_Sparfuchs_ohne_SB": first(t["rrv_jw_spf_ohne"], (t["rrv_jw_spf_ohne"]["Reisepreis bis"] >= preis) & (t["rrv_jw_spf_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_spf_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reisekranken_Einmal_mit_SB": kv_einmal("kv_ew_mit"),
            "Reisekranken_Einmal_ohne_SB": kv_einmal("kv_ew_ohne"),
            "Reisekranken_Jahres_mit_SB": first(t["kv_jw_mit"], (t["kv_jw_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["kv_jw_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reisekranken_Jahres_ohne_SB": first(t["kv_jw_ohne"], (t["kv_jw_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["kv_jw_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Einmal_mit_SB": first(t["rus_ew_mit"], (t["rus_ew_mit"]["Reisepreis bis"] >= preis) & (t["rus_ew_mit"]["Zielgebiet"].str.lower().str.strip() == zielgebiet.lower()), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Einmal_ohne_SB": first(t["rus_ew_ohne"], (t["rus_ew_ohne"]["Reisepreis bis"] >= preis) & (t["rus_ew_ohne"]["Zielgebiet"].str.lower().str.strip() == zielgebiet.lower()) & (t["rus_ew_ohne"]["Altersgruppe"].str.strip() == altersgruppe), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Jahres_mit_SB": first(t["rus_jw_mit"], (t["rus_jw_mit"]["Reisepreis bis"] >= preis) & (t["rus_jw_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rus_jw_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Jahres_ohne_SB": first(t["rus_jw_ohne"], (t["rus_jw_ohne"]["Reisepreis bis"] >= preis) & (t["rus_jw_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rus_jw_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
        }

        st.session_state["word_daten"] = daten

    except Exception as e:
        st.error(f"❌ Fehler: {e}")
