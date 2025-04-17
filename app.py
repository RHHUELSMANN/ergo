
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
    if alter <= 40: return "bis 40 Jahre"
    elif 41 <= alter <= 64: return "41–64 Jahre"
    return "ab 65 Jahre"

def ermittle_personengruppe(alter_liste):
    if len(alter_liste) == 1: return "Einzelperson"
    elif len(alter_liste) == 2: return "Paar"
    return "Familie"

def berechne_reisetage(von, bis):
    return max(1, (bis - von).days + 1)

def lade_absätze_aus_pdf(pfad):
    doc = fitz.open(pfad)
    absatz_liste = []
    for seite in doc:
        text = seite.get_text()
        nummer = seite.number + 1
        absätze = [a.strip() for a in text.split("\n\n") if len(a.strip()) > 50]
        for absatz in absätze:
            absatz_liste.append({ "seite": nummer, "text": absatz })
    return absatz_liste

def suche_passende_absätze(frage, absätze, anzahl=3):
    frage = frage.lower()
    scored = []
    for absatz in absätze:
        text = absatz["text"].lower()
        score = difflib.SequenceMatcher(None, frage, text).ratio()
        if any(wort in text for wort in frage.split()):
            score += 0.2
        scored.append((score, absatz))
    scored.sort(reverse=True, key=lambda x: x[0])
    return [a for _, a in scored[:anzahl]]

def frage_an_gpt(frage, absatz_liste):
    relevante_abschnitte = suche_passende_absätze(frage, absatz_liste)
    kontext = "\n\n".join([f"Seite {a['seite']}:\n{a['text']}" for a in relevante_abschnitte])
    system_prompt = (
        "Du bist ein digitaler Versicherungsberater für Reisebüro Hülsmann. "
        "Beantworte ausschließlich Fragen zu Reiserücktritts-, Reisekranken- oder RundumSorglos-Versicherungen "
        "auf Grundlage der folgenden PDF-Auszüge.\n\n"
        "Berücksichtige bei der Interpretation auch Begriffe mit ähnlicher Bedeutung.\n"
        "- Selbstbeteiligung ≈ Selbstbehalt ≈ SB ≈ Eigenanteil\n"
        "- Reiserücktritt ≈ Rücktritt ≈ Stornierung\n"
        "- Krankheit ≈ Corona ≈ COVID ≈ Quarantäne\n\n"
        "Wenn du keine ausreichende Information findest, sage bitte klar: 'Dazu liegt mir keine Information vor.'"
    )
    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Frage: {frage}\n\nPDF-Auszüge:\n{kontext}"}
        ],
        temperature=0.3
    )
    return response.choices[0].message.content, relevante_abschnitte

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")

st.image("logo.png", width=250)
st.markdown("<h2 style='font-size:28px;'>Der Ergo Chuck – Berechnung & Angebot</h2>", unsafe_allow_html=True)

# Formular
with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    alter_text = st.text_input("Alter (z. B. 45 48)")

    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
    gb_eingaben = []
    for i, col in enumerate([col1, col2, col3, col4]):
        with col:
            gb_eingaben.append(st.text_input("", key=f"gb{i}", label_visibility="collapsed", placeholder=f"Geb. {i+1}"))

    heute = date.today()
    geburts_alter = []
    for geb in gb_eingaben:
        gebdat = parse_geburtstag(geb)
        if gebdat:
            alter = heute.year - gebdat.year - ((heute.month, heute.day) < (gebdat.month, gebdat.day))
            geburts_alter.append(alter)

    if geburts_alter:
        st.markdown(f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(str(a) for a in geburts_alter)}</small>", unsafe_allow_html=True)
        alter_text = " ".join(str(a) for a in geburts_alter)

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage_gpt = st.text_input("GPT-Frage zur Versicherung", placeholder="Was ist bei Corona versichert?")
    submit = st.form_submit_button("Tarife anzeigen")
