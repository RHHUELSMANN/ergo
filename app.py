
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

# OpenAI-Client initialisieren
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Hilfsfunktionen
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
        elif "." in text:
            parts = text.split(".")
            if len(parts[-1]) == 2:
                return datetime.strptime(text, "%d.%m.%y").date()
            return datetime.strptime(text, "%d.%m.%Y").date()
    except:
        return None

def lade_absätze_aus_pdf(pfad):
    doc = fitz.open(pfad)
    absatz_liste = []
    for seite in doc:
        text = seite.get_text()
        nummer = seite.number + 1
        absätze = [a.strip() for a in text.split("\n\n") if len(a.strip()) > 50]
        for absatz in absätze:
            absatz_liste.append({"seite": nummer, "text": absatz})
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
    relevante = suche_passende_absätze(frage, absatz_liste)
    kontext = "\n\n".join([f"Seite {a['seite']}:\n{a['text']}" for a in relevante])
    system_prompt = (
        "Du bist ein digitaler Versicherungsberater für Reisebüro Hülsmann. "
        "Beantworte ausschließlich Fragen zu Reiseversicherungen auf Grundlage der "
        "folgenden PDF-Auszüge. Wenn du es nicht beantworten kannst, sage: "
        "'Dazu liegt mir keine Information vor.'"
    )
    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Frage: {frage}\n\nPDF-Auszüge:\n{kontext}"}
        ],
        temperature=0.3
    )
    return response.choices[0].message.content, relevante

# UI-Setup
st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=250)
st.markdown("<h2 style='font-size:28px;'>Der Ergo Chuck – Berechnung & Angebot</h2>", unsafe_allow_html=True)

# Formular
with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    alter_text = st.text_input("Alter (z. B. 45 48)")

    # Geburtstagsfelder
    cols = st.columns(4)
    geb_eingaben = []
    for idx, col in enumerate(cols):
        with col:
            geb_eingaben.append(st.text_input("", key=f"gb{idx}", label_visibility="collapsed", placeholder=f"Geb. {idx+1}"))

    # Altersberechnung
    heute = date.today()
    geburts_alter = []
    for geb in geb_eingaben:
        d = parse_geburtstag(geb)
        if d:
            geburts_alter.append(heute.year - d.year - ((heute.month, heute.day) < (d.month, d.day)))
    if geburts_alter:
        st.markdown(f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(map(str, geburts_alter))}</small>", unsafe_allow_html=True)
        alter_text = " ".join(map(str, geburts_alter))

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    submit = st.form_submit_button("Tarife anzeigen")

# Nach Absenden
if submit:
    # Datum parsen
    von = parse_datum(von_raw)
    bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben.")
        st.stop()

    # Tariftabelle (Beispielwerte)
    df = pd.DataFrame([
        ["Reiserücktritt", "Einmal", "✓", "✓"],
        ["", "Jahres", "✓", "✓"],
        ["", "Sparfuchs", "✓", "✓"],
        ["Reisekranken", "Einmal", "✓", "✓"],
        ["", "Jahres", "✓", "✓"],
        ["RundumSorglos", "Einmal", "✓", "✓"],
        ["", "Jahres", "✓", "✓"],
    ], columns=["Produktgruppe", "Tarif", "mit SB", "ohne SB"])
    st.subheader("📊 Gruppierte Tarifübersicht")
    st.table(df)

    # Word-Angebot
    daten = {
        "Kundenname": name,
        "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
        "Reisepreis": f"{preis:,.2f} €".replace(".", ","),
        "Alter": alter_text,
        "Reiseziel": zielgebiet
    }
    st.session_state["word_daten"] = daten
    st.subheader("📄 Word-Angebot")
    if st.button("📄 Word-Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Vorlage 'angebot.docx' fehlt!")
        else:
            file_path = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", file_path, daten)
            with open(file_path, "rb") as f:
                btn_bytes = f.read()
            st.download_button("📥 Angebot herunterladen", btn_bytes, file_name=file_path)

    # GPT-Block
    frage = st.text_input("Frage an Versicherung (GPT)", placeholder="z. B. Was ist bei Corona versichert?")
    if frage:
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fundstellen = frage_an_gpt(frage, absätze)
        with st.expander("📄 Verwendete Textstellen aus PDF"):
            for i, a in enumerate(fundstellen, 1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"], width=600, placeholder=" …"), unsafe_allow_html=True)
        st.success(antwort)
