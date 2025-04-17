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

# ————————————————
# OpenAI-Client initialisieren
# Stelle sicher, dass in Streamlit-Secrets OPENAI_API_KEY gesetzt ist
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ————————————————
# Hilfsfunktionen

def parse_datum(text):
    """Parst 'TT.MM.JJJJ', 'TTMMJJJJ' oder 'TTMM' (Jahr heute)"""
    try:
        t = text.strip()
        if "." in t:
            #  dd.mm.yyyy
            return datetime.strptime(t, "%d.%m.%Y").date()
        elif len(t) == 8:
            # ddmmyyyy
            return datetime.strptime(t, "%d%m%Y").date()
        elif len(t) == 4:
            # ddMM – Jahr = heute.year
            dd, mm = int(t[:2]), int(t[2:])
            return date(date.today().year, mm, dd)
    except:
        return None

def parse_geburtstag(text):
    """Parst Geburtstage in Formaten 6-stellig (ddmmyy), 8-stellig (ddmmyyyy)"""
    try:
        t = text.strip().replace(" ", "").replace(".", "")
        if len(t) == 6:
            return datetime.strptime(t, "%d%m%y").date()
        elif len(t) == 8:
            return datetime.strptime(t, "%d%m%Y").date()
    except:
        return None

def lade_absätze_aus_pdf(pfad):
    doc = fitz.open(pfad)
    absätze = []
    for seite in doc:
        text = seite.get_text()
        seiten_nr = seite.number + 1
        # Absätze nach Doppelte Newlines
        for absatz in text.split("\n\n"):
            a = absatz.strip().replace("\n", " ")
            if len(a) > 50:
                absätze.append({"seite": seiten_nr, "text": a})
    return absätze

def suche_passende_absätze(frage, absätze, topk=3):
    frage_lc = frage.lower()
    scores = []
    for a in absätze:
        txt = a["text"].lower()
        # Einfaches Similarity-Scoring
        sc = difflib.SequenceMatcher(None, frage_lc, txt).ratio()
        if any(w in txt for w in frage_lc.split()):
            sc += 0.2
        scores.append((sc, a))
    scores.sort(reverse=True, key=lambda x: x[0])
    return [a for (_, a) in scores[:topk]]

def frage_an_gpt(frage, absätze):
    relevant = suche_passende_absätze(frage, absätze)
    kontext = "\n\n".join([f"Seite {a['seite']}:\n{a['text']}" for a in relevant])
    system_prompt = (
        "Du bist ein digitaler Versicherungsberater für Reisebüro Hülsmann. "
        "Beantworte ausschließlich Fragen zu Reiserücktritts-, Reisekranken- oder RundumSorglos-Versicherungen "
        "auf Grundlage der folgenden PDF-Auszüge. Wenn du es nicht beantworten kannst, sage: "
        "'Dazu liegt mir keine Information vor.'"
    )
    resp = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"Frage: {frage}\n\nPDF-Auszüge:\n{kontext}"}
        ],
        temperature=0.3
    )
    return resp.choices[0].message.content, relevant

# ————————————————
# Streamlit-Oberfläche

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=250)
st.markdown("<h2>Der Ergo Chuck – Berechnung & Angebot</h2>", unsafe_allow_html=True)

# — Formular — 
with st.form("eingabeformular"):
    name        = st.text_input("Kundenname")
    zielgebiet  = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis       = st.number_input("Reisepreis (€)", min_value=0.0)
    # Fallback manueller Alterseintrag
    alter_text  = st.text_input("Alter (z. B. 45 48)", key="alter_manual")

    # ─── Geburtstagsfelder (4 Spalten) ───
    cols = st.columns(4)
    geb_eingaben = []
    for i, col in enumerate(cols, start=1):
        with col:
            geb_eingaben.append(
                st.text_input(
                    "", 
                    key=f"gb{i}", 
                    label_visibility="collapsed", 
                    placeholder=f"Geb. {i}"
                )
            )
    # ──────────────────────────────────────

    # Automatische Altersberechnung aus den Geburtstagen
    heute = date.today()
    geburts_alter = []
    for geb in geb_eingaben:
        d = parse_geburtstag(geb)
        if d:
            age = heute.year - d.year - ((heute.month, heute.day) < (d.month, d.day))
            geburts_alter.append(age)
    if geburts_alter:
        st.markdown(
            f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(map(str, geburts_alter))}</small>",
            unsafe_allow_html=True
        )
        alter_text = " ".join(map(str, geburts_alter))

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage   = st.text_input("GPT-Frage zur Versicherung", placeholder="z. B. Was ist bei Corona versichert?")
    submit  = st.form_submit_button("Tarife anzeigen")

# — Ausführung nach Klick —
if submit:
    # Datum validieren
    von = parse_datum(von_raw)
    bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben.")
        st.stop()

    # Beispiel: Lade Excel, berechne Tarife (hier ersetzen durch deine Logik)
    excel = pd.ExcelFile("ergo.xlsx")
    # … deine Tarif-Logik hier …

    # Beispiel-Tariftabelle
    df = pd.DataFrame([
        ["Reiserücktritt", "Einmal", "123,45 €", "234,56 €"],
        ["",              "Jahres", "345,67 €", "456,78 €"],
        ["Reisekranken",   "Einmal", " 89,01 €", " 90,12 €"],
        ["RundumSorglos",  "Einmal", " 34,56 €", " 45,67 €"]
    ], columns=["Produktgruppe", "Tarif", "mit SB", "ohne SB"])

    st.subheader("📊 Gruppierte Tarifübersicht")
    st.table(df)

    # Word-Export
    daten = {
        "Kundenname": name,
        "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
        "Reisepreis": f"{preis:,.2f} €".replace(".", ","),
        "Alter": alter_text,
        "Reiseziel": zielgebiet
    }
    st.session_state["word_daten"] = daten
    st.subheader("📄 Word-Angebot")
    if st.button("📄 Word-Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Vorlage 'angebot.docx' fehlt!")
        else:
            fp = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", fp, daten)
            with open(fp, "rb") as f:
                content = f.read()
            st.download_button("📥 Angebot herunterladen", content, file_name=fp)

    # GPT-Antwort unter der Tabelle
    if frage.strip():
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fund = frage_an_gpt(frage, absätze)
        with st.expander("📄 Verwendete Textstellen aus PDF"):
            for i, a in enumerate(fund, 1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"], width=600, placeholder=" …"),
                            unsafe_allow_html=True)
        st.success(antwort)
