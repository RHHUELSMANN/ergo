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
        t = text.strip().replace(" ", "").replace(".", "")
        if len(t) == 6:
            return datetime.strptime(t, "%d%m%y").date()
        elif len(t) == 8:
            return datetime.strptime(t, "%d%m%Y").date()
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
        txt = absatz["text"].lower()
        score = difflib.SequenceMatcher(None, frage, txt).ratio()
        if any(w in txt for w in frage.split()):
            score += 0.2
        scored.append((score, absatz))
    scored.sort(reverse=True, key=lambda x: x[0])
    return [a for _, a in scored[:anzahl]]

def frage_an_gpt(frage, absatz_liste):
    relevante = suche_passende_absätze(frage, absatz_liste)
    kontext = "\n\n".join([f"Seite {a['seite']}:\n{a['text']}" for a in relevante])
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
    return resp.choices[0].message.content, relevante

# Streamlit UI
st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=250)
st.markdown("<h2 style='font-size:28px;'>Der Ergo Chuck – Berechnung & Angebot</h2>", unsafe_allow_html=True)

# Formularbereich
with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    # manuelles Alter als Fallback
    alter_text = st.text_input("Alter (z. B. 45 48)", key="alter_manual")

    # ─── Geburtstagsfelder ───
    cols = st.columns(4)
    geb_eingaben = []
    for idx, col in enumerate(cols, start=1):
        with col:
            geb_eingaben.append(
                st.text_input(
                    "", 
                    key=f"gb{idx}", 
                    label_visibility="collapsed", 
                    placeholder=f"Geb. {idx}"
                )
            )
    # ──────────────────────────

    # Automatische Altersberechnung
    heute = date.today()
    geburts_alter = []
    for geb in geb_eingaben:
        d = parse_geburtstag(geb)
        if d:
            geburts_alter.append(heute.year - d.year - ((heute.month, heute.day) < (d.month, d.day)))
    if geburts_alter:
        st.markdown(
            f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(map(str, geburts_alter))}</small>",
            unsafe_allow_html=True
        )
        alter_text = " ".join(map(str, geburts_alter))

    # ─── Reisezeit und GPT-Frage ───
    von_raw   = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw   = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage     = st.text_input("GPT-Frage zur Versicherung", placeholder="z. B. Was ist bei dasdsadsadasversichert?")
    submit    = st.form_submit_button("Tarife anzeigen")

if submit:
    # Datumsprüfung
    von = parse_datum(von_raw)
    bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben.")
        st.stop()

    # Excel laden und Tarifberechnung (Beispiel)
    excel = pd.ExcelFile("ergo.xlsx")
    # ... hier deine Logik ...

    # Beispiel-Tariftabelle
    df = pd.DataFrame([
        ["Reiserücktritt", "Einmal", "123,45 €", "234,56 €"],
        ["", "Jahres", "345,67 €", "456,78 €"],
        ["Reisekranken", "Einmal", "89,01 €", "90,12 €"],
        ["RundumSorglos", "Einmal", "34,56 €", "45,67 €"]
    ], columns=["Produktgruppe", "Tarif", "mit SB", "ohne SB"])
    st.subheader("📊 Gruppierte Tarifübersicht")
    st.table(df)

    # Word-Angebot
    daten = {
        "Kundenname": name,
        "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
        "Reisepreis": f"{preis:,.2f} €".replace('.', ','),
        "Alter": alter_text,
        "Reiseziel": zielgebiet
    }
    st.session_state["word_daten"] = daten
    st.subheader("📄 Word-Angebot")
    if st.button("📄 Word-Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Vorlage fehlt!")
        else:
            fp = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", fp, daten)
            with open(fp, "rb") as f:
                bytes_data = f.read()
            st.download_button("📥 Angebot herunterladen", bytes_data, file_name=fp)

    # GPT‑Antwort
    if frage.strip():
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fund = frage_an_gpt(frage, absätze)
        with st.expander("📄 Verwendete Textstellen aus PDF"):
            for i, a in enumerate(fund, 1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"], width=600, placeholder=" …"), unsafe_allow_html=True)
        st.success(antwort)
