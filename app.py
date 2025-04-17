import os
import textwrap
import difflib
import streamlit as st
import pandas as pd
import fitz
from openai import OpenAI
from word_styling import export_doc
from datetime import datetime, date

# --- Setup ---
st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=200)
st.write("_Wenn Chuck Norris eine Reise plant, versichert sich das Zielland._")
st.markdown("## Der Ergo Chuck // Berechnung – Angebot – Information")

# OpenAI
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Hilfsfunktionen (parse_datum, lade_absätze_aus_pdf, frage_an_gpt, etc.) wie gehabt …

# --- Tabs ---
tab1, tab2, tab3 = st.tabs(["🧮 Berechnung", "📄 Angebot", "🤖 Information"])

with tab1:
    st.subheader("Berechnung")
    with st.form("calc_form"):
        name       = st.text_input("Kundenname")
        zielgebiet = st.radio("Zielgebiet", ["Europa","Welt"])
        preis      = st.number_input("Reisepreis (€)", min_value=0.0)
        
        # Geburtstage als Date-Input
        geb1 = st.date_input("Geburtstag 1", key="g1")
        geb2 = st.date_input("Geburtstag 2", key="g2")
        geb3 = st.date_input("Geburtstag 3", key="g3")
        geb4 = st.date_input("Geburtstag 4", key="g4")
        
        # Altersberechnung
        heute = date.today()
        alters = []
        for d in (geb1, geb2, geb3, geb4):
            if isinstance(d, date):
                age = heute.year - d.year - ((heute.month, heute.day) < (d.month, d.day))
                alters.append(age)
        if alters:
            st.write(f"👥 Berechnete Alter: {', '.join(map(str,alters))}")
            alter_text = " ".join(map(str,alters))
        else:
            alter_text = ""
        
        von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
        bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
        submit  = st.form_submit_button("Tarife anzeigen")
    
    if submit:
        # hier parse_datum, Excel–Logik, df erstellen …
        df = pd.DataFrame([...], columns=[…])  # Deine Tarif‑Daten
        st.table(df)

with tab2:
    st.subheader("Word‑Angebot")
    if "word_daten" not in st.session_state:
        st.info("Bitte erst in ➜ Berechnung die Tarife berechnen.")
    else:
        if st.button("Erstelle Word‑Angebot"):
            daten = st.session_state["word_daten"]
            fp = "angebot_out.docx"
            export_doc("angebot.docx", fp, daten)
            with open(fp,"rb") as f:
                st.download_button("Download Angebot", f.read(), file_name=fp)

with tab3:
    st.subheader("Beratung zur ERGO‑Reiseversicherung")
    frage = st.text_input("Was möchten Sie wissen?", placeholder="z. B. Corona")
    if frage:
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, relevant = frage_an_gpt(frage, absätze)
        with st.expander("Gefundene Textstellen"):
            for i,a in enumerate(relevant,1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"],600,"…"), unsafe_allow_html=True)
        st.success(antwort)
