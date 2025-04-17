
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
        "Beantworte ausschließlich Fragen zu Reiseversicherungen auf Grundlage der "
        "folgenden PDF-Auszüge. Wenn du es nicht beantworten kannst, sage: "
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

# Streamlit UI aufsetzen
st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=250)
st.markdown("<h2 style='font-size:28px;'>Der Ergo Chuck // Berechnung - Angebot - Information</h2>", unsafe_allow_html=True)

# Eingabeformular
with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    # manuelles Alter als Fallback
    alter_text = st.text_input("Alter (z. B. 45 48)", key="alter_manual")

    # Geburtstagsfelder 4-er Spalte
    cols = st.columns(4)
    geb_eingaben = []
    for idx, col in enumerate(cols, start=1):
        with col:
            geb_eingaben.append(st.text_input("", key=f"gb{idx}", label_visibility="collapsed", placeholder=f"Geb. {idx}"))

    # Automatische Altersberechnung aus Geburtstagen
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

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage = st.text_input("GPT-Frage zur Versicherung", placeholder="z. B. Was ist bei Corona versichert?")
    submit = st.form_submit_button("Tarife anzeigen")

if submit:
    # Datumsprüfung
    von = parse_datum(von_raw)
    bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben.")
        st.stop()

    # Excel laden und Tariflogik
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

    def fmt(p, c, preis):
        val = p * preis if p < 1.0 else p
        return f"{val:.2f} € ({c})"

    def first_hit(df):
        col = [c for c in df.columns if "reisepreis bis" in c.lower()]
        return df.sort_values(col[0]).iloc[0] if col else df.iloc[0]

    def tarif(df, cond, keys):
        f = df[cond]
        if f.empty: return "–"
        row = first_hit(f)
        return fmt(row[keys[0]], row[keys[1]], preis)

    def kv_einmal(name):
        df_k = t[name]
        f = df_k[
            (df_k["Zielgebiet"].str.strip().str.lower()==zielgebiet.lower()) &
            (df_k["Altersgruppe"].str.strip()==alter_text.split()[0]) 
        ]
        if not f.empty:
            tages = float(f.iloc[0]["Tagesprämie"])
            return fmt(tages * (bis-von).days+1, f.iloc[0]["Tarifcode"], preis)
        return "–"

    # Zusammenstellen der Daten
    daten = {
        "Kundenname": name,
        "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
        "Reisepreis": f"{preis:,.2f} €".replace('.', ','),
        "Alter": alter_text,
        "Reiseziel": zielgebiet,
        "Reiserücktritt_Einmal_mit_SB": tarif(t["rrv_ew_mit"], t["rrv_ew_mit"]["Reisepreis bis"]>=preis, ["Prämie","Tarifcode"]),
        "Reiserücktritt_Jahres_mit_SB": tarif(t["rrv_jw_mit"], t["rrv_jw_mit"]["Reisepreis bis"]>=preis, ["Prämie","Tarifcode"]),
        "Reisekranken_Einmal_mit_SB": kv_einmal("kv_ew_mit"),
        "RundumSorglos_Einmal_mit_SB": tarif(t["rus_ew_mit"], t["rus_ew_mit"]["Reisepreis bis"]>=preis, ["Prämie","Tarifcode"])
    }

    # Tabellenanzeige
    df_view = pd.DataFrame([
        ["Reiserücktritt", "Einmal", daten["Reiserücktritt_Einmal_mit_SB"], ""],
        ["", "Jahres", daten["Reiserücktritt_Jahres_mit_SB"], ""],
        ["Reisekranken", "Einmal", daten["Reisekranken_Einmal_mit_SB"], ""],
        ["RundumSorglos", "Einmal", daten["RundumSorglos_Einmal_mit_SB"], ""]
    ], columns=["Produktgruppe","Tarif","mit SB","ohne SB"])
    st.subheader("📊 Gruppierte Tarifübersicht")
    st.table(df_view)

    # Word Export
    st.session_state["word_daten"] = daten
    st.subheader("📄 Word-Angebot")
    if st.button("📄 Word-Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Vorlage 'angebot.docx' fehlt!")
        else:
            fp = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", fp, daten)
            with open(fp, "rb") as f: btn = f.read()
            st.download_button("📥 Angebot herunterladen", btn, file_name=fp)

    # GPT-Antwort
    if frage.strip():
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fund = frage_an_gpt(frage, absätze)
        with st.expander("📄 Verwendete Textstellen aus PDF"):
            for i, a in enumerate(fund,1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"], width=600, placeholder=" …"), unsafe_allow_html=True)
        st.success(antwort)
