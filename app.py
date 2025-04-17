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

# ——————————————————————————————————————————————
# OpenAI-Client (API-Key in Streamlit-Secrets hinterlegen)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ——————————————————————————————————————————————
# Hilfsfunktionen

def parse_datum(text: str):
    t = text.strip()
    try:
        if "." in t and len(t.split(".")) == 3:
            return datetime.strptime(t, "%d.%m.%Y").date()
        if len(t) == 8 and t.isdigit():
            return datetime.strptime(t, "%d%m%Y").date()
        if len(t) == 4 and t.isdigit():
            dd, mm = int(t[:2]), int(t[2:])
            return date(date.today().year, mm, dd)
    except:
        return None
    return None

def parse_geburtstag(text: str):
    t = text.strip().replace(".", "").replace(" ", "")
    try:
        if len(t) == 6 and t.isdigit():
            return datetime.strptime(t, "%d%m%y").date()
        if len(t) == 8 and t.isdigit():
            return datetime.strptime(t, "%d%m%Y").date()
    except:
        return None
    return None

def berechne_reisetage(von: date, bis: date):
    return max(1, (bis - von).days + 1)

def ermittle_altersgruppe(a: int):
    if a <= 40: return "bis 40 Jahre"
    if a <= 64: return "41–64 Jahre"
    return "ab 65 Jahre"

def ermittle_personengruppe(lst: list):
    if len(lst) == 1: return "Einzelperson"
    if len(lst) == 2: return "Paar"
    return "Familie"

def first_hit(df: pd.DataFrame):
    cols = [c for c in df.columns if "Reisepreis bis".lower() in c.lower()]
    if cols:
        return df.sort_values(cols[0]).iloc[0]
    return df.iloc[0]

def fmt(p, code, preis):
    p = float(p)
    raus = (p * preis) if p < 1 else p
    return f"{raus:.2f} € ({code})"

# PDF‑Suche / GPT‑Hilfe
def lade_absätze_aus_pdf(pfad: str):
    doc = fitz.open(pfad)
    absz = []
    for page in doc:
        txt = page.get_text()
        seite = page.number + 1
        for para in txt.split("\n\n"):
            p = para.strip().replace("\n", " ")
            if len(p) > 50:
                absz.append({"seite": seite, "text": p})
    return absz

def suche_passende_absätze(frage: str, absätze: list, topk: int = 3):
    f = frage.lower()
    rated = []
    for a in absätze:
        score = difflib.SequenceMatcher(None, f, a["text"].lower()).ratio()
        rated.append((score, a))
    rated.sort(key=lambda x: x[0], reverse=True)
    return [a for (_,a) in rated[:topk]]

def frage_an_gpt(frage: str, absätze: list):
    top = suche_passende_absätze(frage, absätze)
    kontext = "\n\n".join(f"Seite {a['seite']}:\n{a['text']}" for a in top)
    system_prompt = (
        "Du bist ein digitaler Versicherungsberater für Reisebüro Hülsmann. "
        "Beantworte ausschließlich Fragen zu Reiserücktritts-, Reisekranken- oder RundumSorglos-Versicherungen "
        "auf Grundlage der PDF-Auszüge. Wenn keine Info, sage: 'Dazu liegt mir keine Information vor.'"
    )
    resp = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user",   "content": f"Frage: {frage}\n\n{kontext}"}
        ],
        temperature=0.3
    )
    return resp.choices[0].message.content, top

# ——————————————————————————————————————————————
# Streamlit‑App

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=200)
st.markdown("_Wenn Chuck Norris eine Reise plant, versichert sich das Zielland._")
st.markdown("## Der Ergo Chuck – Berechnung | Angebot | Information")

# — Formular —
with st.form("main_form"):
    st.subheader("Berechnung & Tarifübersicht")
    name       = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa","Welt"])
    preis      = st.number_input("Reisepreis (€)", min_value=0.0)

    # → Geburtstagsfelder nebeneinander
    col1,col2,col3,col4 = st.columns(4)
    geb_eingaben = []
    geb_eingaben.append(col1.text_input("", key="gb1", placeholder="Geb. 1", label_visibility="collapsed"))
    geb_eingaben.append(col2.text_input("", key="gb2", placeholder="Geb. 2", label_visibility="collapsed"))
    geb_eingaben.append(col3.text_input("", key="gb3", placeholder="Geb. 3", label_visibility="collapsed"))
    geb_eingaben.append(col4.text_input("", key="gb4", placeholder="Geb. 4", label_visibility="collapsed"))

    # → Alter berechnen
    heute = date.today()
    alters = []
    for t in geb_eingaben:
        d = parse_geburtstag(t)
        if d:
            age = heute.year - d.year - ((heute.month,heute.day) < (d.month,d.day))
            alters.append(age)
    if alters:
        alter_text = " ".join(map(str,alters))
        st.markdown(
            f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(map(str,alters))}</small>",
            unsafe_allow_html=True
        )
    else:
        alter_text = st.text_input("Alter (z. B. 45 48)")

    von_raw   = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw   = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage_gpt  = st.text_input("GPT‑Frage zur Versicherung", placeholder="z. B. Corona")
    submit    = st.form_submit_button("Tarife anzeigen")

if submit:
    von = parse_datum(von_raw); bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben.")
        st.stop()

    # Excel & Tariftabelle
    xls = pd.ExcelFile("ergo.xlsx")
    mapping = {
        "rrv_ew_mit":"rrv-ev-mit","rrv_ew_ohne":"rrv-ev-ohne",
        "rrv_jw_mit":"rrv-jv-mit","rrv_jw_ohne":"rrv-jv-ohne",
        "kv_ew_mit":"kv-ev-mit","kv_ew_ohne":"kv-ev-ohne",
        "kv_jw_mit":"kv-jv-mit","kv_jw_ohne":"kv-jv-ohne",
        "rus_ew_mit":"rus-ev-mit","rus_ew_ohne":"rus-ev-ohne",
        "rus_jw_mit":"rus-jv-mit","rus_jw_ohne":"rus-jv-ohne"
    }
    sheets = {k: xls.parse(v) for k,v in mapping.items()}

    tage   = berechne_reisetage(von,bis)
    liste_alt = [int(a) for a in alter_text.split()] if alter_text else []
    max_alt   = max(liste_alt) if liste_alt else 0
    ag = ermittle_altersgruppe(max_alt)
    pg = ermittle_personengruppe(liste_alt)

    rows = [
        ["Reiserücktritt","Einmal",
         fmt(*first_hit(sheets["rrv_ew_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rrv_ew_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["","Jahres",
         fmt(*first_hit(sheets["rrv_jw_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rrv_jw_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["","Sparfuchs",
         fmt(*first_hit(sheets["rrv_jw_spf_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rrv_jw_spf_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["Reisekranken","Einmal",
         fmt(*first_hit(sheets["kv_ew_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["kv_ew_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["","Jahres",
         fmt(*first_hit(sheets["kv_jw_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["kv_jw_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["RundumSorglos","Einmal",
         fmt(*first_hit(sheets["rus_ew_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rus_ew_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["","Jahres",
         fmt(*first_hit(sheets["rus_jw_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rus_jw_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
    ]
    df = pd.DataFrame(rows, columns=["Produktgruppe","Tarif","mit SB","ohne SB"])

    st.subheader("📊 Gruppierte Tarifübersicht")
    st.table(df)

    # Word‑Export
    daten = {
        "Kundenname": name,
        "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
        "Reisepreis": f"{preis:,.2f} €".replace(".",","),
        "Alter": alter_text,
        "Reiseziel": zielgebiet
    }
    st.session_state["word_daten"] = daten
    st.subheader("📄 Word‑Angebot")
    if st.button("📄 Word‑Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Vorlage 'angebot.docx' fehlt!")
        else:
            fp = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", fp, daten)
            with open(fp,"rb") as f: data = f.read()
            st.download_button("📥 Angebot herunterladen", data, file_name=fp)

    # GPT‑Block
    if frage_gpt:
        absz = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fund = frage_an_gpt(frage_gpt, absz)
        with st.expander("📄 Gefundene PDF‑Textstellen"):
            for i,a in enumerate(fund,1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"],width=600,placeholder=" …"),
                            unsafe_allow_html=True)
        st.success(antwort)

