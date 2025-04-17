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

# ————————————————————————————————
# OpenAI-Client initialisieren (API-Key in deinen Streamlit-Secrets hinterlegen)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ————————————————————————————————
# Hilfsfunktionen

def parse_datum(text: str):
    """Parst TT.MM.JJJJ, TTMMJJJJ oder TTMM (aktuelles Jahr)."""
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
    """Parst Geburtstag ddmmyy oder ddmmyyyy (mit/ohne Punkte)."""
    t = text.strip().replace(".", "").replace(" ", "")
    try:
        if len(t) == 6:
            return datetime.strptime(t, "%d%m%y").date()
        if len(t) == 8:
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
    cols = [c for c in df.columns if "reisepreis bis" in c.lower()]
    if cols:
        return df.sort_values(cols[0]).iloc[0]
    return df.iloc[0]

def fmt(p, code, preis):
    p = float(p)
    val = p * preis if p < 1 else p
    return f"{val:.2f} € ({code})"

def lade_absätze_aus_pdf(pfad: str):
    absz = []
    doc = fitz.open(pfad)
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
    scored = []
    for a in absätze:
        txt = a["text"].lower()
        score = difflib.SequenceMatcher(None, f, txt).ratio()
        if any(w in txt for w in f.split()):
            score += 0.2
        scored.append((score, a))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [a for (_, a) in scored[:topk]]

def frage_an_gpt(frage: str, absätze: list):
    relevant = suche_passende_absätze(frage, absätze)
    kontext = "\n\n".join(f"Seite {a['seite']}:\n{a['text']}" for a in relevant)
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

# ————————————————————————————————
# Streamlit UI

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=200)
st.markdown("_Wenn Chuck Norris eine Reise plant, versichert sich das Zielland._")
st.markdown("## Der Ergo Chuck – Berechnung | Angebot | Information")

# Formular
with st.form("main_form"):
    st.subheader("Berechnung & Tarifübersicht")
    name       = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa","Welt"])
    preis      = st.number_input("Reisepreis (€)", min_value=0.0)

    # Geburtstage
    g1 = st.date_input("Geburtsdatum 1", key="g1")
    g2 = st.date_input("Geburtsdatum 2", key="g2")
    g3 = st.date_input("Geburtsdatum 3", key="g3")
    g4 = st.date_input("Geburtsdatum 4", key="g4")

    heute = date.today()
    alters = []
    for d in (g1,g2,g3,g4):
        if isinstance(d, date):
            age = heute.year - d.year - ((heute.month,heute.day) < (d.month,d.day))
            alters.append(age)
    if alters:
        st.markdown(f"<small style='color:gray;'>👥 Altersberechnung: {', '.join(map(str,alters))}</small>",
                    unsafe_allow_html=True)
        alter_text = " ".join(map(str,alters))
    else:
        alter_text = st.text_input("Alter (z. B. 45 48)")

    von_raw    = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw    = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage_gpt  = st.text_input("GPT‑Frage zur Versicherung", placeholder="z. B. Corona")
    submit     = st.form_submit_button("Tarife anzeigen")

if submit:
    # Datum validieren
    von = parse_datum(von_raw); bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben.")
        st.stop()

    # Excel laden
    xls = pd.ExcelFile("ergo.xlsx")
    sheet_map = {
        "rrv_ew_mit":"rrv-ev-mit","rrv_ew_ohne":"rrv-ev-ohne",
        "rrv_jw_mit":"rrv-jv-mit","rrv_jw_ohne":"rrv-jv-ohne",
        "kv_ew_mit":"kv-ev-mit","kv_ew_ohne":"kv-ev-ohne",
        "kv_jw_mit":"kv-jv-mit","kv_jw_ohne":"kv-jv-ohne",
        "rus_ew_mit":"rus-ev-mit","rus_ew_ohne":"rus-ev-ohne",
        "rus_jw_mit":"rus-jv-mit","rus_jw_ohne":"rus-jv-ohne"
    }
    sheets = {k: xls.parse(v) for k,v in sheet_map.items()}

    reisetage = berechne_reisetage(von,bis)
    liste_alt = [int(a) for a in alter_text.split()] if alter_text else []
    max_alt   = max(liste_alt) if liste_alt else 0
    ag = ermittle_altersgruppe(max_alt)
    pg = ermittle_personengruppe(liste_alt)

    # Tariftabelle aufbauen
    df = pd.DataFrame([
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
    ], columns=["Produktgruppe","Tarif","mit SB","ohne SB"])

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
            fp = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", fp, daten)
            with open(fp, "rb") as f:
                btn = f.read()
            st.download_button("📥 Angebot herunterladen", btn, file_name=fp)

    # GPT‑Block unterhalb der Tariftabelle
    if frage_gpt:
        absz = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fund = frage_an_gpt(frage_gpt, absz)
        with st.expander("📄 Verwendete Textstellen aus PDF"):
            for i,a in enumerate(fund,1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"], width=600, placeholder=" …"), unsafe_allow_html=True)
        st.success(antwort)
