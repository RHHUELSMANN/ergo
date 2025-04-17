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

# ———————————————————————
# OpenAI-Client (Key in Secrets)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ———————————————————————
# Hilfsfunktionen

def parse_datum(text: str):
    """Parst TT.MM.JJJJ, TTMMJJJJ oder TTMM (Jahr heute)."""
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
    """Parst Geburtstag ddmmyy oder ddmmyyyy (mit oder ohne Punkte)."""
    t = text.strip().replace(".", "").replace(" ", "")
    try:
        if len(t) == 6:
            return datetime.strptime(t, "%d%m%y").date()
        if len(t) == 8:
            return datetime.strptime(t, "%d%m%Y").date()
    except:
        return None
    return None

def lade_absätze_aus_pdf(pfad: str):
    """Liest PDF und gibt Liste von Absätzen mit Seitenzahlen zurück."""
    absätze = []
    doc = fitz.open(pfad)
    for page in doc:
        text = page.get_text()
        seite = page.number + 1
        for para in text.split("\n\n"):
            p = para.strip().replace("\n", " ")
            if len(p) > 50:
                absätze.append({"seite": seite, "text": p})
    return absätze

def suche_passende_absätze(frage: str, absätze: list, topk: int = 3):
    """Scored Absätze per SequenceMatcher, gibt topk."""
    f = frage.lower()
    scores = []
    for a in absätze:
        txt = a["text"].lower()
        score = difflib.SequenceMatcher(None, f, txt).ratio()
        # Boost, wenn ein Stichwort drin ist
        if any(w in txt for w in f.split()):
            score += 0.2
        scores.append((score, a))
    scores.sort(key=lambda x: x[0], reverse=True)
    return [a for (_, a) in scores[:topk]]

def frage_an_gpt(frage: str, absätze: list):
    """Sendet System+User-Prompt an GPT und liefert Antwort + Absätze."""
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
            {"role": "user",   "content": f"Frage: {frage}\n\nPDF-Auszüge:\n{kontext}"}
        ],
        temperature=0.3
    )
    return resp.choices[0].message.content, relevant

def ermittle_altersgruppe(alter: int):
    if alter <= 40: return "bis 40 Jahre"
    if alter <= 64: return "41–64 Jahre"
    return "ab 65 Jahre"

def ermittle_personengruppe(alter_liste: list):
    if len(alter_liste) == 1: return "Einzelperson"
    if len(alter_liste) == 2: return "Paar"
    return "Familie"

def berechne_reisetage(von: date, bis: date):
    return max(1, (bis - von).days + 1)

def first_hit(df: pd.DataFrame):
    col = [c for c in df.columns if "reisepreis bis" in c.lower()]
    if col:
        return df.sort_values(col[0]).iloc[0]
    return df.iloc[0]

def fmt(p, c, preis):
    p = float(p)
    wert = p * preis if p < 1.0 else p
    return f"{wert:.2f} € ({c})"

# ———————————————————————
# Streamlit-Oberfläche

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("logo.png", width=200)
st.markdown("_Wenn Chuck Norris eine Reise plant, versichert sich das Zielland._")
st.markdown("## Der Ergo Chuck // Berechnung – Angebot – Information")

tab1, tab2, tab3 = st.tabs(["🧮 Berechnung", "📄 Angebot", "🤖 Information"])

# ─────────── Tab 1: Berechnung ───────────
with tab1:
    st.subheader("Berechnung & Tarif‑Anzeige")
    with st.form("form_calc"):
        name       = st.text_input("Kundenname")
        zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"])
        preis      = st.number_input("Reisepreis (€)", min_value=0.0)

        # Geburtstagsfelder als Date-Input
        geb1 = st.date_input("Geburtsdatum 1", key="g1")
        geb2 = st.date_input("Geburtsdatum 2", key="g2")
        geb3 = st.date_input("Geburtsdatum 3", key="g3")
        geb4 = st.date_input("Geburtsdatum 4", key="g4")

        # Altersberechnung
        heute = date.today()
        alters = []
        for d in (geb1, geb2, geb3, geb4):
            if isinstance(d, date):
                a = heute.year - d.year - ((heute.month, heute.day) < (d.month, d.day))
                alters.append(a)
        if alters:
            st.markdown(f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(map(str,alters))}</small>",
                        unsafe_allow_html=True)
            alter_text = " ".join(map(str,alters))
        else:
            alter_text = st.text_input("Alter (z. B. 45 48)", key="alter_manual")

        von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
        bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")

        submit = st.form_submit_button("Tarife anzeigen")

    if submit:
        von = parse_datum(von_raw)
        bis = parse_datum(bis_raw)
        if not von or not bis:
            st.error("❌ Bitte gültige Datumswerte eingeben.")
            st.stop()

        # Excel laden & sheets parsen
        excel = pd.ExcelFile("ergo.xlsx")
        blatt_map = {
            "rrv_ew_mit": "rrv-ev-mit",    "rrv_ew_ohne": "rrv-ev-ohne",
            "rrv_jw_mit": "rrv-jv-mit",    "rrv_jw_ohne": "rrv-jv-ohne",
            "kv_ew_mit":  "kv-ev-mit",     "kv_ew_ohne":  "kv-ev-ohne",
            "kv_jw_mit":  "kv-jv-mit",     "kv_jw_ohne":  "kv-jv-ohne",
            "rus_ew_mit": "rus-ev-mit",    "rus_ew_ohne":"rus-ev-ohne",
            "rus_jw_mit": "rus-jv-mit",    "rus_jw_ohne":"rus-jv-ohne"
        }
        tables = {k: excel.parse(v) for k,v in blatt_map.items()}

        # Filter & Format helper
        def tages(dframe, cond, keys):
            f = dframe[cond]
            if f.empty: return "–"
            base = float(f.iloc[0]["Tagesprämie"])
            code = f.iloc[0]["Tarifcode"]
            return fmt(round(reisetage * base,2), code, preis)

        # Reise-Tage & Gruppen
        reisetage = berechne_reisetage(von, bis)
        liste_alter = [int(a) for a in alter_text.split()] if alter_text else []
        max_alt = max(liste_alter) if liste_alter else 0
        ag = ermittle_altersgruppe(max_alt)
        pg = ermittle_personengruppe(liste_alter)

        # Beispiel: Baue DataFrame
        df = pd.DataFrame([
            ["Reiserücktritt", "Einmal",
             fmt(*first_hit(tables["rrv_ew_mit"])[["Prämie","Tarifcode"]], preis),
             fmt(*first_hit(tables["rrv_ew_ohne"])[["Prämie","Tarifcode"]], preis)
            ],
            ["Reiserücktritt","Jahres",
             fmt(*first_hit(tables["rrv_jw_mit"])[["Prämie","Tarifcode"]], preis),
             fmt(*first_hit(tables["rrv_jw_ohne"])[["Prämie","Tarifcode"]], preis)
            ],
            ["Reisekranken","Einmal",
             tages(tables["kv_ew_mit"],
                   (tables["kv_ew_mit"]["Altersgruppe"]==ag)&
                   (tables["kv_ew_mit"]["Personengruppe"]==pg),
                   ["Prämie","Tarifcode"]),
             tages(tables["kv_ew_ohne"],
                   (tables["kv_ew_ohne"]["Altersgruppe"]==ag)&
                   (tables["kv_ew_ohne"]["Personengruppe"]==pg),
                   ["Prämie","Tarifcode"])
            ]
        ], columns=["Produktgruppe","Tarif","mit SB","ohne SB"])

        st.subheader("📊 Gruppierte Tarifübersicht")
        st.table(df)

        # Word-Daten merken
        st.session_state["word_daten"] = {
            "Kundenname": name,
            "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
            "Reisepreis": f"{preis:,.2f} €".replace(".",","),
            "Alter": alter_text,
            "Reiseziel": zielgebiet
        }

# ─────────── Tab 2: Angebot ───────────
with tab2:
    st.subheader("Word‑Angebot")
    if "word_daten" not in st.session_state:
        st.info("Bitte erst im Tab „Berechnung“ Tarife anzeigen.")
    else:
        if st.button("📄 Word‑Angebot erstellen"):
            daten = st.session_state["word_daten"]
            if not os.path.exists("angebot.docx"):
                st.warning("⚠️ Vorlage 'angebot.docx' fehlt!")
            else:
                out = "angebot_out.docx"
                export_doc("angebot.docx", out, daten)
                with open(out,"rb") as f:
                    st.download_button("📥 Angebot herunterladen", f.read(), file_name=out)

# ─────────── Tab 3: Information ───────────
with tab3:
    st.subheader("Beratung zur ERGO‑Reiseversicherung")
    frage = st.text_input("Welche Frage haben Sie?", placeholder="z. B. Corona")
    if frage:
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, fund = frage_an_gpt(frage, absätze)
        with st.expander("📄 Gefundene Textstellen"):
            for i,a in enumerate(fund,1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"],600,"…"), unsafe_allow_html=True)
        st.success(antwort)
