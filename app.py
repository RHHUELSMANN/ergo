import os
import textwrap
import difflib
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from openai import OpenAI
from word_styling import export_doc
from datetime import datetime, date

# ————————————————————————————————
# OpenAI-Client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ————————————————————————————————
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

def ermittle_personengruppe(lst):
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
    return f"{val:.2f} € ({code})"

def lade_absätze(pdf_path):
    doc = fitz.open(pdf_path)
    absz = []
    for page in doc:
        txt = page.get_text()
        for para in txt.split("\n\n"):
            p = para.strip().replace("\n"," ")
            if len(p)>50:
                absz.append({"seite": page.number+1, "text": p})
    return absz

def suche_absätze(frage, absz, topk=3):
    f = frage.lower()
    scored = []
    for a in absz:
        txt = a["text"].lower()
        score = difflib.SequenceMatcher(None, f, txt).ratio()
        if any(w in txt for w in f.split()): score += 0.2
        scored.append((score,a))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [a for _,a in scored[:topk]]

def frage_an_gpt(frage, absz):
    rel = suche_absätze(frage, absz)
    ctx = "\n\n".join(f"Seite {a['seite']}:\n{a['text']}" for a in rel)
    sp = (
        "Du bist ein digitaler Versicherungsberater für Reisebüro Hülsmann. "
        "Beantworte ausschließlich Fragen zu Reiserücktritts-, Reisekranken- oder RundumSorglos-Versicherungen "
        "auf Grundlage der folgenden Auszüge. Wenn keine Info, antworte: 'Dazu liegt mir keine Information vor.'"
    )
    resp = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role":"system","content":sp},
            {"role":"user","content":f"Frage: {frage}\n\n{ctx}"}
        ],
        temperature=0.3
    )
    return resp.choices[0].message.content, rel

# ————————————————————————————————
# UI

st.set_page_config("Der Ergo Chuck", "🦾", "centered")
st.image("logo.png", width=200)
st.markdown("_Wenn Chuck Norris eine Reise plant, versichert sich das Zielland._")
st.markdown("### Der Ergo Chuck – Berechnung | Angebot | Information")

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
    al = []
    for d in (g1,g2,g3,g4):
        if isinstance(d,date):
            age = heute.year - d.year - ((heute.month,heute.day)<(d.month,d.day))
            al.append(age)
    if al:
        st.markdown(f"<small style='color:gray;'>👥 Altersberechnung: {', '.join(map(str,al))}</small>",
                    unsafe_allow_html=True)
        alter_text = " ".join(map(str,al))
    else:
        alter_text = st.text_input("Alter (z. B. 45 48)")

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    frage_gpt = st.text_input("GPT‑Frage zur Versicherung", placeholder="z. B. Corona")
    submit    = st.form_submit_button("Tarife anzeigen")

if submit:
    # Datum prüfen
    von = parse_datum(von_raw); bis = parse_datum(bis_raw)
    if not von or not bis:
        st.error("❌ Bitte gültige Datumswerte eingeben."); st.stop()

    # Excel laden
    xls = pd.ExcelFile("ergo.xlsx")
    sheets = {k: xls.parse(v) for k,v in {
        "rrv_ew_mit":"rrv-ev-mit","rrv_ew_ohne":"rrv-ev-ohne",
        "rrv_jw_mit":"rrv-jv-mit","rrv_jw_ohne":"rrv-jv-ohne",
        "kv_ew_mit":"kv-ev-mit","kv_ew_ohne":"kv-ev-ohne",
        "kv_jw_mit":"kv-jv-mit","kv_jw_ohne":"kv-jv-ohne",
        "rus_ew_mit":"rus-ev-mit","rus_ew_ohne":"rus-ev-ohne",
        "rus_jw_mit":"rus-jv-mit","rus_jw_ohne":"rus-jv-ohne"
    }.items()}

    reisetage = berechne_reisetage(von,bis)
    liste_alt = [int(a) for a in alter_text.split()] if alter_text else []
    max_alt   = max(liste_alt) if liste_alt else 0
    ag = ermittle_altersgruppe(max_alt)
    pg = ermittle_personengruppe(liste_alt)

    # DataFrame befüllen (Beispiel mit drei Zeilen; ergänze alle Tarife analog)
    df = pd.DataFrame([
        ["Reiserücktritt","Einmal",
         fmt(*first_hit(sheets["rrv_ew_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rrv_ew_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["Reiserücktritt","Jahres",
         fmt(*first_hit(sheets["rrv_jw_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["rrv_jw_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        ["Reisekranken","Einmal",
         fmt(*first_hit(sheets["kv_ew_mit"])[["Prämie","Tarifcode"]],preis),
         fmt(*first_hit(sheets["kv_ew_ohne"])[["Prämie","Tarifcode"]],preis)
        ],
        # … füge hier noch Jahres-/RundumSorglos‑Zeilen hinzu …
    ], columns=["Produktgruppe","Tarif","mit SB","ohne SB"])

    st.subheader("📊 Gruppierte Tarifübersicht")
    st.table(df)

    # Word‑Export-Daten sichern
    st.session_state["word_daten"] = {
        "Kundenname": name,
        "Reisedatum": f"{von:%d.%m.%Y} – {bis:%d.%m.%Y}",
        "Reisepreis": f"{preis:,.2f} €".replace(".",","),
        "Alter": alter_text,
        "Reiseziel": zielgebiet
    }

    # Word‑Button
    st.markdown("### Word‑Angebot")
    if st.button("📄 Word‑Dokument erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Vorlage fehlt!")
        else:
            out = "angebot_out.docx"
            export_doc("angebot.docx", out, st.session_state["word_daten"])
            with open(out,"rb") as f:
                st.download_button("Download", f.read(), file_name=out)

    # GPT‑Beratung
    st.markdown("### Beratung zur ERGO‑Reiseversicherung")
    if frage_gpt:
        absz = lade_absätze("ergo_tarife.pdf")
        antw, fund = frage_an_gpt(frage_gpt, absz)
        with st.expander("Gefundene Textstellen"):
            for i,a in enumerate(fund,1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"],600,"…"), unsafe_allow_html=True)
        st.success(antw)

