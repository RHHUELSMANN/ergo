import os
import re
import textwrap
import difflib


import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from openai import OpenAI
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def pdf_suche(pfad, suchbegriff):
    doc = fitz.open(pfad)
    treffer = []
    suchbegriffe = [suchbegriff.lower()] + SYNONYME.get(suchbegriff.lower(), [])
    
    for seite in doc:
        text = seite.get_text()
        text_lower = text.lower()
        if any(sb in text_lower for sb in suchbegriffe):
            auszug = re.sub(r"\s+", " ", text.strip())
            treffer.append((seite.number + 1, auszug[:1000]))  # max 1000 Zeichen
    return treffer

def highlight(text, wort):
    clean_text = re.sub(r"[�•⍰]", " ", text)
    pattern = re.compile(f"(?i)({re.escape(wort)})")
    return pattern.sub(r"<mark>\1</mark>", clean_text)
    
from word_styling import export_doc
from datetime import datetime, date
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import qn
from docx.oxml import OxmlElement

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("SmallLogoBW_png.png", width=250)

st.markdown(
    "<div style='font-size:18px; color:#111; font-style:italic; margin-bottom:1em;'>"
    "Wenn Chuck Norris eine Reise plant, versichert sich das Zielland."
    "</div>",
    unsafe_allow_html=True
)

st.markdown(
    "<h2 style='font-size:28px; color:#111;'>Der Ergo Chuck – Berechnung & Angebot</h2>",
    unsafe_allow_html=True
)

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
        
def ermittle_zielgebiet(code):
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    return zielgebiet

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

def fmt(p, c, preis):
    p = float(p)
    betrag = p * preis if p < 1.0 else p
    return f"{betrag:.2f}".replace(".", ",") + f" € ({c})"

def first_hit(df):
    col = [c for c in df.columns if c.strip().lower() == "reisepreis bis"]
    if col:
        return df.sort_values(col[0]).iloc[0]
    return df.iloc[0]

def bereinige_export_daten(daten):
    export_daten = {}
    for key, value in daten.items():
        if isinstance(value, str) and "€" in value and "(" in value and ")" in value:
            betrag = value.split("€")[0].strip()
            export_daten[key] = betrag + " €"
        else:
            export_daten[key] = value
    return export_daten


    def ersetze(text): return text if not text else text.format(**daten)
    for p in doc.paragraphs:
        p.text = ersetze(p.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                cell.text = ersetze(cell.text)

    for para in doc.paragraphs:
        if "{Kundenname}" in para.text:
            for run in para.runs:
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.bold = True
            p = para._element
            pPr = p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '8')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), '000000')
            pBdr.append(bottom)
            pPr.append(pBdr)
        
    kleiner_text = "Diese Übersicht wurde automatisch vom Reisebüro Hülsmann"
    for para in doc.paragraphs:
        if kleiner_text in para.text:
            for run in para.runs:
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(120, 120, 120)
                run.italic = True
    
    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    ueberschriften = [
        "Welche Versicherungen sind wichtig",
        "Abschlussfristen",
        "Reiserücktritts-Versicherung",
        "Reisekranken-Versicherung",
        "RundumSorglos-Schutz"
    ]
    for para in doc.paragraphs:
        if para.text.strip() in ueberschriften:
            for run in para.runs:
                run.bold = True

    doc.save(output_path)
    return output_path

with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    alter_text = st.text_input("Alter (z. B. 45 48)")
    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    submit = st.form_submit_button("Tarife anzeigen")

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
    if alter <= 40:
        return "bis 40 Jahre"
    elif 41 <= alter <= 64:
        return "41–64 Jahre"
    return "ab 65 Jahre"

def ermittle_personengruppe(alter_liste):
    if len(alter_liste) == 1:
        return "Einzelperson"
    elif len(alter_liste) == 2:
        return "Paar"
    return "Familie"

def berechne_reisetage(von, bis):
    return max(1, (bis - von).days + 1)

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")

st.markdown("<h2 style='font-size:28px; color:#111;'>Der Ergo Chuck – Berechnung & Angebot</h2>", unsafe_allow_html=True)

with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    alter_text = st.text_input("Alter (z. B. 45 48)")

    col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
    gb_eingaben = []
    with col1:
        gb_eingaben.append(st.text_input("", key="gb1", label_visibility="collapsed", placeholder="Geb. 1"))
    with col2:
        gb_eingaben.append(st.text_input("", key="gb2", label_visibility="collapsed", placeholder="Geb. 2"))
    with col3:
        gb_eingaben.append(st.text_input("", key="gb3", label_visibility="collapsed", placeholder="Geb. 3"))
    with col4:
        gb_eingaben.append(st.text_input("", key="gb4", label_visibility="collapsed", placeholder="Geb. 4"))

    heute = date.today()
    geburts_alter = []
    for geb in gb_eingaben:
        gebdat = parse_geburtstag(geb)
        if gebdat:
            alter = heute.year - gebdat.year - ((heute.month, heute.day) < (gebdat.month, gebdat.day))
            geburts_alter.append(alter)

    if geburts_alter:
        st.markdown(
            f"<small style='color:gray;'>👥 Berechnete Alter: {', '.join(str(a) for a in geburts_alter)}</small>",
            unsafe_allow_html=True
        )
        alter_text = " ".join(str(a) for a in geburts_alter)

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    submit = st.form_submit_button("Tarife anzeigen")

            
        df = pd.DataFrame([
            ["Reiserücktritt", "Einmal", daten["Reiseruecktritt_Einmal_mit_SB"], daten["Reiseruecktritt_Einmal_ohne_SB"]],
            ["", "Jahres", daten["Reiseruecktritt_Jahres_mit_SB"], daten["Reiseruecktritt_Jahres_ohne_SB"]],
            ["", "Sparfuchs", daten["Reiseruecktritt_Jahres_Sparfuchs_mit_SB"], daten["Reiseruecktritt_Jahres_Sparfuchs_ohne_SB"]],
            ["Reisekranken", "Einmal", daten["Reisekranken_Einmal_mit_SB"], daten["Reisekranken_Einmal_ohne_SB"]],
            ["", "Jahres", daten["Reisekranken_Jahres_mit_SB"], daten["Reisekranken_Jahres_ohne_SB"]],
            ["RundumSorglos", "Einmal", daten["RundumSorglos_Einmal_mit_SB"], daten["RundumSorglos_Einmal_ohne_SB"]],
            ["", "Jahres", daten["RundumSorglos_Jahres_mit_SB"], daten["RundumSorglos_Jahres_ohne_SB"]],
        ], columns=["Produktgruppe", "Tarif", "mit SB", "ohne SB"])
        st.subheader("📊 Gruppierte Tarifübersicht")
        st.table(df)


    except Exception as e:
        st.error(f"❌ Fehler: {e}")

if "word_daten" in st.session_state:
    st.subheader("📄 Word-Angebot")
    if st.button("📄 Word-Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.warning("⚠️ Datei 'angebot.docx' fehlt!")
        else:
            saubere_daten = bereinige_export_daten(st.session_state["word_daten"])
            file_path = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", file_path, saubere_daten)
            with open(file_path, "rb") as f:
                file_bytes = f.read()
            st.download_button("📥 Angebot herunterladen", file_bytes, file_name=file_path)


# PDF laden und Absätze vorbereiten
@st.cache_data
def lade_absätze_aus_pdf(pfad):
    doc = fitz.open(pfad)
    absatz_liste = []

    for seite in doc:
        text = seite.get_text()
        nummer = seite.number + 1
        absätze = [a.strip() for a in text.split("\n\n") if len(a.strip()) > 50]

        for absatz in absätze:
            absatz_liste.append({
                "seite": nummer,
                "text": absatz
            })

    return absatz_liste

def suche_passende_absätze(frage, absätze, anzahl=3):
    frage = frage.lower()
    scored = []

    for absatz in absätze:
        text = absatz["text"].lower()
        score = difflib.SequenceMatcher(None, frage, text).ratio()
        if any(wort in text for wort in frage.split()):
            score += 0.2  # Bonus für direkte Worttreffer
        scored.append((score, absatz))

    scored.sort(reverse=True, key=lambda x: x[0])
    beste = [a for _, a in scored[:anzahl]]
    return beste

def frage_an_gpt(frage, absatz_liste):
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    relevante_abschnitte = suche_passende_absätze(frage, absatz_liste)

    kontext = "\n\n".join([f"Seite {a['seite']}:\n{a['text']}" for a in relevante_abschnitte])

    system_prompt = (
        "Du bist ein digitaler Versicherungsberater für Reisebüro Hülsmann. "
        "Beantworte ausschließlich Fragen zu Reiserücktritts-, Reisekranken- oder RundumSorglos-Versicherungen "
        "auf Grundlage der folgenden PDF-Auszüge.\n\n"
        "Berücksichtige bei der Interpretation auch Begriffe mit ähnlicher Bedeutung. "
        "Zum Beispiel:\n"
        "- Selbstbeteiligung ≈ Selbstbehalt ≈ SB ≈ Eigenanteil\n"
        "- Reiserücktritt ≈ Rücktritt ≈ Stornierung\n"
        "- Krankheit ≈ Corona ≈ COVID ≈ Quarantäne\n\n"
        "Wenn du keine ausreichende Information findest, sage bitte klar: "
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

    return response.choices[0].message.content, relevante_abschnitte

# Streamlit UI
st.subheader("🤖 Beratung zur ERGO-Reiseversicherung")

frage = st.text_input("Welche Frage haben Sie zur Versicherung?", placeholder="z. B. Was ist bei Corona versichert?")
if frage.strip():
    with st.spinner("Suche passende Textstellen und frage GPT …"):
        absätze = lade_absätze_aus_pdf("ergo_tarife.pdf")
        antwort, abschnitte = frage_an_gpt(frage, absätze)

        with st.expander("📄 Verwendete PDF-Auszüge anzeigen"):
            for i, a in enumerate(abschnitte, 1):
                st.markdown(f"**{i}. Seite {a['seite']}**")
                st.markdown(textwrap.shorten(a["text"], width=600, placeholder=" …"), unsafe_allow_html=True)

        st.success(antwort)
