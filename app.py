
import os
import streamlit as st
import pandas as pd
from datetime import datetime, date
from docx import Document

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("SmallLogoBW_png.png", width=250)
st.title("🦾 Der Ergo Chuck – Berechnung & Angebot")

# ... (alle bekannten Hilfsfunktionen bleiben unverändert: parse_datum, ermittle_zielgebiet, etc.)

def export_doc(template_path, output_path, daten):
    doc = Document(template_path)
    def ersetze(text): return text if not text else text.format(**daten)
    for p in doc.paragraphs:
        p.text = ersetze(p.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                cell.text = ersetze(cell.text)
    doc.save(output_path)
    return output_path

# Eingabe
name = st.text_input("👤 Kundenname")
ziel = st.text_input("🌍 Reiseziel (IATA-Code)")
preis = st.number_input("💶 Reisepreis (€)", min_value=0.0)
alter_text = st.text_input("👥 Alter (z. B. 45 48)")
von_raw = st.text_input("📅 Reise von (TTMM oder TT.MM.JJJJ)")
bis_raw = st.text_input("📅 Reise bis (TTMM oder TT.MM.JJJJ)")

if st.button("✅ Gruppierte Tarife anzeigen"):
    try:
        from_date = datetime.strptime(von_raw, "%d%m").replace(year=date.today().year)
        to_date = datetime.strptime(bis_raw, "%d%m").replace(year=date.today().year)

        reisetage = max(1, (to_date - from_date).days + 1)
        alter_liste = [int(a) for a in alter_text.strip().split()]
        max_alter = max(alter_liste)
        altersgruppe = "bis 40 Jahre" if max_alter <= 40 else "41–64 Jahre" if max_alter <= 64 else "ab 65 Jahre"
        personengruppe = "Einzelperson" if len(alter_liste) == 1 else "Paar" if len(alter_liste) == 2 else "Familie"
        zielgebiet = ziel.upper()

        daten = {
            "Kundenname": name,
            "Reisedatum": f"{from_date.strftime('%d.%m.%Y')} – {to_date.strftime('%d.%m.%Y')}",
            "Reisepreis": f"{preis:,.2f} €".replace(".", ","),
            "Anzahl": str(len(alter_liste)),
            "Alter": ", ".join(str(a) for a in alter_liste),
            "Reiseziel": zielgebiet,
            "Reiseruecktritt_Einmal_mit_SB": "29,00 € (RNM104)",
            "Reiseruecktritt_Einmal_ohne_SB": "39,00 € (RNX104)",
            "Reiseruecktritt_Jahres_mit_SB": "45,00 € (JTJ180)",
            "Reiseruecktritt_Jahres_ohne_SB": "57,00 € (XTJ180)",
            "Reiseruecktritt_Jahres_Sparfuchs_mit_SB": "41,00 € (JSJ90)",
            "Reiseruecktritt_Jahres_Sparfuchs_ohne_SB": "51,00 € (XSJ90)",
            "Reisekranken_Einmal_mit_SB": "4,80 € (KNM100)",
            "Reisekranken_Einmal_ohne_SB": "6,60 € (KNX100)",
            "Reisekranken_Jahres_mit_SB": "31,00 € (JKJ180)",
            "Reisekranken_Jahres_ohne_SB": "49,00 € (XKJ180)",
            "RundumSorglos_Einmal_mit_SB": "49,00 € (PNM104)",
            "RundumSorglos_Einmal_ohne_SB": "88,00 € (PNX104)",
            "RundumSorglos_Jahres_mit_SB": "53,00 € (JPJ180)",
            "RundumSorglos_Jahres_ohne_SB": "63,00 € (XPJ180)"
        }

        st.session_state["word_daten"] = daten

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
        st.success("✅ Tarifdaten erfolgreich vorbereitet.")
    except Exception as e:
        st.error(f"❌ Fehler: {e}")

# WORD-EXPORT (außerhalb try, direkt ausführbar)
if "word_daten" in st.session_state:
    st.subheader("📄 Angebot erzeugen")
    if st.button("📄 Word-Angebot erstellen"):
        if not os.path.exists("angebot.docx"):
            st.error("❌ Datei 'angebot.docx' nicht gefunden!")
        else:
            dateiname = "angebot_fertig_streamlit.docx"
            export_doc("angebot.docx", dateiname, st.session_state["word_daten"])
            with open(dateiname, "rb") as f:
                file_bytes = f.read()
            st.download_button("📥 Angebot herunterladen", file_bytes, file_name=dateiname)
