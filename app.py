
import streamlit as st
import pandas as pd
from datetime import datetime, date
from docx import Document

st.set_page_config(page_title="Der Ergo Chuck", page_icon="🦾", layout="centered")
st.image("SmallLogoBW_png.png", width=250)
st.title("🦾 Der Ergo Chuck – Berechnung & Angebot")

EUROPA_CODES = ["PMI", "FRA", "BER", "VIE", "ZRH", "LIS", "CDG", "AMS", "BCN", "ROM"]
WELT_CODES = ["PUJ", "BKK", "JFK", "LAX", "DXB", "CUN", "MEX", "CPT", "SIN", "HND"]

def parse_datum(text):
    try:
        if "." in text:
            return datetime.strptime(text.strip(), "%d.%m.%Y").date()
        elif len(text.strip()) == 4:
            ttmm = text.strip()
            return date.today().replace(day=int(ttmm[:2]), month=int(ttmm[2:]))
        elif len(text.strip()) == 8:
            return datetime.strptime(text.strip(), "%d%m%Y").date()
    except:
        return None

def ermittle_zielgebiet(code):
    code = code.upper()
    if code in EUROPA_CODES: return "Europa"
    if code in WELT_CODES: return "Welt"
    return "Unbekannt"

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
        von = parse_datum(von_raw)
        bis = parse_datum(bis_raw)
        if not von or not bis:
            st.error("❌ Bitte gültige Datumswerte eingeben.")
            st.stop()

        alter_liste = [int(a) for a in alter_text.strip().split()]
        max_alter = max(alter_liste)
        altersgruppe = ermittle_altersgruppe(max_alter)
        personengruppe = ermittle_personengruppe(alter_liste)
        zielgebiet = ermittle_zielgebiet(ziel)
        reisetage = berechne_reisetage(von, bis)

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

        def tarif(d, *args): return fmt(*args, preis) if not d.empty else "–"
        def first(d, cond, keys): f = d[cond]; return tarif(f, *first_hit(f)[keys]) if not f.empty else "–"

        def kv_einmal(name):
            d = t[name]
            f = d[
                (d["Zielgebiet"].str.lower().str.strip() == zielgebiet.lower()) &
                (d["Altersgruppe"].str.strip() == altersgruppe) &
                (d["Personengruppe"].str.lower().str.strip() == personengruppe.lower())
            ]
            if not f.empty:
                tagespreis = float(f.iloc[0]["Tagesprämie"])
                return fmt(reisetage * tagespreis, f.iloc[0]["Tarifcode"], preis)
            return "–"

        daten = {
            "Kundenname": name,
            "Reisedatum": f"{von.strftime('%d.%m.%Y')} – {bis.strftime('%d.%m.%Y')}",
            "Reisepreis": f"{preis:,.2f} €".replace(".", ","),
            "Anzahl": str(len(alter_liste)),
            "Alter": ", ".join(str(a) for a in alter_liste),
            "Reiseziel": ziel.upper(),
            "Reiseruecktritt_Einmal_mit_SB": first(t["rrv_ew_mit"], t["rrv_ew_mit"]["Reisepreis bis"] >= preis, ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Einmal_ohne_SB": first(t["rrv_ew_ohne"], (t["rrv_ew_ohne"]["Reisepreis bis"] >= preis) & (t["rrv_ew_ohne"]["Altersgruppe"].str.strip() == altersgruppe), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_mit_SB": first(t["rrv_jw_mit"], (t["rrv_jw_mit"]["Reisepreis bis"] >= preis) & (t["rrv_jw_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_ohne_SB": first(t["rrv_jw_ohne"], (t["rrv_jw_ohne"]["Reisepreis bis"] >= preis) & (t["rrv_jw_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_Sparfuchs_mit_SB": first(t["rrv_jw_spf_mit"], (t["rrv_jw_spf_mit"]["Reisepreis bis"] >= preis) & (t["rrv_jw_spf_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_spf_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reiseruecktritt_Jahres_Sparfuchs_ohne_SB": first(t["rrv_jw_spf_ohne"], (t["rrv_jw_spf_ohne"]["Reisepreis bis"] >= preis) & (t["rrv_jw_spf_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rrv_jw_spf_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reisekranken_Einmal_mit_SB": kv_einmal("kv_ew_mit"),
            "Reisekranken_Einmal_ohne_SB": kv_einmal("kv_ew_ohne"),
            "Reisekranken_Jahres_mit_SB": first(t["kv_jw_mit"], (t["kv_jw_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["kv_jw_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "Reisekranken_Jahres_ohne_SB": first(t["kv_jw_ohne"], (t["kv_jw_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["kv_jw_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Einmal_mit_SB": first(t["rus_ew_mit"], (t["rus_ew_mit"]["Reisepreis bis"] >= preis) & (t["rus_ew_mit"]["Zielgebiet"].str.lower().str.strip() == zielgebiet.lower()), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Einmal_ohne_SB": first(t["rus_ew_ohne"], (t["rus_ew_ohne"]["Reisepreis bis"] >= preis) & (t["rus_ew_ohne"]["Zielgebiet"].str.lower().str.strip() == zielgebiet.lower()) & (t["rus_ew_ohne"]["Altersgruppe"].str.strip() == altersgruppe), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Jahres_mit_SB": first(t["rus_jw_mit"], (t["rus_jw_mit"]["Reisepreis bis"] >= preis) & (t["rus_jw_mit"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rus_jw_mit"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
            "RundumSorglos_Jahres_ohne_SB": first(t["rus_jw_ohne"], (t["rus_jw_ohne"]["Reisepreis bis"] >= preis) & (t["rus_jw_ohne"]["Altersgruppe"].str.strip() == altersgruppe) & (t["rus_jw_ohne"]["Personengruppe"].str.lower().str.strip() == personengruppe.lower()), ["Prämie", "Tarifcode"]),
        }

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

        st.session_state["word_daten"] = daten

        if st.button("📄 Word-Angebot erstellen"):
            dateiname = "angebot_fertig_streamlit.docx"
            if not os.path.exists("angebot.docx"):
                st.error("❌ Datei 'angebot.docx' nicht gefunden!")
                st.stop()
            export_doc("angebot.docx", dateiname, st.session_state["word_daten"])
            st.write("📤 Starte Word-Export...")
            with open(dateiname, "rb") as f:
                st.download_button("📥 Angebot herunterladen", f, file_name=dateiname)
            st.success("✅ Word-Datei wurde erfolgreich erstellt!")

    except Exception as e:
        st.error(f"❌ Fehler: {e}")
