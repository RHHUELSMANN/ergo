
import streamlit as st
import tempfile
from word_styling import export_doc

def euro_wert(wert):
    return f"{float(wert):,.2f}".replace(".", "X").replace(",", ".").replace("X", ",") + " €"

def export_word_dokument(name, reisepreis, reise_von, reise_bis, zielgebiet, alter_liste, max_alter,
                         rrv_ev_mit, rrv_ev_ohne,
                         rrv_jv_mit, rrv_jv_ohne,
                         spf_mit, spf_ohne,
                         kranken_mit, kranken_ohne,
                         kv_jv_mit, kv_jv_ohne,
                         rus_ev_mit, rus_ev_ohne,
                         rus_jv_mit, rus_jv_ohne):

    daten = {
        "Kundenname": name,
        "Reisedatum": f"{reise_von} – {reise_bis}",
        "Reisepreis": euro_wert(reisepreis),
        "Anzahl": str(len(alter_liste)),
        "Alter": str(max_alter),
        "Reiseziel": zielgebiet,

        "Reiseruecktritt_Einmal_mit_SB": euro_wert(rrv_ev_mit),
        "Reiseruecktritt_Einmal_ohne_SB": euro_wert(rrv_ev_ohne),
        "Reiseruecktritt_Jahres_mit_SB": euro_wert(rrv_jv_mit),
        "Reiseruecktritt_Jahres_ohne_SB": euro_wert(rrv_jv_ohne),
        "Reiseruecktritt_Jahres_Sparfuchs_mit_SB": euro_wert(spf_mit),
        "Reiseruecktritt_Jahres_Sparfuchs_ohne_SB": euro_wert(spf_ohne),

        "Reisekranken_Einmal_mit_SB": euro_wert(kranken_mit),
        "Reisekranken_Einmal_ohne_SB": euro_wert(kranken_ohne),
        "Reisekranken_Jahres_mit_SB": euro_wert(kv_jv_mit),
        "Reisekranken_Jahres_ohne_SB": euro_wert(kv_jv_ohne),

        "RundumSorglos_Einmal_mit_SB": euro_wert(rus_ev_mit),
        "RundumSorglos_Einmal_ohne_SB": euro_wert(rus_ev_ohne),
        "RundumSorglos_Jahres_mit_SB": euro_wert(rus_jv_mit),
        "RundumSorglos_Jahres_ohne_SB": euro_wert(rus_jv_ohne),
    }

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        angebot_path = export_doc("angebot.docx", tmp.name, daten)
        with open(angebot_path, "rb") as f:
            st.download_button(
                label="📄 Word-Datei herunterladen",
                data=f,
                file_name=f"angebot_{name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
