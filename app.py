import streamlit as st
from datetime import datetime, date

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

# Formularbereich
with st.form("eingabeformular"):
    name = st.text_input("Kundenname")
    zielgebiet = st.radio("Zielgebiet", ["Europa", "Welt"], index=0)
    preis = st.number_input("Reisepreis (€)", min_value=0.0)
    alter_text = st.text_input("Alter (z. B. 45 48)")

    col1, col2, col3, col4 = st.columns(4)
    gb_eingaben = []
    with col1:
        gb_eingaben.append(st.text_input("", key="gb1", label_visibility="collapsed"))
    with col2:
        gb_eingaben.append(st.text_input("", key="gb2", label_visibility="collapsed"))
    with col3:
        gb_eingaben.append(st.text_input("", key="gb3", label_visibility="collapsed"))
    with col4:
        gb_eingaben.append(st.text_input("", key="gb4", label_visibility="collapsed"))

    heute = date.today()
    geburts_alter = []
    for geb in gb_eingaben:
        gebdat = parse_geburtstag(geb)
        if gebdat:
            alter = heute.year - gebdat.year - ((heute.month, heute.day) < (gebdat.month, gebdat.day))
            geburts_alter.append(alter)

    if geburts_alter:
        st.markdown(
            f"<small>👥 Berechnete Alter: {', '.join(str(a) for a in geburts_alter)}</small>",
            unsafe_allow_html=True
        )
        alter_text = " ".join(str(a) for a in geburts_alter)

    von_raw = st.text_input("Reise von (TTMM oder TT.MM.JJJJ)")
    bis_raw = st.text_input("Reise bis (TTMM oder TT.MM.JJJJ)")
    submit = st.form_submit_button("Tarife anzeigen")

# Ausgabe zur Kontrolle
if submit:
    st.success("Formular erfolgreich abgeschickt.")
    st.write("Kundenname:", name)
    st.write("Zielgebiet:", zielgebiet)
    st.write("Reisepreis:", preis)
    st.write("Alter:", alter_text)
    st.write("Reisezeit:", von_raw, "-", bis_raw)
