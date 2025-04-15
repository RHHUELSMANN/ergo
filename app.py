
import streamlit as st
import pandas as pd
from datetime import date, timedelta, datetime

# Seiteneinstellungen inkl. Titel & Icon
st.set_page_config(
    page_title="Der Ergo Chuck",
    page_icon="🦾",
    layout="centered"
)

# Logo anzeigen
st.image("SmallLogoBW_png.png", width=250)

# Titel
st.title("🦾 Der Ergo Chuck")

# Eingabemaske
name = st.text_input("👤 Kundenname")
preis = st.number_input("💶 Reisepreis (€)", min_value=0.0)
ziel = st.text_input("🌍 Reiseziel (IATA-Code)")
alter = st.text_input("👥 Alter der Reisenden (z. B. 45 48)")
von = st.date_input("📅 Reisedatum von", value=date.today(), format="DD.MM.YYYY")
bis = st.date_input("📅 Reisedatum bis", value=date.today() + timedelta(days=7), format="DD.MM.YYYY")

if st.button("✅ Tarife anzeigen"):
    st.success("✅ Eingabe verarbeitet:")
    st.write("Kunde:", name)
    st.write("Preis:", preis)
    st.write("Ziel:", ziel)
    st.write("Alter:", alter)
    st.write("Von:", von.strftime("%d.%m.%Y"))
    st.write("Bis:", bis.strftime("%d.%m.%Y"))
