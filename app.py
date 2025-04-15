
import streamlit as st
import pandas as pd
from datetime import date, timedelta, datetime

# Seiteneinstellungen inkl. Favicon
st.set_page_config(
    page_title="ERGO Tarifrechner",
    page_icon="🧾",  # Fallback Emoji, kann durch logo ersetzt werden wenn lokal gehostet
    layout="centered"
)

# Logo anzeigen (zentriert)
st.image("SmallLogoBW_png.png", width=250)

# Titel
st.title("📊 ERGO Tarifrechner (mit Logo)")

# Eingabefelder
name = st.text_input("👤 Kundenname")
preis = st.number_input("💶 Reisepreis (€)", min_value=0.0)
ziel = st.text_input("🌍 Reiseziel (IATA-Code)")
alter = st.text_input("👥 Alter (z. B. 45 48)")
von = st.date_input("📅 Reise von", value=date.today(), format="DD.MM.YYYY")
bis = st.date_input("📅 Reise bis", value=date.today() + timedelta(days=7), format="DD.MM.YYYY")

if st.button("✅ Testdaten anzeigen"):
    st.write("Kunde:", name)
    st.write("Preis:", preis)
    st.write("Ziel:", ziel)
    st.write("Alter:", alter)
    st.write("Von:", von.strftime("%d.%m.%Y"))
    st.write("Bis:", bis.strftime("%d.%m.%Y"))
