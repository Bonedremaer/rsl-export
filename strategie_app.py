# -*- coding: utf-8 -*-
"""RSL Strategie-App — Web-App (Streamlit Community Cloud / iPhone) — APP 2 von 2.

Eigenständige zweite App: bettet die komplette 5-Tab-App
(Signal · Ranking · Journal · Kurz · Doku) aus RSL_Strategie_App.html ein.
Das Journal speichert seine Einträge im Browser (localStorage des iFrames).
Keine Live-Kurse — die Daten sind ein Snapshot (siehe Datenstand in der App).
"""
import os
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="RSL Strategie-App", page_icon="📱",
                   layout="centered", initial_sidebar_state="collapsed")

# Streamlit-Rahmen ausblenden, damit die App-Fläche möglichst groß ist
st.markdown("""
<style>
#MainMenu, header, footer {visibility: hidden;}
.block-container {padding: 0 !important; max-width: 100% !important;}
[data-testid="stAppViewContainer"] > .main {padding: 0 !important;}
</style>
""", unsafe_allow_html=True)

pfad = os.path.join(os.path.dirname(os.path.abspath(__file__)), "RSL_Strategie_App.html")
try:
    with open(pfad, encoding="utf-8") as f:
        html = f.read()
except Exception:
    with open("RSL_Strategie_App.html", encoding="utf-8") as f:
        html = f.read()

components.html(html, height=900, scrolling=True)
