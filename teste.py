import streamlit as st
import pandas as pd
from time import sleep

st.title('Exemplo')

if 'executar' not in st.session_state:
    st.session_state.executar = False

if st.session_state.executar:
    progress_bar = st.progress(0, 'Carregando')
    for x in range(100):
        progress_bar.progress(x, 'Carregando')
        sleep(0.01)
    st.button('Hello')

if st.button('Executar') and not st.session_state.executar:
    st.session_state.executar = True