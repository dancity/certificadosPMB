import streamlit as st
import pandas as pd

st.title('Exemplo')

if 'generate_button' not in st.session_state:
    st.session_state.generate_button = st.button('Clique aqui')