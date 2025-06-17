import streamlit as st


st.title("Por favor, selecione o cliente.")

clientes = {
    "AW": "Alexander Wang",
    "AS": "AllSaints",
    "MH": "MadHappy",
    "MC":"Mochino",
}

# Primeira linha com 2 colunas
col1, col2 = st.columns(2)
with col1:
    if st.button(clientes['AW']):
        st.session_state['cliente_selecionado'] = "Cliente A"
with col2:
    if st.button(clientes['AS']):
        st.session_state['cliente_selecionado'] = "Cliente B"

# Segunda linha com 2 colunas
col3, col4 = st.columns(2)
with col3:
    if st.button(clientes['MH']):
        st.session_state['cliente_selecionado'] = "Cliente C"
with col4:
    if st.button(clientes['MC']):
        st.session_state['cliente_selecionado'] = "Cliente D"

cliente = st.session_state.get('cliente_selecionado', None)

if cliente:
    st.write(f"Cliente selecionado: **{cliente}**")
