import streamlit as st


st.title("Seleciona o Cliente")

clientes = {
    "Cliente A": "👩‍💼",
    "Cliente B": "👨‍🔧",
    "Cliente C": "🏢",
}

col1, col2, col3 = st.columns(3)

with col1:
    if st.button(f"{clientes['Cliente A']} Cliente A"):
        st.session_state['cliente_selecionado'] = "Cliente A"

with col2:
    if st.button(f"{clientes['Cliente B']} Cliente B"):
        st.session_state['cliente_selecionado'] = "Cliente B"

with col3:
    if st.button(f"{clientes['Cliente C']} Cliente C"):
        st.session_state['cliente_selecionado'] = "Cliente C"

cliente = st.session_state.get('cliente_selecionado', None)

if cliente:
    st.write(f"Cliente selecionado: **{cliente}**")
