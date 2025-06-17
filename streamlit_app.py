import streamlit as st


st.title("Por favor, selecione o cliente.")

clientes = {
    "AW": "Alexander Wang",
    "AS": "AllSaints",
    "MH": "MadHappy",
    "MC":"Mochino",
}

if 'mostrar_uploader' not in st.session_state:
    st.session_state['mostrar_uploader'] = False
if 'cliente_selecionado' not in st.session_state:
    st.session_state['cliente_selecionado'] = None

# Linha 1 dos clientes
col1, col2, col_uploader = st.columns([1,1,2])  # col_uploader para uploader aparecer ao lado

with col1:
    if st.button(clientes['AW']):
        st.session_state['cliente_selecionado'] = clientes['AW']
        st.session_state['mostrar_uploader'] = True
with col2:
    if st.button(clientes['AS']):
        st.session_state['cliente_selecionado'] = clientes['AS']
        st.session_state['mostrar_uploader'] = True

# Linha 2 dos clientes
col3, col4 = st.columns([1,1])
with col3:
    if st.button(clientes['MH']):
        st.session_state['cliente_selecionado'] = clientes['MH']
        st.session_state['mostrar_uploader'] = True
with col4:
    if st.button(clientes['MC']):
        st.session_state['cliente_selecionado'] = clientes['MC']
        st.session_state['mostrar_uploader'] = True

# Mostrar uploader na coluna da direita se ativado
with col_uploader:
    if st.session_state['mostrar_uploader']:
        st.write(f"Cliente selecionado: **{st.session_state['cliente_selecionado']}**")
        uploaded_file = st.file_uploader("Carregue o ficheiro", key='upload')
        # Botão para fechar a área do uploader (aqui o "x")
        if st.button("❌ Fechar upload"):
            st.session_state['mostrar_uploader'] = False
            st.session_state['cliente_selecionado'] = None
            st.experimental_rerun()  # força atualizar para limpar uploader

        if uploaded_file is not None:
            st.write(f"Ficheiro {uploaded_file.name} carregado!")
