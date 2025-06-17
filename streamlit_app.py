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
        st.session_state['cliente_selecionado'] = clientes["AW"]
with col2:
    if st.button(clientes['AS']):
        st.session_state['cliente_selecionado'] = clientes['AS']

# Segunda linha com 2 colunas
col3, col4 = st.columns(2)
with col3:
    if st.button(clientes['MH']):
        st.session_state['cliente_selecionado'] = clientes['MH']
with col4:
    if st.button(clientes['MC']):
        st.session_state['cliente_selecionado'] = clientes['MC']

cliente = st.session_state.get('cliente_selecionado', None)

#para exibir por escrito qual cliente está a ser selecionado no momento
#if cliente:
#    st.write(f"Cliente selecionado: **{cliente}**")
    

if cliente:

    # Upload do ficheiro (Excel, PDF, etc)
    uploaded_file = st.file_uploader("Carrega o ficheiro para processar")

    if uploaded_file is not None:
        st.write(f"Ficheiro {uploaded_file.name} carregado!")

        # Aqui podes chamar a tua função/script para processar o ficheiro
        # Exemplo simples: ler conteúdo se for texto
        # content = uploaded_file.read()
        # st.write(content)
