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
    uploaded_file = st.file_uploader(f"Carregue o ficheiro para o cliente {clientes[cliente]}")

    if uploaded_file is not None:
        st.write(f"Ficheiro {uploaded_file.name} carregado!")

        # Agora roda o código do cliente AW se for esse o selecionado
        if cliente == 'AW':
            # Exemplo: importar funções do script AW (que já tem no teu projeto)
            from alexander_wang import pdf_to_excel, convert_selected_columns, formatar_excel, remove_zeros, add_info

            import tempfile
            import os

            # Salva o ficheiro PDF temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.read())
                temp_pdf_path = temp_pdf.name

            excel_entrada = temp_pdf_path.replace(".pdf", ".xlsx")
            excel_saida = temp_pdf_path.replace(".pdf", "_processed.xlsx")

            # Executa funções do script
            styles, sample_sizes = pdf_to_excel(temp_pdf_path, excel_entrada)
            convert_selected_columns(excel_entrada, excel_saida)
            formatar_excel(excel_saida)
            remove_zeros(excel_saida)
            add_info(excel_saida, styles, sample_sizes)

            st.success("Processamento terminado!")

            # Oferecer download do arquivo final
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))

            # Apaga arquivos temporários
            os.remove(temp_pdf_path)
            os.remove(excel_entrada)
            os.remove(excel_saida)
