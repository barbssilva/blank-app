import streamlit as st
import tempfile


st.title("Por favor, selecione o cliente.")

clientes = {
    "AW": "Alexander Wang",
    "AS": "AllSaints",
    "MH": "MadHappy",
    "MC":"Mochino",
}

# Variável para controlar se o uploader deve aparecer
if 'mostrar_uploader' not in st.session_state:
    st.session_state['mostrar_uploader'] = False

col1, col2 = st.columns(2)
with col1:
    if st.button(clientes['AW']):
        st.session_state['cliente_selecionado'] = clientes['AW']
        st.session_state['mostrar_uploader'] = True  # mostra uploader
with col2:
    if st.button(clientes['AS']):
        st.session_state['cliente_selecionado'] = clientes['AS']
        st.session_state['mostrar_uploader'] = True

col3, col4 = st.columns(2)
with col3:
    if st.button(clientes['MH']):
        st.session_state['cliente_selecionado'] = clientes['MH']
        st.session_state['mostrar_uploader'] = True
with col4:
    if st.button(clientes['MC']):
        st.session_state['cliente_selecionado'] = clientes['MC']
        st.session_state['mostrar_uploader'] = True

# Mostrar cliente selecionado
cliente = st.session_state.get('cliente_selecionado', None)

# Mostrar uploader só se mostrar_uploader for True
if st.session_state['mostrar_uploader']:
    uploaded_file = st.file_uploader("Carregue o ficheiro")

    if uploaded_file is not None:
        st.write(f"Ficheiro {uploaded_file.name} carregado!")
        
    if cliente:
        if cliente == 'Alexander Wang':
            # Exemplo: importar funções do script alexander_wang
            from alexander_wang import pdf_to_excel, convert_selected_columns, formatar_excel, remove_zeros, add_info

            # Salva o ficheiro PDF temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file)
                temp_pdf_path = temp_pdf.name
        
            excel_entrada = temp_pdf_path.replace(".pdf", ".xlsx")
            excel_saida = temp_pdf_path.replace(".pdf", "_processed.xlsx")
        
            styles, sample_sizes = pdf_to_excel(temp_pdf_path, excel_entrada)
            convert_selected_columns(excel_entrada, excel_saida)
            formatar_excel(excel_saida)
            remove_zeros(excel_saida)
            add_info(excel_saida, styles, sample_sizes)
        
            st.success("Processo terminado!")
        
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))
        
            os.remove(temp_pdf_path)
            os.remove(excel_entrada)
            os.remove(excel_saida)
