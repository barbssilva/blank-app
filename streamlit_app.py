import streamlit as st
import shutil
import tempfile
import os


st.title("Por favor, selecione o cliente.")

clientes = {
    "AW": "Alexander Wang",
    "AS": "AllSaints",
    "MH": "MadHappy",
    "MC":"Mochino",
}

col1, col2 = st.columns(2)
with col1:
    if st.button(clientes['AW']):
        st.session_state['cliente_selecionado'] = clientes['AW']
with col2:
    if st.button(clientes['AS']):
        st.session_state['cliente_selecionado'] = clientes['AS']

col3, col4 = st.columns(2)
with col3:
    if st.button(clientes['MH']):
        st.session_state['cliente_selecionado'] = clientes['MH']
with col4:
    if st.button(clientes['MC']):
        st.session_state['cliente_selecionado'] = clientes['MC']

# Mostrar cliente selecionado
cliente = st.session_state.get('cliente_selecionado', None)

if cliente:
    if cliente == 'Alexander Wang':
        # Exemplo: importar funções do script alexander_wang
        from alexander_wang import pdf_to_excel, convert_selected_columns, formatar_excel, remove_zeros, add_info
        uploaded_file = st.file_uploader("Carregue o PDF", type=["pdf"])
        
        # Obter nome original do ficheiro carregado
        original_filename = uploaded_file.name
        base_filename = os.path.splitext(original_filename)[0]  # sem extensão

        if uploaded_file is not None:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                # Copia diretamente o conteúdo do ficheiro carregado para o ficheiro temporário
                shutil.copyfileobj(uploaded_file, temp_pdf)
                #temp_pdf_path = temp_pdf.name

                # Criar caminho temporário usando o mesmo nome base
                temp_dir = tempfile.gettempdir()
                temp_pdf_path = os.path.join(temp_dir, original_filename)

            # Guardar o conteúdo do uploaded_file nesse caminho
            with open(temp_pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Agora o Excel vai usar o mesmo nome base
            excel_entrada = os.path.join(temp_dir, f"{base_filename}.xlsx")
            excel_saida = os.path.join(temp_dir, f"{base_filename}_processed.xlsx")
        
            # Agora já podes usar o ficheiro temporário no teu código:
            #excel_entrada = temp_pdf_path.replace(".pdf", ".xlsx")
            #excel_saida = temp_pdf_path.replace(".pdf", "_processed.xlsx")
        
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
