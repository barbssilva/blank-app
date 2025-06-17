import streamlit as st
import shutil
import tempfile
import os


st.title("Por favor, selecione o cliente.")

clientes = {
    "AW": "Alexander Wang",
    "AS": "AllSaints",
    "MH": "Madhappy",
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
    if cliente == 'Madhappy':
        # Exemplo: importar funções do script madhappy
        from madhappy import inches_to_cm, decimal_para_fracao, selecionar_tabelas, convert_selected_columns, formatar_excel
        uploaded_file = st.file_uploader("Carregue o Excel", type=["xls", "xlsx"])

        if uploaded_file is not None:
            base_name = os.path.splitext(uploaded_file.name)[0]
        
            # Criar ficheiro excel temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel:
                temp_excel.write(uploaded_file.read())
                temp_excel_path = temp_pdf.name
        
            # Agora cria o excel_entrada e excel_saida no mesmo diretório do ficheiro temporário,
            # mas com nomes baseados no ficheiro original:
            temp_dir = os.path.dirname(temp_excel_path)
            #definir o nome do ficheiro excel para o qual será transferida a informação do pdf
            excel_entrada = os.path.join(temp_dir, base_name + ".xlsx")
            #definir o nome do ficheiro excel que irá conter as alterações: conversão para cm e calculo da diferença entre tamanhos
            excel_saida = os.path.join(temp_dir, base_name + "_processed.xlsx")

            keywords= ['1st proto', 'sms','size chart','spec']
            output_file1 = selecionar_tabelas(excel_entrada,keywords,excel_saida)

            #processamento do excel criado
            convert_selected_columns(excel_saida)

            #formatar excel
            formatar_excel(excel_saida)

            st.success("Processo terminado!")
        
            # Abrir o ficheiro Excel processado para download
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))
        
    if cliente == 'Alexander Wang':
        # Exemplo: importar funções do script alexander_wang
        from alexander_wang import pdf_to_excel, convert_selected_columns, formatar_excel, remove_zeros, add_info
        uploaded_file = st.file_uploader("Carregue o excel", type=["pdf"])

        if uploaded_file is not None:
            base_name = os.path.splitext(uploaded_file.name)[0]
        
            # Criar ficheiro PDF temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.read())
                temp_pdf_path = temp_pdf.name
        
            # Agora cria o excel_entrada e excel_saida no mesmo diretório do ficheiro temporário,
            # mas com nomes baseados no ficheiro original:
            temp_dir = os.path.dirname(temp_pdf_path)
            excel_entrada = os.path.join(temp_dir, base_name + ".xlsx")
            excel_saida = os.path.join(temp_dir, base_name + "_processed.xlsx")
        
            # Executar processamento (exemplo)
            styles, sample_sizes = pdf_to_excel(temp_pdf_path, excel_entrada)
            convert_selected_columns(excel_entrada, excel_saida)
            formatar_excel(excel_saida)
            remove_zeros(excel_saida)
            add_info(excel_saida, styles, sample_sizes)
        
            st.success("Processo terminado!")
        
            # Abrir o ficheiro Excel processado para download
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))
        
            # Apagar ficheiros temporários
            os.remove(temp_pdf_path)
            os.remove(excel_entrada)
            os.remove(excel_saida)
