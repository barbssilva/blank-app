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
    "MR": "Moncler",
}

col1, col2 = st.columns(2)
with col1:
    if st.button(clientes['AW']):
        st.session_state['cliente_selecionado'] = clientes['AW']

with col1:
    if st.button(clientes['MR']):  
        st.session_state['cliente_selecionado'] = clientes['MR']
        
with col1:
    if st.button(clientes['MH']):
        st.session_state['cliente_selecionado'] = clientes['MH']
        
with col2:
    if st.button(clientes['AS']):
        st.session_state['cliente_selecionado'] = clientes['AS']

with col2:
    if st.button(clientes['MC']):
        st.session_state['cliente_selecionado'] = clientes['MC']

# Mostrar cliente selecionado
cliente = st.session_state.get('cliente_selecionado', None)

if cliente:
    if cliente == "Moncler":
        # Exemplo: importar funções do script moncler
        from moncler import pdf_to_excel, excel_processing, dif_calc, formatar_excel, add_images
        uploaded_file = st.file_uploader("Carregue o pdf", type=["pdf"])
        
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


            #conversão de pdf para excel
            #texto_df é guardado caso depois seja preciso adicionar o nome do modelo ao excel no fim
            texto_df=pdf_to_excel(temp_pdf_path,excel_entrada)
            inf_modelo = [texto_df.iloc[0].item(), texto_df.iloc[1].item()]
            # procurar pelo nome do modelo

            excel_processing(excel_entrada, excel_saida)    
        
            dif_calc(excel_saida)
        
            formatar_excel(excel_saida)

            add_images(temp_pdf_path,excel_saida, inf_modelo)
                    
            st.success("Processo terminado!")
        
            # Abrir o ficheiro Excel processado para download
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))
        
            #Remover o primeiro ficheiro excel criado uma vez que já não será utilizado
            os.remove(excel_entrada)

    
    if cliente == 'Madhappy':
        # Exemplo: importar funções do script madhappy
        from madhappy import inches_to_cm, decimal_para_fracao, selecionar_tabelas, convert_selected_columns, formatar_excel
        uploaded_file = st.file_uploader("Carregue o Excel", type=["xls", "xlsx"])

        if uploaded_file is not None:
            base_name = os.path.splitext(uploaded_file.name)[0]
        
            # Criar ficheiro excel temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel:
                temp_excel.write(uploaded_file.read())
                temp_excel_path = temp_excel.name
        
            # Agora cria o excel_entrada e excel_saida no mesmo diretório do ficheiro temporário,
            # mas com nomes baseados no ficheiro original:
            temp_dir = os.path.dirname(temp_excel_path)
            #definir o nome do ficheiro excel que irá conter as alterações: conversão para cm e calculo da diferença entre tamanhos
            excel_saida = os.path.join(temp_dir, base_name + "_processed.xlsx")

            # Criar um novo arquivo temporário para salvar as alterações
            base_name = os.path.basename(excel_saida)               # só o nome do ficheiro
            new_name = base_name.replace(".xlsx", "_temp.xlsx")     # troca a extensão
            tempor_said = os.path.join(temp_dir, new_name)  

            #nome ficheiro auxiliar
            output_file = os.path.join(temp_dir, base_name + "_aux.xlsx")
            
            keywords= ['1st proto', 'sms','size chart','spec']
            output_file1 = selecionar_tabelas(temp_excel_path,keywords,excel_saida,output_file)

            #processamento do excel criado
            convert_selected_columns(excel_saida)

            #formatar excel
            formatar_excel(excel_saida)
            
            st.success("Processo terminado!")
        
            # Abrir o ficheiro Excel processado para download
            with open(excel_saida, "rb") as f:
                st.download_button("Descarregar Excel Processado", f, file_name=os.path.basename(excel_saida))

            os.remove(output_file)
        
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
