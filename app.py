import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from gerador import gerar_certificado

# Função para verificar a validade do Dashboard
def verify_pmb_dashboard(user_file):
    try:
        # Ler o arquivo Excel
        excelData = pd.ExcelFile(user_file)

        # Verificar se a aba "Backend" existe
        if "Backend" not in excelData.sheet_names:
            return "O arquivo enviado não foi reconhecido como um Dashboard do PMB: Aba 'Backend' não encontrada. Verifique se enviou o arquivo correto."

        # Ler a aba Backend sem cabeçalhos
        df_backend = pd.read_excel(user_file, sheet_name="Backend", header=None)

        # Verificar o valor da célula A2
        if df_backend.iloc[1, 0] not in ["Starters", "Movers", "Flyers"]:
            return "O arquivo enviado não foi reconhecido como um Dashboard válido: Verifique se o dashboard enviado está atualizado"
        
        return "Arquivo reconhecido como um Dashboard do PMB"
    
    except Exception as e:
        return f"Erro ao processar o arquivo: {e}"

# Função para calcular os Scores (Scores são as estrelas do certificado)
def calculate_scores(user_file):

    # Ler a aba backend e verificar o nível do Mock test do dashboard
    df_backend = pd.read_excel(user_file, sheet_name="Backend", header=None)
    mock_level = df_backend.iloc[1, 0]

    # Limiar para cada nível
    # Sintaxe do dicionário: Nº de Troféus: Limiar
    # O Limiar (threshold) é o mínimo de pontos para obter determinado troféu
    thresholds = {
        "Starters": {
            "ls": {1: 0, 2: 11, 3: 13, 4: 16, 5: 18},
            "rw": {1: 0, 2: 13, 3: 16, 4: 19, 5: 21},
            "sp": {1: 0, 2: 3, 3: 7, 4: 10, 5: 12}
        },
        "Movers": {
            "ls": {1: 0, 2: 11, 3: 14, 4: 18, 5: 21},
            "rw": {1: 0, 2: 18, 3: 24, 4: 29, 5: 33},
            "sp": {1: 0, 2: 3, 3: 7, 4: 10, 5: 12}
        },
        "Flyers": {
            "ls": {1: 0, 2: 14, 3: 17, 4: 20, 5: 23},
            "rw": {1: 0, 2: 24, 3: 30, 4: 36, 5: 42},
            "sp": {1: 0, 2: 3, 3: 7, 4: 10, 5: 12}
        }
    }

    # Ler a aba Table
    df_table = pd.read_excel(user_file, sheet_name="Table", skiprows=19)

    # Verificar o nível do Mock test e calcular os scores conforme o nível
    if mock_level == "Movers":
        # Converter colunas para números
        df_table["Resultado do Estudante Listening"] = pd.to_numeric(df_table["Resultado do Estudante Listening"], errors='coerce')
        df_table["Resultado do Estudante RW"] = pd.to_numeric(df_table["Resultado do Estudante RW"], errors='coerce')
        df_table["Resultado do Estudante Speaking"] = pd.to_numeric(df_table["Resultado do Estudante Speaking"], errors='coerce')

        # Criar o dataframe com as colunas necessárias e as modificações
        df_students = pd.DataFrame({
            'Aluno': df_table.iloc[:, 0],
            'Ano': df_table.iloc[:, 1],
            'Turma': df_table.iloc[:, 2],
            'Listening': df_table["Resultado do Estudante Listening"] * 25,
            'Reading & Writing': df_table["Resultado do Estudante RW"] * 35,
            'Speaking': df_table["Resultado do Estudante Speaking"] * 15
        })  

        # Função para calcular os troféus
        def get_trofes(value, thresholds):
            if pd.isna(value):
                return 0  # ou qualquer valor padrão que faça sentido para você
            for t, th in sorted(thresholds.items(), reverse=True):
                if value >= th:
                    return t
            return 0

        # Calcular os troféus com base nos limiares
        df_students['Troféus Listening'] = df_students['Listening'].apply(lambda x: get_trofes(x, thresholds["Movers"]["ls"]))
        df_students['Troféus Reading & Writing'] = df_students['Reading & Writing'].apply(lambda x: get_trofes(x, thresholds["Movers"]["rw"]))
        df_students['Troféus Speaking'] = df_students['Speaking'].apply(lambda x: get_trofes(x, thresholds["Movers"]["sp"]))

        st.dataframe(df_students)
    return df_students

### COMEÇO DO STREAMLIT APP ###
# Título do App
st.title('PMB Statement of Results Generator')
st.text('Current version 1.0')

# Importar o arquivo de Excel
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")

if uploaded_file is not None:
    # Chamar a função de verificação de erros
    result = verify_pmb_dashboard(uploaded_file)

    # Exibir o resultado da verificação
    if result == "Arquivo reconhecido como um Dashboard do PMB":
        st.success(result)
        # Chamar a função de cálculo de scores
        df_scores = calculate_scores(uploaded_file)
        # Exibir o dataframe no app
        st.dataframe(df_scores)

        # Gerar certificados e criar arquivo ZIP
        with tempfile.TemporaryDirectory() as tempdir:
            pdf_files = []
            for index, row in df_scores.iterrows():
                pdf_file = gerar_certificado(
                    aluno=row['Aluno'],
                    ls_stars=row['Troféus Listening'],
                    rw_stars=row['Troféus Reading & Writing'],
                    sp_stars=row['Troféus Speaking'],
                    output_dir=tempdir
                )
                if pdf_file:
                    pdf_files.append(pdf_file)

            # Criar arquivo ZIP
            zip_path = os.path.join(tempdir, "certificados.zip")
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for file in pdf_files:
                    zipf.write(file, os.path.basename(file))

            # Exibir botão para download do arquivo ZIP
            with open(zip_path, "rb") as f:
                bytes = f.read()
                st.download_button(
                    label="Baixar certificados",
                    data=bytes,
                    file_name="certificados.zip",
                    mime="application/zip"
                )
    else:
        st.error(result)
