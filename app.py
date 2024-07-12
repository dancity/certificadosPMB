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
        
        # Verificar se a tabela de alunos está preenchida
        df_table = pd.read_excel(user_file, sheet_name="Table", skiprows=19)
        if len(df_table) < 1:
            return "A tabela de alunos está em branco!"

        return "Sucesso! Arquivo reconhecido como um Dashboard do PMB"
    
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

    # Multiplicadores para cada nível
    max_scores = {
        "Starters": {"Listening": 20, "Reading & Writing": 30, "Speaking": 10},
        "Movers": {"Listening": 25, "Reading & Writing": 35, "Speaking": 15},
        "Flyers": {"Listening": 30, "Reading & Writing": 40, "Speaking": 20}
    }

    # Ler a aba Table
    df_table = pd.read_excel(user_file, sheet_name="Table", skiprows=19)

    # Verificar o nível do Mock test e calcular os scores conforme o nível
    if mock_level in max_scores:
        # Converter colunas para números
        df_table["Resultado do Estudante Listening"] = pd.to_numeric(df_table["Resultado do Estudante Listening"], errors='coerce')
        df_table["Resultado do Estudante RW"] = pd.to_numeric(df_table["Resultado do Estudante RW"], errors='coerce')
        df_table["Resultado do Estudante Speaking"] = pd.to_numeric(df_table["Resultado do Estudante Speaking"], errors='coerce')

        # Aplicar os multiplicadores com base no nível
        max_score = max_scores[mock_level]
        df_students = pd.DataFrame({
            'Aluno': df_table.iloc[:, 0],
            'Ano': df_table.iloc[:, 1],
            'Turma': df_table.iloc[:, 2],
            'Listening': df_table["Resultado do Estudante Listening"] * max_score["Listening"],
            'Reading & Writing': df_table["Resultado do Estudante RW"] * max_score["Reading & Writing"],
            'Speaking': df_table["Resultado do Estudante Speaking"] * max_score["Speaking"]
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
        df_students['Troféus Listening'] = df_students['Listening'].apply(lambda x: get_trofes(x, thresholds[mock_level]["ls"]))
        df_students['Troféus Reading & Writing'] = df_students['Reading & Writing'].apply(lambda x: get_trofes(x, thresholds[mock_level]["rw"]))
        df_students['Troféus Speaking'] = df_students['Speaking'].apply(lambda x: get_trofes(x, thresholds[mock_level]["sp"]))

    return df_students, mock_level

### COMEÇO DO STREAMLIT APP ###
# Título do App
st.title('PMB Statement of Results Generator')
st.text('Current version 1.0')

# Inicializar estados se não existirem
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'valid_file' not in st.session_state:
    st.session_state.valid_file = False
if 'generate_certificates' not in st.session_state:
    st.session_state.generate_certificates = False

# Importar o arquivo de Excel
if st.session_state.uploaded_file is None:
    uploaded_file = st.file_uploader("Faça upload do Dashboard Mocktest Preenchido", type="xlsx")

    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file

if st.session_state.uploaded_file is not None:
    # Chamar a função de verificação de erros
    result = verify_pmb_dashboard(st.session_state.uploaded_file)

    # Exibir o resultado da verificação
    if result == "Sucesso! Arquivo reconhecido como um Dashboard do PMB":
        st.success(result)
        st.session_state.valid_file = True

        # Chamar a função de cálculo de scores
        df_scores, mock_level = calculate_scores(st.session_state.uploaded_file)
        if df_scores is not None:
            st.session_state.df_scores = df_scores
            # Botão para gerar certificados
            if not st.session_state.generate_certificates:
                if st.button("Gerar certificados"):
                    st.toast('Certificados serão gerados!', icon="✨")
                    st.session_state.generate_certificates = True

    else:
        st.error(result)
        st.session_state.uploaded_file = None
        st.session_state.valid_file = False

if st.session_state.valid_file and st.session_state.generate_certificates:
    loader_bar = st.progress(0, text="Gerando certificados")

    # Gerar certificados e criar arquivo ZIP
    with tempfile.TemporaryDirectory() as tempdir:
        pdf_files = []
        total_alunos = len(st.session_state.df_scores)
        for index, row in st.session_state.df_scores.iterrows():
            loader_bar.progress((index + 1) / total_alunos, text=f"Gerando o certificado do aluno(a) {row['Aluno']}")
            pdf_file = gerar_certificado(
                aluno=row['Aluno'],
                ls_stars=row['Troféus Listening'],
                rw_stars=row['Troféus Reading & Writing'],
                sp_stars=row['Troféus Speaking'],
                output_dir=tempdir,
                mock_level=mock_level
            )
            if pdf_file:
                pdf_files.append(pdf_file)

        loader_bar.empty()

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
                mime="application/zip",
                type='primary'
            )