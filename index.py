import streamlit as st
import pandas as pd
import zipfile
import os
import tempfile
from gerador import gerar_certificado

def verify_pmb_dashboard(user_file):
    #Função para verificar a validade do Dashboard.
    try:
        excelData = pd.ExcelFile(user_file)
        if "Backend" not in excelData.sheet_names:
            return False, "O arquivo enviado não foi reconhecido como um Dashboard do PMB: Aba 'Backend' não encontrada. Verifique se enviou o arquivo correto."

        df_backend = pd.read_excel(user_file, sheet_name="Backend", header=None)
        if df_backend.iloc[1, 0] not in ["Starters", "Movers", "Flyers"]:
            return False, "O arquivo enviado não foi reconhecido como um Dashboard válido: Verifique se o dashboard enviado está atualizado"
        
        df_table = pd.read_excel(user_file, sheet_name="Table", skiprows=19)
        if len(df_table) < 1:
            return False, "A tabela de alunos está em branco!"

        return True, "Sucesso! Arquivo reconhecido como um Dashboard do PMB"
    
    except Exception as e:
        return False, f"Erro ao processar o arquivo: {e}"

def calculate_scores(user_file):
    # Função para calcular os Scores (Scores são as estrelas do certificado).
    df_backend = pd.read_excel(user_file, sheet_name="Backend", header=None)
    mock_level = df_backend.iloc[1, 0]

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

    # Nota: Os dados do Movers e Flyers precisam ser alterados!
    max_scores = {
        "Starters": {"Listening": 20, "Reading & Writing": 30, "Speaking": 10},
        "Movers": {"Listening": 25, "Reading & Writing": 35, "Speaking": 15},
        "Flyers": {"Listening": 30, "Reading & Writing": 40, "Speaking": 20}
    }

    # Cria um dataframe com os dados de cada aluno e seus troféus
    df_table = pd.read_excel(user_file, sheet_name="Table", skiprows=19)
    if mock_level in max_scores:
        df_table["Resultado do Estudante Listening"] = pd.to_numeric(df_table["Resultado do Estudante Listening"], errors='coerce')
        df_table["Resultado do Estudante RW"] = pd.to_numeric(df_table["Resultado do Estudante RW"], errors='coerce')
        df_table["Resultado do Estudante Speaking"] = pd.to_numeric(df_table["Resultado do Estudante Speaking"], errors='coerce')

        max_score = max_scores[mock_level]
        df_students = pd.DataFrame({
            'Aluno': df_table.iloc[:, 0],
            'Ano': df_table.iloc[:, 1],
            'Turma': df_table.iloc[:, 2],
            'Listening': df_table["Resultado do Estudante Listening"] * max_score["Listening"],
            'Reading & Writing': df_table["Resultado do Estudante RW"] * max_score["Reading & Writing"],
            'Speaking': df_table["Resultado do Estudante Speaking"] * max_score["Speaking"]
        })  

        def get_trofes(value, thresholds):
            if pd.isna(value):
                return 0
            for t, th in sorted(thresholds.items(), reverse=True):
                if value >= th:
                    return t
            return 0

        df_students['Troféus Listening'] = df_students['Listening'].apply(lambda x: get_trofes(x, thresholds[mock_level]["ls"]))
        df_students['Troféus Reading & Writing'] = df_students['Reading & Writing'].apply(lambda x: get_trofes(x, thresholds[mock_level]["rw"]))
        df_students['Troféus Speaking'] = df_students['Speaking'].apply(lambda x: get_trofes(x, thresholds[mock_level]["sp"]))

    return df_students, mock_level

def generate_certificates(df_scores, mock_level):
    #Função para gerar certificados e criar arquivo ZIP.
    tempdir = tempfile.mkdtemp()
    pdf_files = []
    total_alunos = len(df_scores)
    loader_bar = st.progress(0, text="Gerando certificados")
    for index, row in df_scores.iterrows():
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

    zip_path = os.path.join(tempdir, "certificados.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in pdf_files:
            zipf.write(file, os.path.basename(file))

    st.session_state.zip_path = zip_path

def main():
    #Função principal para execução do Streamlit.
    st.title('PMB Statement of Results Generator')
    st.text('Current version 1.0')

    uploaded_file = st.file_uploader("Faça upload do Dashboard Mocktest Preenchido", type="xlsx")
    action_button = st.empty()

    if 'zip_path' not in st.session_state:
        st.session_state.zip_path = ""

    if uploaded_file is not None:
        valid, result = verify_pmb_dashboard(uploaded_file)
        st.toast(result)

        if valid:
            df_scores, mock_level = calculate_scores(uploaded_file)

            generate_btn = action_button.button('Gerar certificados', disabled=False, key='1')

            if generate_btn:
                generate_btn = action_button.button('Gerar certificados', disabled=True, key='2')
                generate_certificates(df_scores, mock_level)
                with open(st.session_state.zip_path, "rb") as f:
                    bytes = f.read()
                    download_button = action_button.download_button(
                        label="Baixar certificados",
                        data=bytes,
                        file_name="certificados.zip",
                        mime="application/zip",
                        type='primary'
                    )

if __name__ == "__main__":
    main()
