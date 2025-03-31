import streamlit as st
import pandas as pd
import datetime
import io
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential

# === Configurações do SharePoint ===
username = st.secrets["sharepoint"]["username"]
password = st.secrets["sharepoint"]["password"]
site_url = st.secrets["sharepoint"]["site_url"]  
file_name = st.secrets["sharepoint"]["file_name"] 

# Função para ler o arquivo Excel do SharePoint
def get_sharepoint_file():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, file_name)
        return pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo no SharePoint: {e}")
        return pd.DataFrame()

# Função para atualizar (fazer upload) do arquivo Excel no SharePoint
def update_sharepoint_file(df):
    try:
        # Converte o DataFrame para Excel em memória
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        file_content = output.read()
        
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        # Extrai o caminho da pasta e o nome do arquivo
        folder_path = "/".join(file_name.split("/")[:-1])
        file_name_only = file_name.split("/")[-1]
        target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        target_folder.upload_file(file_name_only, file_content).execute_query()
        st.success("Apontamento salvo com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar o arquivo no SharePoint: {e}")

# === Aplicativo Streamlit com duas abas: Formulário e Lista de Apontamentos ===
tabs = st.tabs(["Formulário", "Lista de Apontamentos"])

with tabs[0]:
    st.title("Criar Apontamento")

    def form():
        with st.form(key="Despesa"):
            cod_estudo = st.text_input("Código do Estudo")
            data_apontamento = st.date_input("Data do Apontamento", format="DD/MM/YYYY")
            responsavel = st.text_input("Responsável")
            origem = st.selectbox("Origem Do Apontamento", [
                "Documentação Clínica", "Excelência Operacional", "Operações Clínicas", 
                "Patrocinador / Monitor", "Garantia Da Qualidade"
            ])
            documento = st.selectbox("Documentos", [
                "Acompanhamento da Administração da Medicação",
                "Ajuste dos Relógios",
                "Anotação de enfermagem",
                "Aplicação do TCLE",
                "Ausência de Período",
                "Avaliação Clínica Pré Internação",
                "Avaliação de Alta Clínica",
                "Controle de Eliminações fisiológicas", 
                "Controle de Glicemia",
                "Controle de Ausente de Período",
                "Controle de DropOut",
                "Critérios de Inclusão e Exclusão",
                "Desvio de ambulação",
                "Dieta",
                "Diretrizes do Protocolo",
                "Tabela de Controle de Preparo de Heparina",
                "TIME",
                "TCLE",
                "ECG",
                "Escala de Enfermagem",
                "Evento Adverso",
                "Ficha de internação",
                "Formulário de conferência das amostras",
                "Teste de HCG",
                "Teste de Drogas",
                "Teste de Álcool",
                "Término Prematuro",
                "Medicação para tratamento dos Eventos Adversos",
                "Orientação por escrito",
                "Prescrição Médica",
                "Registro de Temperatura da Enfermaria",
                "Relação dos Profissionais",
                "Sinais Vitais Pós Estudo",
                "SAE",
                "SINEB",
                "FOR 104",
                "FOR 123",
                "FOR 166",
                "FOR 217",
                "FOR 233",
                "FOR 234",
                "FOR 235",
                "FOR 236",
                "FOR 240",
                "FOR 241",
                "FOR 367"
            ])
            participante = st.text_input("Participante")
            periodo = st.text_input("Período")
            criticidade = st.selectbox("Grau De Criticidade Do Apontamento", ["Baixo", "Médio", "Alto"])
            prazo = st.date_input("Prazo Pra Resolução", format="DD/MM/YYYY")
            apontamento = st.text_area("Apontamento")
            
            status = st.selectbox("Status", [
                "REALIZADO DURANTE A CONDUÇÃO",
                "REALIZADO",
                "VERIFICANDO",
                "PENDENTE",
                "NÃO APLICÁVEL"
            ])
            
            verificador_nome = ""
            verificador_data = None
            justificativa = ""
            
            if status == "VERIFICANDO":
                verificador_nome = st.text_input("Nome de quem está verificando")
                verificador_data = st.date_input("Data de verificação", format="DD/MM/YYYY", key="verif_data")
            if status == "NÃO APLICÁVEL":
                justificativa = st.text_input("Justificativa")
            
            submit_button = st.form_submit_button(label="Enviar")
            
            if submit_button:
                if cod_estudo.strip() == "" or participante.strip() == "" or apontamento.strip() == "":
                    st.error("Por favor, preencha os campos obrigatórios: Código do Estudo, Participante e Apontamento.")
                    return
                
                if status == "VERIFICANDO" and verificador_nome.strip() == "":
                    st.error("Por favor, preencha o campo 'Nome de quem está verificando'.")
                    return
                if status == "NÃO APLICÁVEL" and justificativa.strip() == "":
                    st.error("Por favor, preencha o campo 'Justificativa'.")
                    return
                
                novo_apontamento = {
                    "Código do Estudo": cod_estudo,
                    "Data do Apontamento": data_apontamento,
                    "Responsável Pelo Apontamento": responsavel,
                    "Origem Do Apontamento": origem,
                    "Documentos": documento,
                    "Participante": participante,
                    "Período": periodo,
                    "Grau De Criticidade Do Apontamento": criticidade,
                    "Prazo Pra Resolução": prazo,
                    "Apontamento": apontamento,
                    "Status": status,
                    "Verificador": verificador_nome,
                    "Data de Verificação": verificador_data,
                    "Justificativa": justificativa,
                    "Responsável Pela Correção": "",
                    "Plantão": "",
                    "Departamento": "",
                    "Tempo de casa": ""
                }
                
                df = get_sharepoint_file()
                duplicado = df[
                    (df["Código do Estudo"] == cod_estudo) &
                    (df["Documentos"] == documento) &
                    (df["Participante"] == participante)
                ]
                
                if not duplicado.empty:
                    data_existente = duplicado.iloc[0]["Data do Apontamento"]
                    st.warning(f"Apontamento já existe. Data do Apontamento: {data_existente}")
                else:
                    novo_df = pd.DataFrame([novo_apontamento])
                    df = pd.concat([df, novo_df], ignore_index=True)
                    update_sharepoint_file(df)
                    
    form()

with tabs[1]:
    st.title("Lista de Apontamentos")
    df = get_sharepoint_file()
    if df.empty:
        st.info("Nenhum apontamento encontrado!")
    else:
        date_cols = ["Data do Apontamento", "Prazo Pra Resolução", "Data de Verificação"]
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        df.index = range(1, len(df) + 1)
        st.dataframe(df)