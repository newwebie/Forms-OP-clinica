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
bio_file = st.secrets["sharepoint"]["bio_file"]

# Função para ler o arquivo Excel (Apontamentos) do SharePoint com cache
@st.cache_data
def get_sharepoint_file():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, file_name)
        return pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo no SharePoint: {e}")
        return pd.DataFrame()

# Função para ler o arquivo CSV (Estudos) do SharePoint com cache
@st.cache_data
def get_sharepoint_file_estudos_csv():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, bio_file)
        return pd.read_csv(io.BytesIO(response.content))
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo CSV de estudos no SharePoint: {e}")
        return pd.DataFrame()

# Função para atualizar o arquivo Excel (Apontamentos) no SharePoint
def update_sharepoint_file(df):
    try:
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        file_content = output.read()
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        folder_path = "/".join(file_name.split("/")[:-1])
        file_name_only = file_name.split("/")[-1]
        target_folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        target_folder.upload_file(file_name_only, file_content).execute_query()
        st.success("Apontamento salvo com sucesso!")
    except Exception as e:
        st.error(f"Erro ao salvar o arquivo no SharePoint: {e}")

# Carregar dados iniciais
df_study = get_sharepoint_file_estudos_csv()

# Inicializar o DataFrame de apontamentos no session_state
if "df_apontamentos" not in st.session_state:
    st.session_state["df_apontamentos"] = get_sharepoint_file()

# Configurar session_state para campos condicionais
if "status" not in st.session_state:
    st.session_state["status"] = "REALIZADO DURANTE A CONDUÇÃO"
if "enable_verificador" not in st.session_state:
    st.session_state["enable_verificador"] = False
if "enable_justificativa" not in st.session_state:
    st.session_state["enable_justificativa"] = False

def update_status_fields():
    s = st.session_state["status"]
    if s == "VERIFICANDO":
        st.session_state["enable_verificador"] = True
        st.session_state["enable_justificativa"] = False
    elif s == "NÃO APLICÁVEL":
        st.session_state["enable_verificador"] = False
        st.session_state["enable_justificativa"] = True
    else:
        st.session_state["enable_verificador"] = False
        st.session_state["enable_justificativa"] = False

# Início da tela principal
tabs = st.tabs(["Formulário", "Lista de Apontamentos"])

with tabs[0]:
    st.title("Criar Apontamento")
    
    if df_study.empty:
        st.error("Arquivo CSV de estudos não carregado. Verifique o caminho do arquivo.")
    else:
        protocol_options = ["Digite o codigo do estudo"] + df_study["NUMERO_DO_PROTOCOLO"].tolist()
        selected_protocol = st.selectbox("Código do Estudo", options=protocol_options, key="selected_protocol")
        if selected_protocol != "Digite o codigo do estudo":
            research_name = df_study.loc[df_study["NUMERO_DO_PROTOCOLO"] == selected_protocol, "NOME_DA_PESQUISA"].iloc[0]
        else:
            research_name = ""
        st.text_input("Nome da Pesquisa", value=research_name, disabled=True)
        
        responsavel = st.text_input("Responsável", key="responsavel")
        origem = st.selectbox("Origem Do Apontamento", 
                              ["Documentação Clínica", "Excelência Operacional", "Operações Clínicas", 
                               "Patrocinador / Monitor", "Garantia Da Qualidade"], key="origem")
        documento = st.selectbox("Documentos", [
            "Acompanhamento da Administração da Medicação", "Ajuste dos Relógios", "Anotação de enfermagem",
            "Aplicação do TCLE", "Ausência de Período", "Avaliação Clínica Pré Internação", "Avaliação de Alta Clínica",
            "Controle de Eliminações fisiológicas", "Controle de Glicemia", "Controle de Ausente de Período",
            "Controle de DropOut", "Critérios de Inclusão e Exclusão", "Desvio de ambulação", "Dieta",
            "Diretrizes do Protocolo", "Tabela de Controle de Preparo de Heparina", "TIME", "TCLE", "ECG",
            "Escala de Enfermagem", "Evento Adverso", "Ficha de internação", "Formulário de conferência das amostras",
            "Teste de HCG", "Teste de Drogas", "Teste de Álcool", "Término Prematuro",
            "Medicação para tratamento dos Eventos Adversos", "Orientação por escrito", "Prescrição Médica",
            "Registro de Temperatura da Enfermaria", "Relação dos Profissionais", "Sinais Vitais Pós Estudo",
            "SAE", "SINEB", "FOR 104", "FOR 123", "FOR 166", "FOR 217", "FOR 233", "FOR 234", "FOR 235",
            "FOR 236", "FOR 240", "FOR 241", "FOR 367"
        ], key="documento")
        participante = st.selectbox("Participante", 
            ['PP01', 'PP02', 'PP03', 'PP04', 'PP05', 'PP06', 'PP07', 'PP08', 'PP09', 'PP10', 
             'PP11', 'PP12', 'PP13', 'PP14', 'PP15', 'PP16', 'PP17', 'PP18', 'PP19', 'PP20', 
             'PP21', 'PP22', 'PP23', 'PP24', 'PP25', 'PP26', 'PP27', 'PP28', 'PP29', 'PP30', 
             'PP31', 'PP32', 'PP33', 'PP34', 'PP35', 'PP36', 'PP37', 'PP38', 'PP39', 'PP40', 
             'PP41', 'PP42', 'PP43', 'PP44', 'PP45', 'PP46', 'PP47', 'PP48', 'PP49', 'PP50', 
             'PP51', 'PP52', 'PP53', 'PP54', 'PP55', 'PP56', 'PP57', 'PP58', 'PP59', 'PP60', 
             'PP61', 'PP62', 'PP63', 'PP64', 'PP65', 'PP66', 'PP67', 'PP68', 'PP69', 'PP70', 
             'PP71', 'PP72', 'PP73', 'PP74', 'PP75', 'PP76', 'PP77', 'PP78', 'PP79', 'PP80', 
             'PP81', 'PP82', 'PP83', 'PP84', 'PP85', 'PP86', 'PP87', 'PP88', 'PP89', 'PP90', 
             'PP91', 'PP92', 'PP93', 'PP94', 'PP95', 'PP96', 'PP97', 'PP98', 'PP99'], key="participante")
             
        periodo = st.selectbox("Período", ['1° Período', '2° Período', '3° Período', '3° Período',
                                           '4° Período', '5° Período', '6° Período', '7° Período', 
                                           '8° Período', '9° Período', '10° Período' ], key="periodo")
        criticidade = st.selectbox("Grau De Criticidade Do Apontamento", ["Baixo", "Médio", "Alto"], key="criticidade")
        prazo = st.date_input("Prazo Pra Resolução", format="DD/MM/YYYY", key="prazo")
        apontamento = st.text_area("Apontamento", key="apontamento")
        
        # Campo de Status com callback
        status = st.selectbox("Status", [
            "REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"
        ], key="status", on_change=update_status_fields)
        
        if st.session_state["enable_verificador"]:
            verificador_nome = st.text_input("Nome de quem está verificando", key="verificador_nome")
            verificador_data = st.date_input("Data de verificação", format="DD/MM/YYYY", key="verificador_data")
            justificativa = ""
        elif st.session_state["enable_justificativa"]:
            justificativa = st.text_input("Justificativa", key="justificativa")
            verificador_nome = ""
            verificador_data = None
        else:
            verificador_nome = ""
            verificador_data = None
            justificativa = ""
        
        submit = st.button("Enviar")
        
        if submit:
            # Validação dos campos obrigatórios
            if selected_protocol == "Digite o codigo do estudo" or participante.strip() == "" or apontamento.strip() == "":
                st.error("Por favor, preencha os campos obrigatórios: Código do Estudo, Participante e Apontamento.")
            elif status == "VERIFICANDO" and verificador_nome.strip() == "":
                st.error("Por favor, preencha o campo 'Nome de quem está verificando'.")
            elif status == "NÃO APLICÁVEL" and justificativa.strip() == "":
                st.error("Por favor, preencha o campo 'Justificativa'!")
            else:
                novo_apontamento = {
                    "Código do Estudo": selected_protocol,
                    "Nome da Pesquisa": research_name,
                    "Data do Apontamento": datetime.date.today(),
                    "Responsável Pelo Apontamento": responsavel,
                    "Origem Do Apontamento": st.session_state["origem"],
                    "Documentos": st.session_state["documento"],
                    "Participante": st.session_state["participante"],
                    "Período": st.session_state["periodo"],
                    "Grau De Criticidade Do Apontamento": st.session_state["criticidade"],
                    "Prazo Pra Resolução": st.session_state["prazo"],
                    "Apontamento": st.session_state["apontamento"],
                    "Status": st.session_state["status"],
                    "Verificador": st.session_state.get("verificador_nome", ""),
                    "Data de Verificação": st.session_state.get("verificador_data", None),
                    "Justificativa": st.session_state.get("justificativa", ""),
                    "Responsável Pela Correção": "",
                    "Plantão": "",
                    "Departamento": "",
                    "Tempo de casa": ""
                }
                
                df = st.session_state["df_apontamentos"]
                duplicado_set = set(zip(df["Código do Estudo"], df["Documentos"], df["Participante"]))
                chave_nova = (selected_protocol, st.session_state["documento"], st.session_state["participante"])
                
                if chave_nova in duplicado_set:
                    duplicado = df[
                        (df["Código do Estudo"] == selected_protocol) &
                        (df["Documentos"] == st.session_state["documento"]) &
                        (df["Participante"] == st.session_state["participante"])
                    ]
                    data_existente = duplicado.iloc[0]["Data do Apontamento"]
                    st.warning(f"Apontamento já existe. Data do Apontamento: {data_existente}")
                else:
                    novo_df = pd.DataFrame([novo_apontamento])
                    df = pd.concat([df, novo_df], ignore_index=True)
                    update_sharepoint_file(df)
                    # Limpa o cache para forçar a recarga do arquivo do SharePoint
                    st.cache_data.clear()
                    # Recarrega os dados e atualiza o session_state
                    st.session_state["df_apontamentos"] = get_sharepoint_file()
                    st.success("Apontamento enviado com sucesso!")

with tabs[1]:
    st.title("Lista de Apontamentos")
    df = st.session_state["df_apontamentos"]
    if df.empty:
        st.info("Nenhum apontamento encontrado!")
    else:
        date_cols = ["Data do Apontamento", "Prazo Pra Resolução", "Data de Verificação"]
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        df.index = range(1, len(df) + 1)
        st.dataframe(df)
