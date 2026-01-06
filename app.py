import streamlit as st
import pandas as pd
from datetime import datetime
import io
import time
import random
import string
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# >>> usa o conector (precisa do arquivo sp_connector.py no repo)
from sp_connector import SPConnector


from auth_microsoft import (
    AuthManager,
    MicrosoftAuth,
    create_login_page,
    create_user_header,
)



# === Configurações via secrets (SharePoint - site servicosclinicos) ===
TENANT_ID = st.secrets["graph"]["tenant_id"]
CLIENT_ID = st.secrets["graph"]["client_id"]
CLIENT_SECRET = st.secrets["graph"]["client_secret"]
HOSTNAME = st.secrets["graph"]["hostname"]           
SITE_PATH = st.secrets["graph"]["site_path"]        
LIBRARY   = st.secrets["graph"]["library_name"]     

APONTAMENTOS  = st.secrets["files"]["apontamentos"]
ESTUDOS_CSV   = st.secrets["files"]["estudos_csv"]
COLABORADORES = st.secrets["files"]["colaboradores"]  # 'SANDRA/PROJETO_DASHBOARD/base_cargo.xlsx'

# Configurações de email para alertas
EMAIL_SMTP_SERVER = st.secrets.get("email", {}).get("smtp_server", "smtp.office365.com")
EMAIL_SMTP_PORT = st.secrets.get("email", {}).get("smtp_port", 587)
EMAIL_SENDER = st.secrets.get("email", {}).get("sender", "")
EMAIL_PASSWORD = st.secrets.get("email", {}).get("password", "")
EMAIL_ALERTS = ["susanna.bernardes@synvia.com", "washington.gouvea@synvia.com"]

# Instância única do conector (cacheada)
@st.cache_resource
def _sp():
    return SPConnector(
        TENANT_ID, CLIENT_ID, CLIENT_SECRET,
        hostname=HOSTNAME, site_path=SITE_PATH, library_name=LIBRARY
    )

# ============================================================
# SISTEMA DE LOGGING E MONITORAMENTO
# ============================================================


def enviar_email_alerta(assunto: str, corpo: str, anexos: list = None):
    """Envia email de alerta para os responsáveis"""
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        st.warning("Configurações de email não definidas nos secrets. Email não enviado.")
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = ", ".join(EMAIL_ALERTS)
        msg['Subject'] = assunto

        msg.attach(MIMEText(corpo, 'html'))

        # Adiciona anexos se houver
        if anexos:
            for nome_arquivo, dados in anexos:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(dados)
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={nome_arquivo}')
                msg.attach(part)

        # Envia o email
        with smtplib.SMTP(EMAIL_SMTP_SERVER, EMAIL_SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.send_message(msg)

        return True
    except Exception as e:
        st.error(f"Erro ao enviar email: {e}")
        return False

def verificar_diminuicao_dados(df_novo: pd.DataFrame, sheet_name: str = "apontamentos") -> tuple[bool, int, int]:
    """
    Verifica se o número de linhas diminuiu em relação ao último arquivo salvo
    Retorna: (diminuiu: bool, linhas_antes: int, linhas_depois: int)
    """
    try:
        # Carrega o arquivo atual do SharePoint
        data = _sp().download(APONTAMENTOS)
        xls = pd.ExcelFile(io.BytesIO(data))

        if sheet_name in xls.sheet_names:
            df_atual = pd.read_excel(xls, sheet_name=sheet_name)
            linhas_antes = len(df_atual)
            linhas_depois = len(df_novo)

            return (linhas_depois < linhas_antes, linhas_antes, linhas_depois)
    except:
        pass

    return (False, 0, len(df_novo))

def criar_backup_dataframe(df: pd.DataFrame) -> bytes:
    """Cria um backup do DataFrame em formato Excel"""
    output = io.BytesIO()
    df.to_excel(output, index=False, sheet_name="backup")
    output.seek(0)
    return output.getvalue()

# Função para ler o arquivo Excel (Apontamentos) do SharePoint com cache
@st.cache_data
def get_sharepoint_file(sheet_name: str = "apontamentos"):
    """
    Lê o arquivo Excel do SharePoint que contém múltiplas sheets:
    - 'apontamentos': dados principais
    - 'log': histórico de operações
    """
    try:
        data = _sp().download(APONTAMENTOS)
        xls = pd.ExcelFile(io.BytesIO(data))

        if sheet_name in xls.sheet_names:
            return pd.read_excel(xls, sheet_name=sheet_name)
        else:
            # Se a sheet não existir, retorna DataFrame vazio
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo no SharePoint (Graph): {e}")
        return pd.DataFrame()

# Função para ler o arquivo CSV (Estudos) do SharePoint com cache
@st.cache_data
def get_sharepoint_file_estudos_csv():
    try:
        return _sp().read_csv(ESTUDOS_CSV)
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo CSV de estudos no SharePoint (Graph): {e}")
        return pd.DataFrame()

@st.cache_data
def colaboradores_excel():
    try:
        data = _sp().download(COLABORADORES)
        xls = pd.ExcelFile(io.BytesIO(data))
        colaboradores_df = pd.read_excel(xls, sheet_name="Colaboradores")
        return colaboradores_df
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo ou ler as planilhas no SharePoint (Graph): {e}")
        return pd.DataFrame()
    

def generate_custom_id(existing_ids: set[str]) -> str:
    while True:
        digits = random.choices(string.digits, k=3)
        letters = random.choices(string.ascii_uppercase, k=2)
        chars = digits + letters
        random.shuffle(chars)
        new_id = "".join(chars)
        if new_id not in existing_ids:
            return new_id
        

# Função para atualizar o arquivo Excel (Apontamentos) no SharePoint
def update_sharepoint_file(df: pd.DataFrame, usuario: str = "", operacao: str = "ATUALIZAÇÃO", responsavel_indicado: str = "") -> pd.DataFrame | None:
    """
    Atualiza o arquivo Excel no SharePoint de forma segura, com logging e monitoramento.

    Estratégia:
    1. Carrega a versão mais recente do arquivo (ambas sheets: apontamentos e log)
    2. Para linhas existentes: atualiza APENAS as colunas que foram modificadas
    3. Para linhas novas: adiciona ao final
    4. Registra operação no log
    5. Verifica diminuição de dados e dispara alerta se necessário
    6. Salva arquivo com ambas as sheets
    7. Tenta novamente em caso de conflito de versão
    """
    attempts = 0
    while True:
        try:
            # Carrega versão mais recente do arquivo
            data = _sp().download(APONTAMENTOS)
            xls = pd.ExcelFile(io.BytesIO(data))

            # Carrega sheet de apontamentos
            if "apontamentos" in xls.sheet_names:
                base_df = pd.read_excel(xls, sheet_name="apontamentos")
            else:
                base_df = pd.DataFrame()

            # Carrega sheet de log (ou cria vazio)
            if "log" in xls.sheet_names:
                log_df = pd.read_excel(xls, sheet_name="log")
                # Adiciona coluna "Responsável Indicado" se não existir
                if "Responsável Indicado" not in log_df.columns:
                    log_df["Responsável Indicado"] = ""
            else:
                log_df = pd.DataFrame(columns=["Data", "ID", "Estudo", "Operação", "Campo", "Valor Anterior", "Valor Depois", "Responsável", "Responsável Indicado"])

            df_to_save = df.copy()
            if "ID" not in df_to_save.columns:
                st.error("DataFrame sem coluna ID!")
                return None

            df_to_save["ID"] = df_to_save["ID"].astype(str)

            # Variáveis para logging
            ids_novos = []
            ids_atualizados = []

            if not base_df.empty:
                base_df["ID"] = base_df["ID"].astype(str)

                # Separa registros novos dos existentes
                ids_to_save = set(df_to_save["ID"].tolist())
                existing_ids = set(base_df["ID"].tolist())

                new_ids = ids_to_save - existing_ids
                update_ids = ids_to_save & existing_ids

                # Adiciona registros completamente novos
                if new_ids:
                    new_rows = df_to_save[df_to_save["ID"].isin(new_ids)]
                    base_df = pd.concat([base_df, new_rows], ignore_index=True)
                    ids_novos = list(new_ids)

                # Atualiza registros existentes coluna por coluna
                for id_val in update_ids:
                    idx_base = base_df.index[base_df["ID"] == id_val].tolist()
                    idx_update = df_to_save.index[df_to_save["ID"] == id_val].tolist()

                    if idx_base and idx_update:
                        idx_b = idx_base[0]
                        idx_u = idx_update[0]

                        # Atualiza apenas as colunas que existem em ambos
                        for col in df_to_save.columns:
                            if col in base_df.columns:
                                base_df.at[idx_b, col] = df_to_save.at[idx_u, col]
                        ids_atualizados.append(id_val)
            else:
                # Se o arquivo está vazio, salva tudo
                base_df = df_to_save.copy()
                ids_novos = df_to_save["ID"].tolist()

            # === ADICIONA ENTRADAS NO LOG ===
            # Para cada ID criado, adiciona uma entrada no log
            for id_val in ids_novos:
                # Pega o estudo do apontamento
                estudo = df_to_save.loc[df_to_save["ID"] == id_val, "Código do Estudo"].iloc[0] if "Código do Estudo" in df_to_save.columns else ""
                status_novo = df_to_save.loc[df_to_save["ID"] == id_val, "Status"].iloc[0] if "Status" in df_to_save.columns else "PENDENTE"

                # Pega o responsável pela correção do apontamento
                resp_correcao = df_to_save.loc[df_to_save["ID"] == id_val, "Responsável Pela Correção"].iloc[0] if "Responsável Pela Correção" in df_to_save.columns else ""

                nova_log_entry = pd.DataFrame([{
                    "Data": datetime.now(),
                    "ID": id_val,
                    "Estudo": estudo,
                    "Operação": "edição",
                    "Campo": "Status",
                    "Valor Anterior": "",
                    "Valor Depois": status_novo,
                    "Responsável": usuario if usuario else "Sistema",
                    "Responsável Indicado": responsavel_indicado if responsavel_indicado else resp_correcao
                }])
                log_df = pd.concat([log_df, nova_log_entry], ignore_index=True)

            # Para cada ID atualizado, adiciona uma entrada no log
            for id_val in ids_atualizados:
                # Pega valores antigos e novos
                valor_anterior = base_df.loc[base_df["ID"] == id_val, "Status"].iloc[0] if "Status" in base_df.columns else ""
                valor_depois = df_to_save.loc[df_to_save["ID"] == id_val, "Status"].iloc[0] if "Status" in df_to_save.columns else ""
                estudo = df_to_save.loc[df_to_save["ID"] == id_val, "Código do Estudo"].iloc[0] if "Código do Estudo" in df_to_save.columns else ""

                # Só registra se houver mudança
                if valor_anterior != valor_depois:
                    # Pega o responsável pela correção do apontamento
                    resp_correcao = df_to_save.loc[df_to_save["ID"] == id_val, "Responsável Pela Correção"].iloc[0] if "Responsável Pela Correção" in df_to_save.columns else ""

                    nova_log_entry = pd.DataFrame([{
                        "Data": datetime.now(),
                        "ID": id_val,
                        "Estudo": estudo,
                        "Operação": "edição",
                        "Campo": "Status",
                        "Valor Anterior": valor_anterior,
                        "Valor Depois": valor_depois,
                        "Responsável": usuario if usuario else "Sistema",
                        "Responsável Indicado": responsavel_indicado if responsavel_indicado else resp_correcao
                    }])
                    log_df = pd.concat([log_df, nova_log_entry], ignore_index=True)

            # === VERIFICA DIMINUIÇÃO DE DADOS ===
            diminuiu, linhas_antes, linhas_depois = verificar_diminuicao_dados(base_df, "apontamentos")

            if diminuiu:
                # Cria backup do DataFrame anterior
                try:
                    df_anterior = pd.read_excel(xls, sheet_name="apontamentos")
                    backup_bytes = criar_backup_dataframe(df_anterior)

                    # Pega últimos 10 logs
                    ultimos_logs = log_df.tail(10).to_html(index=False)

                    # Monta email de alerta
                    assunto = f"⚠️ ALERTA: Diminuição de Dados Detectada - Sistema de Apontamentos"
                    corpo = f"""
                    <html>
                    <body>
                        <h2 style="color: #d9534f;">Alerta de Diminuição de Dados</h2>
                        <p>O sistema detectou uma diminuição no número de registros do arquivo de apontamentos.</p>

                        <h3>Detalhes:</h3>
                        <ul>
                            <li><strong>Linhas antes:</strong> {linhas_antes}</li>
                            <li><strong>Linhas depois:</strong> {linhas_depois}</li>
                            <li><strong>Diferença:</strong> {linhas_antes - linhas_depois} linhas removidas</li>
                            <li><strong>Data/Hora:</strong> {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</li>
                            <li><strong>Usuário:</strong> {usuario if usuario else "Sistema"}</li>
                            <li><strong>Operação:</strong> {operacao}</li>
                        </ul>

                        <h3>Últimos Logs:</h3>
                        {ultimos_logs}

                        <p style="margin-top: 20px;">
                            <strong>ATENÇÃO:</strong> Um backup do DataFrame anterior está anexado a este email.
                        </p>

                        <p style="color: #666; font-size: 12px; margin-top: 30px;">
                            Este é um email automático do Sistema de Apontamentos.<br>
                            Em caso de dúvidas, entre em contato com a equipe de TI.
                        </p>
                    </body>
                    </html>
                    """

                    anexos = [
                        (f"backup_apontamentos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", backup_bytes)
                    ]

                    # Adiciona log de alerta enviado
                    log_alerta = pd.DataFrame([{
                        "Data": datetime.now(),
                        "ID": "",
                        "Estudo": "",
                        "Operação": "ALERTA_EMAIL",
                        "Campo": "Sistema",
                        "Valor Anterior": str(linhas_antes),
                        "Valor Depois": str(linhas_depois),
                        "Responsável": "Sistema",
                        "Responsável Indicado": ""
                    }])
                    log_df = pd.concat([log_alerta, log_df], ignore_index=True)

                    # Envia email (em background, não bloqueia)
                    enviar_email_alerta(assunto, corpo, anexos)

                except Exception as e_email:
                    # Log de erro no envio de email
                    log_erro = pd.DataFrame([{
                        "Data": datetime.now(),
                        "ID": "",
                        "Estudo": "",
                        "Operação": "ALERTA_EMAIL",
                        "Campo": "Sistema",
                        "Valor Anterior": "",
                        "Valor Depois": "",
                        "Responsável": "Sistema",
                        "Responsável Indicado": ""
                    }])
                    log_df = pd.concat([log_erro, log_df], ignore_index=True)

            # === SALVA O ARQUIVO COM MÚLTIPLAS SHEETS ===
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                base_df.to_excel(writer, sheet_name='apontamentos', index=False)
                log_df.to_excel(writer, sheet_name='log', index=False)
            output.seek(0)

            _sp().upload_small(APONTAMENTOS, output.getvalue(), overwrite=True)

            st.success("Mudanças submetidas com sucesso! Recarregue a página para ver as mudanças")
            return base_df

        except Exception as e:
            attempts += 1
            msg = str(e)

            # Adiciona log de erro
            try:
                log_erro = pd.DataFrame([{
                    "Data": datetime.now(),
                    "ID": "",
                    "Estudo": "",
                    "Operação": operacao,
                    "Campo": "Sistema",
                    "Valor Anterior": "",
                    "Valor Depois": "",
                    "Responsável": usuario if usuario else "Sistema",
                    "Responsável Indicado": ""
                }])
                if 'log_df' in locals():
                    log_df = pd.concat([log_erro, log_df], ignore_index=True)
            except:
                pass

            # 409/412 = conflito de versão | 429 = throttling
            if any(x in msg for x in ["409", "412", "429"]) and attempts < 5:
                st.warning("Outra pessoa está salvando ou limite de chamadas. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue

            st.error(f"Erro ao salvar no SharePoint (Graph): {msg}")
            return None


# -------------------------------------------------
# Autenticação e contexto do usuário
# -------------------------------------------------
auth = MicrosoftAuth()
logged_in = create_login_page(auth)
if not logged_in:
    st.stop()

# Garantir token válido durante a sessão
AuthManager.check_and_refresh_token(auth)
create_user_header()

user = AuthManager.get_current_user() or {}
display_name = user.get("displayName", "Usuário")
user_email = (user.get("mail") or user.get("userPrincipalName") or "").lower()


st.session_state["display_name"] = display_name
st.session_state["user_email"] = user_email


# Carregar dados iniciais
with st.spinner("Carregando dados do SharePoint..."):
    df_study = get_sharepoint_file_estudos_csv()
    colaboradores_df = colaboradores_excel()


# Inicializar o DataFrame de apontamentos no session_state
if "df_apontamentos" not in st.session_state:
    with st.spinner("Carregando apontamentos..."):
        df_loaded = get_sharepoint_file()
    
    # Fill missing or invalid IDs to prevent NaN issues
    if not df_loaded.empty:
        if "ID" not in df_loaded.columns:
            dexisting = set()
            df_loaded["ID"] = [generate_custom_id(existing) for _ in range(len(df_loaded))]
        else:
            df_loaded["ID"] = df_loaded["ID"].astype(str)
            existing = set(df_loaded["ID"])
            mask = df_loaded["ID"].str.lower().isin(["nan", "none", "", "nat"])
            for idx in df_loaded.index[mask]:
                new_id = generate_custom_id(existing)
                df_loaded.at[idx, "ID"] = new_id
                existing.add(new_id)
    
    st.session_state["df_apontamentos"] = df_loaded

    # Gerando o ID do apontamento atual
    existing_ids = set(df_loaded["ID"].astype(str)) if not df_loaded.empty else set()
    st.session_state["generated_id"] = generate_custom_id(existing_ids)

# Configurar session_state para campos condicionais
if "status" not in st.session_state:
    st.session_state["status"] = ""
if "enable_data_resolucao" not in st.session_state:
    st.session_state["enable_data_resolucao"] = False
if "enable_nao_aplicavel" not in st.session_state:
    st.session_state["enable_nao_aplicavel"] = False



def update_status_fields():
    s = st.session_state["status"]

    if s == "VERIFICANDO":
        st.info("Esse staus só pode ser preenchido pelo Guilherme Goncalves")

    elif s == "REALIZADO": 
        st.session_state["enable_data_resolucao"] = True
        st.session_state["enable_nao_aplicavel"] = False
    
    elif s == "NÃO APLICÁVEL":
        st.session_state["enable_data_resolucao"] = False
        st.session_state["enable_nao_aplicavel"] = True

    else:                                       # PENDENTE, REALIZADO DURANTE A CONDUÇÃO …
        st.session_state["enable_data_resolucao"] = False
        st.session_state["enable_nao_aplicavel"] = False

def pegar_dados_colab(nome_colab: str, df: pd.DataFrame, campos: list[str]):
    """
    Retorna os dados solicitados de um colaborador, baseado nos nomes dos campos.

    Parâmetros:
        nome_colab (str): Nome do colaborador.
        df (pd.DataFrame): DataFrame contendo os dados.
        campos (list[str]): Lista de nomes de colunas a serem retornadas.

    Retorna:
        tuple: Valores dos campos solicitados, na ordem da lista `campos`.
    """
    linha = df.loc[df["Nome Completo do Profissional"] == nome_colab]
    if linha.empty:
        return tuple("" for _ in campos)
    
    lin = linha.iloc[0]
    return tuple(lin[campo] if campo in lin else "" for campo in campos)


# Início da tela principal
tab_names = ["Formulário", "Lista de Apontamentos"]
if "active_tab" not in st.session_state:
    st.session_state.active_tab = tab_names[0]

tab_option = st.radio(
    label="",  
    options=tab_names,
    horizontal=True,
    key="active_tab",
)

if tab_option == "Formulário":
    st.title("Criar Apontamento")
    
    if df_study.empty:
        st.error("Arquivo CSV de estudos não carregado. Verifique o caminho do arquivo.")
    else:

        if "generated_id" not in st.session_state:
            df_ids = st.session_state.get("df_apontamentos", pd.DataFrame())
            existing = set(df_ids["ID"].astype(str)) if not df_ids.empty else set()
            st.session_state["generated_id"] = generate_custom_id(existing)

        st.text_input("ID do Apontamento", value=st.session_state["generated_id"], disabled=True)
        protocol_options = ["Digite o codigo do estudo"] + df_study["NUMERO_DO_PROTOCOLO"].tolist()
        selected_protocol = st.selectbox("Código do Estudo", options=protocol_options, key="selected_protocol")
        
        if selected_protocol != "Digite o codigo do estudo":
            research_name = df_study.loc[df_study["NUMERO_DO_PROTOCOLO"] == selected_protocol, "NOME_DA_PESQUISA"].iloc[0]
        else:
            research_name = ""
        st.text_input("Nome da Pesquisa", value=research_name, disabled=True)
        
        
        origem = st.selectbox(
            "Origem Do Apontamento", 
            ["Documentação Clínica", "Excelência Operacional", "Operações Clínicas", 
             "Patrocinador / Monitor", "Garantia Da Qualidade"], 
            key="origem"
        )
        
        # Selectbox para documentos com opção "Outros"
        doc = st.selectbox("Documentos", [
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
            "FOR 236", "FOR 240", "FOR 241", "FOR 367", "Outros"
        ], key="documento")

        
        
        # Se o usuário selecionar "Outros", exibe um input extra para informar o documento
        if st.session_state["documento"] == "Outros":
            st.text_input("Indique o documento", key="doc_custom")
        
        
        
        # Função que retorna o valor final do documento
        def get_final_documento():
            doc_value = st.session_state.get("documento", "")
            if doc_value == "Outros":
                return st.session_state.get("doc_custom", "").strip()
            return doc_value
        
        
        
        # Obtém o valor final do documento usando a função
        documento_final = get_final_documento()

        pp_options = ["N/A", "Outros"] + [f"PP{i:02d}" for i in range(1, 100)] + [f"PP{i}" for i in range(100, 1000)]

        participante = st.selectbox("Participante", pp_options, key="participante")
        


        
        if st.session_state["participante"] == "Outros":
            st.text_input("Indique os PPs", key="pp_custom", placeholder='Neste formato: PP01, PP02')

        def get_final_pp():
            pp_value = st.session_state.get("participante", "")
            if pp_value == "Outros":
                return st.session_state.get("pp_custom", "").strip()
            return pp_value
        
        pp_final = get_final_pp()
             
        periodo = st.selectbox("Período", ["N/A", "Pós",
            '1° Período', '2° Período', '3° Período',
            '4° Período', '5° Período', '6° Período', '7° Período', 
            '8° Período', '9° Período', '10° Período'
        ], key="periodo")
        

        prazo = st.date_input("Prazo Para Resolução", format="DD/MM/YYYY", key="prazo")
        apontamento = st.text_area("Apontamento", key="apontamento")

        
        responsavel_options = ["Selecione um colaborador"] + colaboradores_df["Nome Completo do Profissional"].tolist()
        correcao = st.selectbox("Responsável pela Correção", options=responsavel_options, key="responsavel")

        plantao, status_prof, departamento = pegar_dados_colab(correcao, colaboradores_df, ["Plantão", "Tempo De Casa","Departamento"])



        # Campo de Status com callback (supondo que a função update_status_fields esteja definida)
        opts = ["PENDENTE","REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "NÃO APLICÁVEL"]
        key = "status"

        def _norm(x):
            if x is None: return None
            s = str(x).strip()
            return s if s else None  # trata "" como None

        cur = _norm(st.session_state.get(key))

        # se o valor atual não é uma opção válida, remove do session_state
        if (cur is None) or (cur not in opts):
            st.session_state.pop(key, None)



        status = st.selectbox(
            "Status",
            opts,
            key=key,
            on_change=update_status_fields,
            index=None,
            placeholder="Selecione um Status"
        )
        

        if st.session_state["enable_nao_aplicavel"]:
            justificativa = st.text_input("Justificativa", key="justificativa")
            resolucao = st.date_input("Data da resolução", format="DD/MM/YYYY")
            verificador_nome = ""
            verificador_data = None
        elif st.session_state["enable_data_resolucao"]:
            resolucao = st.date_input("Data da resolução", format="DD/MM/YYYY")
            justificativa = "N/A"


        else:
            verificador_nome = ""
            verificador_data = None
            justificativa = "N/A"
            resolucao = None

        submit = st.button("Enviar")

        if submit:
            # Validação dos campos obrigatórios
            if selected_protocol == "Digite o codigo do estudo" or participante.strip() == "" or apontamento.strip() == "":
                st.error("Por favor, preencha os campos obrigatórios: Código do Estudo, Participante, Responsável e Apontamento.")
            elif status == "VERIFICANDO" and verificador_nome.strip() == "":
                st.error("Somente o Guilherme Gonçalves pode usar esse status!.")
            elif  status == "Selecione um Status":
                st.error("Por favor, defina um status antes de submeter o apontamento!")
            elif status == "NÃO APLICÁVEL" and justificativa.strip() == "":
                st.error("Por favor, preencha o campo 'Justificativa'!")
                st.stop()
            elif correcao == "Selecione um colaborador":
                st.warning("Por favor, selecione o colaborador responsável pela correção antes de salvar.")
                st.stop()
            else:
                with st.spinner("Salvando apontamento..."):
                    data_atual = datetime.now()

                    if st.session_state["status"] == "REALIZADO DURANTE A CONDUÇÃO":
                        resolucao = data_atual
                    
                    df = st.session_state["df_apontamentos"]
    
                    # Usa o ID gerado previamente para este apontamento
                    next_id = st.session_state.get("generated_id")
    
                    responsavel_nome = st.session_state.get("display_name")
                    
                    
    
                    novo_apontamento = {
                        "ID": next_id,
                        "Código do Estudo": selected_protocol,
                        "Nome da Pesquisa": research_name,
                        "Data do Apontamento": data_atual,
                        "Responsável Pelo Apontamento": responsavel_nome,
                        "Origem Do Apontamento": st.session_state["origem"],
                        "Documentos": documento_final,  # Aqui utiliza o valor final (customizado se "Outros")
                        "Participante": pp_final,
                        "Período": st.session_state["periodo"],
                        "Prazo Para Resolução": st.session_state["prazo"],
                        "Apontamento": st.session_state["apontamento"],
                        "Status": st.session_state["status"],
                        "Verificador": st.session_state.get("verificador_nome", ""),
                        "Disponibilizado para Verificação": st.session_state.get("verificador_data", None),
                        "Justificativa": st.session_state.get("justificativa", ""),
                        "Responsável Pela Correção": correcao,
                        "Data Resolução": resolucao,
                        "Plantão": plantao,
                        "Departamento": departamento,
                        "Tempo de casa": status_prof,
                        # Colunas de controle (preenchidas posteriormente)
                        "Responsável Indicado": "",
                        "Grau De Criticidade Do Apontamento": "",
                        "Responsável Atualização": "",
                        "Data Atualização": None,
                        "Data Início Verificação": None
                    }
                    
    
    
                    novo_df = pd.DataFrame([novo_apontamento])
                    df_atualizado = update_sharepoint_file(
                        novo_df,
                        usuario=display_name,
                        operacao="CRIAR_APONTAMENTO"
                    )
                    if df_atualizado is not None:
                        st.session_state["df_apontamentos"] = df_atualizado
                        st.session_state["generated_id"] = generate_custom_id(
                            set(df_atualizado["ID"].astype(str))
                        )
                



if tab_option == "Lista de Apontamentos":
    # ─────────────────────────────────────────────────────────────
    # 1️⃣  Garante índice interno e coluna visível de ID
    # ─────────────────────────────────────────────────────────────
    df = st.session_state["df_apontamentos"]

    # Cria coluna/índice inicial na primeira execução
    if "orig_idx" not in df.columns:
        df.insert(0, "orig_idx", range(len(df)))  # índice técnico permanente
        df.set_index("orig_idx", inplace=True)

    # Cria a coluna ID visível caso não exista
    if "ID" not in df.columns:
        existing = set()
        df["ID"] = [generate_custom_id(existing) for _ in range(len(df))]

    # ─────────────────────────────────────────────────────────────
    # 2️⃣  Estado da interface
    # ─────────────────────────────────────────────────────────────
    st.session_state.setdefault("mostrar_campos_finais", False)
    st.session_state.setdefault("indices_alterados", [])


    st.title("Lista de Apontamentos")

    col_btn1, *_ = st.columns(6)
    with col_btn1:
        if st.button("🔄 Atualizar"):
            st.cache_data.clear()      
            st.cache_resource.clear()
            st.rerun()  

    # ─────────────────────────────────────────────────────────────
    # 3️⃣  Filtros rápidos / seletor de estudo
    # ─────────────────────────────────────────────────────────────
    if df.empty:
        st.info("Nenhum apontamento encontrado!")
        st.stop()

    df_filtrado = df.copy()

    st.markdown("")


        # 🔎 Filtro por ID (linha inteira)
    id_busca = st.text_input(
        "Buscar por ID",
        placeholder="Digite o ID",
    )

    if id_busca:
        df_filtrado = df_filtrado[
            df_filtrado["ID"].astype(str).str.contains(id_busca, case=False, na=False)
        ]

    # Linha com 2 colunas: Estudo (esquerda) e Status (direita)
    col_filtro_estudo, col_filtro_status = st.columns(2)

        # Linha com 2 colunas: Estudo (esquerda) e Status (direita)
    col_filtro_estudo, col_filtro_status = st.columns(2)

    with col_filtro_estudo:
        opcoes_estudos = ["Todos"] + sorted(
            df["Código do Estudo"].dropna().unique().tolist()
        )
        estudo_sel = st.selectbox("Selecione o Estudo", options=opcoes_estudos)

    with col_filtro_status:
        opcoes_status = ["Todos"] + sorted(
            df["Status"].dropna().unique().tolist()
        )
        status_sel = st.selectbox("Filtrar por Status", options=opcoes_status)

    # Aplica filtros
    if estudo_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Código do Estudo"] == estudo_sel]

    if status_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Status"] == status_sel]


    # Colunas visíveis (ID primeiro)
    cols_display = [
        "ID", "Status", "Código do Estudo", "Responsável Pela Correção", "Plantão",
        "Participante", "Período", "Documentos", "Apontamento",
        "Prazo Para Resolução", "Data Resolução", "Justificativa",
        "Responsável Pelo Apontamento", "Origem Do Apontamento",
    ]
    df_filtrado = df_filtrado[cols_display]

    # Converte colunas de data
    colunas_data = ["Data do Apontamento", "Prazo Para Resolução", "Data Resolução"]
    for col in colunas_data:
        if col in df_filtrado.columns:
            df_filtrado[col] = pd.to_datetime(df_filtrado[col], errors="coerce")

    # ─────────────────────────────────────────────────────────────
    # 4️⃣  Config do editor (ID bloqueado, Status editável)
    # ─────────────────────────────────────────────────────────────
    columns_config = {
        "ID": st.column_config.TextColumn("ID", disabled=True)
    }

    for col in df_filtrado.columns:
        if col == "Status":
            columns_config[col] = st.column_config.SelectboxColumn(
                "Status",
                options=[
                    "REALIZADO DURANTE A CONDUÇÃO", "REALIZADO",
                    "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"
                ],
                disabled=False,
            )
        elif col in colunas_data:
            columns_config[col] = st.column_config.DateColumn(col, disabled=True, format="DD/MM/YYYY")
        elif col != "ID":
            columns_config[col] = st.column_config.TextColumn(col, disabled=True)

    df_editado = st.data_editor(
        df_filtrado,
        column_config=columns_config,
        num_rows="fixed",
        key="data_editor",
        hide_index=True,  # esconde orig_idx e numeração lateral
    )

    # ─────────────────────────────────────────────────────────────
    # 5️⃣  Detecta alterações de Status usando a coluna ID
    # ─────────────────────────────────────────────────────────────
    if not st.session_state.mostrar_campos_finais:
        if st.button("Status modificados"):
            alterado = False
            indices_alterados = []
            df_atualizado = df.copy()

            for i in range(len(df_filtrado)):
                status_original = df_filtrado.iloc[i]["Status"]
                status_novo = df_editado.iloc[i]["Status"]

                if status_novo != status_original:
                    alterado = True
                    id_val = df_filtrado.iloc[i]["ID"]    # pega o ID visível

                    # Atualiza no DataFrame base usando a coluna ID
                    df.loc[df["ID"] == id_val, "Status"] = status_novo
                    indices_alterados.append(id_val)

            if not alterado:
                st.warning("Nenhuma alteração de status detectada.")
            else:
                st.session_state.mostrar_campos_finais = True
                st.session_state.indices_alterados = indices_alterados


    # ─────────────────────────────────────────────────────────────
    # 6️⃣  Campos finais obrigatórios + submissão
    # ─────────────────────────────────────────────────────────────
    if st.session_state.mostrar_campos_finais:
        df = st.session_state["df_apontamentos"]
        indices_alterados = st.session_state.indices_alterados
        linhas_faltando = []

        st.markdown("### Preencha os campos obrigatórios")

        for id_val in indices_alterados:
            status_novo = df.loc[df["ID"] == id_val, "Status"].iloc[0]
            st.markdown(f"#### Apontamento ID {id_val}")

            if status_novo in ["REALIZADO", "NÃO APLICÁVEL"]:
                key_data = f"data_conclusao_{id_val}"
                data_concl = st.date_input("Data de Resolução", key=key_data, format="DD/MM/YYYY")
                if not data_concl:
                    linhas_faltando.append(f"[ID {id_val}] Data de Resolução")
                else:
                    df.loc[df["ID"] == id_val, "Data Resolução"] = data_concl

            if status_novo == "NÃO APLICÁVEL":
                key_just = f"justificativa_{id_val}"
                justificativa = st.text_area("Justificativa obrigatória:", key=key_just)
                if not justificativa.strip():
                    linhas_faltando.append(f"[ID {id_val}] Justificativa")
                else:
                    df.loc[df["ID"] == id_val, "Justificativa"] = justificativa

            st.markdown("---")

        # Responsável pela atualização
        colaboradores_eo = colaboradores_df[colaboradores_df["Departamento"] == "Excelência Operacional"]
        resp_opts = ["Selecione um Colaborador"] + colaboradores_eo["Nome Completo do Profissional"].tolist()
        responsavel = st.selectbox("Responsável pela Atualização", options=resp_opts, key="responsavel_final")

        if st.button("Submeter mudanças"):
            if linhas_faltando:
                st.error("Campos obrigatórios pendentes:\n\n" + "\n".join(linhas_faltando))
            elif responsavel == "Selecione um Colaborador":
                st.warning("Por favor, selecione um responsável!")
            else:
                with st.spinner("Salvando mudanças..."):
                    df.loc[df["ID"].isin(indices_alterados), "Verificador"] = responsavel

                    # Salva de volta
                    rows_to_save = df[df["ID"].isin(indices_alterados)]
                    df_atualizado = update_sharepoint_file(
                        rows_to_save,
                        usuario=responsavel,
                        operacao="ATUALIZAR_STATUS",
                        responsavel_indicado=responsavel
                    )
                    if df_atualizado is not None:
                        st.session_state["df_apontamentos"] = df_atualizado

                    # Limpa estados
                    st.session_state.mostrar_campos_finais = False
                    st.session_state.indices_alterados = []
                    st.session_state.df_atualizado = None
