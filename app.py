import streamlit as st
import pandas as pd
from datetime import datetime
import io
import time
import random
import string

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


# Instância única do conector (cacheada)
@st.cache_resource
def _sp():
    return SPConnector(
        TENANT_ID, CLIENT_ID, CLIENT_SECRET,
        hostname=HOSTNAME, site_path=SITE_PATH, library_name=LIBRARY
    )



LOG_SHEET = "logs"
APONT_SHEET = "apontamentos"

LOG_COLUMNS = ["Data", "ID", "Estudo","Operação", "Campo", "Valor Antes", "Valor Depois", "Responsável Indicado", "Responsável"]

@st.cache_data(show_spinner=False)
def get_sharepoint_workbook():
    """
    Lê as duas abas e devolve (df_apontamentos, df_logs).
    """
    try:
        x = _sp().read_excel(APONTAMENTOS, sheet_name=[APONT_SHEET, LOG_SHEET])

        df_ap = x.get(APONT_SHEET, pd.DataFrame())
        df_lg = x.get(LOG_SHEET, pd.DataFrame())

        # garante colunas do log
        if df_lg.empty:
            df_lg = pd.DataFrame(columns=LOG_COLUMNS)
        else:
            for c in LOG_COLUMNS:
                if c not in df_lg.columns:
                    df_lg[c] = None
            df_lg = df_lg[LOG_COLUMNS]

        return df_ap, df_lg

    except Exception as e:
        st.error(f"Erro ao ler workbook (apontamentos/logs): {e}")
        return pd.DataFrame(), pd.DataFrame(columns=LOG_COLUMNS)


def build_log_rows(
    *,
    id_apontamento: str,
    estudo: str,
    operacao: str,
    campo: str,
    valor_antes,
    valor_depois,
    responsavel_nome: str,
    responsavel_indicado: str,
    when: datetime | None = None
) -> dict:
    when = when or datetime.now()
    return {
        "Data": when.strftime("%Y-%m-%d %H:%M:%S"),
        "ID": str(id_apontamento),
        "Estudo": estudo,
        "Operação": operacao,
        "Campo": campo,
        "Valor Antes": "" if valor_antes is None else str(valor_antes),
        "Valor Depois": "" if valor_depois is None else str(valor_depois),
        "Responsável": responsavel_nome,
        "Responsável Indicado": responsavel_indicado,
    }


def update_sharepoint_workbook(
    *,
    apontamentos_delta: pd.DataFrame | None = None,   # linhas para update/append em apontamentos
    logs_append: pd.DataFrame | None = None           # linhas NOVAS para append em logs
) -> tuple[pd.DataFrame, pd.DataFrame] | None:
    """
    Atualiza a aba apontamentos (update + append por ID) e faz append na aba logs.
    Depois escreve o workbook inteiro (duas abas) e sobe no SharePoint.
    """
    attempts = 0
    while True:
        try:
            df_ap_base, df_logs_base = get_sharepoint_workbook()

            # --- APONTAMENTOS: update/append por ID ---
            if df_ap_base.empty:
                df_ap_base = pd.DataFrame()

            if apontamentos_delta is not None and not apontamentos_delta.empty:
                ap = df_ap_base.copy()
                delta = apontamentos_delta.copy()

                # normaliza ID
                if "ID" in ap.columns:
                    ap["ID"] = ap["ID"].astype(str)
                if "ID" in delta.columns:
                    delta["ID"] = delta["ID"].astype(str)

                # usa ID como chave
                if "ID" not in ap.columns:
                    ap = ap.copy()
                    ap["ID"] = None

                ap = ap.set_index("ID", drop=False)
                delta = delta.set_index("ID", drop=False)

                ap.update(delta)  # atualiza colunas existentes
                novos = delta.index.difference(ap.index)
                if len(novos) > 0:
                    ap = pd.concat([ap, delta.loc[novos]], axis=0)

                df_ap_base = ap.reset_index(drop=True)

            # --- LOGS: append ---
            if logs_append is not None and not logs_append.empty:
                lg = df_logs_base.copy()
                add = logs_append.copy()

                # garante colunas e ordem
                for c in LOG_COLUMNS:
                    if c not in lg.columns:
                        lg[c] = None
                    if c not in add.columns:
                        add[c] = None
                lg = lg[LOG_COLUMNS]
                add = add[LOG_COLUMNS]

                df_logs_base = pd.concat([lg, add], ignore_index=True)

            # --- escreve workbook inteiro (2 abas) ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_ap_base.to_excel(writer, index=False, sheet_name=APONT_SHEET)
                df_logs_base.to_excel(writer, index=False, sheet_name=LOG_SHEET)

            output.seek(0)
            _sp().upload_small(APONTAMENTOS, output.getvalue(), overwrite=True)

            # limpa cache de leitura (porque mudou)
            st.cache_data.clear()

            return df_ap_base, df_logs_base

        except Exception as e:
            attempts += 1
            msg = str(e)

            # 1) arquivo bloqueado / aberto no Excel (mensagens variam)
            lock_signals = ["locked", "lock", "423", "resourceLocked", "is locked", "The resource you are attempting to access is locked"]
            conflict_signals = ["409", "412", "precondition", "etag", "conflict"]

            if any(s.lower() in msg.lower() for s in lock_signals):
                st.warning(
                    "⚠️ Não foi possível salvar porque o arquivo parece estar aberto no Excel (bloqueado para edição).\n\n"
                    "👉 Feche o arquivo no computador (e em qualquer outro lugar onde esteja aberto) e tente salvar novamente."
                )
                return None

            # 2) conflitos/transientes (mantém retry)
            if any(s.lower() in msg.lower() for s in conflict_signals) and attempts < 5:
                st.warning("Conflito detectado no SharePoint. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue

            if "429" in msg and attempts < 5:
                st.warning("Muitas requisições ao SharePoint. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue

            st.error(f"Erro ao salvar o apontamento: {msg}")
            return None



# Função para ler o arquivo CSV (Estudos) do SharePoint com cache
@st.cache_data(show_spinner=False)
def get_sharepoint_file_estudos_csv():
    try:
        return _sp().read_csv(ESTUDOS_CSV)
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo CSV de estudos no SharePoint (Graph): {e}")
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
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
        
# Função para carregar o workbook com spinner
def load_workbook_with_spinner(msg="📘 Carregando apontamentos..."):
    with st.spinner(msg):
        return get_sharepoint_workbook()


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
with st.spinner("📥 Carregando bases do SharePoint..."):
    df_study = get_sharepoint_file_estudos_csv()
    colaboradores_df = colaboradores_excel()



# Inicializar o DataFrame de apontamentos no session_state
if "df_apontamentos" not in st.session_state:
    df_loaded, df_logs_loaded = load_workbook_with_spinner(
    "📘 Carregando apontamentos e logs..."
)


    
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
        opts = ["Selecione um Status","PENDENTE","REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "NÃO APLICÁVEL"]
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
            on_change=update_status_fields
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
            elif correcao == "Selecione um colaborador":
                st.warning("Por favor, selecione o colaborador responsável pela correção antes de salvar.")
                st.stop()
            else:
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
                    "Responsável Indicado": None
                }
                


                novo_df = pd.DataFrame([novo_apontamento])
                # operador da ação (quem clicou/salvou)

                
                

                # log de criação (você disse que logs são só de status — aqui registrei o status inicial)
                log_row = build_log_rows(
                    id_apontamento=next_id,
                    estudo=selected_protocol,
                    operacao="criação",
                    campo="Status",
                    valor_antes="",
                    valor_depois=st.session_state["status"],
                    responsavel_nome=responsavel_nome,
                    when=data_atual,
                    responsavel_indicado=None,
                )
                df_logs_add = pd.DataFrame([log_row])

                with st.spinner("💾 Salvando apontamento no SharePoint..."):
                    res = update_sharepoint_workbook(apontamentos_delta=novo_df, logs_append=df_logs_add)

                if res is not None:
                    df_atualizado, df_logs_atualizado = res
                    st.session_state["df_apontamentos"] = df_atualizado
                    st.session_state["generated_id"] = generate_custom_id(set(df_atualizado["ID"].astype(str)))
                    st.session_state["df_logs"] = df_logs_atualizado  # <<< mantém logs atualizados localmente
                    st.success("✅ Apontamento submetido com sucesso!")



                



if tab_option == "Lista de Apontamentos":

    def hard_refresh():
        st.cache_data.clear()
        df_loaded, df_logs_loaded = load_workbook_with_spinner(
            "🔄 Atualizando dados do SharePoint..."
        )
        st.session_state["df_apontamentos"] = df_loaded
        st.session_state["df_logs"] = df_logs_loaded



    if st.session_state.get("_do_refresh"):
        st.session_state["_do_refresh"] = False
        hard_refresh()
        st.rerun()


    def limpar_filtros_e_refresh():
        st.session_state["filtro_id"] = ""
        st.session_state["filtro_estudo"] = "Todos"
        st.session_state["filtro_status"] = "Todos"

        # marca pra rodar refresh no próximo ciclo (evita mexer em cache + rerun dentro do clique)
        st.session_state["_do_refresh"] = True


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
        st.button("🔄 Atualizar", on_click=limpar_filtros_e_refresh)

    
    # --- reset de filtros (precisa rodar antes dos widgets existirem) ---
    if st.session_state.get("_reset_filtros", False):
        st.session_state["_reset_filtros"] = False
        st.session_state["filtro_id"] = ""
        st.session_state["filtro_estudo"] = "Todos"
        st.session_state["filtro_status"] = "Todos"

        hard_refresh()
        st.rerun()





    # ─────────────────────────────────────────────────────────────
    # 3️⃣  Filtros rápidos 
    # ─────────────────────────────────────────────────────────────
    if df.empty:
        st.info("Nenhum apontamento encontrado!")
        st.stop()

    df_filtrado = df.copy()

    campo_id = st.text

    st.markdown("")


        # 🔎 Filtro por ID (linha inteira)
    id_busca = st.text_input(
        "Buscar por ID",
        placeholder="Digite o ID",
        key="filtro_id",
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
        estudo_sel = st.selectbox("Selecione o Estudo", options=opcoes_estudos, key="filtro_estudo",)

    with col_filtro_status:
        opcoes_status = ["Todos"] + sorted(
            df["Status"].dropna().unique().tolist()
        )
        status_sel = st.selectbox("Filtrar por Status", options=opcoes_status,key="filtro_status")

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

    # key muda quando o filtro muda => evita DOM crash do editor antigo
    signature = (
        str(estudo_sel),
        str(status_sel),
        str(id_busca),
        len(df_filtrado),
        str(df_filtrado["ID"].iloc[0]) if len(df_filtrado) else "empty",
    )
    editor_key = "data_editor_" + "_".join(map(lambda x: str(x).replace(" ", ""), signature))


    df_editado = st.data_editor(
        df_filtrado,
        column_config=columns_config,
        num_rows="fixed",
        key=editor_key,
        hide_index=True,  # esconde orig_idx e numeração lateral
    )

    # ─────────────────────────────────────────────────────────────
    # 5️⃣  Detecta alterações de Status usando a coluna ID
    # ─────────────────────────────────────────────────────────────
    if not st.session_state.mostrar_campos_finais:
        if st.button("Status modificados"):
            alterado = False
            indices_alterados = []
            pending_logs = []

            responsavel_nome = st.session_state.get("display_name")
            agora = datetime.now()

            # cria um mapa do status original por ID (do df_filtrado)
            status_original_por_id = dict(zip(df_filtrado["ID"].astype(str), df_filtrado["Status"]))
            st.session_state["status_original_por_id"] = status_original_por_id


            for i in range(len(df_filtrado)):
                id_val = str(df_filtrado.iloc[i]["ID"])
                status_original = status_original_por_id.get(id_val, "")
                status_novo = df_editado.iloc[i]["Status"]

                if status_novo != status_original:
                    alterado = True
                    indices_alterados.append(id_val)

                    estudo_da_linha = df.loc[df["ID"].astype(str) == id_val, "Código do Estudo"].iloc[0]


                    pending_logs.append(build_log_rows(
                        id_apontamento=id_val,
                        estudo=estudo_da_linha,
                        operacao="edição",
                        campo="Status",
                        valor_antes=status_original,
                        valor_depois=status_novo,
                        responsavel_nome=responsavel_nome,
                        when=agora,
                        responsavel_indicado=None,
                    ))

                    # aplica a mudança no df base (session)
                    df.loc[df["ID"].astype(str) == id_val, "Status"] = status_novo

            if not alterado:
                st.warning("Nenhuma alteração de status detectada.")
            else:
                st.session_state.mostrar_campos_finais = True
                st.session_state.indices_alterados = indices_alterados
                st.session_state.pending_logs = pending_logs



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

            if status_novo == "VERIFICANDO":
                st.warning(
                    "⚠️ Não é possível alterar o status para **VERIFICANDO** pelo link de apontamentos. "
                    "Essa mudança só pode ser feita no **Painel ADM**."
                )
                st.markdown("---")
                continue  # Pula para o próximo apontamento

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
        responsavel_indicado = st.selectbox("Responsável pela Atualização", options=resp_opts, key="responsavel_final")


        if st.button("Submeter mudanças"):
            responsavel_nome = st.session_state.get("display_name")
            agora = datetime.now()

            if linhas_faltando:
                st.error("Campos obrigatórios pendentes:\n\n" + "\n".join(linhas_faltando))
                st.warning("Por favor, selecione um responsável!")
                st.stop()

            if responsavel_nome == "Selecione um Colaborador":
                st.error("Selecione um responsável válido em 'Responsável pela Atualização'.")
                st.stop()

            # garante coluna
            if "Responsável Indicado" not in df.columns:
                df["Responsável Indicado"] = ""

            # IMPORTANTÍSSIMO: normaliza IDs pra string
            ids_alterados_str = [str(x) for x in indices_alterados]

            pending_logs = st.session_state.get("pending_logs", [])
            status_original_por_id = st.session_state.get("status_original_por_id", {})

            for id_val in ids_alterados_str:
                mask_id = df["ID"].astype(str) == id_val
                if not mask_id.any():
                    continue

                # valor_antes = status original
                valor_antes = status_original_por_id.get(str(id_val), "")

                # valor_depois = novo status (já aplicado no passo 5 no df base)
                status_novo = df.loc[mask_id, "Status"].iloc[0]



                estudo_da_linha = df.loc[mask_id, "Código do Estudo"].iloc[0]

                pending_logs.append(build_log_rows(
                    id_apontamento=id_val,
                    estudo=estudo_da_linha,
                    operacao="edição",
                    campo="Status",
                    valor_antes=valor_antes,          # status original
                    valor_depois=status_novo,         # novo status
                    responsavel_nome=responsavel_nome,
                    when=agora,
                    responsavel_indicado=responsavel_indicado, # parâmetro separado
                ))


            # auditoria padrão
            mask_all = df["ID"].astype(str).isin(ids_alterados_str)
            df.loc[mask_all, "Responsável Atualização"] = responsavel_nome
            df.loc[mask_all, "Data Atualização"] = agora.strftime("%Y-%m-%d %H:%M:%S")

            rows_to_save = df.loc[mask_all].copy()
            df_logs_add = pd.DataFrame(pending_logs) if pending_logs else pd.DataFrame(columns=LOG_COLUMNS)

            with st.spinner("💾 Salvando apontamento no SharePoint..."):
                res = update_sharepoint_workbook(apontamentos_delta=rows_to_save, logs_append=df_logs_add)
                if res is not None:
                    df_atualizado, df_logs_atualizado = res
                    st.session_state["df_apontamentos"] = df_atualizado
                    st.session_state["df_logs"] = df_logs_atualizado
                    st.success("✅ Apontamento submetido com sucesso!")




                # limpa estados
                
                st.session_state.mostrar_campos_finais = False
                st.session_state.indices_alterados = []
                st.session_state.pending_logs = []
                st.session_state["_reset_filtros"] = True
                st.rerun()