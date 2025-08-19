import streamlit as st
import pandas as pd
from datetime import datetime
import io
import time

# >>> usa o conector (precisa do arquivo sp_connector.py no repo)
from sp_connector import SPConnector

# === Configurações via secrets (SharePoint - site servicosclinicos) ===
TENANT_ID = st.secrets["graph"]["tenant_id"]
CLIENT_ID = st.secrets["graph"]["client_id"]
CLIENT_SECRET = st.secrets["graph"]["client_secret"]
HOSTNAME = st.secrets["graph"]["hostname"]            # 'synviagroup.sharepoint.com'
SITE_PATH = st.secrets["graph"]["site_path"]          # 'sites/servicosclinicos'
LIBRARY   = st.secrets["graph"]["library_name"]       # 'Acompanhamento dos Estudos'

APONTAMENTOS  = st.secrets["files"]["apontamentos"]   # 'SANDRA/PROJETO_DASHBOARD/base_apontamentos - Copia.xlsx'
ESTUDOS_CSV   = st.secrets["files"]["estudos_csv"]    # 'SANDRA/PROJETO_DASHBOARD/ESTUDOS_BIO.csv'
COLABORADORES = st.secrets["files"]["colaboradores"]  # 'SANDRA/PROJETO_DASHBOARD/base_cargo.xlsx'

# Instância única do conector (cacheada)
@st.cache_resource
def _sp():
    return SPConnector(
        TENANT_ID, CLIENT_ID, CLIENT_SECRET,
        hostname=HOSTNAME, site_path=SITE_PATH, library_name=LIBRARY
    )

# Função para ler o arquivo Excel (Apontamentos) do SharePoint com cache
@st.cache_data
def get_sharepoint_file():
    try:
        return _sp().read_excel(APONTAMENTOS)
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
def update_sharepoint_file(df: pd.DataFrame):
    attempts = 0
    while True:
        try:
            output = io.BytesIO()
            df.to_excel(output, index=False)  # grava um único sheet (Sheet1)
            output.seek(0)
            _sp().upload_small(APONTAMENTOS, output.getvalue(), overwrite=True)
            st.cache_data.clear()
            st.session_state["df_apontamentos"] = df
            st.success("Mudanças submetidas com sucesso! Recarregue a página para ver as mudanças")
            break
        except Exception as e:
            attempts += 1
            msg = str(e)
            # 409/412 = conflito de versão | 429 = throttling
            if any(x in msg for x in ["409", "412", "429"]) and attempts < 5:
                st.warning("Outra pessoa está salvando ou limite de chamadas. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue
            st.error(f"Erro ao salvar no SharePoint (Graph): {msg}")
            break

# Carregar dados iniciais
df_study = get_sharepoint_file_estudos_csv()
colaboradores_df = colaboradores_excel()


# Inicializar o DataFrame de apontamentos no session_state
if "df_apontamentos" not in st.session_state:
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
    label="Navegação",  
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
        st.text_input("Nome da Pesquisa", value=research_name, disabled=True, icon="🔍")
        
        responsavel = st.text_input("Responsável pelo Apontamento", key="responsavel_apontamento", icon="👤")
        
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
        

        participante = st.selectbox("Participante", 
            ['N/A', 'Outros', 'PP01', 'PP02', 'PP03', 'PP04', 'PP05', 'PP06', 'PP07', 'PP08', 'PP09', 'PP10', 'PP11', 'PP12', 'PP13', 'PP14', 'PP15', 'PP16', 'PP17', 'PP18', 'PP19', 'PP20', 'PP21', 'PP22', 'PP23', 'PP24', 'PP25', 'PP26', 'PP27', 'PP28', 'PP29', 'PP30', 'PP31', 'PP32', 'PP33', 'PP34', 'PP35', 'PP36', 'PP37', 'PP38', 'PP39', 'PP40', 'PP41', 'PP42', 'PP43', 'PP44', 'PP45', 'PP46', 'PP47', 'PP48', 'PP49', 'PP50', 'PP51', 'PP52', 'PP53', 'PP54', 'PP55', 'PP56', 'PP57', 'PP58', 'PP59', 'PP60', 'PP61', 'PP62', 'PP63', 'PP64', 'PP65', 'PP66', 'PP67', 'PP68', 'PP69', 'PP70', 'PP71', 'PP72', 'PP73', 'PP74', 'PP75', 'PP76', 'PP77', 'PP78', 'PP79', 'PP80', 'PP81', 'PP82', 'PP83', 'PP84', 'PP85', 'PP86', 'PP87', 'PP88', 'PP89', 'PP90', 'PP91', 'PP92', 'PP93', 'PP94', 'PP95', 'PP96', 'PP97', 'PP98', 'PP99', 'PP100', 'PP101', 'PP102', 'PP103', 'PP104', 'PP105', 'PP106', 'PP107', 'PP108', 'PP109', 'PP110', 'PP111', 'PP112', 'PP113', 'PP114', 'PP115', 'PP116', 'PP117', 'PP118', 'PP119', 'PP120', 'PP121', 'PP122', 'PP123', 'PP124', 'PP125', 'PP126', 'PP127', 'PP128', 'PP129', 'PP130', 'PP131', 'PP132', 'PP133', 'PP134', 'PP135', 'PP136', 'PP137', 'PP138', 'PP139', 'PP140', 'PP141', 'PP142', 'PP143', 'PP144', 'PP145', 'PP146', 'PP147', 'PP148', 'PP149', 'PP150', 'PP151', 'PP152', 'PP153', 'PP154', 'PP155', 'PP156', 'PP157', 'PP158', 'PP159', 'PP160', 'PP161', 'PP162', 'PP163', 'PP164', 'PP165', 'PP166', 'PP167', 'PP168', 'PP169', 'PP170', 'PP171', 'PP172', 'PP173', 'PP174', 'PP175', 'PP176', 'PP177', 'PP178', 'PP179', 'PP180', 'PP181', 'PP182', 'PP183', 'PP184', 'PP185', 'PP186', 'PP187', 'PP188', 'PP189', 'PP190', 'PP191', 'PP192', 'PP193', 'PP194', 'PP195', 'PP196', 'PP197', 'PP198', 'PP199', 'PP200', 'PP201', 'PP202', 'PP203', 'PP204', 'PP205', 'PP206', 'PP207', 'PP208', 'PP209', 'PP210', 'PP211', 'PP212', 'PP213', 'PP214', 'PP215', 'PP216', 'PP217', 'PP218', 'PP219', 'PP220', 'PP221', 'PP222', 'PP223', 'PP224', 'PP225', 'PP226', 'PP227', 'PP228', 'PP229', 'PP230', 'PP231', 'PP232', 'PP233', 'PP234', 'PP235', 'PP236', 'PP237', 'PP238', 'PP239', 'PP240', 'PP241', 'PP242', 'PP243', 'PP244', 'PP245', 'PP246', 'PP247', 'PP248', 'PP249', 'PP250', 'PP251', 'PP252', 'PP253', 'PP254', 'PP255', 'PP256', 'PP257', 'PP258', 'PP259', 'PP260', 'PP261', 'PP262', 'PP263', 'PP264', 'PP265', 'PP266', 'PP267', 'PP268', 'PP269', 'PP270', 'PP271', 'PP272', 'PP273', 'PP274', 'PP275', 'PP276', 'PP277', 'PP278', 'PP279', 'PP280', 'PP281', 'PP282', 'PP283', 'PP284', 'PP285', 'PP286', 'PP287', 'PP288', 'PP289', 'PP290', 'PP291', 'PP292', 'PP293', 'PP294', 'PP295', 'PP296', 'PP297', 'PP298', 'PP299', 'PP300', 'PP301', 'PP302', 'PP303', 'PP304', 'PP305', 'PP306', 'PP307', 'PP308', 'PP309', 'PP310', 'PP311', 'PP312', 'PP313', 'PP314', 'PP315', 'PP316', 'PP317', 'PP318', 'PP319', 'PP320', 'PP321', 'PP322', 'PP323', 'PP324', 'PP325', 'PP326', 'PP327', 'PP328', 'PP329', 'PP330', 'PP331', 'PP332', 'PP333', 'PP334', 'PP335', 'PP336', 'PP337', 'PP338', 'PP339', 'PP340', 'PP341', 'PP342', 'PP343', 'PP344', 'PP345', 'PP346', 'PP347', 'PP348', 'PP349', 'PP350', 'PP351', 'PP352', 'PP353', 'PP354', 'PP355', 'PP356', 'PP357', 'PP358', 'PP359', 'PP360', 'PP361', 'PP362', 'PP363', 'PP364', 'PP365', 'PP366', 'PP367', 'PP368', 'PP369', 'PP370', 'PP371', 'PP372', 'PP373', 'PP374', 'PP375', 'PP376', 'PP377', 'PP378', 'PP379', 'PP380', 'PP381', 'PP382', 'PP383', 'PP384', 'PP385', 'PP386', 'PP387', 'PP388', 'PP389', 'PP390', 'PP391', 'PP392', 'PP393', 'PP394', 'PP395', 'PP396', 'PP397', 'PP398', 'PP399', 'PP400', 'PP401', 'PP402', 'PP403', 'PP404', 'PP405', 'PP406', 'PP407', 'PP408', 'PP409', 'PP410', 'PP411', 'PP412', 'PP413', 'PP414', 'PP415', 'PP416', 'PP417', 'PP418', 'PP419', 'PP420', 'PP421', 'PP422', 'PP423', 'PP424', 'PP425', 'PP426', 'PP427', 'PP428', 'PP429', 'PP430', 'PP431', 'PP432', 'PP433', 'PP434', 'PP435', 'PP436', 'PP437', 'PP438', 'PP439', 'PP440', 'PP441', 'PP442', 'PP443', 'PP444', 'PP445', 'PP446', 'PP447', 'PP448', 'PP449', 'PP450', 'PP451', 'PP452', 'PP453', 'PP454', 'PP455', 'PP456', 'PP457', 'PP458', 'PP459', 'PP460', 'PP461', 'PP462', 'PP463', 'PP464', 'PP465', 'PP466', 'PP467', 'PP468', 'PP469', 'PP470', 'PP471', 'PP472', 'PP473', 'PP474', 'PP475', 'PP476', 'PP477', 'PP478', 'PP479', 'PP480', 'PP481', 'PP482', 'PP483', 'PP484', 'PP485', 'PP486', 'PP487', 'PP488', 'PP489', 'PP490', 'PP491', 'PP492', 'PP493', 'PP494', 'PP495', 'PP496', 'PP497', 'PP498', 'PP499', 'PP500', 'PP501', 'PP502', 'PP503', 'PP504', 'PP505', 'PP506', 'PP507', 'PP508', 'PP509', 'PP510', 'PP511', 'PP512', 'PP513', 'PP514', 'PP515', 'PP516', 'PP517', 'PP518', 'PP519', 'PP520', 'PP521', 'PP522', 'PP523', 'PP524', 'PP525', 'PP526', 'PP527', 'PP528', 'PP529', 'PP530', 'PP531', 'PP532', 'PP533', 'PP534', 'PP535', 'PP536', 'PP537', 'PP538', 'PP539', 'PP540', 'PP541', 'PP542', 'PP543', 'PP544', 'PP545', 'PP546', 'PP547', 'PP548', 'PP549', 'PP550', 'PP551', 'PP552', 'PP553', 'PP554', 'PP555', 'PP556', 'PP557', 'PP558', 'PP559', 'PP560', 'PP561', 'PP562', 'PP563', 'PP564', 'PP565', 'PP566', 'PP567', 'PP568', 'PP569', 'PP570', 'PP571', 'PP572', 'PP573', 'PP574', 'PP575', 'PP576', 'PP577', 'PP578', 'PP579', 'PP580', 'PP581', 'PP582', 'PP583', 'PP584', 'PP585', 'PP586', 'PP587', 'PP588', 'PP589', 'PP590', 'PP591', 'PP592', 'PP593', 'PP594', 'PP595', 'PP596', 'PP597', 'PP598', 'PP599', 'PP600', 'PP601', 'PP602', 'PP603', 'PP604', 'PP605', 'PP606', 'PP607', 'PP608', 'PP609', 'PP610', 'PP611', 'PP612', 'PP613', 'PP614', 'PP615', 'PP616', 'PP617', 'PP618', 'PP619', 'PP620', 'PP621', 'PP622', 'PP623', 'PP624', 'PP625', 'PP626', 'PP627', 'PP628', 'PP629', 'PP630', 'PP631', 'PP632', 'PP633', 'PP634', 'PP635', 'PP636', 'PP637', 'PP638', 'PP639', 'PP640', 'PP641', 'PP642', 'PP643', 'PP644', 'PP645', 'PP646', 'PP647', 'PP648', 'PP649', 'PP650', 'PP651', 'PP652', 'PP653', 'PP654', 'PP655', 'PP656', 'PP657', 'PP658', 'PP659', 'PP660', 'PP661', 'PP662', 'PP663', 'PP664', 'PP665', 'PP666', 'PP667', 'PP668', 'PP669', 'PP670', 'PP671', 'PP672', 'PP673', 'PP674', 'PP675', 'PP676', 'PP677', 'PP678', 'PP679', 'PP680', 'PP681', 'PP682', 'PP683', 'PP684', 'PP685', 'PP686', 'PP687', 'PP688', 'PP689', 'PP690', 'PP691', 'PP692', 'PP693', 'PP694', 'PP695', 'PP696', 'PP697', 'PP698', 'PP699', 'PP700', 'PP701', 'PP702', 'PP703', 'PP704', 'PP705', 'PP706', 'PP707', 'PP708', 'PP709', 'PP710', 'PP711', 'PP712', 'PP713', 'PP714', 'PP715', 'PP716', 'PP717', 'PP718', 'PP719', 'PP720', 'PP721', 'PP722', 'PP723', 'PP724', 'PP725', 'PP726', 'PP727', 'PP728', 'PP729', 'PP730', 'PP731', 'PP732', 'PP733', 'PP734', 'PP735', 'PP736', 'PP737', 'PP738', 'PP739', 'PP740', 'PP741', 'PP742', 'PP743', 'PP744', 'PP745', 'PP746', 'PP747', 'PP748', 'PP749', 'PP750', 'PP751', 'PP752', 'PP753', 'PP754', 'PP755', 'PP756', 'PP757', 'PP758', 'PP759', 'PP760', 'PP761', 'PP762', 'PP763', 'PP764', 'PP765', 'PP766', 'PP767', 'PP768', 'PP769', 'PP770', 'PP771', 'PP772', 'PP773', 'PP774', 'PP775', 'PP776', 'PP777', 'PP778', 'PP779', 'PP780', 'PP781', 'PP782', 'PP783', 'PP784', 'PP785', 'PP786', 'PP787', 'PP788', 'PP789', 'PP790', 'PP791', 'PP792', 'PP793', 'PP794', 'PP795', 'PP796', 'PP797', 'PP798', 'PP799', 'PP800', 'PP801', 'PP802', 'PP803', 'PP804', 'PP805', 'PP806', 'PP807', 'PP808', 'PP809', 'PP810', 'PP811', 'PP812', 'PP813', 'PP814', 'PP815', 'PP816', 'PP817', 'PP818', 'PP819', 'PP820', 'PP821', 'PP822', 'PP823', 'PP824', 'PP825', 'PP826', 'PP827', 'PP828', 'PP829', 'PP830', 'PP831', 'PP832', 'PP833', 'PP834', 'PP835', 'PP836', 'PP837', 'PP838', 'PP839', 'PP840', 'PP841', 'PP842', 'PP843', 'PP844', 'PP845', 'PP846', 'PP847', 'PP848', 'PP849', 'PP850', 'PP851', 'PP852', 'PP853', 'PP854', 'PP855', 'PP856', 'PP857', 'PP858', 'PP859', 'PP860', 'PP861', 'PP862', 'PP863', 'PP864', 'PP865', 'PP866', 'PP867', 'PP868', 'PP869', 'PP870', 'PP871', 'PP872', 'PP873', 'PP874', 'PP875', 'PP876', 'PP877', 'PP878', 'PP879', 'PP880', 'PP881', 'PP882', 'PP883', 'PP884', 'PP885', 'PP886', 'PP887', 'PP888', 'PP889', 'PP890', 'PP891', 'PP892', 'PP893', 'PP894', 'PP895', 'PP896', 'PP897', 'PP898', 'PP899', 'PP900', 'PP901', 'PP902', 'PP903', 'PP904', 'PP905', 'PP906', 'PP907', 'PP908', 'PP909', 'PP910', 'PP911', 'PP912', 'PP913', 'PP914', 'PP915', 'PP916', 'PP917', 'PP918', 'PP919', 'PP920', 'PP921', 'PP922', 'PP923', 'PP924', 'PP925', 'PP926', 'PP927', 'PP928', 'PP929', 'PP930', 'PP931', 'PP932', 'PP933', 'PP934', 'PP935', 'PP936', 'PP937', 'PP938', 'PP939', 'PP940', 'PP941', 'PP942', 'PP943', 'PP944', 'PP945', 'PP946', 'PP947', 'PP948', 'PP949', 'PP950', 'PP951', 'PP952', 'PP953', 'PP954', 'PP955', 'PP956', 'PP957', 'PP958', 'PP959', 'PP960', 'PP961', 'PP962', 'PP963', 'PP964', 'PP965', 'PP966', 'PP967', 'PP968', 'PP969', 'PP970', 'PP971', 'PP972', 'PP973', 'PP974', 'PP975', 'PP976', 'PP977', 'PP978', 'PP979', 'PP980', 'PP981', 'PP982', 'PP983', 'PP984', 'PP985', 'PP986', 'PP987', 'PP988', 'PP989', 'PP990', 'PP991', 'PP992', 'PP993', 'PP994', 'PP995', 'PP996', 'PP997', 'PP998', 'PP999']
            , key="participante")
        
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
        status = st.selectbox("Status", [
            "REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"
        ], key="status", on_change=update_status_fields)
        

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
            if selected_protocol == "Digite o codigo do estudo" or participante.strip() == "" or apontamento.strip() == "" or responsavel.strip() == "":
                st.error("Por favor, preencha os campos obrigatórios: Código do Estudo, Participante e Responsável.")
            elif status == "VERIFICANDO" and verificador_nome.strip() == "":
                st.error("Somente o Guilherme Gonçalves pode usar esse status!.")
            elif  status == "":
                st.error("Por favor, defina um status antes de submeter o apontamento!")
            elif status == "NÃO APLICÁVEL" and justificativa.strip() == "":
                st.error("Por favor, preencha o campo 'Justificativa'!")
            elif responsavel == "Selecione um colaborador":
                st.warning("Por favor, selecione o colaborador responsável antes de salvar.")
                st.stop()
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
                
                

                novo_apontamento = {
                    "ID": next_id,
                    "Código do Estudo": selected_protocol,
                    "Nome da Pesquisa": research_name,
                    "Data do Apontamento": data_atual,
                    "Responsável Pelo Apontamento": responsavel,
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
                    "Tempo de casa": status_prof
                }
                


                novo_df = pd.DataFrame([novo_apontamento])
                df = pd.concat([df, novo_df], ignore_index=True)
                update_sharepoint_file(df)
                st.session_state["df_apontamentos"] = df
                st.session_state["generated_id"] = generate_custom_id(set(df["ID"].astype(str)))
                



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
    st.session_state.setdefault("df_atualizado", None)

    st.title("Lista de Apontamentos")

    col_btn1, *_ = st.columns(6)
    with col_btn1:
        if st.button("🔄  Atualizar", key="btn_clear_cache"):
            st.cache_data.clear()

    # ─────────────────────────────────────────────────────────────
    # 3️⃣  Filtros rápidos / seletor de estudo
    # ─────────────────────────────────────────────────────────────
    if df.empty:
        st.info("Nenhum apontamento encontrado!")
        st.stop()

    df_filtrado = df.copy()

    campo_id = st.text
    opcoes_estudos = ["Todos"] + sorted(df["Código do Estudo"].dropna().unique().tolist())
    estudo_sel = st.selectbox("Selecione o Estudo", options=opcoes_estudos)

    if estudo_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Código do Estudo"] == estudo_sel]

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
                    df_atualizado.loc[df_atualizado["ID"] == id_val, "Status"] = status_novo
                    indices_alterados.append(id_val)

            if not alterado:
                st.warning("Nenhuma alteração de status detectada.")
            else:
                st.session_state.mostrar_campos_finais = True
                st.session_state.indices_alterados = indices_alterados
                st.session_state.df_atualizado = df_atualizado

    # ─────────────────────────────────────────────────────────────
    # 6️⃣  Campos finais obrigatórios + submissão
    # ─────────────────────────────────────────────────────────────
    if st.session_state.mostrar_campos_finais:
        df = st.session_state.df_atualizado
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
                df.loc[df["ID"].isin(indices_alterados), "Verificador"] = responsavel

                # Salva de volta
                update_sharepoint_file(df)
                st.session_state["df_apontamentos"] = df

                # Limpa estados
                st.session_state.mostrar_campos_finais = False
                st.session_state.indices_alterados = []
                st.session_state.df_atualizado = None
