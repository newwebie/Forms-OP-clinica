import streamlit as st
import pandas as pd
from datetime import datetime
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
        st.cache_data.clear()
        st.success("Para ver as mudanças ou submeter novas alterações, tecle F5")
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
            ['N/A','PP01', 'PP02', 'PP03', 'PP04', 'PP05', 'PP06', 'PP07', 'PP08', 'PP09', 'PP10', 'PP11', 'PP12', 'PP13', 'PP14', 'PP15', 'PP16', 'PP17', 'PP18', 'PP19', 'PP20', 'PP21', 'PP22', 'PP23', 'PP24', 'PP25', 'PP26', 'PP27', 'PP28', 'PP29', 'PP30', 'PP31', 'PP32', 'PP33', 'PP34', 'PP35', 'PP36', 'PP37', 'PP38', 'PP39', 'PP40', 'PP41', 'PP42', 'PP43', 'PP44', 'PP45', 'PP46', 'PP47', 'PP48', 'PP49', 'PP50', 'PP51', 'PP52', 'PP53', 'PP54', 'PP55', 'PP56', 'PP57', 'PP58', 'PP59', 'PP60', 'PP61', 'PP62', 'PP63', 'PP64', 'PP65', 'PP66', 'PP67', 'PP68', 'PP69', 'PP70', 'PP71', 'PP72', 'PP73', 'PP74', 'PP75', 'PP76', 'PP77', 'PP78', 'PP79', 'PP80', 'PP81', 'PP82', 'PP83', 'PP84', 'PP85', 'PP86', 'PP87', 'PP88', 'PP89', 'PP90', 'PP91', 'PP92', 'PP93', 'PP94', 'PP95', 'PP96', 'PP97', 'PP98', 'PP99', 'PP100', 'PP101', 'PP102', 'PP103', 'PP104', 'PP105', 'PP106', 'PP107', 'PP108', 'PP109', 'PP110', 'PP111', 'PP112', 'PP113', 'PP114', 'PP115', 'PP116', 'PP117', 'PP118', 'PP119', 'PP120', 'PP121', 'PP122', 'PP123', 'PP124', 'PP125', 'PP126', 'PP127', 'PP128', 'PP129', 'PP130', 'PP131', 'PP132', 'PP133', 'PP134', 'PP135', 'PP136', 'PP137', 'PP138', 'PP139', 'PP140', 'PP141', 'PP142', 'PP143', 'PP144', 'PP145', 'PP146', 'PP147', 'PP148', 'PP149', 'PP150', 'PP151', 'PP152', 'PP153', 'PP154', 'PP155', 'PP156', 'PP157', 'PP158', 'PP159', 'PP160', 'PP161', 'PP162', 'PP163', 'PP164', 'PP165', 'PP166', 'PP167', 'PP168', 'PP169', 'PP170', 'PP171', 'PP172', 'PP173', 'PP174', 'PP175', 'PP176', 'PP177', 'PP178', 'PP179', 'PP180', 'PP181', 'PP182', 'PP183', 'PP184', 'PP185', 'PP186', 'PP187', 'PP188', 'PP189', 'PP190', 'PP191', 'PP192', 'PP193', 'PP194', 'PP195', 'PP196', 'PP197', 'PP198', 'PP199', 'PP200', 'PP201', 'PP202', 'PP203', 'PP204', 'PP205', 'PP206', 'PP207', 'PP208', 'PP209', 'PP210', 'PP211', 'PP212', 'PP213', 'PP214', 'PP215', 'PP216', 'PP217', 'PP218', 'PP219', 'PP220', 'PP221', 'PP222', 'PP223', 'PP224', 'PP225', 'PP226', 'PP227', 'PP228', 'PP229', 'PP230', 'PP231', 'PP232', 'PP233', 'PP234', 'PP235', 'PP236', 'PP237', 'PP238', 'PP239', 'PP240', 'PP241', 'PP242', 'PP243', 'PP244', 'PP245', 'PP246', 'PP247', 'PP248', 'PP249', 'PP250', 'PP251', 'PP252', 'PP253', 'PP254', 'PP255', 'PP256', 'PP257', 'PP258', 'PP259', 'PP260', 'PP261', 'PP262', 'PP263', 'PP264', 'PP265', 'PP266', 'PP267', 'PP268', 'PP269', 'PP270', 'PP271', 'PP272', 'PP273', 'PP274', 'PP275', 'PP276', 'PP277', 'PP278', 'PP279', 'PP280', 'PP281', 'PP282', 'PP283', 'PP284', 'PP285', 'PP286', 'PP287', 'PP288', 'PP289', 'PP290', 'PP291', 'PP292', 'PP293', 'PP294', 'PP295', 'PP296', 'PP297', 'PP298', 'PP299', 'PP300', 'PP301', 'PP302', 'PP303', 'PP304', 'PP305', 'PP306', 'PP307', 'PP308', 'PP309', 'PP310', 'PP311', 'PP312', 'PP313', 'PP314', 'PP315', 'PP316', 'PP317', 'PP318', 'PP319', 'PP320', 'PP321', 'PP322', 'PP323', 'PP324', 'PP325', 'PP326', 'PP327', 'PP328', 'PP329', 'PP330', 'PP331', 'PP332', 'PP333', 'PP334', 'PP335', 'PP336', 'PP337', 'PP338', 'PP339', 'PP340', 'PP341', 'PP342', 'PP343', 'PP344', 'PP345', 'PP346', 'PP347', 'PP348', 'PP349', 'PP350', 'PP351', 'PP352', 'PP353', 'PP354', 'PP355', 'PP356', 'PP357', 'PP358', 'PP359', 'PP360', 'PP361', 'PP362', 'PP363', 'PP364', 'PP365', 'PP366', 'PP367', 'PP368', 'PP369', 'PP370', 'PP371', 'PP372', 'PP373', 'PP374', 'PP375', 'PP376', 'PP377', 'PP378', 'PP379', 'PP380', 'PP381', 'PP382', 'PP383', 'PP384', 'PP385', 'PP386', 'PP387', 'PP388', 'PP389', 'PP390', 'PP391', 'PP392', 'PP393', 'PP394', 'PP395', 'PP396', 'PP397', 'PP398', 'PP399', 'PP400', 'PP401', 'PP402', 'PP403', 'PP404', 'PP405', 'PP406', 'PP407', 'PP408', 'PP409', 'PP410', 'PP411', 'PP412', 'PP413', 'PP414', 'PP415', 'PP416', 'PP417', 'PP418', 'PP419', 'PP420', 'PP421', 'PP422', 'PP423', 'PP424', 'PP425', 'PP426', 'PP427', 'PP428', 'PP429', 'PP430', 'PP431', 'PP432', 'PP433', 'PP434', 'PP435', 'PP436', 'PP437', 'PP438', 'PP439', 'PP440', 'PP441', 'PP442', 'PP443', 'PP444', 'PP445', 'PP446', 'PP447', 'PP448', 'PP449', 'PP450', 'PP451', 'PP452', 'PP453', 'PP454', 'PP455', 'PP456', 'PP457', 'PP458', 'PP459', 'PP460', 'PP461', 'PP462', 'PP463', 'PP464', 'PP465', 'PP466', 'PP467', 'PP468', 'PP469', 'PP470', 'PP471', 'PP472', 'PP473', 'PP474', 'PP475', 'PP476', 'PP477', 'PP478', 'PP479', 'PP480', 'PP481', 'PP482', 'PP483', 'PP484', 'PP485', 'PP486', 'PP487', 'PP488', 'PP489', 'PP490', 'PP491', 'PP492', 'PP493', 'PP494', 'PP495', 'PP496', 'PP497', 'PP498', 'PP499', 'PP500', 'PP501', 'PP502', 'PP503', 'PP504', 'PP505', 'PP506', 'PP507', 'PP508', 'PP509', 'PP510', 'PP511', 'PP512', 'PP513', 'PP514', 'PP515', 'PP516', 'PP517', 'PP518', 'PP519', 'PP520', 'PP521', 'PP522', 'PP523', 'PP524', 'PP525', 'PP526', 'PP527', 'PP528', 'PP529', 'PP530', 'PP531', 'PP532', 'PP533', 'PP534', 'PP535', 'PP536', 'PP537', 'PP538', 'PP539', 'PP540', 'PP541', 'PP542', 'PP543', 'PP544', 'PP545', 'PP546', 'PP547', 'PP548', 'PP549', 'PP550', 'PP551', 'PP552', 'PP553', 'PP554', 'PP555', 'PP556', 'PP557', 'PP558', 'PP559', 'PP560', 'PP561', 'PP562', 'PP563', 'PP564', 'PP565', 'PP566', 'PP567', 'PP568', 'PP569', 'PP570', 'PP571', 'PP572', 'PP573', 'PP574', 'PP575', 'PP576', 'PP577', 'PP578', 'PP579', 'PP580', 'PP581', 'PP582', 'PP583', 'PP584', 'PP585', 'PP586', 'PP587', 'PP588', 'PP589', 'PP590', 'PP591', 'PP592', 'PP593', 'PP594', 'PP595', 'PP596', 'PP597', 'PP598', 'PP599', 'PP600', 'PP601', 'PP602', 'PP603', 'PP604', 'PP605', 'PP606', 'PP607', 'PP608', 'PP609', 'PP610', 'PP611', 'PP612', 'PP613', 'PP614', 'PP615', 'PP616', 'PP617', 'PP618', 'PP619', 'PP620', 'PP621', 'PP622', 'PP623', 'PP624', 'PP625', 'PP626', 'PP627', 'PP628', 'PP629', 'PP630', 'PP631', 'PP632', 'PP633', 'PP634', 'PP635', 'PP636', 'PP637', 'PP638', 'PP639', 'PP640', 'PP641', 'PP642', 'PP643', 'PP644', 'PP645', 'PP646', 'PP647', 'PP648', 'PP649', 'PP650', 'PP651', 'PP652', 'PP653', 'PP654', 'PP655', 'PP656', 'PP657', 'PP658', 'PP659', 'PP660', 'PP661', 'PP662', 'PP663', 'PP664', 'PP665', 'PP666', 'PP667', 'PP668', 'PP669', 'PP670', 'PP671', 'PP672', 'PP673', 'PP674', 'PP675', 'PP676', 'PP677', 'PP678', 'PP679', 'PP680', 'PP681', 'PP682', 'PP683', 'PP684', 'PP685', 'PP686', 'PP687', 'PP688', 'PP689', 'PP690', 'PP691', 'PP692', 'PP693', 'PP694', 'PP695', 'PP696', 'PP697', 'PP698', 'PP699', 'PP700', 'PP701', 'PP702', 'PP703', 'PP704', 'PP705', 'PP706', 'PP707', 'PP708', 'PP709', 'PP710', 'PP711', 'PP712', 'PP713', 'PP714', 'PP715', 'PP716', 'PP717', 'PP718', 'PP719', 'PP720', 'PP721', 'PP722', 'PP723', 'PP724', 'PP725', 'PP726', 'PP727', 'PP728', 'PP729', 'PP730', 'PP731', 'PP732', 'PP733', 'PP734', 'PP735', 'PP736', 'PP737', 'PP738', 'PP739', 'PP740', 'PP741', 'PP742', 'PP743', 'PP744', 'PP745', 'PP746', 'PP747', 'PP748', 'PP749', 'PP750', 'PP751', 'PP752', 'PP753', 'PP754', 'PP755', 'PP756', 'PP757', 'PP758', 'PP759', 'PP760', 'PP761', 'PP762', 'PP763', 'PP764', 'PP765', 'PP766', 'PP767', 'PP768', 'PP769', 'PP770', 'PP771', 'PP772', 'PP773', 'PP774', 'PP775', 'PP776', 'PP777', 'PP778', 'PP779', 'PP780', 'PP781', 'PP782', 'PP783', 'PP784', 'PP785', 'PP786', 'PP787', 'PP788', 'PP789', 'PP790', 'PP791', 'PP792', 'PP793', 'PP794', 'PP795', 'PP796', 'PP797', 'PP798', 'PP799', 'PP800', 'PP801', 'PP802', 'PP803', 'PP804', 'PP805', 'PP806', 'PP807', 'PP808', 'PP809', 'PP810', 'PP811', 'PP812', 'PP813', 'PP814', 'PP815', 'PP816', 'PP817', 'PP818', 'PP819', 'PP820', 'PP821', 'PP822', 'PP823', 'PP824', 'PP825', 'PP826', 'PP827', 'PP828', 'PP829', 'PP830', 'PP831', 'PP832', 'PP833', 'PP834', 'PP835', 'PP836', 'PP837', 'PP838', 'PP839', 'PP840', 'PP841', 'PP842', 'PP843', 'PP844', 'PP845', 'PP846', 'PP847', 'PP848', 'PP849', 'PP850', 'PP851', 'PP852', 'PP853', 'PP854', 'PP855', 'PP856', 'PP857', 'PP858', 'PP859', 'PP860', 'PP861', 'PP862', 'PP863', 'PP864', 'PP865', 'PP866', 'PP867', 'PP868', 'PP869', 'PP870', 'PP871', 'PP872', 'PP873', 'PP874', 'PP875', 'PP876', 'PP877', 'PP878', 'PP879', 'PP880', 'PP881', 'PP882', 'PP883', 'PP884', 'PP885', 'PP886', 'PP887', 'PP888', 'PP889', 'PP890', 'PP891', 'PP892', 'PP893', 'PP894', 'PP895', 'PP896', 'PP897', 'PP898', 'PP899', 'PP900', 'PP901', 'PP902', 'PP903', 'PP904', 'PP905', 'PP906', 'PP907', 'PP908', 'PP909', 'PP910', 'PP911', 'PP912', 'PP913', 'PP914', 'PP915', 'PP916', 'PP917', 'PP918', 'PP919', 'PP920', 'PP921', 'PP922', 'PP923', 'PP924', 'PP925', 'PP926', 'PP927', 'PP928', 'PP929', 'PP930', 'PP931', 'PP932', 'PP933', 'PP934', 'PP935', 'PP936', 'PP937', 'PP938', 'PP939', 'PP940', 'PP941', 'PP942', 'PP943', 'PP944', 'PP945', 'PP946', 'PP947', 'PP948', 'PP949', 'PP950', 'PP951', 'PP952', 'PP953', 'PP954', 'PP955', 'PP956', 'PP957', 'PP958', 'PP959', 'PP960', 'PP961', 'PP962', 'PP963', 'PP964', 'PP965', 'PP966', 'PP967', 'PP968', 'PP969', 'PP970', 'PP971', 'PP972', 'PP973', 'PP974', 'PP975', 'PP976', 'PP977', 'PP978', 'PP979', 'PP980', 'PP981', 'PP982', 'PP983', 'PP984', 'PP985', 'PP986', 'PP987', 'PP988', 'PP989', 'PP990', 'PP991', 'PP992', 'PP993', 'PP994', 'PP995', 'PP996', 'PP997', 'PP998', 'PP999']
            , key="participante")
             
             
        periodo = st.selectbox("Período", [
            '1° Período', '2° Período', '3° Período',
            '4° Período', '5° Período', '6° Período', '7° Período', 
            '8° Período', '9° Período', '10° Período'
        ], key="periodo")
        
        criticidade = st.selectbox("Grau De Criticidade Do Apontamento", ["Baixo", "Médio", "Alto"], key="criticidade")
        st.text('Baixo: O apontamento tem rastreabilidade e não gera impacto no RC \nMédio: O apontamento tem rastreabilidade e gera impacto no RC, precisa de correção em sistema \nAlta: O apontamento não tem rastreabilidade')
        prazo = st.date_input("Prazo Para Resolução", format="DD/MM/YYYY", key="prazo")
        apontamento = st.text_area("Apontamento", key="apontamento")
        
        # Campo de Status com callback (supondo que a função update_status_fields esteja definida)
        status = st.selectbox("Status", [
            "REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"
        ], key="status", on_change=update_status_fields)
        
        if st.session_state["enable_verificador"]:
            verificador_nome = st.text_input("Responsável pela verificação", key="verificador_nome")
            verificador_data = st.date_input("Data de início da verificação", format="DD/MM/YYYY", key="verificador_data")
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
            if selected_protocol == "Digite o codigo do estudo" or participante.strip() == "" or apontamento.strip() == "" or responsavel.strip() == "":
                st.error("Por favor, preencha os campos obrigatórios: Código do Estudo, Participante e Responsável.")
            elif status == "VERIFICANDO" and verificador_nome.strip() == "":
                st.error("Por favor, preencha o campo 'Responsável pela verificação'.")
            elif status == "NÃO APLICÁVEL" and justificativa.strip() == "":
                st.error("Por favor, preencha o campo 'Justificativa'!")
            else:
                novo_apontamento = {
                    "Código do Estudo": selected_protocol,
                    "Nome da Pesquisa": research_name,
                    "Data do Apontamento": datetime.date.today(),
                    "Responsável Pelo Apontamento": responsavel,
                    "Origem Do Apontamento": st.session_state["origem"],
                    "Documentos": documento_final,  # Aqui utiliza o valor final (customizado se "Outros")
                    "Participante": st.session_state["participante"],
                    "Período": st.session_state["periodo"],
                    "Grau De Criticidade Do Apontamento": st.session_state["criticidade"],
                    "Prazo Para Resolução": st.session_state["prazo"],
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
                chave_nova = (selected_protocol, documento_final, st.session_state["participante"])
                
                if chave_nova in duplicado_set:
                    duplicado = df[
                        (df["Código do Estudo"] == selected_protocol) &
                        (df["Documentos"] == documento_final) &
                        (df["Participante"] == st.session_state["participante"])
                    ]
                    data_existente = duplicado.iloc[0]["Data do Apontamento"]
                    st.warning(f"Apontamento já existe. Data do Apontamento: {data_existente}")
                else:
                    novo_df = pd.DataFrame([novo_apontamento])
                    df = pd.concat([df, novo_df], ignore_index=True)
                    update_sharepoint_file(df)
                    st.cache_data.clear()  # Limpa o cache para recarregar os dados atualizados
                    st.session_state["df_apontamentos"] = get_sharepoint_file()
                    st.success("Apontamento enviado com sucesso!")
                    

with tabs[1]:
    st.title("Lista de Apontamentos")
    df = get_sharepoint_file()

    if df.empty:
        st.info("Nenhum apontamento encontrado!")
    else:
        # Garante que as colunas de atualização existam
        for col in ["Atualização", "Responsável Atualização"]:
            if col not in df.columns:
                df[col] = ""

        # Campos de input
        st.markdown("### Informações da Atualização")
        responsavel = st.text_input("Responsável pela atualização", placeholder="Digite seu nome completo")
        justificativa = st.text_area("Justificativa", placeholder="Preencher se for o novo Status for 'NÃO APLICÁVEL'")

        st.divider()

        # Configuração do editor
        columns_config = {}
        for col in df.columns:
            if col == "Status":
                columns_config[col] = st.column_config.SelectboxColumn(
                    "Status",
                    options=["REALIZADO DURANTE A CONDUÇÃO", "REALIZADO", "VERIFICANDO", "PENDENTE", "NÃO APLICÁVEL"],
                    disabled=False
                )
            elif col in ["Data do Apontamento", "Prazo Para Resolução", "Data de Verificação"]:
                columns_config[col] = st.column_config.DateColumn(
                    col,
                    disabled=True,
                    format="DD/MM/YYYY"
                )
            else:
                columns_config[col] = st.column_config.TextColumn(col, disabled=True)

        df_editado = st.data_editor(
            df,
            column_config=columns_config,
            num_rows="fixed",
            key="data_editor"
        )

        if st.button("Submeter Edições"):
            if not responsavel.strip():
                st.warning("Por favor, preencha o nome do responsável pela atualização.")
            else:
                df_atualizado = df.copy()
                alterado = False
                erro_verificando = False
                erro_justificativa = False

                for i in range(len(df)):
                    status_original = df.loc[i, "Status"]
                    status_novo = df_editado.loc[i, "Status"]

                    if status_novo != status_original:
                        alterado = True
                        df_atualizado.loc[i, "Status"] = status_novo
                        df_atualizado.loc[i, "Atualização"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        df_atualizado.loc[i, "Responsável Atualização"] = responsavel

                        if status_novo == "NÃO APLICÁVEL" and not justificativa.strip():
                            erro_justificativa = True

                # Mensagens de erro específicas
                if not alterado:
                    st.warning("Nenhuma alteração foi feita nos status. Nada será submetido.")
                elif erro_justificativas:
                    if erro_justificativa:
                        st.error("Por favor, preencha o campo 'Justificativa' para os apontamentos marcados como 'NÃO APLICÁVEL'.")
                else:
                    update_sharepoint_file(df_atualizado)
                    st.success("Alterações salvas com sucesso!")
