"""
Sistema de Autenticação Microsoft para Streamlit
Implementa login via Azure AD com MSAL
"""

import base64
import os
from functools import lru_cache
from html import escape
from pathlib import Path
from typing import Optional, Dict, Any

import msal
import requests
import streamlit as st
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Paleta de cores alinhada ao projeto Syntox Churn
PRIMARY_COLOR = "#6BBF47"
SECONDARY_COLOR = "#52B54B"
ACCENT_DARK = "#0F1C16"
MUTED_TEXT = "#5B6770"

# Estilos dedicados para a tela de login
LOGIN_PAGE_CSS = f"""
<style>

.login-inner {{
    position: relative;
    width: 100%;
    max-width: 540px;
    z-index: 1;

    /* centraliza vertical e horizontalmente */
    display: flex;
    justify-content: center;
    align-items: center;
    min-height: 100vh; /* ocupa a altura inteira da tela */
    margin: 0 auto; /* centraliza horizontalmente */
}}
.login-card {{
    background: #ffffff;
    border-radius: 22px;
    padding: 2.75rem 3rem;
    box-shadow: 0 26px 64px rgba(15, 28, 22, 0.22);
    overflow: hidden;
}}
.login-logo {{
    width: 180px;
    margin: 0 auto 1.5rem;
    display: block;
}}
.login-logo-placeholder {{
    width: 100%;
    height: 50px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 1.75rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    color: {PRIMARY_COLOR};
    background: rgba(107, 191, 71, 0.1);
    border-radius: 14px;
}}
.login-title {{
    font-size: 1.9rem;
    font-weight: 700;
    color: {ACCENT_DARK};
    margin-bottom: 0.35rem;
}}
.login-subtitle {{
    font-size: 1.05rem;
    color: {MUTED_TEXT};
    margin-bottom: 1.75rem;
    line-height: 1.6;
}}
.login-highlights {{
    display: grid;
    gap: 0.75rem;
    margin-bottom: 1.5rem;
}}
.highlight-item {{
    display: flex;
    gap: 0.75rem;
    padding: 0.9rem 1rem;
    border-radius: 14px;
    background: rgba(107, 191, 71, 0.1);
    border: 1px solid rgba(82, 181, 75, 0.18);
}}
.highlight-icon {{
    font-size: 1.35rem;
}}
.highlight-text strong {{
    display: block;
    font-size: 0.98rem;
    color: {ACCENT_DARK};
}}
.highlight-text p {{
    margin: 0.15rem 0 0;
    color: {MUTED_TEXT};
    font-size: 0.92rem;
    line-height: 1.45;
}}
.login-alert {{
    margin-bottom: 1.25rem;
    padding: 0.9rem 1rem;
    border-radius: 12px;
    font-weight: 600;
    font-size: 0.95rem;
    border: 1px solid rgba(234, 179, 8, 0.35);
    background: rgba(245, 158, 11, 0.08);
    color: #9a3412;
}}
.login-alert.danger {{
    border-color: rgba(220, 38, 38, 0.4);
    background: rgba(248, 113, 113, 0.12);
    color: #991b1b;
}}
.login-button {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 0.55rem;
    width: 100%;
    margin-top: 0.25rem;
    padding: 1rem 1.25rem;
    border-radius: 14px;
    background: linear-gradient(135deg, {PRIMARY_COLOR}, {SECONDARY_COLOR});
    color: #ffffff;
    font-weight: 700;
    letter-spacing: 0.01em;
    text-decoration: none;
    border: none;
    transition: transform 0.25s ease, box-shadow 0.25s ease;
    box-shadow: 0 18px 36px rgba(82, 181, 75, 0.28);
}}
.login-button:hover {{
    transform: translateY(-2px);
    box-shadow: 0 24px 46px rgba(82, 181, 75, 0.33);
    text-decoration: none;
}}
.login-button-icon {{
    font-size: 1.25rem;
    line-height: 1;
}}
.login-meta {{
    margin-top: 1.75rem;
    text-align: center;
    color: {MUTED_TEXT};
    font-size: 0.94rem;
    line-height: 1.55;
}}
.login-badge {{
    display: inline-block;
    margin-bottom: 0.65rem;
    padding: 0.35rem 0.85rem;
    background: rgba(82, 181, 75, 0.12);
    color: {SECONDARY_COLOR};
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}}
.login-help {{
    margin-top: 1.75rem;
    padding-top: 1.4rem;
    border-top: 1px solid rgba(15, 28, 22, 0.08);
    color: {MUTED_TEXT};
    font-size: 0.9rem;
    line-height: 1.6;
}}
.login-help a {{
    color: {SECONDARY_COLOR};
    font-weight: 600;
    text-decoration: none;
}}
.login-help a:hover {{
    text-decoration: underline;
}}
@media (max-width: 640px) {{
    .login-card {{
        padding: 2.25rem 1.8rem;
        border-radius: 18px;
    }}
    .login-title {{
        font-size: 1.6rem;
    }}
    .login-subtitle {{
        font-size: 1rem;
    }}
}}
</style>
"""


@lru_cache(maxsize=1)
def _get_login_logo_base64() -> Optional[str]:
    """Carrega o logo da Synvia e retorna em base64 para uso inline."""
    logo_dir = Path(__file__).resolve().parent / "logo"
    if not logo_dir.exists():
        logger.warning("Diretório de logos não encontrado para a tela de login")
        return None

    candidates = [
        "SYNVIA_CAEPTOX_LOGO_PADRÃO(nome-preto).png",
        "SYNVIA_CAEPTOX_LOGO_PADRÃO(nome-branco).png",
        "SYNVIA_CAEPTOX_LOGO_PADRÃO(nome-branco).png",
    ]

    for candidate in candidates:
        path_candidate = logo_dir / candidate
        if path_candidate.exists():
            try:
                return base64.b64encode(path_candidate.read_bytes()).decode()
            except Exception as exc:  # pragma: no cover - leitura de arquivo
                logger.warning(f"Falha ao carregar logo {candidate}: {exc}")

    # Fallback: primeiro PNG disponível
    for path_candidate in sorted(logo_dir.glob("*.png")):
        try:
            return base64.b64encode(path_candidate.read_bytes()).decode()
        except Exception as exc:  # pragma: no cover
            logger.warning(f"Falha ao carregar logo {path_candidate.name}: {exc}")

    return None

class MicrosoftAuth:
    """Classe para gerenciar autenticação Microsoft via Azure AD"""

    def __init__(self):
        """Inicializar com configurações do Streamlit secrets"""
        try:
            auth_config = st.secrets.get("auth", {})

            self.client_id = auth_config.get("client_id", os.getenv("AZURE_CLIENT_ID"))
            self.client_secret = auth_config.get("client_secret", os.getenv("AZURE_CLIENT_SECRET"))
            self.tenant_id = auth_config.get("tenant_id", os.getenv("AZURE_TENANT_ID"))
            self.redirect_uri_local = auth_config.get("redirect_uri_local", "http://localhost:8501")
            self.redirect_uri_prod = auth_config.get("redirect_uri_prod", "https://apontamentos.streamlit.app/")
            self.authority = auth_config.get("authority", f"https://login.microsoftonline.com/{self.tenant_id}")
            self.scope = auth_config.get("scope", ["https://graph.microsoft.com/User.Read"])

            # Determinar redirect URI baseado no ambiente
            self.redirect_uri = self._get_redirect_uri()

            # Validar configurações
            if not all([self.client_id, self.client_secret, self.tenant_id]):
                raise ValueError("Configurações de autenticação Microsoft incompletas")

        except Exception as e:
            logger.error(f"Erro ao inicializar MicrosoftAuth: {e}")
            raise

    def _get_redirect_uri(self) -> str:
        """Determinar URI de redirecionamento baseado no ambiente"""
        try:
            # Estratégia simplificada: detectar se estamos no Streamlit Cloud
            # No Streamlit Cloud, várias variáveis específicas estão presentes
            streamlit_env_vars = [
                "STREAMLIT_RUNTIME_VERSION",
                "IS_STREAMLIT_CLOUD",
                "STREAMLIT_SERVER_BASE_URL_PATH"
            ]

            # Verificar se pelo menos uma variável específica do Streamlit Cloud existe
            is_streamlit_cloud = any(os.getenv(var) for var in streamlit_env_vars)

            # Verificar se o hostname sugere produção
            hostname = os.getenv("HOSTNAME", "")
            is_production_hostname = "streamlit" in hostname.lower() or hostname.startswith("pod-")

            # Verificar se estamos rodando em um path que sugere produção
            base_url_path = os.getenv("STREAMLIT_SERVER_BASE_URL_PATH", "")
            is_production_url = "streamlit.app" in base_url_path

            # Combinação final: se qualquer indicador de produção for verdadeiro
            is_production = is_streamlit_cloud or is_production_hostname or is_production_url

            # Log detalhado
            logger.info(f"Detecção produção - Cloud: {is_streamlit_cloud}, Hostname: {hostname}, URL: {base_url_path}")
            logger.info(f"IS_PRODUCTION: {is_production}")

            if is_production:
                # Em produção (Streamlit Cloud), usar URI de produção
                logger.info("Ambiente de PRODUÇÃO detectado - usando redirect_uri_prod")
                return self.redirect_uri_prod
            else:
                # Em desenvolvimento local, usar URI local se disponível
                logger.warning("Ambiente de DESENVOLVIMENTO detectado")
                logger.warning("IMPORTANTE: Adicione 'http://localhost:8501' como Redirect URI no Azure AD")

                # Usar URI local para desenvolvimento
                if hasattr(self, 'redirect_uri_local') and self.redirect_uri_local:
                    logger.info(f"Usando URI local: {self.redirect_uri_local}")
                    return self.redirect_uri_local

                # Fallback para produção (não funcionará completamente, mas pelo menos tentará)
                logger.warning("FALLBACK: usando redirect_uri_prod (adicione localhost no Azure AD)")
                return self.redirect_uri_prod

        except Exception as e:
            logger.error(f"Erro ao determinar redirect URI: {e}")
            return self.redirect_uri_prod

    def get_login_url(self) -> str:
        """Gera URL de autenticação Microsoft"""
        try:
            app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret
            )

            # MSAL automaticamente solicita offline_access quando usado dessa forma
            auth_url = app.get_authorization_request_url(
                self.scope,
                redirect_uri=self.redirect_uri,
                prompt="select_account"  # Permite seleção de conta
            )
            return auth_url
        except Exception as e:
            logger.error(f"Erro ao gerar URL de login: {e}")
            raise

    def get_token_from_code(self, code: str) -> Optional[Dict[str, Any]]:
        """Troca código de autorização por token de acesso e refresh token"""
        try:
            app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret
            )

            # MSAL automaticamente retorna refresh_token quando disponível
            result = app.acquire_token_by_authorization_code(
                code,
                scopes=self.scope,
                redirect_uri=self.redirect_uri
            )

            if "access_token" in result:
                return {
                    "access_token": result["access_token"],
                    "refresh_token": result.get("refresh_token"),
                    "expires_in": result.get("expires_in", 3600)
                }

            if "error" in result:
                logger.error(f"Erro na autenticação: {result['error_description']}")
                return None

            return None

        except Exception as e:
            logger.error(f"Erro ao obter token: {e}")
            return None
    
    def refresh_access_token(self, refresh_token: str) -> Optional[Dict[str, Any]]:
        """Renova o access token usando refresh token"""
        try:
            app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret
            )

            # MSAL automaticamente retorna novo refresh_token
            result = app.acquire_token_by_refresh_token(
                refresh_token,
                scopes=self.scope
            )

            if "access_token" in result:
                logger.info("Token renovado com sucesso")
                return {
                    "access_token": result["access_token"],
                    "refresh_token": result.get("refresh_token", refresh_token),
                    "expires_in": result.get("expires_in", 3600)
                }

            if "error" in result:
                logger.error(f"Erro ao renovar token: {result.get('error_description')}")
                return None

            return None

        except Exception as e:
            logger.error(f"Erro ao renovar token: {e}")
            return None

    def get_user_info(self, token: str) -> Optional[Dict[str, Any]]:
        """Obtém informações do usuário autenticado via Microsoft Graph"""
        try:
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }

            response = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers=headers,
                timeout=10
            )

            if response.status_code == 200:
                user_data = response.json()
                # Adicionar campos calculados
                user_data['domain'] = user_data.get('userPrincipalName', '').split('@')[-1] if user_data.get('userPrincipalName') else ''
                return user_data
            else:
                logger.error(f"Erro ao obter informações do usuário: {response.status_code} - {response.text}")
                return None

        except requests.exceptions.RequestException as e:
            logger.error(f"Erro de rede ao obter informações do usuário: {e}")
            return None
        except Exception as e:
            logger.error(f"Erro inesperado ao obter informações do usuário: {e}")
            return None

    def validate_token(self, token: str) -> bool:
        """Valida se o token ainda é válido"""
        try:
            headers = {"Authorization": f"Bearer {token}"}
            response = requests.get(
                "https://graph.microsoft.com/v1.0/me",
                headers=headers,
                timeout=5
            )
            return response.status_code == 200
        except:
            return False


class AuthManager:
    """Gerenciador de estado de autenticação para Streamlit"""

    @staticmethod
    def init_session_state():
        """Inicializar estado da sessão para autenticação"""
        if "authenticated" not in st.session_state:
            st.session_state.authenticated = False
        if "user_info" not in st.session_state:
            st.session_state.user_info = None
        if "token" not in st.session_state:
            st.session_state.token = None
        if "refresh_token" not in st.session_state:
            st.session_state.refresh_token = None
        if "token_expiry" not in st.session_state:
            st.session_state.token_expiry = None
        if "login_attempts" not in st.session_state:
            st.session_state.login_attempts = 0

    @staticmethod
    def login(user_info: Dict[str, Any], token: str, refresh_token: str = None, expires_in: int = 3600):
        """Realizar login do usuário"""
        import datetime
        st.session_state.authenticated = True
        st.session_state.user_info = user_info
        st.session_state.token = token
        st.session_state.refresh_token = refresh_token
        st.session_state.token_expiry = datetime.datetime.now() + datetime.timedelta(seconds=expires_in)
        st.session_state.login_attempts = 0
        logger.info(f"Usuário {user_info.get('displayName')} fez login com sucesso")

    @staticmethod
    def logout():
        """Realizar logout do usuário"""
        user_name = st.session_state.user_info.get('displayName') if st.session_state.user_info else 'Unknown'
        logger.info(f"Usuário {user_name} fez logout")

        # Limpar estado da sessão
        st.session_state.authenticated = False
        st.session_state.user_info = None
        st.session_state.token = None
        st.session_state.login_attempts = 0

    @staticmethod
    def is_authenticated() -> bool:
        """Verificar se usuário está autenticado"""
        return st.session_state.get("authenticated", False)

    @staticmethod
    def get_current_user() -> Optional[Dict[str, Any]]:
        """Obter informações do usuário atual"""
        return st.session_state.get("user_info")

    @staticmethod
    def require_auth():
        """Exigir autenticação - redirecionar para login se não autenticado"""
        if not AuthManager.is_authenticated():
            st.error("🔐 Acesso negado. Faça login para continuar.")
            st.stop()

    @staticmethod
    def get_token() -> Optional[str]:
        """Obter token atual"""
        return st.session_state.get("token")

    @staticmethod
    def increment_login_attempts():
        """Incrementar contador de tentativas de login"""
        st.session_state.login_attempts = st.session_state.get("login_attempts", 0) + 1

    @staticmethod
    def get_login_attempts() -> int:
        """Obter número de tentativas de login"""
        return st.session_state.get("login_attempts", 0)
    
    @staticmethod
    def check_and_refresh_token(auth: 'MicrosoftAuth') -> bool:
        """
        Verifica se o token está próximo de expirar e renova automaticamente.
        Retorna True se o token está válido (renovado ou ainda válido).
        """
        import datetime
        
        if not AuthManager.is_authenticated():
            return False
        
        # Verificar se temos refresh_token
        refresh_token = st.session_state.get("refresh_token")
        if not refresh_token:
            logger.warning("Sem refresh_token disponível. Usuário precisará fazer login novamente.")
            return True  # Token atual ainda pode estar válido
        
        # Verificar expiração do token
        token_expiry = st.session_state.get("token_expiry")
        if not token_expiry:
            return True  # Sem informação de expiração, assumir válido
        
        # Renovar token se faltar menos de 5 minutos para expirar
        time_until_expiry = token_expiry - datetime.datetime.now()
        if time_until_expiry.total_seconds() < 300:  # 5 minutos
            logger.info(f"Token expira em {time_until_expiry.total_seconds():.0f}s. Renovando...")
            
            # Tentar renovar token
            new_token_data = auth.refresh_access_token(refresh_token)
            if new_token_data:
                # Atualizar session_state com novo token
                st.session_state.token = new_token_data["access_token"]
                st.session_state.refresh_token = new_token_data.get("refresh_token", refresh_token)
                st.session_state.token_expiry = datetime.datetime.now() + datetime.timedelta(
                    seconds=new_token_data.get("expires_in", 3600)
                )
                logger.info("Token renovado com sucesso!")
                return True
            else:
                logger.error("Falha ao renovar token. Usuário precisará fazer login novamente.")
                # Limpar autenticação se não conseguir renovar
                AuthManager.logout()
                return False
        
        return True


def create_login_page(auth: MicrosoftAuth) -> bool:
    """
    Criar página de login Microsoft
    Retorna True se login foi bem-sucedido
    """
    AuthManager.init_session_state()

    # Se já autenticado, não mostrar página de login
    if AuthManager.is_authenticated():
        return True

    st.markdown(LOGIN_PAGE_CSS, unsafe_allow_html=True)

    # Verificar se há retorno de autenticação na URL
    query_params = st.query_params

    if "code" in query_params:
        with st.spinner("🔄 Autenticando..."):
            code = query_params["code"]
            token_data = auth.get_token_from_code(code)

            if token_data and token_data.get("access_token"):
                access_token = token_data["access_token"]
                refresh_token = token_data.get("refresh_token")
                expires_in = token_data.get("expires_in", 3600)

                user_info = auth.get_user_info(access_token)
                if user_info:
                    AuthManager.login(user_info, access_token, refresh_token, expires_in)
                    st.success("✅ Login realizado com sucesso!")
                    st.balloons()
                    st.query_params.clear()
                    return True

                AuthManager.increment_login_attempts()
                st.error("❌ Erro ao obter informações do usuário. Tente novamente em instantes.")
            else:
                AuthManager.increment_login_attempts()
                st.error("❌ Falha na autenticação. Verifique suas credenciais e tente novamente.")

        if AuthManager.get_login_attempts() >= 3:
            st.warning("⚠️ Muitas tentativas falharam. Atualize a página ou entre em contato com o suporte.")

        st.query_params.clear()

    elif "error" in query_params:
        error = query_params.get("error", [""])[0]
        error_description = query_params.get("error_description", ["Erro desconhecido"])[0]
        st.error(f"❌ Erro de autenticação: {error}")
        st.warning(f"Detalhes: {error_description}")
        st.query_params.clear()

    raw_login_url = auth.get_login_url()
    login_url = escape(raw_login_url, quote=True)
    logo_base64 = _get_login_logo_base64()
    if logo_base64:
        logo_html = f'<img src="data:image/png;base64,{logo_base64}" alt="Synvia" class="login-logo" />'
    else:
        logo_html = '<div class="login-logo-placeholder">Bem-vindo!</div>'

    highlights_html = "\n".join([
        '<div class="login-highlights">',
        '<div class="highlight-item">',
        '<div class="highlight-icon">📊</div>',
        '<div class="highlight-text">',
        '<strong>Cadastre, edite e visualise apontamentos!</strong>',
        '<p>Sistema unificado para controle e armazenamentos dos apontamentos.</p>',
        '</div>',
        '</div>',
        '<div class="highlight-item">',
        '<div class="highlight-icon">🔒</div>',
        '<div class="highlight-text">',
        '<strong>Segurança integrada</strong>',
        '<p>Autenticação Microsoft garante sigilo e rastreabilidade de acessos.</p>',
        '</div>',
        '</div>',
        '<div class="highlight-item">',
        '<div class="highlight-icon">🤝</div>',
        '<div class="highlight-text">',
        '<strong>Experiência unificada</strong>',
        '<p>Acesso único aos dados consolidados do ecossistema Synvia.</p>',
        '</div>',
        '</div>',
        '</div>',
    ])

    attempts = AuthManager.get_login_attempts()
    attempts_html = ""
    if attempts >= 3:
        attempts_html = '<div class="login-alert danger">Muitas tentativas foram detectadas. Atualize a página ou procure o time de Data Analytics.</div>'
    elif attempts > 0:
        attempts_html = f'<div class="login-alert">Tentativa {attempts} registrada. Caso o erro persista, limpe o cache do navegador e tente novamente.</div>'

    html_parts = [
        '<div class="login-wrapper">',
        '<div class="login-inner">',
        '<div class="login-card">',
        f'{logo_html}',
        '<h1 class="login-title" style="text-align: center; color: black;">  CONTROLE DE   APONTAMENTOS</h1>',
        '<p class="login-subtitle"></br></p>',
        f'{highlights_html}',
        attempts_html if attempts_html else '',
        f'<a class="login-button" href="{login_url}">',
        '<span>Entrar com Microsoft</span>',
        '<span class="login-button-icon" aria-hidden="true">&rarr;</span>',
        '</a>',
        '<div class="login-meta">',
        '<span class="login-badge">Acesso restrito Synvia</span>',
        '<p>Use sua conta Microsoft corporativa <strong>@synvia.com</strong> para continuar.</p>',
        '</div>',
        '<div class="login-help">',
        '<p>Ao acessar, você concorda em manter a confidencialidade das informações exibidas.</p>',
        '<p>Problemas no login? Libere pop-ups, limpe o cache do navegador ou procure o time de Data Analytics.</p>',
        '</div>',
        '</div>',
        '</div>',
        '</div>',
    ]

    html_content = "\n".join(part for part in html_parts if part)

    st.markdown(html_content, unsafe_allow_html=True)

    return False


def create_user_header():
    """Mostrar informações do usuário e botão de logout na sidebar (discreto)."""
    if not AuthManager.is_authenticated():
        return

    user = AuthManager.get_current_user()
    if not user:
        return

    with st.sidebar:
        display_name = user.get('displayName', 'Usuário')
        email = user.get('mail') or user.get('userPrincipalName', '')

        # Bloco discreto de conta na sidebar
        st.caption("Conta")
        st.markdown(f"👤 {display_name}")
        if email:
            st.caption(f"📧 {email}")

        if st.button("🚪 Logout", key="logout_sidebar", type="secondary", help="Sair da conta"):
            AuthManager.logout()
            st.rerun()


# Função de compatibilidade para código existente
def check_authentication():
    """Função de compatibilidade - verificar se usuário está autenticado"""
    return AuthManager.is_authenticated()


def get_current_user_info():
    """Função de compatibilidade - obter informações do usuário atual"""
    return AuthManager.get_current_user()