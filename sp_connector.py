# sp_connector.py
import io, time, requests, msal, pandas as pd
from urllib.parse import quote

GRAPH = "https://graph.microsoft.com/v1.0"

class SPConnector:
    """
    Conecta no SharePoint/OneDrive via Microsoft Graph (app-only).
    Suporta:
      - SharePoint Site: hostname + site_path + library_name
      - OneDrive do usuário: user_upn (ex: "susanna.bernardes@synvia.com")
    Caminho do arquivo:
      - OneDrive: RELATIVO a Documents (ex: "Pasta/arquivo.xlsx")
        (aceita tb /personal/<upn>/Documents/... que será normalizado)
      - SharePoint: RELATIVO à biblioteca (ex: "Pasta/arquivo.xlsx")
        (aceita tb server-relative /sites/<site>/<lib>/... que será normalizado)
    """

    def __init__(self, tenant_id, client_id, client_secret,
                 hostname=None, site_path=None, library_name=None, user_upn=None):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret

        self.hostname = hostname or ""
        self.site_path = site_path or ""
        self.library_name = library_name or ""
        self.user_upn = user_upn or ""          # se presente, opera em OneDrive

        self._app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )
        self._tok = None
        self._exp = 0
        self._site_id_cache = None
        self._drive_id_cache = None

    # -------- Auth --------
    def _token(self):
        now = time.time()
        if self._tok and now < self._exp:
            return self._tok
        res = self._app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in res:
            raise RuntimeError(res.get("error_description") or res)
        self._tok = res["access_token"]
        self._exp = now + int(res.get("expires_in", 3600)) - 60
        return self._tok

    def _headers(self):
        return {"Authorization": f"Bearer {self._token()}"}

    # -------- Modo --------
    @property
    def is_onedrive(self) -> bool:
        return bool(self.user_upn)

    # -------- Descoberta (apenas p/ SharePoint Site) --------
    def _site_id(self):
        if self.is_onedrive:
            return None
        if self._site_id_cache:
            return self._site_id_cache
        url = f"{GRAPH}/sites/{self.hostname}:/{self.site_path}"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        self._site_id_cache = r.json()["id"]
        return self._site_id_cache

    def _drive_id(self):
        if self.is_onedrive:
            return None
        if self._drive_id_cache:
            return self._drive_id_cache
        url = f"{GRAPH}/sites/{self._site_id()}/drives"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        drives = r.json().get("value", [])
        for d in drives:
            if d.get("name", "").lower() == self.library_name.lower():
                self._drive_id_cache = d["id"]
                return self._drive_id_cache
        for d in drives:
            if d.get("driveType") == "documentLibrary":
                self._drive_id_cache = d["id"]
                return self._drive_id_cache
        raise RuntimeError(f"Biblioteca '{self.library_name}' não encontrada em {self.site_path}")

    # -------- Normalização de caminho --------
    def normalize_path(self, path: str) -> str:
        """
        Retorna o caminho relativo correto ao "root" usado em cada modo:
          - OneDrive: relativo a Documents/
          - SharePoint: relativo à biblioteca
        Aceita caminhos server-relative e normaliza.
        """
        if not path:
            raise ValueError("Caminho vazio.")
        path = path.strip()

        if self.is_onedrive:
            # Aceita: "Pasta/arquivo.xlsx" OU "/personal/<upn>/Documents/Pasta/arquivo.xlsx"
            if path.startswith("/"):
                marker = "/Documents/"
                idx = path.lower().find(marker.lower())
                if idx == -1:
                    raise ValueError("Para OneDrive, o server-relative precisa conter /Documents/.")
                return path[idx + len(marker):]
            return path
        else:
            # SharePoint site
            if path.startswith("/"):
                prefix = f"/{self.site_path}/{self.library_name}/"
                if not path.startswith(prefix):
                    raise ValueError(
                        "file_path server-relative não bate com site/biblioteca dos secrets.\n"
                        f"Esperado prefixo: {prefix}\nRecebido: {path}"
                    )
                return path[len(prefix):]
            return path

    # -------- Download / Upload --------
    def download(self, path: str) -> bytes:
        rel = quote(self.normalize_path(path), safe="/")
        if self.is_onedrive:
            url = f"{GRAPH}/users/{self.user_upn}/drive/root:/{rel}:/content"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{rel}:/content"
        r = requests.get(url, headers=self._headers(), timeout=180)
        if r.status_code == 404:
            raise FileNotFoundError(path)
        r.raise_for_status()
        return r.content

    def upload_small(self, path: str, content: bytes, overwrite: bool = True):
        rel = quote(self.normalize_path(path), safe="/")
        params = {"@microsoft.graph.conflictBehavior": "replace" if overwrite else "fail"}
        if self.is_onedrive:
            url = f"{GRAPH}/users/{self.user_upn}/drive/root:/{rel}:/content"
        else:
            url = f"{GRAPH}/drives/{self._drive_id()}/root:/{rel}:/content"
        r = requests.put(url, headers=self._headers(), params=params, data=content, timeout=300)
        r.raise_for_status()
        return r.json()

    # -------- Conveniências DataFrame --------
    def read_excel(self, path: str, **kw) -> pd.DataFrame:
        return pd.read_excel(io.BytesIO(self.download(path)), **kw)

    def read_csv(self, path: str, **kw) -> pd.DataFrame:
        return pd.read_csv(io.BytesIO(self.download(path)), **kw)

    def write_excel(self, df: pd.DataFrame, path: str, overwrite: bool = True):
        bio = io.BytesIO()
        df.to_excel(bio, index=False)
        return self.upload_small(path, bio.getvalue(), overwrite=overwrite)
