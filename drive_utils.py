"""
Utilitários para leitura de arquivos do Google Drive via Service Account.

Estrutura esperada no Drive (dentro da pasta raiz GDRIVE_FOLDER_ID):
  Base Móvel/
    falhas/    ← CSVs de falhas
    sucessos/  ← CSVs de sucessos
  Cross Sell/
    falhas/
    sucessos/
  RelatorioDFTOctane.xlsx
  RelatorioUSOctane.xlsx

Variáveis de ambiente / secrets necessários:
  GDRIVE_FOLDER_ID      → ID da pasta raiz no Drive
  GDRIVE_TYPE           → "service_account"
  GDRIVE_PROJECT_ID
  GDRIVE_PRIVATE_KEY_ID
  GDRIVE_PRIVATE_KEY    → chave privada (com \n reais)
  GDRIVE_CLIENT_EMAIL
  GDRIVE_TOKEN_URI      → "https://oauth2.googleapis.com/token"
"""

import io, os, tempfile
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


def _creds():
    """Monta credenciais a partir dos secrets do Streamlit ou variáveis de ambiente."""
    def _get(key):
        try:
            return st.secrets["google"][key]
        except Exception:
            return os.environ.get(f"GDRIVE_{key.upper()}", "")

    info = {
        "type":                        _get("type"),
        "project_id":                  _get("project_id"),
        "private_key_id":              _get("private_key_id"),
        "private_key":                 _get("private_key").replace("\\n", "\n"),
        "client_email":                _get("client_email"),
        "token_uri":                   _get("token_uri") or "https://oauth2.googleapis.com/token",
        "auth_uri":                    "https://accounts.google.com/o/oauth2/auth",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url":        "",
    }
    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)


def _service():
    return build("drive", "v3", credentials=_creds(), cache_discovery=False)


def drive_configurado():
    """Retorna True se as credenciais do Drive estão presentes."""
    try:
        key = st.secrets["google"].get("client_email", "")
    except Exception:
        key = os.environ.get("GDRIVE_CLIENT_EMAIL", "")
    return bool(key)


def _listar_pasta(service, parent_id):
    """Retorna dict {nome: id} dos itens diretos de uma pasta."""
    result = service.files().list(
        q=f"'{parent_id}' in parents and trashed=false",
        fields="files(id, name, mimeType)",
        pageSize=200,
    ).execute()
    return {f["name"]: f for f in result.get("files", [])}


def _download_bytes(service, file_id):
    buf = io.BytesIO()
    req = service.files().get_media(fileId=file_id)
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf


def _baixar_pasta_para_temp(service, pasta_id, dest_dir):
    """Baixa todos os arquivos de uma pasta Drive para dest_dir."""
    os.makedirs(dest_dir, exist_ok=True)
    itens = _listar_pasta(service, pasta_id)
    caminhos = []
    for nome, meta in itens.items():
        if meta["mimeType"] == "application/vnd.google-apps.folder":
            continue
        dest = os.path.join(dest_dir, nome)
        if not os.path.exists(dest):
            dados = _download_bytes(service, meta["id"])
            with open(dest, "wb") as f:
                f.write(dados.read())
        caminhos.append(dest)
    return caminhos


@st.cache_data(show_spinner="Baixando arquivos do Google Drive...", ttl=1800)
def baixar_arquivos_drive():
    """
    Baixa todos os arquivos necessários do Drive para uma pasta temporária.
    Retorna o caminho base (equivalente ao diretório do projeto local).
    """
    svc      = _service()
    root_id  = st.secrets["google"].get("folder_id") or os.environ.get("GDRIVE_FOLDER_ID", "")
    if not root_id:
        st.error("`folder_id` não configurado nos secrets. Adicione o ID da pasta raiz do Drive.")
        st.stop()

    raiz = _listar_pasta(svc, root_id)
    if not raiz:
        st.error(
            f"A pasta do Drive (ID `{root_id}`) está vazia ou não foi compartilhada com a "
            f"service account. Compartilhe-a com o e-mail da service account (permissão Viewer)."
        )
        st.stop()

    _esperado = {"Base Móvel", "Cross Sell"}
    if not (_esperado & set(raiz.keys())):
        st.error(
            "Nenhuma pasta de jornada encontrada no Drive. "
            f"Esperado: {sorted(_esperado)}. Encontrado na pasta raiz: {sorted(raiz.keys())}"
        )
        st.stop()

    tmp = tempfile.mkdtemp(prefix="mayh_drive_")

    # Excel do Octane
    for nome in ["RelatorioDFTOctane.xlsx", "RelatorioUSOctane.xlsx"]:
        if nome in raiz:
            dest = os.path.join(tmp, nome)
            if not os.path.exists(dest):
                dados = _download_bytes(svc, raiz[nome]["id"])
                with open(dest, "wb") as f:
                    f.write(dados.read())

    # Pastas de jornadas
    for jornada in ["Base Móvel", "Cross Sell"]:
        if jornada not in raiz:
            continue
        jornada_id = raiz[jornada]["id"]
        sub = _listar_pasta(svc, jornada_id)
        for subpasta in ["falhas", "sucessos"]:
            if subpasta not in sub:
                continue
            dest_dir = os.path.join(tmp, "extrações", jornada, subpasta)
            _baixar_pasta_para_temp(svc, sub[subpasta]["id"], dest_dir)

    return tmp
