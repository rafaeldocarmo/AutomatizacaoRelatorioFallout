# -*- coding: utf-8 -*-
"""
Espelho local da pasta de extrações do Google Drive.

O app continua lendo arquivos do disco: este módulo baixa a pasta do Drive
para `extracoes/_drive/`, com a mesma estrutura de lá, e `fallout_core.raiz_dados`
passa a apontar para ela. Tudo o que já funciona com a pasta local (conversão
dos XLSX, data de corte pelo nome do arquivo, fotografias) segue igual.

Configuração, ao lado do app.py:

    drive_chave.json   a chave JSON da service account, como o Google Cloud
                       entrega. A presença dela é o que liga o Drive.
    drive_pasta.txt    o link (ou o ID) da pasta do Drive. A pasta precisa
                       estar compartilhada com o e-mail da service account
                       ("client_email" da chave), como Leitor.

A sincronização é incremental: só baixa o que mudou desde a última (pelo MD5
que o Drive informa), e apaga do espelho o que saiu do Drive. Cada arquivo
baixado fica com a data de modificação do Drive, porque os caches e a
conversão dos XLSX se guiam por ela.

O espelho persiste entre aberturas: sem internet, o app usa a última cópia.

    python drive_sync.py        sincroniza e mostra o que mudou
"""

import contextlib
import datetime as _dt
import glob
import json
import os
import re
import sys
import threading
from dataclasses import dataclass, field

import fallout_core

ARQ_PASTA = "drive_pasta.txt"
ARQ_MANIFESTO = ".manifesto.json"
ESCOPO = "https://www.googleapis.com/auth/drive.readonly"
API = "https://www.googleapis.com/drive/v3/files"

MIME_PASTA = "application/vnd.google-apps.folder"
MIME_ATALHO = "application/vnd.google-apps.shortcut"
# Planilha nativa do Google não tem bytes para baixar: sai exportada em XLSX.
MIME_PLANILHA = "application/vnd.google-apps.spreadsheet"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Pastas do Drive que não descem: as fotografias são gravadas no computador
# (extracoes/_fotografias), e uma cópia delas no Drive só pesaria no download.
IGNORAR_PASTAS = {"_fotografias"}

_TRAVA = threading.Lock()


@dataclass
class Resultado:
    quando: _dt.datetime
    pasta: str = ""
    total: int = 0
    baixados: list = field(default_factory=list)
    removidos: list = field(default_factory=list)
    falhas: list = field(default_factory=list)

    @property
    def mudou(self):
        return bool(self.baixados or self.removidos)


class ErroDrive(Exception):
    """Falha com mensagem pronta para a tela."""


def configurado(base_dir):
    return fallout_core.drive_disponivel(base_dir)


def pasta_espelho(base_dir):
    return os.path.join(base_dir, fallout_core.PASTA_DADOS, fallout_core.PASTA_DRIVE)


def id_da_pasta(texto):
    """ID da pasta a partir do link do Drive (…/folders/<id>, …?id=<id>) ou dele mesmo."""
    texto = (texto or "").strip()
    for padrao in (r"/folders/([\w-]+)", r"[?&]id=([\w-]+)", r"^([\w-]{10,})$"):
        achado = re.search(padrao, texto)
        if achado:
            return achado.group(1)
    return None


def _ler_pasta(base_dir):
    caminho = os.path.join(base_dir, ARQ_PASTA)
    if not os.path.isfile(caminho):
        raise ErroDrive(f"Falta o `{ARQ_PASTA}` ao lado do app, com o link da pasta do Drive.")
    with open(caminho, encoding="utf-8-sig") as f:
        linhas = [l.strip() for l in f if l.strip() and not l.lstrip().startswith("#")]
    pasta = id_da_pasta(linhas[0]) if linhas else None
    if not pasta:
        raise ErroDrive(f"O `{ARQ_PASTA}` não tem um link de pasta do Drive reconhecível.")
    return pasta


def _sessao(base_dir):
    """Sessão HTTP autenticada com a service account, mais o e-mail dela."""
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2 import service_account

    chave = os.path.join(base_dir, fallout_core.ARQ_CHAVE_DRIVE)
    try:
        with open(chave, encoding="utf-8") as f:
            email = json.load(f).get("client_email", "")
        cred = service_account.Credentials.from_service_account_file(chave, scopes=[ESCOPO])
    except (OSError, ValueError, KeyError) as e:
        raise ErroDrive(f"A chave `{fallout_core.ARQ_CHAVE_DRIVE}` não é uma chave de "
                        f"service account válida ({e}).") from e
    return AuthorizedSession(cred), email


def _get(sessao, url, email, **params):
    params.setdefault("supportsAllDrives", "true")
    resp = sessao.get(url, params=params, timeout=120, stream=True)
    if resp.status_code == 404:
        raise ErroDrive("A pasta (ou um arquivo dela) não foi encontrada. Confira o link e "
                        f"se ela está compartilhada com {email}.")
    if resp.status_code == 403:
        raise ErroDrive(f"O Drive recusou o acesso de {email}: a Google Drive API precisa "
                        "estar ativada no projeto da service account, e a pasta "
                        f"compartilhada com ela. Resposta: {resp.text[:300]}")
    resp.raise_for_status()
    return resp


def _nome_local(nome):
    """O nome do Drive aceito pelo Windows: sem \\ / : * ? " < > | nem ponto/espaço no fim."""
    return re.sub(r'[\\/:*?"<>|]', "_", nome).rstrip(". ") or "_"


def _listar(sessao, email, pasta_id, prefixo=""):
    """
    {caminho relativo: metadados} de todos os arquivos abaixo de `pasta_id`.
    Atalhos valem pelo que apontam; nomes repetidos na mesma pasta ficam com o
    mais recente, como o Drive mostra no topo.
    """
    arquivos, token = {}, None
    while True:
        params = {"q": f"'{pasta_id}' in parents and trashed=false", "pageSize": 1000,
                  "includeItemsFromAllDrives": "true",
                  "fields": "nextPageToken, files(id, name, mimeType, md5Checksum, size, "
                            "modifiedTime, shortcutDetails)"}
        if token:
            params["pageToken"] = token
        dados = _get(sessao, API, email, **params).json()
        for item in dados.get("files", []):
            mime, alvo = item["mimeType"], item["id"]
            if mime == MIME_ATALHO:
                det = item.get("shortcutDetails") or {}
                mime, alvo = det.get("targetMimeType", ""), det.get("targetId", "")
                if mime != MIME_PASTA:
                    item = _get(sessao, f"{API}/{alvo}", email,
                                fields="id, name, mimeType, md5Checksum, size, modifiedTime"
                                ).json() | {"name": item["name"]}
            rel = prefixo + _nome_local(item["name"])
            if mime == MIME_PASTA:
                if item["name"] not in IGNORAR_PASTAS:
                    arquivos.update(_listar(sessao, email, alvo, rel + "/"))
                continue
            if mime == MIME_PLANILHA and not rel.lower().endswith(".xlsx"):
                rel += ".xlsx"
            meta = {"id": item["id"], "mime": mime, "md5": item.get("md5Checksum"),
                    "modificado": item.get("modifiedTime"), "tamanho": item.get("size")}
            atual = arquivos.get(rel)
            if atual is None or (meta["modificado"] or "") > (atual["modificado"] or ""):
                arquivos[rel] = meta
        token = dados.get("nextPageToken")
        if not token:
            return arquivos


def _versao(meta):
    """O que identifica o conteúdo: o MD5, ou a data para o que o Drive não dá MD5."""
    return meta.get("md5") or meta.get("modificado")


def _baixar(sessao, email, meta, destino):
    """Baixa para um temporário e troca de uma vez, com a data do Drive."""
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    if meta["mime"] == MIME_PLANILHA:
        resp = _get(sessao, f"{API}/{meta['id']}/export", email, mimeType=MIME_XLSX,
                    supportsAllDrives=None)
    else:
        resp = _get(sessao, f"{API}/{meta['id']}", email, alt="media")
    parcial = destino + ".drive-parcial"
    try:
        with open(parcial, "wb") as f:
            for bloco in resp.iter_content(1 << 20):
                f.write(bloco)
        if meta.get("modificado"):
            ts = _dt.datetime.fromisoformat(meta["modificado"].replace("Z", "+00:00")).timestamp()
            os.utime(parcial, (ts, ts))
        os.replace(parcial, destino)
    finally:
        with contextlib.suppress(OSError):
            os.remove(parcial)


def _remover(destino):
    """Apaga o arquivo e os CSVs que o app gerou a partir dele (<nome>__*.csv)."""
    for caminho in [destino] + glob.glob(glob.escape(os.path.splitext(destino)[0]) + "__*.csv"):
        with contextlib.suppress(FileNotFoundError):
            os.remove(caminho)


def _limpar_pastas_vazias(raiz):
    for pasta, _subs, _arqs in sorted(os.walk(raiz), key=lambda t: -len(t[0])):
        if pasta != raiz:
            with contextlib.suppress(OSError):
                os.rmdir(pasta)          # só sai se estiver vazia


def sincronizar(base_dir, progresso=None):
    """
    Deixa `extracoes/_drive` igual à pasta do Drive e devolve um `Resultado`.
    `progresso(fração, texto)` é chamado a cada arquivo baixado.

    Levanta `ErroDrive` quando não dá para listar a pasta; falha num arquivo
    só vai para `Resultado.falhas` e o resto segue.
    """
    with _TRAVA:
        pasta_id = _ler_pasta(base_dir)
        sessao, email = _sessao(base_dir)
        raiz = _get(sessao, f"{API}/{pasta_id}", email, fields="id, name, mimeType").json()
        if raiz.get("mimeType") != MIME_PASTA:
            raise ErroDrive(f"O link do `{ARQ_PASTA}` não é de uma pasta do Drive.")
        remotos = _listar(sessao, email, pasta_id)

        espelho = pasta_espelho(base_dir)
        os.makedirs(espelho, exist_ok=True)
        arq_manifesto = os.path.join(espelho, ARQ_MANIFESTO)
        try:
            with open(arq_manifesto, encoding="utf-8") as f:
                manifesto = json.load(f)
        except (OSError, ValueError):
            manifesto = {}

        res = Resultado(quando=_dt.datetime.now(), pasta=raiz.get("name", ""), total=len(remotos))
        for rel in sorted(set(manifesto) - set(remotos)):
            _remover(os.path.join(espelho, *rel.split("/")))
            manifesto.pop(rel)
            res.removidos.append(rel)

        novos = [rel for rel, meta in sorted(remotos.items())
                 if _versao(manifesto.get(rel, {})) != _versao(meta)
                 or not os.path.isfile(os.path.join(espelho, *rel.split("/")))]
        for i, rel in enumerate(novos):
            if progresso:
                progresso(i / len(novos), f"Baixando do Drive: {rel}")
            try:
                _baixar(sessao, email, remotos[rel], os.path.join(espelho, *rel.split("/")))
            except Exception as e:                       # noqa: BLE001
                res.falhas.append(f"{rel} ({type(e).__name__}: {e})")
                continue
            manifesto[rel] = remotos[rel]
            res.baixados.append(rel)
        if progresso and novos:
            progresso(1.0, "Drive sincronizado")

        with open(arq_manifesto, "w", encoding="utf-8") as f:
            json.dump(manifesto, f, ensure_ascii=False, indent=1)
        _limpar_pastas_vazias(espelho)
        return res


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    base = os.path.dirname(os.path.abspath(__file__))
    if not configurado(base):
        sys.exit(f"Sem {fallout_core.ARQ_CHAVE_DRIVE} ao lado do app: o Drive está desligado.")
    try:
        r = sincronizar(base, lambda _f, t: print(" ", t))
    except ErroDrive as e:
        sys.exit(str(e))
    print(f"Pasta '{r.pasta}': {r.total} arquivos, {len(r.baixados)} baixados, "
          f"{len(r.removidos)} removidos, {len(r.falhas)} com falha.")
    for linha in r.removidos:
        print("  removido:", linha)
    for linha in r.falhas:
        print("  FALHA:", linha)
