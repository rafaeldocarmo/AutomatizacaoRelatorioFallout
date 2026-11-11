# -*- coding: utf-8 -*-
"""
Fonte do "Explorador de Fallout.exe", o atalho clicável da versão portátil.

Este arquivo NÃO empacota o aplicativo: ele só dispara o Python que já vai
dentro da pasta distribuída. É de propósito — congelar o Streamlit inteiro num
executável custaria caro e daria pouco:

  * `app.py` acha a pasta `extracoes` por `__file__`; dentro de um bundle isso
    aponta para o diretório temporário de extração, que some ao fechar;
  * o Streamlit carrega o `app.py` como script e lê os próprios metadados de
    pacote, então congelá-lo exige runtime hooks que quebram a cada versão;
  * atualizar o programa passaria a ser reenviar centenas de MB em vez de
    trocar alguns arquivos `.py`.

Assim o executável fica com poucos MB, abre instantaneamente e o Python que
roda de fato é o oficial, assinado pela PSF — o que importa em máquina
corporativa com antivírus apertado.

Compilado pelo build_portatil.py; veja _compilar_atalho() lá.
"""
import ctypes
import os
import subprocess
import sys

PYTHONW = os.path.join("python", "pythonw.exe")
LAUNCHER = "launcher.py"

AVISO_FALTANDO = (
    "Não encontrei os arquivos do programa ao lado deste atalho:\n\n"
    "    {faltando}\n\n"
    'O atalho precisa ficar dentro da pasta "ExploradorDeFallout", junto com a '
    'pasta "python". Se você copiou só o atalho, copie a pasta inteira.'
)

AVISO_FALHA = (
    "O atalho não conseguiu iniciar o programa.\n\n"
    "{erro}\n\n"
    'Tente abrir pela pasta "Se nao abrir", que fica ao lado deste atalho.'
)


def _erro(msg):
    """Caixa de erro do Windows — sem console, é o único canal de aviso."""
    ctypes.windll.user32.MessageBoxW(None, msg, "Explorador de Fallout", 0x10)


def _raiz():
    """Pasta que contém o atalho: congelado é a do .exe, solto é a do .py."""
    referencia = sys.executable if getattr(sys, "frozen", False) else __file__
    return os.path.dirname(os.path.abspath(referencia))


def main():
    raiz = _raiz()
    py = os.path.join(raiz, PYTHONW)
    app = os.path.join(raiz, LAUNCHER)

    faltando = [c for c in (py, app) if not os.path.exists(c)]
    if faltando:
        _erro(AVISO_FALTANDO.format(
            faltando="\n    ".join(os.path.relpath(c, raiz) for c in faltando)
        ))
        return 1

    try:
        # DETACHED_PROCESS: o atalho sai na hora e quem segura o aplicativo de
        # pé é o launcher. DEVNULL nos três canais porque um .exe sem console
        # não tem handles válidos para repassar ao filho.
        subprocess.Popen(
            [py, app], cwd=raiz,
            creationflags=subprocess.DETACHED_PROCESS,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
    except OSError as e:
        _erro(AVISO_FALHA.format(erro=f"{type(e).__name__}: {e}"))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
