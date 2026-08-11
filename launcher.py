# -*- coding: utf-8 -*-
"""
Abre o Explorador de Fallout em uma janela própria, sem passar pelo navegador.

Sobe o Streamlit em modo headless numa porta livre de 127.0.0.1, espera o
servidor responder e então mostra a janela (WebView2 no Windows). Ao fechar a
janela, o servidor é encerrado junto.

Uso:  pythonw.exe launcher.py
"""
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request

import webview

APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(APP_DIR, "app.py")
TITULO = "Explorador de Fallout"
LOG = os.path.join(APP_DIR, "launcher.log")

# Sem isso o WebView2 engole os downloads e os botões de PowerPoint/Excel
# não fazem nada.
webview.settings["ALLOW_DOWNLOADS"] = True


def _log(msg):
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}  {msg}\n")


def _porta_livre():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _subir_streamlit(porta):
    """Inicia o servidor Streamlit como processo filho, sem janela de console."""
    cmd = [
        sys.executable, "-m", "streamlit", "run", APP,
        "--server.headless=true",
        "--server.address=127.0.0.1",
        f"--server.port={porta}",
        "--server.fileWatcherType=none",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]
    # pythonw não tem console; sem isso o filho abriria uma janela preta.
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    _log(f"iniciando: {' '.join(cmd)}")
    return subprocess.Popen(
        cmd, cwd=APP_DIR, creationflags=flags,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
    )


def _esperar(porta, proc, timeout=120):
    """Aguarda o health check do Streamlit responder."""
    url = f"http://127.0.0.1:{porta}/_stcore/health"
    limite = time.time() + timeout
    while time.time() < limite:
        if proc.poll() is not None:
            raise RuntimeError(f"o servidor encerrou sozinho (código {proc.returncode})")
        try:
            with urllib.request.urlopen(url, timeout=2) as r:
                if r.status == 200:
                    return True
        except (urllib.error.URLError, OSError):
            time.sleep(0.4)
    raise TimeoutError(f"servidor não respondeu em {timeout}s")


def main():
    proc = None
    try:
        porta = _porta_livre()
        proc = _subir_streamlit(porta)
        _esperar(porta, proc)
        _log(f"servidor pronto em 127.0.0.1:{porta}")

        janela = webview.create_window(
            TITULO, f"http://127.0.0.1:{porta}",
            width=1500, height=950, min_size=(1100, 700),
        )

        def _encerrar():
            _log("janela fechada, encerrando o servidor")
            if proc and proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()

        janela.events.closed += _encerrar
        webview.start()
    except Exception as e:                      # noqa: BLE001 - vai para o log
        _log(f"ERRO: {type(e).__name__}: {e}")
        if proc and proc.poll() is None:
            proc.kill()
        # Sem console (pythonw), a janela é a única forma de avisar.
        try:
            webview.create_window(
                f"{TITULO} — erro",
                html=f"<body style='font-family:sans-serif;padding:24px'>"
                     f"<h2>Não foi possível iniciar</h2><p>{e}</p>"
                     f"<p style='color:#666'>Detalhes em launcher.log</p></body>",
                width=640, height=320,
            )
            webview.start()
        except Exception:
            pass
        return 1
    finally:
        if proc and proc.poll() is None:
            proc.kill()
    return 0


if __name__ == "__main__":
    sys.exit(main())
