# -*- coding: utf-8 -*-
"""
Monta a versão portátil do Explorador de Fallout em dist/.

Gera uma pasta que roda em qualquer Windows x64 sem Python instalado:
baixa o Python "embeddable" oficial, habilita o pip nele, instala as
dependências e copia o app junto com um atalho clicável.

Uso:  python build_portatil.py
"""
import os
import shutil
import subprocess
import sys
import urllib.request
import zipfile

# No Windows o console usa cp1252 e quebra nos prints com "→"/"✓".
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

PY_VER = "3.13.3"
PY_ZIP = f"python-{PY_VER}-embed-amd64.zip"
PY_URL = f"https://www.python.org/ftp/python/{PY_VER}/{PY_ZIP}"
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"

RAIZ = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(RAIZ, "dist", "ExploradorDeFallout")
PY_DIR = os.path.join(DIST, "python")
CACHE = os.path.join(RAIZ, "dist", "_cache")

FONTES = ["app.py", "fallout_core.py", "pipeline.py",
          "gerar_pptx.py", "gerar_slide3col.py", "launcher.py"]

# Versões travadas: um app distribuído tem que sair igual a cada build. O
# streamlit em especial não pode flutuar — o CSS do app depende do DOM interno
# dele (ver comentário no requirements.txt).
DEPS = ["streamlit==1.58.0", "pandas==2.2.3", "numpy==2.1.3",
        "matplotlib==3.10.9", "openpyxl==3.1.5", "python-pptx==1.0.2",
        "pywebview==6.2.1"]

ATALHO = r"""@echo off
rem Abre o Explorador de Fallout em janela propria, sem console.
start "" "%~dp0python\pythonw.exe" "%~dp0launcher.py"
"""

# Windows marca arquivos vindos de outra maquina (Mark of the Web) e o .NET
# recusa DLLs nesse estado — e ai a janela nativa nao sobe. Unblock-File
# resolve e nao precisa de admin.
DESBLOQUEAR = r"""@echo off
echo Liberando os arquivos copiados de outra maquina...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-ChildItem -LiteralPath '%~dp0.' -Recurse -File | Unblock-File"
echo.
echo Pronto. Abra o "Explorador de Fallout".
pause
"""

# Plano B sem pywebview: Streamlit no navegador, com console visivel para
# poder encerrar fechando a janela preta.
NAVEGADOR = r"""@echo off
title Explorador de Fallout (navegador) - feche esta janela para encerrar
"%~dp0python\python.exe" -m streamlit run "%~dp0app.py" --server.address=127.0.0.1 --browser.gatherUsageStats=false
pause
"""

LEIAME = """Explorador de Fallout — versão portátil
=======================================

Como usar
---------
Dê um duplo clique em "Explorador de Fallout.bat".

Não é preciso ter Python instalado: ele já vai junto, dentro da pasta
"python". A primeira abertura leva uns 10 segundos.

Onde ficam os dados
-------------------
Na pasta "extracoes", ao lado do atalho. A estrutura é:

    extracoes/
      RelatorioDFTOctane.xlsx
      RelatorioUSOctane.xlsx
      <Nome da Jornada>/
        falhas/     <- CSVs de falhas, um por mês
        sucessos/   <- CSVs de sucessos, um por mês

Para atualizar o relatório, troque os arquivos dessa pasta e clique em
"Atualizar dados" dentro do app.

Para acrescentar uma jornada nova, basta criar a pasta dela com as
subpastas "falhas" e "sucessos" — o app descobre sozinho.

Se não abrir
------------
1. Rode "Desbloquear arquivos.bat" e tente de novo.

   O Windows marca como bloqueado tudo que vem de outra máquina (zip, rede,
   download) e o .NET recusa carregar as bibliotecas nesse estado — é a causa
   mais comum da janela não abrir. Não precisa de administrador.

2. Se ainda assim não abrir, use "Abrir no navegador.bat". O app é o mesmo;
   só roda em aba do navegador em vez de janela própria. Para encerrar,
   feche a janela preta do console.

3. O arquivo "launcher.log", ao lado do atalho, registra o que aconteceu.
"""


def _baixar(url, destino):
    if os.path.exists(destino):
        print(f"  (cache) {os.path.basename(destino)}")
        return destino
    print(f"  baixando {url}")
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    urllib.request.urlretrieve(url, destino)
    return destino


def _py(*args):
    """Roda o Python embutido da dist."""
    exe = os.path.join(PY_DIR, "python.exe")
    subprocess.run([exe, *args], check=True)


def main():
    print(f"→ destino: {DIST}")
    if os.path.exists(DIST):
        print("  limpando build anterior")
        shutil.rmtree(DIST)
    os.makedirs(PY_DIR, exist_ok=True)

    print("\n[1/5] Python embeddable")
    zip_local = _baixar(PY_URL, os.path.join(CACHE, PY_ZIP))
    with zipfile.ZipFile(zip_local) as z:
        z.extractall(PY_DIR)

    # O embeddable vem com site-packages desligado; sem isso o pip instala
    # mas nada é importável.
    pth = next(f for f in os.listdir(PY_DIR) if f.endswith("._pth"))
    caminho_pth = os.path.join(PY_DIR, pth)
    with open(caminho_pth, encoding="utf-8") as f:
        linhas = f.read().splitlines()
    linhas = [("import site" if l.strip() == "#import site" else l) for l in linhas]
    if "Lib\\site-packages" not in linhas:
        linhas.insert(len(linhas) - 1, "Lib\\site-packages")
    with open(caminho_pth, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas) + "\n")
    print(f"  site-packages habilitado em {pth}")

    print("\n[2/5] pip")
    get_pip = _baixar(GET_PIP_URL, os.path.join(CACHE, "get-pip.py"))
    _py(get_pip, "--no-warn-script-location")

    print("\n[3/5] dependências (demora alguns minutos)")
    _py("-m", "pip", "install", "--no-warn-script-location", *DEPS)

    print("\n[4/5] aplicação")
    for nome in FONTES:
        shutil.copy2(os.path.join(RAIZ, nome), os.path.join(DIST, nome))
        print(f"  {nome}")

    origem_dados = os.path.join(RAIZ, "extracoes")
    destino_dados = os.path.join(DIST, "extracoes")
    if os.path.isdir(origem_dados):
        shutil.copytree(origem_dados, destino_dados)
        print("  extracoes/ (dados atuais)")
    else:
        os.makedirs(destino_dados, exist_ok=True)
        print("  extracoes/ (vazia — copie os dados para cá)")

    print("\n[5/5] atalhos e leia-me")
    # cp1252: os .bat são lidos pelo cmd.exe, não por Python.
    for nome, conteudo in [
        ("Explorador de Fallout.bat", ATALHO),
        ("Desbloquear arquivos.bat", DESBLOQUEAR),
        ("Abrir no navegador.bat", NAVEGADOR),
    ]:
        with open(os.path.join(DIST, nome), "w", encoding="cp1252") as f:
            f.write(conteudo)
        print(f"  {nome}")
    with open(os.path.join(DIST, "LEIA-ME.txt"), "w", encoding="utf-8") as f:
        f.write(LEIAME)

    tam = sum(
        os.path.getsize(os.path.join(r, a))
        for r, _, arqs in os.walk(DIST) for a in arqs
    )
    print(f"\n✓ pronto: {DIST}")
    print(f"  {tam / 1024 / 1024:,.0f} MB")
    print('  teste com um duplo clique em "Explorador de Fallout.bat"')


if __name__ == "__main__":
    if os.name != "nt":
        sys.exit("Este build é específico para Windows.")
    main()
