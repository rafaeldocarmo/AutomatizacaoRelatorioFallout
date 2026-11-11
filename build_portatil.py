# -*- coding: utf-8 -*-
"""
Monta a versão portátil do Explorador de Fallout em dist/.

Gera uma pasta que roda em qualquer Windows x64 sem Python instalado e sem
direitos de administrador: baixa o Python "embeddable" oficial, habilita o pip
nele, instala as dependências e copia o app junto com um atalho clicável.

O atalho é um .exe pequeno (fonte em atalho.py) que só dispara o Python de
dentro da pasta — o aplicativo em si continua sendo os .py soltos, então
atualizar é trocar arquivo, não recompilar. Compilar o atalho pede o
PyInstaller na máquina de build (`pip install pyinstaller`); sem ele o build
segue e a pasta sai só com os .bat.

A pasta sai com a "extracoes" VAZIA. É de propósito: o pacote existe para ser
enviado a outras pessoas, e as extrações não podem circular. Quem recebe põe os
próprios arquivos lá dentro. Use --com-dados só para gerar uma cópia para você
mesmo.

Uso:  python build_portatil.py [--zip] [--com-pyc] [--com-dados]

  --zip        também gera o dist/ExploradorDeFallout.zip pronto para enviar
  --com-pyc    mantém os .pyc (pasta ~95 MB maior, primeira abertura mais rápida)
  --com-dados  copia a extracoes/ atual para dentro do pacote — NÃO use no que
               for distribuído: o pacote passa a levar os dados reais dentro
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

# Os .bat são plano B: ficam fora do caminho para a pasta abrir com um único
# item clicável, e para não haver dois "Explorador de Fallout" lado a lado
# (o Windows esconde a extensão por padrão).
PLANO_B = os.path.join(DIST, "Se nao abrir")

# Todo módulo que o app importa em tempo de execução, inclusive os importados
# dentro de função: `gerar_ofensores_xlsx` só é carregado ao clicar em "Baixar
# Ofensores (Excel)", e por não estar aqui o botão quebrava com ModuleNotFound
# na versão portátil, sem quebrar nada na máquina de desenvolvimento.
# Os `converter_*.py` e o `reconstruir_fotografia.py` ficam de fora de
# propósito: são ferramentas de quem prepara a extração, e não do app que a lê.
FONTES = ["app.py", "fallout_core.py", "pipeline.py", "fotografia.py", "parametros.py",
          "distribuicao.py", "previsao.py", "drive_sync.py",
          "comparacao.py", "gerar_pptx.py", "gerar_slide3col.py",
          "gerar_ofensores_xlsx.py", "launcher.py"]

EXE = "Explorador de Fallout.exe"
# Aparece nas propriedades do .exe e no Gerenciador de Tarefas. Ter esses
# campos preenchidos também reduz falso positivo de antivírus em executável
# sem assinatura digital.
EMPRESA = "MAYH"
VERSAO = (1, 0, 0, 0)
# Opcional: se existir um icone.ico na raiz do projeto, ele é usado no atalho.
ICONE = os.path.join(RAIZ, "icone.ico")

# Versões travadas: um app distribuído tem que sair igual a cada build. O
# streamlit em especial não pode flutuar — o CSS do app depende do DOM interno
# dele (ver comentário no requirements.txt).
DEPS = ["streamlit==1.58.0", "pandas==2.2.3", "numpy==2.1.3",
        "matplotlib==3.10.9", "openpyxl==3.1.5", "python-pptx==1.0.2",
        "pywebview==6.2.1", "google-auth==2.56.0"]

# Os .bat abaixo moram em "Se nao abrir\", um nivel abaixo da raiz da dist —
# dai o "%~dp0.." em todos os caminhos.
ATALHO = r"""@echo off
rem Mesmo que o atalho .exe: janela propria, sem console.
start "" "%~dp0..\python\pythonw.exe" "%~dp0..\launcher.py"
"""

# Windows marca arquivos vindos de outra maquina (Mark of the Web) e o .NET
# recusa DLLs nesse estado — e ai a janela nativa nao sobe. Unblock-File
# resolve e nao precisa de admin.
DESBLOQUEAR = r"""@echo off
echo Liberando os arquivos copiados de outra maquina...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Get-ChildItem -LiteralPath '%~dp0..' -Recurse -File | Unblock-File"
echo.
echo Pronto. Abra o "Explorador de Fallout".
pause
"""

# Ultimo recurso, sem pywebview: Streamlit no navegador, com console visivel
# para poder encerrar fechando a janela preta.
NAVEGADOR = r"""@echo off
cd /d "%~dp0.."
title Explorador de Fallout (navegador) - feche esta janela para encerrar
"python\python.exe" -m streamlit run "app.py" --server.address=127.0.0.1 --browser.gatherUsageStats=false
pause
"""

# Recurso VERSIONINFO do .exe. 0416/04B0 = pt-BR, Unicode.
VERSION_FILE = """\
VSVersionInfo(
  ffi=FixedFileInfo(filevers={v}, prodvers={v}, mask=0x3f, flags=0x0,
                    OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041604B0', [
      StringStruct('CompanyName', '{empresa}'),
      StringStruct('FileDescription', 'Explorador de Fallout'),
      StringStruct('FileVersion', '{versao}'),
      StringStruct('InternalName', 'ExploradorDeFallout'),
      StringStruct('LegalCopyright', ''),
      StringStruct('OriginalFilename', '{exe}'),
      StringStruct('ProductName', 'Explorador de Fallout'),
      StringStruct('ProductVersion', '{versao}'),
    ])]),
    VarFileInfo([VarStruct('Translation', [1046, 1200])]),
  ],
)
"""

LEIAME = """Explorador de Fallout — versão portátil
=======================================

Antes de tudo, se você recebeu isto num .zip
--------------------------------------------
Se a pasta veio de um arquivo .zip, o ideal é desbloquear ANTES de extrair:
clique com o botão direito no .zip, escolha "Propriedades", marque
"Desbloquear" no rodapé da janela e só então extraia.

O Windows marca como bloqueado tudo que vem de outra máquina, e isso impede
o programa de abrir na janela própria. Desbloquear o .zip resolve de uma vez
só, e não precisa de administrador. Se já extraiu, dá para consertar depois
(veja "Se não abrir", mais abaixo).

Como usar
---------
Dê um duplo clique em "Explorador de Fallout".

Não é preciso ter Python instalado nem permissão de administrador: o Python
já vai junto, dentro da pasta "python". A primeira abertura leva uns 10
segundos; as seguintes são mais rápidas.

Onde ficam os dados
-------------------
Na pasta "extracoes", ao lado do atalho. Ela vem VAZIA de propósito: os
dados não viajam junto com o programa. Peça os arquivos a quem prepara a
extração e coloque-os aqui.

Na raiz da pasta ficam cinco arquivos:

    extracoes/
      RelatorioDFTOctane.xlsx   <- obrigatório: defeitos do Octane
      RelatorioUSOctane.xlsx    <- obrigatório: User Stories do Octane
      RelatorioJira.xlsx        <- opcional: chamados PDST- e INC-
      RelatorioBRJ.xlsx         <- opcional: chamados BRJ- (aba "Defeitos - BRJ")
      historico.xlsx            <- opcional: meses já fechados

Sem os dois primeiros o app não carrega e avisa qual está faltando. Os
outros três podem faltar: sem o do Jira, os pedidos ligados a chamado
ficam sem status; sem o de BRJ, o Base Residencial usa só a aba
"Defeitos" da própria planilha; sem o histórico, os meses antigos somem do gráfico.

Depois, uma pasta por jornada. Existem três formatos, e o app reconhece
sozinho qual é o de cada pasta:

 1. Extração — Base Móvel, Cross Sell, Prospect, PME Fixa

        <Jornada>/falhas/     <- um arquivo por mês
        <Jornada>/sucessos/   <- um arquivo por mês

 2. Base Residencial — a pasta se chama "NBA"

        Um .xlsx por mês, com as abas "Analitico" e "Defeitos".

 3. PME Móvel — a pasta se chama "PME"

        Um arquivo só, .csv ou .xlsx com a aba "Base", trazendo
        sucessos e falhas na mesma planilha.

Para acrescentar uma jornada nova, crie a pasta dela em um desses
formatos. O nome da pasta é o nome que aparece no app.

Para atualizar o relatório, troque os arquivos dessa pasta e clique em
"Atualizar dados" dentro do app.

Fotografias de cada extração
----------------------------
O botão "Gravar fotografia" guarda o resultado da extração que está na
tela em "extracoes\\_fotografias", numa pasta com a data dos dados. É o
que permite comparar uma semana com a outra, na aba "Comparação": grave
uma a cada extração nova, depois de clicar em "Atualizar dados". Gravar
de novo a mesma extração substitui a fotografia dela.

A aba "Comparação" compara a fotografia mais recente com a de uns 7
dias antes: o que mudou na distribuição, na projeção, nos defeitos e
nos erros, com um texto pronto para e-mail e um Excel com os detalhes.

As fotografias trazem o detalhe pedido a pedido, como as extrações:
valem para elas os mesmos cuidados, e elas não devem circular.

Por que a primeira abertura depois de um arquivo novo demora
------------------------------------------------------------
Ler .xlsx é lento: os arquivos grandes do Base Residencial e da PME
levam de 10 segundos a mais de um minuto cada um. Por isso, ao ler um
.xlsx pela primeira vez, o app grava ao lado dele um .csv com o mesmo
conteúdo ("<nome>__base.csv", ou o par "<nome>__analitico.csv" e
"<nome>__defeitos.csv") e passa a usar esse .csv, que abre em segundos.

Então a primeira abertura depois de colocar um arquivo novo é a
demorada; as seguintes são rápidas. Os .xlsx originais podem ficar onde
estão — o app passa a ignorá-los —, e apagar os .csv só faz a conversão
acontecer de novo.

Os arquivos precisam vir prontos
--------------------------------
Este programa apenas LÊ os dados; ele não converte a extração bruta do
portal. As ferramentas de conversão ficam com quem prepara os dados.

Então, se você baixou um relatório novo do portal e o app não o lê, não
tente ajustar o arquivo aqui: peça a versão já convertida.

Se não abrir
------------
Tudo o que você precisa está na pasta "Se nao abrir", e é para usar nesta
ordem. Nada ali pede administrador.

1. "1 - Desbloquear arquivos.bat"

   Resolve o caso mais comum: os arquivos vieram de outra máquina e o
   Windows os marcou como bloqueados, e aí as bibliotecas da janela não
   carregam. Rode e tente abrir de novo.

2. "2 - Abrir sem o atalho.bat"

   Faz exatamente o mesmo que o atalho, sem passar pelo .exe. Serve para
   quando o antivírus da máquina barra o executável.

3. "3 - Abrir no navegador.bat"

   O programa é o mesmo; só roda numa aba do navegador em vez de janela
   própria. Para encerrar, feche a janela preta do console.

O arquivo "launcher.log", ao lado do atalho, registra o que aconteceu em
cada tentativa — é o que ajuda a descobrir o motivo se nada acima resolver.
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


def _tamanho(caminho):
    return sum(
        os.path.getsize(os.path.join(r, a))
        for r, _, arqs in os.walk(caminho) for a in arqs
    )


# Pacotes que só o instalador usa, ou que o Streamlit traz por dependência mas
# este app não exercita. A remoção de altair/pydeck/narwhals foi verificada
# rodando o app inteiro sem eles: as abas e as tabelas continuam funcionando.
PACOTES_DESCARTAVEIS = [
    "pip", "setuptools", "wheel", "pkg_resources",   # só servem para instalar
    "altair", "pydeck", "narwhals",                  # gráficos nativos não usados
]


def _enxugar(remover_pyc):
    """Tira do site-packages o que não é usado em tempo de execução."""
    sp = os.path.join(PY_DIR, "Lib", "site-packages")
    antes = _tamanho(PY_DIR)

    for nome in PACOTES_DESCARTAVEIS:
        for alvo in (os.path.join(sp, nome), *__import__("glob").glob(
                os.path.join(sp, nome + "-*.dist-info"))):
            if os.path.isdir(alvo):
                shutil.rmtree(alvo, ignore_errors=True)

    # Suítes de teste, headers e libs estáticas: nada disso roda.
    for raiz, dirs, arqs in os.walk(sp, topdown=True):
        for d in list(dirs):
            if d in ("tests", "test", "testing", "include", "includes"):
                shutil.rmtree(os.path.join(raiz, d), ignore_errors=True)
                dirs.remove(d)
        for a in arqs:
            # .pyi = stubs de tipagem (só IDE); .lib/.a = compilar contra o Arrow
            if a.endswith((".pyi", ".lib", ".a", ".pdb", ".chm")):
                try:
                    os.remove(os.path.join(raiz, a))
                except OSError:
                    pass

    if remover_pyc:
        # Trocam ~95 MB por alguns segundos a mais só na primeira abertura,
        # já que o Python regrava os .pyc conforme importa.
        for raiz, dirs, _ in os.walk(PY_DIR, topdown=True):
            for d in list(dirs):
                if d == "__pycache__":
                    shutil.rmtree(os.path.join(raiz, d), ignore_errors=True)
                    dirs.remove(d)

    depois = _tamanho(PY_DIR)
    print(f"  {(antes - depois) / 1024 / 1024:,.0f} MB removidos")


def _compilar_atalho():
    """
    Compila atalho.py no .exe clicável da dist. Devolve True se conseguiu.

    Não derruba o build ao falhar: sem o .exe a pasta continua inteiramente
    utilizável pelos .bat de "Se nao abrir", que fazem a mesma coisa.
    """
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("  PyInstaller não instalado — a pasta sai só com os .bat.")
        print("  Para ter o atalho clicável:  pip install pyinstaller")
        return False

    trabalho = os.path.join(CACHE, "pyinstaller")
    os.makedirs(trabalho, exist_ok=True)
    arq_versao = os.path.join(trabalho, "versao.txt")
    versao = ".".join(str(n) for n in VERSAO)
    with open(arq_versao, "w", encoding="utf-8") as f:
        f.write(VERSION_FILE.format(
            v=VERSAO, versao=versao, empresa=EMPRESA, exe=EXE))

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",    # são poucos MB: extrair custa menos de um segundo
        "--noconsole",  # sem janela preta ao clicar
        "--clean",
        "--noconfirm",
        # o PyInstaller loga ~100 linhas de INFO em stderr; só interessam
        # os avisos e os erros.
        "--log-level=WARN",
        "--name", os.path.splitext(EXE)[0],
        "--distpath", DIST,
        "--workpath", os.path.join(trabalho, "build"),
        "--specpath", trabalho,
        "--version-file", arq_versao,
    ]
    if os.path.exists(ICONE):
        cmd += ["--icon", ICONE]
    cmd.append(os.path.join(RAIZ, "atalho.py"))

    try:
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        print(f"  PyInstaller falhou (código {e.returncode})"
              " — a pasta sai só com os .bat")
        return False

    mb = os.path.getsize(os.path.join(DIST, EXE)) / 1024 / 1024
    icone = os.path.basename(ICONE) if os.path.exists(ICONE) else "padrão"
    print(f"  {EXE} ({mb:,.1f} MB, ícone {icone})")
    return True


def _zipar():
    """Empacota a dist no .zip que vai ser enviado."""
    print("\n[+] compactando (são centenas de MB, leva alguns minutos)")
    arq = shutil.make_archive(
        DIST, "zip",
        root_dir=os.path.dirname(DIST), base_dir=os.path.basename(DIST),
    )
    print(f"  {arq}")
    print(f"  {os.path.getsize(arq) / 1024 / 1024:,.0f} MB")


def main():
    com_dados = "--com-dados" in sys.argv
    print(f"→ destino: {DIST}")
    if os.path.exists(DIST):
        print("  limpando build anterior")
        shutil.rmtree(DIST)
    os.makedirs(PY_DIR, exist_ok=True)

    print("\n[1/7] Python embeddable")
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

    print("\n[2/7] pip")
    get_pip = _baixar(GET_PIP_URL, os.path.join(CACHE, "get-pip.py"))
    _py(get_pip, "--no-warn-script-location")

    print("\n[3/7] dependências (demora alguns minutos)")
    _py("-m", "pip", "install", "--no-warn-script-location", *DEPS)

    print("\n[4/7] enxugando")
    _enxugar(remover_pyc="--com-pyc" not in sys.argv)

    print("\n[5/7] aplicação")
    for nome in FONTES:
        shutil.copy2(os.path.join(RAIZ, nome), os.path.join(DIST, nome))
        print(f"  {nome}")

    # A pasta de dados sai vazia por padrão. O pacote é feito para ser enviado,
    # e copiar a extracoes/ para dentro dele mandaria junto o detalhe pedido a
    # pedido de todas as jornadas — 155 MB de dado real no .zip que vai por
    # e-mail. Quem recebe põe os próprios arquivos.
    origem_dados = os.path.join(RAIZ, "extracoes")
    destino_dados = os.path.join(DIST, "extracoes")
    if com_dados and os.path.isdir(origem_dados):
        shutil.copytree(origem_dados, destino_dados)
        mb = _tamanho(destino_dados) / 1024 / 1024
        print(f"  extracoes/ ({mb:,.0f} MB de DADOS REAIS — cópia para uso próprio)")
    else:
        os.makedirs(destino_dados, exist_ok=True)
        if com_dados:
            print(f"  extracoes/ (vazia — não há {origem_dados} para copiar)")
        else:
            print("  extracoes/ (vazia — quem receber põe os dados; "
                  "--com-dados copia os seus)")

    print("\n[6/7] atalho clicável")
    tem_exe = _compilar_atalho()

    print("\n[7/7] planos B e leia-me")
    os.makedirs(PLANO_B, exist_ok=True)
    # cp1252: os .bat são lidos pelo cmd.exe, não por Python.
    for nome, conteudo in [
        ("1 - Desbloquear arquivos.bat", DESBLOQUEAR),
        ("2 - Abrir sem o atalho.bat", ATALHO),
        ("3 - Abrir no navegador.bat", NAVEGADOR),
    ]:
        with open(os.path.join(PLANO_B, nome), "w", encoding="cp1252") as f:
            f.write(conteudo)
        print(f"  Se nao abrir\\{nome}")
    # utf-8-sig: sem o BOM o Bloco de Notas do Windows lê como ANSI e os
    # acentos viram "DÃª".
    with open(os.path.join(DIST, "LEIA-ME.txt"), "w", encoding="utf-8-sig") as f:
        f.write(LEIAME)
    print("  LEIA-ME.txt")

    tam = sum(
        os.path.getsize(os.path.join(r, a))
        for r, _, arqs in os.walk(DIST) for a in arqs
    )
    abertura = EXE if tem_exe else r"Se nao abrir\2 - Abrir sem o atalho.bat"
    print(f"\n✓ pronto: {DIST}")
    print(f"  {tam / 1024 / 1024:,.0f} MB")
    print(f'  teste com um duplo clique em "{abertura}"')

    if "--zip" in sys.argv:
        _zipar()
    else:
        print("  rode com --zip para já sair compactado para envio")
    print('\n  Ao enviar, avise: desbloquear o .zip (botão direito >')
    print("  Propriedades > Desbloquear) ANTES de extrair.")

    if com_dados:
        print("\n  ATENÇÃO: este pacote leva as extrações reais dentro dele.")
        print("  É uma cópia para uso próprio — não envie para outras pessoas.")
    else:
        print("\n  A pasta 'extracoes' saiu vazia. Quem receber precisa colocar")
        print("  os arquivos lá dentro; o LEIA-ME.txt explica a estrutura.")


if __name__ == "__main__":
    if os.name != "nt":
        sys.exit("Este build é específico para Windows.")
    main()
