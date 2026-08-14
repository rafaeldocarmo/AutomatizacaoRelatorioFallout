# -*- coding: utf-8 -*-
"""
Núcleo de dados compartilhado entre app.py (dashboard interativo) e pipeline.py
(geração de relatório/PNG/PPTX): leitura das bases, join com Octane e
categorização do fallout por mês.

Este módulo não depende de Streamlit nem de estado global — todas as funções
recebem `base_dir` explicitamente e retornam dados, para poder ser chamado
com segurança a partir de múltiplas sessões/threads simultâneas.
"""
import csv
import glob
import os
import re

import pandas as pd

ABREV_MES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
             "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
MESES_PT = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
            7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
MESES_LABEL = {m: f"{abrev}-26" for abrev, m in ABREV_MES.items()}

FASES_CORRIGIDO = {"Corrigido", "Fechado"}
FASES_MOPS = {"Cancelado", "Rejeitado"}
# "Legado Aberto" existe apenas no Base Residencial, marcado na leitura a partir
# da Classificação Original. Nas jornadas de extração o defeito cancelado no
# Octane continua sendo avaliação por MOPs, como sempre foi.
FASES_FORA_TRATAMENTO = FASES_MOPS
# Fase única das jornadas com catálogo próprio (NBA): todo pedido com defeito
# associado está em tratamento.
FASE_TRATAMENTO_NBA = "Em Tratamento"

# Pasta da jornada → nome mostrado no app e no título do slide. A pasta segue
# com o nome da origem do arquivo; só a exibição muda.
NOMES_EXIBICAO = {"NBA": "Base Residencial", "PME": "PME Móvel"}

# Pasta local (dentro do projeto) com as jornadas e as planilhas do Octane.
PASTA_DADOS = "extracoes"
ARQ_DFT = "RelatorioDFTOctane.xlsx"
ARQ_US = "RelatorioUSOctane.xlsx"

COLUNAS_OCTANE_BASE = ["ID", "Name", "Phase", "Bugfix Milestone", "Team", "Type"]
COL_US_MELHORIA = "US de Melhoria"

# ── Jornadas em formato "RPA" ────────────────────────────────────────────────
# Algumas jornadas (PME) não vêm da extração Salesforce/Vlocity e sim de um
# relatório de RPA: um CSV único, separado por ';' e em cp1252, com sucessos e
# falhas na mesma planilha. O de-para abaixo traduz para o formato interno.
RPA_SEP = ";"
RPA_ENCODING = "cp1252"
RPA_COL_RESULTADO = "SUCESSO OU FALHA"   # coluna U
RPA_VALOR_FALHA = "Erro"
RPA_COL_DATA = "DATA_CRIACAO"
RPA_COL_ENTREGA = "DATA CORREÇÃO"        # coluna S
RPA_COL_DEFEITO = "DEFEITO"              # coluna M
RENAME_RPA = {
    "PEDIDO_SF": "OrderNumber",
    "DESCRICAO_ERRO": "ErrorHandled__c",
    "SEGMENTO": "Channel",
    "TERRITORIO": "Segment",
    "DESCRICAO": "OrderStatus",
    "CLASSIFICAÇÃO": "SubStatus",
}
# Rótulos que não são defeito: têm categoria própria na regra de negócio.
RPA_ROTULOS_NAO_DEFEITO = {
    "enviado e-mail",                  # → Em Avaliação por MOPs
    "enviado e-mail - outros times",   # → Em avaliação - Outros times
    "erro de processo",                # → Falha no Processo Usuário
}

# ── Jornadas em formato "NBA" ────────────────────────────────────────────────
# Planilha XLSX com duas abas: "Analitico" (um pedido por linha) e "Defeitos",
# que é o catálogo local — aqui os defeitos são BRJ-*, não existem no Octane.
NBA_ABA_ANALITICO = "Analitico"
NBA_ABA_DEFEITOS = "Defeitos"
# Ler XLSX com openpyxl é lento (~40 s para os três meses). converter_nba.py
# grava cada planilha como um par de CSVs, que o leitor prefere quando existe.
NBA_SUF_ANALITICO = "__analitico.csv"
NBA_SUF_DEFEITOS = "__defeitos.csv"
NBA_COL_STATUS = "Status"
NBA_VALOR_SUCESSO = "SUCESSO"
NBA_VALOR_CANCELADO = "CANCELADO"        # vira a categoria "Legado Aberto"
NBA_COL_DATA = "Data"
NBA_COL_DEFEITO = "Defeito"
RENAME_NBA = {
    "CD": "OrderNumber",
    "Erro agrupado": "ErrorHandled__c",
    "Canal": "Channel",
    "Segmento Comercial": "Segment",
    "Estado": "State",
    "Classificação Original": "SubStatus",
}

RENAME_FALHAS = {
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.OrderNumber": "OrderNumber",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Channel__c": "Channel",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Segment__c": "Segment",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Status": "OrderStatus",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.SubStatus__c": "SubStatus",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.BiometryStatus__c": "BiometryStatus",
    "vlocity_cmt__State__c": "State",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.vlocity_cmt__Reason__c": "Reason",
}

# A partir de ago/26 o mesmo relatório de falhas passou a ser exportado com o
# rótulo em português no lugar do nome de API. Os dois layouts convivem na pasta
# porque os meses anteriores continuam com os arquivos antigos.
COL_DATA_FALHAS_PT = "Data de Criação"
RENAME_FALHAS_PT = {
    COL_DATA_FALHAS_PT: "CreatedDate",
    "Número do pedido": "OrderNumber",
    "Canal": "Channel",
    "Segmento Comercial": "Segment",
    "Status": "OrderStatus",
    "Status Biometria": "BiometryStatus",
    "Número Defeito": "DefectNumber__c",
    "Nome do Ofensor": "OffenderName__c",
    "Erro Tratado": "ErrorHandled__c",
}

# Sucessos, também em dois layouts: até jul/26 um agregado de uma linha só
# (coluna "expr0" com o total do mês); de ago/26 em diante um relatório com
# cabeçalho de filtros e uma linha por dia ("01/08/2026,1555").
SUC_COL_AGREGADA = "expr0"
SUC_DATA_DIA = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")


def mes_do_arquivo(path):
    """Deduz o número do mês (1-12) a partir do nome do arquivo (ex.: '...jan26.csv')."""
    nome = os.path.splitext(os.path.basename(path))[0].lower()
    for abrev, num in ABREV_MES.items():
        if abrev in nome:
            return num
    return None


def _arquivos(pasta, padrao):
    """
    Lista arquivos ignorando os temporários de bloqueio do Excel ("~$nome.xlsx"),
    que aparecem enquanto a planilha está aberta e não são planilhas de verdade.
    """
    return sorted(
        c for c in glob.glob(os.path.join(pasta, padrao))
        if not os.path.basename(c).startswith(("~$", ".~"))
    )


def _subpastas(caminho):
    if not os.path.isdir(caminho):
        return []
    return sorted(d for d in os.listdir(caminho) if os.path.isdir(os.path.join(caminho, d)))


def formato_jornada(pasta):
    """
    Descobre como a jornada guarda os dados:
      "extracao" → subpastas falhas/ e sucessos/ (export Salesforce/Vlocity)
      "rpa"      → um CSV solto na pasta, com sucessos e falhas juntos
      "nba"      → XLSX com as abas "Analitico" e "Defeitos"
      None       → não parece uma jornada
    """
    if os.path.isdir(os.path.join(pasta, "falhas")):
        return "extracao"
    if not os.path.isdir(pasta):
        return None
    # O par __defeitos.csv identifica o NBA já convertido; sem ele, um CSV solto
    # é uma jornada de RPA.
    if (_arquivos(pasta, "*" + NBA_SUF_DEFEITOS)
            or _arquivos(pasta, "*.xlsx")):
        return "nba"
    if _arquivos(pasta, "*.csv"):
        return "rpa"
    return None


def nome_exibicao(jornada):
    """Nome que o usuário vê; por padrão é o próprio nome da pasta."""
    return NOMES_EXIBICAO.get(jornada, jornada)


def _tem_jornadas(caminho):
    """True se a pasta contém diretórios de jornada, em qualquer um dos formatos."""
    return any(formato_jornada(os.path.join(caminho, d)) for d in _subpastas(caminho))


def raiz_dados(base_dir):
    """
    Pasta que guarda as jornadas e as planilhas do Octane: `base_dir/extracoes`.

    Se as jornadas não estiverem soltas ali dentro e sim agrupadas em um único
    diretório (é o que acontece ao copiar a pasta exportada inteira, ex.:
    extracoes/RelatoriosAutomatização/...), desce mais um nível automaticamente.
    """
    raiz = os.path.join(base_dir, PASTA_DADOS)
    if _tem_jornadas(raiz):
        return raiz
    for d in _subpastas(raiz):
        if _tem_jornadas(os.path.join(raiz, d)):
            return os.path.join(raiz, d)
    return raiz


def jornadas_disponiveis(base_dir):
    """Pastas de jornada disponíveis, em qualquer um dos formatos suportados."""
    raiz = raiz_dados(base_dir)
    if not os.path.isdir(raiz):
        return []
    return sorted(
        d for d in os.listdir(raiz)
        if formato_jornada(os.path.join(raiz, d))
    )


def jornadas_consolidaveis(base_dir):
    """
    Jornadas que entram na soma do "Consolidado" — só as de extração.

    As de outros formatos (RPA, NBA) ficam de fora de propósito: cada uma conta
    o denominador do seu jeito — "pedidos processados pelo robô", por exemplo —
    então somar com as demais produziria um fallout que não quer dizer nada.
    Elas continuam disponíveis individualmente, com aba, KPI e slide próprios.
    """
    raiz = raiz_dados(base_dir)
    return [
        j for j in jornadas_disponiveis(base_dir)
        if formato_jornada(os.path.join(raiz, j)) == "extracao"
    ]


def resolver_jornadas(jornada, base_dir):
    """Traduz o nome escolhido nas pastas que devem ser lidas."""
    if jornada == "Consolidado":
        return jornadas_consolidaveis(base_dir)
    if jornada == "Base + Cross Sell":
        return ["Base Móvel", "Cross Sell"]
    return [jornada]


def _limpar_defeito(serie):
    """Texto do campo de defeito sem espaços não-quebráveis nem sobras."""
    return serie.fillna("").astype(str).str.replace("\xa0", " ", regex=False).str.strip()


def _ler_jornada_rpa(pasta):
    """
    Lê uma jornada em formato RPA e devolve (df_falhas, sucessos_por_mes).

    O CSV traz sucessos e falhas juntos: a coluna "SUCESSO OU FALHA" separa os
    dois. As falhas viram linhas no formato interno; os sucessos são apenas
    contados por mês, que é o que o cálculo do fallout precisa.
    """
    partes = [
        pd.read_csv(f, sep=RPA_SEP, encoding=RPA_ENCODING, dtype=str, low_memory=False)
        for f in _arquivos(pasta, "*.csv")
    ]
    bruto = pd.concat(partes, ignore_index=True)
    bruto = bruto.loc[:, ~bruto.columns.str.startswith("Unnamed:")]

    data = pd.to_datetime(bruto[RPA_COL_DATA], format="%d/%m/%Y", errors="coerce")
    bruto = bruto[data.notna()].copy()
    bruto["CreatedDate"] = (
        data[data.notna()].dt.tz_localize("America/Sao_Paulo", ambiguous=True,
                                          nonexistent="shift_forward").dt.tz_convert("UTC")
    )

    resultado = bruto[RPA_COL_RESULTADO].fillna("").str.strip().str.lower()
    e_falha = resultado == RPA_VALOR_FALHA.lower()

    sucessos = (
        bruto[~e_falha]
        .assign(Mes=lambda d: d["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month)
        .groupby("Mes").size().rename("Sucessos")
    )

    df = bruto[e_falha].rename(columns=RENAME_RPA).copy()
    df["DefectNumber_orig"] = _limpar_defeito(df[RPA_COL_DEFEITO])
    # Aqui "0" é a marca de "sem defeito"; internamente isso é -1.
    numero = pd.to_numeric(df["DefectNumber_orig"], errors="coerce")
    df["DefectNumber__c"] = numero.where(numero > 0, -1).fillna(-1).astype(int)

    # Rótulos como "Paliativo RPA - 6" ou "NFCOM" são defeitos de verdade, só que
    # controlados fora do Octane — contam como defeito mesmo sem número.
    rotulo = df["DefectNumber_orig"].str.lower()
    df["TemDefeito"] = (
        (df["DefectNumber__c"] > 0) & (df["DefectNumber__c"] != 999999)
    ) | (
        (df["DefectNumber__c"] == -1)
        & (df["DefectNumber_orig"] != "")
        & (df["DefectNumber_orig"] != "0")
        & ~rotulo.isin(RPA_ROTULOS_NAO_DEFEITO)
    )

    # Data de entrega do próprio relatório: só entra onde o Octane não tiver.
    df["MilestoneRPA"] = pd.to_datetime(
        df[RPA_COL_ENTREGA].fillna("").str.strip(), format="%d/%m/%y", errors="coerce"
    )
    return df, sucessos


def _catalogo_nba(cat):
    """
    Normaliza a aba/CSV "Defeitos": o catálogo local de defeitos BRJ-*, que faz
    aqui o papel do Octane. Devolve um DataFrame no formato das colunas DFT_*.
    """
    tipo = cat["Tipo de Chamado"].fillna("").str.strip().str.upper()
    return pd.DataFrame({
        "DefectNumber_orig": cat[NBA_COL_DEFEITO].fillna("").str.strip(),
        # O título vem com tabulação/quebra no meio, que o matplotlib desenha
        # como caixinha de glifo ausente no PNG do dashboard.
        "DFT_Name": (cat["Título do Defeito"].fillna("")
                     .str.replace(r"\s+", " ", regex=True).str.strip()),
        # Todo pedido com defeito está em tratamento; o "Status do Chamado" da
        # planilha não entra na conta.
        "DFT_Phase": FASE_TRATAMENTO_NBA,
        # A data de implantação é o equivalente ao Bugfix Milestone: separa o
        # que é "Planejado" do que fica "s/ data" e alimenta o Planejamento
        # de Redução.
        "DFT_BugfixMilestone": pd.to_datetime(
            cat["Data de Implantação"], format="%d/%m/%Y", errors="coerce"),
        # Nomes vêm com caixa e espaços inconsistentes ("Gilberto", "gilberto",
        # " Gilberto"), o que criaria times diferentes na tabela do slide.
        "DFT_Team": cat["Responsável pelo Chamado"].fillna("").str.strip().str.title(),
        "DFT_Type": tipo.map(lambda t: "User Story" if t.startswith("MELHORIA") else "Defect"),
    }).drop_duplicates(subset="DefectNumber_orig")


def _ler_jornada_nba(pasta):
    """
    Lê uma jornada em formato NBA e devolve (df_falhas, sucessos_por_mes).

    A aba "Analitico" traz um pedido por linha, com sucessos e falhas juntos; a
    aba "Defeitos" entra no lugar do Octane, já que os defeitos são BRJ-* e não
    existem lá.
    """
    partes, catalogos = [], []

    def _registrar(analitico, defeitos):
        partes.append(analitico)
        # Guarda até quando o arquivo vai, para saber qual catálogo é o mais novo.
        ate = pd.to_datetime(analitico[NBA_COL_DATA], format="%d/%m/%Y", errors="coerce").max()
        catalogos.append((ate, _catalogo_nba(defeitos)))

    # Pares já convertidos (rápido).
    convertidos = set()
    for arq_def in _arquivos(pasta, "*" + NBA_SUF_DEFEITOS):
        base = arq_def[: -len(NBA_SUF_DEFEITOS)]
        arq_an = base + NBA_SUF_ANALITICO
        if not os.path.isfile(arq_an):
            raise FileNotFoundError(f"Falta o par de {os.path.basename(arq_def)}: {arq_an}")
        convertidos.add(os.path.basename(base))
        _registrar(pd.read_csv(arq_an, dtype=str), pd.read_csv(arq_def, dtype=str))

    # XLSX ainda não convertidos continuam sendo lidos direto, para que largar um
    # arquivo novo na pasta funcione sem passar pelo converter.
    for arq in _arquivos(pasta, "*.xlsx"):
        if os.path.splitext(os.path.basename(arq))[0] in convertidos:
            continue
        _registrar(
            pd.read_excel(arq, sheet_name=NBA_ABA_ANALITICO, dtype=str),
            pd.read_excel(arq, sheet_name=NBA_ABA_DEFEITOS, dtype=str),
        )

    if not partes:
        raise FileNotFoundError(f"Nenhuma planilha do NBA encontrada em {pasta}")
    bruto = pd.concat(partes, ignore_index=True)

    # O mesmo defeito aparece em vários arquivos e muda de status ao longo do
    # tempo (avaliação por MOPs primeiro, tratamento depois). Vale o do arquivo
    # mais recente — ordenar por nome daria o resultado errado.
    catalogos.sort(key=lambda par: par[0], reverse=True)
    catalogo = (
        pd.concat([c for _, c in catalogos], ignore_index=True)
        .drop_duplicates(subset="DefectNumber_orig")
    )

    data = pd.to_datetime(bruto[NBA_COL_DATA], format="%d/%m/%Y", errors="coerce")
    bruto = bruto[data.notna()].copy()
    bruto["CreatedDate"] = (
        data[data.notna()].dt.tz_localize("America/Sao_Paulo", ambiguous=True,
                                          nonexistent="shift_forward").dt.tz_convert("UTC")
    )

    e_sucesso = bruto[NBA_COL_STATUS].fillna("").str.strip().str.upper() == NBA_VALOR_SUCESSO
    sucessos = (
        bruto[e_sucesso]
        .assign(Mes=lambda d: d["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month)
        .groupby("Mes").size().rename("Sucessos")
    )

    df = bruto[~e_sucesso].rename(columns=RENAME_NBA).copy()
    # Alguns pedidos trazem mais de um defeito na mesma célula
    # ("BRJ-385|BRJ-325"); o pedido é contado no primeiro.
    df["DefectNumber_orig"] = (
        df[NBA_COL_DEFEITO].fillna("").astype(str)
        .str.split("|").str[0].str.strip()
    )
    numero = pd.to_numeric(df["DefectNumber_orig"], errors="coerce")
    df["DefectNumber__c"] = numero.fillna(-1).astype(int)   # preserva o 999999
    df["TemDefeito"] = (df["DefectNumber_orig"] != "") & (df["DefectNumber__c"] != 999999)
    # Categoria exclusiva desta jornada: o que a origem marca como CANCELADO é
    # legado em aberto. A marca sai daqui, e não do valor da coluna, porque
    # "Cancelado" também aparece no SubStatus de outras jornadas.
    df["LegadoAberto"] = (
        df["SubStatus"].fillna("").astype(str).str.strip().str.upper() == NBA_VALOR_CANCELADO
    )

    df = df.merge(catalogo, on="DefectNumber_orig", how="left")
    return df, sucessos


def _caminho_octane(raiz, base_dir, nome):
    """
    Localiza uma planilha do Octane: primeiro junto das jornadas (é onde a
    exportação a coloca), com a raiz do projeto como alternativa.
    """
    for pasta in (raiz, base_dir):
        caminho = os.path.join(pasta, nome)
        if os.path.isfile(caminho):
            return caminho
    raise FileNotFoundError(
        f"'{nome}' não encontrado. Coloque o arquivo em {raiz} (junto das "
        f"pastas de jornada) ou em {base_dir}."
    )


def _norm_id(v):
    """Normaliza um ID que pode vir como float ('182366.0') para string inteira."""
    try:
        return str(int(float(str(v).strip())))
    except (TypeError, ValueError):
        return str(v).strip()


def _ler_falhas_extracao(caminho):
    """
    Lê um CSV da pasta falhas/ e devolve as colunas já no formato interno,
    aceitando tanto o export com nome de API quanto o com rótulo em português.
    """
    df = pd.read_csv(caminho)
    if COL_DATA_FALHAS_PT in df.columns:
        df = df.rename(columns=RENAME_FALHAS_PT)
        # O layout em português traz só a data, já no fuso de Brasília, enquanto
        # o antigo traz o timestamp com offset. Localizar antes de normalizar
        # para UTC evita que os pedidos do dia 1º caiam no mês anterior quando
        # o resto do código converte de volta para America/Sao_Paulo.
        df["CreatedDate"] = (
            pd.to_datetime(df["CreatedDate"], errors="coerce")
            .dt.tz_localize("America/Sao_Paulo")
            .dt.tz_convert("UTC")
        )
    else:
        df = df.rename(columns=RENAME_FALHAS)
        df["CreatedDate"] = pd.to_datetime(df["CreatedDate"], utc=True, errors="coerce")
    return df


def _ler_sucessos(caminho, mes):
    """
    Total de sucessos do mês `mes` em um CSV da pasta sucessos/, nos dois
    layouts de export.

    Na quebra diária só entram os dias do próprio mês do arquivo: o relatório
    cobre mês atual e anterior, e o anterior já tem CSV próprio — contar os dois
    somaria o mesmo dia duas vezes.
    """
    with open(caminho, encoding="utf-8-sig", newline="") as fh:
        linhas = list(csv.reader(fh))

    if any(SUC_COL_AGREGADA in celula for celula in (linhas[0] if linhas else [])):
        col = pd.read_csv(caminho)[SUC_COL_AGREGADA]
        return int(pd.to_numeric(col, errors="coerce").fillna(0).sum())

    total = 0
    for linha in linhas:
        celulas = [c.strip() for c in linha if c.strip()]
        if len(celulas) != 2:
            continue
        dia = SUC_DATA_DIA.match(celulas[0])
        if dia and int(dia.group(2)) == mes:
            total += int(re.sub(r"\D", "", celulas[1]) or 0)
    return total


def carregar_base(base_dir, jornada):
    """
    Lê falhas + sucessos (CSV) e Octane (Excel) para a jornada informada,
    faz o join e devolve (df, df_octane, resumo, jornadas_combo).

    `jornada` pode ser um nome de pasta única, "Base + Cross Sell" ou
    "Consolidado" — a resolução para as pastas reais é feita aqui.
    """
    raiz = raiz_dados(base_dir)
    jornadas_combo = resolver_jornadas(jornada, base_dir)
    if not jornadas_combo:
        raise FileNotFoundError(
            f"Nenhuma pasta de jornada encontrada em {raiz}. Cada jornada deve ser "
            f"uma pasta com as subpastas 'falhas' e 'sucessos'."
        )

    # ── Falhas ───────────────────────────────────────────────────────────
    # Cada jornada é lida conforme o seu formato; o resultado é sempre o mesmo
    # conjunto de colunas internas, então daqui para baixo tanto faz a origem.
    # `partes_prontas` são as jornadas que já trazem os dados do defeito (NBA
    # tem catálogo próprio) e por isso não passam pelo join com o Octane.
    partes_falhas, partes_prontas, sucessos_rpa = [], [], []
    arquivos_extracao = []
    for j in jornadas_combo:
        pasta = os.path.join(raiz, j)
        formato = formato_jornada(pasta)
        if formato == "nba":
            df_nba, suc_nba = _ler_jornada_nba(pasta)
            partes_prontas.append(df_nba)
            sucessos_rpa.append(suc_nba)
        elif formato == "rpa":
            df_rpa, suc_rpa = _ler_jornada_rpa(pasta)
            partes_falhas.append(df_rpa)
            sucessos_rpa.append(suc_rpa)
        elif formato == "extracao":
            arqs = sorted(glob.glob(os.path.join(pasta, "falhas", "*.csv")))
            if not arqs:
                raise FileNotFoundError(f"Nenhum CSV encontrado em {os.path.join(pasta, 'falhas')}")
            arquivos_extracao += arqs
        else:
            raise FileNotFoundError(f"{pasta} não parece uma jornada (sem falhas/ e sem CSV)")

    if arquivos_extracao:
        df_ext = pd.concat(
            [_ler_falhas_extracao(f) for f in arquivos_extracao], ignore_index=True
        )

        # O export do CRM às vezes traz o número do defeito com espaços não-quebráveis
        # no fim (ex.: "233538\xa0\xa0"). Sem limpar, pd.to_numeric devolve NaN e o
        # pedido acaba contado como "Falta Associar ao Defeito/US" mesmo tendo DFT.
        df_ext["DefectNumber_orig"] = _limpar_defeito(df_ext["DefectNumber__c"])
        # Alguns registros vêm com o prefixo do próprio sistema ("DFT 232143"). Só o
        # prefixo DFT é removido: INC-, DDP-, PDST- e OS referenciam outros sistemas,
        # não existem no Octane e devem continuar sem defeito associado.
        df_ext["DefectNumber__c"] = (
            pd.to_numeric(
                df_ext["DefectNumber_orig"].str.replace(
                    r"^DFT\s*(?=\d)", "", regex=True, case=False),
                errors="coerce",
            )
            .fillna(-1).astype(int)
        )
        # 999999 é "Falha Pontual", que tem categoria própria e não conta como
        # defeito em tratamento.
        df_ext["TemDefeito"] = (
            (df_ext["DefectNumber__c"] > 0) & (df_ext["DefectNumber__c"] != 999999)
        )
        partes_falhas.append(df_ext)

    if not partes_falhas and not partes_prontas:
        raise FileNotFoundError(f"Nenhum dado de falha encontrado para {jornada}")
    df_falhas = (
        pd.concat(partes_falhas, ignore_index=True) if partes_falhas
        else pd.DataFrame(columns=["DefectNumber__c"])
    )

    # ── Octane: DFTs + US, com resolução de "US de Melhoria" ───────────────
    df_dft = pd.read_excel(_caminho_octane(raiz, base_dir, ARQ_DFT))
    df_us = pd.read_excel(_caminho_octane(raiz, base_dir, ARQ_US))

    df_dft_prep = df_dft[
        COLUNAS_OCTANE_BASE + ([COL_US_MELHORIA] if COL_US_MELHORIA in df_dft.columns else [])
    ].copy()

    if COL_US_MELHORIA in df_dft_prep.columns:
        df_us_idx = df_us.set_index("ID")[
            ["Name", "Phase", "Bugfix Milestone", "Team", "Type"]
        ].copy()
        df_us_idx.index = [_norm_id(v) for v in df_us_idx.index]

        df_dft_prep[COL_US_MELHORIA] = df_dft_prep[COL_US_MELHORIA].apply(_norm_id)
        _has_us = df_dft_prep[COL_US_MELHORIA].notna() & ~df_dft_prep[COL_US_MELHORIA].isin(["", "nan", "None"])
        for i, row in df_dft_prep[_has_us].iterrows():
            us_id = row[COL_US_MELHORIA]
            if us_id in df_us_idx.index:
                us = df_us_idx.loc[us_id]
                for col in ["Name", "Phase", "Bugfix Milestone", "Team"]:
                    if pd.notna(us[col]):
                        df_dft_prep.at[i, col] = us[col]
                df_dft_prep.at[i, "Type"] = "User Story"

    df_octane = (
        pd.concat([df_dft_prep[COLUNAS_OCTANE_BASE], df_us[COLUNAS_OCTANE_BASE]], ignore_index=True)
        .drop_duplicates(subset="ID")
        .rename(columns={
            "ID": "DefectNumber__c", "Name": "DFT_Name", "Phase": "DFT_Phase",
            "Bugfix Milestone": "DFT_BugfixMilestone", "Team": "DFT_Team", "Type": "DFT_Type",
        })
    )
    df_octane["DefectNumber__c"] = (
        pd.to_numeric(df_octane["DefectNumber__c"], errors="coerce")
        .fillna(-1).astype(int)
    )

    # ── Join falhas ← Octane ────────────────────────────────────────────
    df = df_falhas.merge(df_octane, on="DefectNumber__c", how="left")

    # As jornadas de catálogo próprio (NBA) entram já resolvidas, depois do join.
    if partes_prontas:
        df = pd.concat([df] + partes_prontas, ignore_index=True) if len(df) else \
             pd.concat(partes_prontas, ignore_index=True)

    # Jornadas RPA trazem a data de entrega no próprio relatório. O Octane é a
    # fonte oficial, então ela só preenche onde o Octane não se posicionou —
    # inclusive nos rótulos que nem existem lá ("Paliativo RPA - 6", "NFCOM").
    if "MilestoneRPA" in df.columns:
        df["DFT_BugfixMilestone"] = df["DFT_BugfixMilestone"].fillna(df["MilestoneRPA"])
        df = df.drop(columns="MilestoneRPA")
    if "TemDefeito" not in df.columns:
        df["TemDefeito"] = (df["DefectNumber__c"] > 0) & (df["DefectNumber__c"] != 999999)
    df["TemDefeito"] = df["TemDefeito"].fillna(False).astype(bool)
    if "LegadoAberto" not in df.columns:
        df["LegadoAberto"] = False
    df["LegadoAberto"] = df["LegadoAberto"].fillna(False).astype(bool)
    # Identificador do defeito para agrupar e exibir: o número quando existe,
    # senão o próprio rótulo — é o caso dos BRJ-* do NBA e dos "Paliativo RPA"
    # do PME, que são defeitos de verdade sem número.
    df["DefectKey"] = (
        df["DefectNumber__c"].astype(str)
        .where(df["DefectNumber__c"] > 0, df["DefectNumber_orig"])
    )

    # Desduplicar por OrderNumber mantendo a linha mais informativa:
    #   1º) DFT real (número > 0);
    #   2º) no empate entre linhas sem DFT (todas viram -1), vence a que tem
    #       rótulo explícito ("Enviado e-mail - Outros Times", "Erro de
    #       Processo", "INC-…") em vez da linha em branco.
    # `kind="stable"` evita que o desempate mude de execução para execução — o
    # quicksort padrão do pandas reordenava empates de forma arbitrária.
    _sem_rotulo = df["DefectNumber_orig"].isin(["", "nan", "None", "-1"])
    df = (
        df.assign(_tem_rotulo=(~_sem_rotulo).astype(int))
          .sort_values(["DefectNumber__c", "_tem_rotulo"],
                       ascending=[False, False], kind="stable")
          .drop_duplicates(subset=["OrderNumber"])
          .drop(columns="_tem_rotulo")
          .reset_index(drop=True)
    )
    df["Mes"] = df["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month

    # ── Sucessos por mês (mesmo denominador do fallout rate) ────────────
    # Nas jornadas RPA os sucessos já vieram contados junto com as falhas; nas
    # de extração eles estão em CSVs próprios, com o mês no nome do arquivo.
    partes_suc = []
    for j in jornadas_combo:
        if formato_jornada(os.path.join(raiz, j)) != "extracao":
            continue   # RPA e NBA já contaram os sucessos junto com as falhas
        arqs_suc = sorted(glob.glob(os.path.join(raiz, j, "sucessos", "*.csv")))
        if not arqs_suc:
            raise FileNotFoundError(f"Nenhum CSV encontrado em {os.path.join(raiz, j, 'sucessos')}")
        for f in arqs_suc:
            mes = mes_do_arquivo(f)
            if mes is None:
                continue
            partes_suc.append(
                pd.Series({mes: _ler_sucessos(f, mes)}, name="Sucessos", dtype="int64")
            )

    series_suc = partes_suc + sucessos_rpa
    if not series_suc:
        raise FileNotFoundError("Nenhum dado de sucesso encontrado para a jornada")
    sucessos_mes = (
        pd.concat(series_suc).groupby(level=0).sum().rename("Sucessos")
    )

    falhas_mes = df.groupby("Mes").size().rename("Falhas")
    resumo = pd.DataFrame({"Falhas": falhas_mes, "Sucessos": sucessos_mes}).fillna(0)
    resumo["Total"] = resumo["Falhas"] + resumo["Sucessos"]
    resumo["Pct"] = resumo["Falhas"] / resumo["Total"] * 100

    return df, df_octane, resumo, jornadas_combo


def categorizar(df, mes, hoje=None):
    """
    Divide os registros de falha de um mês nas categorias de negócio do fallout.
    Retorna (df_mes, cats, total), onde `cats` é um dict {nome_categoria: DataFrame}.

    `hoje` pode ser fixado (ex.: para reprodutibilidade em testes); por padrão
    usa a data/hora atual normalizada para meia-noite UTC.
    """
    if hoje is None:
        hoje = pd.Timestamp.now(tz="UTC").normalize()
    quinze_dias = hoje - pd.Timedelta(days=15)
    df_mes = df[df["Mes"] == mes].copy()

    milestone_dt = pd.to_datetime(df_mes["DFT_BugfixMilestone"], utc=True, errors="coerce")
    _corrigido = df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO)
    _tem_ms = df_mes["DFT_BugfixMilestone"].notna()
    _encerrado = _corrigido & _tem_ms & (milestone_dt <= hoje)
    _rotulo = df_mes["DefectNumber_orig"].str.strip().str.lower()
    _outros_mask = _rotulo == "enviado e-mail - outros times"
    _erro_proc_mask = _rotulo == "erro de processo"
    # No relatório de PME, "Enviado e-mail" (sem sufixo) significa que o caso
    # está em avaliação por MOPs — é rótulo diferente de "…- Outros Times".
    _email_mops_mask = _rotulo == "enviado e-mail"

    # Jornadas RPA controlam parte dos defeitos fora do Octane ("Paliativo RPA - 6",
    # "NFCOM"), então a marca de "tem defeito" vem da leitura e não do número.
    if "TemDefeito" in df_mes.columns:
        _tem_defeito = df_mes["TemDefeito"].fillna(False).astype(bool)
    else:
        _tem_defeito = (df_mes["DefectNumber__c"] > 0) & (df_mes["DefectNumber__c"] != 999999)

    _fase = df_mes["DFT_Phase"].fillna("").str.strip()

    # Legado em aberto é exclusivo do Base Residencial, marcado na leitura a
    # partir da Classificação Original — não vem de fase do Octane. A marca sai
    # de lá, e não do valor da coluna aqui, porque "Cancelado" também aparece no
    # SubStatus do Prospect, onde a regra não vale.
    if "LegadoAberto" in df_mes.columns:
        _legado = df_mes["LegadoAberto"].fillna(False).astype(bool)
    else:
        _legado = pd.Series(False, index=df_mes.index)

    cats = {
        "Em Tratamento/Avaliação pela Squad": df_mes[
            _tem_defeito &
            ~_fase.isin(FASES_FORA_TRATAMENTO) &
            ~_legado &
            ~_email_mops_mask &
            ~_encerrado
        ],
        "Resolvido": df_mes[_corrigido & _tem_ms & (milestone_dt < quinze_dias)],
        "Falha Pontual": df_mes[df_mes["DefectNumber__c"] == 999999],
        "Falta Associar ao Defeito/US": df_mes[
            ~_tem_defeito &
            ~_legado &
            (df_mes["DefectNumber__c"] != 999999) &
            ~_outros_mask & ~_erro_proc_mask & ~_email_mops_mask
        ],
        "Tratado - Em avaliação de eficácia": df_mes[
            _corrigido & _tem_ms & (milestone_dt >= quinze_dias) & (milestone_dt <= hoje)
        ],
        "Em Avaliação por MOPs": df_mes[
            (_tem_defeito & _fase.isin(FASES_MOPS)) | _email_mops_mask
        ],
        "Legado Aberto": df_mes[_legado],
        "Em avaliação - Outros times": df_mes[_outros_mask],
        "Falha no Processo Usuário": df_mes[_erro_proc_mask],
    }
    total = len(df_mes)
    return df_mes, cats, total
