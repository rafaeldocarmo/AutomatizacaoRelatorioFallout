# -*- coding: utf-8 -*-
"""
Núcleo de dados compartilhado entre app.py (dashboard interativo) e pipeline.py
(geração de relatório/PNG/PPTX): leitura das bases, join com Octane e
categorização do fallout por mês.

Este módulo não depende de Streamlit nem de estado global — todas as funções
recebem `base_dir` explicitamente e retornam dados, para poder ser chamado
com segurança a partir de múltiplas sessões/threads simultâneas.
"""
import contextlib
import csv
import datetime as _dt
import functools
import glob
import io
import json
import os
import re
import threading
import zipfile
from xml.etree import ElementTree

import pandas as pd

ABREV_MES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
             "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
MESES_PT = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
            7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
ABREV_POR_MES = {num: abrev for abrev, num in ABREV_MES.items()}


# ── Chave de mês ──────────────────────────────────────────────────────────────
# A chave é AAAAMM (202608). Antes era só o número do mês, o que bastava
# enquanto o histórico era "os CSVs que estão na pasta" — todos do mesmo ano.
# Com o histórico digitado a série passa de 12 meses e jan/25 colidiria com
# jan/26. Inteiro em vez de Period do pandas para continuar servindo de índice,
# ordenar sozinho e sobreviver à ida e volta do Excel.
def chave_mes(ano, mes):
    return int(ano) * 100 + int(mes)


def ano_do(chave):
    return int(chave) // 100


def mes_do(chave):
    return int(chave) % 100


def mes_anterior(chave):
    ano, mes = ano_do(chave), mes_do(chave)
    return chave_mes(ano - 1, 12) if mes == 1 else chave_mes(ano, mes - 1)


def proximo_mes(chave):
    ano, mes = ano_do(chave), mes_do(chave)
    return chave_mes(ano + 1, 1) if mes == 12 else chave_mes(ano, mes + 1)


def chave_de_datas(serie):
    """Série de datas (já no fuso de Brasília) -> série de chaves AAAAMM."""
    return serie.dt.year * 100 + serie.dt.month


# Dias distintos com falha para um mês ter "corpo" na comparação entre
# fotografias (ver comparacao._maduro): comparar um mês de dois dias numa delas
# com o mesmo mês inteiro na outra dava variações sem sentido. Já não decide o
# mês do relatório — ver `mes_do_relatorio`.
MIN_DIAS_MES = 7


def mes_do_relatorio(df, resumo):
    """
    O mês do relatório: o mais recente que tem falhas na extração — em todas
    as jornadas da visão — e algum sucesso no resumo; no mês fechado digitado só
    com a taxa, sem sucesso nenhum, vale a taxa. Sem nenhum assim, o último do
    resumo; None com o resumo vazio.

    Antes o mês novo só virava o do relatório com MIN_DIAS_MES dias de dados (e,
    sem isso, com 10 registros): em 08/10 o volume de outubro não aparecia nos
    slides, e o Base + Cross Sell, com set/26 já fechado no histórico, seguia
    comparando set/26 com set/26. A exigência de sucesso é o que sobrou da
    trava — no primeiro dia, com falhas e sem o arquivo de sucessos, o mês
    sairia com 100% de fallout.
    """
    if not len(resumo.index):
        return None
    # Numa visão combinada, o mês tem de estar em todas as jornadas dela: em
    # 08/10 o Prospect ainda não tinha outubro, e o Consolidado de out/26 saía
    # com 2,53% — só o Base + Cross Sell.
    meses_df = None
    for _jornada, parte in (df.groupby("Jornada") if "Jornada" in df.columns else [(None, df)]):
        meses = set(parte["Mes"].dropna().astype(int).unique().tolist())
        meses_df = meses if meses_df is None else meses_df & meses
    meses_df = meses_df or set()
    validos = []
    for m in resumo.index:
        if int(m) not in meses_df:
            continue
        sucessos = resumo.loc[m, "Sucessos"]
        if pd.isna(sucessos) or sucessos > 0:
            validos.append(int(m))
    return max(validos) if validos else int(max(resumo.index))


def dias_por_mes(df):
    """
    Dias distintos com falha em cada mês (chave AAAAMM), no fuso de Brasília.

    Dias, e não registros: um mês com 300 falhas em dois dias continua sendo um
    mês de dois dias, e é isso que segura a virada do mês.
    """
    datas = df["CreatedDate"].dt.tz_convert("America/Sao_Paulo")
    tabela = pd.DataFrame({"mes": chave_de_datas(datas), "dia": datas.dt.date}).dropna()
    # A chave sai como float quando há linha sem data; o índice tem que ser int
    # para casar com as chaves inteiras que vêm do resumo.
    return tabela.astype({"mes": "int64"}).groupby("mes")["dia"].nunique()


def fmt_inteiro(valor):
    """1234 -> '1.234'. Devolve "" no ausente: os meses que vêm da planilha de
    acompanhamento têm a taxa mas não o volume, e a célula fica em branco em
    vez de exibir um número que ninguém digitou."""
    if valor is None or valor != valor:      # NaN
        return ""
    return f"{int(valor):,}".replace(",", ".")


def fmt_pct(valor):
    """2.6 -> '2,60%'. Devolve "" no ausente, como `fmt_inteiro`."""
    if valor is None or valor != valor:
        return ""
    return f"{valor:.2f}%".replace(".", ",")


class _RotulosMes(dict):
    """
    `MESES_LABEL[202608]` -> "ago-26".

    É um dict com `__missing__` em vez de uma função para os pontos de chamada
    seguirem escrevendo `MESES_LABEL[m]` como sempre fizeram — são dez lugares
    entre pipeline, gerar_pptx e gerar_slide3col.
    """

    def __missing__(self, chave):
        return f"{ABREV_POR_MES[mes_do(chave)]}-{ano_do(chave) % 100:02d}"


MESES_LABEL = _RotulosMes()

# Vermelho dos cabeçalhos de slide e do dashboard. É o mesmo acento do app
# (COR_ACCENT em app.py), para o PowerPoint e a tela não saírem com dois
# vermelhos diferentes. Estava escrito à mão em três arquivos.
COR_VERMELHO = "#EC3013"
# Verde da linha "Planejado" — a meta digitada no histórico, desenhada tanto no
# slide de 3 colunas quanto no de jornada única.
COR_VERDE = "#2E7D32"

# "Resolvido" é a fase dos chamados do Jira (ver ARQ_JIRA), cuja base só tem
# "Em Atendimento" e "Resolvido". Pôr o segundo aqui é o que faz a regra dos 15
# dias de `categorizar` valer para ele sem nenhum caso especial.
# "Done" e "Implantado Prod" são as fases de entrega do relatório de US: sem
# elas, os pedidos de uma US já em produção ficavam em Tratamento para sempre,
# em vez de passarem por Tratado e Resolvido como os de um DFT corrigido. A US
# não preenche a Bugfix Milestone; a data de entrega dela sai de `_entrega_us`.
FASES_US_ENTREGUE = {"Done", "Implantado Prod"}
# "Concluído" é o equivalente no relatório de chamados BRJ (ver ARQ_BRJ).
FASES_BRJ_ENTREGUE = {"Concluído"}
FASES_CORRIGIDO = ({"Corrigido", "Fechado", "Resolvido"} | FASES_US_ENTREGUE
                   | FASES_BRJ_ENTREGUE)
FASES_MOPS = {"Cancelado", "Rejeitado"}
# "Legado Aberto" existe apenas no Base Residencial, marcado na leitura a partir
# da Classificação Original. Nas jornadas de extração o defeito cancelado no
# Octane continua sendo avaliação por MOPs, como sempre foi.
FASES_FORA_TRATAMENTO = FASES_MOPS

# "BSIM" é exclusiva do Cross Sell: pedido cujo defeito é do time de Confirmação
# Eletrônica, ainda em andamento, sai da categoria em que estaria e vai para cá.
# Entregue a correção, o pedido segue o caminho de qualquer defeito corrigido
# (Tratado, depois Resolvido) — o BSIM é para o que ainda depende do time. O mesmo time aparece
# no Prospect, onde a regra não vale — por isso a marca depende da jornada, e a
# jornada precisa ser gravada na leitura: o `df` junta várias pastas e depois
# não há como saber de onde cada linha veio.
CAT_BSIM = "BSIM"
JORNADA_BSIM = "Cross Sell"
TIME_BSIM = "Confirmação Eletrônica"

# "Incidente" é rótulo no lugar do número do defeito, como "Erro de Processo" e
# "Enviado e-mail": o pedido não tem DFT, o caso está aberto como incidente e
# tem categoria própria. Ainda não apareceu nenhum na extração — a regra entra
# pronta para quando aparecer.
CAT_INCIDENTE = "Incidente"
# "Legado Aberto" existe apenas no Base Residencial, com duas subcategorias
# (ver `subgrupo_legado`), as duas marcadas na leitura:
#   "Feito manualmente no Legado" — a origem marca CANCELADO (`LegadoAberto`);
#   "Funcionalidade idêntica ao legado" — o "Erro agrupado" começa com
#   "ErroProcesso" (`ErroLegado`), o mesmo erro do legado.
# Cancelado e ErroProcesso ao mesmo tempo é "Feito manualmente". A categoria
# tem precedência sobre as outras — com ou sem defeito associado —, fica fora
# da tabela de ofensores (ver pipeline) e sai inteira da "Projeção s/ Legado
# Aberto" do slide. Até 15/10 as duas eram categorias separadas, e a primeira
# se chamava "Legado Aberto".
CAT_LEGADO = "Legado Aberto"
SUB_LEGADO_MANUAL = "Feito manualmente no Legado"
SUB_LEGADO_FUNCIONALIDADE = "Funcionalidade idêntica ao legado"
CAT_ERRO_LEGADO = SUB_LEGADO_FUNCIONALIDADE        # nome antigo, das fotografias
# Também só no Base Residencial (desde 11/11/2026), marcadas na leitura
# (`BaseResidencial` e `ErroAgendamento`):
#   o defeito 999999 é erro do script da automação, e não Falha Pontual — no
#   export de 09/11 os 129 pedidos 999999 são exatamente os de "Erro no script
#   da automação" no "Erro agrupado";
#   o "Erro agrupado" que começa com "ErroAgendamento" é erro no agendamento,
#   com precedência sobre as demais categorias menos o Legado Aberto (o
#   cancelado continua lá, como na Funcionalidade idêntica ao legado);
#   o pedido sem defeito associado é Falha Pontual, e não "Falta Associar".
CAT_SCRIPT_RPA = "Erro no Script da Automação - RPA"
CAT_AGENDAMENTO = "Erro no Agendamento"
# Comparados sem espaços e em minúsculo, como o NBA_PREFIXO_ERRO_LEGADO.
NBA_PREFIXOS_AGENDAMENTO = ("erroagendamento", "erronoagendamento")
# A primeira categoria da Distribuição. Ela e o Legado Aberto têm subgrupos.
CAT_TRATAMENTO = "Em Tratamento/Avaliação pela Squad"
# Os subgrupos do Em Tratamento, em dois níveis: o tipo do item do Octane (US
# ou DFT) e, dentro dele, com ou sem data de entrega (ver `subgrupo_tratamento`).
# Os rótulos do segundo nível levam o tipo na frente para serem únicos — as
# fotografias e a comparação guardam a Distribuição por rótulo —, e quem
# desenha mostra só o "c/ data"/"s/ data" (ver `rotulo_dist`). Até 11/11/2026
# eram três: "Planejado" (com data, de qualquer tipo), "US s/ data" e
# "DFT s/ data"; os dois últimos continuam com o mesmo nome e significado.
SUB_TRAT_US, SUB_TRAT_DFT = "US", "DFT"
SUBGRUPOS_TRATAMENTO = {
    SUB_TRAT_US: ("US c/ data", "US s/ data"),
    SUB_TRAT_DFT: ("DFT c/ data", "DFT s/ data"),
}
ROTULO_INCIDENTE = "incidente"          # comparado em minúsculo, como os demais
# Aba "Precisa de Correção": fases em que ninguém mais vai mexer no defeito.
# Serve para rotular o pedido, e não para escolhê-lo: "Fechado" está aqui, mas
# sozinho não põe ninguém na lista. Lá houve entrega, então o pedido criado
# antes dela já foi resolvido por ela — só vira caso quando é criado depois, e
# aí entra pelo critério do pós-milestone. Cancelado e Rejeitado é que entram em
# qualquer data: nenhuma correção vai chegar, e a Bugfix Milestone de um defeito
# cancelado é só um resto de planejamento (no Cross Sell eram 532 pedidos de
# março atribuídos a DFTs cancelados com entrega prevista para junho e julho).
FASES_ENCERRADAS = {"Fechado"} | FASES_MOPS
# Fases do relatório de US em que a US saiu de cena: entregue ou cancelada. Nas
# demais ela está em andamento (ver `us_em_aberto`).
FASES_US_ENCERRADAS = FASES_US_ENTREGUE | {"Liberado para Uso", "cancelar"}
# Fase única das jornadas com catálogo próprio (NBA): todo pedido com defeito
# associado está em tratamento. O BRJ que está no RelatorioBRJ (ver ARQ_BRJ)
# troca esta fase pelo status de lá.
FASE_TRATAMENTO_NBA = "Em Tratamento"

# Pasta da jornada → nome mostrado no app e no título do slide. A pasta segue
# com o nome da origem do arquivo; só a exibição muda.
NOMES_EXIBICAO = {"NBA": "Base Residencial", "PME": "PME Móvel",
                  "PME Fixa": "PME Fixa"}

# Jornadas que usam o formato de extração mas não compõem o Consolidado. A
# PME Fixa nasceu depois da série histórica: somá-la faria o Consolidado de
# agosto deixar de ser comparável com os meses anteriores, que vêm da planilha
# de acompanhamento e não a incluem.
FORA_DO_CONSOLIDADO = {"PME Fixa"}

# Pasta local (dentro do projeto) com as jornadas e as planilhas do Octane.
PASTA_DADOS = "extracoes"
# Pasta com "_" na frente, dentro dos dados, é do próprio app e não jornada —
# hoje, só a das fotografias (ver fotografia.py). A descoberta de jornadas pula
# essas pastas: uma planilha solta lá dentro faria a pasta parecer uma jornada
# RPA e aparecer no seletor.
PREFIXO_INTERNO = "_"
# Espelho da pasta do Google Drive, dentro dos dados (ver drive_sync.py). Vale
# no lugar das extrações locais quando a chave do Drive está ao lado do app e a
# fonte escolhida no app (ARQ_FONTE) não é o PC.
PASTA_DRIVE = "_drive"
ARQ_CHAVE_DRIVE = "drive_chave.json"
ARQ_FONTE = "fonte_dados.txt"
FONTE_PC, FONTE_DRIVE = "pc", "drive"
ARQ_DFT = "RelatorioDFTOctane.xlsx"
# ID do item do Octane como texto em maiúsculas, no catálogo: a chave dos que
# não têm número ("DDP-366"), casados pelo rótulo do pedido (ver carregar_base).
COL_CHAVE_OCTANE = "_ChaveOctane"
ARQ_US = "RelatorioUSOctane.xlsx"

# ── Chamados BRJ ─────────────────────────────────────────────────────────────
# Extração do Jira que faz o papel do Octane para os BRJ-* do Base Residencial
# e da PME: o status vira a fase e o campo personalizado de data — a data de
# implantação — vira a Bugfix Milestone. Opcional, como o do Jira: sem ele, o
# Base Residencial segue com o catálogo da própria planilha (ver `_catalogo_nba`).
ARQ_BRJ = "RelatorioBRJ.xlsx"
BRJ_ABA = "Defeitos - BRJ"
BRJ_COL_CHAVE = "Chave da item"
BRJ_COL_STATUS = "Status"
BRJ_COL_RESUMO = "Resumo"
BRJ_COL_RESPONSAVEL = "Responsável"
BRJ_COL_TIPO = "Tipo de item"
BRJ_COL_RESOLVIDO = "Resolvido"
# O Jira exporta o campo como "Campo personalizado (Data e Hora Limite)"; a
# coluna é achada pelo prefixo, que é a parte que o export mantém.
BRJ_PREFIXO_DATA = "Campo personalizado"
BRJ_TIPOS_US = {"História", "Epic"}

# ── Chamados do Jira ─────────────────────────────────────────────────────────
# Base de defeitos bem mais simples que a do Octane: chamado, status e data de
# fechamento, só isso. Cobre justamente o que o Octane não conhece — todo pedido
# que casa aqui estava sem fase nenhuma, parado em "Falta Associar ao Defeito".
ARQ_JIRA = "RelatorioJira.xlsx"
JIRA_COL_CHAMADO = "Chamado"
JIRA_COL_STATUS = "Status"
JIRA_COL_FECHAMENTO = "Data fechamento"
JIRA_COLS = (JIRA_COL_CHAMADO, JIRA_COL_STATUS, JIRA_COL_FECHAMENTO)
# A aba e a linha do cabeçalho não são fixas: o export já veio como "Planilha2"
# com uma coluna vazia à esquerda (cabeçalho na segunda linha) e como
# "Chamados JIRA" começando direto na primeira. Em vez de fixar as duas coisas,
# procuramos as três colunas acima nas primeiras linhas de cada aba.
JIRA_MAX_LINHAS_CABECALHO = 10
# Só PDST e INC contam. A mesma planilha lista GML-, PRB- e um número solto, que
# referenciam outros sistemas e seguem sem defeito associado.
JIRA_PREFIXOS = ("PDST-", "INC-")
JIRA_RESOLVIDO = "Resolvido"
JIRA_EM_ATENDIMENTO = "Em Atendimento"
# Igual à fase do Octane de propósito: estando em FASES_MOPS, o chamado
# cancelado cai em "Em Avaliação por MOPs" e na aba "Precisa de Correção" como
# qualquer DFT cancelado.
JIRA_CANCELADO = "Cancelado"

COLUNAS_OCTANE_BASE = ["ID", "Name", "Phase", "Bugfix Milestone", "Team", "Type"]
COL_US_MELHORIA = "US de Melhoria"
COL_US_RESOLVIDA = "_us_resolvida"

# ── Jornadas em formato "RPA" ────────────────────────────────────────────────
# Algumas jornadas (PME) não vêm da extração Salesforce/Vlocity e sim de um
# relatório de RPA: um CSV único, separado por ';' e em cp1252, com sucessos e
# falhas na mesma planilha. O de-para abaixo traduz para o formato interno.
RPA_SEP = ";"
RPA_ENCODING = "cp1252"
# O export em XLSX (aba "Base") também é lido, mas é lento, e vira um CSV com
# este sufixo na primeira leitura (ver `_gravar_espelho`). Esse CSV sai em UTF-8,
# e não em cp1252 como o do RPA, porque o XLSX traz caractere que o cp1252 não
# tem — um espaço de largura zero dentro da mensagem de erro, que é a chave de
# agrupamento da aba de erros e não pode ser trocado.
RPA_SUF_CONVERTIDO = "__base.csv"
RPA_ENCODING_CONVERTIDO = "utf-8-sig"
RPA_COL_RESULTADO = "SUCESSO OU FALHA"   # coluna U
RPA_VALOR_FALHA = "Erro"
RPA_COL_DATA = "DATA_CRIACAO"
RPA_ABA = "Base"                          # aba do export em Excel
NBA_ABA_ANALITICO = "Analitico"           # aba que identifica o export do NBA
NBA_ABA_DEFEITOS = "Defeitos"
RPA_COL_ENTREGA = "DATA CORREÇÃO"        # coluna S
RPA_COL_DEFEITO = "DEFEITO"              # coluna M
# Nome do defeito para os rótulos que o Octane não conhece ("Paliativo RPA - 7",
# "NFCOM", GML-…): a descrição mais frequente do rótulo (ver carregar_base).
RPA_COL_DESCRICAO = "DESCRIÇÃO DO DEFEITO"
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
    ROTULO_INCIDENTE,                  # → Incidente
}


def _encoding_rpa(caminho):
    """Encoding de um CSV da jornada RPA: o convertido é UTF-8, o original cp1252."""
    if caminho.endswith(RPA_SUF_CONVERTIDO):
        return RPA_ENCODING_CONVERTIDO
    return RPA_ENCODING


def _espelho_rpa(xlsx):
    """CSV gerado a partir de um export XLSX da jornada RPA."""
    return os.path.splitext(xlsx)[0] + RPA_SUF_CONVERTIDO

# ── Jornadas em formato "NBA" ────────────────────────────────────────────────
# Planilha XLSX com duas abas: "Analitico" (um pedido por linha) e "Defeitos",
# que é o catálogo local — aqui os defeitos são BRJ-*, não existem no Octane.
NBA_ABA_ANALITICO = "Analitico"
NBA_ABA_DEFEITOS = "Defeitos"
# Ler XLSX com openpyxl é lento (~40 s para os três meses). Cada planilha vira
# um par de CSVs na primeira leitura (ver `_gravar_espelho`), que o leitor
# prefere daí em diante. O converter_nba.py faz o mesmo à mão.
NBA_SUF_ANALITICO = "__analitico.csv"
NBA_SUF_DEFEITOS = "__defeitos.csv"
NBA_ENCODING_CONVERTIDO = "utf-8-sig"


def _espelhos_nba(xlsx):
    """Par de CSVs (analítico, defeitos) gerado a partir de um export XLSX do NBA."""
    base = os.path.splitext(xlsx)[0]
    return base + NBA_SUF_ANALITICO, base + NBA_SUF_DEFEITOS
NBA_COL_STATUS = "Status"
NBA_VALOR_SUCESSO = "SUCESSO"
NBA_VALOR_CANCELADO = "CANCELADO"        # vira a categoria "Legado Aberto"
# Início do "Erro agrupado" que vira CAT_ERRO_LEGADO. Comparado sem espaços e
# em minúsculo: a planilha escreve "ErroProcesso - …", e "Erro Processo" também
# tem que valer.
NBA_PREFIXO_ERRO_LEGADO = "erroprocesso"
NBA_COL_DATA = "Data"
NBA_COL_DEFEITO = "Defeito"
NBA_COL_CRIACAO = "Data de Criação"      # aba Defeitos: quando o BRJ foi aberto
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

# Sucessos, em três layouts: até jul/26 um agregado de uma linha só (coluna
# "expr0" com o total do mês); em ago/26 um relatório com cabeçalho de filtros e
# uma linha por dia ("01/08/2026,1555"); de set/26 em diante o agregado do
# Salesforce quebrado por dia — "expr0" com o total e "expr1" com o dia
# ("2026-09-01"). Os dois últimos são diários e passam pelo corte.
SUC_COL_AGREGADA = "expr0"
SUC_COL_DIA = "expr1"
SUC_DATA_DIA = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
SUC_DATA_ISO = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")


RE_MES_ARQUIVO = re.compile("(" + "|".join(ABREV_MES) + r")(\d{2})")
# Os exports do portal trazem a data da extração no nome, em vez do mês:
#   "...Base (Orch.Item)-2026-08-27-12-06-12.xlsx"  ->  2026-08
#   "Completo_26_08_2026.xlsx"                      ->  2026-08
RE_DATA_ISO = re.compile(r"(20\d{2})-(\d{2})-(\d{2})")
RE_DATA_BR = re.compile(r"(\d{2})_(\d{2})_(20\d{2})")


def mes_do_arquivo(path):
    """
    Chave AAAAMM a partir do nome do arquivo (ex.: '...jan26.csv' -> 202601).

    O ano sai dos dois dígitos que seguem a abreviação; sem eles não dá para
    distinguir jan/25 de jan/26 quando a série passa de um ano.
    """
    nome = os.path.splitext(os.path.basename(path))[0].lower()
    achado = RE_MES_ARQUIVO.search(nome)
    if achado:
        return chave_mes(2000 + int(achado.group(2)), ABREV_MES[achado.group(1)])
    achado = RE_DATA_ISO.search(nome)
    if achado:
        return chave_mes(int(achado.group(1)), int(achado.group(2)))
    achado = RE_DATA_BR.search(nome)
    if achado:
        return chave_mes(int(achado.group(3)), int(achado.group(2)))
    return None


def data_do_nome(path):
    """
    Dia da extração escrito no nome do arquivo, ou None se não houver um.

    "Completo_16_09_2026.xlsx" -> 16/09/2026. Só os dois formatos com o dia
    completo contam; "set26" diz o mês, e não quando o arquivo foi tirado.
    """
    nome = os.path.basename(path)
    achado = RE_DATA_ISO.search(nome)
    if achado:
        ano, mes, dia = (int(g) for g in achado.groups())
    else:
        achado = RE_DATA_BR.search(nome)
        if not achado:
            return None
        dia, mes, ano = (int(g) for g in achado.groups())
    try:
        return _dt.date(ano, mes, dia)
    except ValueError:
        return None


# ── Qual export é o mais recente ─────────────────────────────────────────────
# As pastas acumulam exports, e em todo formato vale o mais recente. A data de
# modificação do arquivo não serve para decidir isso: ela não sobrevive ao envio
# da pasta. Extrair um .zip baixado (e-mail, Teams, navegador) pelo Explorer
# grava em cada arquivo a hora da extração, na ordem alfabética, e aí "o mais
# recente" vira o último nome do alfabeto — foi assim que quem recebeu a pasta
# via o Cross Sell só até 11/09, com o export de 24/09 ao lado. Por isso cada
# formato decide pela data que está nos dados (ou, no NBA, no nome), e a data de
# modificação só desempata dois exports do mesmo dia.
SEM_DATA = _dt.date.min   # ordena atrás de qualquer export com data


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
    """Subpastas de `caminho`, sem as do próprio app (ver PREFIXO_INTERNO)."""
    if not os.path.isdir(caminho):
        return []
    return sorted(d for d in os.listdir(caminho)
                  if os.path.isdir(os.path.join(caminho, d))
                  and not d.startswith(PREFIXO_INTERNO))


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
    # A pasta com os relatórios do Octane é a dos dados, e não uma jornada. Com
    # o RelatorioParametros.csv solto ali, a regra do "*.csv" abaixo a tomava por
    # RPA: copiada aninhada (extracoes\extracoes\...), ela virava a única
    # "jornada", sem data nenhuma, e o PowerPoint quebrava sem data de corte.
    if any(os.path.isfile(os.path.join(pasta, arq)) for arq in (ARQ_DFT, ARQ_US)):
        return None
    # O par __defeitos.csv identifica o NBA já convertido.
    if _arquivos(pasta, "*" + NBA_SUF_DEFEITOS):
        return "nba"
    # E o __base.csv, o RPA convertido — que só sai de um XLSX com a aba "Base".
    # Olhar ele antes poupa abrir o XLSX de 43 MB só para ler o nome das abas.
    if _arquivos(pasta, "*" + RPA_SUF_CONVERTIDO):
        return "rpa"
    # Os dois formatos exportam .xlsx, então a extensão não distingue: o NBA
    # traz as abas "Analitico"/"Defeitos" e o RPA traz "Base".
    for arq in _arquivos(pasta, "*.xlsx"):
        abas = _abas(arq)
        if NBA_ABA_ANALITICO in abas:
            return "nba"
        if RPA_ABA in abas:
            return "rpa"
    if _arquivos(pasta, "*.csv"):
        return "rpa"
    return None


_NS_PLANILHA = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@functools.lru_cache(maxsize=64)
def _abas(caminho):
    """
    Nomes das abas de um .xlsx, sem carregar as células.

    Lê só o `xl/workbook.xml` de dentro do zip: o `pd.ExcelFile` abria o
    workbook pelo openpyxl, que no export de 43 MB da PME levava ~1,5 s só para
    dizer o nome das abas. O que não sair assim (outro namespace, arquivo
    estranho) cai no `pd.ExcelFile`, que é o caminho de antes.
    """
    try:
        with zipfile.ZipFile(caminho) as z:
            raiz = ElementTree.fromstring(z.read("xl/workbook.xml"))
        abas = raiz.find("m:sheets", _NS_PLANILHA)
        nomes = tuple(a.get("name") for a in abas.findall("m:sheet", _NS_PLANILHA))
        if nomes and all(nomes):
            return nomes
    except Exception:                                    # noqa: BLE001
        pass
    try:
        # `with`: o ExcelFile segura o arquivo aberto até ser fechado, e no
        # Windows isso impede trocar a planilha enquanto o app está no ar.
        with pd.ExcelFile(caminho) as planilha:
            return tuple(planilha.sheet_names)
    except Exception:                                    # noqa: BLE001
        return ()


def nome_exibicao(jornada):
    """Nome que o usuário vê; por padrão é o próprio nome da pasta."""
    return NOMES_EXIBICAO.get(jornada, jornada)


def _tem_jornadas(caminho):
    """True se a pasta contém diretórios de jornada, em qualquer um dos formatos."""
    return any(formato_jornada(os.path.join(caminho, d)) for d in _subpastas(caminho))


def drive_disponivel(base_dir):
    """True se a chave do Google Drive está ao lado do app (ver drive_sync.py)."""
    return os.path.isfile(os.path.join(base_dir, ARQ_CHAVE_DRIVE))


def fonte_escolhida(base_dir):
    """
    A fonte dos dados escolhida no app: FONTE_DRIVE ou FONTE_PC. Fica num
    arquivo, e não no estado da sessão, porque vale para tudo o que lê os dados
    (PowerPoint, fotografia, linha de comando). Sem escolha gravada, o Drive
    ganha quando está disponível.
    """
    try:
        with open(os.path.join(base_dir, ARQ_FONTE), encoding="utf-8") as f:
            escolha = f.read().strip().lower()
    except OSError:
        escolha = ""
    if escolha not in (FONTE_PC, FONTE_DRIVE):
        escolha = FONTE_DRIVE
    return escolha if drive_disponivel(base_dir) else FONTE_PC


def gravar_fonte(base_dir, fonte):
    with open(os.path.join(base_dir, ARQ_FONTE), "w", encoding="utf-8") as f:
        f.write(fonte)


def usa_drive(base_dir):
    """True se os dados vêm do espelho do Google Drive, e não de `extracoes/`."""
    return fonte_escolhida(base_dir) == FONTE_DRIVE


def raiz_dados(base_dir):
    """
    Pasta que guarda as jornadas e as planilhas do Octane: `base_dir/extracoes`,
    ou o espelho do Google Drive (`extracoes/_drive`) quando a fonte escolhida é
    o Drive. Sem volta para o PC quando o espelho está vazio: com o Drive
    escolhido, mostrar os dados do PC seria mostrar a fonte errada.

    Se as jornadas não estiverem soltas ali dentro e sim agrupadas em um único
    diretório (é o que acontece ao copiar a pasta exportada inteira, ex.:
    extracoes/RelatoriosAutomatização/...), desce mais um nível automaticamente.
    """
    raiz = os.path.join(base_dir, PASTA_DADOS)
    if usa_drive(base_dir):
        raiz = os.path.join(raiz, PASTA_DRIVE)
    if _tem_jornadas(raiz):
        return raiz
    for d in _subpastas(raiz):
        if _tem_jornadas(os.path.join(raiz, d)):
            return os.path.join(raiz, d)
    return raiz


def jornadas_disponiveis(base_dir):
    """Pastas de jornada disponíveis, em qualquer um dos formatos suportados."""
    raiz = raiz_dados(base_dir)
    return [d for d in _subpastas(raiz) if formato_jornada(os.path.join(raiz, d))]


def visoes_disponiveis(base_dir):
    """
    Tudo o que o seletor de jornada do app oferece: cada jornada, mais
    "Base + Cross Sell" e "Consolidado" quando as partes existem.

    Mora aqui, e não no app, porque a fotografia (fotografia.py) precisa gravar
    exatamente as visões que o seletor mostra — duas listas mantidas à mão
    acabariam divergindo.
    """
    jornadas = jornadas_disponiveis(base_dir)
    visoes = list(jornadas)
    if "Base Móvel" in jornadas and "Cross Sell" in jornadas:
        visoes.append("Base + Cross Sell")
    if len(jornadas) > 1:
        visoes.append("Consolidado")   # todas as jornadas juntas
    return visoes


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
        and j not in FORA_DO_CONSOLIDADO
    ]


def resolver_jornadas(jornada, base_dir):
    """Traduz o nome escolhido nas pastas que devem ser lidas."""
    if jornada == "Consolidado":
        return jornadas_consolidaveis(base_dir)
    if jornada == "Base + Cross Sell":
        return ["Base Móvel", "Cross Sell"]
    return [jornada]


def arquivos_de_entrada(base_dir):
    """
    Todo arquivo de dados que o app enxerga, como [{"arquivo", "modificado",
    "tamanho"}]: as planilhas da raiz (Octane, Jira, histórico) e o conteúdo
    das pastas de jornada, com o caminho relativo à raiz dos dados e a data de
    modificação em ISO.

    É o registro de onde saiu cada fotografia. Comparando duas, dá para dizer
    que o Octane não foi atualizado entre elas, ou que uma jornada não recebeu
    arquivo novo — e aí "nada mudou" quer dizer "ninguém atualizou", e não "o
    fallout ficou parado".

    Lista a pasta inteira, e não só o arquivo que cada leitor escolhe: a regra
    de escolha muda por formato (o mais recente no RPA, um por mês no NBA, um
    por mês e por etapa na extração) e já mora nos leitores. Para dizer o que
    mudou entre duas semanas, a pasta inteira basta.
    """
    raiz = raiz_dados(base_dir)
    caminhos = []
    for nome in (ARQ_DFT, ARQ_US, ARQ_JIRA, ARQ_BRJ):
        try:
            caminhos.append(_caminho_octane(raiz, base_dir, nome))
        except FileNotFoundError:
            pass                    # o Jira é opcional; a falta do Octane aparece na leitura
    caminhos.append(os.path.join(raiz, ARQ_HISTORICO))
    for j in jornadas_disponiveis(base_dir):
        for pasta, _subdirs, nomes in os.walk(os.path.join(raiz, j)):
            caminhos += [os.path.join(pasta, n) for n in nomes]

    registros = []
    for caminho in sorted(caminhos):
        nome = os.path.basename(caminho)
        # Fora: a trava do Excel, o temporário de uma conversão em andamento
        # (ver `_gravar_espelho`) e o índice de datas do RPA, que não são dados.
        if (not os.path.isfile(caminho) or nome.startswith(("~$", ".~"))
                or nome.endswith(".parcial") or nome == ARQ_INDICE_RPA):
            continue
        info = os.stat(caminho)
        registros.append({
            "arquivo": os.path.relpath(caminho, raiz).replace(os.sep, "/"),
            "modificado": _dt.datetime.fromtimestamp(info.st_mtime).isoformat(timespec="seconds"),
            "tamanho": info.st_size,
        })
    return registros


def _limpar_defeito(serie):
    """Texto do campo de defeito sem espaços não-quebráveis nem sobras."""
    return serie.fillna("").astype(str).str.replace("\xa0", " ", regex=False).str.strip()


# ── Corte no último dia completo ──────────────────────────────────────────────
# Uma extração feita no meio do dia que traga esse dia o traz pela metade: no
# exemplo de 13/08 os sucessos da Base caíram de 2.869 (dia 12) para 935.
# Deixar esse dia entrar afunda o total do mês e, como falhas e sucessos são
# contados de fontes diferentes, mexe no fallout rate de forma difícil de
# perceber. Por isso o corte vale para tudo que é lido — falhas e sucessos.

# "10_6_2026" no nome dos exports do Salesforce (mês_dia_ano, sem zero à
# esquerda) e "16_09_2026" nos do NBA (dia_mês_ano). Os dois cabem aqui; quem
# desfaz a ambiguidade é `_dia_da_extracao`.
RE_DATA_SUBLINHADO = re.compile(r"(?<!\d)(\d{1,2})_(\d{1,2})_(20\d{2})(?!\d)")


def _dia_da_extracao(caminho, fim_dos_dados):
    """
    Dia em que o arquivo foi tirado, pelo nome, ou None se o nome não diz.

    "…PF-2026-10-01-14-01-53.xlsx" é 01/10. Em "order-10_6_2026.csv" dia e mês
    podem estar em qualquer ordem; vale a leitura que não é anterior ao último
    dia que o próprio arquivo traz (`fim_dos_dados`) — ninguém tira um arquivo
    antes do dia em que os dados dele terminam. Sobrando duas, a mais cedo.
    """
    nome = os.path.basename(caminho)
    achado = RE_DATA_ISO.search(nome)
    if achado:
        try:
            return _dt.date(*(int(g) for g in achado.groups()))
        except ValueError:
            return None
    candidatos = []
    for a, b, ano in RE_DATA_SUBLINHADO.findall(nome):
        for dia, mes in ((int(a), int(b)), (int(b), int(a))):
            try:
                data = _dt.date(int(ano), mes, dia)
            except ValueError:
                continue
            if data >= fim_dos_dados:
                candidatos.append(data)
    return min(candidatos) if candidatos else None


def _max_data_pasta(pasta, formato):
    """
    (último dia, completo?) de uma jornada, lendo só a coluna de data.

    `completo` só é verdadeiro quando se sabe que a extração foi feita depois
    do último dia: algum arquivo que chega a esse dia traz no nome a data em
    que foi tirado, e ela é posterior. Tirado no próprio dia, ou sem data no
    nome, o dia conta como incompleto — era a regra de sempre.

    Cada máximo vira `datetime.date` na hora: os layouts trazem umas datas com
    fuso e outras sem, e comparar Timestamp aware com naive levanta exceção.
    """
    datas = []

    def _guardar(serie_datas):
        maximo = serie_datas.max()
        if pd.notna(maximo):
            datas.append(maximo.date())

    if formato == "extracao":
        fins = []                       # (último dia do arquivo, caminho)
        for f in _arquivos_da_etapa(pasta, "falhas", EXT_FALHAS):
            # Reusa o leitor: ele já resolve HTML/Excel/CSV e os dois formatos
            # de data. Custa mais do que ler só a coluna, mas o resultado fica
            # em cache (por arquivo) e é o mesmo que o relatório lê em seguida.
            try:
                maximo = (_ler_falhas_extracao(f)["CreatedDate"]
                          .dt.tz_convert("America/Sao_Paulo").max())
            except (KeyError, ValueError):
                continue
            if pd.notna(maximo):
                fins.append((maximo.date(), f))
        if not fins:
            return None, False
        ultimo = max(d for d, _f in fins)
        # Os sucessos não decidem o último dia, mas o nome deles costuma trazer
        # a hora da extração ("…-2026-10-01-14-01-53.xlsx"), e a das falhas não
        # ("Falhas PF.xls").
        for f in _arquivos_da_etapa(pasta, "sucessos", EXT_SUCESSOS):
            dias = [d for d, _total in _linhas_diarias(f)]
            if dias:
                fins.append((max(dias), f))
        tirados = [_dia_da_extracao(f, d) for d, f in fins if d == ultimo]
        tirados = [t for t in tirados if t is not None]
        return ultimo, bool(tirados) and all(t > ultimo for t in tirados)
    elif formato == "rpa":
        # Só o arquivo que o leitor usa. Antes entrava todo CSV da pasta —
        # exports antigos e espelhos de XLSX que perderam para um mais novo —, e
        # o corte mudava conforme o XLSX já tivesse sido convertido ou não. O
        # último dia dele já foi apurado na escolha (ver `_ultimo_dia_rpa`).
        escolhido, _ignorados = _escolher_rpa(pasta)
        dia = _ultimo_dia_rpa(escolhido)
        if dia is not None:
            datas.append(dia)
    else:
        # NBA: os exports que valem mês a mês, como no leitor (ver acima).
        for tipo, caminho in _escolher_nba(pasta)[0]:
            if tipo == "csv":
                serie = pd.read_csv(caminho + NBA_SUF_ANALITICO, encoding="utf-8",
                                    usecols=[NBA_COL_DATA], dtype=str,
                                    low_memory=False)[NBA_COL_DATA]
            else:
                serie = _ler_abas_nba(tipo, caminho)[0][NBA_COL_DATA]
            _guardar(pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce"))

    return (max(datas) if datas else None), False


def _assinatura_jornadas(base_dir):
    """
    Nome, tamanho e data de todo arquivo das pastas de jornada: a chave do
    cache do corte. Trocar um arquivo muda a chave e o corte é refeito — antes
    ele ficava preso até o "Atualizar dados", e a fotografia gravada logo
    depois de colocar uma extração nova saía com o corte da anterior.
    """
    raiz = raiz_dados(base_dir)
    itens = []
    for j in jornadas_disponiveis(base_dir):
        for pasta, _subdirs, nomes in os.walk(os.path.join(raiz, j)):
            for n in nomes:
                if n.startswith(PREFIXO_INTERNO):      # o índice do RPA, do próprio app
                    continue
                caminho = os.path.join(pasta, n)
                try:
                    info = os.stat(caminho)
                except OSError:
                    continue
                itens.append((os.path.relpath(caminho, raiz), info.st_size, info.st_mtime_ns))
    return tuple(sorted(itens))


@functools.lru_cache(maxsize=8)
def _fim_das_jornadas(base_dir, _assinatura):
    """
    (jornada, último dia, completo?) de cada jornada (ver `_max_data_pasta`).

    Cacheado porque varre as colunas de data de todas as jornadas (~3s com o
    volume atual); `_assinatura` só compõe a chave.
    """
    raiz = raiz_dados(base_dir)
    fins = []
    for j in jornadas_disponiveis(base_dir):
        pasta = os.path.join(raiz, j)
        try:
            d, completo = _max_data_pasta(pasta, formato_jornada(pasta))
        except (OSError, ValueError, KeyError) as e:
            # Uma jornada ilegível não pode derrubar o relatório inteiro; ela
            # apenas não participa da escolha do corte.
            print(f"  [corte] {j}: não deu para ler a data ({type(e).__name__}: {e})")
            continue
        if d is not None:
            fins.append((j, d, completo))
    return tuple(fins)


def fim_das_jornadas(base_dir):
    """(jornada, último dia, completo?) de cada jornada, com os arquivos de agora."""
    return _fim_das_jornadas(base_dir, _assinatura_jornadas(base_dir))


def data_corte(base_dir, hoje=None):
    """
    Último dia completo do conjunto: o mais recente entre as jornadas, contando
    em cada uma o último dia só quando se sabe que ele está completo, e nunca
    depois de ontem.

    Saber vem do nome dos arquivos (ver `_max_data_pasta`): a extração de 01/10
    do Prospect ("…-2026-10-01-14-01-53.xlsx") termina em 30/09, então 30/09 está
    completo; a de 06/10 do Base Móvel ("order-10_6_2026.csv") traz o próprio
    dia 06/10, pela metade, e ele sai. Sem data no nome, o último dia sai, como
    sempre foi.

    O corte depende só dos arquivos. Antes ele era "o último dia, mas nunca
    depois de ontem", e a mesma extração dava cortes diferentes conforme o dia
    da gravação: gravada em 06/10 era a fotografia de 05/10, e em 07/10 virava a
    de 06/10 — com o dia incompleto dentro. O "nunca depois de ontem" segue só
    como trava. `hoje` é a data de referência (ver `hoje_utc`); a fotografia de
    uma extração antiga passa o dia dela.

    O corte é único para o relatório inteiro em vez de um por jornada. As
    jornadas são extraídas em momentos diferentes (as de extração vão até 13/08,
    NBA e PME até 11/08); um corte por jornada descartaria de NBA e PME um dia
    que está completo e deixaria cada uma olhando uma janela diferente, o que
    quebra a comparação no Consolidado.

    O último dia sai dos pedidos com falha; dos sucessos só se lê quando a
    extração foi feita. Eles não decidem o dia nem quando saíram de uma
    extração mais antiga — se o arquivo de
    sucessos parar antes do corte, o dia que faltar entra com o volume que
    houver.

    Devolve `datetime.date` ou None quando não há data legível em lugar nenhum.
    """
    fins = fim_das_jornadas(base_dir)
    if not fins:
        return None
    um_dia = _dt.timedelta(days=1)
    corte = max(d if completo else d - um_dia for _j, d, completo in fins)
    return min(corte, hoje_utc(hoje).date() - um_dia)


def limpar_cache_corte():
    """
    Esquece a data de corte e todas as leituras em cache — usar depois de trocar
    os arquivos de dados.

    As leituras já se renovam sozinhas quando o arquivo muda (a chave é a data
    e o tamanho dele); limpar aqui é para o "Atualizar dados" reler tudo mesmo,
    inclusive um arquivo trocado por outro de mesma data e tamanho.
    """
    _fim_das_jornadas.cache_clear()
    _abas.cache_clear()
    _ultimo_dia_rpa_cacheado.cache_clear()
    for cache in (_linhas_sucessos_cacheadas, _octane_cacheado, _jira_cacheado, _brj_cacheado,
                  _falhas_extracao_cacheada, _jornada_rpa_cacheada,
                  _jornada_nba_cacheada):
        cache.cache_clear()


# ── CSV espelho dos XLSX lentos ──────────────────────────────────────────────
# Ler os exports XLSX da PME e do NBA com openpyxl leva de 10 s a mais de um
# minuto por arquivo, e isso se repetia em toda abertura do app, em todo
# "Atualizar dados" e em todo relatório. Por isso os leitores dessas jornadas,
# ao ler um XLSX, gravam ao lado dele um CSV com o mesmo conteúdo — o mesmo que
# converter_pme.py e converter_nba.py geram à mão —, e das vezes seguintes leem
# o CSV. Só o export que o leitor escolheu é convertido: os que perdem para um
# mais recente nunca são lidos e ficam como estão.
#
# O CSV herda a data de modificação do XLSX. Os leitores escolhem o export mais
# recente por essa data; com a data da conversão, um export antigo convertido
# depois passaria à frente de um mais novo.

# Diferença de data tolerada para o CSV contar como do mesmo XLSX: cópia entre
# discos arredonda o horário.
FOLGA_MTIME = 2.0


def _assinatura(caminho):
    """
    (data de modificação em ns, tamanho) de um arquivo — a chave dos caches de
    leitura deste módulo. Trocar o arquivo muda a chave, e a leitura seguinte
    relê sozinha, sem depender do "Atualizar dados".
    """
    info = os.stat(caminho)
    return info.st_mtime_ns, info.st_size


def _espelho_em_dia(origem, *espelhos):
    """True se todos os `espelhos` existem e saíram desta versão de `origem`."""
    alvo = os.path.getmtime(origem)
    return all(os.path.isfile(e) and abs(os.path.getmtime(e) - alvo) <= FOLGA_MTIME
               for e in espelhos)


def _gravar_espelho(origem, espelho, df, **to_csv):
    """
    Grava `df` em `espelho` com a data de modificação de `origem`.

    Devolve True se gravou. Falhar aqui não derruba a leitura, que já tem os
    dados em mãos: pasta sem permissão de escrita ou disco cheio só custam a
    lentidão de ler o XLSX de novo da próxima vez.
    """
    # Temporário com nome próprio por processo e thread, datado e trocado de uma
    # vez no fim: uma gravação interrompida, ou duas sessões convertendo o mesmo
    # arquivo juntas, não deixam CSV pela metade nem com a data errada para ser
    # lido no lugar do XLSX. A extensão ".parcial" fica fora dos globs de *.csv.
    parcial = f"{espelho}.{os.getpid()}-{threading.get_ident()}.parcial"
    try:
        df.to_csv(parcial, index=False, **to_csv)
        mtime = os.path.getmtime(origem)
        os.utime(parcial, (mtime, mtime))
        os.replace(parcial, espelho)
        return True
    except OSError as e:
        print(f"  [conversão] não deu para gravar {os.path.basename(espelho)} "
              f"({type(e).__name__}: {e}); o XLSX continua sendo lido direto")
        return False
    finally:
        with contextlib.suppress(OSError):
            os.remove(parcial)


def converter_rpa(xlsx, df=None):
    """
    Grava o CSV espelho de um export XLSX da jornada RPA. Devolve True se gravou.

    `df` é a aba "Base" já lida, quando quem chama já a tem; senão ela é lida.
    """
    if df is None:
        df = pd.read_excel(xlsx, sheet_name=RPA_ABA, dtype=str)
    return _gravar_espelho(xlsx, _espelho_rpa(xlsx), df,
                           sep=RPA_SEP, encoding=RPA_ENCODING_CONVERTIDO)


def converter_nba(xlsx, abas=None):
    """
    Grava o par de CSVs espelho de um export XLSX do NBA. Devolve True se gravou.

    `abas` é o dict {aba: DataFrame} já lido, quando quem chama já o tem.
    """
    if abas is None:
        abas = pd.read_excel(xlsx, sheet_name=[NBA_ABA_ANALITICO, NBA_ABA_DEFEITOS],
                             dtype=str)
    analitico, defeitos = _espelhos_nba(xlsx)
    # O analítico primeiro: é o CSV de defeitos que marca o par como pronto
    # (`_ler_jornada_nba` procura por ele e exige o analítico ao lado). Se a
    # segunda gravação falhar, sobra um analítico solto: o leitor o ignora e a
    # próxima leitura do XLSX regrava os dois.
    return (_gravar_espelho(xlsx, analitico, abas[NBA_ABA_ANALITICO],
                            encoding=NBA_ENCODING_CONVERTIDO)
            and _gravar_espelho(xlsx, defeitos, abas[NBA_ABA_DEFEITOS],
                                encoding=NBA_ENCODING_CONVERTIDO))


def _escolher_rpa(pasta):
    """
    Arquivo que vale numa jornada RPA, como (escolhido, ignorados).

    Um arquivo só: o export do RPA cobre a série inteira, então dois deles na
    pasta somariam tudo em dobro. Vale o mais recente — o que chega mais longe
    nos dados (ver "Qual export é o mais recente") —, e trocar de fonte é jogar
    o novo arquivo lá dentro. O leitor e a data de corte usam esta mesma escolha,
    para o corte não olhar arquivo que o relatório não lê.
    """
    candidatos = _arquivos(pasta, "*.csv") + _arquivos(pasta, "*.xlsx")
    if not candidatos:
        raise FileNotFoundError(f"Nenhum CSV ou XLSX em {pasta}")
    # Um XLSX e o próprio espelho empatam no último dia; aí decide a data de
    # modificação (o espelho herda a do XLSX, ver `_gravar_espelho`) e, no
    # empate dela, o CSV: ler o XLSX custa ~35 s contra ~2 s do CSV.
    escolhido = max(candidatos,
                    key=lambda c: (_ultimo_dia_rpa(c) or SEM_DATA,
                                   os.path.getmtime(c), c.lower().endswith(".csv")))
    # A folga cobre o espelho que ficou segundos mais velho que o XLSX numa
    # cópia entre discos e por isso perdeu o desempate acima.
    if escolhido.lower().endswith(".xlsx") and _espelho_em_dia(escolhido, _espelho_rpa(escolhido)):
        escolhido = _espelho_rpa(escolhido)
    return escolhido, [f for f in candidatos if f != escolhido]


# Último dia de cada export RPA, guardado na própria pasta da jornada. Para
# escolher o mais recente é preciso a data de todos os exports da pasta, e não só
# a do escolhido: ~0,7 s por CSV de 35 MB, e de 35 s a um minuto para um XLSX
# que nunca foi convertido. O índice paga isso uma vez por arquivo. A chave é
# nome e tamanho, e não a data de modificação — que é justamente o que não
# sobrevive ao envio da pasta.
ARQ_INDICE_RPA = "_ultimo_dia_exports.json"
_TRAVA_INDICE_RPA = threading.Lock()


def _ultimo_dia_rpa(caminho):
    """Último dia com pedido num export RPA (CSV ou XLSX), ou None se não houver."""
    return _ultimo_dia_rpa_cacheado(caminho, os.path.getsize(caminho))


@functools.lru_cache(maxsize=64)
def _ultimo_dia_rpa_cacheado(caminho, tamanho):
    """Corpo de `_ultimo_dia_rpa`: primeiro o índice da pasta, senão os dados."""
    pasta, nome = os.path.split(caminho)
    indice = os.path.join(pasta, ARQ_INDICE_RPA)
    chave = f"{nome}|{tamanho}"
    with _TRAVA_INDICE_RPA:
        guardado = _ler_indice_rpa(indice)
    if chave in guardado:
        valor = guardado[chave]
        return _dt.date.fromisoformat(valor) if valor else None

    dia = _apurar_ultimo_dia_rpa(caminho)
    with _TRAVA_INDICE_RPA:
        # Relido dentro da trava: outra sessão pode ter gravado no meio tempo.
        guardado = _ler_indice_rpa(indice)
        guardado[chave] = dia.isoformat() if dia else None
        # Sai do índice o que não está mais na pasta, para ele não só crescer.
        presentes = set(os.listdir(pasta))
        guardado = {k: v for k, v in guardado.items() if k.rsplit("|", 1)[0] in presentes}
        _gravar_indice_rpa(indice, guardado)
    return dia


def _apurar_ultimo_dia_rpa(caminho):
    """Lê a coluna de data do export; um XLSX sem espelho já sai convertido."""
    try:
        if caminho.lower().endswith(".xlsx"):
            espelho = _espelho_rpa(caminho)
            if _espelho_em_dia(caminho, espelho):
                return _ultimo_dia_rpa(espelho)
            serie = _ler_bruto_rpa(caminho)[RPA_COL_DATA]
        else:
            serie = pd.read_csv(caminho, sep=RPA_SEP, encoding=_encoding_rpa(caminho),
                                usecols=[RPA_COL_DATA], dtype=str,
                                low_memory=False)[RPA_COL_DATA]
    except (KeyError, ValueError) as e:
        # Planilha sem a aba "Base" ou sem a coluna de data: não é um export
        # que o leitor aceite, e sem data ela fica atrás de todos.
        print(f"  [rpa] {os.path.basename(caminho)}: sem data legível "
              f"({type(e).__name__}: {e})")
        return None
    maximo = _datas_br_ou_iso(serie).max()
    return maximo.date() if pd.notna(maximo) else None


def _ler_indice_rpa(indice):
    try:
        with open(indice, encoding="utf-8") as fh:
            dados = json.load(fh)
        return dados if isinstance(dados, dict) else {}
    except (OSError, ValueError):
        return {}


def _gravar_indice_rpa(indice, dados):
    """Grava o índice de uma vez; sem permissão de escrita, só se perde o atalho."""
    parcial = f"{indice}.{os.getpid()}-{threading.get_ident()}.parcial"
    try:
        with open(parcial, "w", encoding="utf-8") as fh:
            json.dump(dados, fh, ensure_ascii=False, indent=1, sort_keys=True)
        os.replace(parcial, indice)
    except OSError as e:
        print(f"  [rpa] não deu para gravar {ARQ_INDICE_RPA} ({type(e).__name__}: {e})")
    finally:
        with contextlib.suppress(OSError):
            os.remove(parcial)


def _ler_bruto_rpa(escolhido):
    """Planilha inteira do export RPA escolhido; um XLSX já é convertido na leitura."""
    if escolhido.lower().endswith(".xlsx"):
        bruto = pd.read_excel(escolhido, sheet_name=RPA_ABA, dtype=str)
        if converter_rpa(escolhido, bruto):
            print(f"  [conversão] {os.path.basename(escolhido)} -> "
                  f"{os.path.basename(_espelho_rpa(escolhido))}")
        return bruto
    return pd.read_csv(escolhido, sep=RPA_SEP, encoding=_encoding_rpa(escolhido),
                       dtype=str, low_memory=False)


def _ler_jornada_rpa(pasta, corte=None):
    """
    Lê uma jornada em formato RPA e devolve (df_falhas, sucessos_por_mes).

    O CSV traz sucessos e falhas juntos: a coluna "SUCESSO OU FALHA" separa os
    dois. As falhas viram linhas no formato interno; os sucessos são apenas
    contados por mês, que é o que o cálculo do fallout precisa. `corte` descarta
    os dias posteriores ao último dia completo (ver `data_corte`).

    O resultado fica em cache pelo arquivo escolhido (data e tamanho) e pelo
    corte: o PowerPoint e o Excel de ofensores passam por aqui uma vez por
    combinação de jornadas, e cada passada relia a planilha inteira.
    """
    escolhido, ignorados = _escolher_rpa(pasta)
    for f in ignorados:
        print(f"  [rpa] ignorado (mais antigo): {os.path.basename(f)}")
    df, sucessos = _jornada_rpa_cacheada(escolhido, *_assinatura(escolhido), corte)
    return df.copy(), sucessos.copy()


@functools.lru_cache(maxsize=4)
def _jornada_rpa_cacheada(escolhido, _mtime, _tamanho, corte):
    """Corpo de `_ler_jornada_rpa`. Data e tamanho do arquivo só compõem a chave."""
    bruto = _ler_bruto_rpa(escolhido)
    bruto = bruto.loc[:, ~bruto.columns.str.startswith("Unnamed:")]

    # O formato da data muda com a exportação: o "Radar Ofensores RPA" traz
    # dd/mm/aaaa e o "Solar - Apoio PME Móvel" traz ISO com hora. Fixar um
    # formato só zerava a coluna inteira do outro, e aí o `where` abaixo
    # comparava uma coluna só de NaT com uma date — que estoura com "Invalid
    # comparison between dtype=datetime64[ns] and date".
    data = _datas_br_ou_iso(bruto[RPA_COL_DATA])
    if not data.notna().any():
        amostra = bruto[RPA_COL_DATA].dropna().astype(str).head(3).tolist()
        raise ValueError(
            f"{os.path.basename(escolhido)}: nenhuma data reconhecida na coluna "
            f"{RPA_COL_DATA!r}. Amostra do que veio: {amostra}"
        )
    if corte is not None:
        # `where` em vez de filtrar: o índice tem que continuar alinhado com
        # `bruto`, que é indexado logo abaixo por esta mesma máscara.
        # A comparação é datetime64 contra Timestamp, e não `.dt.date` contra
        # date: aquela passava por um array de objetos e quebrava justamente
        # quando nenhuma data era reconhecida, escondendo a causa real.
        data = data.where(data.dt.normalize() <= pd.Timestamp(corte))
    bruto = bruto[data.notna()].copy()
    bruto["CreatedDate"] = (
        data[data.notna()].dt.tz_localize("America/Sao_Paulo", ambiguous=True,
                                          nonexistent="shift_forward").dt.tz_convert("UTC")
    )

    resultado = bruto[RPA_COL_RESULTADO].fillna("").str.strip().str.lower()
    e_falha = resultado == RPA_VALOR_FALHA.lower()

    sucessos = (
        bruto[~e_falha]
        .assign(Mes=lambda d: chave_de_datas(
            d["CreatedDate"].dt.tz_convert("America/Sao_Paulo")))
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
        # Quando o chamado foi aberto: só esta jornada traz, e o slide dela
        # mostra numa coluna própria (ver gerar_pptx._montar_dados).
        "DFT_CreatedDate": pd.to_datetime(
            cat[NBA_COL_CRIACAO] if NBA_COL_CRIACAO in cat.columns
            else pd.Series(pd.NA, index=cat.index),
            format="%d/%m/%Y", errors="coerce"),
    }).drop_duplicates(subset="DefectNumber_orig")


def _ler_jornada_nba(pasta, corte=None):
    """
    Lê uma jornada em formato NBA e devolve (df_falhas, sucessos_por_mes).

    A aba "Analitico" traz um pedido por linha, com sucessos e falhas juntos; a
    aba "Defeitos" entra no lugar do Octane, já que os defeitos são BRJ-* e não
    existem lá. `corte` descarta os dias posteriores ao último dia completo
    (ver `data_corte`) — só dos pedidos, nunca do catálogo de defeitos.

    O resultado fica em cache pelos arquivos escolhidos (data e tamanho) e pelo
    corte, como em `_ler_jornada_rpa`.
    """
    escolhas, ignorados = _escolher_nba(pasta)
    for ignorado in ignorados:
        print(f"  [nba] ignorado (mais antigo, mesmo mês): {os.path.basename(ignorado)}")
    if not escolhas:
        raise FileNotFoundError(f"Nenhuma planilha do NBA encontrada em {pasta}")
    chave = tuple((tipo, caminho) + _assinatura_nba(tipo, caminho) for tipo, caminho in escolhas)
    df, sucessos = _jornada_nba_cacheada(chave, corte)
    return df.copy(), sucessos.copy()


def _escolher_nba(pasta):
    """
    Exports que valem numa jornada NBA, como ([(tipo, caminho), ...], ignorados).

    `tipo` é "csv" (o caminho é a base do par de CSVs) ou "xlsx". Cada export
    cobre um mês. Dois arquivos do mesmo mês na pasta — o de uma extração
    anterior e o da nova — somariam o mês em dobro, então vale só o mais recente.
    É o que faz "jogar o arquivo novo na pasta" bastar, sem precisar apagar o
    antigo. O leitor e a data de corte usam esta mesma escolha.

    "Mais recente" é pela data da extração no nome ("Completo_16_09_2026"), e
    não pela data de modificação (ver "Qual export é o mais recente"). Ler a
    data dos dados custaria abrir os XLSX que perdem, e que hoje nunca são lidos.
    """
    candidatos = {}          # mês -> (data no nome, mtime, tipo, caminho base)
    for arq_def in _arquivos(pasta, "*" + NBA_SUF_DEFEITOS):
        base = arq_def[: -len(NBA_SUF_DEFEITOS)]
        if not os.path.isfile(base + NBA_SUF_ANALITICO):
            raise FileNotFoundError(
                f"Falta o par de {os.path.basename(arq_def)}: {base + NBA_SUF_ANALITICO}")
        candidatos.setdefault(mes_do_arquivo(arq_def), []).append(
            (data_do_nome(base) or SEM_DATA, os.path.getmtime(arq_def), "csv", base))
    for arq in _arquivos(pasta, "*.xlsx"):
        candidatos.setdefault(mes_do_arquivo(arq), []).append(
            (data_do_nome(arq) or SEM_DATA, os.path.getmtime(arq), "xlsx", arq))

    escolhas, ignorados = [], []
    for mes, opcoes in sorted(candidatos.items(), key=lambda kv: (kv[0] is None, kv[0])):
        # Um XLSX e o próprio espelho têm o mesmo nome, então empatam na data
        # dele; aí decide a de modificação e, no empate dela, o CSV: o espelho
        # herda a data do XLSX (ver `_gravar_espelho`), e sem o desempate
        # explícito "xlsx" > "csv" na ordenação da tupla e o XLSX lento ganharia.
        opcoes.sort(key=lambda o: (o[0], o[1], o[2] == "csv"), reverse=True)
        _, _, tipo, escolhido = opcoes[0]
        # A folga cobre o espelho que ficou segundos mais velho que o XLSX numa
        # cópia entre discos e por isso perdeu o desempate acima.
        if tipo == "xlsx" and _espelho_em_dia(escolhido, *_espelhos_nba(escolhido)):
            tipo, escolhido = "csv", os.path.splitext(escolhido)[0]
        ignorados += [ign for _, _, _t, ign in opcoes if ign != escolhido]
        escolhas.append((tipo, escolhido))
    return escolhas, ignorados


def _assinatura_nba(tipo, caminho):
    """Data e tamanho dos arquivos de um export NBA, para chave de cache."""
    if tipo == "csv":
        return (_assinatura(caminho + NBA_SUF_ANALITICO)
                + _assinatura(caminho + NBA_SUF_DEFEITOS))
    return _assinatura(caminho)


def _ler_abas_nba(tipo, caminho):
    """(analítico, defeitos) de um export NBA; um XLSX já é convertido na leitura."""
    if tipo == "csv":
        return (pd.read_csv(caminho + NBA_SUF_ANALITICO, dtype=str),
                pd.read_csv(caminho + NBA_SUF_DEFEITOS, dtype=str))
    # As duas abas numa chamada só, para abrir o arquivo uma vez.
    abas = pd.read_excel(caminho, sheet_name=[NBA_ABA_ANALITICO, NBA_ABA_DEFEITOS], dtype=str)
    if converter_nba(caminho, abas):
        print(f"  [conversão] {os.path.basename(caminho)} -> "
              f"par {NBA_SUF_ANALITICO} / {NBA_SUF_DEFEITOS}")
    return abas[NBA_ABA_ANALITICO], abas[NBA_ABA_DEFEITOS]


@functools.lru_cache(maxsize=4)
def _jornada_nba_cacheada(escolhas, corte):
    """Corpo de `_ler_jornada_nba`. Data e tamanho dos arquivos só compõem a chave."""
    partes, catalogos = [], []

    def _registrar(analitico, defeitos):
        partes.append(analitico)
        # Guarda até quando o arquivo vai, para saber qual catálogo é o mais novo.
        ate = pd.to_datetime(analitico[NBA_COL_DATA], format="%d/%m/%Y", errors="coerce").max()
        catalogos.append((ate, _catalogo_nba(defeitos)))

    for tipo, caminho, *_assinatura_arquivos in escolhas:
        _registrar(*_ler_abas_nba(tipo, caminho))
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
    if corte is not None:
        # `where` em vez de filtrar: o índice tem que continuar alinhado com
        # `bruto`, que é indexado logo abaixo por esta mesma máscara.
        data = data.where(data.dt.date <= corte)
    bruto = bruto[data.notna()].copy()
    bruto["CreatedDate"] = (
        data[data.notna()].dt.tz_localize("America/Sao_Paulo", ambiguous=True,
                                          nonexistent="shift_forward").dt.tz_convert("UTC")
    )

    e_sucesso = bruto[NBA_COL_STATUS].fillna("").str.strip().str.upper() == NBA_VALOR_SUCESSO
    sucessos = (
        bruto[e_sucesso]
        .assign(Mes=lambda d: chave_de_datas(
            d["CreatedDate"].dt.tz_convert("America/Sao_Paulo")))
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
    # Outra exclusiva: o erro que a origem agrupa como "ErroProcesso - …" é o
    # mesmo que o legado dá (ver CAT_ERRO_LEGADO).
    _erro = (df["ErrorHandled__c"].fillna("").astype(str)
             .str.replace(r"\s+", "", regex=True).str.lower())
    df["ErroLegado"] = _erro.str.startswith(NBA_PREFIXO_ERRO_LEGADO)
    # E as regras de CAT_SCRIPT_RPA, CAT_AGENDAMENTO e da Falha Pontual sem
    # defeito, que valem só nesta jornada.
    df["ErroAgendamento"] = _erro.str.startswith(NBA_PREFIXOS_AGENDAMENTO)
    df["BaseResidencial"] = True

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


def _octane(raiz, base_dir, hoje=None):
    """
    Catálogo do Octane (DFTs + US) no formato das colunas DFT_*.

    Cada combinação de jornadas passa por `carregar_base` e precisa do catálogo,
    que antes era relido a cada vez: ao abrir o app eram 8 leituras de cada
    planilha, ~2,5 s por par. Agora o par é lido uma vez por processo e volta a
    ser lido sozinho quando uma das planilhas muda — a chave do cache é a data de
    modificação e o tamanho de cada uma.

    Devolve uma cópia, para quem recebe o catálogo não alterar o que está no
    cache. O dia (`hoje`, ou o de hoje) também entra na chave: a Milestone de
    uma US aberta só vale como data enquanto não passou (ver `_entrega_us`).
    """
    caminhos = (_caminho_octane(raiz, base_dir, ARQ_DFT),
                _caminho_octane(raiz, base_dir, ARQ_US))
    chave = []
    for c in caminhos:
        chave += [c, *_assinatura(c)]
    # O dia de `hoje_utc`, o mesmo das outras regras, sem fuso: a Milestone vem
    # sem fuso também.
    dia = hoje_utc(hoje).tz_localize(None)
    return _octane_cacheado(*chave, dia).copy()


TIPO_US = "User Story"
# Como cada export do Octane escreve o tipo. O de DFT de 08/10 passou a trazer
# histórias e defeitos juntos, em minúsculas ("story", "defect"), e o app só
# reconhecia "User Story": as US 233426, 234139 e 225194 caíam em "DFT s/ data".
_TIPOS_OCTANE = {"story": TIPO_US, "user story": TIPO_US, "defect": "Defect"}


def _tipo_octane(serie):
    """A coluna Type no padrão do app: "User Story" e "Defect"."""
    texto = serie.astype("string").str.strip()
    return texto.str.lower().map(_TIPOS_OCTANE).fillna(texto).astype(object)


def _entrega_us(df_us, hoje):
    """
    Bugfix Milestone das US, preenchida pela coluna Milestone quando vazia.

    A US não usa a Bugfix Milestone — nenhuma das 616 em Done ou Implantado
    Prod a trazia em 25/09/2026 —, e sem data o pedido não tem como passar por
    Tratado e Resolvido. Na entregue (FASES_US_ENTREGUE) vale a data de
    implantação da coluna Milestone ("Implantação_2026.09.12"), que é a de
    produção; sem ela, a Done date, que é quando a US foi fechada no Octane e
    costuma vir dias depois.

    Desde 11/11/2026 a Milestone vale também na US em andamento (fora de
    FASES_US_ENCERRADAS e de FASES_MOPS), mas só a partir de `hoje`: uma
    implantação que já passou com a US ainda aberta não é previsão de nada, e
    ela fica sem data. Era o caso da 234139 (Implantação_2026.12.05), que
    aparecia em "US s/ data" com a entrega marcada.
    """
    ms = pd.to_datetime(df_us["Bugfix Milestone"], errors="coerce")
    fase = df_us["Phase"].fillna("").astype(str).str.strip()
    entregue = fase.isin(FASES_US_ENTREGUE)
    aberta = ~fase.isin(FASES_US_ENCERRADAS | FASES_MOPS)
    vazio = pd.Series(pd.NA, index=df_us.index, dtype=object)
    partes = (df_us.get("Milestone", vazio).astype(str)
              .str.extract(r"(\d{4})\.(\d{2})\.(\d{2})"))
    implantacao = pd.to_datetime(partes[0] + "-" + partes[1] + "-" + partes[2],
                                 errors="coerce")
    feito = pd.to_datetime(df_us.get("Done date", vazio), errors="coerce")
    ms = ms.where(ms.notna() | ~entregue, implantacao.fillna(feito))
    return ms.where(ms.notna() | ~aberta | ~(implantacao >= hoje), implantacao)


@functools.lru_cache(maxsize=4)
def _octane_cacheado(caminho_dft, _mtime_dft, _tam_dft, caminho_us, _mtime_us, _tam_us,
                     hoje):
    """
    Lê e monta o catálogo do Octane. A data e o tamanho só compõem a chave;
    `hoje` (Timestamp do dia) decide quais Milestones de US abertas valem.
    """
    df_dft = pd.read_excel(caminho_dft)
    df_us = pd.read_excel(caminho_us)
    for tabela in (df_dft, df_us):
        tabela["Type"] = _tipo_octane(tabela["Type"])
    # O export de DFT de 08/10 traz também as histórias ("story"), e mais novas
    # que as do relatório de US: das 1.890 que estão nos dois, as fases dele
    # andaram (Implantado Prod → Done, Backlog → Aguardando Testes Integrados) e
    # o relatório de US, de 29/09, não. A história vale então pelo export de
    # DFT; o de US fica para as que só existem nele. `df_us` passa a ser essa
    # junção, e é dela que saem a troca pela US de Melhoria e o catálogo.
    historias = df_dft["Type"] == TIPO_US
    df_us = (pd.concat([df_dft[historias], df_us], ignore_index=True)
               .drop_duplicates(subset="ID").reset_index(drop=True))
    df_dft = df_dft[~historias]
    # Antes da troca pela US de Melhoria, logo abaixo: o DFT que aponta para
    # uma US entregue herda a data dela.
    df_us["Bugfix Milestone"] = _entrega_us(df_us, hoje)

    df_dft_prep = df_dft[
        COLUNAS_OCTANE_BASE + ([COL_US_MELHORIA] if COL_US_MELHORIA in df_dft.columns else [])
    ].copy()
    # Número da US que tomou o lugar do DFT, só quando ela foi achada no
    # relatório de US — o mesmo caso em que nome, fase e time viram os dela.
    # O "Top Ofensores" mostra esse número no lugar do DFT.
    df_dft_prep[COL_US_RESOLVIDA] = pd.NA

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
                for col in ["Name", "Phase", "Team"]:
                    if pd.notna(us[col]):
                        df_dft_prep.at[i, col] = us[col]
                # A data de entrega é sempre a da US, mesmo vazia: a do DFT é de
                # um planejamento que a US substituiu. Só copiada quando
                # preenchida, a do DFT ficava — o 214711, cancelado, passava a
                # data de 23/04/2025 para a US 216759, que está em Backlog.
                df_dft_prep.at[i, "Bugfix Milestone"] = us["Bugfix Milestone"]
                df_dft_prep.at[i, "Type"] = "User Story"
                df_dft_prep.at[i, COL_US_RESOLVIDA] = us_id

    df_octane = (
        pd.concat([df_dft_prep[COLUNAS_OCTANE_BASE], df_us[COLUNAS_OCTANE_BASE]], ignore_index=True)
        .drop_duplicates(subset="ID")
        .rename(columns={
            "ID": "DefectNumber__c", "Name": "DFT_Name", "Phase": "DFT_Phase",
            "Bugfix Milestone": "DFT_BugfixMilestone", "Team": "DFT_Team", "Type": "DFT_Type",
        })
    )
    # O ID como texto, para os itens que não têm número: o export de 08/10 traz
    # 627 "DDP-796" e afins. Viravam -1 no número e casavam com todo pedido sem
    # DFT — 8.440 da PME saíam com o nome "DDP - Bloqueio de contestação técnica
    # de duas pontas…". `carregar_base` casa pelo número só os de número, e
    # estes pelo rótulo (ver `COL_CHAVE_OCTANE`).
    df_octane[COL_CHAVE_OCTANE] = df_octane["DefectNumber__c"].astype(str).str.strip().str.upper()
    df_octane["DefectNumber__c"] = (
        pd.to_numeric(df_octane["DefectNumber__c"], errors="coerce")
        .fillna(-1).astype(int)
    )
    _resolvidas = df_dft_prep.dropna(subset=[COL_US_RESOLVIDA])
    df_octane["DFT_US"] = df_octane["DefectNumber__c"].map(dict(zip(
        pd.to_numeric(_resolvidas["ID"], errors="coerce"), _resolvidas[COL_US_RESOLVIDA])))
    return df_octane


def _catalogo_jira(raiz, base_dir):
    """
    Chamados do Jira no formato interno, indexados pelo número do chamado.

    O status vira `DFT_Phase` e a data de fechamento vira `DFT_BugfixMilestone`.
    Isso é de propósito, e não atalho: com essas duas colunas preenchidas,
    `categorizar` já produz a regra pedida sem caso especial nenhum —
    "Em Atendimento" fica sem data de previsão e cai em Em Tratamento pela
    Squad; "Resolvido" cai em avaliação de eficácia enquanto o fechamento
    estiver dentro dos 15 dias, e em Resolvido daí em diante; "Cancelado" cai
    em Em Avaliação por MOPs, como o DFT cancelado no Octane.

    Devolve None quando a planilha não existe: ela é opcional, e quem não
    depende dela (o Base Residencial tem catálogo próprio) segue igual.
    """
    try:
        caminho = _caminho_octane(raiz, base_dir, ARQ_JIRA)
    except FileNotFoundError:
        return None
    # Em cache por data e tamanho, como o Octane: era relida a cada combinação.
    return _jira_cacheado(caminho, *_assinatura(caminho)).copy()


def _catalogo_brj(raiz, base_dir):
    """
    Chamados BRJ no formato das colunas DFT_*, indexados pela chave em
    maiúsculas ("BRJ-640"), mais a coluna `_resolvido`. None sem a planilha.
    """
    try:
        caminho = _caminho_octane(raiz, base_dir, ARQ_BRJ)
    except FileNotFoundError:
        return None
    return _brj_cacheado(caminho, *_assinatura(caminho)).copy()


@functools.lru_cache(maxsize=4)
def _brj_cacheado(caminho, _mtime, _tamanho):
    """Corpo de `_catalogo_brj`. Data e tamanho só compõem a chave."""
    bruto = pd.read_excel(caminho, sheet_name=BRJ_ABA)
    col_data = next((c for c in bruto.columns
                     if str(c).strip().startswith(BRJ_PREFIXO_DATA)), None)
    if col_data is None:
        raise ValueError(f"{ARQ_BRJ}: a aba {BRJ_ABA!r} não tem a coluna "
                         f"{BRJ_PREFIXO_DATA!r} com a data de implantação.")
    status = bruto[BRJ_COL_STATUS].fillna("").astype(str).str.strip()
    tipo = bruto[BRJ_COL_TIPO].fillna("").astype(str).str.strip()
    texto = lambda col: (bruto[col].fillna("").astype(str)
                         .str.replace(r"\s+", " ", regex=True).str.strip()
                         .replace("", pd.NA))
    cat = pd.DataFrame({
        "DFT_Name": texto(BRJ_COL_RESUMO),
        "DFT_Phase": status.replace("", pd.NA),
        "DFT_BugfixMilestone": pd.to_datetime(bruto[col_data], errors="coerce"),
        # Data do fechamento, para o concluído sem data de implantação — o
        # mesmo papel da Done date nas US (ver `_entrega_us`).
        "_resolvido": pd.to_datetime(bruto[BRJ_COL_RESOLVIDO], errors="coerce")
                        .where(status.isin(FASES_BRJ_ENTREGUE)),
        "DFT_Team": texto(BRJ_COL_RESPONSAVEL).str.title(),
        "DFT_Type": tipo.map(lambda t: "User Story" if t in BRJ_TIPOS_US else "Defect"),
    })
    cat.index = bruto[BRJ_COL_CHAVE].fillna("").astype(str).str.strip().str.upper()
    cat = cat[cat.index.str.startswith("BRJ-")]
    return cat[~cat.index.duplicated()]


def _cabecalho_jira(caminho):
    """Em qual aba e linha está o cabeçalho do relatório do Jira.

    Devolve `(aba, linha)` para passar direto ao `read_excel`. Procurar pelas
    três colunas conhecidas, em vez de fixar aba e linha, aguenta o export
    trocar de nome de aba ou perder a coluna vazia da esquerda — as duas coisas
    já aconteceram.
    """
    with pd.ExcelFile(caminho) as arquivo:
        for aba in arquivo.sheet_names:
            topo = arquivo.parse(aba, header=None, nrows=JIRA_MAX_LINHAS_CABECALHO)
            for linha in range(len(topo)):
                celulas = {str(v).strip() for v in topo.iloc[linha] if pd.notna(v)}
                if set(JIRA_COLS) <= celulas:
                    return aba, linha
        abas = arquivo.sheet_names
    raise ValueError(
        f"{ARQ_JIRA}: não achei as colunas {list(JIRA_COLS)} nas primeiras "
        f"{JIRA_MAX_LINHAS_CABECALHO} linhas de nenhuma aba. Abas: {abas}")


@functools.lru_cache(maxsize=4)
def _jira_cacheado(caminho, _mtime, _tamanho):
    """Corpo de `_catalogo_jira`. Data e tamanho só compõem a chave."""
    aba, linha = _cabecalho_jira(caminho)
    bruto = pd.read_excel(caminho, sheet_name=aba, header=linha)

    cat = pd.DataFrame({
        "DefectNumber_orig": _limpar_defeito(bruto[JIRA_COL_CHAMADO]),
        "DFT_Phase": _limpar_defeito(bruto[JIRA_COL_STATUS]),
        "DFT_BugfixMilestone": pd.to_datetime(bruto[JIRA_COL_FECHAMENTO], errors="coerce"),
    })
    cat = cat[cat["DefectNumber_orig"].str.upper().str.startswith(JIRA_PREFIXOS, na=False)]

    previstos = {JIRA_RESOLVIDO, JIRA_EM_ATENDIMENTO, JIRA_CANCELADO}
    outros = set(cat["DFT_Phase"].unique()) - previstos
    if outros:
        print(f"  [jira] status fora do previsto, ignorados: {sorted(outros)}")
        cat = cat[cat["DFT_Phase"].isin(previstos)]

    # "Em Atendimento" não tem previsão de entrega, e o cancelado não vai ter
    # entrega nenhuma; se vier data, ela não vale.
    cat.loc[cat["DFT_Phase"].isin([JIRA_EM_ATENDIMENTO, JIRA_CANCELADO]),
            "DFT_BugfixMilestone"] = pd.NaT
    # O chamado se repete na planilha, às vezes uma vez com data e outra sem
    # (PDST-1990710, PDST-2013844). Vence a linha com data e, entre duas com
    # data, o fechamento mais recente — senão um "Resolvido" perderia a data e
    # voltaria a contar como em tratamento.
    cat = (cat.sort_values("DFT_BugfixMilestone", na_position="first")
              .drop_duplicates(subset="DefectNumber_orig", keep="last"))
    # A origem não traz tipo: são chamados, então entram como defeito, e não
    # como User Story — é o que separa "DFT s/ data" de "US s/ data" no app.
    cat["DFT_Type"] = "Defect"
    return cat.set_index("DefectNumber_orig")


def _norm_id(v):
    """Normaliza um ID que pode vir como float ('182366.0') para string inteira."""
    try:
        return str(int(float(str(v).strip())))
    except (TypeError, ValueError):
        return str(v).strip()


# O portal exporta as falhas como ".xls" que é uma tabela HTML em ISO-8859-1, e
# os sucessos como .xlsx de verdade. Os CSVs antigos continuam sendo lidos.
EXT_FALHAS = (".csv", ".xls", ".xlsx", ".htm", ".html")
EXT_SUCESSOS = (".csv", ".xlsx", ".xls")
RE_TAG = re.compile(r"<[^>]+>")


def _tabela_html(caminho):
    """Tabela HTML exportada com extensão .xls -> DataFrame de texto."""
    html = io.open(caminho, encoding="ISO-8859-1").read()
    linhas = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S)
    if not linhas:
        raise ValueError(f"{caminho}: nenhuma linha <tr> encontrada")

    def _celulas(bruto, tag):
        return [RE_TAG.sub("", c).replace("&amp;", "&").replace("&lt;", "<")
                .replace("&gt;", ">").replace("&nbsp;", " ").strip()
                for c in re.findall(rf"<{tag}[^>]*>(.*?)</{tag}>", bruto, re.S)]

    cabecalho = _celulas(linhas[0], "th")
    dados = []
    for bruto in linhas[1:]:
        cel = _celulas(bruto, "td")
        if cel:
            dados.append((cel + [""] * len(cabecalho))[:len(cabecalho)])
    return pd.DataFrame(dados, columns=cabecalho, dtype=str)


def _datas_br_ou_iso(serie):
    """
    Converte uma coluna de data aceitando os dois formatos que aparecem.

    A mesma coluna chega ora em dd/mm/aaaa ("25/08/2026"), ora em ISO
    ("2026-08-13", ou "2026-07-01 00:00:00" quando a origem é uma célula de
    data do Excel lida como texto). Sem separar os dois, `pd.to_datetime` lê
    "01/08/2026" como 8 de janeiro — todo dia até o 12 cairia no mês errado,
    sem erro nenhum aparecendo.

    Vale para as falhas das jornadas de extração e para o relatório de RPA,
    que trocou de formato ao mudar de exportação.
    """
    texto = serie.astype(str).str.strip()
    br = texto.str.match(r"^\d{2}/\d{2}/\d{4}$", na=False)
    datas = pd.Series(pd.NaT, index=texto.index, dtype="datetime64[ns]")
    if br.any():
        datas[br] = pd.to_datetime(texto[br], format="%d/%m/%Y", errors="coerce")
    if (~br).any():
        datas[~br] = pd.to_datetime(texto[~br], errors="coerce")
    return datas


def _arquivos_da_etapa(pasta, subpasta, extensoes):
    """Arquivos de falhas/ ou sucessos/ em qualquer formato aceito."""
    achados = []
    for ext in extensoes:
        achados += _arquivos(os.path.join(pasta, subpasta), "*" + ext)
    return sorted(set(achados))


def _numero_pedido(serie):
    """
    OrderNumber sempre como texto, com o mesmo formato em todos os leitores.

    O CSV é lido pelo pandas e vira int64; a tabela HTML vem como string. Sem
    normalizar, "52118713" e 52118713 são valores diferentes, o
    `drop_duplicates(OrderNumber)` não casa os dois, e o mesmo pedido presente
    no arquivo antigo e no novo é contado duas vezes.
    """
    numero = pd.to_numeric(serie, errors="coerce")
    texto = serie.astype(str).str.strip()
    return numero.astype("Int64").astype(str).where(numero.notna(), texto)


def _ler_falhas_extracao(caminho):
    """
    Lê um CSV da pasta falhas/ e devolve as colunas já no formato interno,
    aceitando tanto o export com nome de API quanto o com rótulo em português.

    Em cache por arquivo (data e tamanho): a data de corte, cada combinação de
    jornadas e cada relatório passavam pelo mesmo arquivo — num PowerPoint, o
    ago26.xls da Base era processado 9 vezes. Devolve uma cópia, para quem
    recebe não alterar o que está no cache.
    """
    return _falhas_extracao_cacheada(caminho, *_assinatura(caminho)).copy()


def _ultimo_dia_falhas(caminho):
    """Dia do pedido mais recente de um arquivo de falhas, ou None se não houver."""
    datas = _falhas_extracao_cacheada(caminho, *_assinatura(caminho))["CreatedDate"]
    maximo = datas.dt.tz_convert("America/Sao_Paulo").max()
    return maximo.date() if pd.notna(maximo) else None


@functools.lru_cache(maxsize=128)
def _falhas_extracao_cacheada(caminho, _mtime, _tamanho):
    """Corpo de `_ler_falhas_extracao`. Data e tamanho só compõem a chave."""
    if caminho.lower().endswith((".xls", ".htm", ".html")):
        df = _tabela_html(caminho)
    elif caminho.lower().endswith(".xlsx"):
        df = pd.read_excel(caminho, dtype=str)
    else:
        df = pd.read_csv(caminho)
    if COL_DATA_FALHAS_PT in df.columns:
        df = df.rename(columns=RENAME_FALHAS_PT)
        df["OrderNumber"] = _numero_pedido(df["OrderNumber"])
        # O layout em português traz só a data, já no fuso de Brasília, enquanto
        # o antigo traz o timestamp com offset. Localizar antes de normalizar
        # para UTC evita que os pedidos do dia 1º caiam no mês anterior quando
        # o resto do código converte de volta para America/Sao_Paulo.
        df["CreatedDate"] = (
            _datas_br_ou_iso(df["CreatedDate"])
            .dt.tz_localize("America/Sao_Paulo")
            .dt.tz_convert("UTC")
        )
    else:
        df = df.rename(columns=RENAME_FALHAS)
        # Um export de outra consulta (em 25/09/2026 veio um só com Id, defeito,
        # ofensor, erro e data) dava "KeyError: 'OrderNumber'" sem dizer de onde.
        faltando = [c for c in ("OrderNumber", "CreatedDate") if c not in df.columns]
        if faltando:
            raise ValueError(
                f"{os.path.basename(os.path.dirname(os.path.dirname(caminho)))}/falhas/"
                f"{os.path.basename(caminho)}: o arquivo não tem as colunas "
                f"{', '.join(faltando)} — não parece o export de falhas. Colunas "
                f"que vieram: {', '.join(map(str, df.columns))}.")
        df["CreatedDate"] = pd.to_datetime(df["CreatedDate"], utc=True, errors="coerce")
        df["OrderNumber"] = _numero_pedido(df["OrderNumber"])
    return df


def _ler_falhas_jornada(pasta):
    """
    Falhas de uma jornada de extração, com um arquivo mandando em cada mês.

    A pasta acumula exports: o CSV mensal antigo fica ao lado do export novo,
    que cobre mês atual e anterior. Concatenar tudo não erra só por contar o
    mesmo pedido duas vezes — isso o `drop_duplicates(OrderNumber)` resolveria.
    O problema é o pedido que a extração nova **tirou**: ele existe só no
    arquivo velho, sobrevive ao concat e continua no relatório. Em jul/26 o
    Prospect ficava com 94 pedidos a mais, 27 deles já reclassificados na Base
    pela extração nova.

    Vence, mês a mês, o arquivo mais recente que cobre aquele mês — inclusive
    nos meses antigos, onde a extração nova também manda. Os meses realmente
    fechados não dependem disso: o histórico manual os substitui depois (ver
    `historico_manual`).

    A escolha é por jornada, e não sobre a pasta inteira: os meses de uma não
    podem ser calados pelo export mais novo de outra.

    "Mais recente" é o que chega mais longe nos dados: cada export cobre o mês
    atual e o anterior até o dia em que foi tirado, então o último pedido dele
    marca a data da extração (ver "Qual export é o mais recente"). Não custa
    leitura a mais: todo arquivo da pasta já é lido aqui e na data de corte.
    """
    arquivos = _arquivos_da_etapa(pasta, "falhas", EXT_FALHAS)
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum CSV encontrado em {os.path.join(pasta, 'falhas')}")

    dono, partes = {}, []
    for caminho in sorted(arquivos, reverse=True,
                          key=lambda c: (_ultimo_dia_falhas(c) or SEM_DATA,
                                         os.path.getmtime(c))):
        df = _ler_falhas_extracao(caminho)
        # 0 é o mês das linhas sem data legível: elas seguem a mesma regra, em
        # vez de sumirem por não casar com mês nenhum.
        mes = chave_de_datas(
            df["CreatedDate"].dt.tz_convert("America/Sao_Paulo")
        ).fillna(0).astype(int)
        meus = set(mes.unique())
        novos = meus - dono.keys()
        for m in sorted(meus - novos):
            print(f"  [falhas] {os.path.basename(caminho)}: "
                  f"{'sem data' if m == 0 else m} ignorado, já veio de "
                  f"{os.path.basename(dono[m])} (mais recente)")
        for m in novos:
            dono[m] = caminho
        partes.append(df[mes.isin(novos)])
    return pd.concat(partes, ignore_index=True)


def _linhas_sucessos(caminho):
    """
    Linhas cruas do relatório de sucessos, do CSV ou do Excel.

    O relatório em Excel tem o mesmo desenho do CSV: cabeçalho, depois a quebra
    diária. Vira lista de listas para o resto seguir igual nos dois casos.

    A chave do cache inclui data e tamanho do arquivo: só pelo caminho, trocar
    o arquivo com o app aberto deixava os sucessos antigos no PowerPoint.
    """
    return _linhas_sucessos_cacheadas(caminho, *_assinatura(caminho))


@functools.lru_cache(maxsize=64)
def _linhas_sucessos_cacheadas(caminho, _mtime, _tamanho):
    """Corpo de `_linhas_sucessos`; devolve tuplas, que não precisam de cópia."""
    if caminho.lower().endswith((".xlsx", ".xls")):
        tabela = pd.read_excel(caminho, header=None, dtype=str)
        return tuple(tuple("" if pd.isna(v) else str(v) for v in linha)
                     for _, linha in tabela.iterrows())
    with open(caminho, encoding="utf-8-sig", newline="") as fh:
        return tuple(tuple(l) for l in csv.reader(fh))


def _dias_cobertos(caminho, mes, corte=None):
    """
    Quantos dias daquele mês o arquivo traz, já descontado o corte.

    Zero no layout agregado, que só tem o total do mês. É esta contagem que
    distingue "cobre o mês inteiro" de "a extração parou no dia 27" — coisa
    que o nome do arquivo não conta.
    """
    return len({dia for dia, _total in _linhas_diarias(caminho)
                if (corte is None or dia <= corte)
                and chave_mes(dia.year, dia.month) == mes})


def _linhas_diarias(caminho):
    """(dia, total em texto) de cada linha da quebra diária; nada no agregado."""
    linhas = _linhas_sucessos(caminho)
    cabecalho = [str(c).strip() for c in linhas[0]] if linhas else []
    if SUC_COL_AGREGADA in cabecalho and SUC_COL_DIA in cabecalho:
        i_total, i_dia = cabecalho.index(SUC_COL_AGREGADA), cabecalho.index(SUC_COL_DIA)
        for linha in linhas[1:]:
            if len(linha) <= max(i_total, i_dia):
                continue
            achado = SUC_DATA_ISO.match(str(linha[i_dia]).strip())
            if achado:
                a, m, d = (int(g) for g in achado.groups())
                yield _dt.date(a, m, d), str(linha[i_total]).strip()
        return
    for linha in linhas:
        celulas = [str(c).strip() for c in linha if str(c).strip()]
        if len(celulas) != 2:
            continue
        achado = SUC_DATA_DIA.match(celulas[0])
        if achado:
            d, m, a = (int(g) for g in achado.groups())
            yield _dt.date(a, m, d), celulas[1]


def _data_export_sucessos(caminho):
    """
    Até quando vai um relatório de sucessos: o último dia da quebra diária.

    O agregado não diz até quando vai, e aí vale o dia da extração no nome
    ("order-18_09_2026.csv"). Sem nenhum dos dois (o "jul26.csv" antigo), None:
    ele fica atrás dos exports com data, o que não o impede de mandar no mês do
    próprio nome (ver a regra em `carregar_base`).
    """
    return (max((dia for dia, _total in _linhas_diarias(caminho)), default=None)
            or data_do_nome(caminho))


def sucessos_por_mes(caminho, corte=None):
    """
    Total de sucessos por mês (chave AAAAMM) contido em um arquivo.

    Um arquivo cobre mês atual e anterior, e antes cada um era atribuído a um
    mês só — o do nome. Nas jornadas antigas isso passava, porque o mês anterior
    tinha arquivo próprio; na PME Fixa, que nasceu com um arquivo só, julho
    ficava com 106 sucessos descartados e 2 falhas contra zero: 100%.

    `corte` descarta os dias posteriores ao último dia completo.
    """
    linhas = _linhas_sucessos(caminho)
    diarias = list(_linhas_diarias(caminho))

    # O agregado quebrado por dia também tem "expr0" no cabeçalho, mas é
    # diário: só o agregado sem dias cai aqui.
    if not diarias and any(SUC_COL_AGREGADA in str(celula)
                           for celula in (linhas[0] if linhas else [])):
        # Layout agregado: um total só, sem os dias, então só resta creditar ao
        # mês do nome do arquivo.
        mes = mes_do_arquivo(caminho)
        if mes is None:
            return {}
        if corte is not None and mes == chave_mes(corte.year, corte.month):
            print(f"  [corte] {os.path.basename(caminho)}: export agregado, sem "
                  f"quebra por dia — não foi possível cortar em {corte:%d/%m}")
        col = pd.read_excel(caminho)[SUC_COL_AGREGADA] if caminho.lower().endswith(
            (".xlsx", ".xls")) else pd.read_csv(caminho)[SUC_COL_AGREGADA]
        return {mes: int(pd.to_numeric(col, errors="coerce").fillna(0).sum())}

    totais = {}
    for dia, total in diarias:
        if corte is not None and dia > corte:
            continue
        chave = chave_mes(dia.year, dia.month)
        totais[chave] = totais.get(chave, 0) + int(re.sub(r"\D", "", total) or 0)
    return totais


# ── Histórico manual (meses fechados) ─────────────────────────────────────────
# Os meses antigos deixam de ser recalculados a partir das extrações e passam a
# vir da planilha de acompanhamento, convertida por converter_historico.py. As
# extrações de um mês fechado não são estáveis — o Octane continua mexendo em
# defeito antigo, e o relatório do mês passado mudava sozinho de uma semana para
# a outra.
ARQ_HISTORICO = "historico.xlsx"
COLS_HISTORICO = ["Jornada", "Ano", "Mes", "Realizado", "ProRata", "Planejado",
                  "Sucessos", "Falhas"]
# Coluna opcional: o histórico da linha "Projeção s/ Legado Aberto" do Base
# Residencial (verde-escura), como o "Planejado" é o da linha verde.
COL_PLANEJADO_SEM_LEGADO = "PlanejadoSemLegado"


def _historico_bruto(base_dir):
    """Lê o arquivo de histórico manual. Devolve () quando ele não existe."""
    caminho = os.path.join(raiz_dados(base_dir), ARQ_HISTORICO)
    if not os.path.isfile(caminho):
        return ()
    return _historico_cacheado(caminho, *_assinatura(caminho))


@functools.lru_cache(maxsize=8)
def _historico_cacheado(caminho, _mtime, _tamanho):
    """
    Corpo de `_historico_bruto`. Data e tamanho do arquivo só compõem a chave,
    como nos outros leitores: o histórico editado é relido na hora. Antes a
    chave era só a pasta, e a fotografia gravada logo depois de digitar o
    fechamento de set/26 (3,24% no Base + Cross Sell) saiu com o histórico
    antigo — 3,40%, recalculado das extrações.
    """
    tabela = pd.read_excel(caminho)
    faltando = [c for c in COLS_HISTORICO if c not in tabela.columns]
    if faltando:
        raise ValueError(
            f"{ARQ_HISTORICO} sem as colunas {faltando}. "
            f"Esperado: {', '.join(COLS_HISTORICO)}"
        )

    def _num(linha, col):
        v = pd.to_numeric(linha[col], errors="coerce")
        return None if pd.isna(v) else float(v)

    registros = []
    for _, linha in tabela.iterrows():
        jornada = str(linha["Jornada"]).strip()
        ano = pd.to_numeric(linha["Ano"], errors="coerce")
        mes = pd.to_numeric(linha["Mes"], errors="coerce")
        if not jornada or jornada.lower() == "nan" or pd.isna(ano) or pd.isna(mes):
            continue
        registros.append({
            "jornada": jornada,
            "chave": chave_mes(int(ano), int(mes)),
            "realizado": _num(linha, "Realizado"),
            "prorata": _num(linha, "ProRata"),
            "planejado": _num(linha, "Planejado"),
            # Opcional: só o Base Residencial usa, e históricos antigos não têm.
            "planejado_sem_legado": (_num(linha, COL_PLANEJADO_SEM_LEGADO)
                                     if COL_PLANEJADO_SEM_LEGADO in tabela.columns else None),
            "sucessos": _num(linha, "Sucessos"),
            "falhas": _num(linha, "Falhas"),
        })
    return tuple(registros)


def historico_manual(base_dir, jornada, jornadas_combo):
    """
    Meses fechados digitados para a jornada pedida, como {chave: {...}}.

    Cada mês traz `realizado` (a taxa de fallout), `prorata` e `planejado` (as
    linhas do gráfico) e, quando informados, `sucessos` e `falhas`.

    A planilha de acompanhamento é a fonte oficial dos meses fechados e traz
    percentual, não volume — por isso volume é opcional e esses meses aparecem
    sem a linha de "Volume de Pedidos" até alguém preenchê-lo.

    A linha da própria `jornada` vence; só na falta dela o volume sai da soma
    das jornadas cruas de `jornadas_combo`. A combinação precisa de linha
    própria porque **ela não é a soma das partes**: um mesmo pedido pode falhar
    em mais de uma jornada, e o Consolidado descarta a repetição por
    OrderNumber. Com os dados de hoje são 66 pedidos nessa situação — somar as
    jornadas inflava as falhas de maio em 28 e as de junho em 11.
    """
    cruas = set(jornadas_combo)
    proprios, das_partes = {}, {}

    for r in _historico_bruto(base_dir):
        if all(r[c] is None for c in ("realizado", "prorata", "planejado",
                                      "planejado_sem_legado")):
            continue
        if r["jornada"] == jornada:
            proprios[r["chave"]] = {
                "realizado": r["realizado"], "prorata": r["prorata"],
                "planejado": r["planejado"],
                "planejado_sem_legado": r["planejado_sem_legado"],
                "sucessos": r["sucessos"], "falhas": r["falhas"],
            }
        elif r["jornada"] in cruas and r["sucessos"] is not None and r["falhas"] is not None:
            acc = das_partes.setdefault(r["chave"], {"sucessos": 0.0, "falhas": 0.0,
                                                     "realizado": None, "prorata": None,
                                                     "planejado": None,
                                                     "planejado_sem_legado": None})
            acc["sucessos"] += r["sucessos"]
            acc["falhas"] += r["falhas"]

    for chave, acc in das_partes.items():
        if chave in proprios:
            continue
        total = acc["sucessos"] + acc["falhas"]
        acc["realizado"] = acc["falhas"] / total * 100 if total else 0.0
        proprios[chave] = acc

    # Devolve todos os meses declarados, inclusive os que só trazem a meta
    # "Planejado". Quem decide o que é mês fechado — e portanto sai do bruto —
    # é `carregar_base`, pelo `realizado`.
    return proprios


def _tem_volume(v):
    """True se o mês do histórico traz sucessos e falhas digitados."""
    return v["sucessos"] is not None and v["falhas"] is not None


def meses_do_historico(declarados):
    """
    Dos meses de `historico_manual`, os que o histórico fecha — os que saem dos
    agregados da extração e passam a vir da planilha (ver `carregar_base`).

    Mês fechado é o que traz a taxa digitada **ou** o volume. Os demais apenas
    acrescentam a meta "Planejado" e seguem vindo da extração. Volume e taxa são
    independentes de propósito — a planilha de acompanhamento traz a taxa, e o
    volume é digitado à parte quando existir.

    Função à parte para a fotografia (fotografia.py) marcar de onde veio cada
    mês com a mesma regra do cálculo.
    """
    return {c: v for c, v in declarados.items()
            if v["realizado"] is not None or _tem_volume(v)}


def limpar_cache_historico():
    """
    Esquece o histórico manual. A edição já é relida sozinha (a chave é a data
    e o tamanho do arquivo); isto fica para o "Atualizar dados" reler tudo.
    """
    _historico_cacheado.cache_clear()


def carregar_base(base_dir, jornada, hoje=None):
    """
    Lê falhas + sucessos (CSV) e Octane (Excel) para a jornada informada,
    faz o join e devolve (df, df_octane, resumo, jornadas_combo).

    `jornada` pode ser um nome de pasta única, "Base + Cross Sell" ou
    "Consolidado" — a resolução para as pastas reais é feita aqui. `hoje` só
    decide o corte (ver `data_corte`).
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
    # Um corte só para o relatório inteiro, calculado sobre todas as jornadas —
    # e não sobre `jornadas_combo` — para que ver o NBA sozinho não corte num dia
    # diferente de vê-lo dentro do Consolidado.
    corte = data_corte(base_dir, hoje)

    partes_falhas, partes_prontas, sucessos_rpa = [], [], []
    partes_extracao = []
    for j in jornadas_combo:
        pasta = os.path.join(raiz, j)
        formato = formato_jornada(pasta)
        # `Jornada` guarda a pasta de origem de cada linha. Sem isso, dentro do
        # Consolidado ou de "Base + Cross Sell" não dá para aplicar uma regra que
        # vale só para uma jornada — é o caso do BSIM (ver CAT_BSIM).
        if formato == "nba":
            df_nba, suc_nba = _ler_jornada_nba(pasta, corte=corte)
            partes_prontas.append(df_nba.assign(Jornada=j))
            sucessos_rpa.append(suc_nba)
        elif formato == "rpa":
            df_rpa, suc_rpa = _ler_jornada_rpa(pasta, corte=corte)
            partes_falhas.append(df_rpa.assign(Jornada=j))
            sucessos_rpa.append(suc_rpa)
        elif formato == "extracao":
            partes_extracao.append(_ler_falhas_jornada(pasta).assign(Jornada=j))
        else:
            raise FileNotFoundError(f"{pasta} não parece uma jornada (sem falhas/ e sem CSV)")

    if partes_extracao:
        df_ext = pd.concat(partes_extracao, ignore_index=True)

        # O export do CRM às vezes traz o número do defeito com espaços não-quebráveis
        # no fim (ex.: "233538\xa0\xa0"). Sem limpar, pd.to_numeric devolve NaN e o
        # pedido acaba contado como "Falta Associar ao Defeito/US" mesmo tendo DFT.
        df_ext["DefectNumber_orig"] = _limpar_defeito(df_ext["DefectNumber__c"])
        # Alguns registros vêm com o prefixo do próprio sistema ("DFT 232143"). Só o
        # prefixo DFT é removido: os demais referenciam outros sistemas e não têm
        # número no Octane. INC- e PDST- ganham defeito mais abaixo, pelo catálogo
        # do Jira; DDP- e OS seguem sem defeito associado.
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
    df_octane = _octane(raiz, base_dir, hoje)

    # ── Join falhas ← Octane ────────────────────────────────────────────
    # Pelo número, só os itens que têm número: os de ID em texto (DDP-796) ficam
    # com -1, o mesmo do pedido sem DFT, e o join casava os dois.
    _por_numero = df_octane[df_octane["DefectNumber__c"] > 0].drop(columns=COL_CHAVE_OCTANE)
    df = df_falhas.merge(_por_numero, on="DefectNumber__c", how="left")

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
    for marca in ("LegadoAberto", "ErroLegado", "ErroAgendamento", "BaseResidencial"):
        if marca not in df.columns:
            df[marca] = False
        df[marca] = df[marca].fillna(False).astype(bool)

    # Itens do Octane sem número (DDP-366): casam pelo rótulo do pedido, como os
    # chamados do Jira e os BRJ logo abaixo. Só onde o join por número não
    # trouxe nada.
    _por_rotulo = (df_octane[df_octane["DefectNumber__c"] <= 0]
                   .drop_duplicates(subset=COL_CHAVE_OCTANE).set_index(COL_CHAVE_OCTANE))
    if len(_por_rotulo):
        _rotulo = df["DefectNumber_orig"].astype(str).str.strip().str.upper()
        _casa_rotulo = _rotulo.isin(_por_rotulo.index) & df["DFT_Phase"].isna()
        if _casa_rotulo.any():
            _k = _rotulo[_casa_rotulo]
            for _col in ("DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Team",
                         "DFT_Type", "DFT_US"):
                df.loc[_casa_rotulo, _col] = _k.map(_por_rotulo[_col])
            df.loc[_casa_rotulo, "TemDefeito"] = True
            print(f"  [octane] {int(_casa_rotulo.sum())} pedidos associados a "
                  f"{_k.nunique()} itens sem número (DDP-…)")

    # Chamados do Jira, pela mesma lógica da MilestoneRPA acima: preenche só
    # onde o Octane não se posicionou. Hoje não há sobreposição — todo pedido
    # que casa aqui está sem fase —, mas o Octane segue sendo a fonte oficial.
    _jira = _catalogo_jira(raiz, base_dir)
    if _jira is not None and len(_jira):
        _chamado = df["DefectNumber_orig"].astype(str).str.strip()
        _casa_jira = _chamado.isin(_jira.index) & df["DFT_Phase"].isna()
        if _casa_jira.any():
            for _col in ("DFT_Phase", "DFT_BugfixMilestone", "DFT_Type"):
                df.loc[_casa_jira, _col] = _chamado[_casa_jira].map(_jira[_col])
            # Chamado do Jira é defeito como outro qualquer. Sem isto o pedido
            # continuaria em "Falta Associar ao Defeito/US", que é onde ele
            # estava justamente por o Octane não conhecer INC-* nem PDST-*.
            df.loc[_casa_jira, "TemDefeito"] = True
            print(f"  [jira] {int(_casa_jira.sum())} pedidos associados a "
                  f"{_chamado[_casa_jira].nunique()} chamados")
    # Chamados BRJ: o relatório faz o papel do Octane (ver ARQ_BRJ), então a
    # fase e a data de implantação dele valem sobre as da jornada — no Base
    # Residencial, a fase fixa "Em Tratamento" e a Data de Implantação da aba
    # Defeitos, que só fica onde o relatório não tem data. Nome, time e tipo
    # só preenchem o que a jornada não trouxe.
    _brj = _catalogo_brj(raiz, base_dir)
    if _brj is not None and len(_brj):
        _chave_brj = df["DefectNumber_orig"].astype(str).str.strip().str.upper()
        _casa_brj = _chave_brj.isin(_brj.index)
        if _casa_brj.any():
            _k = _chave_brj[_casa_brj]
            df.loc[_casa_brj, "DFT_Phase"] = _k.map(_brj["DFT_Phase"])
            df.loc[_casa_brj, "DFT_BugfixMilestone"] = (
                _k.map(_brj["DFT_BugfixMilestone"])
                .fillna(pd.to_datetime(df.loc[_casa_brj, "DFT_BugfixMilestone"], errors="coerce"))
                .fillna(_k.map(_brj["_resolvido"]))
            )
            for _col in ("DFT_Name", "DFT_Team", "DFT_Type"):
                _atual = df.loc[_casa_brj, _col].replace("", pd.NA)
                df.loc[_casa_brj, _col] = _atual.fillna(_k.map(_brj[_col]))
            df.loc[_casa_brj, "TemDefeito"] = True
            print(f"  [brj] {int(_casa_brj.sum())} pedidos associados a "
                  f"{_k.nunique()} chamados BRJ")
    # PME: o defeito que nenhum catálogo nomeou ("Paliativo RPA - 7", "NFCOM",
    # GML-…) leva o nome da coluna DESCRIÇÃO DO DEFEITO do próprio relatório. A
    # descrição muda de pedido para pedido no mesmo rótulo; vale a mais
    # frequente, para o defeito ter um nome só e não virar várias linhas na
    # tabela de ofensores.
    if RPA_COL_DESCRICAO in df.columns:
        _desc = df[RPA_COL_DESCRICAO].astype("string").str.strip().replace("", pd.NA)
        _rot = df["DefectNumber_orig"].astype(str).str.strip()
        _com_desc = _desc.notna()
        if _com_desc.any():
            _nome_do_rotulo = (_desc[_com_desc].groupby(_rot[_com_desc])
                               .agg(lambda s: s.value_counts().index[0]))
            _sem_nome = (df["DFT_Name"].astype("string").str.strip().replace("", pd.NA).isna()
                         & _rot.isin(_nome_do_rotulo.index))
            df.loc[_sem_nome, "DFT_Name"] = _rot[_sem_nome].map(_nome_do_rotulo)
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
    # Corte no último dia completo. As jornadas RPA/NBA já vieram cortadas dos
    # seus leitores (lá o mesmo arquivo dá falhas e sucessos, e os dois têm que
    # sair do mesmo recorte); aqui entram as de extração, e reaplicar nas outras
    # não muda nada.
    if corte is not None:
        _dia = df["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.date
        df = df[_dia.isna() | (_dia <= corte)].reset_index(drop=True)

    df["Mes"] = chave_de_datas(df["CreatedDate"].dt.tz_convert("America/Sao_Paulo"))

    # ── Sucessos por mês (mesmo denominador do fallout rate) ────────────
    # Nas jornadas RPA os sucessos já vieram contados junto com as falhas; nas
    # de extração eles estão em CSVs próprios, com o mês no nome do arquivo.
    partes_suc = []
    for j in jornadas_combo:
        if formato_jornada(os.path.join(raiz, j)) != "extracao":
            continue   # RPA e NBA já contaram os sucessos junto com as falhas
        arqs_suc = _arquivos_da_etapa(os.path.join(raiz, j), "sucessos", EXT_SUCESSOS)
        if not arqs_suc:
            raise FileNotFoundError(f"Nenhum arquivo em {os.path.join(raiz, j, 'sucessos')}")
        # Mapa mês -> total, e não soma: os arquivos se sobrepõem (cada um cobre
        # mês atual e anterior), então somá-los contaria o mesmo dia duas vezes.
        #
        # Vence o arquivo do próprio mês — o "jul26" manda em julho, mesmo que o
        # export de agosto também o contenha. Os três discordam (no Prospect,
        # 64.163 contra 76.561), e trocar a fonte mudaria meses já publicados.
        # Só quando não existe arquivo do mês é que outro entra: é o caso da PME
        # Fixa, que nasceu com um arquivo só e teria julho zerado.
        #
        # "Mais recente" é pela data que o export traz, e não pela data de
        # modificação (ver "Qual export é o mais recente" e
        # `_data_export_sucessos`).
        cands = {}
        for f in sorted(arqs_suc, reverse=True,
                        key=lambda c: (_data_export_sucessos(c) or SEM_DATA,
                                       os.path.getmtime(c))):
            for mes, total in sucessos_por_mes(f, corte=corte).items():
                cands.setdefault(mes, []).append((f, total))

        por_mes = {}
        for mes, lista in cands.items():
            # A lista vem do mais recente para o mais antigo, então o primeiro
            # do próprio mês é o vencedor da regra acima.
            arq, total = ([c for c in lista if mes_do_arquivo(c[0]) == mes] or lista)[0]
            # Mas quando o vencedor é um export com quebra por dia, o nome não
            # prova nada: outro export diário pode cobrir mais dias do mesmo
            # mês. Foi o que aconteceu no Prospect em ago/26 — o export tirado
            # em 27/08 "é de agosto" pelo nome e ganhava do de 03/09, que tem o
            # mês fechado; agosto saía 9.925 sucessos menor, e o fallout do mês
            # subia junto. Entre dois diários, portanto, vence a cobertura.
            #
            # Agregado não entra nessa disputa: ele traz um total só, sem dizer
            # quantos dias cobre, e trocar a fonte mexeria em meses publicados
            # (no Prospect, jul/26 fecha em 64.163 pelo agregado e 76.570 pelo
            # diário — diferença de fonte, não de cobertura).
            dias_atual = _dias_cobertos(arq, mes, corte)
            if dias_atual:
                melhor, dias_melhor = max(
                    ((c, _dias_cobertos(c[0], mes, corte)) for c in lista),
                    key=lambda par: par[1],
                )
                if dias_melhor > dias_atual:
                    arq, total = melhor
            por_mes[mes] = total
        if por_mes:
            partes_suc.append(pd.Series(por_mes, name="Sucessos", dtype="int64"))

    series_suc = partes_suc + sucessos_rpa
    if not series_suc:
        raise FileNotFoundError("Nenhum dado de sucesso encontrado para a jornada")
    sucessos_mes = (
        pd.concat(series_suc).groupby(level=0).sum().rename("Sucessos")
    )

    # ── Histórico manual vence a extração, mas só nos agregados ────────────
    # O corte é entre `df` e `resumo`, e não entre meses:
    #   `df`     é a extração pedido a pedido e continua inteira, inclusive nos
    #            meses fechados — é dela que vivem as abas do app (Distribuição,
    #            Erros) e a contagem de ocorrências por defeito dos Ofensores.
    #   `resumo` é o número do relatório, e aí o histórico vence: `falhas_mes` e
    #            `sucessos_mes` ignoram o mês fechado para ele não ser contado
    #            duas vezes, uma pela extração e outra pelo que foi digitado.
    # Por isso o detalhe de um mês fechado pode não somar exatamente a taxa
    # exibida: a taxa é a oficial da planilha de acompanhamento, e a extração
    # continua se mexendo depois do fechamento.
    declarados = historico_manual(base_dir, jornada, jornadas_combo)
    manual = meses_do_historico(declarados)
    # Total de pedidos que a extração tem em cada mês, fechado ou não: é o
    # denominador da Distribuição e dos Ofensores (que contam o `df`) quando o
    # mês fechado foi digitado só com a taxa. Sem ele o "%" de cada categoria
    # saía NaN e as abas Projeção e Ofensores quebravam — set/26 do Base + Cross
    # Sell, fechado em 3,24% sem Sucessos/Falhas.
    total_extracao = (df.groupby("Mes").size()
                      .add(sucessos_mes, fill_value=0).astype("float64"))
    if manual:
        sucessos_mes = sucessos_mes[~sucessos_mes.index.isin(manual)]

    falhas_mes = df[~df["Mes"].isin(manual)].groupby("Mes").size().rename("Falhas")
    resumo = pd.DataFrame({"Falhas": falhas_mes, "Sucessos": sucessos_mes}).fillna(0)
    resumo = resumo.astype("float64")
    resumo["Total"] = resumo["Falhas"] + resumo["Sucessos"]
    resumo["Pct"] = resumo["Falhas"] / resumo["Total"] * 100
    resumo["PctEsperada"] = pd.Series(dtype="float64")

    for chave, valores in manual.items():
        # Volume fica NaN quando não foi digitado; a taxa não é inventada a
        # partir dele nem o contrário.
        tem_vol = _tem_volume(valores)
        resumo.loc[chave, "Falhas"] = valores["falhas"]
        resumo.loc[chave, "Sucessos"] = valores["sucessos"]
        resumo.loc[chave, "Total"] = (
            valores["falhas"] + valores["sucessos"] if tem_vol else float("nan")
        )

        # A taxa da planilha manda: ela é o número oficial e o volume pode estar
        # arredondado ou vir de um recorte diferente. Sem taxa, ela sai do
        # volume — é o único jeito de o mês existir no gráfico.
        taxa = valores["realizado"]
        if taxa is None and tem_vol:
            total_chave = valores["falhas"] + valores["sucessos"]
            taxa = valores["falhas"] / total_chave * 100 if total_chave else 0.0
        elif taxa is not None and tem_vol:
            total_chave = valores["falhas"] + valores["sucessos"]
            taxa_do_volume = valores["falhas"] / total_chave * 100 if total_chave else 0.0
            if abs(taxa_do_volume - taxa) > 0.05:
                print(f"  [historico] {jornada} {MESES_LABEL[chave]}: a taxa digitada "
                      f"({taxa:.2f}%) não bate com o volume "
                      f"({taxa_do_volume:.2f}% = {valores['falhas']:.0f} de "
                      f"{total_chave:.0f}). Vale a taxa; confira o volume.")
        resumo.loc[chave, "Pct"] = taxa
        # Projeção congelada: o ponto de cada mês sai do bruto do mês anterior,
        # que nos meses fechados não existe mais. Sem o valor digitado o cálculo
        # não falha — devolve a curva real defasada em um mês, que parece certa
        # e não é.
        resumo.loc[chave, "PctEsperada"] = valores["prorata"]

    # Só depois do laço acima, quando as linhas dos meses fechados já existem:
    # atribuir a coluna antes alinharia a Série apenas aos meses de extração e
    # o histórico inteiro sairia vazio.
    resumo["PctPlanejado"] = pd.Series(
        {c: v["planejado"] for c, v in declarados.items() if v["planejado"] is not None},
        dtype="float64")
    resumo["PctPlanejadoSemLegado"] = pd.Series(
        {c: v["planejado_sem_legado"] for c, v in declarados.items()
         if v.get("planejado_sem_legado") is not None},
        dtype="float64")
    resumo["TotalExtracao"] = total_extracao
    resumo = resumo.sort_index()

    return df, df_octane, resumo, jornadas_combo


def total_do_mes(resumo, mes):
    """
    Pedidos do mês para servir de denominador às contagens do `df`: o total do
    resumo e, no mês fechado digitado só com a taxa, o total da extração.
    NaN quando nenhum dos dois existe.
    """
    if mes not in resumo.index:
        return float("nan")
    total = resumo.loc[mes, "Total"]
    if pd.isna(total) and "TotalExtracao" in resumo.columns:
        total = resumo.loc[mes, "TotalExtracao"]
    return float(total)


def dia_local(serie):
    """
    Coluna de data reduzida ao dia e sem fuso, para datas de origens diferentes
    se compararem.

    `CreatedDate` vem em UTC e `DFT_BugfixMilestone` vem sem fuso — é data de
    calendário, não instante. Comparar as duas direto levanta exceção, e passar
    o milestone para UTC o empurraria para o dia anterior no fuso de Brasília:
    o pedido criado no próprio dia da entrega passaria a contar como posterior
    a ela.
    """
    dt = pd.to_datetime(serie, errors="coerce")
    if isinstance(dt.dtype, pd.DatetimeTZDtype):
        dt = dt.dt.tz_convert("America/Sao_Paulo").dt.tz_localize(None)
    return dt.dt.normalize()


# Defeitos que o Octane dá como corrigidos mas cuja correção não resolveu: o
# erro continuou depois da entrega. Todos os pedidos deles vão para "Em
# Avaliação por MOPs" e para a aba "Precisa de Correção" — inclusive os
# criados antes da entrega, que só pela regra do pós-entrega ficariam em
# Tratado, Resolvido ou BSIM. A marcação é manual: acrescente o defeito aqui,
# como texto e do jeito que aparece no DefectKey (vale rótulo, como "BRJ-464").
#
# O 232882 (Confirmação Eletrônica) saiu da lista em 18/09/2026: foi marcado
# com a entrega em 12/09, e os pedidos de 13/09 pareciam o erro seguindo.
# Com a entrega corrigida para 14/09 no Octane, nenhum pedido veio depois dela.
# Se o erro voltar, a regra do pós-entrega pega os pedidos novos sozinha.
DFTS_CORRECAO_INEFICAZ = set()


def us_em_aberto(df):
    """
    Pedidos cujo DFT tem US de Melhoria ainda em andamento.

    Nesses o catálogo já trocou fase e data de entrega pelas da US (ver
    `_octane_cacheado`), e a data da US em andamento é planejamento, não
    entrega: o 214711 aponta para a US 216759, em Backlog com entrega em
    23/04/2025, e todo pedido dele caía como criado depois da entrega. Sem a
    coluna DFT_US (fotografia gravada antes dela existir) não marca ninguém.
    """
    if "DFT_US" not in df.columns:
        return pd.Series(False, index=df.index)
    fase = df["DFT_Phase"].fillna("").astype(str).str.strip()
    return df["DFT_US"].notna() & ~fase.isin(FASES_US_ENCERRADAS)


def precisa_correcao(df):
    """
    Pedidos cujo defeito associado não vai resolvê-los sozinho, como três
    máscaras sobre `df`: (fase_mops, pos_entrega, ineficaz).

      fase_mops   — o defeito foi cancelado ou rejeitado (FASES_MOPS);
      pos_entrega — o pedido foi criado depois da data de entrega do defeito:
                    a correção entrou em produção e o erro continuou. Não vale
                    para DFT com US de Melhoria em andamento (`us_em_aberto`):
                    ali a data é a da US, que ainda não foi entregue;
      ineficaz    — o defeito está em DFTS_CORRECAO_INEFICAZ: a correção não
                    resolveu, então vale para todos os pedidos dele.

    É a lista da aba "Precisa de Correção" do app e, na Distribuição, o que vai
    para "Em Avaliação por MOPs". A mesma função nos dois lugares para a aba e
    a categoria não divergirem, como acontecia quando cada um tinha a sua
    regra: o pós-entrega ficava em Resolvido, Tratado ou BSIM — no PME de
    set/26 eram 219 pedidos em Resolvido.

    Só conta pedido com defeito (`TemDefeito`) e fora do Legado Aberto e da
    Funcionalidade idêntica ao legado, que no Base Residencial têm categoria
    própria.
    """
    if "TemDefeito" in df.columns:
        elegivel = df["TemDefeito"].fillna(False).astype(bool)
    else:
        elegivel = (df["DefectNumber__c"] > 0) & (df["DefectNumber__c"] != 999999)
    for marca in ("LegadoAberto", "ErroLegado", "ErroAgendamento"):
        if marca in df.columns:
            elegivel = elegivel & ~df[marca].fillna(False).astype(bool)
    fase_mops = elegivel & df["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS)
    entrega = dia_local(df["DFT_BugfixMilestone"])
    pos_entrega = (elegivel & entrega.notna() & (dia_local(df["CreatedDate"]) > entrega)
                   & ~us_em_aberto(df))
    chave = df["DefectKey"] if "DefectKey" in df.columns else df["DefectNumber__c"]
    ineficaz = elegivel & chave.astype(str).str.strip().isin(DFTS_CORRECAO_INEFICAZ)
    return fase_mops, pos_entrega, ineficaz


def hoje_utc(valor=None):
    """
    A data de referência das regras que dependem de "hoje" — os 15 dias da
    avaliação de eficácia, a entrega que já passou —, como meia-noite UTC.

    Sem `valor`, é hoje, como sempre foi. Com ele (data, texto AAAA-MM-DD ou
    Timestamp), é o dia informado: a fotografia de uma extração antiga precisa
    das categorias como elas estavam naquele dia, e não como estariam hoje.
    Um Timestamp sem fuso é lido como UTC; comparar um sem fuso com o marco de
    entrega, que `categorizar` converte para UTC, levantaria exceção.
    """
    if valor is None:
        return pd.Timestamp.now(tz="UTC").normalize()
    ts = pd.Timestamp(valor)
    ts = ts.tz_localize("UTC") if ts.tzinfo is None else ts.tz_convert("UTC")
    return ts.normalize()


def categorizar(df, mes, hoje=None):
    """
    Divide os registros de falha de um mês nas categorias de negócio do fallout.
    Retorna (df_mes, cats, total), onde `cats` é um dict {nome_categoria: DataFrame}.

    `hoje` pode ser fixado (ex.: para reprodutibilidade em testes ou para a
    fotografia de uma extração antiga); por padrão usa a data atual. Ver
    `hoje_utc`.
    """
    hoje = hoje_utc(hoje)
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
    _incidente_mask = _rotulo == ROTULO_INCIDENTE

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
    if "ErroLegado" in df_mes.columns:
        # Cancelado na origem e "ErroProcesso" ao mesmo tempo é Legado Aberto: o
        # cancelamento vence (pedido de 08/10).
        _erro_legado = df_mes["ErroLegado"].fillna(False).astype(bool) & ~_legado
    else:
        _erro_legado = pd.Series(False, index=df_mes.index)
    # As regras só do Base Residencial (ver CAT_SCRIPT_RPA e CAT_AGENDAMENTO).
    def _marca(coluna):
        return (df_mes[coluna].fillna(False).astype(bool) if coluna in df_mes.columns
                else pd.Series(False, index=df_mes.index))
    _residencial = _marca("BaseResidencial")
    _agendamento = _marca("ErroAgendamento") & ~_legado & ~_erro_legado
    _999999 = df_mes["DefectNumber__c"] == 999999

    # BSIM tem precedência: o pedido do Cross Sell cujo defeito é do time de
    # Confirmação Eletrônica sai da categoria em que cairia e vem para cá. Só as
    # quatro categorias abaixo precisam excluí-lo — as outras dependem de rótulo
    # sem defeito (Falta Associar, Outros times, Erro de Processo), de número
    # reservado (Falha Pontual) ou de outra jornada (Legado Aberto), e em nenhum
    # desses casos existe time do Octane para casar.
    #
    # Só enquanto o defeito está em andamento: com a correção entregue
    # (`_encerrado`), o pedido vai para Tratado ou Resolvido como os demais. Antes
    # o BSIM valia em qualquer fase, e um defeito corrigido cujo erro parou (o
    # 232882, entregue em 14/09/2026) ficava em BSIM em vez de em eficácia.
    if {"Jornada", "DFT_Team"} <= set(df_mes.columns):
        _bsim = (
            (df_mes["Jornada"] == JORNADA_BSIM)
            & (df_mes["DFT_Team"].fillna("").astype(str).str.strip() == TIME_BSIM)
            & ~_encerrado
        )
    else:
        _bsim = pd.Series(False, index=df_mes.index)

    # Tudo o que a aba "Precisa de Correção" lista vai para MOPs e sai de onde
    # estaria — inclusive do BSIM, cuja precedência acaba aqui: pedido criado
    # depois da entrega é sinal de que a correção não pegou, seja de qual time
    # for o defeito.
    _fase_mops, _pos_entrega, _ineficaz = precisa_correcao(df_mes)
    _precisa = _fase_mops | _pos_entrega | _ineficaz

    cats = {
        "Em Tratamento/Avaliação pela Squad": df_mes[
            _tem_defeito &
            ~_fase.isin(FASES_FORA_TRATAMENTO) &
            ~_legado &
            ~_email_mops_mask &
            ~_incidente_mask &
            ~_encerrado &
            ~_bsim &
            ~_precisa
        ],
        # `~_legado` nas duas seguintes: com a fase vinda do RelatorioBRJ, um
        # pedido do Legado Aberto pode ter o chamado concluído, e entraria em
        # duas categorias.
        "Resolvido": df_mes[
            _corrigido & _tem_ms & (milestone_dt < quinze_dias) & ~_bsim & ~_precisa
            & ~_legado],
        # No Base Residencial, Falha Pontual é o pedido sem defeito associado; o
        # 999999 de lá é CAT_SCRIPT_RPA.
        "Falha Pontual": df_mes[
            (_999999 & ~_residencial)
            | (_residencial & ~_tem_defeito & ~_999999 & ~_legado)],
        "Falta Associar ao Defeito/US": df_mes[
            ~_tem_defeito &
            ~_legado &
            ~_999999 &
            ~_residencial &
            ~_outros_mask & ~_erro_proc_mask & ~_email_mops_mask &
            ~_incidente_mask
        ],
        "Tratado - Em avaliação de eficácia": df_mes[
            _corrigido & _tem_ms & (milestone_dt >= quinze_dias)
            & (milestone_dt <= hoje) & ~_bsim & ~_precisa & ~_legado
        ],
        "Em Avaliação por MOPs": df_mes[_precisa | (_email_mops_mask & ~_bsim)],
        CAT_BSIM: df_mes[_bsim & ~_precisa],
        # As duas subcategorias juntas (ver `subgrupo_legado`).
        CAT_LEGADO: df_mes[_legado | _erro_legado],
        "Em avaliação - Outros times": df_mes[_outros_mask],
        "Falha no Processo Usuário": df_mes[_erro_proc_mask],
        CAT_INCIDENTE: df_mes[_incidente_mask],
        CAT_SCRIPT_RPA: df_mes[_residencial & _999999 & ~_legado],
        CAT_AGENDAMENTO: df_mes[_agendamento],
    }
    # A Funcionalidade idêntica ao legado tem precedência sobre as demais: o
    # pedido sai de onde estaria, tenha defeito ou não, e fica no Legado Aberto.
    # O Erro no Agendamento vem logo depois: sai de todas, menos do Legado.
    cats = {nome: (parte if nome == CAT_LEGADO else parte[~_erro_legado[parte.index]])
            for nome, parte in cats.items()}
    cats = {nome: (parte if nome in (CAT_LEGADO, CAT_AGENDAMENTO)
                   else parte[~_agendamento[parte.index]])
            for nome, parte in cats.items()}
    total = len(df_mes)
    return df_mes, cats, total


def categorias_por_pedido(df, hoje=None):
    """
    Categoria e subgrupo de cada pedido, em todos os meses de `df` — e não só no
    mês do relatório: a comparação de uma semana com a outra pode olhar
    qualquer mês que as duas extrações tenham.

    É a mesma classificação da Distribuição, pedido a pedido: a fotografia grava
    com ela, e o slide diário a soma por dia.

    Devolve (categoria, subgrupo, sobrepostos). Pela regra, as categorias de
    `categorizar` não se sobrepõem; se um pedido cair em duas, ele fica com a
    primeira (na ordem da Distribuição) e entra em `sobrepostos`, que vai para
    o meta.json. Sem essa conta, a fotografia contaria o pedido uma vez e o
    app duas, e ninguém saberia por quê.
    """
    categoria = pd.Series(pd.NA, index=df.index, dtype="string")
    subgrupo = pd.Series(pd.NA, index=df.index, dtype="string")
    sobrepostos = 0
    for mes in sorted(df["Mes"].dropna().unique()):
        _df_mes, cats, _total = categorizar(df, mes, hoje=hoje)
        for nome, linhas in cats.items():
            livres = linhas.index[categoria.loc[linhas.index].isna()]
            sobrepostos += len(linhas) - len(livres)
            categoria.loc[livres] = nome
        em_trat = cats[CAT_TRATAMENTO]
        subgrupo.loc[em_trat.index] = subgrupo_tratamento(em_trat)
        legado = cats[CAT_LEGADO]
        subgrupo.loc[legado.index] = subgrupo_legado(legado)
    return categoria, subgrupo, sobrepostos


# ── Fallout por dia (slides diários) ─────────────────────────────────────────
DIAS_FALLOUT_DIARIO = 15


def sucessos_por_dia(base_dir, jornada):
    """
    {dia: sucessos} de uma jornada de extração, tirado dos exports com quebra
    diária. Um dia que aparece em mais de um export fica com o mais recente, na
    mesma ordem que `carregar_base` usa para escolher o export do mês.
    """
    pasta = os.path.join(raiz_dados(base_dir), jornada)
    arquivos = _arquivos_da_etapa(pasta, "sucessos", EXT_SUCESSOS)
    por_dia = {}
    for f in sorted(arquivos, reverse=True,
                    key=lambda c: (_data_export_sucessos(c) or SEM_DATA,
                                   os.path.getmtime(c))):
        for dia, total in _linhas_diarias(f):
            por_dia.setdefault(dia, int(re.sub(r"\D", "", total) or 0))
    return por_dia


def fim_da_visao(base_dir, visao, hoje=None):
    """
    Último dia completo que todas as jornadas da visão têm, sem passar do corte
    geral (`data_corte`). É onde terminam as janelas por dia (slide diário, mapa
    de calor): em 28/10 o Prospect parava em 07/10, e o Consolidado de 13 a
    27/10 era só o Base + Cross Sell — e o Prospect sozinho, uma janela vazia.
    None sem data nenhuma.
    """
    corte = data_corte(base_dir, hoje)
    if corte is None:
        return None
    um_dia = _dt.timedelta(days=1)
    fins = {j: (d if completo else d - um_dia) for j, d, completo in fim_das_jornadas(base_dir)}
    da_visao = [fins[j] for j in resolver_jornadas(visao, base_dir) if j in fins]
    return min(da_visao + [corte])


# No fallout por dia, "Em Tratamento/Avaliação pela Squad" sai em duas colunas:
# o que está com User Story e o que está com defeito (DFT, inclusive sem tipo).
CAT_TRAT_US = "Em Tratamento pela Squad – US"
CAT_TRAT_DFT = "Em Tratamento pela Squad – DFT"


def fallout_por_dia(df, base_dir, visao, dias=DIAS_FALLOUT_DIARIO, hoje=None):
    """
    Os `dias` dias até o corte, um por linha: uma coluna por categoria da
    Distribuição com as falhas do dia, e "Falhas", "Sucessos" e "Pedidos".
    Em Tratamento vem dividido em CAT_TRAT_US e CAT_TRAT_DFT, pelo tipo do item
    do Octane, a mesma regra de "US s/ data" em `subgrupo_tratamento`.

    `df` é o de `carregar_base` da visão — já cortado e sem pedido repetido — e
    a categoria de cada pedido é a da Distribuição (`categorias_por_pedido`),
    vista hoje. Dia sem sucessos em alguma jornada da visão fica com
    "Sucessos" e "Pedidos" vazios: o total estaria incompleto, e o percentual
    do dia sairia inflado.
    """
    corte = fim_da_visao(base_dir, visao, hoje)
    if corte is None:
        raise ValueError("Nenhuma data legível nas extrações: confira se as pastas das "
                         f"jornadas estão direto em {raiz_dados(base_dir)}.")
    janela = [corte - _dt.timedelta(days=i) for i in range(dias - 1, -1, -1)]
    dia = dia_local(df["CreatedDate"]).dt.date
    no_periodo = df[dia.isin(set(janela))]
    categoria, _subgrupo, _sobrepostos = categorias_por_pedido(no_periodo, hoje)
    em_trat = (categoria == CAT_TRATAMENTO).fillna(False).astype(bool)
    e_us = no_periodo["DFT_Type"].fillna("").astype(str).str.strip() == TIPO_US
    categoria = categoria.copy()
    categoria[em_trat & e_us] = CAT_TRAT_US
    categoria[em_trat & ~e_us] = CAT_TRAT_DFT
    falhas = (pd.crosstab(dia[no_periodo.index], categoria)
                .reindex(index=janela, fill_value=0))
    tabela = falhas.reindex(columns=[c for c in falhas.columns if falhas[c].sum() > 0])
    tabela["Falhas"] = falhas.sum(axis=1)

    sucessos = pd.Series(0.0, index=janela)
    for jornada in resolver_jornadas(visao, base_dir):
        da_jornada = sucessos_por_dia(base_dir, jornada)
        sucessos = sucessos + pd.Series([da_jornada.get(d, float("nan")) for d in janela],
                                        index=janela)
    tabela["Sucessos"] = sucessos
    tabela["Pedidos"] = tabela["Falhas"] + tabela["Sucessos"]
    tabela.index.name, tabela.columns.name = "Dia", None
    return tabela


def subgrupo_legado(legado):
    """
    Subcategoria de cada pedido do Legado Aberto, alinhada ao índice de
    `legado`: "Feito manualmente no Legado" quando a origem marca CANCELADO
    (`LegadoAberto`, que vence quando os dois valem) e "Funcionalidade idêntica
    ao legado" quando o erro agrupado é "ErroProcesso" (`ErroLegado`).
    """
    manual = (legado["LegadoAberto"].fillna(False).astype(bool)
              if "LegadoAberto" in legado.columns
              else pd.Series(False, index=legado.index))
    subgrupo = pd.Series(SUB_LEGADO_FUNCIONALIDADE, index=legado.index, dtype=object)
    subgrupo[manual] = SUB_LEGADO_MANUAL
    return subgrupo


def subgrupo_tratamento(em_trat):
    """
    Subgrupo de cada pedido de "Em Tratamento/Avaliação pela Squad", alinhado
    ao índice de `em_trat`: um dos quatro de SUBGRUPOS_TRATAMENTO — US se o item
    do Octane é User Story, DFT nos demais (inclusive sem tipo), e "c/ data"
    quando tem data de entrega.

    Uma função só para a distribuição do slide (pipeline.py) e a fotografia
    (fotografia.py) não dividirem o tratamento de jeitos diferentes.
    """
    tem_data = em_trat["DFT_BugfixMilestone"].notna()
    # `astype(str)` antes do `.str`: numa jornada sem nenhum casamento com o
    # Octane a coluna é toda NaN (float), e vazia ela não aceita o `.str`.
    e_us = em_trat["DFT_Type"].fillna("").astype(str).str.strip() == TIPO_US
    subgrupo = pd.Series("DFT s/ data", index=em_trat.index, dtype=object)
    subgrupo[tem_data & ~e_us] = "DFT c/ data"
    subgrupo[~tem_data & e_us] = "US s/ data"
    subgrupo[tem_data & e_us] = "US c/ data"
    return subgrupo


def tipo_do_subgrupo(subgrupo):
    """US ou DFT, o primeiro nível de um subgrupo de `subgrupo_tratamento`."""
    return subgrupo.astype(str).str.split(" ", n=1).str[0]


def so_rotulo(rotulo, nivel):
    """
    True no US e no DFT do Em Tratamento: a linha só agrupa o c/ data e o s/ data
    de baixo, e aparece sem o %. O valor segue na lista, para a ordem e para a
    linha sumir quando o tipo não tem pedido.
    """
    return int(nivel) == 1 and rotulo in SUBGRUPOS_TRATAMENTO


def rotulo_dist(rotulo, nivel):
    """
    O rótulo como aparece na Distribuição: no segundo nível do Em Tratamento,
    sem o tipo na frente ("US c/ data" → "c/ data"), que já está na linha de cima.
    """
    if int(nivel) >= 2:
        for tipo in SUBGRUPOS_TRATAMENTO:
            if rotulo.startswith(tipo + " "):
                return rotulo[len(tipo) + 1:]
    return rotulo
