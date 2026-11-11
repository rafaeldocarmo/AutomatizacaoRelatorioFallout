# -*- coding: utf-8 -*-
"""
Planilha de Ofensores: uma tabela dinâmica com todas as jornadas.

São duas abas:

* `Dinâmica` — a distribuição do fallout na mesma hierarquia da tabela do slide:
  classificação (Sucesso / Falha / Falha Pontual) → agrupamento do fallout →
  defeito, status, data de previsão, time e tipo. Nas colunas, o volume de
  pedidos e a participação no mês. **Jornada** e **motivo do pedido**
  (`vlocity_cmt__Reason__c`) são campos de filtro: escolher um motivo tira da
  tabela quem não tem pedido nele.
* `Dados` — uma linha por combinação dessas colunas, que é a origem da dinâmica.

Por a participação ser sobre o total de pedidos do mês, e não só sobre o
fallout, entram também os **sucessos**: uma linha por jornada e mês, com a
contagem que vem de `resumo`. É o que faz a coluna do mês fechar em 100%.
Sucesso não tem defeito nem motivo — a extração de sucessos é um agregado por
dia, não uma linha por pedido —, então essa linha fica com "—" em tudo que não
seja classificação, e some quando se escolhe um motivo.

Diferente da tabela do app e do slide, entram os defeitos dos dois meses, e não
só os que reaparecem no mais recente: sem os que saíram de cena o mês anterior
não fecharia em 100%.

Os números saem de `pipeline.calcular_relatorio`, o mesmo que alimenta o app e o
PowerPoint, para as três saídas não divergirem.
"""

import contextlib
import io
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.pivot.cache import (CacheDefinition, CacheField, CacheSource,
                                  SharedItems, WorksheetSource)
from openpyxl.pivot.fields import Index, Number, Text
from openpyxl.pivot.record import Record, RecordList
from openpyxl.pivot.table import (DataField, FieldItem, Location, PageField,
                                  PivotField, RowColField, TableDefinition)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import fallout_core

VERMELHO = fallout_core.COR_VERMELHO.lstrip("#")
BRANCO = "FFFFFF"
BORDA_CINZA = "BFBFBF"

SEM_INFO = "(não informado)"
# Linha do fallout que não caiu em nenhuma categoria da distribuição. Não
# deve acontecer; existe para o total fechar em vez de sumir gente no meio.
SEM_CATEGORIA = "(sem categoria)"

# Marca de "não se aplica" nas colunas descritivas. A dinâmica agrupa por texto,
# e item em branco vira uma linha "(vazio)" que não diz nada a quem lê.
TRACO = "—"

# 1º nível da tabela. "Falha Pontual" sai de dentro de Falha por ser o
# agrupamento que não se cobra de ninguém — é a leitura da planilha de
# acompanhamento, onde ela aparece separada das demais falhas.
CLASSE_SUCESSO = "Sucesso"
CLASSE_FALHA = "Falha"
CLASSE_PONTUAL = "Falha Pontual"          # também é nome de agrupamento
ORDEM_CLASSE = {CLASSE_SUCESSO: 0, CLASSE_FALHA: 1, CLASSE_PONTUAL: 2}

ABA_DADOS = "Dados"
ABA_DINAMICA = "Dinâmica"

_LADO = Side(style="thin", color=BORDA_CINZA)
BORDA = Border(left=_LADO, right=_LADO, top=_LADO, bottom=_LADO)

# ── Aba `Dados` ──────────────────────────────────────────────────────────────
# A ordem das colunas é a ordem dos campos no cache da dinâmica: os índices em
# `PIV_*` apontam para cá. Coluna nova entra no fim, nunca no meio.
CAB_DADOS = ["Jornada", "Classificação", "Agrupamento", "Defeito / US",
             "Data de previsão", "Time", "Tipo", "Status Defeito/US",
             "Nome Defeito", "Erro Tratado", "Motivo (Reason)", "Mês",
             "Pedidos"]
LARG_DADOS = [18, 14, 34, 14, 15, 24, 10, 20, 60, 60, 34, 9, 9]
NUM_DADOS = frozenset({12})                # só "Pedidos" é número
IDX_PEDIDOS = 12
COL_PEDIDOS = "Pedidos"
COLS_CHAVE = [c for c in CAB_DADOS if c != COL_PEDIDOS]

# ── Tabela dinâmica ──────────────────────────────────────────────────────────
# Índices em `CAB_DADOS` dos campos que entram no leiaute inicial.
PIV_PAGINA = [0, 10]                       # Jornada, Motivo (Reason)
# A ordem aqui é a ordem das colunas na tela, e não precisa ser a de
# `CAB_DADOS`: o status vem logo depois do número do defeito, que é como se lê
# ("233585, Aberto, previsão tal, time tal").
PIV_LINHAS = [1, 2, 3, 7, 4, 5, 6]
PIV_COLUNA = 11                            # Mês
PIV_VALOR = 12                             # Pedidos
# Índice reservado do OOXML para o eixo de valores: com dois campos de valor
# (volume e %), é ele que diz onde a dupla se repete — aqui, dentro de cada mês.
PIV_EIXO_VALORES = -2
# Formatos embutidos do Excel: 3 = "#,##0", 10 = "0.00%".
FMT_VOLUME, FMT_PCT = 3, 10
# A dinâmica começa aqui. Acima ficam o título e os avisos (linhas 1 a 3) e,
# logo antes dela, os dois campos de filtro, que o Excel desenha nas duas linhas
# acima de `ref` mais uma em branco.
LIN_DINAMICA = 7


def _rotulo_mes(chave):
    """202608 -> "ago/26" — o formato da planilha de acompanhamento."""
    return fallout_core.MESES_LABEL[chave].replace("-", "/")


def _texto(valor):
    """Descritiva pronta para a dinâmica: vazio e NaN viram "—"."""
    if pd.isna(valor):
        return TRACO
    return str(valor).strip() or TRACO


def _por_valor(fn):
    """
    `fn` memorizada pelos argumentos, para formatar cada valor distinto uma vez.

    Argumento que não serve de chave (não hasheável) é calculado direto; NaN que
    não casa com outro NaN só perde o atalho — o resultado é sempre o de `fn`.
    """
    prontos = {}

    def chamar(*args):
        try:
            return prontos[args]
        except KeyError:
            valor = prontos[args] = fn(*args)
            return valor
        except TypeError:
            return fn(*args)
    return chamar


# ── Dados da tabela ──────────────────────────────────────────────────────────
def dados_ofensores(ns, jornada):
    """
    Uma linha por (classificação, agrupamento, defeito, erro, motivo, mês) com a
    contagem de pedidos, mais uma linha de sucesso por mês.

    O **agrupamento** vem de `fallout_core.categorizar` e vale para toda linha de
    falha, tenha ela defeito ou não — é o que põe cada defeito debaixo de "Em
    Tratamento/Avaliação pela Squad", "Resolvido" e afins. `categorizar` roda por
    mês e as categorias dela podem se sobrepor numa linha ou outra; vale a
    primeira em que a linha aparece, que é a ordem de precedência do próprio
    dicionário. Sem esse desempate o mesmo pedido entraria em dois agrupamentos e
    a soma passaria do total de falhas do mês.

    A ordem das linhas é a ordem em que a dinâmica exibe os itens (o Excel numera
    cada campo pela primeira aparição na origem): classificação, agrupamento na
    ordem canônica de `categorizar`, e dentro dele o defeito do maior para o
    menor no mês mais recente.
    """
    df, meses_tab, mes_atual = ns["df"], ns["meses_tab"], ns["mes_atual"]
    fmt_id, fmt_ms, hoje = ns["fmt_id"], ns["fmt_milestone"], ns["hoje"]

    base = df[df["Mes"].isin(meses_tab)].copy()
    if base.empty:
        return pd.DataFrame(columns=CAB_DADOS)

    for col in ("ErrorHandled__c", "Reason"):
        if col not in base.columns:
            base[col] = pd.NA
        base[col] = (base[col].fillna("").astype(str).str.strip()
                     .replace({"": SEM_INFO, "nan": SEM_INFO, "None": SEM_INFO}))

    categoria, ordem_agrup = {}, []
    for m in meses_tab:
        _dfm, cats, _t = fallout_core.categorizar(df, m, hoje=hoje)
        for nome, sub in cats.items():
            if nome not in ordem_agrup:
                ordem_agrup.append(nome)
            for i in sub.index:
                categoria.setdefault(i, nome)

    agrup = [categoria.get(i, SEM_CATEGORIA) for i in base.index]
    eh_defeito = base["TemDefeito"].fillna(False).astype(bool)

    # O mesmo defeito se repete em centenas de pedidos, e `fmt_id`/`fmt_ms`
    # convertem a data com `pd.to_datetime` a cada chamada: linha a linha eram
    # ~40 mil conversões por planilha. Formatar cada valor uma vez só dá o mesmo
    # texto em uma fração do tempo.
    fmt_id_valor = _por_valor(lambda k, ms: fmt_id({"DefectKey": k, "DFT_BugfixMilestone": ms}))
    fmt_ms_valor = _por_valor(lambda ms: _texto(fmt_ms(ms)))

    base["Jornada"] = jornada
    base["Agrupamento"] = agrup
    base["Classificação"] = [CLASSE_PONTUAL if a == CLASSE_PONTUAL else CLASSE_FALHA
                             for a in agrup]
    # Sem defeito não há o que descrever: a linha é um pedaço do fallout que
    # ainda não foi endereçado a ninguém, e o agrupamento já diz qual.
    base["Defeito / US"] = [
        fmt_id_valor(k, ms) if e else TRACO
        for k, ms, e in zip(base["DefectKey"], base["DFT_BugfixMilestone"], eh_defeito)]
    base["Data de previsão"] = [
        (fmt_ms_valor(ms) if e else TRACO)
        for ms, e in zip(base["DFT_BugfixMilestone"], eh_defeito)]
    base["Time"] = [_texto(v) if e else TRACO
                    for v, e in zip(base["DFT_Team"], eh_defeito)]
    base["Tipo"] = [(("US" if str(t).strip() == "User Story" else "Defeito") if e
                     else TRACO) for t, e in zip(base["DFT_Type"], eh_defeito)]
    base["Status Defeito/US"] = [_texto(v) if e else TRACO
                                 for v, e in zip(base["DFT_Phase"], eh_defeito)]
    base["Nome Defeito"] = [_texto(v) if e else TRACO
                            for v, e in zip(base["DFT_Name"], eh_defeito)]
    base["Erro Tratado"] = base["ErrorHandled__c"]
    base["Motivo (Reason)"] = base["Reason"]
    base["Mês"] = [_rotulo_mes(m) for m in base["Mes"]]

    longo = base.groupby(COLS_CHAVE, dropna=False).size().reset_index(name=COL_PEDIDOS)
    # O mês entra na ordenação pela chave numérica, e não pelo rótulo: em ordem
    # alfabética "ago/26" viria antes de "jul/26", e as colunas da dinâmica
    # sairiam trocadas — o Excel numera o item do campo pela primeira aparição.
    ordem_mes = {_rotulo_mes(m): i for i, m in enumerate(sorted(meses_tab))}
    return _ordenar(pd.concat([longo, _sucessos(ns, jornada)], ignore_index=True),
                    ordem_agrup, ordem_mes, _rotulo_mes(mes_atual))


def _sucessos(ns, jornada):
    """
    Uma linha por mês com os pedidos que não falharam.

    Sem ela a coluna do mês fecharia em 100% do fallout, e não em 100% dos
    pedidos — que é a leitura da planilha de acompanhamento, onde o sucesso é a
    primeira linha da tabela.
    """
    resumo = ns["resumo"]
    linhas = []
    for m in ns["meses_tab"]:
        if m not in resumo.index:
            continue
        n = resumo.loc[m, "Sucessos"]
        if pd.isna(n) or n <= 0:
            continue
        reg = dict.fromkeys(COLS_CHAVE, TRACO)
        reg.update({"Jornada": jornada, "Classificação": CLASSE_SUCESSO,
                    "Mês": _rotulo_mes(m), COL_PEDIDOS: int(n)})
        linhas.append(reg)
    return pd.DataFrame(linhas, columns=CAB_DADOS)


def _ordenar(longo, ordem_agrup, ordem_mes, rotulo_atual):
    """Ordem de exibição da dinâmica (ver `dados_ofensores`)."""
    if not len(longo):
        return longo[CAB_DADOS]

    pos = {nome: i for i, nome in enumerate(ordem_agrup)}
    peso = (longo[longo["Mês"] == rotulo_atual]
            .groupby(["Agrupamento", "Defeito / US"])[COL_PEDIDOS].sum().to_dict())
    longo = longo.assign(
        _classe=[ORDEM_CLASSE.get(c, 9) for c in longo["Classificação"]],
        _agrup=[pos.get(a, 99) for a in longo["Agrupamento"]],
        _peso=[-peso.get((a, d), 0) for a, d in
               zip(longo["Agrupamento"], longo["Defeito / US"])],
        _mes=[ordem_mes.get(m, 99) for m in longo["Mês"]])
    return (longo.sort_values(["_classe", "_agrup", "_peso", "Defeito / US", "_mes"])
            [CAB_DADOS])


def _canonizar(linhas):
    """
    Uma só grafia por valor em cada campo, ignorando a caixa.

    O Excel compara item de tabela dinâmica **sem diferenciar maiúscula de
    minúscula**. Ter "Em Tratamento" e "Em tratamento" como dois itens do mesmo
    campo deixa o cache inválido, e ele recusa o arquivo inteiro — a caixa de
    "conteúdo ilegível", que descarta a dinâmica. O Octane devolve as duas
    grafias em `DFT_Phase`, e o mesmo acontece em `ErrorHandled__c` ("Erro
    Submeter" / "Erro submeter").

    O WPS não faz essa checagem e abre a planilha assim mesmo, o que faz o
    problema passar despercebido em quem gera. Por isso a normalização é feita
    aqui, na origem, e não confiando no leitor.

    Vence a grafia com mais pedidos; no empate, a primeira que apareceu. As
    linhas que ficam idênticas depois da troca são somadas, senão a aba `Dados`
    mostraria duas linhas iguais lado a lado.
    """
    grafias = [defaultdict(lambda: defaultdict(int)) for _ in CAB_DADOS]
    for r in linhas:
        for i, v in enumerate(r):
            if i not in NUM_DADOS:
                grafias[i][v.casefold()][v] += r[IDX_PEDIDOS]

    canone = [{} for _ in CAB_DADOS]
    for i, grupos in enumerate(grafias):
        for variantes in grupos.values():
            vencedora = max(variantes, key=variantes.get)
            for v in variantes:
                canone[i][v] = vencedora

    juntas = {}
    for r in linhas:
        nova = [v if i in NUM_DADOS else canone[i][v] for i, v in enumerate(r)]
        chave = tuple(v for i, v in enumerate(nova) if i not in NUM_DADOS)
        if chave in juntas:
            juntas[chave][IDX_PEDIDOS] += nova[IDX_PEDIDOS]
        else:
            juntas[chave] = nova
    return list(juntas.values())


# ── Escrita ──────────────────────────────────────────────────────────────────
def _cabecalho(ws, linha, textos):
    """Faixa vermelha de cabeçalho, com o texto branco em negrito."""
    for col, texto in enumerate(textos, start=1):
        c = ws.cell(row=linha, column=col, value=texto)
        c.fill = PatternFill("solid", fgColor=VERMELHO)
        c.font = Font(bold=True, color=BRANCO, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BORDA


def _escrever_dados(wb, blocos):
    """
    Aba `Dados`: a origem da dinâmica, uma linha por combinação das colunas.

    Não dá para tirá-la: a dinâmica pede uma atualização ao abrir, e atualizar
    significa reler este intervalo. Fica visível porque uma origem escondida só
    aparece quando alguém clica em "Atualizar" e leva um erro sem entender de
    onde veio.

    Devolve a lista de linhas já normalizadas (texto sem vazio, número como int),
    que é o que alimenta o cache da dinâmica — reaproveitar daqui garante que a
    dinâmica e a aba mostrem exatamente o mesmo.
    """
    linhas = []
    for bloco in blocos:
        # `itertuples` e não `iterrows`: este monta uma Series por linha.
        for r in bloco[CAB_DADOS].itertuples(index=False, name=None):
            linhas.append([int(v) if i in NUM_DADOS else _texto(v)
                           for i, v in enumerate(r)])
    valores = _canonizar(linhas)

    ws = wb.create_sheet(title=ABA_DADOS)
    _cabecalho(ws, 1, CAB_DADOS)
    linha = 2
    for registro in valores:
        for i, valor in enumerate(registro, start=1):
            ws.cell(row=linha, column=i, value=valor)
        linha += 1

    for i, larg in enumerate(LARG_DADOS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = (f"A1:{get_column_letter(len(CAB_DADOS))}"
                          f"{max(linha - 1, 1)}")
    return valores


# ── Tabela dinâmica ──────────────────────────────────────────────────────────
def _dominios(valores):
    """
    Valores distintos de cada coluna de texto de `Dados`, na ordem de primeira
    aparição, como {coluna: {valor: índice}}.

    A ordem importa: é ela que numera os itens no cache, e o Excel exibe o campo
    nessa numeração. Como `_ordenar` já pôs as linhas na ordem de leitura, os
    itens saem na ordem certa sem precisar de regra de classificação na dinâmica.
    """
    dom = {}
    for i in range(len(CAB_DADOS)):
        if i in NUM_DADOS:
            continue
        vistos = {}
        for registro in valores:
            vistos.setdefault(registro[i], len(vistos))
        dom[i] = vistos
    return dom


def _cache_dinamica(valores, dom, n_linhas):
    """Cache da dinâmica: os campos e uma cópia das linhas de `Dados`."""
    campos = []
    for i, nome in enumerate(CAB_DADOS):
        if i in NUM_DADOS:
            nums = [r[i] for r in valores]
            campos.append(CacheField(name=nome, sharedItems=SharedItems(
                containsSemiMixedTypes=False, containsString=False,
                containsNumber=True, containsInteger=True,
                minValue=min(nums), maxValue=max(nums))))
        else:
            itens = list(dom[i])
            # `longText` é obrigatório quando algum item passa de 255
            # caracteres: sem ele o Excel recusa o cache inteiro e descarta a
            # dinâmica. Acontece em "Erro Tratado", onde a mensagem do OM chega
            # a 380 caracteres. O WPS não faz essa checagem, então a planilha
            # abre normalmente lá mesmo estando errada.
            campos.append(CacheField(name=nome, sharedItems=SharedItems(
                _fields=[Text(v=v) for v in itens], count=len(itens),
                longText=True if any(len(v) > 255 for v in itens) else None)))

    registros = RecordList(r=[
        Record(_fields=[Number(v=r[i]) if i in NUM_DADOS else Index(v=dom[i][r[i]])
                        for i in range(len(CAB_DADOS))])
        for r in valores])

    # `refreshOnLoad` é o que faz a dinâmica abrir já desenhada: o openpyxl grava
    # a definição, mas não o resultado célula a célula, e sem o pedido de
    # atualização a área da dinâmica apareceria vazia até alguém clicar em
    # "Atualizar".
    cache = CacheDefinition(
        refreshOnLoad=True, refreshedBy="Automatização MAYH",
        recordCount=len(valores), createdVersion=8, refreshedVersion=8,
        minRefreshableVersion=3,
        cacheSource=CacheSource(
            type="worksheet",
            worksheetSource=WorksheetSource(
                ref=f"A1:{get_column_letter(len(CAB_DADOS))}{n_linhas}",
                sheet=ABA_DADOS)),
        cacheFields=campos)
    cache.records = registros
    return cache


def _escrever_dinamica(wb, valores, corte):
    """
    Aba `Dinâmica`: a distribuição do fallout, com jornada e motivo como filtros.

    Jornada é filtro, e não coluna, porque "Consolidado" e "Base + Cross Sell"
    já somam outras jornadas: com duas selecionadas o total contaria o mesmo
    pedido duas vezes — e a participação deixaria de fechar em 100%.

    Sem subtotal em nenhum nível de linha: a tabela é uma lista, como a do
    acompanhamento, e o rótulo do agrupamento já aparece na primeira linha do
    grupo. Quem quiser o total por agrupamento liga em duas: botão direito no
    campo, "Subtotal de …".
    """
    ws = wb.create_sheet(title=ABA_DINAMICA, index=0)
    if not valores:
        return

    dom = _dominios(valores)
    cache = _cache_dinamica(valores, dom, len(valores) + 1)

    def _campo(i):
        if i == PIV_VALOR:
            return PivotField(dataField=True, showAll=False,
                              compact=False, outline=False)
        eixo = ("axisPage" if i in PIV_PAGINA else
                "axisRow" if i in PIV_LINHAS else
                "axisCol" if i == PIV_COLUNA else None)
        if eixo is None:
            return PivotField(showAll=False, compact=False, outline=False)
        subtotal = eixo != "axisRow"
        itens = [FieldItem(x=k) for k in range(len(dom[i]))]
        if subtotal:
            itens.append(FieldItem(t="default"))
        return PivotField(axis=eixo, showAll=False, compact=False, outline=False,
                          defaultSubtotal=subtotal, items=itens)

    # A jornada começa numa só, e não em "(Todos)", pela mesma razão da soma
    # dupla acima. "Consolidado" já é o retrato de todas juntas, então é ele
    # quando existe.
    jorn = dom[PIV_PAGINA[0]]
    inicial = "Consolidado" if "Consolidado" in jorn else next(iter(jorn))

    # Duas colunas de valor (volume e %) por mês.
    n_valor = 2 * len(dom[PIV_COLUNA])
    largura = get_column_letter(len(PIV_LINHAS) + n_valor)
    ultima = LIN_DINAMICA + 3 + len(valores)
    pivot = TableDefinition(
        name="OfensoresDinamica", cacheId=1, dataCaption="Valores",
        rowHeaderCaption="Ofensor", updatedVersion=8, createdVersion=8,
        minRefreshableVersion=3, useAutoFormatting=True, indent=0,
        compact=False, compactData=False, outline=False, outlineData=False,
        gridDropZones=False, itemPrintTitles=True, multipleFieldFilters=False,
        # Sem "Total Volume" / "Total % do mês" na ponta direita: somar dois
        # meses de volume não diz nada, e o % desse total é a participação nos
        # dois meses juntos, que não é leitura de ninguém. `colGrandTotals` é
        # que governa essas colunas; `rowGrandTotals` é a linha "Total geral" do
        # rodapé, que fica — é ela que fecha cada mês em 100%.
        rowGrandTotals=True, colGrandTotals=False,
        location=Location(ref=f"A{LIN_DINAMICA}:{largura}{ultima}",
                          firstHeaderRow=1, firstDataRow=2,
                          firstDataCol=len(PIV_LINHAS),
                          rowPageCount=len(PIV_PAGINA), colPageCount=1),
        pivotFields=[_campo(i) for i in range(len(CAB_DADOS))],
        rowFields=[RowColField(x=i) for i in PIV_LINHAS],
        # O eixo de valores depois do mês: cada mês abre em volume e %, em vez de
        # a planilha inteira abrir em dois blocos de meses.
        colFields=[RowColField(x=PIV_COLUNA),
                   RowColField(x=PIV_EIXO_VALORES)],
        colItems=[],
        pageFields=[PageField(fld=PIV_PAGINA[0], hier=-1, item=jorn[inicial]),
                    PageField(fld=PIV_PAGINA[1], hier=-1)],
        # "% do total da coluna" é a participação dentro do mês, que é o que
        # fecha em 100%. O nome não pode ser "Pedidos": ele já é o nome da coluna
        # de origem, e o Excel desempata sozinho pondo "Pedidos2".
        dataFields=[
            DataField(name="Volume", fld=PIV_VALOR, subtotal="sum",
                      baseField=0, baseItem=0, numFmtId=FMT_VOLUME),
            DataField(name="% do mês", fld=PIV_VALOR, subtotal="sum",
                      showDataAs="percentOfCol", baseField=0, baseItem=0,
                      numFmtId=FMT_PCT)])
    pivot.cache = cache
    ws._pivots.append(pivot)

    titulo = "Ofensores — distribuição do fallout"
    if corte is not None:
        titulo += f" · dados até {corte:%d/%m/%Y}"
    t = ws.cell(row=1, column=1, value=titulo)
    t.font = Font(bold=True, size=12)
    # Só três linhas de aviso cabem: o Excel desenha os dois campos de filtro
    # nas linhas 5 e 6, logo acima da tabela.
    for i, texto in enumerate((
            "«% do mês» é a participação no total de pedidos do mês, sucesso "
            "incluído — por isso cada mês fecha em 100%.",
            "Deixe «Jornada» numa só: Consolidado e Base + Cross Sell já somam "
            "outras jornadas, e com duas marcadas o mesmo pedido conta duas "
            "vezes. Consolidado cobre só as jornadas de extração — as demais "
            "medem o denominador de outro jeito e ficam fora dele.",
            "Escolher um «Motivo (Reason)» tira a linha de Sucesso: a extração "
            "de sucessos é um total por dia, sem motivo. Aí o % passa a ser "
            "dentro das falhas daquele motivo."), start=2):
        c = ws.cell(row=i, column=1, value=texto)
        c.font = Font(italic=True, size=9, color="808080")

    # A largura de cada campo de linha sai da mesma tabela da aba `Dados`, para
    # não haver duas listas de larguras que precisem ser mexidas juntas quando
    # `PIV_LINHAS` muda de ordem ou ganha um campo.
    for coluna, campo in enumerate(PIV_LINHAS, start=1):
        ws.column_dimensions[get_column_letter(coluna)].width = LARG_DADOS[campo]
    for i in range(len(PIV_LINHAS) + 1, len(PIV_LINHAS) + n_valor + 1):
        ws.column_dimensions[get_column_letter(i)].width = 12


# ── Montagem do arquivo ──────────────────────────────────────────────────────
def _jornada(jornada, base_dir):
    """Linhas de uma jornada, já descritas. E quantos defeitos/US ela tem."""
    import pipeline

    # Só os números: `calcular_relatorio` não monta a figura do slide, que aqui
    # seria desenhada para ser descartada. O diagnóstico que ele imprime no
    # stdout também não interessa.
    with contextlib.redirect_stdout(io.StringIO()):
        ns = pipeline.calcular_relatorio(base_dir, jornada)

    linhas = dados_ofensores(ns, jornada)
    n_def = len(set(linhas["Defeito / US"]) - {TRACO}) if len(linhas) else 0
    return linhas, n_def


def gerar(base_dir, jornadas, progresso=None):
    """
    Devolve (bytes do .xlsx, jornadas feitas, avisos).

    Uma jornada que falhar não derruba o arquivo: ela fica de fora e o motivo
    volta na lista de avisos.

    `progresso(feitos, total, rotulo)`, quando passado, é chamado antes de cada
    jornada e antes da escrita da planilha.
    """
    wb = Workbook()
    wb.remove(wb.active)

    total = len(jornadas) + 1
    feitas, avisos, blocos = [], [], []
    for i, j in enumerate(jornadas):
        if progresso:
            progresso(i, total, fallout_core.nome_exibicao(j))
        try:
            linhas, n_def = _jornada(j, base_dir)
        except Exception as e:                              # noqa: BLE001
            avisos.append(f"{j}: {type(e).__name__}: {e}")
            continue
        if len(linhas):
            blocos.append(linhas)
        feitas.append((j, n_def))

    if not feitas:
        raise RuntimeError("Nenhuma jornada pôde ser gerada. " + " | ".join(avisos))

    if progresso:
        progresso(len(jornadas), total, "Escrevendo a planilha")
    valores = _escrever_dados(wb, blocos)
    _escrever_dinamica(wb, valores, fallout_core.data_corte(base_dir))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), feitas, avisos
