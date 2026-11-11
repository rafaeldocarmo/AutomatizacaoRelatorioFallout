# -*- coding: utf-8 -*-
"""
Comparação entre duas fotografias de extração (ver fotografia.py): o que mudou
de uma semana para a outra na distribuição, na projeção, nos defeitos e nos
erros.

Como o fallout_core, não depende de Streamlit nem de estado global: recebe as
duas fotografias e devolve tabelas e texto. A aba "Comparação" do app e a
linha de comando só mostram o que sai daqui.

Toda variação é dividida em partes que somam exatamente o total, e cada
comparação confere isso (`Comparacao.verificacoes`). Uma explicação que não
fecha com o número que ela explica é pior do que nenhuma: ninguém percebe que
está errada.

Uso pela linha de comando:
    python comparacao.py                        a fotografia mais recente contra a de ~7 dias antes
    python comparacao.py --visao "Base Móvel"
    python comparacao.py --base 2026-09-10 --atual 2026-09-16
"""
import sys

# No Windows o console usa cp1252 e quebra nos prints com acento e "→".
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import dataclasses
import datetime as _dt
import html
import re

import pandas as pd

import fallout_core
import fotografia

FUSO = "America/Sao_Paulo"
SEM_MENSAGEM = "(sem mensagem de erro)"
CAT_TRATAMENTO = fotografia.CAT_TRATAMENTO
CAT_FALTA_ASSOCIAR = "Falta Associar ao Defeito/US"
# A ordem da Distribuição — a do dicionário de `fallout_core.categorizar`.
# Categoria que aparecer nos dados e não estiver aqui entra no fim: nada some
# da comparação por esta lista ficar desatualizada.
ORDEM_CATEGORIAS = (
    CAT_TRATAMENTO, "Resolvido", "Falha Pontual", CAT_FALTA_ASSOCIAR,
    "Tratado - Em avaliação de eficácia", "Em Avaliação por MOPs",
    # A Funcionalidade idêntica ao legado foi categoria até 15/10 e hoje é
    # subcategoria do Legado Aberto; fica na lista pelas fotografias antigas.
    fallout_core.CAT_BSIM, fallout_core.CAT_LEGADO, fallout_core.CAT_ERRO_LEGADO,
    "Em avaliação - Outros times",
    "Falha no Processo Usuário", fallout_core.CAT_INCIDENTE,
    fallout_core.CAT_SCRIPT_RPA, fallout_core.CAT_AGENDAMENTO,
)
_VAZIOS = {"", "nan", "none", "<na>"}
# Diferença aceita nas conferências: as partes são somas de floats.
TOLERANCIA = 1e-6


@dataclasses.dataclass(frozen=True)
class Limites:
    """
    O que conta como alerta. Os padrões são os combinados em 18/09/2026: um
    erro é significativo a partir de 20 ocorrências, e a base da comparação é
    a fotografia de cerca de 7 dias antes.
    """
    # Erro novo, em alta ou que sumiu: ocorrências mínimas para entrar.
    min_ocorrencias: int = 20
    # "Em alta": a média diária do período novo chega a este múltiplo da média
    # das semanas anteriores.
    fator_alta: float = 2.0
    # A janela de referência do "em alta" e do "sumiu": os dias que terminam na
    # data de corte da fotografia anterior. Quatro semanas cheias, para os fins
    # de semana pesarem nela o mesmo que no período novo.
    dias_referencia: int = 28
    # Distância, em dias, da fotografia base padrão (ver `escolher_base`).
    dias_base: int = 7
    # Pedidos criados no período novo depois da data de entrega do defeito
    # para ele virar alerta. Um basta: depois da entrega o erro não deveria
    # voltar.
    min_pos_entrega: int = 1
    # Queda de falhas num mês que as duas fotografias têm, como fração do que
    # havia antes, a partir da qual a extração nova é apontada como suspeita.
    queda_suspeita: float = 0.05


@dataclasses.dataclass
class Comparacao:
    """Tudo o que muda entre duas fotografias, para uma visão."""
    visao: str
    base_id: str
    atual_id: str
    mes: int | None
    contexto: dict
    distribuicao: pd.DataFrame
    migracao: pd.DataFrame
    pedidos_migrados: pd.DataFrame
    projecao: dict
    defeitos: pd.DataFrame
    erros_novos: pd.DataFrame
    erros_em_alta: pd.DataFrame
    erros_sumidos: pd.DataFrame
    texto: list
    verificacoes: list
    limites: "Limites" = None


# ── Formatação ────────────────────────────────────────────────────────────────
def _pp(valor):
    """+0.25 -> '+0,25 pp'."""
    if valor is None or valor != valor:
        return "—"
    return f"{valor:+.2f} pp".replace(".", ",")


def _pct(valor):
    return fallout_core.fmt_pct(valor) or "—"


def _int(valor):
    return fallout_core.fmt_inteiro(valor) or "0"


def _dia(data):
    """date -> '12/09'; ausente -> '—'."""
    return "—" if data is None or pd.isna(data) else f"{data:%d/%m}"


def _texto(valor):
    """Texto de uma célula; ausente vira ''."""
    return "" if valor is None or pd.isna(valor) else str(valor).strip()


def _curto(texto, limite=80):
    """Corta no último espaço antes do limite, se ele não estiver longe demais."""
    texto = " ".join(str(texto).split())
    if len(texto) <= limite:
        return texto
    corte = texto[:limite - 1]
    espaco = corte.rfind(" ")
    if espaco >= limite - 20:
        corte = corte[:espaco]
    return corte.rstrip(" ,;:-–·") + "…"


def _n(quantos, singular, plural):
    """(1, 'pedido novo', 'pedidos novos') -> '1 pedido novo'; 1234 -> '1.234 pedidos novos'."""
    return f"{_int(quantos)} {singular if int(quantos) == 1 else plural}"


def _lista(itens):
    """['a', 'b', 'c'] -> 'a, b e c'."""
    itens = list(itens)
    return itens[0] if len(itens) == 1 else ", ".join(itens[:-1]) + " e " + itens[-1]


# Prefixos que se repetem no nome dos defeitos e não distinguem um do outro:
# "[JORNADA - BASE] ERRO NO OM – ETAPA …", "Cópia de …", "SOLAR - …".
_RE_PREFIXO_NOME = re.compile(
    r"^(?:\s*(?:c[oó]pia de|copy of)\s+|\s*\[[^\]]*\]\s*[-–]?|\s*solar\s*[-–]"
    r"|\s*erro no om\s*[-–]|\s*etapa\s+)+",
    re.IGNORECASE)


def nome_curto(nome, limite=70):
    """O nome do defeito sem os prefixos repetidos, para caber numa linha."""
    texto = " ".join(_texto(nome).split())
    return _curto(_RE_PREFIXO_NOME.sub("", texto).strip(" -–") or texto, limite)


# As mensagens das APIs vêm num envelope fixo — "message: … | link: … |
# detailed message: … | error code: … | http code: …" — e o que distingue uma
# da outra está no "detailed message". Cortada pelo começo, toda mensagem vira
# o mesmo "message: unprocessable entity | link: https://…".
_RE_ENVELOPE = re.compile(
    r"^message:\s*(?P<cabeca>.*?)\s*\|\s*link:\s*\S+\s*\|\s*detailed message:\s*(?P<detalhe>.*)$",
    re.IGNORECASE | re.DOTALL)
_RE_RABO = re.compile(r"\|\s*error code:.*$", re.IGNORECASE | re.DOTALL)


def resumir_mensagem(mensagem, limite=120):
    """
    A parte da mensagem de erro que a identifica, para caber numa linha:
    o "detailed message" do envelope das APIs, sem o código HTTP no fim.
    Só para exibir — as tabelas guardam a mensagem inteira.
    """
    texto = html.unescape(" ".join(str(mensagem).split()))
    achado = _RE_ENVELOPE.match(texto)
    if achado:
        detalhe = _RE_RABO.sub("", achado.group("detalhe")).replace("|error;", "")
        detalhe = re.sub(r"\s*\|\s*", " · ", detalhe.strip(" |;"))
        texto = (detalhe if detalhe and detalhe.lower() != "null"
                 else f"message: {achado.group('cabeca')} (sem detalhe)")
    return _curto(texto, limite)


def _rotulo(mes):
    return fallout_core.MESES_LABEL[int(mes)]


def _chave_do_rotulo(rotulo):
    """'out-26' -> 202610, o inverso de `fallout_core.MESES_LABEL`."""
    abrev, ano = rotulo.split("-")
    return fallout_core.chave_mes(2000 + int(ano), fallout_core.ABREV_MES[abrev.lower()])


def _defeito(chave):
    """'232882' -> 'DFT 232882'; rótulo (BRJ-218, INC-…) fica como veio."""
    chave = _texto(chave)
    return f"DFT {chave}" if chave.isdigit() else chave


def _verif(nome, ok, detalhe=""):
    return {"Verificação": nome, "OK": bool(ok), "Detalhe": detalhe}


# ── Uma visão dentro de uma fotografia ────────────────────────────────────────
@dataclasses.dataclass
class _Lado:
    """As tabelas de uma visão numa fotografia, prontas para comparar."""
    foto: fotografia.Fotografia
    visao: str
    pedidos: pd.DataFrame
    resumo: pd.DataFrame          # indexado pelo mês
    reducao: pd.DataFrame
    ofensores: pd.DataFrame
    rel: dict

    @property
    def corte(self):
        return _dt.date.fromisoformat(self.foto.meta["corte"])

    @property
    def hoje(self):
        return fallout_core.hoje_utc(self.foto.meta["hoje"])

    def total(self, mes):
        """
        Pedidos do mês, sucesso incluído — o denominador da distribuição —, ou
        None. No mês fechado digitado só com a taxa, o total da extração (ver
        `fallout_core.total_do_mes`): sem ele, o set/26 do Base + Cross Sell
        fechado em 3,24% sumia da comparação, que passava a mostrar ago/26.
        Fotografia antiga, gravada antes dessa coluna, segue sem o total.
        """
        if mes is None or mes not in self.resumo.index:
            return None
        valor = fallout_core.total_do_mes(self.resumo, mes)
        return None if pd.isna(valor) else float(valor)

    def taxa(self, mes):
        """Fallout do mês, como o relatório mostra, ou None."""
        if mes is None or mes not in self.resumo.index:
            return None
        valor = self.resumo.loc[mes, "Pct"]
        return None if pd.isna(valor) else float(valor)

    def fonte(self, mes):
        return self.resumo.loc[mes, "Fonte"] if mes in self.resumo.index else None


def _lado(foto, visao):
    """A visão `visao` da fotografia, ou None quando ela não foi fotografada."""
    if visao not in foto.meta["relatorios"]:
        return None
    pedidos = foto.pedidos[foto.pedidos["Visao"] == visao].copy()
    pedidos["Dia"] = pedidos["CreatedDate"].dt.tz_convert(FUSO).dt.date
    # O pedido sem número não casa com nada: recebe uma chave só dele, que não
    # se repete na outra fotografia. Casar "vazio" com "vazio" juntaria pedidos
    # diferentes numa migração que não existiu.
    numero = pedidos["OrderNumber"].astype("string").str.strip()
    sem_numero = numero.fillna("").str.lower().isin(_VAZIOS)
    substituta = pd.Series([f"sem número #{i}" for i in range(len(pedidos))],
                           index=pedidos.index, dtype="string")
    pedidos["Chave"] = numero.where(~sem_numero, substituta)
    # A mensagem é a chave da análise de erros, como na aba do app; vazio e
    # ausente viram a mesma coisa.
    erro = pedidos["ErrorHandled__c"].astype("string").str.strip()
    pedidos["Erro"] = erro.where(~erro.fillna("").eq(""), SEM_MENSAGEM)
    pedidos["Categoria"] = pedidos["Categoria"].fillna("(sem categoria)")

    resumo = foto.resumo[foto.resumo["Visao"] == visao].copy()
    resumo["Mes"] = resumo["Mes"].astype(int)
    return _Lado(
        foto=foto, visao=visao, pedidos=pedidos,
        resumo=resumo.set_index("Mes"),
        reducao=foto.reducao[foto.reducao["Visao"] == visao],
        ofensores=foto.ofensores[foto.ofensores["Visao"] == visao],
        rel=foto.meta["relatorios"][visao],
    )


def _atributos(lado):
    """Nome, fase, data de entrega, time e tipo de cada defeito da fotografia."""
    colunas = ["DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Team", "DFT_Type"]
    com_defeito = lado.pedidos[lado.pedidos["TemDefeito"]]
    if not len(com_defeito):
        return pd.DataFrame(columns=colunas)
    return com_defeito.groupby("DefectKey")[colunas].first()


def _entrega(valor):
    """A data de entrega como o `categorizar` a enxerga: instante em UTC."""
    ts = pd.to_datetime(valor, utc=True, errors="coerce")
    return None if pd.isna(ts) else ts


def escolher_base(fotos, id_atual, dias=Limites.dias_base):
    """
    A fotografia padrão para comparar com `id_atual`: entre as anteriores, a de
    data de corte mais próxima de `dias` antes — no empate, a mais recente.

    `fotos` é o que `fotografia.listar` devolve. Devolve o id, ou None quando
    não há fotografia anterior.
    """
    atual = _dt.date.fromisoformat(id_atual)
    anteriores = [f for f in fotos if f["corte"] < atual]
    if not anteriores:
        return None
    alvo = atual - _dt.timedelta(days=dias)
    return min(anteriores,
               key=lambda f: (abs((f["corte"] - alvo).days), -f["corte"].toordinal()))["id"]


# ── Fontes e qualidade dos dados ──────────────────────────────────────────────
def comparar_fontes(base, atual):
    """
    Os arquivos de entrada das duas fotografias, lado a lado, com a situação de
    cada um: novo, removido, alterado ou igual.

    "Alterado" é pelo tamanho. A data de modificação não sobrevive a uma cópia
    da pasta (o Explorer grava a hora da cópia em todo arquivo), e com ela um
    arquivo idêntico, só copiado de novo, aparecia como alterado — a jornada
    passava por "atualizada" e o aviso de Octane igual nunca saía. Ela segue na
    tabela, só para consulta.
    """
    colunas = ["arquivo", "modificado", "tamanho"]
    antes = pd.DataFrame(base.meta["fontes"], columns=colunas)
    agora = pd.DataFrame(atual.meta["fontes"], columns=colunas)
    tab = antes.merge(agora, on="arquivo", how="outer", suffixes=(" antes", " agora"),
                      indicator=True)
    situacao = pd.Series("igual", index=tab.index)
    situacao[tab["_merge"] == "right_only"] = "novo"
    situacao[tab["_merge"] == "left_only"] = "removido"
    mudou = (tab["_merge"] == "both") & (tab["tamanho antes"] != tab["tamanho agora"])
    situacao[mudou] = "alterado"
    tab["Situação"] = situacao
    # Os arquivos da raiz (Octane, Jira, histórico) não têm pasta de jornada.
    tab["Jornada"] = [c.split("/")[0] if "/" in c and not c.startswith("..") else "(raiz)"
                      for c in tab["arquivo"]]
    return (tab[["Jornada", "arquivo", "Situação", "modificado antes", "modificado agora",
                 "tamanho antes", "tamanho agora"]]
            .rename(columns={"arquivo": "Arquivo"})
            .sort_values(["Jornada", "Arquivo"]).reset_index(drop=True))


# Queda a partir da qual um mês que perdeu falhas deixa de ser "vale conferir"
# e passa a ser "provável problema no arquivo": 30% do que havia.
QUEDA_FORTE = 0.30


def _qualidade(a, b, limites):
    """
    Falhas por jornada e mês nas duas fotografias, e os meses que perderam
    falhas de uma para a outra.

    Perder pedido num mês que já existia acontece: a extração tira o pedido
    reclassificado em outra jornada, e um export novo pode redistribuir
    pedidos entre meses (no Prospect, ago-26 perdeu 241 e jul-26 ganhou 102 de
    uma semana para a outra). Mas uma queda forte quase sempre é o arquivo — foi
    assim que o export da PME de 17/09 apareceu, sem o número do pedido em
    13.847 falhas. Por isso a tabela separa as duas coisas.
    """
    antes = a.pedidos.groupby(["Jornada", "Mes"]).size()
    agora = b.pedidos.groupby(["Jornada", "Mes"]).size()
    chaves = sorted(set(antes.index) | set(agora.index))
    tab = pd.DataFrame({
        "Jornada": [j for j, _m in chaves],
        "Mes": [int(m) for _j, m in chaves],
        "Falhas antes": [int(antes.get(k, 0)) for k in chaves],
        "Falhas agora": [int(agora.get(k, 0)) for k in chaves],
    })
    tab["Diferença"] = tab["Falhas agora"] - tab["Falhas antes"]
    tab["Suspeito"] = ((-tab["Diferença"] >= limites.min_ocorrencias)
                       & (tab["Falhas agora"] < tab["Falhas antes"] * (1 - limites.queda_suspeita)))
    tab["Queda forte"] = tab["Suspeito"] & (tab["Falhas agora"]
                                            < tab["Falhas antes"] * (1 - QUEDA_FORTE))
    tab.insert(1, "Mês", [_rotulo(m) for m in tab["Mes"]])
    return tab


def _avisos_qualidade(qualidade):
    """Um aviso por jornada com mês que perdeu falhas."""
    avisos = []
    suspeitos = qualidade[qualidade["Suspeito"]]
    for jornada, linhas in suspeitos.groupby("Jornada", sort=False):
        meses = "; ".join(
            f"{l['Mês']} ({_int(l['Falhas antes'])} → {_int(l['Falhas agora'])})"
            for _, l in linhas.iterrows())
        forte = linhas["Queda forte"].any()
        avisos.append(
            f"{fallout_core.nome_exibicao(jornada)}: a extração nova tem menos falhas em "
            f"{'meses' if len(linhas) > 1 else 'mês'} que já existia"
            f"{'m' if len(linhas) > 1 else ''} — {meses}. "
            + ("Queda forte: provável problema no arquivo novo, e não melhora."
               if forte else
               "Pode ser pedido reclassificado na origem; vale conferir antes de ler "
               "a variação desse mês."))
    return avisos


def _contexto(base, atual, a, b, mes, limites):
    """
    Datas, fontes, qualidade e o que vale saber para ler a comparação: os
    avisos por extenso, e as ressalvas — os mesmos avisos em poucas palavras,
    para a linha única do fim do texto.
    """
    nome = fallout_core.nome_exibicao(b.visao)
    fontes = comparar_fontes(base, atual)
    avisos, ressalvas = [], []

    reconstruidas = []
    for foto, quando in ((base, "anterior"), (atual, "atual")):
        if foto.meta["origem"].startswith("reconstruída"):
            avisos.append(f"A fotografia {quando} foi {foto.meta['origem']}.")
            reconstruidas.append(quando)
        # Os avisos gravados na fotografia são os da reconstrução: a ressalva
        # abaixo já os resume.
        avisos += [f"Fotografia {quando}: {aviso}" for aviso in foto.meta["avisos"]]
    if reconstruidas:
        ressalvas.append("as duas fotografias foram reconstruídas com arquivos antigos"
                         if len(reconstruidas) == 2 else
                         f"a fotografia {reconstruidas[0]} foi reconstruída com arquivos antigos")

    if base.meta["regras"] != atual.meta["regras"]:
        avisos.append("O código das regras mudou entre as duas fotografias: parte da "
                      "diferença pode ser da regra, e não dos dados.")
        ressalvas.append("as regras mudaram entre as duas fotografias")

    octane = fontes[(fontes["Jornada"] == "(raiz)")
                    & fontes["Arquivo"].str.endswith((fallout_core.ARQ_DFT, fallout_core.ARQ_US))]
    if len(octane) and (octane["Situação"] == "igual").all():
        # Sem a data do arquivo: é a da última cópia da pasta, e não a do export.
        avisos.append("O Octane não mudou entre as duas fotografias: mudanças de fase e "
                      "de data dos defeitos não aparecem nesta comparação.")
        ressalvas.append("o Octane é o mesmo nas duas, então fase e data dos defeitos não mudam")

    jornadas = sorted(set(a.rel["jornadas"]) | set(b.rel["jornadas"]))
    situacao = []
    for j in jornadas:
        linhas = fontes[fontes["Jornada"] == j]
        mudou = linhas["Situação"].isin(["novo", "alterado", "removido"])
        situacao.append({
            "Jornada": fallout_core.nome_exibicao(j),
            "Situação": "atualizada" if mudou.any() else "sem arquivo novo",
            "Arquivos novos ou alterados": int(mudou.sum()),
        })
        if not mudou.any():
            avisos.append(f"{fallout_core.nome_exibicao(j)} não recebeu arquivo novo entre "
                          f"as duas fotografias: o que não mudou pode ser só falta de "
                          f"atualização.")
    sem_arquivo = [s["Jornada"] for s in situacao if s["Situação"] == "sem arquivo novo"]
    if sem_arquivo:
        ressalvas.append(f"{_lista(sem_arquivo)} não "
                         f"{'recebeu' if len(sem_arquivo) == 1 else 'receberam'} arquivo novo")

    for m in sorted(set(a.resumo.index) & set(b.resumo.index)):
        if a.fonte(m) != b.fonte(m):
            avisos.append(f"{_rotulo(m)} passou a vir do {b.fonte(m)} (antes vinha da "
                          f"{a.fonte(m)}): a variação desse mês é troca de fonte, e não "
                          f"do fallout.")
            ressalvas.append(f"{_rotulo(m)} passou a vir do {b.fonte(m)}")

    qualidade = _qualidade(a, b, limites)
    avisos += _avisos_qualidade(qualidade)
    suspeitos = qualidade[qualidade["Suspeito"]]
    if len(suspeitos):
        jornadas_s = [fallout_core.nome_exibicao(j) for j in suspeitos["Jornada"].unique()]
        ressalvas.insert(0, f"{_lista(jornadas_s)} "
                         f"{'perdeu' if len(jornadas_s) == 1 else 'perderam'} falhas em "
                         f"meses que já existiam"
                         + (" (queda forte: provável problema no arquivo novo)"
                            if suspeitos["Queda forte"].any() else "")
                         + " — ver Qualidade dos dados")

    if b.rel["mes_atual"] < a.rel["mes_atual"]:
        avisos.append(f"{nome}: o mês do relatório voltou de {_rotulo(a.rel['mes_atual'])} "
                      f"para {_rotulo(b.rel['mes_atual'])} — sinal de pedidos faltando na "
                      f"extração nova.")
        ressalvas.insert(0, f"o mês do relatório voltou de {_rotulo(a.rel['mes_atual'])} "
                            f"para {_rotulo(b.rel['mes_atual'])}")

    return {
        "dias": (b.corte - a.corte).days,
        "corte_antes": a.corte, "corte_agora": b.corte,
        "hoje_antes": a.hoje.date(), "hoje_agora": b.hoje.date(),
        "origem_antes": base.meta["origem"], "origem_agora": atual.meta["origem"],
        "fontes": fontes,
        "jornadas": pd.DataFrame(situacao),
        "qualidade": qualidade,
        "avisos": avisos,
        "ressalvas": ressalvas,
    }


# ── Distribuição ──────────────────────────────────────────────────────────────
def _ordenar_categorias(presentes):
    presentes = set(presentes)
    conhecidas = [c for c in ORDEM_CATEGORIAS if c in presentes]
    return conhecidas + sorted(presentes - set(ORDEM_CATEGORIAS))


def _motivo_reclassificacao(linha, hoje_a, hoje_b):
    """
    Por que um pedido mudou de categoria, a partir do que mudou nele ou no
    defeito. Quando nada mudou, foi a passagem do tempo: a entrega chegou, os
    15 dias da avaliação de eficácia terminaram — ou a regra mudou.
    """
    partes = []
    def_a, def_b = _texto(linha["DefectKey antes"]), _texto(linha["DefectKey agora"])
    sem_a, sem_b = def_a.lower() in _VAZIOS | {"-1"}, def_b.lower() in _VAZIOS | {"-1"}
    ms_b = _entrega(linha["DFT_BugfixMilestone agora"])
    if def_a != def_b:
        # 999999 não é defeito: é a marca de Falha Pontual.
        if def_b == "999999":
            partes.append("virou Falha Pontual (999999)")
        elif def_a == "999999":
            partes.append(f"deixou de ser Falha Pontual: {_defeito(def_b)}")
        elif sem_a:
            partes.append(f"ganhou o defeito {_defeito(def_b)}")
        elif sem_b:
            partes.append(f"perdeu o defeito {_defeito(def_a)}")
        else:
            partes.append(f"defeito trocado: {_defeito(def_a)} → {_defeito(def_b)}")
    else:
        # Fase e data só contam quando o defeito é o mesmo: trocado o defeito,
        # a fase "nova" é a de outro defeito, e não uma mudança.
        fase_a, fase_b = _texto(linha["DFT_Phase antes"]), _texto(linha["DFT_Phase agora"])
        if fase_a != fase_b:
            partes.append(f"fase: {fase_a or '—'} → {fase_b or '—'}")
        ms_a = _entrega(linha["DFT_BugfixMilestone antes"])
        if (ms_a is None) != (ms_b is None) or (ms_a is not None and ms_a.date() != ms_b.date()):
            partes.append(f"entrega: {_dia(ms_a)} → {_dia(ms_b)}")
    # A classificação na origem só decide categoria no Legado Aberto; fora
    # dele, a diferença costuma ser só o layout do export (o de agosto em diante
    # nem traz a coluna).
    if "Legado Aberto" in (linha["Categoria antes"], linha["Categoria agora"]):
        sub_a, sub_b = _texto(linha["SubStatus antes"]), _texto(linha["SubStatus agora"])
        if sub_a != sub_b:
            partes.append(f"classificação na origem: {sub_a or '—'} → {sub_b or '—'}")
    if partes:
        return "; ".join(partes)
    if ms_b is not None and hoje_a < ms_b <= hoje_b:
        return f"a data de entrega ({_dia(ms_b)}) chegou"
    quinze = pd.Timedelta(days=15)
    if ms_b is not None and hoje_a - quinze <= ms_b < hoje_b - quinze:
        return f"terminaram os 15 dias da avaliação de eficácia (entrega {_dia(ms_b)})"
    return "mudança de regra"


def _fluxo(migracao, categoria, sentido):
    """
    Os pedidos que entraram na categoria vindos de outra (`sentido` > 0) ou
    que saíram dela para outra, em poucas palavras, com o maior fluxo à vista.
    """
    reclass = migracao[~migracao["De"].str.startswith("(") & ~migracao["Para"].str.startswith("(")]
    if sentido > 0:
        fluxos, outra, verbo, prep = reclass[reclass["Para"] == categoria], "De", ("veio", "vieram"), "de"
    else:
        fluxos, outra, verbo, prep = reclass[reclass["De"] == categoria], "Para", ("passou", "passaram"), "para"
    if not len(fluxos):
        return ""
    total, maior = int(fluxos["Pedidos"].sum()), fluxos.iloc[0]      # value_counts: do maior ao menor
    verbo = verbo[0] if total == 1 else verbo[1]
    return f"{_int(total)} {verbo} {prep} {maior[outra]}" + (" e outras" if len(fluxos) > 1 else "")


def _explicar_categoria(linha, migracao, periodo):
    """
    O que explica a variação da categoria, em poucas palavras: as parcelas que
    puxam no mesmo sentido da variação, da maior para a menor, no máximo duas.
    A parcela contrária fica na conta completa — aqui ela só confundiria.
    """
    total = linha["Variação (pp)"]
    if total != total or abs(total) < 0.005:
        return ""
    sentido = 1 if total > 0 else -1
    parcelas = sorted(
        ((abs(linha[col]), col) for col in ("Período novo (pp)", "Reclassificação (pp)", "Removidos (pp)")
         if linha[col] * sentido > 0 and abs(linha[col]) >= 0.005),
        reverse=True)
    frases = []
    for _, col in parcelas[:2]:
        if col == "Período novo (pp)":
            novos = int(linha["Novos"])
            if sentido > 0:
                frases.append(f"{_n(novos, 'pedido novo', 'pedidos novos')} {periodo}")
            elif novos == 0:
                frases.append(f"nenhum pedido novo {periodo}")
            else:
                frases.append(f"só {_n(novos, 'pedido novo', 'pedidos novos')} {periodo}, "
                              f"abaixo do ritmo do mês")
        elif col == "Reclassificação (pp)":
            frases.append(_fluxo(migracao, linha["Categoria"], sentido))
        else:
            frases.append(f"{_n(linha['Removidos'], 'pedido saiu', 'pedidos saíram')} da extração")
    return "; ".join(f for f in frases if f)


def comparar_distribuicao(a, b, mes):
    """
    A Distribuição de `mes` antes e agora, com a variação de cada categoria
    dividida em partes que somam exatamente a variação.

    Os pedidos das duas fotografias são casados pelo número. Cada pedido do
    mês é então mantido na mesma categoria, reclassificado, novo (só na atual)
    ou removido (só na anterior). Sendo Tₐ e Tᵦ o total de pedidos do mês
    (sucesso incluído) e nₐ a contagem da categoria na anterior, a variação em
    pp é exatamente

        (novos − removidos + entraram − saíram) / Tᵦ  +  nₐ · (1/Tᵦ − 1/Tₐ)

    O último termo é o efeito do volume: o mês cresceu, e as mesmas falhas
    pesam menos. Ele e os novos contam juntos a história dos dias novos — por
    isso a tabela traz também "Período novo", a soma dos dois: positiva quando
    os dias novos tiveram, proporcionalmente, mais pedidos da categoria que o
    período que já existia.

    Devolve (tabela, migração, pedidos que mudaram, verificações).
    """
    colunas = ["Chave", "Jornada", "Dia", "Erro", "DefectKey", "DFT_Phase",
               "DFT_BugfixMilestone", "SubStatus", "Categoria"]
    ant = a.pedidos.loc[a.pedidos["Mes"] == mes, colunas]
    atu = b.pedidos.loc[b.pedidos["Mes"] == mes, colunas]
    t_a, t_b = a.total(mes), b.total(mes)

    par = ant.merge(atu, on="Chave", how="outer", suffixes=(" antes", " agora"), indicator=True)
    # O lado que falta vem vazio; como texto vazio, toda comparação com ele dá
    # falso — com o NA do pandas, daria NA e a máscara quebraria.
    cat_a = par["Categoria antes"].fillna("")
    cat_b = par["Categoria agora"].fillna("")
    ambos = par["_merge"] == "both"
    tipo = pd.Series("mantido", index=par.index)
    tipo[ambos & (cat_a != cat_b)] = "reclassificado"
    tipo[par["_merge"] == "right_only"] = "novo"
    tipo[par["_merge"] == "left_only"] = "removido"
    par["Tipo"] = tipo
    reclass = tipo == "reclassificado"

    linhas = []
    for cat in _ordenar_categorias(set(ant["Categoria"]) | set(atu["Categoria"])):
        n_a = int((ant["Categoria"] == cat).sum())
        n_b = int((atu["Categoria"] == cat).sum())
        saiu = int((reclass & (cat_a == cat)).sum())
        entrou = int((reclass & (cat_b == cat)).sum())
        novos = int(((tipo == "novo") & (cat_b == cat)).sum())
        removidos = int(((tipo == "removido") & (cat_a == cat)).sum())
        linha = {"Categoria": cat, "Pedidos antes": n_a, "Pedidos agora": n_b,
                 "Novos": novos, "Removidos": removidos, "Entraram": entrou, "Saíram": saiu}
        if t_a and t_b:
            linha.update({
                "% antes": n_a / t_a * 100,
                "% agora": n_b / t_b * 100,
                "Variação (pp)": (n_b / t_b - n_a / t_a) * 100,
                "Novos (pp)": novos / t_b * 100,
                "Volume do mês (pp)": n_a * (1 / t_b - 1 / t_a) * 100,
                "Reclassificação (pp)": (entrou - saiu) / t_b * 100,
                "Removidos (pp)": -removidos / t_b * 100,
            })
            linha["Período novo (pp)"] = linha["Novos (pp)"] + linha["Volume do mês (pp)"]
        linhas.append(linha)
    tabela = pd.DataFrame(linhas)

    verificacoes = []
    if len(tabela):
        saldo = (tabela["Pedidos antes"] - tabela["Saíram"] + tabela["Entraram"]
                 + tabela["Novos"] - tabela["Removidos"] - tabela["Pedidos agora"])
        verificacoes.append(_verif(
            "distribuição: a migração fecha com as contagens", (saldo == 0).all(),
            f"{int((saldo != 0).sum())} categorias fora"))
        verificacoes.append(_verif(
            "distribuição: as contagens somam os pedidos do mês",
            tabela["Pedidos antes"].sum() == len(ant) and tabela["Pedidos agora"].sum() == len(atu),
            f"{tabela['Pedidos antes'].sum()}/{len(ant)} e {tabela['Pedidos agora'].sum()}/{len(atu)}"))
    if "Variação (pp)" in tabela:
        partes = tabela[["Novos (pp)", "Volume do mês (pp)", "Reclassificação (pp)",
                         "Removidos (pp)"]].sum(axis=1)
        erro = (partes - tabela["Variação (pp)"]).abs().max()
        verificacoes.append(_verif("distribuição: as partes somam a variação de cada categoria",
                                   erro < TOLERANCIA, f"maior diferença {erro:.2e} pp"))
        # Nos meses que vêm da extração, a taxa é falhas/total — e as falhas são
        # exatamente as categorias. Nos do histórico a taxa é a digitada, que
        # não precisa bater com a soma.
        if a.fonte(mes) == "extração" and b.fonte(mes) == "extração":
            delta = b.taxa(mes) - a.taxa(mes)
            soma = tabela["Variação (pp)"].sum()
            verificacoes.append(_verif(
                "distribuição: as categorias somam a variação do fallout",
                abs(soma - delta) < TOLERANCIA, f"{soma:.6f} x {delta:.6f} pp"))

    migracao = (pd.DataFrame({
        "De": par.loc[tipo != "mantido", "Categoria antes"].fillna("(pedido novo)"),
        "Para": par.loc[tipo != "mantido", "Categoria agora"].fillna("(saiu da extração)"),
    }).value_counts().rename("Pedidos").reset_index())
    if "Variação (pp)" in tabela:
        periodo = "na semana" if (b.corte - a.corte).days <= 8 else "no período"
        tabela.insert(tabela.columns.get_loc("Variação (pp)") + 1, "O que explica",
                      [_explicar_categoria(l, migracao, periodo) for _, l in tabela.iterrows()])

    mudaram = par[tipo != "mantido"].copy()
    motivo = pd.Series("", index=mudaram.index, dtype=object)
    rec = mudaram["Tipo"] == "reclassificado"
    motivo[rec] = [_motivo_reclassificacao(l, a.hoje, b.hoje) for _, l in mudaram[rec].iterrows()]
    dia = mudaram["Dia agora"].where(mudaram["Dia agora"].notna(), mudaram["Dia antes"])
    # O pedido novo com data até o corte da anterior já existia quando ela foi
    # tirada: a extração daquela jornada parava antes do corte (o corte é o
    # mesmo para todas), ou o pedido entrou no export depois.
    motivo[mudaram["Tipo"] == "novo"] = [
        "criado depois da fotografia anterior" if d > a.corte
        else "já existia na data da anterior, mas não estava nela"
        for d in dia[mudaram["Tipo"] == "novo"]]
    motivo[mudaram["Tipo"] == "removido"] = "saiu da extração"
    pedidos = pd.DataFrame({
        "Pedido": mudaram["Chave"],
        "Jornada": mudaram["Jornada agora"].fillna(mudaram["Jornada antes"]),
        "Criado em": dia,
        "Tipo": mudaram["Tipo"],
        "De": mudaram["Categoria antes"],
        "Para": mudaram["Categoria agora"],
        "Motivo": motivo,
        "Defeito antes": mudaram["DefectKey antes"],
        "Defeito agora": mudaram["DefectKey agora"],
        "Fase antes": mudaram["DFT_Phase antes"],
        "Fase agora": mudaram["DFT_Phase agora"],
        "Entrega antes": mudaram["DFT_BugfixMilestone antes"],
        "Entrega agora": mudaram["DFT_BugfixMilestone agora"],
        "Mensagem de erro": mudaram["Erro agora"].fillna(mudaram["Erro antes"]),
    }).sort_values(["Tipo", "De", "Para", "Pedido"], na_position="first").reset_index(drop=True)
    return tabela, migracao, pedidos, verificacoes


# ── Projeção ──────────────────────────────────────────────────────────────────
def _contribuicoes(lado, chave_mes):
    """
    Quanto cada defeito do Planejamento de Redução tira do mês projetado
    `chave_mes`, pela regra do pipeline: cheio (`Pct_plena`) se a entrega é de
    um mês anterior, pro-rata (`Pct_mes`) se é do próprio mês, nada se é
    depois. Um defeito por linha, somando as datas quando há mais de uma.
    """
    red = lado.reducao
    colunas = ["Entrega", "Pedidos", "Peso", "Contribuição"]
    if not len(red):
        return pd.DataFrame(columns=colunas, index=pd.Index([], name="DefectKey"))
    entrega = pd.to_datetime(red["MilestoneDate"])
    mes_entrega = entrega.dt.year * 100 + entrega.dt.month
    contrib = red["Pct_plena"].where(mes_entrega < chave_mes,
                                     red["Pct_mes"].where(mes_entrega == chave_mes, 0.0))
    # "Peso" é a fatia do mês do relatório que o defeito representa (o
    # `Pct_plena`). É ele, e não a contagem, que a projeção desconta: com o mês
    # crescendo, um defeito pode ganhar pedidos e ainda assim pesar menos.
    return (pd.DataFrame({"DefectKey": red["DefectKey"].astype(str),
                          "Entrega": entrega.dt.date, "Pedidos": red["n"],
                          "Peso": red["Pct_plena"], "Contribuição": contrib})
            .groupby("DefectKey")
            .agg(**{"Entrega": ("Entrega", "max"), "Pedidos": ("Pedidos", "sum"),
                    "Peso": ("Peso", "sum"), "Contribuição": ("Contribuição", "sum")}))


def _no_plano(chave, antes, agora, attrs_a, attrs_b, hoje_a, hoje_b):
    """
    Tipo e detalhe da mudança de um defeito no Planejamento de Redução.

    O plano só tem defeito em tratamento com entrega futura; sair dele quer
    dizer que a entrega chegou, a data passou sem a correção, a data sumiu ou
    os pedidos foram para outra categoria.
    """
    esteve, esta = pd.notna(antes["Entrega"]), pd.notna(agora["Entrega"])
    if esteve and esta:
        if antes["Entrega"] != agora["Entrega"]:
            sentido = "adiada" if agora["Entrega"] > antes["Entrega"] else "antecipada"
            return f"data {sentido}", f"{_dia(antes['Entrega'])} → {_dia(agora['Entrega'])}"
        return "peso mudou", (f"{_pct(antes['Peso'])} → {_pct(agora['Peso'])} do mês, "
                              f"com {_int(antes['Pedidos'])} → {_int(agora['Pedidos'])} pedidos")
    if esta:
        if chave not in attrs_a.index:
            return "entrou no plano", f"defeito novo na extração, entrega {_dia(agora['Entrega'])}"
        if pd.isna(attrs_a.loc[chave, "DFT_BugfixMilestone"]):
            return "entrou no plano", f"ganhou data de entrega: {_dia(agora['Entrega'])}"
        return "entrou no plano", f"entrega {_dia(agora['Entrega'])}"
    if chave not in attrs_b.index:
        return "saiu do plano", "não tem mais pedidos"
    fase = _texto(attrs_b.loc[chave, "DFT_Phase"])
    entrega = _entrega(attrs_b.loc[chave, "DFT_BugfixMilestone"])
    if entrega is None:
        return "saiu do plano", "perdeu a data de entrega"
    if entrega <= hoje_b:
        if fase in fallout_core.FASES_CORRIGIDO:
            return "saiu do plano", f"entregue em {_dia(entrega)} ({fase})"
        return "saiu do plano", f"a data de entrega ({_dia(entrega)}) passou; fase: {fase or '—'}"
    if fase in fallout_core.FASES_MOPS:
        return "saiu do plano", f"defeito {fase.lower()}"
    return "saiu do plano", "os pedidos foram para outra categoria"


def comparar_projecao(a, b):
    """
    As duas projeções e, para cada mês projetado que as duas têm, a variação
    dividida em partes que somam exatamente a variação.

    A projeção do pipeline é, para o mês m,

        max(F − N − Σ contribuição de cada defeito em m, 0)

    com F o fallout do mês do relatório e N o que não se repete (Resolvido,
    Tratado e Incidente). A variação se divide então em: o ponto de partida
    (ΔF), o que não se repete (−ΔN), cada defeito do plano (−Δcontribuição) e,
    quando alguma das duas bateu no zero, o piso.

    Antes de dividir, a conta do pipeline é refeita com o que a fotografia
    guardou e conferida contra a projeção dela — senão a explicação poderia
    estar explicando outro número.
    """
    ra, rb = a.rel, b.rel
    proj_a = dict(zip(ra["proj_labels"][1:], ra["proj_vals"][1:]))
    proj_b = dict(zip(rb["proj_labels"][1:], rb["proj_vals"][1:]))
    attrs_a, attrs_b = _atributos(a), _atributos(b)
    fa, fb = ra["fallout_pct"], rb["fallout_pct"]
    na_, nb_ = ra["pct_nao_recorrente"], rb["pct_nao_recorrente"]

    meses, itens, verificacoes, notas = [], [], [], []
    for rot in [m for m in proj_b if m in proj_a]:
        va, vb = proj_a[rot], proj_b[rot]
        if None in (va, vb, fa, fb, na_, nb_):
            notas.append(f"{rot}: uma das projeções está vazia — sem como explicar a variação.")
            continue
        chave = _chave_do_rotulo(rot)
        ca, cb = _contribuicoes(a, chave), _contribuicoes(b, chave)
        for quando, f, n, c, v in (("anterior", fa, na_, ca, va), ("atual", fb, nb_, cb, vb)):
            refeita = max(f - n - c["Contribuição"].sum(), 0.0)
            verificacoes.append(_verif(
                f"projeção: a conta de {rot} refeita bate ({quando})",
                abs(refeita - v) < TOLERANCIA, f"{refeita:.6f} x {v:.6f}"))

        partida, nao_rec = fb - fa, -(nb_ - na_)
        defeitos = ca.join(cb, how="outer", lsuffix=" antes", rsuffix=" agora")
        defeitos["Efeito"] = (defeitos["Contribuição antes"].fillna(0)
                              - defeitos["Contribuição agora"].fillna(0))
        piso = (vb - va) - (partida + nao_rec + defeitos["Efeito"].sum())
        chegou_ao_zero = min(va, vb) <= TOLERANCIA
        verificacoes.append(_verif(
            f"projeção: as partes somam a variação de {rot}",
            chegou_ao_zero or abs(piso) < TOLERANCIA,
            f"resto {piso:.2e} pp" + (" (piso de zero)" if chegou_ao_zero else "")))

        meses.append({"Mês": rot, "Antes": va, "Agora": vb, "Variação (pp)": vb - va,
                      "Ponto de partida (pp)": partida, "Já corrigido (pp)": nao_rec,
                      "Defeitos do plano (pp)": defeitos["Efeito"].sum(),
                      "Piso de zero (pp)": piso if abs(piso) >= TOLERANCIA else 0.0})

        mes_rel_a, mes_rel_b = _rotulo(ra["mes_atual"]), _rotulo(rb["mes_atual"])
        itens.append({"Mês": rot, "Item": "Ponto de partida", "Tipo": "fallout do mês do relatório",
                      "Detalhe": f"{mes_rel_a} {_pct(fa)} → {mes_rel_b} {_pct(fb)}",
                      "Efeito (pp)": partida})
        itens.append({"Mês": rot, "Item": "Já corrigido", "Tipo": "Resolvido, Tratado e Incidente",
                      "Detalhe": f"{_pct(na_)} → {_pct(nb_)} do mês",
                      "Efeito (pp)": nao_rec})
        for chave_def, linha in defeitos.iterrows():
            if abs(linha["Efeito"]) < 1e-9:
                continue
            antes = {"Entrega": linha["Entrega antes"], "Pedidos": linha["Pedidos antes"],
                     "Peso": linha["Peso antes"]}
            agora = {"Entrega": linha["Entrega agora"], "Pedidos": linha["Pedidos agora"],
                     "Peso": linha["Peso agora"]}
            tipo, detalhe = _no_plano(chave_def, antes, agora, attrs_a, attrs_b, a.hoje, b.hoje)
            itens.append({"Mês": rot, "Item": _defeito(chave_def), "Tipo": tipo,
                          "Detalhe": detalhe, "Efeito (pp)": linha["Efeito"]})
        if abs(piso) >= TOLERANCIA:
            itens.append({"Mês": rot, "Item": "Piso de zero", "Tipo": "a projeção não fica negativa",
                          "Detalhe": "", "Efeito (pp)": piso})

    if not meses and not notas:
        notas.append("As duas fotografias não projetam nenhum mês em comum.")

    def _curva(r):
        return {"labels_hist": r["labels_hist"], "vals_hist": r["vals_hist"],
                "plan_hist": r["plan_hist"], "proj_labels": r["proj_labels"],
                "proj_vals": r["proj_vals"]}

    itens = pd.DataFrame(itens, columns=["Mês", "Item", "Tipo", "Detalhe", "Efeito (pp)"])
    if len(itens):
        itens = (itens.assign(_ordem=itens["Efeito (pp)"].abs())
                 .sort_values(["Mês", "_ordem"], ascending=[True, False], kind="stable")
                 .drop(columns="_ordem").reset_index(drop=True))
    return {"curvas": {"antes": _curva(ra), "agora": _curva(rb)},
            "meses": pd.DataFrame(meses), "itens": itens,
            "notas": notas}, verificacoes


# ── Defeitos ──────────────────────────────────────────────────────────────────
# O sinal de cada defeito, do mais grave ao bom: o que precisa de ação, o que
# pede atenção e o que melhorou. Um defeito leva só o mais grave dos que tem;
# os outros continuam nas Observações.
SINAIS_DEFEITO = {
    "Correção não pegou": "pior",
    "Passou da data": "pior",
    "Fora do Octane": "atencao",
    "Novo ofensor": "atencao",
    "Entrega adiada": "atencao",
    "Perdeu a data": "atencao",
    "Saiu dos ofensores": "melhor",
}


def _sinal_defeito(situacao, pos, fase, ms_a, ms_b, pedidos, no_rel_antes, no_rel_agora):
    """
    (sinal, detalhe curto) do defeito; ('', '') sem sinal. `pos` são os
    pedidos depois da entrega quando eles são alerta, e 0 quando não são.
    Entrar ou sair dos ofensores só vira sinal com pedidos no mês do
    relatório: a tabela de ofensores tem vários meses, e um defeito que entra
    nela por um mês antigo, com zero pedidos no atual, não é notícia.
    """
    tem = lambda inicio: any(s.startswith(inicio) for s in situacao)
    if pos and fase in fallout_core.FASES_CORRIGIDO:
        return ("Correção não pegou",
                f"entregue em {_dia(ms_b)}; {_n(pos, 'pedido', 'pedidos')} depois da entrega")
    if pos:
        return ("Passou da data",
                f"a entrega era {_dia(ms_b)}, fase {fase or '—'}; {_n(pos, 'pedido', 'pedidos')} depois")
    if tem("não está no Octane") and pedidos:
        return "Fora do Octane", "sem nome, fase nem data"
    if tem("novo ofensor") and no_rel_agora:
        entrega = f"entrega {_dia(ms_b)}" if ms_b is not None else "sem data de entrega"
        return "Novo ofensor", " · ".join(p for p in (fase, entrega) if p)
    if tem("entrega adiada"):
        return "Entrega adiada", f"{_dia(ms_a)} → {_dia(ms_b)}"
    if tem("perdeu a data"):
        return "Perdeu a data", f"era {_dia(ms_a)}"
    if tem("saiu dos ofensores") and no_rel_antes:
        return "Saiu dos ofensores", (f"entregue em {_dia(ms_b)}"
                                      if ms_b is not None and fase in fallout_core.FASES_CORRIGIDO
                                      else fase)
    return "", ""


def comparar_defeitos(a, b, mes, limites):
    """
    Cada defeito com pedido no mês ou na tabela de ofensores de alguma das
    duas fotografias: volume, fase e data de entrega antes e agora, e o que
    merece atenção.

    O alerta mais forte é "pedidos depois da entrega": pedidos criados no
    período novo depois da data de entrega do defeito. Se a fase é de
    corrigido, a correção não pegou; se não é, a data passou e o erro
    continua. É o que hoje se descobre à mão para `DFTS_CORRECAO_INEFICAZ`.
    """
    attrs_a, attrs_b = _atributos(a), _atributos(b)

    def _no_mes(lado):
        p = lado.pedidos
        return p[(p["Mes"] == mes) & p["TemDefeito"]].groupby("DefectKey").size()

    n_a, n_b = _no_mes(a), _no_mes(b)
    ofe_a = set(a.ofensores["DefectKey"].dropna())
    ofe_b = set(b.ofensores["DefectKey"].dropna())

    # A tabela de ofensores é do mês do relatório de cada fotografia, que pode
    # não ser `mes` (na virada do mês, a anterior ainda está no mês passado).
    def _no_relatorio(lado):
        o = lado.ofensores
        return o[o["Mes"] == lado.rel["mes_atual"]].groupby("DefectKey")["Pedidos"].sum()

    rel_a, rel_b = _no_relatorio(a), _no_relatorio(b)

    # Pedidos do período novo criados depois da data de entrega do defeito —
    # a mesma comparação por dia de `fallout_core.precisa_correcao`.
    novos = b.pedidos[(b.pedidos["Dia"] > a.corte) & b.pedidos["TemDefeito"]]
    entrega_dia = fallout_core.dia_local(novos["DFT_BugfixMilestone"])
    depois = (entrega_dia.notna() & (pd.to_datetime(novos["Dia"]) > entrega_dia)
              & ~fallout_core.us_em_aberto(novos))
    pos_entrega = novos[depois].groupby("DefectKey").size()

    t_a, t_b = a.total(mes), b.total(mes)
    linhas = []
    for chave in sorted(set(n_a.index) | set(n_b.index) | ofe_a | ofe_b):
        at_a = attrs_a.loc[chave] if chave in attrs_a.index else None
        at_b = attrs_b.loc[chave] if chave in attrs_b.index else None
        atual_ou_antigo = at_b if at_b is not None else at_a

        def _campo(at, col):
            return None if at is None else at[col]

        fase_a, fase_b = _texto(_campo(at_a, "DFT_Phase")), _texto(_campo(at_b, "DFT_Phase"))
        ms_a, ms_b = _entrega(_campo(at_a, "DFT_BugfixMilestone")), _entrega(_campo(at_b, "DFT_BugfixMilestone"))
        ped_a, ped_b = int(n_a.get(chave, 0)), int(n_b.get(chave, 0))
        pos = int(pos_entrega.get(chave, 0))

        situacao = []
        alerta_pos = pos >= limites.min_pos_entrega and ms_b is not None and ms_b <= b.hoje
        if alerta_pos:
            if fase_b in fallout_core.FASES_CORRIGIDO:
                situacao.append(f"correção não pegou: {pos} pedidos novos depois da entrega de {_dia(ms_b)}")
            else:
                situacao.append(f"{pos} pedidos novos depois da data de entrega ({_dia(ms_b)}), "
                                f"fase {fase_b or '—'}")
        if chave in ofe_b and chave not in ofe_a:
            situacao.append("novo ofensor")
        if chave in ofe_a and chave not in ofe_b:
            situacao.append("saiu dos ofensores")
        if fase_a != fase_b and at_a is not None and at_b is not None:
            situacao.append(f"fase: {fase_a or '—'} → {fase_b or '—'}")
        if at_a is not None and at_b is not None:
            if ms_a is None and ms_b is not None:
                situacao.append(f"ganhou data de entrega: {_dia(ms_b)}")
            elif ms_a is not None and ms_b is None:
                situacao.append("perdeu a data de entrega")
            elif ms_a is not None and ms_a.date() != ms_b.date():
                sentido = "adiada" if ms_b > ms_a else "antecipada"
                situacao.append(f"entrega {sentido}: {_dia(ms_a)} → {_dia(ms_b)}")
        if ms_b is not None and a.hoje < ms_b <= b.hoje:
            situacao.append(f"data de entrega na semana ({_dia(ms_b)})")
        # Número de DFT sem fase é defeito que o Octane da pasta não conhece —
        # em geral, aberto depois da última exportação dele. Os pedidos ficam
        # em tratamento sem data, e a projeção não conta com ele.
        if at_b is not None and chave.isdigit() and not fase_b:
            situacao.append("não está no Octane")
        if chave in fallout_core.DFTS_CORRECAO_INEFICAZ:
            situacao.append("já marcado como correção que não resolveu")
        sinal, detalhe = _sinal_defeito(situacao, pos if alerta_pos else 0, fase_b,
                                        ms_a, ms_b, ped_b,
                                        int(rel_a.get(chave, 0)), int(rel_b.get(chave, 0)))

        linha = {
            "Defeito": _defeito(chave), "DefectKey": chave,
            "Nome": _texto(_campo(atual_ou_antigo, "DFT_Name")),
            "Time": _texto(_campo(atual_ou_antigo, "DFT_Team")),
            "Tipo": "US" if _texto(_campo(atual_ou_antigo, "DFT_Type")) == "User Story" else "Defeito",
            "Fase antes": fase_a, "Fase agora": fase_b,
            "Entrega antes": None if ms_a is None else ms_a.date(),
            "Entrega agora": None if ms_b is None else ms_b.date(),
            "Pedidos antes": ped_a, "Pedidos agora": ped_b, "Variação": ped_b - ped_a,
            "% antes": ped_a / t_a * 100 if t_a else None,
            "% agora": ped_b / t_b * 100 if t_b else None,
            "Pós-entrega na semana": pos if ms_b is not None and ms_b <= b.hoje else 0,
            "Ofensor antes": chave in ofe_a, "Ofensor agora": chave in ofe_b,
            "No relatório antes": int(rel_a.get(chave, 0)),
            "No relatório agora": int(rel_b.get(chave, 0)),
            "Sinal": sinal, "Detalhe do sinal": detalhe,
            "Observações": "; ".join(situacao),
        }
        linhas.append(linha)

    tabela = pd.DataFrame(linhas)
    if not len(tabela):
        return tabela
    alerta = tabela["Pós-entrega na semana"] >= limites.min_pos_entrega
    return (tabela.assign(_alerta=alerta, _tamanho=tabela["Variação"].abs())
            .sort_values(["_alerta", "_tamanho", "Pedidos agora"], ascending=False, kind="stable")
            .drop(columns=["_alerta", "_tamanho"]).reset_index(drop=True))


# ── Erros ─────────────────────────────────────────────────────────────────────
def _descrever_erros(pedidos, mensagens):
    """Categoria principal, fatia sem defeito e defeitos associados de cada mensagem."""
    mensagens = list(mensagens)
    if not mensagens:
        return pd.DataFrame(columns=["Sem defeito associado (%)", "Categoria principal", "Defeitos"])
    parte = pedidos[pedidos["Erro"].isin(mensagens)]
    grupos = parte.groupby("Erro")
    sem_defeito = grupos["Categoria"].agg(lambda c: (c == CAT_FALTA_ASSOCIAR).mean() * 100)
    principal = grupos["Categoria"].agg(lambda c: c.value_counts().index[0])

    def _defeitos(linhas):
        contagem = linhas.loc[linhas["TemDefeito"], "DefectKey"].value_counts().head(3)
        return ", ".join(f"{_defeito(k)} ({n})" for k, n in contagem.items())

    defeitos = grupos[["TemDefeito", "DefectKey"]].apply(_defeitos)
    return pd.DataFrame({"Sem defeito associado (%)": sem_defeito,
                         "Categoria principal": principal,
                         "Defeitos": defeitos}).reindex(mensagens)


def comparar_erros(a, b, mes, limites):
    """
    Mensagens de erro novas, em alta e que sumiram.

      nova     — não aparece em pedido nenhum da fotografia anterior e tem
                 pelo menos `min_ocorrencias` na atual;
      em alta  — no período novo (depois da data de corte da anterior), a
                 média diária chega a `fator_alta` vezes a das
                 `dias_referencia` anteriores, com pelo menos
                 `min_ocorrencias` ocorrências; se não aparecia na referência
                 mas existia antes, ela "voltou";
      sumiu    — tinha pelo menos `min_ocorrencias` na referência e nenhuma
                 no período novo.

    O "em alta" e o "sumiu" olham só a extração atual, nas duas janelas: a
    fonte é uma só, e uma mudança no arquivo não vira tendência. A anterior
    entra só para dizer o que é novo.

    Devolve (novos, em alta, sumidos, informações do período).
    """
    ant, atu = a.pedidos, b.pedidos
    vazio = pd.DataFrame()
    dias_com_dado = atu["Dia"].dropna()
    ultimo = min(b.corte, max(dias_com_dado)) if len(dias_com_dado) else a.corte
    dias_novos = (ultimo - a.corte).days
    info = {"inicio": a.corte + _dt.timedelta(days=1), "fim": ultimo, "dias": dias_novos}
    if dias_novos <= 0:
        info["nota"] = ("A extração atual não tem pedido depois da data de corte da "
                        "anterior: não há período novo para comparar os erros.")
        return vazio, vazio, vazio, info

    inicio_ref = a.corte - _dt.timedelta(days=limites.dias_referencia - 1)
    primeiro_dia = min(dias_com_dado)
    dias_ref = (a.corte - max(inicio_ref, primeiro_dia)).days + 1
    periodo = atu[(atu["Dia"] > a.corte) & (atu["Dia"] <= ultimo)]
    referencia = atu[(atu["Dia"] >= inicio_ref) & (atu["Dia"] <= a.corte)]
    info.update({"inicio_ref": max(inicio_ref, primeiro_dia), "dias_ref": dias_ref})

    no_periodo = periodo["Erro"].value_counts()
    na_referencia = referencia["Erro"].value_counts()
    total_atual = atu["Erro"].value_counts()
    desde = atu.groupby("Erro")["Dia"].min()
    ultima = atu.groupby("Erro")["Dia"].max()
    existiam = set(ant["Erro"])
    t_mes = b.total(mes)
    no_mes = atu[atu["Mes"] == mes]["Erro"].value_counts()

    def _pp_mes(msg):
        return no_mes.get(msg, 0) / t_mes * 100 if t_mes else None

    # Novos
    msgs = [m for m in total_atual.index
            if m not in existiam and total_atual[m] >= limites.min_ocorrencias]
    novos = pd.DataFrame({
        "Mensagem de erro": msgs,
        "Ocorrências": [int(total_atual[m]) for m in msgs],
        "No período novo": [int(no_periodo.get(m, 0)) for m in msgs],
        "Desde": [desde[m] for m in msgs],
        "Peso no mês (pp)": [_pp_mes(m) for m in msgs],
    })
    if len(novos):
        descr = _descrever_erros(atu, msgs)
        novos = novos.join(descr.reset_index(drop=True))

    # Em alta (e as que voltaram)
    linhas = []
    for msg, qtd in no_periodo.items():
        if msg not in existiam or qtd < limites.min_ocorrencias:
            continue
        ref = int(na_referencia.get(msg, 0))
        media_ref = ref / dias_ref if dias_ref > 0 else 0.0
        media_nova = qtd / dias_novos
        if media_ref == 0:
            vezes = None                      # voltou depois de sumir da referência
        elif media_nova >= limites.fator_alta * media_ref:
            vezes = media_nova / media_ref
        else:
            continue
        linhas.append({"Mensagem de erro": msg, "No período novo": int(qtd),
                       "Na referência": ref,
                       "Média diária antes": media_ref, "Média diária agora": media_nova,
                       "Vezes": vezes,
                       "Situação": "voltou" if vezes is None else "em alta",
                       "Acima do esperado": qtd - media_ref * dias_novos,
                       "Peso no mês (pp)": _pp_mes(msg)})
    em_alta = pd.DataFrame(linhas)
    if len(em_alta):
        em_alta = (em_alta.sort_values("Acima do esperado", ascending=False)
                   .reset_index(drop=True))
        em_alta = em_alta.join(
            _descrever_erros(periodo, em_alta["Mensagem de erro"]).reset_index(drop=True))

    # Sumiram. Sem a "(sem mensagem de erro)": ela some quando o export passa a
    # preencher a mensagem (o layout de agosto em diante preenche), e isso é
    # mudança de arquivo, não erro que parou de acontecer.
    msgs = [m for m, qtd in na_referencia.items()
            if qtd >= limites.min_ocorrencias and no_periodo.get(m, 0) == 0
            and m != SEM_MENSAGEM]
    sumidos = pd.DataFrame({
        "Mensagem de erro": msgs,
        "Na referência": [int(na_referencia[m]) for m in msgs],
        "Última ocorrência": [ultima[m] for m in msgs],
    })
    if len(sumidos):
        sumidos = sumidos.join(_descrever_erros(referencia, msgs).reset_index(drop=True))

    if len(novos):
        novos = novos.sort_values("Ocorrências", ascending=False).reset_index(drop=True)
    return novos, em_alta, sumidos, info


def _janela(dias):
    """28 -> 'nas 4 semanas'; 10 -> 'nos 10 dias'."""
    return f"nas {dias // 7} semanas" if dias % 7 == 0 and dias >= 14 else f"nos {dias} dias"


def _onde_caiu(linha):
    """Onde o erro está classificado: o defeito mais comum, ou a categoria."""
    if linha["Sem defeito associado (%)"] >= 50:
        return "sem defeito associado"
    defeitos = _texto(linha["Defeitos"])
    return defeitos.split(" (")[0] if defeitos else linha["Categoria principal"]


def erros_em_atencao(c):
    """
    Os erros novos e os em alta numa tabela só, do maior volume ao menor:
    sinal, quantos pedidos, a comparação com antes e onde o erro caiu.
    `Pedidos` é o total do erro novo (a anterior não tinha nenhum) e, no
    em alta, o do período novo.
    """
    colunas = ["Sinal", "Mensagem de erro", "Pedidos", "Comparado a antes", "Classificado como"]
    dias_ref = c.contexto["periodo_erros"].get("dias_ref", c.limites.dias_referencia)
    linhas = []
    for _, l in c.erros_novos.iterrows():
        linhas.append(["Novo", l["Mensagem de erro"], int(l["Ocorrências"]),
                       f"desde {_dia(l['Desde'])}", _onde_caiu(l)])
    for _, l in c.erros_em_alta.iterrows():
        voltou = l["Situação"] == "voltou"
        if voltou:
            antes = f"nenhum {_janela(dias_ref)} antes"
        else:
            vezes = f"{l['Vezes']:.1f}".replace(".", ",")
            antes = f"{vezes}× a média {_janela(dias_ref).replace('nas ', 'das ').replace('nos ', 'dos ')} antes"
        linhas.append(["Voltou" if voltou else "Em alta", l["Mensagem de erro"],
                       int(l["No período novo"]), antes, _onde_caiu(l)])
    tabela = pd.DataFrame(linhas, columns=colunas)
    return tabela.sort_values("Pedidos", ascending=False, kind="stable").reset_index(drop=True)


# ── Texto ─────────────────────────────────────────────────────────────────────
# Variação abaixo disto, em pp, é "estável" no texto — no fallout do mês e,
# com a folga maior de quem ainda é estimativa, na projeção.
ESTAVEL = 0.05
ESTAVEL_PROJECAO = 0.10


def _pedidos(n):
    """1 -> '1 pedido'; 1234 -> '1.234 pedidos'."""
    return _n(n, "pedido", "pedidos")


def _item_projecao(item, a, b):
    """Um item da explicação da projeção, em palavras."""
    nome, efeito, tipo = item["Item"], item["Efeito (pp)"], item["Tipo"]
    if nome == "Ponto de partida":
        ma, mb = a.rel["mes_atual"], b.rel["mes_atual"]
        if ma != mb:
            return f"a virada do mês do relatório ({_rotulo(ma)} → {_rotulo(mb)})"
        return f"a {'alta' if efeito > 0 else 'queda'} do fallout de {_rotulo(mb)}"
    if nome == "Já corrigido":
        return "o que já está corrigido" if efeito < 0 else "a queda do que já está corrigido"
    if nome == "Piso de zero":
        return "o piso de zero"
    if tipo == "saiu do plano":
        return f"{nome} fora do plano"
    if tipo == "entrou no plano":
        return f"{nome} no plano"
    if tipo.startswith("data "):
        return f"{nome} com entrega {tipo.split()[1]}"
    return f"{nome} com {'mais' if efeito < 0 else 'menos'} pedidos a corrigir"


def _por(texto):
    """'a queda…' -> 'pela queda…'; 'o que…' -> 'pelo que…'; o resto, 'por …'."""
    for artigo, contraido in (("a ", "pela "), ("o ", "pelo ")):
        if texto.startswith(artigo):
            return contraido + texto[len(artigo):]
    return "por " + texto


# Categorias que a projeção desconta inteiras no mês seguinte (ver
# `pct_nao_recorrente` no pipeline): o que sobra é o que se repete.
CATEGORIAS_NAO_RECORRENTES = {"Resolvido", "Tratado - Em avaliação de eficácia",
                              fallout_core.CAT_INCIDENTE}


def _texto_virada(c, a, b, itens):
    """
    A explicação da projeção quando o mês do relatório virou, ou None.

    Na virada, "ponto de partida" e "já corrigido" mudam muito e quase se
    anulam — no Base + Cross Sell de 07/10, −0,71 pp e +1,10 pp —, e a frase
    antiga dizia só que um "pesou para cima" e o outro "para baixo". O que move
    a projeção é a diferença dos dois: o fallout que se repete de um mês para o
    outro (o do relatório menos Resolvido, Tratado e Incidente). Ela é dita
    com as categorias que mais a mudaram, tiradas da distribuição do relatório
    de cada fotografia.
    """
    ra, rb = a.rel, b.rel
    if None in (ra["fallout_pct"], rb["fallout_pct"], ra["pct_nao_recorrente"],
                rb["pct_nao_recorrente"]):
        return None
    base_a = ra["fallout_pct"] - ra["pct_nao_recorrente"]
    base_b = rb["fallout_pct"] - rb["pct_nao_recorrente"]
    dif_base = base_b - base_a
    partes = [f"o fallout que se repete foi de {_pct(base_a)} em {_rotulo(ra['mes_atual'])} "
              f"(sem Resolvido, Tratado e Incidente) para {_pct(base_b)} em "
              f"{_rotulo(rb['mes_atual'])} ({_pp(dif_base)})"]

    dist = c.contexto.get("relatorio", {}).get("distribuicao")
    if dist is not None and len(dist) and abs(dif_base) >= ESTAVEL:
        recorrentes = dist[~dist["Subgrupo"]
                           & ~dist["Categoria"].isin(CATEGORIAS_NAO_RECORRENTES)]
        sentido = 1 if dif_base > 0 else -1
        puxaram = recorrentes[recorrentes["Diferença (pp)"] * sentido >= ESTAVEL]
        puxaram = puxaram.reindex(puxaram["Diferença (pp)"].abs()
                                  .sort_values(ascending=False).index).head(3)
        if len(puxaram):
            partes[0] += ", " + ("puxado por " if sentido > 0 else "com a queda de ") + _lista(
                f"{l['Categoria']} ({_pp(l['Diferença (pp)'])})" for _, l in puxaram.iterrows())

    # O que sobra são os defeitos do plano (e o piso), com a regra de sempre.
    resto = itens[~itens["Item"].isin(["Ponto de partida", "Já corrigido"])]
    efeito_resto = resto["Efeito (pp)"].sum()
    if abs(efeito_resto) >= ESTAVEL:
        partes.append(f"as entregas planejadas tiram {_pct(abs(efeito_resto)).replace('%', ' pp')} "
                      f"{'a mais' if efeito_resto < 0 else 'a menos'}")
    dias = c.contexto.get("relatorio", {}).get("dias_agora")
    aviso = (f" {_rotulo(rb['mes_atual']).capitalize()} ainda tem só {_n(dias, 'dia', 'dias')} "
             f"de dados." if dias is not None and dias < fallout_core.MIN_DIAS_MES * 2 else "")
    return "; ".join(partes) + "." + aviso


def _texto_projecao(c, a, b):
    """A linha da projeção: o primeiro mês projetado e o que o moveu."""
    meses = c.projecao["meses"]
    if not len(meses):
        return None
    m = meses.iloc[0]
    dif = m["Variação (pp)"]
    if abs(dif) < ESTAVEL_PROJECAO:
        estado = f"ficou quase estável em {_pct(m['Agora'])}"
    else:
        estado = f"{'subiu' if dif > 0 else 'caiu'} para {_pct(m['Agora'])}"
    # Os itens vão sem o número: o quanto cada um pesou está na aba Projeção,
    # e aqui ele só alongaria a frase.
    itens = c.projecao["itens"]
    itens = itens[itens["Mês"] == m["Mês"]]
    if a.rel["mes_atual"] != b.rel["mes_atual"]:
        virada = _texto_virada(c, a, b, itens)
        if virada:
            return f"Projeção de {m['Mês']} {estado} ({_pp(dif)}): {virada}"
    fortes = [(i["Efeito (pp)"], _item_projecao(i, a, b))
              for _, i in itens.iterrows()
              if abs(i["Efeito (pp)"]) >= ESTAVEL]
    sobe = [d for e, d in fortes if e > 0][:2]
    desce = [d for e, d in fortes if e < 0][:2]
    if sobe and desce:
        explica = (f"{_lista(sobe)} {'pesou' if len(sobe) == 1 else 'pesaram'} para cima; "
                   f"{_lista(desce)}, para baixo")
    elif sobe or desce:
        explica = "puxada " + _por(_lista(sobe or desce))
    else:
        explica = ""
    return f"Projeção de {m['Mês']} {estado} ({_pp(dif)})" + (f": {explica}." if explica else ".")


def _textos_erros(c, limites, info_erros):
    """As linhas dos erros: o maior em alta (ou que voltou) e os novos, o maior volume primeiro."""
    if info_erros.get("nota"):
        return [info_erros["nota"]]
    atencao = erros_em_atencao(c)
    periodo = "na semana" if info_erros["dias"] <= 8 else "no período"
    linhas = []
    alta = atencao[atencao["Sinal"] != "Novo"]
    if len(alta):
        l = alta.iloc[0]
        msg = resumir_mensagem(l["Mensagem de erro"], 70)
        if l["Sinal"] == "Voltou":
            dias_ref = info_erros.get("dias_ref", limites.dias_referencia)
            janela = _janela(dias_ref).split(" ", 1)[1]            # "4 semanas"
            frase = (f"O erro \"{msg}\" voltou: {_pedidos(l['Pedidos'])} {periodo}, "
                     f"depois de {janela} sem aparecer.")
        else:
            frase = (f"O erro \"{msg}\" está em alta: {_pedidos(l['Pedidos'])} {periodo}, "
                     f"{l['Comparado a antes']}.")
        if len(alta) > 1:
            frase += f" Mais {len(alta) - 1} em alta."
        linhas.append((l["Pedidos"], frase))
    novos = atencao[atencao["Sinal"] == "Novo"]
    if len(novos):
        l = novos.iloc[0]
        msg = resumir_mensagem(l["Mensagem de erro"], 70)
        quantos = _n(len(novos), "erro novo", "erros novos")
        linhas.append((l["Pedidos"],
                       f"{quantos} com {limites.min_ocorrencias}+ ocorrências"
                       + (f": \"{msg}\" ({_int(l['Pedidos'])})." if len(novos) == 1 else
                          f"; o maior é \"{msg}\" ({_int(l['Pedidos'])}).")))
    return [frase for _, frase in sorted(linhas, key=lambda x: -x[0])]


def _textos_defeitos(c):
    """As linhas dos defeitos: novos ofensores e entregas adiadas, depois o que fazer."""
    d = c.defeitos
    if not len(d):
        return [], []
    info, acoes = [], []

    def _dfts(tabela, quantos=4):
        nomes = list(tabela["Defeito"].head(quantos))
        resto = len(tabela) - len(nomes)
        return _lista(nomes + ([f"mais {resto}"] if resto else []))

    # Pela entrada na tabela, e não pelo Sinal: um novo ofensor fora do Octane
    # leva o sinal do Octane, e continua sendo novo ofensor.
    novos_of = (d[d["Ofensor agora"] & ~d["Ofensor antes"] & (d["No relatório agora"] > 0)]
                .sort_values("No relatório agora", ascending=False))
    if len(novos_of):
        partes = [f"{l['Defeito']} ({_pedidos(l['No relatório agora']) if i == 0 else _int(l['No relatório agora'])})"
                  for i, (_, l) in enumerate(novos_of.head(2).iterrows())]
        if len(novos_of) > 2:
            partes.append(f"mais {len(novos_of) - 2}")
        info.append(f"{'Novo ofensor' if len(novos_of) == 1 else 'Novos ofensores'}: "
                    f"{_lista(partes)}.")
    adiadas = d[d["Sinal"] == "Entrega adiada"]
    if len(adiadas):
        info.append(f"{'Entrega adiada' if len(adiadas) == 1 else 'Entregas adiadas'}: " + _lista(
            [f"{l['Defeito']} ({l['Detalhe do sinal']})" for _, l in adiadas.head(3).iterrows()]
            + ([f"mais {len(adiadas) - 3}"] if len(adiadas) > 3 else [])) + ".")

    nao_pegou = d[d["Sinal"] == "Correção não pegou"]
    if len(nao_pegou) == 1:
        l = nao_pegou.iloc[0]
        acoes.append(f"A fazer: rever {l['Defeito']} ({nome_curto(l['Nome'], 45)}), corrigido "
                     f"em {_dia(l['Entrega agora'])}, que teve "
                     f"{_pedidos(l['Pós-entrega na semana'])} depois da entrega.")
    elif len(nao_pegou):
        acoes.append(f"A fazer: rever {_dfts(nao_pegou)}: corrigidos, mas com pedidos depois "
                     f"da entrega.")
    passou = d[d["Sinal"] == "Passou da data"]
    if len(passou) == 1:
        l = passou.iloc[0]
        acoes.append(f"A fazer: cobrar {l['Defeito']}: a entrega era {_dia(l['Entrega agora'])} "
                     f"e o erro continua ({_pedidos(l['Pós-entrega na semana'])} depois).")
    elif len(passou):
        acoes.append(f"A fazer: cobrar {_dfts(passou)}: a data de entrega passou e o erro continua.")
    fora = d[d["Sinal"] == "Fora do Octane"].sort_values("Pedidos agora", ascending=False)
    if len(fora):
        acoes.append(f"A fazer: incluir {_dfts(fora)} no Octane "
                     f"({_pedidos(fora['Pedidos agora'].sum())} sem fase nem data).")
    return info, acoes


def _textos(c, a, b, limites, info_erros):
    """
    O resumo em texto, pronto para colar num e-mail. A primeira linha é o
    título, com o resultado; depois, poucos tópicos curtos — o que explica,
    os riscos novos, o que fazer ("A fazer: …") —, cada um saído das tabelas
    da própria comparação, sem conta nova; por fim, uma "Ressalva: …" com o
    que vale saber para ler os números. O que as tabelas já mostram por
    inteiro fica nelas.
    """
    nome = fallout_core.nome_exibicao(c.visao)
    linhas = []

    # Título. Com o relatório virando de mês, vem antes a virada — é o número
    # que se acompanha —, e depois o mês que as duas fotografias têm.
    rel = c.contexto.get("relatorio", {})
    if rel.get("virou") and rel["taxa_antes"] is not None and rel["taxa_agora"] is not None:
        dif = rel["taxa_agora"] - rel["taxa_antes"]
        linhas.append(
            f"{nome}: o mês do relatório virou — {_rotulo(rel['mes_antes'])} estava em "
            f"{_pct(rel['taxa_antes'])} e {_rotulo(rel['mes_agora'])} está em "
            f"{_pct(rel['taxa_agora'])} ({_pp(dif)}), com "
            f"{_n(rel['dias_agora'], 'dia', 'dias')} de dados.")
    info_mes = c.contexto.get("mes", {})
    ta, tb = info_mes.get("taxa_antes"), info_mes.get("taxa_agora")
    if c.mes is None:
        linhas.append(f"{nome}: as duas fotografias não têm mês em comum para comparar.")
    elif ta is None or tb is None:
        linhas.append(f"{nome}: fallout de {_rotulo(c.mes)} em {_pct(tb)}.")
    else:
        dif = tb - ta
        if abs(dif) < ESTAVEL:
            estado = f"ficou estável em {_pct(tb)}"
        else:
            estado = f"{'subiu' if dif > 0 else 'caiu'} de {_pct(ta)} para {_pct(tb)}"
        linhas.append(f"{nome}: fallout de {_rotulo(c.mes)} {estado} ({_pp(dif)}).")

    # O que explica: a categoria que mais puxou no sentido do total
    dist = c.distribuicao
    if c.mes is not None and "Variação (pp)" in dist and len(dist):
        dif = (tb - ta) if ta is not None and tb is not None else dist["Variação (pp)"].sum()
        sentido = 1 if dif >= 0 else -1
        # Com o total estável, vale o maior movimento de qualquer lado.
        mesmo_lado = dist[(dist["Variação (pp)"].abs() if abs(dif) < ESTAVEL
                           else dist["Variação (pp)"] * sentido) >= ESTAVEL]
        if len(mesmo_lado):
            l = mesmo_lado.loc[mesmo_lado["Variação (pp)"].abs().idxmax()]
            explica = l["O que explica"]
            if abs(dif) < ESTAVEL:
                linhas.append(f"Maior movimento: {l['Categoria']} ({_pp(l['Variação (pp)'])})"
                              + (f" — {explica}." if explica else "."))
            else:
                linhas.append(f"A {'alta' if sentido > 0 else 'queda'} veio de {l['Categoria']} "
                              f"({_pp(l['Variação (pp)'])})" + (f": {explica}." if explica else "."))

    projecao = _texto_projecao(c, a, b)
    # As notas da projeção ("não projetam nenhum mês em comum") ficam na aba
    # Projeção: numa jornada sem projeção, repetidas toda semana, seriam ruído.
    if projecao:
        linhas.append(projecao)
    linhas += _textos_erros(c, limites, info_erros)
    info, acoes = _textos_defeitos(c)
    linhas += info + acoes

    ressalvas = c.contexto["ressalvas"]
    if ressalvas:
        linhas.append("Ressalva: " + "; ".join(ressalvas) + ".")
    return linhas


def _distribuicao_do_relatorio(a, b):
    """
    A Distribuição que cada fotografia reportava — a do slide e da aba
    Distribuição do app, no mês do relatório de cada uma —, lado a lado.

    Quando o relatório virou de mês (set/26 na anterior, out/26 na atual), é ela
    que bate com o app; a de `comparar_distribuicao` é a de um mês que as duas
    têm, e mostrava 1,34% em Tratamento contra 1,09% na aba. Meses diferentes
    não têm os mesmos pedidos, então aqui não há migração: só os dois números e
    a diferença.
    """
    def _itens(lado):
        return {rotulo: valor for rotulo, valor, _sub, _neg in lado.rel["dist_items"]
                if valor is not None}

    antes, agora = _itens(a), _itens(b)
    # A ordem é a da fotografia atual (decrescente, subgrupos sob a categoria),
    # e o que só existia na anterior vem no fim.
    ordem = [r for r, *_ in b.rel["dist_items"]] + [r for r in antes if r not in agora]
    # O nível do subgrupo (1 = US/DFT, 2 = c/ e s/ data); a fotografia atual
    # ganha da anterior quando as duas têm o rótulo.
    nivel = {r: int(sub) for r, _v, sub, _n in a.rel["dist_items"] + b.rel["dist_items"]}
    linhas = []
    for rotulo in dict.fromkeys(ordem):
        va, vb = antes.get(rotulo, 0.0), agora.get(rotulo, 0.0)
        if round(va, 2) == 0 and round(vb, 2) == 0:
            continue
        linhas.append({"Categoria": rotulo, "Subgrupo": nivel.get(rotulo, 0) > 0,
                       "Nível": nivel.get(rotulo, 0),
                       "% antes": va, "% agora": vb, "Diferença (pp)": vb - va})
    return pd.DataFrame(linhas, columns=["Categoria", "Subgrupo", "Nível", "% antes",
                                         "% agora", "Diferença (pp)"])


# ── Comparação completa ───────────────────────────────────────────────────────
def _dias(lado, mes):
    """Dias distintos com falha no mês, na fotografia."""
    return int(lado.pedidos.loc[lado.pedidos["Mes"] == mes, "Dia"].nunique())


def _maduro(lado, mes):
    """
    Mês com corpo na fotografia: fechado pelo histórico, ou com pelo menos
    `fallout_core.MIN_DIAS_MES` dias de dados. Na fotografia de 02/09, setembro
    tinha dois dias de falhas e nenhum sucesso, e saía com 100% de fallout. É
    só para escolher o mês da distribuição comparada: o mês do relatório já não
    espera esses dias (ver `fallout_core.mes_do_relatorio`).
    """
    return lado.fonte(mes) == "histórico" or _dias(lado, mes) >= fallout_core.MIN_DIAS_MES


def _comparavel(a, b, mes, estrito=True):
    """
    True se as duas fotografias têm o mês com volume e com pedidos — e, no
    modo estrito, com corpo nas duas (ver `_maduro`).
    """
    def _tem(lado):
        return (lado.total(mes) is not None and (lado.pedidos["Mes"] == mes).any()
                and (not estrito or _maduro(lado, mes)))
    return mes is not None and _tem(a) and _tem(b)


def _motivo_incomparavel(a, b, mes):
    """Por que o mês não entra na comparação, ou None se entra."""
    rotulo = _rotulo(mes)
    for lado, qual in ((a, "anterior"), (b, "atual")):
        if lado.total(mes) is None or not (lado.pedidos["Mes"] == mes).any():
            return f"A fotografia {qual} não tem {rotulo}"
        if not _maduro(lado, mes):
            dias = _dias(lado, mes)
            return (f"A fotografia {qual} tinha só {dias} {'dia' if dias == 1 else 'dias'} "
                    f"de {rotulo}")
    return None


def _mes_padrao(a, b):
    """
    O mês da distribuição: o do relatório atual, se a anterior também o tem com
    corpo; senão o do relatório anterior, se a atual o tem; senão o mais
    recente que as duas têm. Comparar o mês do relatório atual com uma anterior
    que ainda não o tinha — ou tinha dois dias dele — mostraria todo pedido como
    novo. Jornada pequena demais para ter mês com corpo (a PME Fixa tem meses
    de duas falhas) cai na mesma ordem, sem a exigência de corpo.
    """
    comuns = sorted(set(a.resumo.index) & set(b.resumo.index), reverse=True)
    for estrito in (True, False):
        for mes in (b.rel["mes_atual"], a.rel["mes_atual"], *comuns):
            if _comparavel(a, b, mes, estrito):
                return int(mes)
    return None


def comparar(base, atual, visao, mes=None, limites=None):
    """
    Compara a visão `visao` entre a fotografia `base` (anterior) e `atual`.

    `mes` é o mês da distribuição e dos defeitos (chave AAAAMM); por padrão,
    o do relatório atual (ver `_mes_padrao`). A projeção compara os meses
    projetados que as duas têm, e os erros olham o período entre as duas
    datas de corte.
    """
    limites = limites or Limites()
    a, b = _lado(base, visao), _lado(atual, visao)
    if a is None or b is None:
        falta = "anterior" if a is None else "atual"
        raise ValueError(f"{fallout_core.nome_exibicao(visao)} não está na fotografia "
                         f"{falta}: a jornada não existia, ou não pôde ser lida, quando ela "
                         f"foi gravada.")

    # O mês pedido (o do seletor do app) ou, sem pedido, o do relatório atual.
    # Se ele não tem corpo nas duas, vale o padrão, e o aviso diz por quê.
    aviso_mes = ressalva_mes = None
    padrao = _mes_padrao(a, b)
    alvo = int(mes) if mes is not None else int(b.rel["mes_atual"])
    motivo = _motivo_incomparavel(a, b, alvo)
    if motivo is None:
        mes = alvo
    elif padrao is not None and padrao != alvo:
        mes = padrao
        aviso_mes = f"{motivo}: a comparação mostra {_rotulo(padrao)}."
        ressalva_mes = f"a comparação é de {_rotulo(padrao)} porque {motivo[0].lower()}{motivo[1:]}"
    elif _comparavel(a, b, alvo, estrito=False):
        mes = alvo
        aviso_mes = f"{motivo}: a variação desse mês é pouco confiável."
        ressalva_mes = f"{motivo[0].lower()}{motivo[1:]}, então a variação dele é pouco confiável"
    else:
        mes = padrao

    contexto = _contexto(base, atual, a, b, mes, limites)
    if aviso_mes:
        contexto["avisos"].append(aviso_mes)
        contexto["ressalvas"].insert(0, ressalva_mes)
    # Os números do mês comparado, para quem mostra a comparação não ter de
    # abrir as fotografias de novo.
    contexto["mes"] = {} if mes is None else {
        "taxa_antes": a.taxa(mes), "taxa_agora": b.taxa(mes),
        "total_antes": a.total(mes), "total_agora": b.total(mes),
        "fonte_antes": a.fonte(mes), "fonte_agora": b.fonte(mes),
    }
    # O fallout que cada fotografia reportava, mês do relatório contra mês do
    # relatório. A distribuição precisa de um mês que as duas tenham, mas quando
    # o relatório virou de mês é esta a variação que se acompanha de uma semana
    # para a outra: set/26 fechado em 3,24% → out/26 em 2,53%.
    contexto["relatorio"] = {
        "mes_antes": int(a.rel["mes_atual"]), "taxa_antes": a.rel["fallout_pct"],
        "mes_agora": int(b.rel["mes_atual"]), "taxa_agora": b.rel["fallout_pct"],
        "dias_agora": _dias(b, int(b.rel["mes_atual"])),
        "virou": int(b.rel["mes_atual"]) > int(a.rel["mes_atual"]),
        "distribuicao": _distribuicao_do_relatorio(a, b),
    }
    verificacoes = []
    if mes is not None:
        distribuicao, migracao, migrados, v = comparar_distribuicao(a, b, mes)
        verificacoes += v
    else:
        distribuicao = migracao = migrados = pd.DataFrame()
    projecao, v = comparar_projecao(a, b)
    verificacoes += v
    defeitos = comparar_defeitos(a, b, mes, limites) if mes is not None else pd.DataFrame()
    novos, em_alta, sumidos, info_erros = comparar_erros(a, b, mes, limites)

    c = Comparacao(visao=visao, base_id=base.id, atual_id=atual.id, mes=mes,
                   contexto=contexto, distribuicao=distribuicao, migracao=migracao,
                   pedidos_migrados=migrados, projecao=projecao, defeitos=defeitos,
                   erros_novos=novos, erros_em_alta=em_alta, erros_sumidos=sumidos,
                   texto=[], verificacoes=verificacoes, limites=limites)
    c.contexto["periodo_erros"] = info_erros
    c.texto = _textos(c, a, b, limites, info_erros)
    return c


def resumo_semana(base, atual, limites=None):
    """
    Uma linha por visão: o fallout do mês do relatório atual e a projeção do
    mês seguinte, antes e agora, e quantos alertas cada uma tem.

    Devolve (tabela, {visão: Comparacao}) — as comparações já calculadas, para
    quem mostra o resumo não refazer tudo ao abrir uma visão.
    """
    limites = limites or Limites()
    linhas, comparacoes = [], {}
    for visao in atual.meta["visoes"]:
        nome = fallout_core.nome_exibicao(visao)
        if visao not in base.meta["relatorios"]:
            linhas.append({"Visão": nome, "Observação": "sem fotografia anterior desta visão"})
            continue
        c = comparar(base, atual, visao, limites=limites)
        comparacoes[visao] = c
        a, b = _lado(base, visao), _lado(atual, visao)
        mes = b.rel["mes_atual"]
        # A taxa da anterior só entra se o mês tinha corpo nela: dois dias de
        # setembro sem nenhum sucesso davam 100%, e a variação saía -97 pp.
        ta = a.taxa(mes) if a.taxa(mes) is not None and _maduro(a, mes) else None
        tb = b.taxa(mes)
        obs = []
        if mes < a.rel["mes_atual"]:
            obs.append(f"o mês do relatório voltou de {_rotulo(a.rel['mes_atual'])} para "
                       f"{_rotulo(mes)}")
        elif ta is None:
            tinha = ("não tinha" if a.taxa(mes) is None
                     else f"tinha só {_dias(a, mes)} dias de")
            obs.append(f"a anterior {tinha} {_rotulo(mes)} (o relatório dela era "
                       f"{_rotulo(a.rel['mes_atual'])}, {_pct(a.rel['fallout_pct'])})")
        proj = c.projecao["meses"]
        prox = proj.iloc[0] if len(proj) else None
        qualidade = c.contexto["qualidade"]
        suspeitos = qualidade[qualidade["Suspeito"]]
        if len(suspeitos):
            forte = " (queda forte)" if suspeitos["Queda forte"].any() else ""
            obs.append(f"{len(suspeitos)} {'mês perdeu' if len(suspeitos) == 1 else 'meses perderam'} "
                       f"falhas{forte}")
        linhas.append({
            "Visão": nome,
            "Mês": _rotulo(mes),
            "Fallout antes": ta, "Fallout agora": tb,
            "Variação (pp)": None if ta is None or tb is None else tb - ta,
            "Projeção": None if prox is None else prox["Mês"],
            "Projeção antes": None if prox is None else prox["Antes"],
            "Projeção agora": None if prox is None else prox["Agora"],
            "Variação da projeção (pp)": None if prox is None else prox["Variação (pp)"],
            "Erros novos": len(c.erros_novos),
            "Erros em alta": len(c.erros_em_alta),
            "Pós-entrega": int((c.defeitos["Pós-entrega na semana"] >= limites.min_pos_entrega).sum())
                           if len(c.defeitos) else 0,
            "Observação": "; ".join(obs),
        })
    return pd.DataFrame(linhas), comparacoes


# ── Excel ─────────────────────────────────────────────────────────────────────
def _para_planilha(tabela):
    """
    Uma tabela pronta para o openpyxl. Data com fuso vira data sem fuso (o
    Excel não tem fuso), os tipos que aceitam vazio do pandas viram tipos
    comuns, e o texto perde os caracteres de controle que o XML não aceita —
    uma mensagem de erro com um deles derrubaria o arquivo inteiro.
    """
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    tabela = tabela.copy()
    for col in tabela.columns:
        serie = tabela[col]
        if isinstance(serie.dtype, pd.DatetimeTZDtype):
            tabela[col] = serie.dt.tz_localize(None)
        elif serie.dtype == object or pd.api.types.is_extension_array_dtype(serie.dtype):
            tabela[col] = [
                None if v is None or v is pd.NA or (isinstance(v, float) and v != v)
                else ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v
                for v in serie.astype(object)]
    return tabela


def planilha(c, resumo=None):
    """
    O Excel da comparação: uma aba por seção, na ordem da aba do app, mais
    "Sobre" com o que foi comparado e com que limites. `resumo` é a tabela de
    `resumo_semana`, quando quem chama já a tem. Devolve os bytes do .xlsx.
    """
    import io

    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    limites = c.limites or Limites()
    ctx = c.contexto
    certas = sum(v["OK"] for v in c.verificacoes)
    sobre = pd.DataFrame([
        ("Visão", fallout_core.nome_exibicao(c.visao)),
        ("Mês da distribuição", _rotulo(c.mes) if c.mes else "—"),
        ("Fotografia anterior", f"dados até {ctx['corte_antes']:%d/%m/%Y} — {ctx['origem_antes']}"),
        ("Fotografia atual", f"dados até {ctx['corte_agora']:%d/%m/%Y} — {ctx['origem_agora']}"),
        ("Erro significativo", f"a partir de {limites.min_ocorrencias} ocorrências"),
        ("Erro em alta", f"média diária de {limites.fator_alta:g} vezes a dos "
                         f"{limites.dias_referencia} dias anteriores".replace(".", ",")),
        ("Conferências", f"{certas} de {len(c.verificacoes)} certas"),
    ], columns=["Campo", "Valor"])

    qualidade = ctx["qualidade"].drop(columns="Mes")
    qualidade["Jornada"] = qualidade["Jornada"].map(fallout_core.nome_exibicao)
    # Movimentos entre categorias e arquivos saíram a pedido (22/09/2026), e o
    # app deixou de passar o resumo de todas as jornadas.
    abas = [("Sobre", sobre), ("O que mudou", pd.DataFrame({"O que mudou": c.texto}))]
    if resumo is not None:
        abas.append(("Resumo da semana", resumo))
    abas += [
        ("Distribuição", c.distribuicao),
        ("Projeção", c.projecao["meses"]),
        ("Projeção por item", c.projecao["itens"]),
        ("Defeitos", c.defeitos.drop(columns="DefectKey", errors="ignore")),
        ("Erros novos", c.erros_novos),
        ("Erros em alta", c.erros_em_alta),
        ("Erros que sumiram", c.erros_sumidos),
        ("Qualidade dos dados", qualidade),
        ("Conferências", pd.DataFrame(c.verificacoes)),
    ]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as escritor:
        for nome, tabela in abas:
            if not len(tabela.columns) or not len(tabela):
                tabela = pd.DataFrame({"": ["(nada nesta comparação)"]})
            tabela = _para_planilha(tabela)
            tabela.to_excel(escritor, sheet_name=nome, index=False)
            aba = escritor.sheets[nome]
            aba.freeze_panes = "A2"
            for i, col in enumerate(tabela.columns, start=1):
                aba.cell(row=1, column=i).font = Font(bold=True)
                amostra = [len(str(col))] + [len(str(v)) for v in tabela[col].head(300)
                                             if v is not None]
                aba.column_dimensions[get_column_letter(i)].width = min(max(amostra) + 2, 70)
                if pd.api.types.is_float_dtype(tabela[col]):
                    for linha in range(2, len(tabela) + 2):
                        aba.cell(row=linha, column=i).number_format = "0.00"
    return buf.getvalue()


# ── Linha de comando ──────────────────────────────────────────────────────────
def main(argv=None):
    import argparse
    import os

    parser = argparse.ArgumentParser(description="Compara duas fotografias de extração.")
    parser.add_argument("--base", help="fotografia anterior (AAAA-MM-DD); padrão: a de ~7 dias antes")
    parser.add_argument("--atual", help="fotografia atual (AAAA-MM-DD); padrão: a mais recente")
    parser.add_argument("--visao", default="Consolidado",
                        help='visão detalhada (nome da pasta, "Base + Cross Sell" ou "Consolidado")')
    args = parser.parse_args(argv)
    base_dir = os.path.dirname(os.path.abspath(__file__))

    fotos = fotografia.listar(base_dir)
    if not fotos:
        print("Nenhuma fotografia gravada.")
        return 1
    atual_id = args.atual or fotos[0]["id"]
    base_id = args.base or escolher_base(fotos, atual_id)
    if base_id is None:
        print(f"Não há fotografia anterior a {atual_id} para comparar.")
        return 1
    base = fotografia.carregar(base_dir, base_id)
    atual = fotografia.carregar(base_dir, atual_id)

    tabela, comparacoes = resumo_semana(base, atual)
    print(f"Comparação: dados até {_dt.date.fromisoformat(base_id):%d/%m/%Y} → "
          f"dados até {_dt.date.fromisoformat(atual_id):%d/%m/%Y}\n")
    with pd.option_context("display.width", 250, "display.max_columns", 20,
                           "display.float_format", lambda v: f"{v:,.2f}"):
        print(tabela.drop(columns="Observação").to_string(index=False))
    for _, l in tabela[tabela["Observação"].fillna("") != ""].iterrows():
        print(f"  {l['Visão']}: {l['Observação']}")

    c = comparacoes.get(args.visao)
    if c is None:
        print(f"\nA visão {args.visao} não está nas duas fotografias.")
        return 1
    print(f"\n── {fallout_core.nome_exibicao(args.visao)} ──")
    for linha in c.texto:
        print(f"• {linha}")

    falhas = [v for c_ in comparacoes.values() for v in c_.verificacoes if not v["OK"]]
    total = sum(len(c_.verificacoes) for c_ in comparacoes.values())
    print(f"\nConferências: {total - len(falhas)} de {total} ok")
    for v in falhas:
        print(f"  FALHOU: {v['Verificação']} — {v['Detalhe']}")
    return 0 if not falhas else 2


if __name__ == "__main__":
    sys.exit(main())
