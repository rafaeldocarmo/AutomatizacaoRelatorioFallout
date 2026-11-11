import sys

# No Windows o console/pipe usa cp1252 e quebra nos prints com "→"/"✓".
# Força UTF-8 na saída (rodando direto, via subprocess ou pelo app).
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import calendar as _cal
import datetime as _dt

import pandas as pd
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
import matplotlib.ticker as mticker

import fallout_core
from fallout_core import FASES_CORRIGIDO, FASES_MOPS, MESES_LABEL


# Janela do gráfico no slide de jornada única. O histórico conta o mês
# corrente: 5 significa os quatro anteriores mais ele.
MESES_NO_GRAFICO = 5


# ── Desenho das linhas do gráfico de fallout ──────────────────────────────────
# As três versões do gráfico (slide da jornada, slide de 3 colunas e aba
# Projeção do app) chamam estas duas funções, para as regras de desenho saírem
# iguais nas três. `xs` pode ser rótulo de mês (eixo categórico) ou posição
# numérica (o slide de 3 colunas desenha por índice).

# Os rótulos de % ficam logo acima do ponto, e a linha que sobe ou desce dali
# passava por cima do número ("12,21%" riscado no Cross Sell). O fundo branco,
# à frente das linhas, deixa o número legível sem mudar a posição dele.
# Projeção sem o Legado Aberto (Base Residencial). Bem mais escura que o verde
# da projeção (#2E7D32): com #1B5E20 as duas linhas se confundiam no slide.
COR_VERDE_ESCURO = "#0B3D17"

ROTULO_SOBRE_LINHA = dict(
    bbox=dict(boxstyle="round,pad=0.12", fc="white", ec="none", alpha=0.85),
    zorder=6,
)


def desenhar_real(ax, xs, ys, linewidth=1.5, markersize=4, zorder=3, fechado=False):
    """
    Fallout realizado: contínuo até o mês passado e pontilhado do mês passado
    para o atual. O mês do relatório ainda está correndo e o número dele muda
    até fechar — o pontilhado é o aviso de que aquele último trecho não é
    definitivo.

    `fechado` (o `mes_fechado` do relatório) é o slide de fechamento: o corte
    chegou ao último dia do mês, o número dele não muda mais e a linha é
    contínua até o fim.
    """
    xs, ys = list(xs), list(ys)
    comum = dict(color="black", linewidth=linewidth, marker="o",
                 markersize=markersize, zorder=zorder)
    if len(xs) >= 2 and not fechado:
        ax.plot(xs[:-1], ys[:-1], label="% Falha Real", **comum)
        ax.plot(xs[-2:], ys[-2:], linestyle=":", label="Mês em andamento", **comum)
    elif xs:
        ax.plot(xs, ys, label="% Falha Real", **comum)


def desenhar_verde(ax, xs_hist, plan_hist, xs_proj, vals_proj,
                   linewidth=1.5, markersize=4, zorder=2):
    """
    Uma linha verde só: o Planejado nos meses que já passaram e a projeção dali
    em diante, como continuação dele.

    A emenda pode dar salto — o Planejado é meta digitada, e a projeção sai do
    fallout realizado. No Cross Sell de set/26 a linha subia de 2,28% para
    16,20% na passagem para out/26. A escolha de uma linha contínua foi
    deliberada: o verde é "para onde vamos", e o salto mostra a distância entre
    o que foi prometido e o que a correção dos defeitos de fato entrega.

    `xs_proj`/`vals_proj` trazem o ponto de partida (o mês do relatório) na
    primeira posição, que aqui é descartado: nesse mês quem vale é o Planejado.
    """
    xs = list(xs_hist) + list(xs_proj)[1:]
    ys = list(plan_hist) + list(vals_proj)[1:]
    if not any(v == v for v in ys):             # nem meta nem projeção
        return
    ax.plot(xs, ys, color=fallout_core.COR_VERDE, linewidth=linewidth,
            marker="o", markersize=markersize, zorder=zorder,
            label="Planejado / Projeção")


def desenhar_sem_legado(ax, xs_hist, plan_hist, xs_proj, vals_proj,
                        fontsize=7, linewidth=1.5, markersize=4, zorder=2):
    """
    A linha verde-escura do Base Residencial, "Projeção s/ Legado Aberto" —
    desenhada pelo slide da jornada e pela aba Projeção do app.

    Como a verde, é uma linha só: o histórico dela (a coluna PlanejadoSemLegado
    do histórico, `plan_hist`) nos meses passados e, do mês do relatório em
    diante, a projeção sem o Legado Aberto (`vals_proj`, com o ponto de
    partida na primeira posição) — no mês do relatório vale o digitado, se
    houver. Os rótulos ficam embaixo dos pontos, para não encostar nos da
    projeção, que ficam em cima. Sem `vals_proj` (as outras jornadas), nada.
    """
    if not vals_proj:
        return
    hist = list(plan_hist) if plan_hist else [float("nan")] * len(xs_hist)
    if hist[-1] != hist[-1]:
        hist[-1] = vals_proj[0]
    ax.plot(list(xs_hist) + list(xs_proj)[1:], hist + list(vals_proj)[1:],
            color=COR_VERDE_ESCURO, linewidth=linewidth, marker="o",
            markersize=markersize, zorder=zorder, label="Projeção s/ Legado Aberto")
    for x, v in zip(xs_proj, [hist[-1]] + list(vals_proj)[1:]):
        ax.annotate(f"{v:.2f}%".replace(".", ","), xy=(x, v), xytext=(0, -11),
                    textcoords="offset points", ha="center", va="top",
                    fontsize=fontsize, color=COR_VERDE_ESCURO, **ROTULO_SOBRE_LINHA)


def tabela_ofensores(df, meses, hoje):
    """
    Pedidos em aberto de cada defeito, um mês por coluna (`meses`): a tabela de
    ofensores do slide (o `pivot2`), sem o filtro e a ordem que o relatório
    aplica depois.

    Função à parte para o Top Ofensores contar cada jornada nos meses da visão:
    com o Base + Cross Sell já em out/26 (set/26 fechado no histórico) e o Base
    Móvel ainda em set/26, o `pivot2` da jornada não tinha outubro, e o slide
    saía com 0,00% em todo defeito.
    """
    hoje = fallout_core.hoje_utc(hoje)
    ms = pd.to_datetime(df["DFT_BugfixMilestone"], utc=True, errors="coerce")
    fechado = (
        df["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO) &
        df["DFT_BugfixMilestone"].notna() &
        (ms <= hoje)
    )
    # `TemDefeito` no lugar do teste pelo número: jornadas com catálogo próprio
    # identificam o defeito por rótulo (BRJ-218, "Paliativo RPA - 6") e ficariam
    # de fora da tabela se o filtro exigisse um número.
    # O Legado Aberto e a Funcionalidade idêntica ao legado, do Base
    # Residencial, também saem: o pedido não está em tratamento, mesmo com
    # defeito associado (ver fallout_core.CAT_ERRO_LEGADO). E o defeito
    # Cancelado ou Rejeitado, como no Top Ofensores: ninguém está tratando, e
    # o pedido está em avaliação por MOPs.
    fora = df["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS)
    for marca in ("LegadoAberto", "ErroLegado"):
        if marca in df.columns:
            fora |= df[marca].fillna(False).astype(bool)
    em_aberto = df[df["Mes"].isin(meses) & df["TemDefeito"] & ~fechado & ~fora]

    idx = ["DefectKey", "DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Type", "DFT_Team"]
    tabela = (
        em_aberto.groupby(idx + ["Mes"], dropna=False)
        .size()
        .reset_index(name="n")
        .set_index(idx + ["Mes"])["n"]
        .unstack("Mes", fill_value=0)
        .reset_index()
    )
    for m in meses:
        if m not in tabela.columns:
            tabela[m] = 0
    return tabela


def _em_ordem_decrescente(itens, nivel=0):
    """
    A Distribuição da maior categoria para a menor, em todos os slides. Os
    subgrupos (`subitem` > 0, o nível, logo abaixo do item-mãe na lista) andam
    junto com ele e também ficam do maior para o menor entre si, em cada nível
    — US e DFT sob o Em Tratamento, e o c/ data e s/ data sob cada um. No
    empate vale a ordem da lista.
    """
    blocos = []                                 # (item, [itens abaixo dele])
    for item in itens:
        if int(item[2]) > nivel and blocos:
            blocos[-1][1].append(item)
        else:
            blocos.append((item, []))
    blocos.sort(key=lambda bloco: -bloco[0][1])
    return [linha for item, abaixo in blocos
            for linha in [item] + _em_ordem_decrescente(abaixo, nivel + 1)]


def gerar_relatorio(base_dir, jornada):
    """
    Calcula o fallout de uma jornada e monta o dashboard consolidado (figura
    matplotlib) e as tabelas usadas pelos geradores de PPTX/PNG.

    Recebe `base_dir` e `jornada` explicitamente (sem depender de cwd/sys.argv
    globais) e usa a API orientada a objetos do matplotlib (Figure() própria,
    não o estado global do pyplot) — pode ser chamada com segurança a partir
    de múltiplas sessões/threads simultâneas do app Streamlit.

    Retorna um dict com os dados consumidos por gerar_pptx.py e
    gerar_slide3col.py, mais a figura `fig` — o dashboard que vira PNG/slide.

    É `calcular_relatorio` seguido de `desenhar_dashboard`. Quem só precisa dos
    números (o app, o Excel de ofensores, as colunas do slide de 3 colunas) chama
    `calcular_relatorio` direto: a figura tem milhares de retângulos e custava
    mais de meio segundo por jornada para ninguém usar.
    """
    ns = calcular_relatorio(base_dir, jornada)
    return desenhar_dashboard(ns, base_dir, jornada)


def calcular_relatorio(base_dir, jornada, base=None, hoje=None):
    """
    Os números do relatório de uma jornada, sem a figura: o mesmo dict de
    `gerar_relatorio` menos `fig` e `ax_chart`.

    `base` é o (df, resumo) de `fallout_core.carregar_base` quando quem chama já
    o tem — o app guarda os dois em cache e não precisa ler a base de novo.

    `hoje` é a data de referência das regras (ver `fallout_core.hoje_utc`);
    sem ela, é a data atual. A fotografia de uma extração antiga passa o dia
    daquela extração: a regra dos 15 dias e a entrega que já passou dependem
    dele, e com a data de hoje as categorias sairiam como estariam agora.

    Cada bloco aqui calcula algo que sai no dict, e a figura sai dele em
    `desenhar_dashboard`. Se um dia voltar a existir tabela/gráfico que não
    aparece em nenhum dos dois, ele é trabalho jogado fora: já houve cinco
    figuras assim, todas removidas.
    """
    print(f"Jornada: {jornada}")

    # ── 1. Leitura das bases (falhas + sucessos + Octane) e join ───────────────
    # O Octane já vem embutido no `df` pelo join; a planilha crua não é usada aqui.
    if base is None:
        df, _df_octane, resumo, jornadas_combo = fallout_core.carregar_base(base_dir, jornada,
                                                                             hoje=hoje)
    else:
        df, resumo = base
        jornadas_combo = fallout_core.resolver_jornadas(jornada, base_dir)
    print(f"  Pastas: {', '.join(jornadas_combo)}")
    print(f"Falhas carregadas: {len(df):,} linhas")

    total = len(df)
    com_octane = df["DFT_Phase"].notna().sum()
    sem_octane = df["DFT_Phase"].isna().sum()
    print(f"Total de registros de falha : {total:,}")
    print(f"Com match no Octane         : {com_octane:,} ({com_octane/total:.1%})")
    print(f"Sem match no Octane         : {sem_octane:,} ({sem_octane/total:.1%})")

    # ── 2. Distribuição do fallout do mês + planejamento de redução ─────────────
    hoje = fallout_core.hoje_utc(hoje)
    _agora_brt = hoje.tz_convert("America/Sao_Paulo")
    # O mais recente com dados (ver fallout_core.mes_do_relatorio); sem nenhum,
    # o do calendário. A mesma regra decide o mês em que o app abre.
    mes_atual = fallout_core.mes_do_relatorio(df, resumo)
    if mes_atual is None:
        mes_atual = fallout_core.chave_mes(_agora_brt.year, _agora_brt.month)

    # Mês fechado: o corte chegou ao último dia do mês do relatório ou já passou
    # dele. A linha dele é contínua (ver `desenhar_real`).
    _corte = fallout_core.data_corte(base_dir, hoje)
    _dia_seguinte = _corte + _dt.timedelta(days=1) if _corte is not None else None
    mes_fechado = bool(
        _dia_seguinte is not None
        and fallout_core.chave_mes(_dia_seguinte.year, _dia_seguinte.month) > mes_atual
    )

    df_mes, cats, _total_cat = fallout_core.categorizar(df, mes_atual, hoje=hoje)
    # O denominador das categorias: num mês fechado digitado só com a taxa, o
    # total da extração (ver fallout_core.total_do_mes).
    total_mes = fallout_core.total_do_mes(resumo, mes_atual)
    fallout_pct = resumo.loc[mes_atual, "Pct"]

    def pct(n):
        return n / total_mes * 100

    em_trat = cats["Em Tratamento/Avaliação pela Squad"]
    resolvido = cats["Resolvido"]
    pontual = cats["Falha Pontual"]
    falta_assoc = cats["Falta Associar ao Defeito/US"]
    tratado = cats["Tratado - Em avaliação de eficácia"]
    mops = cats["Em Avaliação por MOPs"]
    bsim = cats[fallout_core.CAT_BSIM]
    legado = cats[fallout_core.CAT_LEGADO]
    _sub_legado = fallout_core.subgrupo_legado(legado)
    legado_manual = legado[_sub_legado == fallout_core.SUB_LEGADO_MANUAL]
    legado_funcionalidade = legado[_sub_legado == fallout_core.SUB_LEGADO_FUNCIONALIDADE]
    outros_times = cats["Em avaliação - Outros times"]
    erro_processo = cats["Falha no Processo Usuário"]
    incidente = cats[fallout_core.CAT_INCIDENTE]
    script_rpa = cats[fallout_core.CAT_SCRIPT_RPA]
    agendamento = cats[fallout_core.CAT_AGENDAMENTO]

    # ── Diagnóstico DFT específico ───────────────────────────────────────────────
    _DFT_DEBUG = 232288
    _rows_dft = df[df["DefectNumber__c"] == _DFT_DEBUG]
    if len(_rows_dft):
        r = _rows_dft.iloc[0]
        _mes_dft = df[df["DefectNumber__c"] == _DFT_DEBUG]["Mes"].unique().tolist()
        print(f"\n[DEBUG DFT {_DFT_DEBUG}]")
        print(f"  Meses em que aparece : {_mes_dft}  (mes_atual={mes_atual})")
        print(f"  Phase                : {r['DFT_Phase']}")
        print(f"  BugfixMilestone      : {r['DFT_BugfixMilestone']}")
        print(f"  Type                 : {r['DFT_Type']}")
        _in_mops = _rows_dft["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS).any()
        _in_enc = (_rows_dft["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO) & _rows_dft["DFT_BugfixMilestone"].notna()).any()
        print(f"  Phase é MOPs (Rejeitado) : {_in_mops}")
        print(f"  Está encerrado (Corrigido+Milestone): {_in_enc}")
        print(f"  Aparece em em_trat   : {(em_trat['DefectNumber__c'] == _DFT_DEBUG).any()}")
    else:
        print(f"\n[DEBUG DFT {_DFT_DEBUG}] NÃO encontrado em nenhuma linha do df")

    _subgrupo = fallout_core.subgrupo_tratamento(em_trat)

    # Planejamento de Redução: DFTs com milestone futuro (pro-rate por dia de entrega)
    _ms_future = pd.to_datetime(em_trat["DFT_BugfixMilestone"], utc=True, errors="coerce")
    plan_futuro = em_trat[_ms_future > hoje].copy()
    # A data de entrega é de calendário (ver fallout_core.dia_local). Passada
    # para UTC e de volta a Brasília, a data sem fuso do export do Octane de
    # 08/10 (33 mil linhas, histórias e defeitos juntos) caía às 21h da véspera,
    # e todo o Planejamento Redução saía um dia antes: 16/10 no lugar de 17/10.
    plan_futuro["MilestoneDate"] = fallout_core.dia_local(
        plan_futuro["DFT_BugfixMilestone"]).dt.tz_localize("America/Sao_Paulo")

    # Agrupa por DefectKey, não por DefectNumber__c: as jornadas com catálogo
    # próprio identificam o defeito por rótulo (BRJ-218 no Base Residencial,
    # "Desenvolvimento RPA-9" no PME) e ficam todas com DefectNumber__c = -1.
    # Pelo número, esses defeitos colapsavam numa linha só, rotulada "DFT-1".
    dft_counts = (
        plan_futuro.groupby(["MilestoneDate", "DefectKey"])
        .size().reset_index(name="n")
    )
    dft_counts["Pct_bruta"] = dft_counts["n"] / total_mes * 100

    def _prorate(ms_date):
        """Fração de dias restantes no mês após a data de entrega."""
        dias_no_mes = _cal.monthrange(ms_date.year, ms_date.month)[1]
        dias_restantes = dias_no_mes - ms_date.day
        return max(dias_restantes, 0) / dias_no_mes

    # Sem nada em tratamento com entrega futura, o .apply() devolve uma série
    # vazia com dtype de data e a multiplicação abaixo estoura.
    dft_counts["ProRate"] = (
        dft_counts["MilestoneDate"].apply(lambda d: _prorate(d.to_pydatetime()))
        if len(dft_counts) else pd.Series(dtype="float64")
    )
    dft_counts["Pct_mes"] = dft_counts["Pct_bruta"] * dft_counts["ProRate"]
    dft_counts["Pct_plena"] = dft_counts["Pct_bruta"]

    def _rotulo_defeito(chave):
        """DFT123456 quando a chave é um número; o rótulo como veio quando não é."""
        try:
            return f"DFT{int(float(chave))}"
        except (TypeError, ValueError):
            return str(chave)

    reducao = (
        dft_counts.groupby("MilestoneDate")
        .agg(
            DFTs=("DefectKey", lambda x: "\n".join(
                _rotulo_defeito(v) for v in sorted(x.unique(), key=str))),
            n=("n", "sum"),
            Pct_mes=("Pct_mes", "sum"),
            Pct_plena=("Pct_plena", "sum"),
        )
        .reset_index()
        .sort_values("MilestoneDate")
    )
    reducao["Pct"] = reducao["Pct_plena"]

    soma = (len(em_trat) + len(resolvido) + len(pontual) + len(falta_assoc)
            + len(tratado) + len(mops) + len(bsim) + len(legado)
            + len(outros_times) + len(erro_processo) + len(incidente)
            + len(script_rpa) + len(agendamento))
    print(f"\nDistribuição Fallout {MESES_LABEL[mes_atual]} ({fallout_pct:.2f}%)")
    print(f"  Em Tratamento          {pct(len(em_trat)):.2f}%  ("
          + " | ".join(f"{s}:{pct((_subgrupo == s).sum()):.2f}%"
                       for subs in fallout_core.SUBGRUPOS_TRATAMENTO.values() for s in subs)
          + ")")
    print(f"  Resolvido              {pct(len(resolvido)):.2f}%")
    print(f"  Falha Pontual          {pct(len(pontual)):.2f}%")
    print(f"  Falta associar         {pct(len(falta_assoc)):.2f}%")
    print(f"  Tratado - eficácia     {pct(len(tratado)):.2f}%")
    print(f"  Em Avaliação MOPs      {pct(len(mops)):.2f}%")
    print(f"  BSIM                   {pct(len(bsim)):.2f}%")
    print(f"  Legado Aberto          {pct(len(legado)):.2f}%  (Manual:{pct(len(legado_manual)):.2f}% | "
          f"Func. idêntica:{pct(len(legado_funcionalidade)):.2f}%)")
    print(f"  Outros times           {pct(len(outros_times)):.2f}%")
    print(f"  Falha Processo Usuário {pct(len(erro_processo)):.2f}%")
    print(f"  Incidente              {pct(len(incidente)):.2f}%")
    print(f"  Total categorizado: {soma}/{len(df_mes)}")

    # ── 3. Ofensores: defeitos/US ainda em aberto no mês (pivot2) ───────────────
    mes_ant2 = fallout_core.mes_anterior(mes_atual)
    meses_tab = [mes_ant2, mes_atual]
    labels_tab = [MESES_LABEL[m][:3].capitalize() for m in meses_tab]
    totais_tab = {m: fallout_core.total_do_mes(resumo, m) for m in meses_tab
                  if m in resumo.index}

    def pct_m(n, m):
        return n / totais_tab.get(m, 1) * 100

    pivot2 = tabela_ofensores(df, meses_tab, hoje)

    _dft_em_trat = set(em_trat["DefectKey"].unique())
    _dft_pivot_antes = set(pivot2["DefectKey"].unique()) if "DefectKey" in pivot2.columns else set()
    _dft_sumidos = _dft_em_trat - _dft_pivot_antes
    if _dft_sumidos:
        print(f"  [DIAG] DFTs em em_trat mas fora do pivot (antes filtro mes_atual): {sorted(_dft_sumidos)}")
    _dft_sem_mes_atual = set(pivot2[pivot2[mes_atual] == 0]["DefectKey"].unique()) if mes_atual in pivot2.columns else set()
    _dft_filtrados = _dft_em_trat & _dft_sem_mes_atual
    if _dft_filtrados:
        print(f"  [DIAG] DFTs removidos por nao ter ocorrencia em mes_atual ({mes_atual}): {sorted(_dft_filtrados)}")

    pivot2 = pivot2[pivot2[mes_atual] > 0].sort_values(mes_atual, ascending=False)

    def quarter_suffix(milestone):
        if pd.isna(milestone):
            return ""
        dt = pd.to_datetime(milestone, errors="coerce")
        if pd.isna(dt):
            return ""
        q = (dt.month - 1) // 3 + 1
        if q == 3:
            return "*"
        if q == 4:
            return "**"
        return ""

    def fmt_id(row):
        # A chave é o número quando existe e o rótulo (BRJ-218) quando não.
        raw = row.get("DefectKey", row.get("DefectNumber__c"))
        try:
            base = str(int(float(raw)))
        except (TypeError, ValueError):
            base = str(raw)
        return base + quarter_suffix(row["DFT_BugfixMilestone"])

    def fmt_milestone(ms):
        dt = pd.to_datetime(ms, errors="coerce")
        return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else ""

    COL_NAMES = ["Classificação", "Agrupamento", "Tipo", "Defeito / US",
                 "Data Prod\nnão resolvidos", "Nome Defeito", "Time",
                 "Status\nDefeito/US"] + labels_tab

    # Data de abertura de cada defeito, onde a origem traz (o catálogo do Base
    # Residencial). Vazio nas demais jornadas, e aí o slide não ganha a coluna.
    if "DFT_CreatedDate" in df.columns:
        _criado = df.dropna(subset=["DFT_CreatedDate"]).drop_duplicates("DefectKey")
        criado_em = dict(zip(_criado["DefectKey"].astype(str), _criado["DFT_CreatedDate"]))
    else:
        criado_em = {}

    # ── 3b. Volume, distribuição e séries do gráfico ──────────────────────────
    # Os números que o dashboard mostra. Ficam aqui, e não junto do desenho,
    # porque o app, o Excel de ofensores e o slide de 3 colunas usam esses
    # valores sem precisar da figura.
    # Também até o mês do relatório: a Distribuição logo abaixo é do mes_atual,
    # e uma coluna de volume à frente dele — com poucos dias de extração — seria
    # lida como se fosse o mesmo mês.
    meses_vol = [m for m in sorted(resumo.index) if m <= mes_atual][-3:]

    # O terceiro campo é o nível: 0 a categoria, 1 o subgrupo, 2 o subgrupo do
    # subgrupo (o c/ data e s/ data dentro de US e DFT).
    _tipo = fallout_core.tipo_do_subgrupo(_subgrupo)
    dist_items = [
        ("Em Tratamento/Avaliação pela Squad", pct(len(em_trat)), 0, True),
        *[item for tipo, subs in fallout_core.SUBGRUPOS_TRATAMENTO.items()
          for item in [(tipo, pct((_tipo == tipo).sum()), 1, False)]
          + [(s, pct((_subgrupo == s).sum()), 2, False) for s in subs]],
        ("Resolvido", pct(len(resolvido)), False, False),
        ("Falha Pontual", pct(len(pontual)), False, False),
        ("Falta associar problema ao Defeito/US", pct(len(falta_assoc)), False, False),
        ("Tratado - Em avaliação de eficácia", pct(len(tratado)), False, False),
        ("Em Avaliação por MOPs", pct(len(mops)), False, False),
        (fallout_core.CAT_BSIM, pct(len(bsim)), False, False),
        # Negrito como o Em Tratamento: as duas são as categorias com subgrupos.
        (fallout_core.CAT_LEGADO, pct(len(legado)), False, True),
        (fallout_core.SUB_LEGADO_MANUAL, pct(len(legado_manual)), True, False),
        (fallout_core.SUB_LEGADO_FUNCIONALIDADE, pct(len(legado_funcionalidade)), True, False),
        ("Em avaliação - Outros times", pct(len(outros_times)), False, False),
        ("Falha no Processo Usuário", pct(len(erro_processo)), False, False),
        (fallout_core.CAT_SCRIPT_RPA, pct(len(script_rpa)), False, False),
        (fallout_core.CAT_AGENDAMENTO, pct(len(agendamento)), False, False),
        (fallout_core.CAT_INCIDENTE, pct(len(incidente)), False, False),
    ]
    dist_items = _em_ordem_decrescente(dist_items)

    # O slide de jornada mostra uma janela curta: com o histórico digitado a
    # série passou de 25 meses e os rótulos viravam um borrão. O arquivo segue
    # guardando tudo — quem mostra a série inteira é o slide de 3 colunas.
    # A série realizada termina no mês do relatório. O resumo pode ter um mês à
    # frente dele — um mês com falhas e ainda sem sucessos, que não vira o do
    # relatório (ver fallout_core.mes_do_relatorio): ele não é realização, e sem
    # este corte entrava na linha preta como se fosse — e ainda era projetado em
    # seguida, com o mesmo rótulo saindo duas vezes no eixo.
    meses_hist = [m for m in sorted(resumo.index) if m <= mes_atual][-MESES_NO_GRAFICO:]
    vals_hist = [resumo.loc[m, "Pct"] for m in meses_hist]
    labels_hist = [MESES_LABEL[m] for m in meses_hist]

    # "Planejado": a meta digitada no histórico. Só existe onde foi preenchida —
    # os buracos ficam NaN e o matplotlib não desenha o trecho. Ele é desenhado
    # mais abaixo, junto com a projeção, porque os dois formam uma linha só.
    plan_hist = [
        float(resumo.loc[m, "PctPlanejado"]) if "PctPlanejado" in resumo.columns
        else float("nan")
        for m in meses_hist
    ]
    # O mesmo para a linha verde-escura do Base Residencial (projeção sem o
    # Legado Aberto): o histórico dela é a coluna PlanejadoSemLegado.
    plan_sem_legado_hist = [
        float(resumo.loc[m, "PctPlanejadoSemLegado"])
        if "PctPlanejadoSemLegado" in resumo.columns else float("nan")
        for m in meses_hist
    ]

    _ms_chave = (fallout_core.chave_de_datas(reducao["MilestoneDate"])
                 if len(reducao) else pd.Series(dtype="int64"))
    # O que está no número do mês corrente mas não se repete no mês que vem sai
    # inteiro já no primeiro mês projetado. São três categorias:
    #   "Resolvido" e "Tratado - Em avaliação de eficácia", cuja correção já
    #   entrou em produção (milestone <= hoje);
    #   "Incidente", que é falha pontual de um incidente já encerrado e não tem
    #   DFT para entrar no planejamento de redução — mantê-lo projetava para
    #   sempre uma falha que não volta.
    # Nenhuma leva pro-rata: ele existe para quem entrega no meio do mês, e aqui
    # não há entrega a esperar.
    #
    # Sem isso a projeção nascia pessimista, tratando como permanente o que já
    # estava corrigido: no Cross Sell de ago/26 eram 5,29 pp (4,99 de Resolvido
    # mais 0,30 de eficácia) sobre um fallout de 21,31%. O incidente é o mesmo
    # caso — em set/26 são 2,04 pp do Cross Sell, e zero nas demais jornadas.
    #
    # Não há dupla contagem com `reducao`: ela sai de `em_trat`, que exclui
    # tanto o que já está encerrado quanto o incidente. As três são disjuntas.
    pct_nao_recorrente = pct(len(resolvido) + len(tratado) + len(incidente))

    proj_months_num = [m for m in sorted(set(_ms_chave)) if m > mes_atual]
    todos_meses_proj = []
    # Vai até um mês depois da última entrega, para a curva mostrar o patamar em
    # que ela fica quando tudo estiver corrigido, e sempre projeta ao menos o mês
    # seguinte, mesmo sem entrega nem nada corrigido: a PME Móvel de out/26 não
    # tinha nenhuma das duas coisas e o gráfico saía sem projeção.
    limite = fallout_core.proximo_mes(
        proj_months_num[-1] if proj_months_num else mes_atual)
    chave = fallout_core.proximo_mes(mes_atual)
    # Sem limite de meses: a curva segue até a última entrega planejada.
    # Antes parava em dois meses, e as entregas mais distantes não
    # apareciam na projeção.
    while chave <= limite:
        todos_meses_proj.append(chave)
        chave = fallout_core.proximo_mes(chave)

    proj_labels = [labels_hist[-1]]
    proj_vals = [vals_hist[-1]]
    pct_acumulada = (reducao.loc[_ms_chave == mes_atual, "Pct_plena"].sum()
                     + pct_nao_recorrente)

    # `proj_detalhe` guarda a conta de cada mês projetado, parcela a parcela, para
    # a aba de Projeção do app mostrar de onde sai o número sem refazer a regra.
    proj_detalhe = []
    for m in todos_meses_proj:
        red_mes = reducao.loc[_ms_chave == m, "Pct_mes"].sum()
        red_plena = reducao.loc[_ms_chave == m, "Pct_plena"].sum()
        cur_val = max(vals_hist[-1] - pct_acumulada - red_mes, 0)
        proj_detalhe.append({
            "mes": MESES_LABEL[m],
            "acumulado": float(pct_acumulada),   # entregas já contadas + não recorrente
            "entrega_mes": float(red_mes),       # entregas deste mês, pro-rata
            "valor": float(cur_val),
        })
        proj_labels.append(MESES_LABEL[m])
        proj_vals.append(cur_val)
        pct_acumulada += red_plena

    # A mesma projeção sem o Legado Aberto, só onde ele existe (o Base
    # Residencial): a fatia dele no mês do relatório sai de todos os pontos,
    # inclusive do de partida. Vazia nas demais jornadas.
    pct_legado = pct(len(legado))
    proj_sem_legado = ([max(v - pct_legado, 0) for v in proj_vals]
                       if round(pct_legado, 2) > 0 else [])

    return {
        "df": df,
        "resumo": resumo,
        "MESES_LABEL": MESES_LABEL,
        "mes_atual": mes_atual,
        "mes_fechado": mes_fechado,
        "hoje": hoje,
        "fallout_pct": fallout_pct,
        "meses_vol": meses_vol,
        "dist_items": dist_items,
        "reducao": reducao,
        "COL_NAMES": COL_NAMES,
        "criado_em": criado_em,
        "meses_tab": meses_tab,
        "pivot2": pivot2,
        "fmt_id": fmt_id,
        "fmt_milestone": fmt_milestone,
        "pct_m": pct_m,
        "proj_labels": proj_labels,
        "proj_vals": proj_vals,
        "proj_sem_legado": proj_sem_legado,
        "plan_sem_legado_hist": plan_sem_legado_hist,
        # Séries e parcelas do gráfico de projeção, para a aba do app desenhar o
        # mesmo gráfico sem repetir a regra — se ela fosse recalculada lá, app e
        # slide poderiam divergir sem ninguém perceber.
        "labels_hist": labels_hist,
        "vals_hist": vals_hist,
        "plan_hist": plan_hist,
        "pct_nao_recorrente": float(pct_nao_recorrente),
        "proj_detalhe": proj_detalhe,
        "total_mes": total_mes,
        # O Planejamento de Redução defeito a defeito, antes de ser agrupado por
        # data em `reducao`. É o que a fotografia guarda para a comparação
        # explicar a mudança da projeção por defeito ("o DFT X foi adiado").
        "reducao_defeitos": dft_counts,
    }


def desenhar_dashboard(ns, base_dir, jornada):
    """
    Dashboard 16x9 de uma jornada — a figura do slide dela — a partir do
    resultado de `calcular_relatorio`. Devolve o mesmo `ns`, agora com `fig` e
    `ax_chart`. Só desenha: todo número que aparece aqui já vem calculado.
    """
    resumo, mes_atual, hoje = ns["resumo"], ns["mes_atual"], ns["hoje"]
    fallout_pct, dist_items, reducao = ns["fallout_pct"], ns["dist_items"], ns["reducao"]
    meses_vol, meses_tab, pivot2 = ns["meses_vol"], ns["meses_tab"], ns["pivot2"]
    COL_NAMES = ns["COL_NAMES"]
    fmt_id, fmt_milestone, pct_m = ns["fmt_id"], ns["fmt_milestone"], ns["pct_m"]
    labels_hist, vals_hist, plan_hist = ns["labels_hist"], ns["vals_hist"], ns["plan_hist"]
    proj_labels, proj_vals = ns["proj_labels"], ns["proj_vals"]

    COL_W = [1.5, 2.8, 1.3, 1.5, 1.8, 6.2, 2.6, 2.2] + [1.3] * len(meses_tab)
    X = [sum(COL_W[:i]) for i in range(len(COL_W) + 1)]

    COR_HDR = fallout_core.COR_VERMELHO
    COR_W2 = "#FFFFFF"; COR_G2 = "#F2F2F2"; COR_FG2 = "#222222"
    CELL_H = 0.32
    HDR_H = 0.38
    n_data = 1 + 1 + len(pivot2)
    fig_h = HDR_H + n_data * CELL_H + 0.45

    # ── 4. Dashboard consolidado — a única figura que sai daqui ─────────────────
    COR_RED = fallout_core.COR_VERMELHO
    COR_W = "#FFFFFF"; COR_G = "#F0F0F0"
    COR_BLUE = "#1565C0"; COR_FG = "#222222"

    fig = Figure(figsize=(16, 9))
    FigureCanvasAgg(fig)
    gs = fig.add_gridspec(3, 2, height_ratios=[0.09, 1, 1.1], width_ratios=[1, 1.2])

    ax_title = fig.add_subplot(gs[0, :])
    ax_left = fig.add_subplot(gs[1, 0])
    ax_chart = fig.add_subplot(gs[1, 1])
    ax_table = fig.add_subplot(gs[2, :])
    for a in (ax_title, ax_left, ax_chart, ax_table):
        a.set_axis_off()

    ax_title.set_xlim(0, 1)
    ax_title.set_ylim(0, 1)
    ax_title.add_patch(_rect(0, 0, 1, 1, facecolor=COR_RED, transform=ax_title.transAxes, clip_on=False))
    ax_title.text(0.5, 0.5, f"{jornada} – Realização e Projeção de Fallout",
                  transform=ax_title.transAxes, ha="center", va="center",
                  fontsize=16, fontweight="bold", color="white")

    # ── 4a. Bloco esquerdo: Volume + Distribuição + Planejamento ──────────────
    labels_vol = [MESES_LABEL[m].capitalize() for m in meses_vol]

    n_planos = len(reducao)
    n_dft_lines = sum(r["DFTs"].count("\n") + 1 for _, r in reducao.iterrows())
    left_rows_n = (2 + 3 + 1 + 1 + 3 + 1 + 4 + 1 + n_planos + n_dft_lines)
    LW, VW = 5.5, 2.0
    left_w = LW + VW
    row_h = 0.72
    left_h = left_rows_n * row_h
    ax_left.set_xlim(0, left_w)
    ax_left.set_ylim(0, left_h)

    def cellL(x0, y0, w, h, text, bg, fg=COR_FG, bold=False, fs=6.5, align="center"):
        ax_left.add_patch(_rect(x0, y0, w, h, facecolor=bg, edgecolor="white", linewidth=0.6))
        ha = "left" if align == "left" else "center"
        xt = x0 + 0.1 if align == "left" else x0 + w / 2
        ax_left.text(xt, y0 + h / 2, text, ha=ha, va="center", fontsize=fs,
                     color=fg, fontweight="bold" if bold else "normal")

    y = left_h
    n_vol = len(meses_vol)
    VOL_LBL_W = left_w * 0.46
    VOL_COL_W = (left_w - VOL_LBL_W) / n_vol
    DIST_LBL_W = left_w * 0.74
    DIST_VAL_W = left_w - DIST_LBL_W

    y -= row_h
    cellL(0, y, VOL_LBL_W, row_h, "", COR_RED)
    cellL(VOL_LBL_W, y, left_w - VOL_LBL_W, row_h, "Volume de Pedidos", COR_RED, COR_W, bold=True, fs=7.5)

    y -= row_h
    cellL(0, y, VOL_LBL_W, row_h, "", COR_RED)
    for i, lbl in enumerate(labels_vol):
        cellL(VOL_LBL_W + i * VOL_COL_W, y, VOL_COL_W, row_h, lbl, COR_RED, COR_W, bold=True, fs=7)

    vol_data = [
        ("Vendas com Sucesso", "Sucessos", fallout_core.fmt_inteiro),
        ("Falha (Análise Técnica)", "Falhas", fallout_core.fmt_inteiro),
        ("Fallout Rate", "Pct", fallout_core.fmt_pct),
    ]
    for ri, (label, key, fmt) in enumerate(vol_data):
        y -= row_h
        bg = COR_W if ri % 2 == 0 else COR_G
        cellL(0, y, VOL_LBL_W, row_h, label, bg, COR_FG, fs=8, align="left")
        for i, m in enumerate(meses_vol):
            cellL(VOL_LBL_W + i * VOL_COL_W, y, VOL_COL_W, row_h, fmt(resumo.loc[m, key]), bg, COR_FG, fs=8)

    y -= row_h
    cellL(0, y, left_w, row_h, f"Distribuição Fallout ({fallout_pct:.2f}%)", COR_RED, COR_W, bold=True, fs=7.5)

    toggle = False
    _tem_nivel_2 = any(int(it[2]) >= 2 for it in dist_items)
    for label, val, subitem, bold_label in dist_items:
        y -= row_h
        bg = COR_W if not toggle else COR_G
        toggle = not toggle
        val_str = "" if fallout_core.so_rotulo(label, subitem) else f"{val:.2f}%".replace(".", ",")
        if subitem:
            # Subgrupo alinhado à direita; o de primeiro nível com um recuo a
            # mais, para o "c/ data"/"s/ data" de baixo ler como parte dele.
            recuo = 0.15 + (0.8 if int(subitem) == 1 and _tem_nivel_2 else 0)
            ax_left.add_patch(_rect(0, y, DIST_LBL_W, row_h, facecolor=bg, edgecolor="white", linewidth=0.6))
            ax_left.text(DIST_LBL_W - recuo, y + row_h / 2,
                         fallout_core.rotulo_dist(label, subitem), ha="right", va="center",
                         fontsize=6.5, color=COR_FG)
            cellL(DIST_LBL_W, y, DIST_VAL_W, row_h, val_str, bg, COR_FG, fs=6.5)
        else:
            cellL(0, y, DIST_LBL_W, row_h, label, bg, COR_FG, bold=bold_label, fs=6.5, align="left")
            cellL(DIST_LBL_W, y, DIST_VAL_W, row_h, val_str, bg, COR_FG, bold=bold_label, fs=6.5)

    y -= row_h
    cellL(0, y, left_w, row_h, f"Planejamento Redução ({reducao['Pct_plena'].sum():.2f}% pleno)", COR_RED, COR_W, bold=True, fs=7.5)

    PLAN_DATE_W = left_w * 0.28
    PLAN_ID_W = left_w * 0.46
    PLAN_VAL_W = left_w - PLAN_DATE_W - PLAN_ID_W
    toggle = False
    for _, r in reducao.iterrows():
        dft_list = r["DFTs"].split("\n")
        bg = COR_W if not toggle else COR_G
        toggle = not toggle
        y -= row_h
        cellL(0, y, PLAN_DATE_W, row_h, r["MilestoneDate"].strftime("%d/%m/%Y"), bg, COR_FG, fs=6.5)
        cellL(PLAN_DATE_W, y, PLAN_ID_W, row_h, "", bg)
        cellL(PLAN_DATE_W + PLAN_ID_W, y, PLAN_VAL_W, row_h, "", bg)
        for j, dft_id in enumerate(dft_list):
            y -= row_h
            cellL(0, y, PLAN_DATE_W, row_h, "", bg)
            cellL(PLAN_DATE_W, y, PLAN_ID_W, row_h, dft_id.strip(), bg, COR_FG, bold=True, fs=7)
            if j == 0:
                cellL(PLAN_DATE_W + PLAN_ID_W, y, PLAN_VAL_W, row_h, f"{r['Pct']:.2f}%".replace(".", ","), bg, COR_FG, bold=True, fs=7)
            else:
                cellL(PLAN_DATE_W + PLAN_ID_W, y, PLAN_VAL_W, row_h, "", bg)

    # ── 4b. Gráfico de linha com projeção ──────────────────────────────────────
    ax_chart.set_axis_on()
    ax_chart.spines[["top", "right"]].set_visible(False)

    # O real é contínuo até o mês passado e pontilhado do mês passado para o
    # atual: o mês do relatório ainda está correndo, e o número dele muda até
    # fechar. O real vem antes da linha verde de propósito — o eixo é
    # categórico, e a ordem em que os meses aparecem nos plots é a ordem do eixo.
    desenhar_real(ax_chart, labels_hist, vals_hist, fechado=ns["mes_fechado"])
    for m, v in zip(labels_hist, vals_hist):
        ax_chart.annotate(f"{v:.2f}%".replace(".", ","), xy=(m, v), xytext=(0, 7),
                           textcoords="offset points", ha="center", fontsize=7,
                           **ROTULO_SOBRE_LINHA)

    desenhar_verde(ax_chart, labels_hist, plan_hist, proj_labels, proj_vals)
    for lbl, v in zip(proj_labels[1:], proj_vals[1:]):
        ax_chart.annotate(f"{v:.2f}%".replace(".", ","), xy=(lbl, v), xytext=(0, 7),
                          textcoords="offset points", ha="center", va="bottom",
                          fontsize=7, color=fallout_core.COR_VERDE, **ROTULO_SOBRE_LINHA)

    # Base Residencial: a mesma projeção sem o Legado Aberto, em verde-escuro
    # (ver `desenhar_sem_legado`).
    desenhar_sem_legado(ax_chart, labels_hist, ns.get("plan_sem_legado_hist"),
                        proj_labels, ns.get("proj_sem_legado"))

    todas_labels = labels_hist + proj_labels[1:]
    ax_chart.axhline(y=1, color=COR_BLUE, linewidth=1.5, label="Meta (1%)")
    ax_chart.grid(axis="y", linestyle="--", alpha=0.3)

    _todos_vals = (vals_hist + proj_vals
                   + [v for v in plan_hist if v == v])
    _ymin, _ymax = ax_chart.get_ylim()
    _data_max = max(_todos_vals) if _todos_vals else _ymax
    ax_chart.set_ylim(_ymin, max(_ymax, _data_max) * 1.18)

    _yr = ax_chart.get_ylim()[1] - ax_chart.get_ylim()[0]
    _dec = 0 if _yr >= 5 else 1
    ax_chart.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.{_dec}f}%".replace(".", ",")))

    ax_chart.legend(fontsize=7, loc="upper right", framealpha=0.9, edgecolor="none")
    ax_chart.set_xticks(range(len(todas_labels)))
    ax_chart.set_xticklabels(todas_labels, rotation=30, ha="right", fontsize=7)

    # O último dia que entrou nos números (ver fallout_core.data_corte), o mesmo
    # em todos os slides. Rodar o relatório dias depois não muda o rótulo.
    _corte = fallout_core.data_corte(base_dir, hoje) or hoje.tz_convert("America/Sao_Paulo").date()
    ax_chart.text(0.99, 1.04, f"Data de Corte: {_corte.strftime('%d/%b')}",
                  transform=ax_chart.transAxes, ha="right", fontsize=7.5, fontstyle="italic")

    # ── 4c. Tabela detalhada (reaproveitando pivot2) ───────────────────────────
    ax_table.set_xlim(0, sum(COL_W))
    ax_table.set_ylim(0, fig_h)

    _FS = max(0.55, min(1.0, 16.0 / max(n_data, 1)))

    y_t = fig_h - HDR_H
    for ci, col in enumerate(COL_NAMES):
        x = X[ci]; w = sum(COL_W[ci:ci + 1])
        ax_table.add_patch(_rect(x, y_t, w, HDR_H, facecolor=COR_HDR, edgecolor="#FFFFFF", linewidth=0.5))
        ax_table.text(x + w / 2, y_t + HDR_H / 2, col, ha="center", va="center", fontsize=6.5 * _FS,
                      color="#FFFFFF", fontweight="bold", multialignment="center")

    def tcell(ci, y, h, text, bg, fg=COR_FG2, bold=False, fs=6, align="center", span=1):
        x = X[ci]; w = sum(COL_W[ci:ci + span])
        ax_table.add_patch(_rect(x, y, w, h, facecolor=bg, edgecolor="#CCCCCC", linewidth=0.3))
        if not text:
            return
        ha = "left" if align == "left" else "center"
        xt = x + 0.1 if align == "left" else x + w / 2
        ax_table.text(xt, y + h / 2, str(text), ha=ha, va="center", fontsize=fs * _FS,
                      color=fg, fontweight="bold" if bold else "normal",
                      multialignment=ha, clip_on=True)

    y_t -= CELL_H
    for ci in range(len(COL_NAMES) - len(meses_tab)):
        tcell(ci, y_t, CELL_H, "Sucesso" if ci == 0 else "", COR_W2, fs=7)
    for li, m in enumerate(meses_tab):
        suc_pct = pct_m(resumo.loc[m, "Sucessos"] if m in resumo.index else 0, m)
        tcell(len(COL_NAMES) - len(meses_tab) + li, y_t, CELL_H, f"{suc_pct:.2f}%".replace(".", ","), COR_W2, fs=7)

    pivot2_reset = pivot2.reset_index(drop=True)
    for ri in range(len(pivot2_reset)):
        row = pivot2_reset.iloc[ri]
        y_t -= CELL_H
        bg = COR_W2 if ri % 2 == 0 else COR_G2

        tipo = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
        tipo_fmt = "US" if tipo == "User Story" else "Defeito"
        id_str = fmt_id({"DefectKey": row["DefectKey"], "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]})
        ms_fmt = fmt_milestone(row["DFT_BugfixMilestone"])
        nome = str(row["DFT_Name"])[:80] if pd.notna(row["DFT_Name"]) else ""
        team = str(row["DFT_Team"])[:35] if pd.notna(row["DFT_Team"]) else ""
        phase = str(row["DFT_Phase"])[:25] if pd.notna(row["DFT_Phase"]) else ""
        tem_data = bool(ms_fmt)
        team_bold = team == team.upper() and len(team) > 3

        tcell(0, y_t, CELL_H, "Falha" if ri == 0 else "", bg, fs=7)
        tcell(1, y_t, CELL_H, "Em Tratamento/Avaliação\npela Squad" if ri == 0 else "", bg, fs=6.5)
        tcell(2, y_t, CELL_H, tipo_fmt, bg, fs=7)
        tcell(3, y_t, CELL_H, id_str, bg, fs=7)
        tcell(4, y_t, CELL_H, ms_fmt, bg, bold=tem_data, fs=7)
        tcell(5, y_t, CELL_H, nome, bg, fs=5.5, align="left")
        tcell(6, y_t, CELL_H, team, bg, bold=team_bold, fs=6.5)
        tcell(7, y_t, CELL_H, phase, bg, fs=6.5)
        for li, m in enumerate(meses_tab):
            val = int(row.loc[m]) if m in pivot2_reset.columns else 0
            tcell(8 + li, y_t, CELL_H, f"{pct_m(val, m):.2f}%".replace(".", ","), bg, fs=7)

    fig.subplots_adjust(left=0.01, right=0.99, top=0.97, bottom=0.01, hspace=0.18, wspace=0.12)

    ns["fig"] = fig
    ns["ax_chart"] = ax_chart
    return ns


def _rect(x, y, w, h, **kwargs):
    """Retângulo desenhável em qualquer Axes (equivalente ao antigo plt.Rectangle)."""
    from matplotlib.patches import Rectangle
    return Rectangle((x, y), w, h, **kwargs)
