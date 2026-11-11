# -*- coding: utf-8 -*-
"""Relatório completo em PowerPoint editável (tabelas nativas + gráficos como
imagem): slide de 3 colunas, dois slides "Top Ofensores" (Prospect PF e
Base + Cross Sell) e um slide por jornada.

Uso como módulo:  from gerar_slide3col import gerar_pptx_completo
                  buf = gerar_pptx_completo(base_dir)
"""
import io, contextlib
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.ticker as mticker
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor

import distribuicao
import fallout_core
import pipeline
import gerar_pptx
from gerar_pptx import (RED, WHITE, GRAY, FG,
                        _cell, _widths, _no_style, _sem_bordas)

COLUNAS = [("Consolidado",       "Consolidado"),
           ("Prospect",          "Prospect PF"),
           ("Base + Cross Sell", "Base (Móvel) + Cross Sell")]

# Um slide "Top Ofensores" por visão, na ordem do deck, com o sufixo que entra
# no título. O Consolidado não tem o seu: somadas, Prospect e Base + Cross Sell
# disputavam as mesmas 15 linhas e uma escondia a outra.
TOP_OFENSORES = [("Prospect",          "Prospect PF"),
                 ("Base + Cross Sell", "Base")]
# Um slide de "Evolução do Fallout por dia" por visão, logo depois dos Top
# Ofensores e na ordem do slide de 3 colunas, com o começo do título.
FALLOUT_DIARIO = [("Consolidado",       "Consolidado"),
                  ("Prospect",          "Jornada Prospect"),
                  ("Base + Cross Sell", "Jornada Base + Cross Sell")]
# Um slide "Mapa de calor de erros" por visão, depois dos diários.
MAPA_CALOR = [("Base + Cross Sell", "Base + Cross Sell"),
              ("Prospect",          "Prospect PF")]

# Cores do gráfico (matplotlib usa string hex; WHITE/RED do pptx são RGBColor)
COR_BLUE  = "#1565C0"
COR_BRANCO = "#FFFFFF"
COR_SEP   = "#DCDCDC"          # separador entre colunas (cinza esbranquiçado)
COR_VERDE = fallout_core.COR_VERDE     # "Planejado", digitado no histórico
RGB_SEP   = RGBColor(0xDC, 0xDC, 0xDC)


# ── Extração ────────────────────────────────────────────────────────────────
def _extrair_col(jornada, base_dir):
    """
    Roda pipeline.calcular_relatorio() para a jornada e devolve os blocos da
    coluna. Sem a figura do slide de jornada: a coluna desenha o próprio gráfico
    compacto, e o `ns` guardado aqui só ganha a figura se a jornada também tiver
    slide próprio (ver `gerar_pptx_completo`).
    """
    with contextlib.redirect_stdout(io.StringIO()):
        ns = pipeline.calcular_relatorio(base_dir, jornada)

    resumo = ns["resumo"]; ML = ns["MESES_LABEL"]
    # Mesma régua do gráfico da jornada: o realizado termina no mês do
    # relatório. O mês seguinte, com poucos dias de extração, aparece só na
    # projeção — senão o rótulo dele sai duas vezes no eixo e a tabela de
    # volume mostraria um mês de dois dias ao lado de meses fechados.
    todos  = [m for m in sorted(resumo.index.tolist()) if m <= ns["mes_atual"]]
    meses  = todos[-4:]
    fi = fallout_core.fmt_inteiro
    fp = fallout_core.fmt_pct

    volume = {"months": [ML[m] for m in meses], "rows": [
        ("Vendas com Sucesso",      [fi(resumo.loc[m, "Sucessos"]) for m in meses]),
        ("Falha (Análise Técnica)", [fi(resumo.loc[m, "Falhas"])   for m in meses]),
        ("Fallout Rate",            [fp(resumo.loc[m, "Pct"])      for m in meses]),
    ]}

    # A ordem é a do pipeline, já decrescente (ver pipeline._em_ordem_decrescente).
    # Some a linha que arredonda para 0,00%, não só a que é zero: é a régua do
    # slide de jornada (gerar_pptx._montar_dados). Um BSIM de 1 pedido no
    # Consolidado saía como "0,00%", que o leitor não distingue de zero.
    dist = {"pct": ns["fallout_pct"],
            "rows": [(l, fp(v), s, b) for (l, v, s, b) in ns["dist_items"]
                     if round(v, 2) > 0]}

    red = ns["reducao"]
    plan = {"pct": red["Pct_plena"].sum(), "rows": [
        {"date": r["MilestoneDate"].strftime("%d/%m/%Y"),
         "dfts": [d.strip() for d in r["DFTs"].split("\n")],
         "pct":  fp(r["Pct"])} for _, r in red.iterrows()]}

    # "Planejado" só existe onde foi digitado; os buracos viram NaN e o
    # matplotlib simplesmente não desenha o trecho.
    _plan = (resumo["PctPlanejado"].dropna().to_dict()
             if "PctPlanejado" in resumo.columns else {})
    chart = {"labels":  [ML[m] for m in todos],
             "vals":    [resumo.loc[m, "Pct"] for m in todos],
             "planejado": [float(_plan.get(m, float("nan"))) for m in todos],
             "plabels": ns["proj_labels"], "pvals": list(ns["proj_vals"]),
             "fechado": ns["mes_fechado"]}

    return {"volume": volume, "dist": dist, "plan": plan, "chart": chart,
            # O último dia que entrou nos números, o mesmo do slide de jornada.
            "corte": (fallout_core.data_corte(base_dir, ns["hoje"])
                      or ns["hoje"].tz_convert("America/Sao_Paulo")), "ns": ns}


def _chart_png(c, destino, larg_in, alt_in):
    """Gráfico compacto da coluna (real + projeção + meta)."""
    fig = Figure(figsize=(larg_in, alt_in))
    FigureCanvasAgg(fig)
    fig.patch.set_facecolor(COR_BRANCO)
    ax = fig.add_axes([0.13, 0.26, 0.84, 0.66])
    labs = c["labels"] + c["plabels"][1:]

    planejado = c.get("planejado") or []

    # Mesmas regras do slide da jornada (ver pipeline.desenhar_real/_verde):
    # real contínuo até o mês passado e pontilhado até o atual; verde único,
    # Planejado no passado e projeção no futuro. Aqui o eixo é por índice, então
    # a projeção começa na posição do último mês realizado.
    # Marcadores pequenos: com a série histórica inteira são 25 pontos por
    # coluna, e as bolinhas do tamanho anterior encostavam umas nas outras.
    xs_hist = list(range(len(c["labels"])))
    xs_proj = list(range(len(c["labels"]) - 1, len(labs)))
    pipeline.desenhar_real(ax, xs_hist, c["vals"], linewidth=1.3, markersize=2.0,
                           fechado=c.get("fechado", False))
    pipeline.desenhar_verde(ax, xs_hist, planejado, xs_proj, c["pvals"],
                            linewidth=1.3, markersize=2.0)
    # Um rótulo por mês projetado, na vertical como os do realizado. O primeiro
    # ponto da projeção é o mês atual, que já tem o rótulo do realizado.
    for x, v in zip(xs_proj[1:], c["pvals"][1:]):
        ax.annotate(f"{v:.2f}%".replace(".", ","), xy=(x, v), xytext=(0, 6),
                    textcoords="offset points", ha="center", fontsize=6.2,
                    rotation=90, fontweight="bold", color=COR_VERDE,
                    **pipeline.ROTULO_SOBRE_LINHA)
    for i, v in enumerate(c["vals"]):
        ax.annotate(f"{v:.2f}%".replace(".", ","), xy=(i, v), xytext=(0, 6),
                    textcoords="offset points", ha="center", fontsize=6.2, rotation=90,
                    **pipeline.ROTULO_SOBRE_LINHA)

    ax.axhline(y=1, color=COR_BLUE, lw=1.2)
    _para_escala = (list(c["vals"]) + list(c["pvals"])
                    + [v for v in planejado if v == v])
    ax.set_ylim(0, max(_para_escala) * 1.55)
    ax.set_xlim(-0.6, len(labs) - 0.4)
    ax.set_xticks(range(len(labs)))
    ax.set_xticklabels(labs, rotation=45, ha="right", fontsize=5.2)
    ax.tick_params(axis="y", labelsize=6, length=2)
    _dec = 0 if (ax.get_ylim()[1] - ax.get_ylim()[0]) >= 5 else 1
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _, d=_dec: f"{v:.{d}f}%".replace(".", ",")))
    ax.grid(axis="y", ls="--", alpha=0.25)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)

    fig.savefig(destino, dpi=200, facecolor=COR_BRANCO, format="png")


def extrair_todas(base_dir=".", progresso=None):
    """
    Roda o pipeline das 3 visões uma única vez (reaproveitável nos 2 formatos).

    `progresso(rotulo)`, quando passado, é chamado antes de cada visão.
    """
    dados = {}
    for j, _ in COLUNAS:
        if progresso:
            progresso(fallout_core.nome_exibicao(j))
        dados[j] = _extrair_col(j, base_dir)
    return dados


# ── Montagem: PowerPoint ────────────────────────────────────────────────────
def _montar_slide3col(prs, dados):
    """Adiciona o slide de 3 colunas à apresentação `prs` já existente."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Faixa de título
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.52))
    bar.fill.solid(); bar.fill.fore_color.rgb = RED; bar.line.fill.background()
    p = bar.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "Realização e Projeção de Fallout"
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    corte = list(dados.values())[0]["corte"]
    tb = slide.shapes.add_textbox(Inches(0.12), Inches(0.54), Inches(2.2), Inches(0.25))
    p = tb.text_frame.paragraphs[0]
    r = p.add_run(); r.text = f"Data de Corte: {corte.strftime('%d/%b')}"
    r.font.size = Pt(9); r.font.bold = True; r.font.italic = True
    r.font.color.rgb = FG; r.font.name = "Calibri"

    M, GAP = 0.15, 0.22
    CW = (13.333 - 2*M - 2*GAP) / 3
    XS = [M + i*(CW + GAP) for i in range(3)]

    Y_CHART, H_CHART = 0.88, 1.42
    Y_TAB, Y_FUNDO, GAP_BL = 2.42, 7.35, 0.06

    def n_linhas(ci, d):
        n = (2 + len(d["volume"]["rows"])) if ci > 0 else 0
        n += 1 + len(d["dist"]["rows"])
        n += 1 + sum(len(r["dfts"]) for r in d["plan"]["rows"])
        return n

    max_lin = max(n_linhas(ci, dados[j]) for ci, (j, _) in enumerate(COLUNAS))
    RH = min(0.20, (Y_FUNDO - Y_TAB - 2*GAP_BL) / max_lin)
    # Fontes conforme a altura de linha (RH em pt = RH*72; a linha ocupa ~1.2*fonte,
    # e a célula não tem margem: 0.15" = 10.8 pt cabe o título de 9 pt). Era
    # 0.16, e as quatro linhas a mais do US/DFT da Distribuição (0.155") já
    # baixavam tudo para 7 pt.
    if RH >= 0.15:   FD, FT = 8.0, 9.0     # dados, títulos
    elif RH >= 0.13: FD, FT = 7.0, 8.0
    else:            FD, FT = 6.0, 7.0

    # Separadores pontilhados entre as colunas
    for ci in (1, 2):
        xs = XS[ci] - GAP/2
        cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                        Inches(xs), Inches(0.60),
                                        Inches(xs), Inches(Y_FUNDO))
        cn.line.color.rgb = RGB_SEP
        cn.line.width = Pt(1)
        cn.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    for ci, (jornada, titulo) in enumerate(COLUNAS):
        d = dados[jornada]; x0 = XS[ci]

        # Título da coluna
        tb = slide.shapes.add_textbox(Inches(x0), Inches(0.52), Inches(CW), Inches(0.32))
        p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = titulo
        r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = FG; r.font.name = "Calibri"

        # Gráfico — em memória: gravado numa pasta temporária por clique, ele
        # deixava uma "slide3col_*" esquecida no %TEMP% a cada PowerPoint.
        png = io.BytesIO()
        _chart_png(d["chart"], png, CW, H_CHART)
        pic = slide.shapes.add_picture(png, Inches(x0), Inches(Y_CHART), width=Inches(CW))
        # O texto alternativo era o nome do arquivo; de um stream o python-pptx
        # poria "image.png". Mantém o de antes para o slide sair igual.
        pic._element.nvPicPr.cNvPr.set("descr", f"c{ci}.png")

        # As três seções da coluna vão numa tabela só. Eram três tabelas, cada
        # uma numa posição calculada com a altura de linha pedida; quem abre o
        # arquivo pode desenhar a linha mais alta que isso (a fonte pequena tem
        # altura mínima, e cada programa arredonda de um jeito), e aí a tabela de
        # cima crescia por baixo da seguinte — o "Fallout Rate" sumia sob o
        # título da Distribuição. Numa tabela só, o que cresce empurra o resto.
        linhas = []           # (altura, [(início, fim, texto, formato)]) ou (altura, None)

        def secao(titulo):
            linhas.append((RH, [(0.0, 1.0, titulo,
                                 dict(size=FT, bold=True, fg=WHITE, bg=RED))]))

        def espaco():
            linhas.append((GAP_BL, None))

        # ── Volume de Pedidos (colunas 2 e 3) ──
        if ci > 0:
            v = d["volume"]; n = len(v["months"])
            lw_ = 0.40
            bordas = [0.0, lw_] + [lw_ + (1 - lw_) * (i + 1) / n for i in range(n)]
            secao("Volume de Pedidos")
            linhas.append((RH, [(0.0, lw_, "", dict(size=FD, bg=RED))] + [
                (bordas[1 + i], bordas[2 + i], m, dict(size=FD, bold=True, fg=WHITE, bg=RED))
                for i, m in enumerate(v["months"])]))
            for ri, (lbl, vals) in enumerate(v["rows"]):
                bg = WHITE if ri % 2 == 0 else GRAY
                linhas.append((RH, [(0.0, lw_, lbl, dict(size=FD, bg=bg, align="left"))] + [
                    (bordas[1 + i], bordas[2 + i], val, dict(size=FD, bg=bg))
                    for i, val in enumerate(vals)]))
            espaco()

        # ── Distribuição Fallout ──
        VW = 0.26
        secao(f"Distribuição Fallout ({d['dist']['pct']:.2f}%)".replace(".", ","))
        tem_nivel_2 = any(int(r[2]) >= 2 for r in d["dist"]["rows"])
        for ri, (lbl, val, sub, _b) in enumerate(d["dist"]["rows"]):
            bg = WHITE if ri % 2 == 0 else GRAY
            # Toda a Distribuição em negrito; subgrupo alinhado à direita, e o
            # de primeiro nível (US/DFT) um pouco antes do fim, para o c/ data e
            # s/ data abaixo dele lerem como parte dele.
            fim = 1 - VW - (0.08 if int(sub) == 1 and tem_nivel_2 else 0)
            celulas = [(0.0, fim, fallout_core.rotulo_dist(lbl, sub),
                        dict(size=FD, bold=True, bg=bg, align="right" if sub else "left"))]
            if fim < 1 - VW:
                celulas.append((fim, 1 - VW, "", dict(size=FD, bg=bg)))
            linhas.append((RH, celulas + [
                (1 - VW, 1.0, "" if fallout_core.so_rotulo(lbl, sub) else val,
                 dict(size=FD, bold=True, bg=bg, align="right"))]))
        espaco()

        # ── Planejamento Redução ──
        DW, PW = 0.30, 0.22
        secao(f"Planejamento Redução ({d['plan']['pct']:.2f}%)".replace(".", ","))
        for ri, r in enumerate(d["plan"]["rows"]):
            bg = WHITE if ri % 2 == 0 else GRAY
            linhas.append((RH * len(r["dfts"]), [
                (0.0, DW, r["date"], dict(size=FD, bg=bg)),
                (DW, 1 - PW, "\n".join(r["dfts"]), dict(size=FD, bold=True, bg=bg)),
                (1 - PW, 1.0, r["pct"], dict(size=FD, bold=True, bg=bg, align="right"))]))

        _tabela_empilhada(slide, x0, Y_TAB, CW, linhas)


def _tabela_empilhada(slide, x, y, largura, linhas):
    """
    Uma tabela com seções de colunas diferentes, uma embaixo da outra.

    `linhas` é [(altura, células)], com as células como (início, fim, texto,
    formato do `_cell`) e início/fim em fração da largura; `células` None é um
    espaço em branco entre seções. A grade junta as bordas de todas as seções,
    e cada célula ocupa as colunas entre as suas mesclando-as.
    """
    grade = sorted({round(b, 6) for _, cels in linhas if cels
                    for c in cels for b in c[:2]})
    col = {b: i for i, b in enumerate(grade)}
    gf = slide.shapes.add_table(len(linhas), len(grade) - 1, Inches(x), Inches(y),
                                Inches(largura), Inches(sum(h for h, _ in linhas)))
    t = gf.table
    _no_style(t)
    _widths(t, [(b - a) * largura for a, b in zip(grade, grade[1:])])

    # Toda célula começa sem margem e com fonte de 1 pt no fim do parágrafo: a
    # célula vazia (o espaço entre seções e as que a mescla esconde) herdaria
    # os 18 pt padrão, e a linha dela não desceria abaixo disso.
    for linha in t.rows:
        for cel in linha.cells:
            cel.margin_top = cel.margin_bottom = Inches(0)
            cel.text_frame.paragraphs[0]._p.get_or_add_endParaRPr().set("sz", "100")

    for ri, (altura, cels) in enumerate(linhas):
        t.rows[ri].height = Inches(altura)
        if cels is None:
            t.cell(ri, 0).merge(t.cell(ri, len(grade) - 2))
            continue
        for ini, fim, texto, fmt in cels:
            c0, c1 = col[round(ini, 6)], col[round(fim, 6)] - 1
            if c1 > c0:
                t.cell(ri, c0).merge(t.cell(ri, c1))
            _cell(t.cell(ri, c0), texto, **fmt)
    _sem_bordas(t)


# ── Slide "Distribuição do Fallout detalhada" ────────────────────────────────
# Um por visão do slide de 3 colunas, logo depois dele. À esquerda a mesma
# Distribuição da coluna; ao lado, o que preenche as categorias que pedem
# detalhe: os defeitos do Em Tratamento (US e DFT, c/ e s/ data), do Tratado e
# do Resolvido e as mensagens
# mais frequentes do Incidente.
DETALHADO = [("Consolidado",       "Consolidado"),
             ("Prospect",          "Prospect PF"),
             ("Base + Cross Sell", "Base (Móvel) + Cross Sell")]
TOP_INCIDENTES = 5
CAT_TRATADO = "Tratado - Em avaliação de eficácia"
LIMITE_NOME = 120                # o que cabe numa linha na largura do slide
LIMITE_MENSAGEM = 110


def _pct_txt(valor):
    # Um defeito de poucos pedidos arredonda para 0,00%, que se lê como zero.
    if 0 < valor < 0.005:
        return "<0,01%"
    return f"{valor:.2f}%".replace(".", ",")


def _texto(valor, limite=None):
    txt = "" if pd.isna(valor) else " ".join(str(valor).split())
    return txt if limite is None or len(txt) <= limite else txt[:limite - 1].rstrip() + "…"


def _defeitos(pedidos, total):
    """
    Linhas do slide, como (defeito, nome, equipe, entrega, %): o agrupamento de
    `distribuicao.por_defeito`, com os textos cortados para caber.
    """
    tabela = distribuicao.por_defeito(pedidos, total)
    return [(r["Defeito/US"], _texto(r["Nome"], LIMITE_NOME), _texto(r["Equipe"], 40),
             "" if pd.isna(r["Entrega"]) else r["Entrega"].strftime("%d/%m/%Y"), r["%"])
            for _, r in tabela.iterrows()]


def _detalhes(ns):
    """
    O que preenche os subgrupos de Em Tratamento (US e DFT, cada um c/ data e
    s/ data — a chave é o rótulo do subgrupo), Tratado, Resolvido e Incidente
    no mês do relatório.
    """
    _df_mes, cats, _t = fallout_core.categorizar(ns["df"], ns["mes_atual"], hoje=ns["hoje"])
    total = ns["total_mes"]
    em_trat = cats[fallout_core.CAT_TRATAMENTO]
    subgrupo = fallout_core.subgrupo_tratamento(em_trat)
    incidentes = distribuicao.por_mensagem(cats[fallout_core.CAT_INCIDENTE], total,
                                           LIMITE_MENSAGEM)
    partes = {s: em_trat[subgrupo == s] for subs in fallout_core.SUBGRUPOS_TRATAMENTO.values()
              for s in subs}
    return {
        **{s: (len(p) / total * 100, _defeitos(p, total)) for s, p in partes.items()},
        "tratado": (len(cats[CAT_TRATADO]) / total * 100, _defeitos(cats[CAT_TRATADO], total)),
        "resolvido": (len(cats["Resolvido"]) / total * 100, _defeitos(cats["Resolvido"], total)),
        "incidente": (len(cats[fallout_core.CAT_INCIDENTE]) / total * 100,
                      list(zip(incidentes["Mensagem"].head(TOP_INCIDENTES),
                               incidentes["%"].head(TOP_INCIDENTES)))),
    }


def _montar_slide_detalhado(prs, titulo, d, ns):
    """
    Slide "Distribuição do Fallout detalhada" de uma visão (ver DETALHADO): a
    Distribuição numa coluna só, com o detalhe logo abaixo da linha que ele
    explica — os defeitos sob o c/ data e o s/ data de US e de DFT, Tratado e
    Resolvido e as mensagens sob Incidente. O % do detalhe fica na mesma coluna
    do % da categoria.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.52))
    bar.fill.solid(); bar.fill.fore_color.rgb = RED; bar.line.fill.background()
    p = bar.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = f"Distribuição do Fallout detalhada – {titulo}"
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    tb = slide.shapes.add_textbox(Inches(0.12), Inches(0.54), Inches(6), Inches(0.25))
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = (f"{ns['MESES_LABEL'][ns['mes_atual']]} · "
              f"Data de Corte: {d['corte'].strftime('%d/%b')}")
    r.font.size = Pt(9); r.font.bold = True; r.font.italic = True
    r.font.color.rgb = FG; r.font.name = "Calibri"

    det = _detalhes(ns)
    X, LARG = 0.15, 13.333 - 0.3
    Y0, FUNDO = 0.85, 7.42
    # A categoria em 8 pt; o detalhe em 7, numa linha mais baixa, para caber
    # a lista dos cinco blocos no Consolidado.
    RH, RH_DET, RH_TIT, FC, FD, FT = 0.16, 0.14, 0.19, 8.0, 7.0, 9.0
    RECUO = " " * 6                       # subgrupo abaixo da categoria
    # Defeito/US | Nome | Equipe | Entrega | %, com o % alinhado ao da categoria
    F = [0.02, 0.11, 0.62, 0.78, 0.90, 1.0]
    NOMES_DETALHE = {**{s: s for subs in fallout_core.SUBGRUPOS_TRATAMENTO.values()
                        for s in subs},
                     CAT_TRATADO: "tratado", "Resolvido": "resolvido"}

    def linha_categoria(lbl, val, sub):
        bg = WHITE if sub else GRAY
        return (RH, [(0.0, F[4], RECUO * int(sub) + fallout_core.rotulo_dist(lbl, sub),
                      dict(size=FC, bold=True, bg=bg, align="left")),
                     (F[4], 1.0, "" if fallout_core.so_rotulo(lbl, sub) else val,
                      dict(size=FC, bold=True, bg=bg, align="right"))])

    def cabecalho(colunas):
        return (RH_DET, [(0.0, F[0], "", dict(size=FD, bg=WHITE))] + [
            (ini, fim, txt, dict(size=FD, bold=True, fg=FG, bg=WHITE,
                                 align="right" if txt == "%" else
                                 "center" if txt in ("Equipe", "Entrega") else "left"))
            for ini, fim, txt in colunas])

    def linha_defeito(rot, nome, equipe, entrega, p_):
        return (RH_DET, [(0.0, F[0], "", dict(size=FD, bg=WHITE)),
                     (F[0], F[1], rot, dict(size=FD, bold=True, bg=WHITE, align="left")),
                     (F[1], F[2], nome, dict(size=FD, bg=WHITE, align="left")),
                     (F[2], F[3], equipe, dict(size=FD, bg=WHITE)),
                     (F[3], F[4], entrega, dict(size=FD, bg=WHITE)),
                     (F[4], 1.0, _pct_txt(p_), dict(size=FD, bg=WHITE, align="right"))])

    def linha_resto(qtd, p_):
        return (RH_DET, [(0.0, F[0], "", dict(size=FD, bg=WHITE)),
                     (F[0], 1.0, f"… e mais {qtd} {'defeito' if qtd == 1 else 'defeitos'} "
                                 f"({_pct_txt(p_)})",
                      dict(size=FD, bg=WHITE, fg=FG, align="left"))])

    # Quantos defeitos de cada bloco cabem. Tudo o que não é defeito tem altura
    # fixa; se faltar espaço, sai o defeito de menor % entre todos os blocos —
    # o último mostrado de algum deles, já que cada bloco vem do maior para o
    # menor —, e o que saiu de cada bloco vira uma linha "… e mais N".
    rotulos = [l for l, _v, _s, _b in d["dist"]["rows"]]
    blocos = {c: det[c][1] for l, c in NOMES_DETALHE.items() if l in rotulos and det[c][1]}
    fixo = RH_TIT + RH * len(rotulos) + RH_DET * len(blocos)          # + cabeçalhos
    if det["incidente"][1] and fallout_core.CAT_INCIDENTE in rotulos:
        fixo += RH_DET * (1 + len(det["incidente"][1]))
    mostrar = {c: len(lista) for c, lista in blocos.items()}

    def altura_defeitos():
        return sum(RH_DET * (n + (n < len(blocos[c]))) for c, n in mostrar.items())

    while fixo + altura_defeitos() > FUNDO - Y0 and any(mostrar.values()):
        menor = min((c for c in mostrar if mostrar[c]),
                    key=lambda c: blocos[c][mostrar[c] - 1][4])
        mostrar[menor] -= 1

    linhas = [(RH_TIT, [(0.0, 1.0, _pct_txt(d["dist"]["pct"]).join(
        ("Distribuição Fallout (", ")")), dict(size=FT, bold=True, fg=WHITE, bg=RED))])]
    colunas_defeito = list(zip(F[:-1], F[1:], ["Defeito/US", "Nome", "Equipe", "Entrega", "%"]))
    for lbl, val, sub, _b in d["dist"]["rows"]:
        linhas.append(linha_categoria(lbl, val, sub))
        chave = NOMES_DETALHE.get(lbl)
        if chave in blocos:
            defeitos = blocos[chave]
            linhas.append(cabecalho(colunas_defeito))
            linhas += [linha_defeito(*x) for x in defeitos[:mostrar[chave]]]
            resto = defeitos[mostrar[chave]:]
            if resto:
                linhas.append(linha_resto(len(resto), sum(x[4] for x in resto)))
        elif lbl == fallout_core.CAT_INCIDENTE and det["incidente"][1]:
            linhas.append(cabecalho([(F[0], F[4], f"Mensagem de erro · top {TOP_INCIDENTES}"),
                                     (F[4], 1.0, "%")]))
            for msg, p_ in det["incidente"][1]:
                linhas.append((RH_DET, [(0.0, F[0], "", dict(size=FD, bg=WHITE)),
                                        (F[0], F[4], msg, dict(size=FD, bg=WHITE, align="left")),
                                        (F[4], 1.0, _pct_txt(p_),
                                         dict(size=FD, bg=WHITE, align="right"))]))
    _tabela_empilhada(slide, X, Y0, LARG, linhas)


def gerar_pptx_completo(base_dir=".", dados_3col=None, progresso=None):
    """
    PPTX único:
      1. Slide de 3 colunas (Consolidado | Prospect | Base + Cross Sell) e,
         depois dele, um "Distribuição do Fallout detalhada" por coluna
      2. Dois slides "Top Ofensores" — Top 15 defeitos/US do Prospect PF e do
         Base + Cross Sell, cada um em relação ao total da própria visão
      3. Três slides "Evolução do Fallout por dia" — os últimos 15 dias do
         Consolidado, do Prospect e do Base + Cross Sell — e dois "Mapa de calor
         de falhas pontuais", do Base + Cross Sell e do Prospect PF, na mesma
         janela
      4. Um slide no estilo do "print" para cada jornada bruta encontrada em
         extrações/ (ex.: Base Móvel, Cross Sell, Prospect — sem os combinados)
    Devolve os bytes (BytesIO).

    `progresso(feitos, total, rotulo)`, quando passado, é chamado antes de cada
    etapa — uma por visão do slide de 3 colunas, uma por jornada e a montagem.
    """
    jornadas_brutas = fallout_core.jornadas_disponiveis(base_dir)
    etapa = {"feitos": 0, "total": (0 if dados_3col else len(COLUNAS))
             + len(jornadas_brutas) + 1}

    def _avisar(rotulo):
        if progresso:
            progresso(etapa["feitos"], etapa["total"], rotulo)
        etapa["feitos"] += 1

    dados_3col = dados_3col or extrair_todas(base_dir, progresso=_avisar)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    _montar_slide3col(prs, dados_3col)
    for visao, titulo in DETALHADO:
        if visao in dados_3col:
            _montar_slide_detalhado(prs, titulo, dados_3col[visao], dados_3col[visao]["ns"])

    # A jornada que já foi calculada para o slide de 3 colunas (o Prospect) não
    # é calculada de novo: só ganha a figura do slide dela.
    ns_por_jornada = {}
    for j in jornadas_brutas:
        _avisar(fallout_core.nome_exibicao(j))
        pronto = (dados_3col.get(j) or {}).get("ns")
        ns_por_jornada[j] = (pipeline.desenhar_dashboard(pronto, base_dir, j)
                             if pronto is not None else pipeline.gerar_relatorio(base_dir, j))

    _avisar("Montando os slides")
    # As visões são as do slide de 3 colunas, já calculadas; cada linha da
    # tabela vem da jornada bruta, para a coluna "Jornada" separar Base de
    # Cross Sell. As de formato RPA e NBA (PME, Base Residencial) entram apenas
    # com o slide individual delas, mais abaixo.
    ns_visoes = {}

    def _ns_visao(visao):
        if visao not in ns_visoes:
            ns = (dados_3col.get(visao) or {}).get("ns")
            ns_visoes[visao] = ns if ns is not None else pipeline.calcular_relatorio(base_dir, visao)
        return ns_visoes[visao]

    for visao, sufixo in TOP_OFENSORES:
        membros = {j: ns_por_jornada[j]
                   for j in fallout_core.resolver_jornadas(visao, base_dir)
                   if j in ns_por_jornada}
        if not membros:
            continue
        ns_visao = _ns_visao(visao)
        gerar_pptx._montar_slide_top_ofensores(
            prs, ns_visao, membros,
            titulo=f"{gerar_pptx.TITULO_TOP_OFENSORES} - {sufixo}",
        )

    # O df da visão já vem cortado e sem pedido repetido; a categoria de cada
    # pedido é a da Distribuição, vista no mesmo "hoje" do relatório.
    jornadas_existentes = set(jornadas_brutas)
    for visao, nome in FALLOUT_DIARIO:
        if not set(fallout_core.resolver_jornadas(visao, base_dir)) <= jornadas_existentes:
            continue
        ns_visao = _ns_visao(visao)
        tabela = fallout_core.fallout_por_dia(ns_visao["df"], base_dir, visao,
                                              hoje=ns_visao["hoje"])
        gerar_pptx._montar_slide_diario(
            prs, f"{nome} – Evolução do Fallout por dia", tabela)

    # Mapa de calor de erros (a tabela da aba Análise de Erros), na mesma janela
    # de dias dos slides diários.
    for visao, nome in MAPA_CALOR:
        if not set(fallout_core.resolver_jornadas(visao, base_dir)) <= jornadas_existentes:
            continue
        ns_visao = _ns_visao(visao)
        fim = fallout_core.fim_da_visao(base_dir, visao, ns_visao["hoje"])
        if fim is None:
            continue
        gerar_pptx._montar_slide_mapa_calor(
            prs, f"{nome} – Mapa de calor de falhas pontuais",
            gerar_pptx.tabela_mapa_calor(ns_visao["df"], fim, base_dir, visao,
                                          hoje=ns_visao["hoje"]))

    for jornada, ns in ns_por_jornada.items():
        data, chart_png = gerar_pptx._montar_dados(jornada, ns)
        gerar_pptx._montar_slide(prs, data, chart_png)

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf
