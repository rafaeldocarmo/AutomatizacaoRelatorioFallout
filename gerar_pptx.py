# -*- coding: utf-8 -*-
"""Blocos de montagem dos slides em PowerPoint editável (tabelas nativas +
gráfico como imagem), consumidos por gerar_slide3col.gerar_pptx_completo().
"""
import io
import matplotlib
matplotlib.use("Agg")
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

import fallout_core

RED = RGBColor.from_string(fallout_core.COR_VERMELHO.lstrip("#"))
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0xF2, 0xF2, 0xF2); GRAY2 = RGBColor(0xF5, 0xF5, 0xF5)
FG = RGBColor(0x22, 0x22, 0x22)


# ── Extração dos dados via pipeline ──────────────────────────────────────────
# Coluna da tabela de defeitos do slide de jornada que só o Base Residencial tem.
COL_CRIADO = "Criado em"


def _pct_sucesso(resumo, mes, pct_m):
    """
    % de sucesso do mês na linha "Sucesso" das tabelas de defeitos. No mês
    fechado digitado só com a taxa não há Sucessos, e a conta dava "nan%" (set/26
    do Base + Cross Sell): aí vale 100 menos a taxa digitada.
    """
    import pandas as pd

    if mes not in resumo.index:
        return 0.0
    sucessos = resumo.loc[mes, "Sucessos"]
    if pd.isna(sucessos):
        return 100.0 - float(resumo.loc[mes, "Pct"])
    return pct_m(sucessos, mes)


def _linhas_do_nome(nome, largura):
    """
    Quantas linhas o nome do defeito ocupa numa célula de `largura` polegadas,
    em 8 pt — para estimar a altura da linha antes de o PowerPoint esticá-la.

    Cabem ~21 caracteres por polegada em minúsculas e ~16 em maiúsculas, e os
    nomes do Octane vêm quase todos em maiúsculas: contados a 21, o DFT 233292
    (150 caracteres) saía como duas linhas, quebrava em três e a última linha
    da tabela do Base Móvel ficava cortada na borda do slide. A quebra é por
    palavra, e os 8% a mais cobrem a sobra no fim de cada linha.
    """
    texto = str(nome)
    letras = [c for c in texto if c.isalpha()]
    maiusculas = sum(c.isupper() for c in letras) / len(letras) if letras else 0.0
    por_linha = max(int((largura - 0.06) * (21.0 - 5.0 * maiusculas)), 10)
    return max(1, -(-int(len(texto) * 1.08) // por_linha))


def _montar_dados(jornada, ns):
    """Monta (data, chart_png em memória) a partir de um resultado já calculado
    por pipeline.gerar_relatorio()."""
    import pandas as pd
    resumo = ns["resumo"]; ML = ns["MESES_LABEL"]; reducao = ns["reducao"]
    meses_vol = ns["meses_vol"]
    fi = fallout_core.fmt_inteiro
    fp = fallout_core.fmt_pct

    volume = {"months": [ML[m].capitalize() for m in meses_vol], "rows": [
        ["Vendas com Sucesso"] + [fi(resumo.loc[m, "Sucessos"]) for m in meses_vol],
        ["Falha (Análise Técnica)"] + [fi(resumo.loc[m, "Falhas"]) for m in meses_vol],
        ["Fallout Rate"] + [fp(resumo.loc[m, "Pct"]) for m in meses_vol],
    ]}
    # Categorias zeradas ficam de fora do slide de jornada: são 12 linhas fixas,
    # e as que não aconteceram no mês só gastavam altura repetindo "0,00%".
    # O teste é sobre o valor já arredondado para duas casas, que é o que
    # aparece na tabela — assim 0,004% também sai, em vez de virar um "0,00%"
    # que o leitor não consegue distinguir de zero.
    # O slide de 3 colunas monta a lista por conta própria (gerar_slide3col),
    # com a mesma régua.
    dist_itens = [it for it in ns["dist_items"] if round(it[1], 2) > 0]
    dist = {"title": f"Distribuição Fallout ({fp(ns['fallout_pct'])})", "rows": [
        {"label": lbl, "val": fp(val), "sub": sub, "bold": bold}
        for (lbl, val, sub, bold) in dist_itens]}
    plan_rows = []
    for _, r in reducao.iterrows():
        dfts = r["DFTs"].split("\n")
        plan_rows.append({"date": r["MilestoneDate"].strftime("%d/%m/%Y"),
                          "dfts": [{"id": d.strip(), "pct": fp(r["Pct"] / len(dfts))} for d in dfts]})
    plan = {"title": f"Planejamento Redução ({fp(reducao['Pct_plena'].sum())} pleno)", "rows": plan_rows}

    COL_NAMES = list(ns["COL_NAMES"]); meses_tab = ns["meses_tab"]
    pivot2 = ns["pivot2"].reset_index(drop=True)
    fmt_id = ns["fmt_id"]; fmt_ms = ns["fmt_milestone"]; pct_m = ns["pct_m"]
    # "Criado em" logo depois do número, só na jornada que traz a data de
    # abertura do defeito (o Base Residencial, ver `pipeline.criado_em`).
    criado_em = ns.get("criado_em") or {}
    if criado_em:
        COL_NAMES.insert(COL_NAMES.index("Defeito / US") + 1, COL_CRIADO)
    n_fixas = len(COL_NAMES) - len(meses_tab)
    det_rows = []
    suc = ["Sucesso"] + [""] * (n_fixas - 1)
    for m in meses_tab:
        sv = _pct_sucesso(resumo, m, pct_m)
        suc.append(f"{sv:.2f}%".replace(".", ","))
    det_rows.append({"cells": suc, "kind": "sucesso"})
    for ri in range(len(pivot2)):
        row = pivot2.iloc[ri]
        tipo = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
        cells = [
            "Falha" if ri == 0 else "",
            "Em Tratamento/Avaliação pela Squad" if ri == 0 else "",
            "US" if tipo == "User Story" else "Defeito",
            fmt_id({"DefectKey": row["DefectKey"], "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]}),
            *([fmt_ms(criado_em.get(str(row["DefectKey"])))] if criado_em else []),
            fmt_ms(row["DFT_BugfixMilestone"]),
            str(row["DFT_Name"]) if pd.notna(row["DFT_Name"]) else "",
            str(row["DFT_Team"])[:35] if pd.notna(row["DFT_Team"]) else "",
            str(row["DFT_Phase"])[:25] if pd.notna(row["DFT_Phase"]) else "",
        ]
        for m in meses_tab:
            val = int(row.loc[m]) if m in pivot2.columns else 0
            cells.append(f"{pct_m(val, m):.2f}%".replace(".", ","))
        det_rows.append({"cells": cells, "kind": "dft"})
    detalhe = {"cols": COL_NAMES, "rows": det_rows}

    # recorta o gráfico
    fig = ns["fig"]; axc = ns["ax_chart"]
    fig.canvas.draw()
    ext = axc.get_tightbbox(fig.canvas.get_renderer()).transformed(fig.dpi_scale_trans.inverted())
    ext = ext.expanded(1.02, 1.0)
    # Em memória, e não numa pasta de mkdtemp que ninguém apagava.
    chart_png = io.BytesIO()
    fig.savefig(chart_png, dpi=200, bbox_inches=ext, pad_inches=0.08, facecolor="white",
                format="png")

    data = {"title": f"{fallout_core.nome_exibicao(jornada)} – Realização e Projeção de Fallout",
            "volume": volume, "dist": dist, "plan": plan, "detalhe": detalhe}
    return data, chart_png


# ── Montagem do slide ────────────────────────────────────────────────────────
def _no_style(tbl):
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0"); tblPr.set("bandRow", "0")
    for sid in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(sid)

def _cell(cell, text, size, bold=False, fg=FG, bg=WHITE, align="center", valign="middle"):
    cell.fill.solid(); cell.fill.fore_color.rgb = bg
    cell.margin_left = Inches(0.03); cell.margin_right = Inches(0.03)
    cell.margin_top = Inches(0.01); cell.margin_bottom = Inches(0.01)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE if valign == "middle" else MSO_ANCHOR.TOP
    tf = cell.text_frame; tf.word_wrap = True
    for i, ln in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"center": PP_ALIGN.CENTER, "left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT}[align]
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = fg; r.font.name = "Calibri"

def _no_border(cell):
    """Remove as bordas da célula (deixa como o PNG, sem grade)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for tag in ("lnB", "lnT", "lnR", "lnL"):   # insere no início → ordem final lnL,lnR,lnT,lnB
        ln = parse_xml('<a:%s %s w="0" cap="flat"><a:noFill/></a:%s>' % (tag, nsdecls("a"), tag))
        tcPr.insert(0, ln)

def _sem_bordas(tbl):
    for row in tbl.rows:
        for cell in row.cells:
            _no_border(cell)

def _borda_clara(tbl, hexcor="D9D9D9", w=6350):
    """Bordas finas em cinza claro (como o PNG) em todas as células."""
    for row in tbl.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                for el in tcPr.findall(qn(tag)):
                    tcPr.remove(el)
            for tag in ("lnB", "lnT", "lnR", "lnL"):
                ln = parse_xml(
                    '<a:%s %s w="%d" cap="flat"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:%s>'
                    % (tag, nsdecls("a"), w, hexcor, tag))
                tcPr.insert(0, ln)

def _widths(tbl, ws):
    for i, w in enumerate(ws): tbl.columns[i].width = Inches(w)
def _rowh(tbl, h):
    for r in tbl.rows: r.height = Inches(h)

def _montar_slide(prs, d, chart_png):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tb = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.55))
    tb.fill.solid(); tb.fill.fore_color.rgb = RED; tb.line.fill.background()
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = d["title"]
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    def add_tbl(nr, nc, x, y, w, h):
        gf = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
        _no_style(gf.table); return gf.table

    LX, LW = 0.18, 5.75; y = 0.62
    vm = d["volume"]["months"]; drows_d = d["dist"]["rows"]; prows = d["plan"]["rows"]
    # Altura da tabela em múltiplos de RH (não é a contagem de linhas): cada
    # data ocupa uma linha só, alta o suficiente para os defeitos dela. Sem
    # nenhuma entrega prevista sobra uma linha para dizer isso — o cabeçalho
    # vermelho sozinho parecia tabela quebrada.
    sem_plano = not prows
    n_plan = 1 + (1 if sem_plano else sum(len(r["dfts"]) for r in prows))
    GAP = 0.04
    # Fonte fixa em 8pt. A linha fica no mínimo que 8pt exige em vez de esticar
    # para ocupar o espaço: o bloco da esquerda disputa altura com a tabela de
    # detalhe lá embaixo, e esticar aqui comia as linhas de defeito de lá.
    FD, FT = 8.0, 9.0
    RH = 0.16

    vt = add_tbl(2 + len(d["volume"]["rows"]), 1 + len(vm), LX, y, LW, RH * (2 + len(d["volume"]["rows"])))
    vlabw = 2.5; vcolw = (LW - vlabw) / len(vm)
    _widths(vt, [vlabw] + [vcolw] * len(vm)); _rowh(vt, RH)
    _cell(vt.cell(0, 0), "", FT, bg=RED)
    vt.cell(0, 1).merge(vt.cell(0, len(vm)))
    _cell(vt.cell(0, 1), "Volume de Pedidos", FT, bold=True, fg=WHITE, bg=RED)
    _cell(vt.cell(1, 0), "", FT, bg=RED)
    for i, ml in enumerate(vm): _cell(vt.cell(1, 1 + i), ml, FT, bold=True, fg=WHITE, bg=RED)
    for ri, row in enumerate(d["volume"]["rows"]):
        bg = WHITE if ri % 2 == 0 else GRAY
        _cell(vt.cell(2 + ri, 0), row[0], FD, bg=bg, align="left")
        for i in range(len(vm)): _cell(vt.cell(2 + ri, 1 + i), row[1 + i], FD, bg=bg)
    y += RH * (2 + len(d["volume"]["rows"])) + GAP

    dt = add_tbl(1 + len(drows_d), 2, LX, y, LW, RH * (1 + len(drows_d)))
    _widths(dt, [LW - 1.5, 1.5]); _rowh(dt, RH)
    dt.cell(0, 0).merge(dt.cell(0, 1))
    _cell(dt.cell(0, 0), d["dist"]["title"], FT, bold=True, fg=WHITE, bg=RED)
    tgl = False
    for ri, row in enumerate(drows_d):
        bg = WHITE if not tgl else GRAY; tgl = not tgl
        # Cinco espaços de recuo por nível (US/DFT, e o c/ data e s/ data deles).
        lbl = "     " * int(row["sub"]) + fallout_core.rotulo_dist(row["label"], row["sub"])
        _cell(dt.cell(1 + ri, 0), lbl, FD, bold=row["bold"], bg=bg, align="left")
        _cell(dt.cell(1 + ri, 1), "" if fallout_core.so_rotulo(row["label"], row["sub"])
              else row["val"], FD, bold=row["bold"], bg=bg, align="right")
    y += RH * (1 + len(drows_d)) + GAP

    # Data | DFT | %, tudo na mesma linha da tabela. Uma linha por data: os
    # defeitos que caem na mesma entrega ficam empilhados dentro da célula, um
    # por linha, para a data não se repetir. O % continua sendo o de cada DFT,
    # empilhado na mesma ordem — as duas células têm sempre a mesma quantidade
    # de linhas, então cada defeito fica lado a lado com o seu percentual.
    PLAN_DATA_W, PLAN_PCT_W = 1.25, 1.0
    pt = add_tbl(1 + (1 if sem_plano else len(prows)), 3, LX, y, LW, RH * n_plan)
    _widths(pt, [PLAN_DATA_W, LW - PLAN_DATA_W - PLAN_PCT_W, PLAN_PCT_W])
    pt.rows[0].height = Inches(RH)
    pt.cell(0, 0).merge(pt.cell(0, 2))
    _cell(pt.cell(0, 0), d["plan"]["title"], FT, bold=True, fg=WHITE, bg=RED)
    if sem_plano:
        pt.rows[1].height = Inches(RH)
        pt.cell(1, 0).merge(pt.cell(1, 2))
        _cell(pt.cell(1, 0), "Nenhum defeito com data de correção prevista",
              FD, bg=WHITE, align="left")
    for ri, pr in enumerate(prows):
        bg = WHITE if ri % 2 == 0 else GRAY
        # A linha cresce com a quantidade de defeitos da data. Sem fixar isso o
        # PowerPoint estica sozinho para caber o texto, a tabela passa da altura
        # calculada aqui e empurra a tabela de detalhe para fora do slide.
        pt.rows[1 + ri].height = Inches(RH * len(pr["dfts"]))
        _cell(pt.cell(1 + ri, 0), pr["date"], FD, bg=bg, align="left")
        _cell(pt.cell(1 + ri, 1), "\n".join(x["id"] for x in pr["dfts"]),
              FD, bold=True, bg=bg, align="left")
        _cell(pt.cell(1 + ri, 2), "\n".join(x["pct"] for x in pr["dfts"]),
              FD, bold=True, bg=bg, align="right")
    left_bottom = y + RH * n_plan

    # bloco esquerdo sem bordas (igual ao PNG)
    for _t in (vt, dt, pt):
        _sem_bordas(_t)

    CHART_X, CHART_Y, CHART_W = 6.15, 0.72, 6.95
    pic = slide.shapes.add_picture(chart_png, Inches(CHART_X), Inches(CHART_Y), width=Inches(CHART_W))
    # O texto alternativo era o nome do arquivo em disco; de um stream o
    # python-pptx poria "image.png". Mantém o de antes para o slide sair igual.
    pic._element.nvPicPr.cNvPr.set("descr", "chart.png")
    chart_bottom = CHART_Y + CHART_W / 2.08

    cols = d["detalhe"]["cols"]; drows = d["detalhe"]["rows"]
    dety = max(left_bottom, chart_bottom) + GAP
    LARG_DET = 13.333
    # Larguras pelo nome da coluna: a do Base Residencial tem "Criado em" a
    # mais, e a diferença sai do nome do defeito, que quebra em linhas.
    larguras = {"Classificação": 0.95, "Agrupamento": 1.45, "Tipo": 0.7,
                "Defeito / US": 0.9, COL_CRIADO: 0.9, "Data Prod\nnão resolvidos": 1.1,
                "Nome Defeito": 3.9, "Time": 1.65, "Status\nDefeito/US": 1.35}
    fixas = [c for c in cols if c in larguras]
    ws = [larguras[c] for c in fixas]
    COL_NOME_DET = cols.index("Nome Defeito")  # dentro de row["cells"]
    if COL_CRIADO in fixas:
        ws[COL_NOME_DET] -= larguras[COL_CRIADO]
    meses_det = cols[len(fixas):]
    ws += [(LARG_DET - sum(ws)) / len(meses_det)] * len(meses_det)
    COLS_SEC = {cols.index(c) for c in ("Agrupamento", "Time", "Status\nDefeito/US")}
    f_base, f_nome, f_sec = 9.0, 8.0, 8.5

    # O nome do defeito vai inteiro e quebra em várias linhas, e o PowerPoint
    # estica a linha sozinho — era isso que jogava a tabela para fora do slide.
    # A altura de cada linha é estimada pelo tamanho do nome e as que não
    # couberem saem, preservando os defeitos de maior peso, que vêm primeiro.
    def _altura_det(nome):
        linhas = _linhas_do_nome(nome, ws[COL_NOME_DET])
        return max(0.24, linhas * f_nome * 1.25 / 72 + 0.06)

    disponivel = 7.5 - dety - 0.08
    rh_cab = 0.32
    usado_det = rh_cab
    linhas_ok = []
    for row in drows:
        h = rh_cab if row["kind"] == "sucesso" else _altura_det(row["cells"][COL_NOME_DET])
        if usado_det + h > disponivel:
            break
        usado_det += h
        linhas_ok.append((row, h))
    fora_det = len(drows) - len(linhas_ok)
    drows = [r for r, _ in linhas_ok]
    n = 1 + len(drows)

    dt2 = add_tbl(n, len(cols), 0.0, dety, LARG_DET, usado_det)
    _widths(dt2, ws)
    dt2.rows[0].height = Inches(rh_cab)
    for i, (_, h) in enumerate(linhas_ok):
        dt2.rows[1 + i].height = Inches(h)
    if fora_det:
        print(f"  [slide] {d['title'][:28]}: {fora_det} defeito(s) fora do slide, "
              f"descartados por falta de espaço")
    for ci, col in enumerate(cols):
        _cell(dt2.cell(0, ci), col, f_base, bold=True, fg=WHITE, bg=RED)
    for ri, row in enumerate(drows):
        bg = WHITE if row["kind"] == "sucesso" else (WHITE if (ri - 1) % 2 == 0 else GRAY2)
        for ci, val in enumerate(row["cells"]):
            al = "left" if ci == COL_NOME_DET else "center"
            fs = f_nome if ci == COL_NOME_DET else (f_sec if ci in COLS_SEC else f_base)
            bold = (ci == 0 and (row["kind"] == "sucesso" or val == "Falha"))
            _cell(dt2.cell(ri + 1, ci), val, fs, bold=bold, bg=bg, align=al)
    _sem_bordas(dt2)


# ── Slide "Top Ofensores": Top N defeitos/US de uma visão ───────────────────
JORNADA_LABEL_CURTO = {"Base Móvel": "Base"}
TITULO_TOP_OFENSORES = "Consolidado - Top Ofensores - Defeitos com Squad"
COLS_DESCRICAO = ["DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Type", "DFT_Team"]
# Fases que não entram no Top Ofensores: o defeito cancelado ou rejeitado não
# está mais com a squad — são as fases de MOPs. O DFT cancelado que virou US
# entra, pela fase da US (ver `_por_us`).
FASES_FORA_TOP_OFENSORES = fallout_core.FASES_MOPS


def _por_us(pivot2, df, meses):
    """
    `pivot2` sem os defeitos cancelados e rejeitados (FASES_FORA_TOP_OFENSORES)
    e com a coluna "Chave": o número da US quando o DFT tem uma US de Melhoria
    (ver `fallout_core.COL_US_RESOLVIDA`) e o próprio DefectKey nos demais. DFTs
    que apontam para a mesma US viram uma linha, somando os meses; nome, fase,
    data e time já são os da US, iguais entre eles.
    """
    us = (df.dropna(subset=["DFT_US"]).drop_duplicates("DefectKey")
            .set_index("DefectKey")["DFT_US"]) if "DFT_US" in df.columns else {}
    pivot2 = pivot2[~pivot2["DFT_Phase"].fillna("").astype(str).str.strip()
                    .isin(FASES_FORA_TOP_OFENSORES)]
    chave = pivot2["DefectKey"].map(us).fillna(pivot2["DefectKey"])
    agg = {c: "first" for c in COLS_DESCRICAO}
    agg.update({m: "sum" for m in meses if m in pivot2.columns})
    return (pivot2.assign(Chave=chave)
                  .groupby("Chave", sort=False).agg(agg).reset_index())


def _montar_slide_top_ofensores(prs, ns_visao, ns_por_jornada, titulo, top_n=15):
    """
    Adiciona um slide com uma única tabela "Top Ofensores": os `top_n` defeitos/US
    com maior % no mês mais recente, juntando as jornadas brutas da visão em
    `ns_por_jornada` (cada uma vira uma linha "Jornada"), ordenados decrescente.
    `ns_visao` é o relatório da visão inteira (ex.: "Base + Cross Sell"): dá os
    meses, o denominador dos percentuais e a linha Sucesso.
    Reaproveita o mesmo layout/estilo da tabela detalhada de `_montar_slide`.
    """
    import pandas as pd

    meses_tab = ns_visao["meses_tab"]
    MESES_LABEL = ns_visao["MESES_LABEL"]
    labels_tab = [MESES_LABEL[m][:3].capitalize() for m in meses_tab]
    resumo = ns_visao["resumo"]; pct_m_visao = ns_visao["pct_m"]

    import pipeline

    candidatos = []
    for jornada, ns in ns_por_jornada.items():
        # A tabela da jornada recontada nos meses da visão, e não o `pivot2`
        # dela: a visão pode ter outro mês do relatório (ver
        # `pipeline.tabela_ofensores`). Fica o que tem pedido no mês mais recente.
        tabela = pipeline.tabela_ofensores(ns["df"], meses_tab, ns["hoje"])
        tabela = tabela[tabela[meses_tab[-1]] > 0]
        pivot2 = _por_us(tabela, ns["df"], meses_tab)
        fmt_id = ns["fmt_id"]; fmt_ms = ns["fmt_milestone"]
        label_jornada = JORNADA_LABEL_CURTO.get(jornada, jornada)
        for _, row in pivot2.iterrows():
            tipo = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
            tipo_fmt = "US" if tipo == "User Story" else "Defeito"
            id_str = fmt_id({"DefectKey": row["Chave"],
                              "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]})
            ms_fmt = fmt_ms(row["DFT_BugfixMilestone"])
            nome = str(row["DFT_Name"]) if pd.notna(row["DFT_Name"]) else ""
            team = str(row["DFT_Team"])[:35] if pd.notna(row["DFT_Team"]) else ""
            phase = str(row["DFT_Phase"])[:25] if pd.notna(row["DFT_Phase"]) else ""
            # Denominador é o total da visão, não o da jornada de origem: no
            # Base + Cross Sell a tabela soma jornadas de portes diferentes, e
            # dividir cada uma pelo próprio total tornava os números
            # incomparáveis entre linhas — um defeito de uma jornada pequena
            # parecia maior que um de uma jornada grande com muito mais pedidos
            # afetados. É também o que a linha "Sucesso" logo abaixo usa.
            valores = [pct_m_visao(int(row.loc[m]) if m in row.index else 0, m)
                       for m in meses_tab]
            cells = [label_jornada, tipo_fmt, id_str, ms_fmt, nome, team, phase] + \
                    [f"{v:.2f}%".replace(".", ",") for v in valores]
            candidatos.append((valores[-1], cells))   # ordena pelo mês mais recente

    candidatos.sort(key=lambda c: c[0], reverse=True)
    top_linhas = [cells for _, cells in candidatos[:top_n]]

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.55))
    tb.fill.solid(); tb.fill.fore_color.rgb = RED; tb.line.fill.background()
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = titulo
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    def add_tbl(nr, nc, x, y, w, h):
        gf = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
        _no_style(gf.table); return gf.table

    COLS = ["Classificação", "Agrupamento", "Jornada", "Tipo", "Defeito/US",
            "Data Prod\nnão resolvidos", "Nome Defeito", "Time", "Status\nDefeito/US"] + labels_tab
    LARGURA = 13.333                       # ocupa o slide inteiro
    # "Classificação" em 10 pt negrito pede ~0,95": com 0,7 a palavra quebrava
    # no meio ("Classificaç/ão"). A diferença sai do nome, que quebra em linhas.
    ws = [0.95, 1.6, 0.8, 0.6, 0.85, 1.0, 3.5, 1.45, 1.25]
    ws += [(LARGURA - sum(ws)) / len(labels_tab)] * len(labels_tab)
    COL_NOME = 6                           # índice de "Nome Defeito" em COLS

    y = 0.62
    FS_BASE, FS_SEC, FS_NOME = 10.0, 9.0, 8.0
    ALTURA_UTIL = 7.5 - y - 0.08           # até a borda de baixo do slide

    # O nome do defeito não é truncado, então a célula quebra em várias linhas e
    # o PowerPoint estica a linha sozinho — é o que jogava a tabela para fora do
    # slide. Aqui a altura de cada linha é estimada pelo tamanho do nome e as
    # que não couberem são descartadas, mantendo os defeitos de maior peso.
    def _altura(nome):
        linhas = _linhas_do_nome(nome, ws[COL_NOME])
        return max(0.30, linhas * FS_NOME * 1.25 / 72 + 0.06)

    rh_cab = 0.42                          # cabeçalho em duas linhas + Sucesso
    usado = rh_cab * 2
    cabem = []
    for cells in top_linhas:
        h = _altura(cells[4])              # 4 = Nome Defeito dentro de `cells`
        if usado + h > ALTURA_UTIL:
            break
        usado += h
        cabem.append((cells, h))
    cortados = len(top_linhas) - len(cabem)
    top_linhas = [c for c, _ in cabem]

    n_data = 1 + len(top_linhas)           # linha Sucesso + linhas que couberam
    n_total = 1 + n_data                   # + cabeçalho

    tbl = add_tbl(n_total, len(COLS), 0.0, y, LARGURA, usado)
    _widths(tbl, ws)
    tbl.rows[0].height = Inches(rh_cab)
    tbl.rows[1].height = Inches(rh_cab)
    for i, (_, h) in enumerate(cabem):
        tbl.rows[2 + i].height = Inches(h)

    for ci, col in enumerate(COLS):
        _cell(tbl.cell(0, ci), col, FS_BASE, bold=True, fg=WHITE, bg=RED)

    # ── Linha Sucesso (taxa da visão, suas jornadas somadas) ───────────────
    _cell(tbl.cell(1, 0), "Sucesso", FS_BASE, bg=WHITE)
    for ci in range(1, len(COLS) - len(labels_tab)):
        _cell(tbl.cell(1, ci), "", FS_BASE, bg=WHITE)
    for li, m in enumerate(meses_tab):
        suc_pct = _pct_sucesso(resumo, m, pct_m_visao)
        _cell(tbl.cell(1, len(COLS) - len(labels_tab) + li), f"{suc_pct:.2f}%".replace(".", ","), FS_BASE, bg=WHITE)

    # ── Linhas do Top N ──────────────────────────────────────────────────
    for ri, cells in enumerate(top_linhas):
        row_idx = 2 + ri
        bg = WHITE if ri % 2 == 0 else GRAY2
        _cell(tbl.cell(row_idx, 0), "Falha" if ri == 0 else "", FS_BASE, bg=bg, bold=(ri == 0))
        _cell(tbl.cell(row_idx, 1), "Em Tratamento/Avaliação\npela Squad" if ri == 0 else "", FS_SEC, bg=bg)
        for ci, val in enumerate(cells):
            al = "left" if ci == 4 else "center"       # 4 = Nome Defeito
            fs = FS_NOME if ci == 4 else (FS_SEC if ci in (0, 5, 6) else FS_BASE)
            _cell(tbl.cell(row_idx, 2 + ci), val, fs, bg=bg, align=al)

    _sem_bordas(tbl)
    if cortados:
        print(f"  [slide] {titulo}: {cortados} defeito(s) fora do slide, "
              f"descartados por falta de espaço")


# ── Slide "Evolução do Fallout por dia" ──────────────────────────────────────
# Cor fixa por categoria, e não pela posição no dia: a mesma categoria tem a
# mesma cor nos dois slides e em qualquer semana. São as oito cores da paleta
# categórica validada, empilhadas na ordem dela, que é a que passa no teste de
# daltonismo entre vizinhas. O que não tem cor própria soma em "Outras".
# Em Tratamento vem dividido em DFT e US (ver `fallout_core.fallout_por_dia`):
# dois tons do mesmo azul, empilhados juntos, para ler como uma categoria só
# partida em duas — e sem gastar uma nona cor da paleta.
CORES_DIARIO = {
    fallout_core.CAT_TRAT_DFT: "#2a78d6",
    fallout_core.CAT_TRAT_US: "#9ec5f4",
    "Em Avaliação por MOPs": "#eb6834",
    "Tratado - Em avaliação de eficácia": "#1baf7a",
    "Falha Pontual": "#eda100",
    fallout_core.CAT_INCIDENTE: "#e87ba4",
    "Falha no Processo Usuário": "#008300",
    "Falta Associar ao Defeito/US": "#4a3aa7",
    "Em avaliação - Outros times": "#e34948",
}
COR_OUTRAS_DIARIO = "#8c8a85"
# Os nomes da legenda são os da tabela Distribuição dos slides.
ROTULOS_DIARIO = {"Falta Associar ao Defeito/US": "Falta associar problema ao Defeito/US"}
TINTA, TINTA_2, COR_EIXO = "#222222", "#52514e", "#c9c8c4"
ABREV_DIA = {n: a for a, n in fallout_core.ABREV_MES.items()}


def _luminancia(cor):
    """Luminância relativa (WCAG) de uma cor "#rrggbb"."""
    canais = [int(cor[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    r, g, b = [c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4 for c in canais]
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _tinta_sobre(cor):
    """Branco ou tinta escura, o que tiver mais contraste com o fundo `cor`."""
    fundo = _luminancia(cor)
    contra_branco = 1.05 / (fundo + 0.05)
    contra_tinta = (fundo + 0.05) / (_luminancia(TINTA) + 0.05)
    return "white" if contra_branco >= contra_tinta else TINTA


def _grafico_diario(tabela, largura, altura):
    """
    PNG do slide diário: dois painéis alinhados pelo dia, e não um gráfico de
    dois eixos — pedidos (milhares) e fallout (poucos %) não dividem escala, e
    num eixo duplo a altura da linha contra a da barra não quer dizer nada.
    Em cima a linha de pedidos; embaixo as colunas empilhadas do fallout.
    """
    import numpy as np
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.patches import Patch

    dias = list(tabela.index)
    x = np.arange(len(dias))
    pedidos = tabela["Pedidos"].astype(float)
    valido = (pedidos > 0).to_numpy()          # NaN (sem sucessos) também sai

    series = [(c, cor) for c, cor in CORES_DIARIO.items() if c in tabela.columns]
    outras = [c for c in tabela.columns
              if c not in CORES_DIARIO and c not in ("Falhas", "Sucessos", "Pedidos")]
    pct = {c: (tabela[c] / pedidos * 100).where(valido, 0).fillna(0).to_numpy()
           for c, _ in series}
    if outras:
        pct["Outras"] = (tabela[outras].sum(axis=1) / pedidos * 100).where(valido, 0).fillna(0).to_numpy()
        series.append(("Outras", COR_OUTRAS_DIARIO))
    total = sum(pct.values()) if pct else np.zeros(len(dias))

    fig = Figure(figsize=(largura, altura))
    FigureCanvasAgg(fig)
    fig.patch.set_facecolor("white")
    grade = fig.add_gridspec(2, 1, height_ratios=[1, 2.7], hspace=0.28,
                             left=0.02, right=0.98, top=0.86, bottom=0.1)
    ax_p = fig.add_subplot(grade[0])
    ax_f = fig.add_subplot(grade[1], sharex=ax_p)

    # ── Pedidos por dia ──
    ax_p.plot(x[valido], pedidos[valido], color=TINTA, lw=2, marker="o", ms=5,
              markerfacecolor=TINTA, markeredgecolor="white", markeredgewidth=1.2, zorder=3)
    for xi, v in zip(x[valido], pedidos[valido]):
        ax_p.annotate(f"{int(v):,}".replace(",", "."), xy=(xi, v), xytext=(0, 7),
                      textcoords="offset points", ha="center", fontsize=8, color=TINTA_2)
    ax_p.set_ylim(0, (pedidos[valido].max() if valido.any() else 1) * 1.35)
    ax_p.set_title("Pedidos por dia (sucessos + falhas)", loc="left", fontsize=9,
                   color=TINTA_2, pad=2)

    # ── Fallout por dia, empilhado por categoria ──
    teto = (total.max() if len(total) else 1) * 1.2 or 1
    base = np.zeros(len(dias))
    for nome, cor in series:
        h = pct[nome]
        ax_f.bar(x, h, bottom=base, width=0.62, color=cor, edgecolor="white",
                 linewidth=1.0, zorder=2)
        tinta = _tinta_sobre(cor)
        for xi, bi, hi in zip(x, base, h):
            if hi / teto >= 0.06:             # só onde o rótulo cabe no segmento
                ax_f.text(xi, bi + hi / 2, f"{hi:.2f}%".replace(".", ","), ha="center",
                          va="center", fontsize=7, color=tinta)
        base = base + h
    for xi, t, ok in zip(x, total, valido):
        if ok:
            ax_f.text(xi, t + teto * 0.015, f"{t:.2f}%".replace(".", ","), ha="center",
                      va="bottom", fontsize=9, fontweight="bold", color=TINTA)
        else:
            ax_f.text(xi, teto * 0.02, "sem\nsucessos", ha="center", va="bottom",
                      fontsize=7, color=TINTA_2)
    ax_f.set_ylim(0, teto)
    ax_f.set_title("Fallout por dia (% dos pedidos do dia)", loc="left", fontsize=9,
                   color=TINTA_2, pad=2)
    ax_f.set_xticks(x)
    ax_f.set_xticklabels([f"{d.day:02d}/{ABREV_DIA[d.month]}" for d in dias],
                         fontsize=8.5, color=TINTA_2)

    for ax in (ax_p, ax_f):
        ax.set_yticks([])
        ax.tick_params(axis="x", length=0)
        for lado in ("top", "right", "left"):
            ax.spines[lado].set_visible(False)
        ax.spines["bottom"].set_color(COR_EIXO)
    ax_p.tick_params(axis="x", labelbottom=False)
    ax_p.spines["bottom"].set_visible(False)

    # O matplotlib preenche a legenda coluna a coluna; reordenada, ela se lê
    # linha a linha na ordem do empilhamento, de baixo para cima na coluna.
    # Sem nenhuma categoria (janela sem falhas) não há legenda — e o ncol zero
    # derrubava o PowerPoint inteiro.
    ncol = min(len(series), 5)
    if ncol:
        linhas_leg = -(-len(series) // ncol)
        ordem = [r * ncol + c for c in range(ncol) for r in range(linhas_leg)
                 if r * ncol + c < len(series)]
        fig.legend([Patch(facecolor=series[i][1]) for i in ordem],
                   [ROTULOS_DIARIO.get(series[i][0], series[i][0]) for i in ordem],
                   loc="upper center", bbox_to_anchor=(0.5, 0.995), frameon=False,
                   ncol=ncol, fontsize=8.5, labelcolor=TINTA,
                   handlelength=1.1, handleheight=1.1, columnspacing=1.6)
    fig.text(0.02, 0.015,
             f"Últimos {len(dias)} dias até {dias[-1]:%d/%m} (último dia completo). "
             "Cada pedido está na categoria que tem hoje.",
             fontsize=7.5, color=TINTA_2)

    png = io.BytesIO()
    fig.savefig(png, dpi=200, facecolor="white", format="png")
    png.seek(0)
    return png


def _montar_slide_diario(prs, titulo, tabela):
    """Slide com a evolução diária do fallout de uma visão (ver `_grafico_diario`)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.55))
    tb.fill.solid(); tb.fill.fore_color.rgb = RED; tb.line.fill.background()
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = titulo
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    LARG, ALT = 12.93, 6.8
    slide.shapes.add_picture(_grafico_diario(tabela, LARG, ALT),
                             Inches(0.2), Inches(0.62), width=Inches(LARG))


# ── Slide "Mapa de calor de erros pontuais" ──────────────────────────────────
# A tabela da aba Análise de Erros do app (mensagem × dia), só com as Falhas
# Pontuais, na janela do slide diário e com as mensagens mais frequentes. No app
# a cor é verde→vermelho; no slide é uma escala de um tom só (azul, do claro ao
# escuro), em faixas com legenda — leitura de "quanto", e não de alerta, e sem
# disputar com o vermelho da marca nas faixas de título.
MAPA_CATEGORIA = "Falha Pontual"
MAPA_TOP_MENSAGENS = 15
MAPA_LIMITE_MENSAGEM = 90          # uma linha na coluna de 4,95" em 8 pt
# Rampa sequencial azul (passos 100, 200, 300, 400, 550 e 700 da paleta).
MAPA_RAMPA = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#1c5cab", "#0d366b"]
MAPA_DIAS_SEMANA = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]


def _faixas_mapa(valores, n=len(MAPA_RAMPA)):
    """
    Limites inferiores das faixas de cor, em progressão geométrica até o maior
    valor (1, 2, 3, 5, 7, 11… para um máximo de 17). As contagens são muito
    assimétricas: em escala linear quase tudo ficava no tom mais claro, e por
    quantil — como na primeira versão — quase tudo no mais escuro ("5+"),
    sem distinguir 5 de 14.
    """
    maximo = max([x for x in valores if x > 0], default=0)
    if maximo <= 1:
        return [1]
    return sorted({max(1, round(maximo ** (i / n))) for i in range(n)})


def _faixa_de(valor, limites):
    """Índice da faixa do valor (0 = a mais clara); -1 para zero."""
    if valor <= 0:
        return -1
    return max(i for i, lim in enumerate(limites) if valor >= lim)


def tabela_mapa_calor(df, corte, base_dir, visao, dias=None, top_n=MAPA_TOP_MENSAGENS,
                      hoje=None):
    """
    Falhas Pontuais por mensagem de erro e dia, nos `dias` dias até `corte`: as
    `top_n` mensagens mais frequentes da janela, uma linha somando as demais e o
    total do dia. A categoria de cada pedido é a da Distribuição
    (`fallout_core.categorias_por_pedido`), e a mensagem é a resumida
    (`comparacao.resumir_mensagem`); variações que só mudam no trecho cortado
    viram uma.

    O % é sobre o total de pedidos da janela (falhas de todas as categorias +
    sucessos das jornadas da visão), a mesma base da taxa de fallout.

    Devolve um dict com "dias", "linhas" (mensagem, contagens, total, % dos
    pedidos), "demais", "total_dia", "falhas" (as pontuais), "falhas_periodo"
    (todas as falhas da janela) e "pedidos" (None se faltar o sucesso de algum
    dia: o % sairia inflado).
    """
    import datetime as _dt
    import pandas as pd
    import comparacao

    dias = dias or fallout_core.DIAS_FALLOUT_DIARIO
    janela = [corte - _dt.timedelta(days=i) for i in range(dias - 1, -1, -1)]
    dia = fallout_core.dia_local(df["CreatedDate"]).dt.date
    no_periodo = df[dia.isin(set(janela))].copy()
    falhas_periodo = len(no_periodo)
    sucessos = 0
    for jornada in fallout_core.resolver_jornadas(visao, base_dir):
        da_jornada = fallout_core.sucessos_por_dia(base_dir, jornada)
        if any(d not in da_jornada for d in janela):
            sucessos = None
            break
        sucessos += sum(da_jornada[d] for d in janela)
    pedidos = falhas_periodo + sucessos if sucessos is not None else None
    categoria, _subgrupo, _sobrepostos = fallout_core.categorias_por_pedido(no_periodo, hoje)
    no_periodo = no_periodo[categoria == MAPA_CATEGORIA].copy()
    no_periodo["_dia"] = dia[no_periodo.index]
    bruta = (no_periodo["ErrorHandled__c"].astype("string").str.strip()
             .replace("", pd.NA).fillna("(sem mensagem de erro)").astype(str))
    resumo = {m: comparacao.resumir_mensagem(m, MAPA_LIMITE_MENSAGEM) for m in bruta.unique()}
    no_periodo["_msg"] = bruta.map(resumo)

    contagem = (no_periodo.groupby(["_msg", "_dia"]).size().unstack("_dia", fill_value=0)
                .reindex(columns=janela, fill_value=0))
    contagem["_total"] = contagem.sum(axis=1)
    contagem = contagem.sort_values("_total", ascending=False, kind="stable")
    falhas = int(contagem["_total"].sum())

    topo = contagem.head(top_n)
    resto = contagem.iloc[top_n:]
    linhas = [{"mensagem": msg, "contagens": [int(r[d]) for d in janela],
               "total": int(r["_total"]),
               "pct": r["_total"] / pedidos * 100 if pedidos else None}
              for msg, r in topo.iterrows()]
    demais = None
    if len(resto):
        demais = {"mensagem": f"Demais mensagens ({len(resto)})",
                  "contagens": [int(resto[d].sum()) for d in janela],
                  "total": int(resto["_total"].sum()),
                  "pct": resto["_total"].sum() / pedidos * 100 if pedidos else None}
    total_dia = [int(contagem[d].sum()) for d in janela]
    return {"dias": janela, "linhas": linhas, "demais": demais,
            "total_dia": total_dia, "falhas": falhas, "falhas_periodo": falhas_periodo,
            "pedidos": pedidos}


def _pct_pequeno(valor):
    """`fmt_pct`, mas "<0,01%" no que é positivo e arredondaria para 0,00%."""
    if valor is not None and 0 < valor < 0.005:
        return "<0,01%"
    return fallout_core.fmt_pct(valor)


def _borda_branca(cell, largura_emu=19050):
    """As quatro bordas da célula em branco, 1,5 pt."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for tag in ("lnB", "lnT", "lnR", "lnL"):   # insere no início → ordem final lnL,lnR,lnT,lnB
        ln = parse_xml(f'<a:{tag} {nsdecls("a")} w="{largura_emu}" cap="flat">'
                       f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:{tag}>')
        tcPr.insert(0, ln)


def _montar_slide_mapa_calor(prs, titulo, mapa):
    """
    Slide do mapa de calor (ver `tabela_mapa_calor`): tabela nativa, editável,
    com a célula de cada dia na cor da faixa e o número dentro.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.55))
    tb.fill.solid(); tb.fill.fore_color.rgb = RED; tb.line.fill.background()
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = titulo
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    dias, linhas = mapa["dias"], mapa["linhas"]
    pedidos = mapa["pedidos"]
    pct_pontual = mapa["falhas"] / pedidos * 100 if pedidos else None
    pct_top = sum(l["pct"] for l in linhas) if pedidos else None
    sub = slide.shapes.add_textbox(Inches(0.12), Inches(0.58), Inches(13), Inches(0.3))
    r = sub.text_frame.paragraphs[0].add_run()
    r.text = (f"Falhas pontuais por dia · {dias[0]:%d/%m} a {dias[-1]:%d/%m} · "
              f"{fallout_core.fmt_inteiro(mapa['falhas'])} falhas pontuais"
              + (f" em {fallout_core.fmt_inteiro(pedidos)} pedidos "
                 f"({fallout_core.fmt_pct(pct_pontual)} dos pedidos)" if pedidos
                 else " (sem o total de pedidos de algum dia: % não calculado)")
              + (f" · as {len(linhas)} mensagens mais frequentes somam "
                 f"{fallout_core.fmt_pct(pct_top)} dos pedidos"
                 if mapa["demais"] and pedidos else ""))
    r.font.size = Pt(10); r.font.italic = True; r.font.color.rgb = FG; r.font.name = "Calibri"
    if not linhas:
        vazio = slide.shapes.add_textbox(Inches(0.5), Inches(3.3), Inches(12.3), Inches(0.6))
        pv = vazio.text_frame.paragraphs[0]; pv.alignment = PP_ALIGN.CENTER
        rv = pv.add_run(); rv.text = "Nenhuma falha pontual no período."
        rv.font.size = Pt(16); rv.font.color.rgb = FG; rv.font.name = "Calibri"
        return

    limites = _faixas_mapa([c for l in linhas for c in l["contagens"]])
    # Os tons mais escuros da rampa, um por faixa; com uma faixa só (máximo 1)
    # o tom do meio, para não pintar tudo de quase preto.
    cores = MAPA_RAMPA[-len(limites):] if len(limites) > 1 else [MAPA_RAMPA[3]]
    rampa = [RGBColor.from_string(c.lstrip("#")) for c in cores]
    tinta = [WHITE if _tinta_sobre(c) == "white" else FG for c in cores]

    W_MSG, W_DIA, W_TOT, W_PCT = 4.95, 0.47, 0.6, 0.75
    ws = [W_MSG] + [W_DIA] * len(dias) + [W_TOT, W_PCT]
    X0, Y0 = (13.333 - sum(ws)) / 2, 0.95
    n_linhas = 1 + len(linhas) + (1 if mapa["demais"] else 0) + 1
    RH, RH_CAB = 0.315, 0.42
    gf = slide.shapes.add_table(n_linhas, len(ws), Inches(X0), Inches(Y0), Inches(sum(ws)),
                                Inches(RH_CAB + RH * (n_linhas - 1)))
    tbl = gf.table
    _no_style(tbl); _widths(tbl, ws)
    tbl.rows[0].height = Inches(RH_CAB)
    for i in range(1, n_linhas):
        tbl.rows[i].height = Inches(RH)

    FS, FS_CAB, FS_MSG = 8.5, 8.5, 8.0
    CINZA = RGBColor(0xF2, 0xF2, 0xF2)
    cab = (["Mensagem de erro"]
           + [f"{MAPA_DIAS_SEMANA[d.weekday()]}\n{d:%d/%m}" for d in dias]
           + ["Total", "% dos\npedidos"])
    for ci, txt in enumerate(cab):
        _cell(tbl.cell(0, ci), txt, FS_CAB, bold=True, fg=WHITE, bg=RED,
              align="left" if ci == 0 else "center")

    def _linha(ri, item, fundo, negrito=False, aquecer=False):
        _cell(tbl.cell(ri, 0), item["mensagem"], FS_MSG, bold=negrito, bg=fundo, align="left")
        for di, v in enumerate(item["contagens"]):
            f = _faixa_de(v, limites) if aquecer else -1
            _cell(tbl.cell(ri, 1 + di), fallout_core.fmt_inteiro(v) if v else "", FS,
                  bold=negrito, fg=tinta[f] if f >= 0 else FG,
                  bg=rampa[f] if f >= 0 else fundo)
        base = 1 + len(dias)
        _cell(tbl.cell(ri, base), fallout_core.fmt_inteiro(item["total"]), FS, bold=True, bg=fundo)
        _cell(tbl.cell(ri, base + 1),
              _pct_pequeno(item.get("pct")),
              FS, bg=fundo)

    ri = 1
    for item in linhas:
        _linha(ri, item, WHITE, aquecer=True)
        ri += 1
    if mapa["demais"]:
        _linha(ri, mapa["demais"], CINZA)
        ri += 1
    _linha(ri, {"mensagem": "Total do dia", "contagens": mapa["total_dia"],
                "total": mapa["falhas"], "pct": pct_pontual}, CINZA, negrito=True)
    # Borda branca fina em toda célula: separa as vizinhas de mesma cor, que
    # sem ela viravam uma mancha só.
    for row in tbl.rows:
        for cell in row.cells:
            _borda_branca(cell)

    # Legenda da escala: uma caixa por faixa, com o intervalo de ocorrências.
    y_leg = Y0 + RH_CAB + RH * (n_linhas - 1) + 0.15
    lbl = slide.shapes.add_textbox(Inches(X0), Inches(y_leg), Inches(1.9), Inches(0.26))
    r = lbl.text_frame.paragraphs[0].add_run(); r.text = "Ocorrências no dia:"
    r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = FG; r.font.name = "Calibri"
    for i, lim in enumerate(limites):
        fim = limites[i + 1] - 1 if i + 1 < len(limites) else None
        if len(limites) == 1:   # faixa única só existe com máximo 1
            txt = f"{lim}"
        else:
            txt = f"{lim}" if fim == lim else (f"{lim}–{fim}" if fim else f"{lim}+")
        caixa = slide.shapes.add_shape(1, Inches(X0 + 1.55 + i * 0.85), Inches(y_leg + 0.02),
                                       Inches(0.8), Inches(0.22))
        caixa.fill.solid(); caixa.fill.fore_color.rgb = rampa[i]; caixa.line.fill.background()
        caixa.shadow.inherit = False
        pp = caixa.text_frame.paragraphs[0]; pp.alignment = PP_ALIGN.CENTER
        rr = pp.add_run(); rr.text = txt
        rr.font.size = Pt(8); rr.font.color.rgb = tinta[i]; rr.font.name = "Calibri"
    nota = slide.shapes.add_textbox(Inches(X0 + 1.6 + len(limites) * 0.85), Inches(y_leg),
                                    Inches(6.5), Inches(0.26))
    r = nota.text_frame.paragraphs[0].add_run()
    r.text = ("Falha Pontual: a mesma categoria da Distribuição (defeito 999999). "
              "Mensagens agrupadas pelo texto resumido.")
    r.font.size = Pt(8); r.font.italic = True; r.font.color.rgb = FG; r.font.name = "Calibri"
