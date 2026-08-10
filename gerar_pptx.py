# -*- coding: utf-8 -*-
"""Gera o relatório em PowerPoint editável (tabelas nativas + gráfico como imagem)
para uma jornada, reaproveitando a lógica do pipeline.py.

Uso como módulo:  from gerar_pptx import gerar_pptx; buf = gerar_pptx("Base Móvel", base_dir)
Uso via CLI:      python gerar_pptx.py "Base Móvel"
"""
import os, re, sys, io, tempfile, contextlib
import matplotlib
matplotlib.use("Agg")
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

RED = RGBColor(0xC0, 0x39, 0x2B); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0xF2, 0xF2, 0xF2); GRAY2 = RGBColor(0xF5, 0xF5, 0xF5)
FG = RGBColor(0x22, 0x22, 0x22)


# ── Extração dos dados via pipeline ──────────────────────────────────────────
def _extrair(jornada, base_dir):
    """Executa o pipeline na pasta base_dir e devolve (data, caminho_chart_png)."""
    pipeline_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pipeline.py")
    with open(pipeline_path, encoding="utf-8") as f:
        src = f.read()
    src = re.sub(r"plt\.close\([^)]*\)", "pass", src)   # mantém figuras vivas p/ recortar

    old_cwd = os.getcwd(); old_argv = sys.argv
    os.chdir(base_dir); sys.argv = ["pipeline.py", jornada]
    try:
        ns = {"__name__": "__pipeline__"}
        # redireciona stdout p/ buffer unicode (evita crash de cp1252 com os prints ✓/→)
        with contextlib.redirect_stdout(io.StringIO()):
            exec(compile(src, "pipeline.py", "exec"), ns)
    finally:
        os.chdir(old_cwd); sys.argv = old_argv

    import pandas as pd
    resumo = ns["resumo"]; ML = ns["MESES_LABEL"]; reducao = ns["reducao"]
    meses_vol = ns["meses_vol"]
    fi = lambda v: f"{int(v):,}".replace(",", ".")
    fp = lambda v: f"{v:.2f}%".replace(".", ",")

    volume = {"months": [ML[m].capitalize() for m in meses_vol], "rows": [
        ["Vendas com Sucesso"] + [fi(resumo.loc[m, "Sucessos"]) for m in meses_vol],
        ["Falha (Análise Técnica)"] + [fi(resumo.loc[m, "Falhas"]) for m in meses_vol],
        ["Fallout Rate"] + [fp(resumo.loc[m, "Pct"]) for m in meses_vol],
    ]}
    dist = {"title": f"Distribuição Fallout ({ns['fallout_pct']:.2f}%)", "rows": [
        {"label": lbl, "val": fp(val), "sub": sub, "bold": bold}
        for (lbl, val, sub, bold) in ns["dist_items"]]}
    plan_rows = []
    for _, r in reducao.iterrows():
        dfts = r["DFTs"].split("\n")
        plan_rows.append({"date": r["MilestoneDate"].strftime("%d/%m/%Y"),
                          "dfts": [{"id": d.strip(), "pct": fp(r["Pct"] / len(dfts))} for d in dfts]})
    plan = {"title": f"Planejamento Redução ({reducao['Pct_plena'].sum():.2f}% pleno)", "rows": plan_rows}

    COL_NAMES = ns["COL_NAMES"]; meses_tab = ns["meses_tab"]
    pivot2 = ns["pivot2"].reset_index(drop=True)
    fmt_id = ns["fmt_id"]; fmt_ms = ns["fmt_milestone"]; pct_m = ns["pct_m"]
    n_fixas = len(COL_NAMES) - len(meses_tab)
    det_rows = []
    suc = ["Sucesso"] + [""] * (n_fixas - 1)
    for m in meses_tab:
        sv = pct_m(resumo.loc[m, "Sucessos"] if m in resumo.index else 0, m)
        suc.append(f"{sv:.2f}%".replace(".", ","))
    det_rows.append({"cells": suc, "kind": "sucesso"})
    for ri in range(len(pivot2)):
        row = pivot2.iloc[ri]
        tipo = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
        cells = [
            "Falha" if ri == 0 else "",
            "Em Tratamento/Avaliação pela Squad" if ri == 0 else "",
            "US" if tipo == "User Story" else "Defeito",
            fmt_id({"DefectNumber__c": row["DefectNumber__c"], "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]}),
            fmt_ms(row["DFT_BugfixMilestone"]),
            str(row["DFT_Name"])[:80] if pd.notna(row["DFT_Name"]) else "",
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
    chart_png = os.path.join(tempfile.mkdtemp(), "chart.png")
    fig.savefig(chart_png, dpi=200, bbox_inches=ext, pad_inches=0.08, facecolor="white")

    data = {"title": f"{jornada} – Realização e Projeção de Fallout",
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

    tb = slide.shapes.add_shape(1, Inches(0), Inches(0.05), Inches(13.333), Inches(0.5))
    tb.fill.solid(); tb.fill.fore_color.rgb = RED; tb.line.fill.background()
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = d["title"]
    r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = "Calibri"

    def add_tbl(nr, nc, x, y, w, h):
        gf = slide.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h))
        _no_style(gf.table); return gf.table

    LX, LW = 0.18, 5.75; y = 0.68
    vm = d["volume"]["months"]; drows_d = d["dist"]["rows"]; prows = d["plan"]["rows"]
    n_plan = 1 + sum(1 + len(r["dfts"]) for r in prows)
    left_total = (2 + len(d["volume"]["rows"])) + (1 + len(drows_d)) + n_plan
    if left_total <= 24:   RH, FD, FT = 0.135, 6.5, 7.5
    elif left_total <= 30: RH, FD, FT = 0.115, 5.5, 6.3
    else:                  RH, FD, FT = 0.102, 5.0, 5.8
    GAP = 0.04

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
        lbl = ("     " + row["label"]) if row["sub"] else row["label"]
        _cell(dt.cell(1 + ri, 0), lbl, FD, bold=row["bold"], bg=bg, align="left")
        _cell(dt.cell(1 + ri, 1), row["val"], FD, bold=row["bold"], bg=bg, align="right")
    y += RH * (1 + len(drows_d)) + GAP

    pt = add_tbl(n_plan, 2, LX, y, LW, RH * n_plan)
    _widths(pt, [LW - 1.5, 1.5]); _rowh(pt, RH)
    pt.cell(0, 0).merge(pt.cell(0, 1))
    _cell(pt.cell(0, 0), d["plan"]["title"], FT, bold=True, fg=WHITE, bg=RED)
    ridx = 1
    for pr in prows:
        pt.cell(ridx, 0).merge(pt.cell(ridx, 1))
        _cell(pt.cell(ridx, 0), pr["date"], FD, bold=True, bg=GRAY, align="left"); ridx += 1
        for dft in pr["dfts"]:
            _cell(pt.cell(ridx, 0), "     " + dft["id"], FD, bg=WHITE, align="left")
            _cell(pt.cell(ridx, 1), dft["pct"], FD, bg=WHITE, align="right"); ridx += 1
    left_bottom = y + RH * n_plan

    # bloco esquerdo sem bordas (igual ao PNG)
    for _t in (vt, dt, pt):
        _sem_bordas(_t)

    CHART_X, CHART_Y, CHART_W = 6.15, 0.72, 6.95
    slide.shapes.add_picture(chart_png, Inches(CHART_X), Inches(CHART_Y), width=Inches(CHART_W))
    chart_bottom = CHART_Y + CHART_W / 2.08

    cols = d["detalhe"]["cols"]; drows = d["detalhe"]["rows"]
    dety = max(left_bottom, chart_bottom) + GAP
    n = 1 + len(drows)
    rh = max(0.085, min(0.24, (7.42 - dety) / n))
    if rh >= 0.15:   f_base, f_nome, f_sec = 7.0, 5.5, 6.2
    elif rh >= 0.11: f_base, f_nome, f_sec = 6.0, 5.0, 5.5
    else:            f_base, f_nome, f_sec = 5.2, 4.4, 4.8
    ws = [0.95, 1.5, 0.7, 0.95, 1.15, 3.55, 1.7, 1.4]
    ws += [(13.0 - sum(ws)) / len(cols[8:])] * len(cols[8:])
    dt2 = add_tbl(n, len(cols), 0.18, dety, 13.0, rh * n)
    _widths(dt2, ws); _rowh(dt2, rh)
    for ci, col in enumerate(cols):
        _cell(dt2.cell(0, ci), col, f_base, bold=True, fg=WHITE, bg=RED)
    for ri, row in enumerate(drows):
        bg = WHITE if row["kind"] == "sucesso" else (WHITE if (ri - 1) % 2 == 0 else GRAY2)
        for ci, val in enumerate(row["cells"]):
            al = "left" if ci == 5 else "center"
            fs = f_nome if ci == 5 else (f_sec if ci in (1, 6, 7) else f_base)
            bold = (ci == 0 and (row["kind"] == "sucesso" or val == "Falha"))
            _cell(dt2.cell(ri + 1, ci), val, fs, bold=bold, bg=bg, align=al)
    _borda_clara(dt2)   # bordas finas em cinza claro (igual ao PNG)


def gerar_pptx(jornada, base_dir="."):
    """Gera o relatório PPTX da jornada e devolve os bytes (BytesIO)."""
    data, chart_png = _extrair(jornada, base_dir)
    prs = Presentation()
    prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
    _montar_slide(prs, data, chart_png)
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf


if __name__ == "__main__":
    j = sys.argv[1] if len(sys.argv) > 1 else "Base Móvel"
    out = f"relatorio_{j.replace(' + ', '_').replace(' ', '_')}.pptx"
    buf = gerar_pptx(j, ".")
    with open(out, "wb") as f:
        f.write(buf.getbuffer())
    print("Salvo:", out)
