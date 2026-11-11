# -*- coding: utf-8 -*-
"""
Previsto x Realizado: por que o fallout de um mês não ficou na projeção feita
no mês anterior.

A projeção de out/26 que saiu na fotografia de 30/09 (2,12% no Base + Cross
Sell) supõe, categoria por categoria (ver `pipeline.calcular_relatorio`):

  * Resolvido, Tratado e Incidente: zero — são o que não se repete;
  * Em Tratamento: o % de set/26 menos o que as entregas planejadas tiram de
    out/26 (cheio se a entrega é anterior ao mês, pro-rata se é dentro dele);
  * as demais (Falha Pontual, Falta Associar, MOPs, …): o mesmo % de set/26.

A soma dessas suposições é exatamente a projeção. O realizado é a distribuição
do mês com os dados de agora, e a diferença de cada categoria diz onde o mês
saiu do previsto. Em Tratamento abre por defeito (entrega que não veio, data
que mudou, defeito novo) e as categorias sem defeito abrem por mensagem de erro.

Como o comparacao.py, não depende de Streamlit: recebe a fotografia e os dados
de agora e devolve tabelas.
"""
import dataclasses

import pandas as pd

import comparacao
import distribuicao
import fallout_core
import fotografia

# Categorias que a projeção supõe zeradas no mês seguinte.
NAO_RECORRENTES = comparacao.CATEGORIAS_NAO_RECORRENTES
# As que se explicam pela mensagem de erro, como na Distribuição.
POR_MENSAGEM = distribuicao.POR_MENSAGEM
LIMITE_MENSAGEM = 110
# Diferença a partir da qual uma linha entra no texto.
RELEVANTE = 0.01
TOLERANCIA = 1e-6


@dataclasses.dataclass
class PrevistoRealizado:
    visao: str
    mes: int                    # o mês avaliado (AAAAMM)
    foto_id: str                # a fotografia que fez a previsão
    mes_base: int               # o mês do relatório nela, de onde a previsão partiu
    previsto: float
    realizado: float
    dias: int                   # dias com falha no mês, nos dados de agora
    categorias: pd.DataFrame
    defeitos: pd.DataFrame
    mensagens: dict             # categoria -> DataFrame por mensagem
    texto: list
    conferencia: dict


def _rotulo(mes):
    return fallout_core.MESES_LABEL[int(mes)]


def _projecao_de(rel, mes):
    """O valor que a fotografia projetou para `mes`, ou None."""
    return dict(zip(rel["proj_labels"][1:], rel["proj_vals"][1:])).get(_rotulo(mes))


def fotografias_com_previsao(base_dir, visao, mes):
    """
    As fotografias que projetaram `mes` para a visão: as de mês do relatório
    anterior a `mes` com `mes` entre os projetados.

    A primeira é a padrão — a previsão do fechamento do mês anterior: a última
    fotografia com o mês anterior como relatório tirada antes de o mês começar
    (30/09 para out/26); sem ela, a última com o mês anterior como relatório,
    mesmo já dentro do mês (02/09 para set/26, que ainda reportava ago/26).
    As demais vêm depois, da mais recente para a mais antiga.
    """
    anterior = fallout_core.mes_anterior(int(mes))
    candidatas = []
    for f in fotografia.listar(base_dir):                 # mais recente primeiro
        rel = fotografia.ler_meta(base_dir, f["id"])["relatorios"].get(visao)
        if not (rel and int(rel["mes_atual"]) < int(mes)
                and _projecao_de(rel, mes) is not None):
            continue
        do_anterior = int(rel["mes_atual"]) == anterior
        antes_do_mes = fallout_core.chave_mes(f["corte"].year, f["corte"].month) < int(mes)
        prioridade = 0 if do_anterior and antes_do_mes else (1 if do_anterior else 2)
        candidatas.append((prioridade, len(candidatas), f["id"]))
    return [i for _p, _ordem, i in sorted(candidatas)]


def _dia(valor):
    ts = pd.to_datetime(valor, errors="coerce")
    return None if pd.isna(ts) else ts.date()


def _fmt_dia(data):
    return "sem data" if data is None else f"{data:%d/%m/%Y}"


def _por_categoria(pedidos, total):
    return (pedidos["Categoria"].value_counts() / total * 100) if total else pd.Series(dtype=float)


def _atributos(pedidos):
    """Nome, fase, entrega e US de cada defeito, pelo primeiro pedido dele."""
    cols = [c for c in ("DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_US", "DFT_Type")
            if c in pedidos.columns]
    com = pedidos[pedidos["DefectKey"].notna()]
    return com.groupby(com["DefectKey"].astype(str))[cols].first()


def _explicar_defeito(chave, linha, cat_antes, cat_agora, hoje):
    """Por que um defeito de Em Tratamento saiu do previsto, em poucas palavras."""
    ms_a, ms_b = linha["Entrega antes"], linha["Entrega agora"]
    if not linha["Na previsão"] and linha["Entrega prevista (pp)"] == 0:
        de_onde = cat_antes.get(chave)
        return (f"não estava em Em Tratamento no mês da previsão"
                + (f" (estava em {de_onde})" if de_onde else " — defeito novo"))
    partes = []
    if linha["Entrega prevista (pp)"] > 0:
        if ms_a != ms_b:
            partes.append(f"entrega mudou de {_fmt_dia(ms_a)} para {_fmt_dia(ms_b)}")
        elif ms_b is not None and ms_b > hoje:
            partes.append(f"entrega de {_fmt_dia(ms_b)} ainda não chegou")
        elif linha["Realizado"] > 0:
            partes.append(f"entregue em {_fmt_dia(ms_b)}, mas ainda com pedidos em tratamento")
        else:
            partes.append(f"entregue em {_fmt_dia(ms_b)}")
    elif ms_a != ms_b:
        partes.append(f"entrega {_fmt_dia(ms_a)} → {_fmt_dia(ms_b)}")
    fa, fb = linha["Fase antes"], linha["Fase agora"]
    if isinstance(fa, str) and isinstance(fb, str) and fa != fb:
        partes.append(f"fase {fa} → {fb}")
    if linha["Realizado"] == 0 and linha["Previsto"] > 0:
        onde = cat_agora.get(chave)
        partes.append(f"pedidos do mês em {onde}" if onde else "sem pedido no mês até agora")
    elif not partes:
        partes.append("mais pedidos que no mês da previsão" if linha["Diferença (pp)"] > 0
                      else "menos pedidos que no mês da previsão")
    return "; ".join(partes)


def explicar(base_dir, foto_id, visao, df_agora, resumo_agora, mes, hoje=None):
    """
    O Previsto x Realizado de `mes` na visão: a previsão da fotografia
    `foto_id` contra os dados de agora (`df_agora`/`resumo_agora`, os de
    `fallout_core.carregar_base`).
    """
    hoje_dt = fallout_core.hoje_utc(hoje).date()
    foto = fotografia.carregar(base_dir, foto_id)
    lado = comparacao._lado(foto, visao)
    if lado is None:
        raise ValueError(f"{fallout_core.nome_exibicao(visao)} não está na fotografia {foto_id}.")
    rel = lado.rel
    mes_base = int(rel["mes_atual"])
    previsto = _projecao_de(rel, mes)
    if previsto is None:
        raise ValueError(f"A fotografia {foto_id} não projetou {_rotulo(mes)}.")

    # ── Previsto, categoria por categoria ────────────────────────────────────
    t_base = lado.total(mes_base)
    ped_base = lado.pedidos[lado.pedidos["Mes"] == mes_base]
    cats_base = _por_categoria(ped_base, t_base)
    contrib = comparacao._contribuicoes(lado, mes)["Contribuição"]
    prev = {c: (0.0 if c in NAO_RECORRENTES else float(v)) for c, v in cats_base.items()}
    prev[fallout_core.CAT_TRATAMENTO] = prev.get(fallout_core.CAT_TRATAMENTO, 0.0) - contrib.sum()
    # O fechamento digitado no histórico não precisa bater com a extração: a
    # diferença entra na projeção e ganha linha própria.
    ajuste_base = rel["fallout_pct"] - cats_base.sum()

    # ── Realizado ────────────────────────────────────────────────────────────
    t_agora = fallout_core.total_do_mes(resumo_agora, mes)
    _df_mes, cats, _t = fallout_core.categorizar(df_agora, mes, hoje=hoje)
    real = {c: len(p) / t_agora * 100 for c, p in cats.items() if len(p) and t_agora}
    realizado = float(resumo_agora.loc[mes, "Pct"])
    ajuste_agora = realizado - sum(real.values())

    linhas = []
    for cat in comparacao._ordenar_categorias(set(prev) | set(real)):
        p, r = prev.get(cat, 0.0), real.get(cat, 0.0)
        if p == 0 and r == 0:
            continue
        if cat in NAO_RECORRENTES:
            sup = "zero: não se repete no mês seguinte"
        elif cat == fallout_core.CAT_TRATAMENTO:
            sup = (f"{_rotulo(mes_base)} ({cats_base.get(cat, 0):.2f}%) menos as entregas "
                   f"({contrib.sum():.2f} pp)").replace(".", ",")
        else:
            sup = f"igual a {_rotulo(mes_base)}"
        linhas.append({"Categoria": cat, "Previsto": p, "Realizado": r,
                       "Diferença (pp)": r - p, "O que a previsão supôs": sup})
    if abs(ajuste_base) >= TOLERANCIA or abs(ajuste_agora) >= TOLERANCIA:
        linhas.append({"Categoria": "Fechamento digitado x extração", "Previsto": ajuste_base,
                       "Realizado": ajuste_agora, "Diferença (pp)": ajuste_agora - ajuste_base,
                       "O que a previsão supôs": "a taxa digitada no histórico menos a soma "
                                                 "das categorias da extração"})
    else:
        ajuste_base = ajuste_agora = 0.0
    tabela = pd.DataFrame(linhas)
    piso = previsto - tabela["Previsto"].sum()
    if abs(piso) >= TOLERANCIA:
        # A projeção não fica abaixo de zero (ver pipeline).
        tabela.loc[len(tabela)] = {"Categoria": "Piso de zero da projeção", "Previsto": piso,
                                   "Realizado": 0.0, "Diferença (pp)": -piso,
                                   "O que a previsão supôs": "a projeção não fica negativa"}

    # ── Em Tratamento, por defeito ───────────────────────────────────────────
    trat = fallout_core.CAT_TRATAMENTO
    ped_trat_base = ped_base[ped_base["Categoria"] == trat]
    base_def = (ped_trat_base["DefectKey"].astype(str).value_counts() / t_base * 100
                if t_base else pd.Series(dtype=float))
    agora_trat = cats.get(trat, df_agora.iloc[0:0])
    agora_def = (agora_trat["DefectKey"].astype(str).value_counts() / t_agora * 100
                 if t_agora else pd.Series(dtype=float))
    chaves = sorted(set(base_def.index) | set(agora_def.index) | set(contrib.index.astype(str)))
    attr_a, attr_b = _atributos(ped_base), _atributos(df_agora[df_agora["Mes"] == mes])
    attr_b_tudo = _atributos(df_agora)
    # Em que categoria cada defeito estava (no mês da previsão) e está (agora).
    cat_antes = (ped_base[ped_base["Categoria"] != trat]
                 .assign(K=ped_base["DefectKey"].astype(str)).groupby("K")["Categoria"].first()
                 .to_dict())
    pedidos_mes = pd.concat([p.assign(Categoria=c) for c, p in cats.items() if len(p)]) \
        if any(len(p) for p in cats.values()) else df_agora.iloc[0:0].assign(Categoria=None)
    cat_agora = (pedidos_mes[pedidos_mes["Categoria"] != trat]
                 .assign(K=pedidos_mes["DefectKey"].astype(str)).groupby("K")["Categoria"].first()
                 .to_dict())

    def _get(attr, chave, col):
        """O atributo do defeito, ou None — nunca o NA do pandas, que quebra o `or`."""
        if chave not in attr.index or col not in attr.columns:
            return None
        valor = attr.loc[chave, col]
        return None if pd.isna(valor) or str(valor).strip() in ("", "nan") else valor

    linhas_d = []
    for k in chaves:
        if k in ("", "nan", "None", "-1"):
            continue
        entrega_prev = float(contrib.get(k, 0.0))
        p = float(base_def.get(k, 0.0)) - entrega_prev
        r = float(agora_def.get(k, 0.0))
        if round(p, 2) == 0 and round(r, 2) == 0 and entrega_prev == 0:
            continue
        nome = _get(attr_b, k, "DFT_Name") or _get(attr_a, k, "DFT_Name") or _get(attr_b_tudo, k, "DFT_Name")
        us = _get(attr_b, k, "DFT_US") or _get(attr_a, k, "DFT_US")
        tipo = _get(attr_b, k, "DFT_Type") or _get(attr_a, k, "DFT_Type")
        # Como na Distribuição: US pelo número da US, e o DFT que virou US com
        # as duas pontas.
        rotulo = f"US {k}" if tipo == "User Story" and k.isdigit() else comparacao._defeito(k)
        if us is not None and str(us) != k:
            rotulo += f" → US {us}"
        linha = {
            "Defeito": rotulo,
            "Nome": nome,
            "Previsto": p,
            "Realizado": r,
            "Na previsão": k in base_def.index,
            "Entrega prevista (pp)": entrega_prev,
            "Entrega antes": _dia(_get(attr_a, k, "DFT_BugfixMilestone")),
            "Entrega agora": _dia(_get(attr_b, k, "DFT_BugfixMilestone")
                                  or _get(attr_b_tudo, k, "DFT_BugfixMilestone")),
            "Fase antes": _get(attr_a, k, "DFT_Phase"),
            "Fase agora": _get(attr_b, k, "DFT_Phase") or _get(attr_b_tudo, k, "DFT_Phase"),
        }
        linha["Diferença (pp)"] = r - p
        linha["O que aconteceu"] = _explicar_defeito(k, linha, cat_antes, cat_agora, hoje_dt)
        linhas_d.append(linha)
    defeitos = pd.DataFrame(linhas_d)
    if len(defeitos):
        defeitos = defeitos.reindex(defeitos["Diferença (pp)"].abs()
                                    .sort_values(ascending=False).index).reset_index(drop=True)

    # ── Categorias sem defeito, por mensagem ─────────────────────────────────
    mensagens = {}
    for cat in POR_MENSAGEM:
        pb = ped_base[ped_base["Categoria"] == cat]
        pa = cats.get(cat)
        if not len(pb) and (pa is None or not len(pa)):
            continue
        mb = (distribuicao.por_mensagem(pb, t_base, LIMITE_MENSAGEM).set_index("Mensagem")["%"]
              if len(pb) and t_base else pd.Series(dtype=float))
        if cat in NAO_RECORRENTES:
            mb = mb * 0
        ma = (distribuicao.por_mensagem(pa, t_agora, LIMITE_MENSAGEM).set_index("Mensagem")["%"]
              if pa is not None and len(pa) and t_agora else pd.Series(dtype=float))
        tab = pd.DataFrame({"Previsto": mb, "Realizado": ma}).fillna(0.0)
        tab.index = [m if str(m).strip() else "(sem mensagem de erro)" for m in tab.index]
        tab = tab.groupby(level=0).sum()
        tab["Diferença (pp)"] = tab["Realizado"] - tab["Previsto"]
        tab = tab.reindex(tab["Diferença (pp)"].abs().sort_values(ascending=False).index)
        mensagens[cat] = tab.rename_axis("Mensagem").reset_index()

    # As outras categorias com defeito (MOPs, BSIM, Legado…): por defeito, só
    # para dizer quem mexeu nelas. Previsto = o % do mês anterior.
    for cat in set(prev) | set(real):
        if cat in POR_MENSAGEM or cat in NAO_RECORRENTES or cat == trat or cat in mensagens:
            continue
        pb = ped_base[ped_base["Categoria"] == cat]
        pa = cats.get(cat)
        mb = (pb["DefectKey"].astype(str).value_counts() / t_base * 100
              if len(pb) and t_base else pd.Series(dtype=float))
        ma = (pa["DefectKey"].astype(str).value_counts() / t_agora * 100
              if pa is not None and len(pa) and t_agora else pd.Series(dtype=float))
        tab = pd.DataFrame({"Previsto": mb, "Realizado": ma}).fillna(0.0)
        tab["Diferença (pp)"] = tab["Realizado"] - tab["Previsto"]
        tab = tab.reindex(tab["Diferença (pp)"].abs().sort_values(ascending=False).index)
        tab.index = [comparacao._defeito(k) for k in tab.index]
        mensagens[cat] = tab.rename_axis("Mensagem").reset_index()

    dias = int(df_agora.loc[df_agora["Mes"] == mes, "CreatedDate"]
               .dt.tz_convert("America/Sao_Paulo").dt.date.nunique())
    conferencia = {
        "previsto_soma": float(tabela["Previsto"].sum()), "previsto": float(previsto),
        "realizado_soma": float(tabela["Realizado"].sum()), "realizado": realizado,
        "ok": abs(tabela["Previsto"].sum() - previsto) < 1e-6
              and abs(tabela["Realizado"].sum() - realizado) < 1e-6,
    }
    r = PrevistoRealizado(visao=visao, mes=int(mes), foto_id=foto_id, mes_base=mes_base,
                          previsto=float(previsto), realizado=realizado, dias=dias,
                          categorias=tabela, defeitos=defeitos, mensagens=mensagens,
                          texto=[], conferencia=conferencia)
    r.texto = _texto(r)
    tabela["O que explica"] = [_explica_categoria(r, l) for _, l in tabela.iterrows()]
    return r


def _pp(valor):
    return f"{valor:+.2f} pp".replace(".", ",").replace("-", "−")


def _pct(valor):
    return f"{valor:.2f}%".replace(".", ",")


def _explica_categoria(r, linha):
    """O detalhe que mais pesou na diferença da categoria."""
    cat, dif = linha["Categoria"], linha["Diferença (pp)"]
    if abs(dif) < RELEVANTE:
        return ""
    sentido = 1 if dif > 0 else -1
    if cat == fallout_core.CAT_TRATAMENTO and len(r.defeitos):
        d = r.defeitos[r.defeitos["Diferença (pp)"] * sentido >= RELEVANTE].head(2)
        return "; ".join(f"{l['Defeito']} {_pp(l['Diferença (pp)'])}: {l['O que aconteceu']}"
                         for _, l in d.iterrows())
    if cat in r.mensagens:
        m = r.mensagens[cat]
        m = m[m["Diferença (pp)"] * sentido >= RELEVANTE].head(2)
        aspas = '"' if cat in POR_MENSAGEM else ""
        return "; ".join(f"{aspas}{comparacao._curto(l['Mensagem'], 60)}{aspas} "
                         f"{_pp(l['Diferença (pp)'])}" for _, l in m.iterrows())
    if cat in NAO_RECORRENTES:
        return f"pedidos de {_rotulo(r.mes)} já nessa categoria"
    return ""


def _texto(r):
    """O resumo: a diferença total e as categorias que a fazem."""
    nome = fallout_core.nome_exibicao(r.visao)
    dif = r.realizado - r.previsto
    linhas = [f"{nome}: {_rotulo(r.mes)} está em {_pct(r.realizado)}, contra "
              f"{_pct(r.previsto)} previstos na fotografia de "
              f"{pd.Timestamp(r.foto_id):%d/%m} ({_pp(dif)})"
              + (f", com {r.dias} dias de dados." if r.dias < 28 else ".")]
    tab = r.categorias
    sentido = 1 if dif >= 0 else -1
    fortes = tab[tab["Diferença (pp)"] * sentido >= RELEVANTE]
    fortes = fortes.reindex(fortes["Diferença (pp)"].abs().sort_values(ascending=False).index)
    if len(fortes):
        linhas.append(("Acima do previsto: " if sentido > 0 else "Abaixo do previsto: ")
                      + "; ".join(f"{l['Categoria']} {_pp(l['Diferença (pp)'])}"
                                  for _, l in fortes.head(4).iterrows()) + ".")
    contra = tab[tab["Diferença (pp)"] * sentido <= -RELEVANTE]
    if len(contra):
        linhas.append("Compensaram em parte: "
                      + "; ".join(f"{l['Categoria']} {_pp(l['Diferença (pp)'])}"
                                  for _, l in contra.head(3).iterrows()) + ".")
    if len(r.defeitos):
        atrasos = r.defeitos[(r.defeitos["Entrega prevista (pp)"] > 0)
                             & (r.defeitos["Diferença (pp)"] >= RELEVANTE)]
        if len(atrasos):
            linhas.append("Entregas que a previsão contava e não tiraram o que se esperava: "
                          + "; ".join(f"{l['Defeito']} ({l['O que aconteceu']})"
                                      for _, l in atrasos.head(3).iterrows()) + ".")
    return linhas
