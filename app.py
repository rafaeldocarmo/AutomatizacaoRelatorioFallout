import streamlit as st
import pandas as pd
import os
from drive_utils import drive_configurado, baixar_arquivos_drive
import fallout_core
from fallout_core import MESES_PT

st.set_page_config(page_title="Fallout Explorer", layout="wide")

# ── Leitura de dados (cache) ──────────────────────────────────────────────────
def _base_dir():
    """Retorna o diretório base dos arquivos (Drive ou local)."""
    if drive_configurado():
        return baixar_arquivos_drive()
    return os.path.dirname(os.path.abspath(__file__))

def listar_jornadas():
    jornadas = fallout_core.jornadas_disponiveis(_base_dir())
    opcoes = list(jornadas)
    if "Base Móvel" in jornadas and "Cross Sell" in jornadas:
        opcoes.append("Base + Cross Sell")
    if len(jornadas) > 1:
        opcoes.append("Consolidado")   # todas as jornadas juntas
    return opcoes

@st.cache_data(show_spinner="Carregando dados...")
def carregar_dados(jornada: str):
    base = _base_dir()
    df, _df_octane, resumo, _jornadas_combo = fallout_core.carregar_base(base, jornada)
    return df, resumo

# ── Colunas a exibir na tabela de pedidos ─────────────────────────────────────
COLS_EXIBIR = [
    "OrderNumber", "CreatedDate", "DefectNumber_orig", "ErrorHandled__c",
    "DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Team", "DFT_Type",
    "State", "Channel", "Segment",
]

def to_excel_bytes(df_export: pd.DataFrame) -> bytes:
    import io
    df_exp = df_export.copy()
    for col in df_exp.select_dtypes(include=["datetimetz"]).columns:
        df_exp[col] = df_exp[col].dt.tz_localize(None)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_exp.to_excel(writer, index=False)
    return buf.getvalue()

def disparar_download(data: bytes, file_name: str, mime: str):
    """Dispara o download do navegador automaticamente (sem precisar de um 2º clique)."""
    import base64
    b64 = base64.b64encode(data).decode()
    st.components.v1.html(f"""
        <a id="auto-dl" href="data:{mime};base64,{b64}" download="{file_name}"></a>
        <script>document.getElementById('auto-dl').click();</script>
    """, height=0)

# ── Interface ──────────────────────────────────────────────────────────────────
st.markdown("""
<div style='background:#C0392B;padding:18px 24px;border-radius:6px;margin-bottom:16px'>
  <span style='color:white;font-size:22px;font-weight:700'>
    Explorador de Fallout
  </span>
</div>
""", unsafe_allow_html=True)

jornadas = listar_jornadas()
col_jornada, col_sel, col_info, col_info2 = st.columns([2, 2, 2, 2])
with col_jornada:
    jornada_escolhida = st.selectbox("Jornada", options=jornadas,
                                      index=jornadas.index("Base Móvel") if "Base Móvel" in jornadas else 0)

with st.sidebar:
    st.markdown("### Dados")
    if st.button("🔄 Atualizar dados do Drive"):
        st.cache_data.clear()
        st.success("Cache limpo. Recarregando dados mais recentes...")
        st.rerun()
    st.caption("Os dados são cacheados por até 30 min. Use o botão acima após atualizar os arquivos no Drive.")

df, resumo = carregar_dados(jornada_escolhida)
meses_disp = sorted(df["Mes"].unique().tolist())
mes_labels = {m: f"{MESES_PT[m]}-26" for m in meses_disp}

with col_sel:
    mes_escolhido = st.selectbox(
        "Mês de análise",
        options=meses_disp,
        format_func=lambda m: mes_labels[m],
        index=len(meses_disp) - 1,
    )

total_mes = int(resumo.loc[mes_escolhido, "Total"]) if mes_escolhido in resumo.index else 1
fallout_pct = resumo.loc[mes_escolhido, "Falhas"] / total_mes * 100 if mes_escolhido in resumo.index else 0

df_mes, cats, _ = fallout_core.categorizar(df, mes_escolhido)

with col_info:
    st.metric("Total de pedidos no mês", f"{total_mes:,}")
with col_info2:
    st.metric("Fallout Rate", f"{fallout_pct:.2f}%")

st.markdown("---")

# ── Aba: Consolidado / Distribuição ──────────────────────────────────────────
aba_consolidado, aba_distribuicao, aba_erros, aba_defeitos = st.tabs(["📊 Consolidado", "📋 Distribuição Fallout", "🔍 Análise de Erros", "🔗 Erros por Defeito"])

with aba_consolidado:
    _nome_dash = jornada_escolhida.replace(" + ", "_").replace(" ", "_")
    _app_dir   = os.path.dirname(os.path.abspath(__file__))
    _data_dir  = _base_dir()   # local: mesmo que _app_dir | Drive: pasta temporária

    # ── Botão: gerar relatório ────────────────────────────────────────────────
    if st.button(f"🔄 Gerar relatório — {jornada_escolhida}"):
        import subprocess, sys
        with st.spinner(f"Gerando dashboard de {jornada_escolhida}..."):
            _env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
            _res = subprocess.run(
                [sys.executable, os.path.join(_app_dir, "pipeline.py"), jornada_escolhida],
                cwd=_data_dir,               # dados (extrações/ + xlsx) ficam aqui
                capture_output=True, text=True,
                encoding="utf-8", errors="replace", env=_env,
            )
        if _res.returncode == 0:
            st.success("Relatório gerado com sucesso!")
        else:
            st.error("Erro ao gerar o relatório:")
            st.code(_res.stderr[-3000:] if _res.stderr else _res.stdout[-3000:])

    # Procura o PNG: primeiro na pasta de dados (onde o pipeline gera), depois na do app
    dashboard_path = os.path.join(_data_dir, f"dashboard_{_nome_dash}.png")
    if not os.path.exists(dashboard_path):
        dashboard_path = os.path.join(_app_dir, f"dashboard_{_nome_dash}.png")
    if os.path.exists(dashboard_path):
        import base64
        img_b64 = base64.b64encode(open(dashboard_path, "rb").read()).decode()
        img_bytes = open(dashboard_path, "rb").read()
        st.components.v1.html(f"""
        <style>
          * {{ box-sizing: border-box; margin: 0; padding: 0; }}
          body {{ font-family: sans-serif; background: #111; }}
          #fs-controls {{ display: flex; gap: 8px; align-items: center; padding: 8px 12px;
                          background: #222; position: sticky; top: 0; z-index: 10; }}
          #fs-controls button {{ padding: 4px 14px; border: 1px solid #C0392B; background: #333;
                                 color: #fff; border-radius: 4px; cursor: pointer; font-size: 15px; font-weight: bold; }}
          #fs-controls button:hover {{ background: #C0392B; }}
          #fs-controls span {{ color: #aaa; font-size: 13px; }}
          #fs-container {{ overflow: auto; height: calc(100vh - 46px); cursor: grab; }}
          #fs-container:active {{ cursor: grabbing; }}
          #fs-img {{ transform-origin: top left; display: block; }}
        </style>
        <div id="fs-controls">
          <button onclick="fsZoom(-0.2)">−</button>
          <button onclick="fsZoom(0.2)">+</button>
          <button onclick="fsFit()">Ajustar</button>
          <span id="fs-zoom-level">100%</span>
        </div>
        <div id="fs-container">
          <img id="fs-img" src="data:image/png;base64,{img_b64}" />
        </div>
        <script>
          var fsScale = 1;
          var fsImg = document.getElementById('fs-img');
          var fsContainer = document.getElementById('fs-container');
          function fsZoom(delta) {{
            fsScale = Math.min(Math.max(fsScale + delta, 0.2), 5);
            fsImg.style.transform = 'scale(' + fsScale + ')';
            document.getElementById('fs-zoom-level').textContent = Math.round(fsScale * 100) + '%';
          }}
          function fsFit() {{
            var ratio = (window.innerWidth - 20) / fsImg.naturalWidth;
            fsScale = Math.round(ratio * 100) / 100;
            fsImg.style.transform = 'scale(' + fsScale + ')';
            document.getElementById('fs-zoom-level').textContent = Math.round(fsScale * 100) + '%';
            fsContainer.scrollTop = 0; fsContainer.scrollLeft = 0;
          }}
          fsImg.onload = fsFit;
          if (fsImg.complete) fsFit();
        </script>
        """, height=900)
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button("⬇ Baixar PNG", data=img_bytes,
                               file_name=f"dashboard_{_nome_dash}.png", mime="image/png")
        with col_dl2:
            if st.button("📊 Gerar PowerPoint editável"):
                from gerar_pptx import gerar_pptx
                with st.spinner("Gerando PowerPoint..."):
                    buf = gerar_pptx(jornada_escolhida, _data_dir)
                st.session_state["pptx_bytes"] = buf.getvalue()
                st.session_state["pptx_name"]  = f"relatorio_{_nome_dash}.pptx"
            if st.session_state.get("pptx_bytes") and st.session_state.get("pptx_name") == f"relatorio_{_nome_dash}.pptx":
                st.download_button(
                    "⬇ Baixar PowerPoint", data=st.session_state["pptx_bytes"],
                    file_name=st.session_state["pptx_name"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    else:
        st.warning("Dashboard não encontrado. Clique em **Gerar relatório** acima.")

    # ── Slide executivo de 3 colunas (independe da jornada selecionada) ───────
    st.markdown("---")
    st.markdown("**Slide executivo — Consolidado | Prospect | Base + Cross Sell**")
    if st.button("🗂 Gerar slide de 3 colunas"):
        from gerar_slide3col import extrair_todas, gerar_slide3col, gerar_slide3col_png
        with st.spinner("Gerando slide de 3 colunas..."):
            _d3 = extrair_todas(_data_dir)          # roda o pipeline uma só vez
            st.session_state["s3col_bytes"] = gerar_slide3col(dados=_d3).getvalue()
            st.session_state["s3col_png"]   = gerar_slide3col_png(dados=_d3).getvalue()

    if st.session_state.get("s3col_png"):
        st.image(st.session_state["s3col_png"], use_container_width=True)
        _c1, _c2 = st.columns(2)
        with _c1:
            st.download_button(
                "⬇ Baixar slide (PNG)", data=st.session_state["s3col_png"],
                file_name="slide_3colunas.png", mime="image/png")
        with _c2:
            st.download_button(
                "⬇ Baixar slide (PowerPoint)", data=st.session_state["s3col_bytes"],
                file_name="slide_3colunas.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    # ── PowerPoint completo: 3 colunas + Top Ofensores + um slide por jornada ──
    st.markdown("---")
    st.markdown("**PowerPoint completo — 3 colunas + Top Ofensores + um slide por jornada**")
    if st.button("📑 Gerar PowerPoint completo"):
        from gerar_slide3col import gerar_pptx_completo
        with st.spinner("Gerando PowerPoint completo (isso lê todos os dados de novo, pode levar um tempo)..."):
            pptx_bytes = gerar_pptx_completo(_data_dir).getvalue()
        _PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        disparar_download(pptx_bytes, "relatorio_completo.pptx", _PPTX_MIME)
        st.success("PowerPoint completo gerado — o download deve começar automaticamente.")

with aba_distribuicao:
    st.subheader(f"Distribuição Fallout — {mes_labels[mes_escolhido]}")

    # ── Cards clicáveis ───────────────────────────────────────────────────────
    cols = st.columns(3)
    cat_names = list(cats.keys())

    if "categoria_ativa" not in st.session_state:
        st.session_state.categoria_ativa = None

    for i, nome in enumerate(cat_names):
        qtd = len(cats[nome])
        pct = qtd / total_mes * 100 if total_mes > 0 else 0
        ativo = st.session_state.categoria_ativa == nome
        border = "3px solid #C0392B" if ativo else "1px solid #ddd"
        bg     = "#fff5f5" if ativo else "#fafafa"
        with cols[i % 3]:
            st.markdown(f"""
            <div style='border:{border};background:{bg};border-radius:8px;
                        padding:14px 16px;margin-bottom:10px;cursor:pointer'>
              <div style='font-size:13px;color:#555;margin-bottom:4px'>{nome}</div>
              <div style='font-size:26px;font-weight:700;color:#C0392B'>{pct:.2f}%</div>
              <div style='font-size:12px;color:#888'>{qtd:,} pedidos</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"Ver pedidos", key=f"btn_{i}"):
                if st.session_state.categoria_ativa == nome:
                    st.session_state.categoria_ativa = None
                else:
                    st.session_state.categoria_ativa = nome

    # ── Tabela de pedidos da categoria selecionada ────────────────────────────
    cat_ativa = st.session_state.categoria_ativa
    if cat_ativa:
        st.markdown("---")
        df_cat = cats[cat_ativa].copy()

        if cat_ativa == "Em Tratamento/Avaliação pela Squad":
            subfiltro = st.radio(
                "Sub-grupo",
                ["Todos", "Planejado (com milestone)", "US s/ data", "DFT s/ data"],
                horizontal=True,
            )
            if subfiltro == "Planejado (com milestone)":
                df_cat = df_cat[df_cat["DFT_BugfixMilestone"].notna()]
            elif subfiltro == "US s/ data":
                df_cat = df_cat[df_cat["DFT_BugfixMilestone"].isna() &
                                (df_cat["DFT_Type"].fillna("").str.strip() == "User Story")]
            elif subfiltro == "DFT s/ data":
                df_cat = df_cat[df_cat["DFT_BugfixMilestone"].isna() &
                                (df_cat["DFT_Type"].fillna("").str.strip() != "User Story")]

        cols_disp = [c for c in COLS_EXIBIR if c in df_cat.columns]
        df_show = df_cat[cols_disp].copy().reset_index(drop=True)
        if "DefectNumber_orig" in df_show.columns:
            df_show = df_show.rename(columns={"DefectNumber_orig": "DefectNumber__c"})
            df_show["DefectNumber__c"] = df_show["DefectNumber__c"].replace({"nan": "", "-1": "", "999999": "999999 (Pontual)"})

        pct_cat = len(df_show) / total_mes * 100 if total_mes > 0 else 0
        st.markdown(f"#### {cat_ativa} — {len(df_show):,} pedidos ({pct_cat:.2f}% do total)")

        filtro = st.text_input("Filtrar por número de pedido, DFT ou qualquer campo", "")
        if filtro:
            mask = df_show.apply(lambda row: row.astype(str).str.contains(filtro, case=False).any(), axis=1)
            df_show = df_show[mask]

        st.dataframe(df_show, use_container_width=True, height=400)

        excel_bytes = to_excel_bytes(df_show)
        nome_arquivo = f"fallout_{cat_ativa[:20].replace('/', '-').strip()}_{mes_labels[mes_escolhido]}.xlsx"
        st.download_button(
            label="⬇ Baixar Excel",
            data=excel_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with aba_erros:
    st.subheader(f"Análise de Erros — {jornada_escolhida}")

    # Usa todo o histórico (não só o mês selecionado); o período é escolhido abaixo
    df_err = df.copy()
    df_err["ErrorHandled__c"] = df_err["ErrorHandled__c"].fillna("(sem mensagem de erro)")
    _orig_lower = df_err["DefectNumber_orig"].astype(str).str.strip().str.lower()
    df_err["_outros_times"] = _orig_lower == "enviado e-mail - outros times"
    df_err["_pontual"]      = df_err["DefectNumber__c"] == 999999
    df_err["tem_dft"] = (
        df_err["DefectNumber__c"].notna() &
        (df_err["DefectNumber__c"] != -1) &
        (df_err["DefectNumber__c"] != 999999) &
        ~df_err["_outros_times"]
    )
    df_err["Data"] = df_err["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.date

    # ── Filtro de datas (todo o histórico disponível) ──────────────────────────
    datas_disp = sorted(df_err["Data"].dropna().unique())
    if datas_disp:
        # padrão = mês selecionado no topo, mas dá para ampliar até o histórico todo
        _dm = [d for d in datas_disp if d.month == mes_escolhido]
        _ini_pad = _dm[0]  if _dm else datas_disp[0]
        _fim_pad = _dm[-1] if _dm else datas_disp[-1]
        col_d1, col_d2, col_g, col_f1, col_f2 = st.columns([1.4, 1.4, 1.2, 2.6, 1.4])
        with col_d1:
            dt_ini = st.date_input("Data início", value=_ini_pad,
                                   min_value=datas_disp[0], max_value=datas_disp[-1])
        with col_d2:
            dt_fim = st.date_input("Data fim", value=_fim_pad,
                                   min_value=datas_disp[0], max_value=datas_disp[-1])
        with col_g:
            agrupar = st.selectbox("Colunas por", ["Dia", "Mês"])
    else:
        col_f1, col_f2 = st.columns([3, 1.5])
        dt_ini, dt_fim, agrupar = None, None, "Dia"

    # ── Filtros de texto e DFT ─────────────────────────────────────────────────
    with col_f1:
        busca = st.text_input("Buscar na mensagem de erro", "")
    with col_f2:
        filtro_dft = st.selectbox(
            "Filtrar por DFT",
            ["Todos", "Com DFT", "Sem DFT", "Falha Pontual", "Outros Times"],
        )

    if dt_ini and dt_fim:
        df_err = df_err[(df_err["Data"] >= dt_ini) & (df_err["Data"] <= dt_fim)]

    # Denominador do "% do Total": soma dos meses cobertos pelo período escolhido
    _meses_sel = sorted({d.month for d in df_err["Data"].dropna().unique()})
    _total_per = int(sum(resumo.loc[m, "Total"] for m in _meses_sel if m in resumo.index))
    _total_per = _total_per or 1

    if busca:
        df_err = df_err[df_err["ErrorHandled__c"].str.contains(busca, case=False, na=False)]
    if filtro_dft == "Com DFT":
        df_err = df_err[df_err["tem_dft"]]
    elif filtro_dft == "Sem DFT":
        df_err = df_err[~df_err["tem_dft"] & ~df_err["_pontual"] & ~df_err["_outros_times"]]
    elif filtro_dft == "Falha Pontual":
        df_err = df_err[df_err["_pontual"]]
    elif filtro_dft == "Outros Times":
        df_err = df_err[df_err["_outros_times"]]

    # ── Pivot por dia ou por mês ──────────────────────────────────────────────
    if agrupar == "Mês":
        _per = (df_err["CreatedDate"].dt.tz_convert("America/Sao_Paulo")
                .dt.tz_localize(None).dt.to_period("M"))   # naive: evita warning do pandas
        df_err = df_err.assign(_grp=_per)
        _fmt_col = lambda p: f"{MESES_PT[p.month]}-{str(p.year)[-2:]}"
    else:
        df_err = df_err.assign(_grp=df_err["Data"])
        _fmt_col = lambda d: d.strftime("%d/%m/%Y")

    pivot_erros = (
        df_err.groupby(["ErrorHandled__c", "_grp"], dropna=False)
        .size()
        .unstack("_grp", fill_value=0)
        .reset_index()
    )
    pivot_erros.columns.name = None

    day_cols_raw = sorted(c for c in pivot_erros.columns if c != "ErrorHandled__c")
    rename_map   = {c: _fmt_col(c) for c in day_cols_raw}
    pivot_erros  = pivot_erros.rename(columns=rename_map)
    day_col_strs = [rename_map[c] for c in day_cols_raw]

    pivot_erros["Total Geral"] = pivot_erros[day_col_strs].sum(axis=1)
    pivot_erros["% do Total"]  = (pivot_erros["Total Geral"] / _total_per * 100).round(2)
    pivot_erros = pivot_erros.sort_values("Total Geral", ascending=False)

    # ── DFTs por mensagem de erro ──────────────────────────────────────────────
    def agg_dfts(sub):
        pares = {}
        tem_pontual = False
        tem_outros  = False
        for _, row in sub.iterrows():
            orig = str(row.get("DefectNumber_orig", "") or "").strip().lower()
            if orig == "enviado e-mail - outros times":
                tem_outros = True
                continue
            dft = row["DefectNumber__c"]
            if pd.isna(dft) or int(dft) == -1:
                continue
            if int(dft) == 999999:
                tem_pontual = True
                continue
            dft_id = str(int(dft))
            ms_str = pd.to_datetime(row.get("DFT_BugfixMilestone"), errors="coerce")
            ms_fmt = ms_str.strftime("%d/%m/%Y") if not pd.isna(ms_str) else "s/ data"
            phase  = str(row.get("DFT_Phase", "") or "").strip() or "s/ status"
            pares[dft_id] = (ms_fmt, phase)
        partes = [f"DFT{k} · {ms} · {ph}" for k, (ms, ph) in sorted(pares.items())]
        if tem_pontual:
            partes.append("Falha Pontual")
        if tem_outros:
            partes.append("Em avaliação - Outros Times")
        return " | ".join(partes)

    dfts_por_erro = (
        df_err.groupby("ErrorHandled__c", dropna=False)
        .apply(agg_dfts, include_groups=False)
        .reset_index(name="DFTs")
    )
    pivot_erros = pivot_erros.merge(dfts_por_erro, on="ErrorHandled__c", how="left")
    pivot_erros["DFTs"] = pivot_erros["DFTs"].replace("", "—").fillna("—")
    pivot_erros = pivot_erros.rename(columns={"ErrorHandled__c": "Mensagem de Erro"})

    _per_txt = f"{dt_ini.strftime('%d/%m/%Y')} a {dt_fim.strftime('%d/%m/%Y')}" if dt_ini else "—"
    _meses_txt = ", ".join(mes_labels.get(m, str(m)) for m in _meses_sel)
    st.markdown(
        f"**{len(pivot_erros):,} tipos de erro** | "
        f"**{int(pivot_erros['Total Geral'].sum()):,} ocorrências** | "
        f"Período: **{_per_txt}**"
    )
    st.caption(f"% do Total calculado sobre {_total_per:,} pedidos ({_meses_txt}). "
               "Meses parciais usam o total do mês inteiro.".replace(",", "."))
    if agrupar == "Dia" and len(day_col_strs) > 62:
        st.warning(f"O período tem {len(day_col_strs)} dias — troque **Colunas por** para "
                   "**Mês** para uma leitura mais confortável.")

    col_order  = ["Mensagem de Erro"] + day_col_strs + ["Total Geral", "% do Total", "DFTs"]
    col_config = {
        "Mensagem de Erro": st.column_config.TextColumn(width="large"),
        "Total Geral":      st.column_config.NumberColumn(width="small"),
        "% do Total":       st.column_config.NumberColumn(format="%.2f%%", width="small"),
        "DFTs":             st.column_config.TextColumn(width="large"),
        **{d: st.column_config.NumberColumn(width="small") for d in day_col_strs},
    }
    import numpy as _np
    import matplotlib.cm as _cm
    import matplotlib.colors as _mc
    # Trunca o colormap: evita o verde escuro (começa em 30% do RdYlGn_r)
    _cmap = _mc.LinearSegmentedColormap.from_list(
        "RdYlGn_light", _cm.RdYlGn_r(_np.linspace(0.30, 1.0, 256))
    )
    _day_vals = pivot_erros[day_col_strs].values
    _nonzero  = _day_vals[_day_vals > 0]
    _vmin     = int(_nonzero.min()) if len(_nonzero) else 1
    _vmax     = int(_day_vals.max()) if _day_vals.max() > 0 else 1

    def _color_cell(val):
        if val == 0 or pd.isna(val):
            return ""
        norm = (val - _vmin) / (_vmax - _vmin) if _vmax > _vmin else 1.0
        return f"background-color: {_mc.to_hex(_cmap(norm))}"

    styled = pivot_erros[col_order].style.map(_color_cell, subset=day_col_strs)
    st.dataframe(
        styled,
        use_container_width=True,
        height=500,
        column_config=col_config,
    )

    excel_erros = to_excel_bytes(pivot_erros[col_order])
    st.download_button(
        label="⬇ Baixar Excel",
        data=excel_erros,
        file_name=(f"erros_{jornada_escolhida}_"
                   f"{dt_ini.strftime('%Y%m%d')}-{dt_fim.strftime('%Y%m%d')}.xlsx"
                   if dt_ini else f"erros_{jornada_escolhida}.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with aba_defeitos:
    st.subheader(f"Erros por Defeito — {jornada_escolhida} | {mes_labels[mes_escolhido]}")

    df_def = df_mes.copy()
    df_def["ErrorHandled__c"] = df_def["ErrorHandled__c"].fillna("(sem mensagem de erro)")
    df_def["Data"] = df_def["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.date

    # ── Filtros ───────────────────────────────────────────────────────────────
    datas_def = sorted(df_def["Data"].dropna().unique())
    col_d1, col_d2, col_f1, col_f2 = st.columns([1.4, 1.4, 2.5, 1.7])
    with col_d1:
        dt_ini_d = st.date_input("Data início", value=datas_def[0] if datas_def else None,
                                  min_value=datas_def[0] if datas_def else None,
                                  max_value=datas_def[-1] if datas_def else None,
                                  key="def_dt_ini")
    with col_d2:
        dt_fim_d = st.date_input("Data fim", value=datas_def[-1] if datas_def else None,
                                  min_value=datas_def[0] if datas_def else None,
                                  max_value=datas_def[-1] if datas_def else None,
                                  key="def_dt_fim")
    with col_f1:
        busca_def = st.text_input("Buscar por erro ou defeito", "", key="def_busca")
    with col_f2:
        min_occ = st.number_input("Mín. ocorrências por erro", min_value=1, value=1, step=1)

    if datas_def:
        df_def = df_def[(df_def["Data"] >= dt_ini_d) & (df_def["Data"] <= dt_fim_d)]
    if busca_def:
        mask = (
            df_def["ErrorHandled__c"].str.contains(busca_def, case=False, na=False) |
            df_def["DefectNumber_orig"].astype(str).str.contains(busca_def, case=False, na=False)
        )
        df_def = df_def[mask]

    # ── Agrupamento Erro → DFT ────────────────────────────────────────────────
    grp_def = (
        df_def.groupby(
            ["ErrorHandled__c", "DefectNumber_orig",
             "DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Team"],
            dropna=False
        )
        .size()
        .reset_index(name="Qtd")
    )
    total_por_erro = grp_def.groupby("ErrorHandled__c")["Qtd"].sum()
    grp_def = grp_def[grp_def["ErrorHandled__c"].map(total_por_erro) >= min_occ]
    grp_def = grp_def.sort_values(
        ["ErrorHandled__c"],
        key=lambda s: s.map(total_por_erro),
        ascending=False
    )

    # ── Renderização em cards ─────────────────────────────────────────────────
    erros_unicos = grp_def["ErrorHandled__c"].unique()
    st.write(f"**{len(erros_unicos):,} tipos de erro**")

    cards_html = []
    for erro in erros_unicos:
        dfts       = grp_def[grp_def["ErrorHandled__c"] == erro].sort_values("Qtd", ascending=False)
        total_erro = int(dfts["Qtd"].sum())
        pct_erro   = total_erro / total_mes * 100 if total_mes > 0 else 0
        n_dfts     = len(dfts)

        linhas_dft = []
        for _, r in dfts.iterrows():
            orig  = str(r["DefectNumber_orig"] or "").strip()
            qtd   = int(r["Qtd"])
            nome  = str(r["DFT_Name"]  or "").strip()
            phase = str(r["DFT_Phase"] or "").strip()
            team  = str(r["DFT_Team"]  or "").strip()
            ms    = pd.to_datetime(r["DFT_BugfixMilestone"], errors="coerce")
            ms_s  = ms.strftime("%d/%m/%Y") if not pd.isna(ms) else ""

            if orig in ("", "nan", "-1"):
                dft_label = "(sem DFT)"
                info      = "sem defeito associado"
                info_color = "#aaa"
            elif orig == "999999":
                dft_label  = "Pontual"
                info       = ""
                info_color = "#666"
            else:
                dft_label  = "DFT " + orig
                partes = [p for p in [nome, phase, ("entrega " + ms_s) if ms_s else "", team] if p]
                info       = " · ".join(partes)
                info_color = "#555"

            bc = "#C0392B" if qtd >= 10 else "#e67e22" if qtd >= 3 else "#7f8c8d"
            linhas_dft.append(
                "<div style='display:flex;align-items:center;gap:10px;"
                "padding:5px 0;border-bottom:1px solid #f0f0f0'>"
                "<span style='background:" + bc + ";color:#fff;border-radius:4px;"
                "padding:1px 8px;font-size:12px;font-weight:700;min-width:28px;text-align:center'>"
                + str(qtd) +
                "</span>"
                "<span style='font-weight:600;color:#1a1a1a;font-size:13px;min-width:110px'>"
                + dft_label +
                "</span>"
                "<span style='color:" + info_color + ";font-size:12px'>" + info + "</span>"
                "</div>"
            )

        multi = (
            "<span style='background:#1565C0;color:#fff;border-radius:10px;"
            "padding:1px 8px;font-size:11px;margin-left:8px'>" + str(n_dfts) + " DFTs</span>"
        ) if n_dfts > 1 else ""

        cards_html.append(
            "<div style='border:1px solid #ddd;border-radius:8px;"
            "margin-bottom:14px;overflow:hidden;font-family:sans-serif'>"
            "<div style='background:#2c2c2c;padding:10px 14px;"
            "display:flex;justify-content:space-between;align-items:center'>"
            "<span style='color:#fff;font-size:12px;flex:1;margin-right:16px'>" + erro + "</span>"
            "<span style='white-space:nowrap'>"
            "<span style='background:#C0392B;color:#fff;border-radius:4px;"
            "padding:2px 10px;font-size:13px;font-weight:700'>" + str(total_erro) + "</span>"
            "<span style='color:#ccc;font-size:11px;margin-left:6px'>" + f"{pct_erro:.2f}%" + "</span>"
            + multi +
            "</span></div>"
            "<div style='padding:8px 14px 4px;background:#fff'>"
            + "".join(linhas_dft) +
            "</div></div>"
        )

    altura_total = max(400, len(erros_unicos) * 80)
    st.components.v1.html(
        "<div style='font-family:sans-serif;padding:4px'>" + "".join(cards_html) + "</div>",
        height=min(altura_total, 800),
        scrolling=True,
    )
