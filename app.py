import streamlit as st
import pandas as pd
import glob, os
from datetime import timezone, timedelta
from drive_utils import drive_configurado, baixar_arquivos_drive

st.set_page_config(page_title="Fallout Explorer", layout="wide")

# ── Constantes ────────────────────────────────────────────────────────────────
ABREV_MES = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,
             "jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}
MESES_PT  = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
             7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
FASES_CORRIGIDO = {"Corrigido", "Fechado"}

def mes_do_arquivo(path):
    nome = os.path.splitext(os.path.basename(path))[0].lower()
    for abrev, num in ABREV_MES.items():
        if abrev in nome:
            return num
    return None

# ── Leitura de dados (cache) ──────────────────────────────────────────────────
def _base_dir():
    """Retorna o diretório base dos arquivos (Drive ou local)."""
    if drive_configurado():
        return baixar_arquivos_drive()
    return os.path.dirname(os.path.abspath(__file__))

def listar_jornadas():
    pasta_extracoes = os.path.join(_base_dir(), "extrações")
    jornadas = [
        d for d in os.listdir(pasta_extracoes)
        if os.path.isdir(os.path.join(pasta_extracoes, d, "falhas"))
    ]
    jornadas = sorted(jornadas)
    if "Base Móvel" in jornadas and "Cross Sell" in jornadas:
        jornadas.append("Base + Cross Sell")
    return jornadas

@st.cache_data(show_spinner="Carregando dados...")
def carregar_dados(jornada: str):
    base = _base_dir()
    jornadas_combo = ["Base Móvel", "Cross Sell"] if jornada == "Base + Cross Sell" else [jornada]

    arquivos_falhas = []
    for _j in jornadas_combo:
        arquivos_falhas += sorted(glob.glob(os.path.join(base, "extrações", _j, "falhas/*.csv")))
    df_falhas = pd.concat(
        [pd.read_csv(f) for f in arquivos_falhas],
        ignore_index=True
    )
    df_falhas["CreatedDate"] = pd.to_datetime(df_falhas["CreatedDate"], utc=True, errors="coerce")
    df_falhas = df_falhas.rename(columns={
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.OrderNumber": "OrderNumber",
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Channel__c":  "Channel",
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Segment__c":  "Segment",
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Status":      "OrderStatus",
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.SubStatus__c": "SubStatus",
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.BiometryStatus__c": "BiometryStatus",
        "vlocity_cmt__State__c": "State",
        "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.vlocity_cmt__Reason__c": "Reason",
    })
    df_falhas["DefectNumber_orig"] = df_falhas["DefectNumber__c"].astype(str).str.strip()
    df_falhas["DefectNumber__c"] = (
        pd.to_numeric(df_falhas["DefectNumber__c"], errors="coerce")
        .fillna(-1).astype(int)
    )
    df_falhas["Mes"] = df_falhas["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month

    df_dft = pd.read_excel(os.path.join(base, "RelatorioDFTOctane.xlsx"))
    df_us  = pd.read_excel(os.path.join(base, "RelatorioUSOctane.xlsx"))
    colunas_base    = ["ID", "Name", "Phase", "Bugfix Milestone", "Team", "Type"]
    COL_US_MELHORIA = "US de Melhoria"

    df_dft_prep = df_dft[colunas_base + ([COL_US_MELHORIA] if COL_US_MELHORIA in df_dft.columns else [])].copy()

    if COL_US_MELHORIA in df_dft_prep.columns:
        def _norm_id(v):
            try:    return str(int(float(str(v).strip())))
            except: return str(v).strip()

        df_us_idx = df_us.set_index("ID")[["Name", "Phase", "Bugfix Milestone", "Team", "Type"]].copy()
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
        pd.concat([df_dft_prep[colunas_base], df_us[colunas_base]], ignore_index=True)
        .drop_duplicates(subset="ID")
        .rename(columns={"ID":"DefectNumber__c","Name":"DFT_Name","Phase":"DFT_Phase",
                         "Bugfix Milestone":"DFT_BugfixMilestone","Team":"DFT_Team","Type":"DFT_Type"})
    )
    df_octane["DefectNumber__c"] = (
        pd.to_numeric(df_octane["DefectNumber__c"], errors="coerce")
        .fillna(-1).astype(int)
    )

    df = df_falhas.merge(df_octane, on="DefectNumber__c", how="left")

    # Desduplicar por OrderNumber: mantém a linha com DFT real (> 0) se existir
    df = (
        df.sort_values("DefectNumber__c", ascending=False)
          .drop_duplicates(subset=["OrderNumber"])
          .reset_index(drop=True)
    )

    # Sucessos por mês (mesmo denominador do pipeline)
    arquivos_suc = []
    for _j in jornadas_combo:
        arquivos_suc += sorted(glob.glob(os.path.join(base, "extrações", _j, "sucessos/*.csv")))
    partes_suc = []
    for f in arquivos_suc:
        mes = mes_do_arquivo(f)
        if mes is None:
            continue
        df_tmp = pd.read_csv(f).rename(columns={"expr0": "Sucessos"})
        df_tmp["Mes"] = mes
        partes_suc.append(df_tmp[["Mes", "Sucessos"]])
    sucessos_mes = pd.concat(partes_suc, ignore_index=True).groupby("Mes")["Sucessos"].sum()

    falhas_mes = df.groupby("Mes").size().rename("Falhas")
    resumo = pd.DataFrame({"Falhas": falhas_mes, "Sucessos": sucessos_mes}).fillna(0)
    resumo["Total"] = resumo["Falhas"] + resumo["Sucessos"]

    return df, resumo

# ── Categorização para um mês ─────────────────────────────────────────────────
FASES_MOPS = {"Cancelado", "Rejeitado"}

def categorizar(df, mes):
    hoje        = pd.Timestamp.now(tz="UTC")
    quinze_dias = hoje - pd.Timedelta(days=15)
    df_mes      = df[df["Mes"] == mes].copy()

    milestone_dt  = pd.to_datetime(df_mes["DFT_BugfixMilestone"], utc=True, errors="coerce")
    _corrigido    = df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO)
    _tem_ms       = df_mes["DFT_BugfixMilestone"].notna()
    _encerrado    = _corrigido & _tem_ms & (milestone_dt <= hoje)
    _outros_mask  = df_mes["DefectNumber_orig"].str.strip().str.lower() == "enviado e-mail - outros times"

    cats = {
        "Em Tratamento/Avaliação pela Squad": df_mes[
            df_mes["DefectNumber__c"].notna() &
            (df_mes["DefectNumber__c"] != 999999) &
            (df_mes["DefectNumber__c"] != -1) &
            ~df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS) &
            ~_encerrado
        ],
        "Resolvido": df_mes[_corrigido & _tem_ms & (milestone_dt < quinze_dias)],
        "Falha Pontual": df_mes[df_mes["DefectNumber__c"] == 999999],
        "Falta Associar ao Defeito/US": df_mes[
            (df_mes["DefectNumber__c"].isna() | (df_mes["DefectNumber__c"] == -1)) &
            ~_outros_mask
        ],
        "Em avaliação de eficácia": df_mes[
            _corrigido & _tem_ms & (milestone_dt >= quinze_dias) & (milestone_dt <= hoje)
        ],
        "Em Avaliação por MOPs": df_mes[
            (df_mes["DefectNumber__c"] > 0) &
            (df_mes["DefectNumber__c"] != 999999) &
            df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS)
        ],
        "Em avaliação - Outros times": df_mes[_outros_mask],
    }
    total = len(df_mes)
    return df_mes, cats, total

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

df_mes, cats, _ = categorizar(df, mes_escolhido)

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
            _res = subprocess.run(
                [sys.executable, os.path.join(_app_dir, "pipeline.py"), jornada_escolhida],
                cwd=_data_dir,               # dados (extrações/ + xlsx) ficam aqui
                capture_output=True, text=True,
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
        st.download_button("⬇ Baixar PNG", data=img_bytes,
                           file_name=f"dashboard_{_nome_dash}.png", mime="image/png")
    else:
        st.warning("Dashboard não encontrado. Clique em **Gerar relatório** acima.")

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
    st.subheader(f"Análise de Erros — {jornada_escolhida} | {mes_labels[mes_escolhido]}")

    df_err = df_mes.copy()
    df_err["ErrorHandled__c"] = df_err["ErrorHandled__c"].fillna("(sem mensagem de erro)")
    df_err["tem_dft"] = (
        df_err["DefectNumber__c"].notna() &
        (df_err["DefectNumber__c"] != -1) &
        (df_err["DefectNumber__c"] != 999999)
    )
    df_err["Data"] = df_err["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.date

    # ── Filtro de datas ────────────────────────────────────────────────────────
    datas_disp = sorted(df_err["Data"].dropna().unique())
    if datas_disp:
        col_d1, col_d2, col_f1, col_f2 = st.columns([1.5, 1.5, 3, 1.5])
        with col_d1:
            dt_ini = st.date_input("Data início", value=datas_disp[0],
                                   min_value=datas_disp[0], max_value=datas_disp[-1])
        with col_d2:
            dt_fim = st.date_input("Data fim", value=datas_disp[-1],
                                   min_value=datas_disp[0], max_value=datas_disp[-1])
    else:
        col_f1, col_f2 = st.columns([3, 1.5])
        dt_ini, dt_fim = None, None

    # ── Filtros de texto e DFT ─────────────────────────────────────────────────
    with col_f1:
        busca = st.text_input("Buscar na mensagem de erro", "")
    with col_f2:
        filtro_dft = st.selectbox("Filtrar por DFT", ["Todos", "Com DFT", "Sem DFT"])

    if dt_ini and dt_fim:
        df_err = df_err[(df_err["Data"] >= dt_ini) & (df_err["Data"] <= dt_fim)]
    if busca:
        df_err = df_err[df_err["ErrorHandled__c"].str.contains(busca, case=False, na=False)]
    if filtro_dft == "Com DFT":
        df_err = df_err[df_err["tem_dft"]]
    elif filtro_dft == "Sem DFT":
        df_err = df_err[~df_err["tem_dft"]]

    # ── Pivot por dia ─────────────────────────────────────────────────────────
    pivot_erros = (
        df_err.groupby(["ErrorHandled__c", "Data"], dropna=False)
        .size()
        .unstack("Data", fill_value=0)
        .reset_index()
    )
    pivot_erros.columns.name = None

    day_cols_raw = [c for c in pivot_erros.columns if c != "ErrorHandled__c"]
    rename_map   = {c: c.strftime("%d/%m/%Y") for c in day_cols_raw}
    pivot_erros  = pivot_erros.rename(columns=rename_map)
    day_col_strs = [rename_map[c] for c in day_cols_raw]

    pivot_erros["Total Geral"] = pivot_erros[day_col_strs].sum(axis=1)
    pivot_erros["% do Total"]  = (pivot_erros["Total Geral"] / total_mes * 100).round(2)
    pivot_erros = pivot_erros.sort_values("Total Geral", ascending=False)

    # ── DFTs por mensagem de erro ──────────────────────────────────────────────
    def agg_dfts(sub):
        pares = {}
        for _, row in sub.iterrows():
            dft = row["DefectNumber__c"]
            if pd.isna(dft) or int(dft) in (-1, 999999):
                continue
            dft_id = str(int(dft))
            ms_str = pd.to_datetime(row.get("DFT_BugfixMilestone"), errors="coerce")
            ms_fmt = ms_str.strftime("%d/%m/%Y") if not pd.isna(ms_str) else "s/ data"
            phase  = str(row.get("DFT_Phase", "") or "").strip() or "s/ status"
            pares[dft_id] = (ms_fmt, phase)
        if not pares:
            return ""
        return " | ".join(f"DFT{k} · {ms} · {ph}" for k, (ms, ph) in sorted(pares.items()))

    dfts_por_erro = (
        df_err.groupby("ErrorHandled__c", dropna=False)
        .apply(agg_dfts, include_groups=False)
        .reset_index(name="DFTs")
    )
    pivot_erros = pivot_erros.merge(dfts_por_erro, on="ErrorHandled__c", how="left")
    pivot_erros["DFTs"] = pivot_erros["DFTs"].replace("", "—").fillna("—")
    pivot_erros = pivot_erros.rename(columns={"ErrorHandled__c": "Mensagem de Erro"})

    st.markdown(f"**{len(pivot_erros):,} tipos de erro** | **{int(pivot_erros['Total Geral'].sum()):,} ocorrências**")

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
        file_name=f"erros_{jornada_escolhida}_{mes_labels[mes_escolhido]}.xlsx",
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
