import streamlit as st
import pandas as pd
import os
import fallout_core
from fallout_core import MESES_PT

st.set_page_config(page_title="Fallout Explorer", layout="wide")

# ── Design system "Modernist" (apenas visual) ─────────────────────────────────
# Tokens espelhados de _ds/modernist/styles.css: Archivo, raio 0, réguas de 2px,
# hierarquia por alinhamento e divisores em vez de sombra/cor decorativa.
COR_BG        = "#f3f2f2"
COR_SURFACE   = "#eae9e9"
COR_TEXT      = "#201e1d"
COR_ACCENT    = "#ec3013"
COR_ACC_100   = "#fff2ef"
COR_ACC_500   = "#ff563c"
COR_ACC_600   = "#dd2b0f"
COR_ACC_700   = "#ae1800"
COR_ACC_800   = "#7c1405"
COR_DIVIDER   = "rgba(32,30,29,.40)"
COR_MUTED     = "rgba(32,30,29,.55)"
COR_MUTED_2   = "rgba(32,30,29,.68)"
FONTE         = '"Archivo", system-ui, sans-serif'

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;600;800&display=swap');

/* — base — */
[data-testid="stAppViewContainer"] {{ background: {COR_BG}; }}
[data-testid="stHeader"] {{ background: transparent; }}
[data-testid="stMain"] .block-container {{ padding-top: 2.2rem; max-width: 1500px; }}
html, body, [class*="css"], [data-testid="stAppViewContainer"] * {{
  font-family: {FONTE};
  color: {COR_TEXT};
}}
/* Streamlit estiliza os headings com seletor mais específico — daí o !important */
h1, h2, h3, h4, h5, h6,
[data-testid="stHeading"] h1, [data-testid="stHeading"] h2, [data-testid="stHeading"] h3,
[data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3, [data-testid="stMarkdownContainer"] h4 {{
  font-family: {FONTE} !important; font-weight: 800 !important;
  letter-spacing: -.015em; line-height: 1.12;
}}

/* — réguas de 2px em vez de linha fina (Streamlit fixa height:1px) — */
hr, [data-testid="stMarkdownContainer"] hr {{
  border: 0 !important; height: 2px !important; background: {COR_DIVIDER} !important;
  margin: 1.1rem 0;
}}

/* — botões: retângulos sem raio; primário na cor de acento — */
.stButton > button, .stDownloadButton > button {{
  border-radius: 0; border: 1px solid {COR_DIVIDER}; background: transparent;
  color: {COR_TEXT}; font-family: {FONTE}; font-weight: 600; font-size: 13px;
  padding: .55rem 1.1rem; box-shadow: none; transition: background .12s ease;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{
  background: rgba(32,30,29,.07); border-color: {COR_DIVIDER}; color: {COR_TEXT};
}}
.stButton > button:active, .stDownloadButton > button:active {{ background: rgba(32,30,29,.14); }}
.stButton > button:focus, .stDownloadButton > button:focus {{ box-shadow: none; color: {COR_TEXT}; }}
.stDownloadButton > button {{ background: {COR_ACCENT}; border-color: {COR_ACCENT}; color: #fff; }}
.stDownloadButton > button:hover {{ background: {COR_ACC_600}; border-color: {COR_ACC_600}; color: #fff; }}
.stDownloadButton > button p {{ color: #fff !important; }}

/* — abas: sublinhado de 3px, sem pílula — */
.stTabs [data-baseweb="tab-list"] {{
  gap: 8px; border-bottom: 2px solid {COR_DIVIDER}; background: transparent;
}}
.stTabs [data-baseweb="tab"] {{
  border-radius: 0; background: transparent; padding: 14px 18px 12px;
  font-family: {FONTE}; font-weight: 600; font-size: 14px; color: {COR_MUTED};
}}
.stTabs [aria-selected="true"] {{ color: {COR_TEXT}; }}
.stTabs [data-baseweb="tab-highlight"] {{ background: {COR_ACC_700}; height: 3px; }}
.stTabs [data-baseweb="tab-border"] {{ display: none; }}

/* — campos: raio 0, superfície plana — */
[data-baseweb="select"] > div, .stTextInput input, .stDateInput input,
.stNumberInput input, [data-baseweb="input"] {{
  border-radius: 0 !important; background: {COR_SURFACE} !important;
  border: 1px solid {COR_DIVIDER} !important; font-size: 14px;
}}
[data-baseweb="select"] > div:hover, .stTextInput input:hover {{ border-color: rgba(32,30,29,.45) !important; }}
[data-baseweb="popover"] li, [data-baseweb="menu"] {{ border-radius: 0 !important; font-family: {FONTE}; }}
.stTextInput label, .stDateInput label, .stSelectbox label,
.stNumberInput label, .stRadio label {{
  font-size: 12px !important; font-weight: 600 !important; color: {COR_MUTED_2} !important;
  text-transform: uppercase; letter-spacing: .06em;
}}
.stRadio [role="radiogroup"] {{ gap: 14px; }}

/* — expanders: caixa reta com régua — */
[data-testid="stExpander"] {{
  border: 1px solid {COR_DIVIDER}; border-radius: 0; background: #fff; margin-bottom: 2px;
}}
[data-testid="stExpander"] summary {{ border-radius: 0; font-family: {FONTE}; padding: .6rem .9rem; }}
[data-testid="stExpander"] summary:hover {{ background: rgba(32,30,29,.04); }}

/* — dataframe — */
[data-testid="stDataFrame"] {{ border: 1px solid {COR_DIVIDER}; border-radius: 0; }}

/* — alertas — */
[data-testid="stAlert"] {{ border-radius: 0; }}

/* — sparkline e tiles de KPI — */
.kpi-row {{ display: flex; gap: 2px; }}
.kpi-tile {{ flex: 1; padding: 16px 18px; }}
.kpi-label {{
  font-weight: 600; font-size: 10px; letter-spacing: .08em; text-transform: uppercase;
}}
.kpi-value {{ font-weight: 800; font-size: 28px; line-height: 1.1; }}
.kpi-delta {{ font-weight: 600; font-size: 11px; }}
.kpi-spark {{ display: flex; gap: 2px; align-items: flex-end; height: 18px; margin-top: 10px; }}
.kpi-foot {{ font-size: 10px; margin-top: 8px; }}
.chip {{
  display: inline-flex; align-items: center; padding: 4px 10px; margin: 0 6px 6px 0;
  background: {COR_SURFACE}; font: 600 11px ui-monospace, SFMono-Regular, Menlo, monospace;
}}
.raw-msg {{
  font: 12px/1.5 ui-monospace, SFMono-Regular, Menlo, monospace; color: {COR_MUTED_2};
  padding: 8px 12px; border-left: 2px solid {COR_ACC_700}; background: {COR_SURFACE};
  margin-bottom: 12px; word-break: break-word;
}}
</style>
""", unsafe_allow_html=True)


def _sparkline(valores, cor, altura_max=18):
    """Barras inline representando a série já calculada (puramente visual)."""
    vals = [v for v in valores if pd.notna(v)]
    if not vals:
        return ""
    topo = max(vals) or 1
    barras = "".join(
        f"<div style='width:6px;background:{cor};"
        f"height:{max(2, round(v / topo * altura_max))}px'></div>"
        for v in vals
    )
    return f"<div class='kpi-spark'>{barras}</div>"


def _kpi_tile(label, valor, delta_txt, delta_up, serie, destaque, rodape=""):
    """Tile de KPI no estilo Modernist. Recebe valores já calculados."""
    bg    = COR_ACC_100 if destaque else COR_SURFACE
    c_lbl = COR_ACC_800 if destaque else COR_MUTED
    c_val = COR_ACC_800 if destaque else COR_TEXT
    c_dlt = COR_ACC_700 if destaque else COR_MUTED
    c_spk = COR_ACC_500 if destaque else "rgba(32,30,29,.30)"
    seta  = "▲" if delta_up else "▼"
    delta_html = (f"<div class='kpi-delta' style='color:{c_dlt}'>{seta} {delta_txt}</div>"
                  if delta_txt else "")
    rodape_html = f"<div class='kpi-foot' style='color:{c_lbl}'>{rodape}</div>" if rodape else ""
    return (
        f"<div class='kpi-tile' style='background:{bg}'>"
        f"<div class='kpi-label' style='color:{c_lbl}'>{label}</div>"
        f"<div style='display:flex;align-items:baseline;gap:8px;margin-top:6px'>"
        f"<div class='kpi-value' style='color:{c_val}'>{valor}</div>{delta_html}</div>"
        f"{_sparkline(serie, c_spk)}{rodape_html}</div>"
    )


# ── Leitura de dados (cache) ──────────────────────────────────────────────────
def _base_dir():
    """Raiz do projeto — os dados ficam em `extracoes/` dentro dela."""
    return os.path.dirname(os.path.abspath(__file__))

def listar_jornadas():
    jornadas = fallout_core.jornadas_disponiveis(_base_dir())
    opcoes = list(jornadas)
    if "Base Móvel" in jornadas and "Cross Sell" in jornadas:
        opcoes.append("Base + Cross Sell")
    if len(jornadas) > 1:
        opcoes.append("Consolidado")   # todas as jornadas juntas
    return opcoes


def jornadas_kpi():
    """
    Jornadas de negócio exibidas nos KPIs do topo: cada pasta solta vira uma
    jornada e as pastas Base Móvel + Cross Sell contam como a combinada
    "Base + Cross Sell". Cresce sozinha conforme novas pastas aparecerem.
    """
    jornadas = fallout_core.jornadas_disponiveis(_base_dir())
    par_combo = {"Base Móvel", "Cross Sell"}
    tem_combo = par_combo.issubset(set(jornadas))
    soltas = sorted(j for j in jornadas if not (tem_combo and j in par_combo))
    return soltas + (["Base + Cross Sell"] if tem_combo else [])

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
st.markdown(
    f"<div style='font:600 11px/1 {FONTE};letter-spacing:.14em;text-transform:uppercase;"
    f"color:{COR_ACC_700};margin-bottom:6px'>Painel de acompanhamento</div>"
    f"<h2 style='margin:0 0 20px;font-size:30px'>Explorador de Fallout</h2>",
    unsafe_allow_html=True,
)

jornadas = listar_jornadas()

# ── Faixa 1: fallout por jornada + ações ──────────────────────────────────────
# Container reservado aqui e preenchido no fim: o conteúdo depende do mês
# escolhido no seletor abaixo, mas precisa aparecer acima dele.
faixa_topo = st.container()

# ── Faixa 2: filtros globais ──────────────────────────────────────────────────
_f_vazio, col_sel, col_jornada = st.columns([4, 1.6, 1.9])
with col_jornada:
    # O valor continua sendo o nome da pasta; só o rótulo muda (NBA aparece
    # como "Base Residencial").
    jornada_escolhida = st.selectbox(
        "Jornada", options=jornadas,
        format_func=fallout_core.nome_exibicao,
        index=jornadas.index("Base Móvel") if "Base Móvel" in jornadas else 0)
JORNADA_LABEL = fallout_core.nome_exibicao(jornada_escolhida)

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

# ── Tiles de KPI (valores já calculados acima / colunas já existentes) ────────
_meses_ord = sorted(resumo.index.tolist())
_mes_ant   = ([m for m in _meses_ord if m < mes_escolhido] or [None])[-1]


def _num_br(valor, casas=0):
    """Formata no padrão pt-BR: milhar com '.' e decimal com ','."""
    return f"{valor:,.{casas}f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def _delta(coluna, casas=0, sufixo=""):
    """Variação vs. mês anterior sobre valores que o resumo já traz."""
    if _mes_ant is None or mes_escolhido not in resumo.index:
        return "", False
    atual, anterior = resumo.loc[mes_escolhido, coluna], resumo.loc[_mes_ant, coluna]
    dif = atual - anterior
    return _num_br(abs(dif), casas) + sufixo, dif > 0


_serie_pct = [resumo.loc[m, "Pct"] for m in _meses_ord]
_d_pct, _up_pct = _delta("Pct", 2, "pp")
_d_tot, _up_tot = _delta("Total")
_d_suc, _up_suc = _delta("Sucessos")
_d_fal, _up_fal = _delta("Falhas")

_tem_mes = mes_escolhido in resumo.index
_v_suc = _num_br(int(resumo.loc[mes_escolhido, "Sucessos"])) if _tem_mes else "—"
_v_fal = _num_br(int(resumo.loc[mes_escolhido, "Falhas"]))   if _tem_mes else "—"

st.markdown(
    "<div class='kpi-row'>"
    + _kpi_tile("Fallout Rate", _num_br(fallout_pct, 2) + "%", _d_pct, _up_pct,
                _serie_pct, destaque=fallout_pct >= 1, rodape="Meta: abaixo de 1,00%")
    + _kpi_tile("Total de pedidos no mês", _num_br(total_mes), _d_tot, _up_tot,
                [resumo.loc[m, "Total"] for m in _meses_ord], destaque=False)
    + _kpi_tile("Vendas com sucesso", _v_suc, _d_suc, _up_suc,
                [resumo.loc[m, "Sucessos"] for m in _meses_ord], destaque=False)
    + _kpi_tile("Falha (análise técnica)", _v_fal, _d_fal, _up_fal,
                [resumo.loc[m, "Falhas"] for m in _meses_ord], destaque=False)
    + "</div>",
    unsafe_allow_html=True,
)

# ── Preenche a faixa 1 (fallout por jornada + ações) ──────────────────────────
_data_dir = _base_dir()


def _tile_jornada(nome_jornada):
    """Tile de fallout rate de uma jornada, no mesmo estilo dos demais."""
    _, resumo_j = carregar_dados(nome_jornada)
    if mes_escolhido not in resumo_j.index:
        return _kpi_tile(f"Fallout {fallout_core.nome_exibicao(nome_jornada)}", "—", "", False, [],
                         destaque=False, rodape="sem dados no mês")
    meses_j = sorted(resumo_j.index.tolist())
    pct_j   = resumo_j.loc[mes_escolhido, "Pct"]
    ant_j   = [m for m in meses_j if m < mes_escolhido]
    if ant_j:
        dif = pct_j - resumo_j.loc[ant_j[-1], "Pct"]
        delta_txt, delta_up = _num_br(abs(dif), 2) + "pp", dif > 0
    else:
        delta_txt, delta_up = "", False
    return _kpi_tile(
        f"Fallout {fallout_core.nome_exibicao(nome_jornada)}", _num_br(pct_j, 2) + "%", delta_txt, delta_up,
        [resumo_j.loc[m, "Pct"] for m in meses_j],
        destaque=pct_j >= 1, rodape="Meta: abaixo de 1,00%",
    )


with faixa_topo:
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:0 0 10px'>Fallout por jornada · "
        f"{mes_labels[mes_escolhido]}</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div class='kpi-row'>"
        + "".join(_tile_jornada(j) for j in jornadas_kpi())
        + "</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    col_a1, col_a2, _a_vazio = st.columns([1.5, 2.1, 5])
    with col_a1:
        if st.button("Atualizar dados", use_container_width=True):
            st.cache_data.clear()
            st.success("Cache limpo. Recarregando dados mais recentes...")
            st.rerun()
    with col_a2:
        if st.button("Baixar PowerPoint ↓", use_container_width=True):
            from gerar_slide3col import gerar_pptx_completo
            with st.spinner("Gerando PowerPoint completo (isso lê todos os dados, pode levar um tempo)..."):
                pptx_bytes = gerar_pptx_completo(_data_dir).getvalue()
            _PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            disparar_download(pptx_bytes, "relatorio_completo.pptx", _PPTX_MIME)
            st.success("PowerPoint completo gerado — o download deve começar automaticamente.")

    st.caption("PowerPoint: 3 colunas + Top Ofensores + um slide por jornada. "
               "Os dados são lidos da pasta `extracoes/` e ficam em cache — use "
               "**Atualizar dados** depois de trocar os arquivos por lá.")

    st.markdown("<hr>", unsafe_allow_html=True)   # separa faixa 1 da faixa 2

aba_distribuicao, aba_erros, aba_defeitos = st.tabs(
    ["Distribuição Fallout", "Análise de Erros", "Erros por Defeito"])

with aba_distribuicao:
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>"
        f"Distribuição Fallout · {mes_labels[mes_escolhido]}</div>",
        unsafe_allow_html=True,
    )

    # ── Cards clicáveis ───────────────────────────────────────────────────────
    cols = st.columns(3)
    cat_names = list(cats.keys())

    if "categoria_ativa" not in st.session_state:
        st.session_state.categoria_ativa = None

    for i, nome in enumerate(cat_names):
        qtd = len(cats[nome])
        pct = qtd / total_mes * 100 if total_mes > 0 else 0
        ativo = st.session_state.categoria_ativa == nome
        borda   = f"2px solid {COR_ACC_700}" if ativo else f"1px solid {COR_DIVIDER}"
        bg      = COR_ACC_100 if ativo else "#fff"
        c_pct   = COR_ACC_800 if ativo else COR_TEXT
        c_dot   = COR_ACC_700 if ativo else COR_DIVIDER
        qtd_fmt = _num_br(qtd)
        pct_fmt = _num_br(pct, 2)
        with cols[i % 3]:
            st.markdown(f"""
            <div style='border:{borda};background:{bg};padding:18px 20px;margin-bottom:8px'>
              <div style='display:flex;justify-content:space-between;align-items:flex-start;gap:10px'>
                <div style='font:600 14px {FONTE};line-height:1.3'>{nome}</div>
                <div style='width:7px;height:7px;border-radius:50%;flex:none;margin-top:5px;
                            background:{c_dot}'></div>
              </div>
              <div style='display:flex;align-items:baseline;gap:8px;margin-top:12px'>
                <div style='font:800 26px {FONTE};color:{c_pct}'>{pct_fmt}%</div>
                <div style='font:600 12px {FONTE};color:{COR_MUTED}'>do total</div>
              </div>
              <div style='font:600 12px {FONTE};color:{COR_MUTED};margin-top:2px'>{qtd_fmt} pedidos</div>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"Ver pedidos", key=f"btn_{i}", use_container_width=True):
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
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Análise de Erros · {JORNADA_LABEL}</div>",
        unsafe_allow_html=True,
    )

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
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Erros por Defeito · {JORNADA_LABEL} · "
        f"{mes_labels[mes_escolhido]}</div>",
        unsafe_allow_html=True,
    )

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

    # ── Renderização: um expander por grupo de erro ───────────────────────────
    from html import escape as _esc

    erros_unicos = grp_def["ErrorHandled__c"].unique()
    st.markdown(
        f"<div style='font:600 13px {FONTE};margin-bottom:10px'>"
        f"{_num_br(len(erros_unicos))} tipos de erro</div>",
        unsafe_allow_html=True,
    )

    for erro in erros_unicos:
        dfts       = grp_def[grp_def["ErrorHandled__c"] == erro].sort_values("Qtd", ascending=False)
        total_erro = int(dfts["Qtd"].sum())
        pct_erro   = total_erro / total_mes * 100 if total_mes > 0 else 0
        n_dfts     = len(dfts)

        titulo = (erro[:110] + "…") if len(erro) > 110 else erro
        sufixo = f" · {n_dfts} DFTs" if n_dfts > 1 else ""
        with st.expander(f"{titulo}  —  {total_erro} ocorrências ({_num_br(pct_erro, 2)}%){sufixo}"):
            st.markdown(
                f"<div class='raw-msg'>{_esc(str(erro))}</div>", unsafe_allow_html=True)

            chips, linhas_dft = [], []
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
                elif orig == "999999":
                    dft_label = "Pontual"
                    info      = ""
                else:
                    # alguns registros já vêm com o prefixo ("DFT 232143")
                    dft_label = orig if orig.upper().startswith("DFT") else "DFT " + orig
                    partes = [p for p in [nome, phase, ("entrega " + ms_s) if ms_s else "", team] if p]
                    info   = " · ".join(partes)

                chips.append(f"<span class='chip'>{_esc(dft_label)}</span>")
                cor_qtd = COR_ACC_700 if qtd >= 10 else (COR_ACC_500 if qtd >= 3 else COR_MUTED)
                linhas_dft.append(
                    f"<div style='display:flex;align-items:baseline;gap:12px;padding:7px 0;"
                    f"border-bottom:1px solid {COR_DIVIDER}'>"
                    f"<span style='font:800 13px {FONTE};color:{cor_qtd};min-width:34px'>{qtd}</span>"
                    f"<span style='font:600 13px {FONTE};min-width:120px'>{_esc(dft_label)}</span>"
                    f"<span style='font-size:12px;color:{COR_MUTED_2}'>{_esc(info)}</span>"
                    f"</div>"
                )

            st.markdown("<div>" + "".join(chips) + "</div>", unsafe_allow_html=True)
            st.markdown("".join(linhas_dft), unsafe_allow_html=True)
