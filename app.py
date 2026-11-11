import functools

import streamlit as st
import numpy as np
import pandas as pd
import os
from datetime import datetime
import distribuicao
import fallout_core
import fotografia
import parametros
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
/* O seletor acima alcança "*", e com isso também os ícones do Streamlit — que
   são ligaduras da Material Symbols. Trocada a fonte, a ligadura não resolve e
   o nome do ícone vaza como texto: a seta do expander virava a palavra
   "keyboard_arrow_right". Aqui a fonte de ícone é devolvida. */
[data-testid="stIconMaterial"] {{
  font-family: "Material Symbols Rounded" !important;
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
/* Setas de rolagem, quando as abas não cabem na largura: o degradê vem do tema
   escuro do Streamlit e saía como um bloco preto sobre o fundo claro. */
[data-testid="stTabsScrollLeft"], [data-testid="stTabsScrollRight"] {{ color: {COR_TEXT}; }}
[data-testid="stTabsScrollRight"] {{
  background-image: linear-gradient(to right, rgba(243,242,242,0), {COR_BG} 40%) !important;
}}
[data-testid="stTabsScrollLeft"] {{
  background-image: linear-gradient(to left, rgba(243,242,242,0), {COR_BG} 40%) !important;
}}

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
/* — "Erros por Defeito": <details> nativo, não st.expander —
   O expander do Streamlit abre animando `height`: fixa a altura inline e só a
   solta ao fim da animação. Com "animações do Windows" desligadas (o caso desta
   máquina) o navegador reporta prefers-reduced-motion, a animação não roda, a
   altura fica presa no tamanho fechado e o conteúdo aberto fica cortado — a
   seta gira e nada aparece. O corpo daquela aba é HTML puro, sem widget dentro,
   então o <details> do próprio navegador faz o mesmo trabalho sem JS. */
details.erro-det {{
  border: 1px solid {COR_DIVIDER}; background: #fff; margin-bottom: 2px;
}}
details.erro-det > summary {{
  cursor: pointer; padding: .6rem .9rem; font: 600 13px {FONTE};
  list-style: none; display: flex; gap: 10px; align-items: baseline;
}}
details.erro-det > summary::-webkit-details-marker {{ display: none; }}
/* Triângulo em texto, e não ícone de fonte: a seta do Streamlit é uma ligadura
   da Material Symbols, que a regra global de font-family acima derrubava. */
details.erro-det > summary::before {{
  content: "▸"; color: {COR_ACC_700}; font-size: 11px; flex: none;
}}
details.erro-det[open] > summary::before {{ content: "▾"; }}
details.erro-det > summary:hover {{ background: rgba(32,30,29,.04); }}
details.erro-det .erro-corpo {{ padding: .2rem .9rem 1rem; }}

/* — aba Distribuição: uma linha por categoria, como no slide detalhado — */
.dist-cat {{
  display: flex; gap: 10px; align-items: baseline; padding: .6rem .9rem .6rem 1.75rem;
  font: 600 13px {FONTE}; border: 1px solid {COR_DIVIDER}; background: #fff;
  margin-bottom: 2px;
}}
details.erro-det.dist-sub > summary {{ padding-left: 2.2rem; font-weight: 500; }}
details.erro-det.dist-sub {{ margin-left: 1.2rem; }}
/* segundo nível: o c/ data e s/ data dentro de US e DFT */
details.erro-det.dist-sub2 {{ margin-left: 2.6rem; }}
details.erro-det.dist-sub2 > summary {{ font-weight: 400; }}
.dist-num {{ margin-left: auto; white-space: nowrap; font-variant-numeric: tabular-nums; }}
.dist-num .qtd {{ font-weight: 500; color: {COR_MUTED}; margin-left: 10px; }}

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

/* — aba Comparação: tabelas executivas em HTML —
   A regra global acima pinta todo elemento com COR_TEXT, então cada elemento
   com cor própria leva classe (seletor de duas classes, que ganha dela). */
.cmp-tab {{ width: 100%; border-collapse: collapse; table-layout: fixed; background: #fff; }}
.cmp-tab th {{
  font: 600 10px {FONTE}; letter-spacing: .08em; text-transform: uppercase;
  color: {COR_MUTED_2}; text-align: left; padding: 9px 12px;
  border-bottom: 2px solid {COR_DIVIDER};
}}
.cmp-tab td {{
  font-size: 13px; line-height: 1.4; padding: 9px 12px; vertical-align: top;
  border-bottom: 1px solid rgba(32,30,29,.12); overflow-wrap: anywhere;
}}
.cmp-tab .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
.cmp-tab .sub {{ display: block; font-size: 12px; color: {COR_MUTED}; margin-top: 2px; }}
.cmp-tab .fraco {{ color: {COR_MUTED_2}; }}
.cmp-tab tr.total td {{ font-weight: 600; border-top: 2px solid {COR_DIVIDER}; border-bottom: 0; }}
.cmp-tab .pior, .cmp-nota .pior {{ color: {COR_ACC_700}; }}
.cmp-tab .melhor, .cmp-nota .melhor {{ color: #1e6b3a; }}
.cmp-tab .barra {{ display: inline-block; height: 6px; vertical-align: middle; margin-left: 8px; }}
.cmp-tab .barra.pior {{ background: {COR_ACC_500}; }}
.cmp-tab .barra.melhor {{ background: #5aa574; }}
.cmp-tab .sinal {{ display: inline-block; padding: 2px 8px; font: 600 11px {FONTE}; white-space: nowrap; }}
.cmp-tab .sinal.pior {{ background: {COR_ACC_100}; color: {COR_ACC_800}; }}
.cmp-tab .sinal.atencao {{ background: #fbf0d9; color: #6b4500; }}
.cmp-tab .sinal.melhor {{ background: #e3f1e7; color: #1e5a33; }}
.cmp-nota {{ font-size: 12px; line-height: 1.5; color: {COR_MUTED_2}; margin: 8px 0 14px; }}
.cmp-nota b {{ color: {COR_TEXT}; }}
.cmp-nota li {{ color: {COR_MUTED_2}; margin-bottom: 2px; }}
.cmp-resumo {{ margin: 4px 0 12px; }}
.cmp-resumo .titulo {{ font: 600 16px/1.35 {FONTE}; margin-bottom: 6px; }}
.cmp-resumo ul {{ margin: 0; padding-left: 18px; }}
.cmp-resumo li {{ font-size: 13px; line-height: 1.55; margin-bottom: 2px; }}
.cmp-resumo li b {{ color: {COR_ACC_700}; }}
.cmp-resumo .ressalva {{ font-size: 12px; line-height: 1.5; color: {COR_MUTED}; margin-top: 8px; }}
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
    """As opções do seletor — as mesmas visões que a fotografia grava."""
    return fallout_core.visoes_disponiveis(_base_dir())


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

# A chave do cache é só a jornada: os dados valem até o "Atualizar dados", que
# limpa este cache e os de leitura do fallout_core. Não há estado de sessão
# aqui — todo usuário vê os mesmos arquivos —, então compartilhar entre
# sessões é o esperado.
@st.cache_data(show_spinner="Carregando dados... Na primeira abertura depois de "
                            "um export novo em XLSX, a planilha é convertida para "
                            "CSV e isso pode levar até um minuto.")
def carregar_dados(jornada: str):
    base = _base_dir()
    df, _df_octane, resumo, _jornadas_combo = fallout_core.carregar_base(base, jornada)
    return df, resumo


@st.cache_data(show_spinner="Calculando o relatório da jornada...")
def carregar_relatorio(jornada: str):
    """
    Roda `pipeline.calcular_relatorio` uma vez e devolve o que as abas precisam.

    Reaproveita o pipeline em vez de repetir as regras aqui: quem entra na
    tabela de ofensores (defeito em aberto, com ocorrência no mês mais recente)
    e como a projeção é montada são regras sutis, que têm de sair iguais no app
    e no PowerPoint — se divergirem, a conferência contra o slide deixa de
    fechar e ninguém percebe.

    Uma chamada só para as duas abas, sobre a base que `carregar_dados` já tem
    em cache, e sem a figura do slide: antes o pipeline relia a base e montava
    o dashboard inteiro, que o app jogava fora.

    O `ns` do pipeline carrega closures, que não passam pelo cache do
    Streamlit — por isso aqui tudo vira tipo simples.
    """
    import pipeline

    df_j, resumo_base = carregar_dados(jornada)
    ns = pipeline.calcular_relatorio(_base_dir(), jornada, base=(df_j, resumo_base))
    pivot2, meses_tab, resumo_j = ns["pivot2"], ns["meses_tab"], ns["resumo"]
    pct_m, fmt_id, fmt_ms = ns["pct_m"], ns["fmt_id"], ns["fmt_milestone"]
    # meses_tab vem como chave AAAAMM; aqui basta a abreviação, já que são
    # sempre os dois meses mais recentes.
    meses = [fallout_core.MESES_LABEL[m][:3].capitalize() for m in meses_tab]

    linhas = []
    for _, row in pivot2.iterrows():
        tipo = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
        reg = {
            "Tipo": "US" if tipo == "User Story" else "Defeito",
            "Defeito / US": fmt_id({"DefectKey": row["DefectKey"],
                                    "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]}),
            "Data de correção": fmt_ms(row["DFT_BugfixMilestone"]),
            "Nome Defeito": str(row["DFT_Name"]) if pd.notna(row["DFT_Name"]) else "",
            "Time": str(row["DFT_Team"]) if pd.notna(row["DFT_Team"]) else "",
            "Status Defeito/US": str(row["DFT_Phase"]) if pd.notna(row["DFT_Phase"]) else "",
        }
        for m, lbl in zip(meses_tab, meses):
            n = int(row.loc[m]) if m in row.index else 0
            reg[f"{lbl} (pedidos)"] = n
            reg[f"{lbl} (%)"] = pct_m(n, m)
        linhas.append(reg)

    # Taxa de sucesso da jornada nos mesmos meses — no slide ela é a primeira
    # linha da tabela; aqui vira legenda, que numa tabela ordenável uma linha
    # que não é ofensor só atrapalharia.
    sucesso = {
        lbl: pct_m(resumo_j.loc[m, "Sucessos"] if m in resumo_j.index else 0, m)
        for m, lbl in zip(meses_tab, meses)
    }

    projecao = {
        "mes_atual": fallout_core.MESES_LABEL[ns["mes_atual"]].capitalize(),
        "fallout_pct": float(ns["fallout_pct"]),
        "total_mes": int(ns["total_mes"]),
        "pct_nao_recorrente": ns["pct_nao_recorrente"],
        "labels_hist": list(ns["labels_hist"]),
        "vals_hist": [float(v) for v in ns["vals_hist"]],
        "mes_fechado": bool(ns["mes_fechado"]),
        "plan_hist": [float(v) for v in ns["plan_hist"]],
        "proj_labels": list(ns["proj_labels"]),
        "proj_vals": [float(v) for v in ns["proj_vals"]],
        # A linha verde-escura do Base Residencial (vazias nas demais jornadas).
        "proj_sem_legado": [float(v) for v in ns.get("proj_sem_legado", [])],
        "plan_sem_legado_hist": [float(v) for v in ns.get("plan_sem_legado_hist", [])],
        "detalhe": list(ns["proj_detalhe"]),
        "reducao": [
            {"data": r["MilestoneDate"].strftime("%d/%m/%Y"),
             "dfts": [d.strip() for d in str(r["DFTs"]).split("\n")],
             "n": int(r["n"]),
             "pct_mes": float(r["Pct_mes"]),
             "pct_plena": float(r["Pct_plena"])}
            for _, r in ns["reducao"].iterrows()
        ],
    }
    return {"ofensores": (pd.DataFrame(linhas), meses, sucesso),
            "projecao": projecao}


def carregar_ofensores(jornada: str):
    """Ofensores da jornada — ver `carregar_relatorio`."""
    return carregar_relatorio(jornada)["ofensores"]


def carregar_projecao(jornada: str):
    """Série realizada, projeção e as parcelas da conta — ver `carregar_relatorio`."""
    return carregar_relatorio(jornada)["projecao"]


# Recebe, além do id de cada fotografia, a data em que ela foi gravada: gravar
# de novo a mesma extração troca a fotografia sem trocar o id, e o resultado em
# cache tem de trocar junto. Os parâmetros não podem começar com "_", que o
# Streamlit deixa fora da chave do cache.
@st.cache_data(show_spinner="Comparando a jornada...")
def carregar_comparacao(base_id: str, base_gravada, atual_id: str, atual_gravada,
                        visao: str, mes: int, min_ocorrencias: int, fator_alta: float):
    """A comparação completa de uma jornada (ver `comparacao.comparar`)."""
    import comparacao

    limites = comparacao.Limites(min_ocorrencias=min_ocorrencias, fator_alta=fator_alta)
    return comparacao.comparar(
        fotografia.carregar(_base_dir(), base_id),
        fotografia.carregar(_base_dir(), atual_id), visao, mes=mes, limites=limites)

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

class _barra_progresso:
    """
    Barra de progresso para os geradores de relatório, que levam de 5 a 15 s e
    antes só mostravam um spinner — sem dizer em que jornada estavam. Passada
    como `progresso` para eles, que a chamam com (feitos, total, rótulo).
    """

    def __init__(self, titulo):
        self._titulo = titulo
        self._barra = st.progress(0.0, text=f"{titulo}...")

    def __call__(self, feitos, total, rotulo):
        fracao = min(max(feitos / total, 0.0), 1.0) if total else 1.0
        self._barra.progress(fracao, text=f"{self._titulo}: {rotulo} ({feitos + 1}/{total})")

    def limpar(self):
        self._barra.empty()


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

# ── Google Drive (opcional) ───────────────────────────────────────────────────
# Com a chave do Drive ao lado do app, os dados vêm da pasta do Drive, espelhada
# em extracoes/_drive (ver drive_sync.py). Sincroniza na primeira abertura, a
# cada 30 min e no "Atualizar dados". O espelho é o mesmo para todas as
# sessões, por isso o cache_resource.
@st.cache_resource(show_spinner="Sincronizando com o Google Drive...", ttl=1800)
def _sincronizar_drive():
    import drive_sync
    try:
        return drive_sync.sincronizar(_base_dir()), None
    except drive_sync.ErroDrive as e:
        return None, str(e)
    except Exception as e:                               # noqa: BLE001 — sem rede etc.
        return None, f"{type(e).__name__}: {e}"


def _limpar_caches():
    st.cache_data.clear()
    # A data de corte é cacheada fora do Streamlit (lru_cache), então
    # st.cache_data.clear() sozinho deixaria o corte antigo de pé.
    fallout_core.limpar_cache_corte()
    fallout_core.limpar_cache_historico()


_ROTULO_FONTE = {fallout_core.FONTE_PC: "PC", fallout_core.FONTE_DRIVE: "Google Drive"}


def _seletor_fonte():
    """
    PC ou Google Drive. A escolha é gravada ao lado do app (ver
    `fallout_core.fonte_escolhida`) e troca a fonte de tudo, inclusive do
    PowerPoint e da fotografia. Sem a chave do Drive a escolha volta para o PC,
    com um aviso do que falta — desabilitado, o rádio do Streamlit pintava a
    opção desmarcada como se fosse a marcada.
    """
    base = _base_dir()

    def _trocar():
        nova = st.session_state["fonte_dados"]
        fonte = next(k for k, v in _ROTULO_FONTE.items() if v == nova)
        if fonte == fallout_core.FONTE_DRIVE and not fallout_core.drive_disponivel(base):
            st.session_state["_aviso_sem_drive"] = True
            return
        fallout_core.gravar_fonte(base, fonte)
        if fonte == fallout_core.FONTE_DRIVE:
            _sincronizar_drive.clear()
        _limpar_caches()

    # Reposto a cada execução: o arquivo é quem manda.
    st.session_state["fonte_dados"] = _ROTULO_FONTE[fallout_core.fonte_escolhida(base)]
    st.radio("Dados de", list(_ROTULO_FONTE.values()), key="fonte_dados",
             horizontal=True, on_change=_trocar)
    if st.session_state.pop("_aviso_sem_drive", False):
        st.warning(f"O Google Drive ainda não está configurado: coloque "
                   f"`{fallout_core.ARQ_CHAVE_DRIVE}` e `drive_pasta.txt` ao lado "
                   f"do app ({base}).")


# O seletor vem antes de qualquer leitura: a sincronização e o carregamento
# podem demorar, e a escolha precisa estar à mão durante a espera (ou quando o
# Drive falha). A linha de status embaixo dele é preenchida depois.
_c_fonte_vazio, _c_fonte = st.columns([4, 2.6])
with _c_fonte:
    _seletor_fonte()
    _status_fonte = st.empty()

_drive, _erro_drive = (_sincronizar_drive() if fallout_core.usa_drive(_base_dir())
                       else (None, None))
if _drive is not None:
    _status_fonte.caption(f"{_drive.pasta} · sincronizado às {_drive.quando:%H:%M}")
elif not fallout_core.usa_drive(_base_dir()):
    _status_fonte.caption(f"Pasta {fallout_core.PASTA_DADOS} deste computador")
if _erro_drive:
    st.warning(f"Não deu para sincronizar com o Google Drive: {_erro_drive}\n\n"
               "O app segue com a última cópia baixada do Drive, se houver. "
               "Para usar os arquivos do computador, escolha **PC** em "
               "\"Dados de\".")
elif _drive is not None:
    # Arquivo novo no Drive invalida o que já estava em cache — uma vez por
    # sincronização, e antes de qualquer leitura desta execução.
    if _drive.mudou and not getattr(_drive, "caches_limpos", False):
        _drive.caches_limpos = True
        _limpar_caches()
    for _falha in _drive.falhas:
        st.warning(f"Arquivo do Drive não baixado: {_falha}")

jornadas = listar_jornadas()

# Pasta de dados vazia é o primeiro estado de quem acabou de receber a versão
# portátil: ela vai sem as extrações de propósito. Sem esta parada o seletor de
# jornada fica sem opção, `jornada_escolhida` vira None e a pessoa recebe um
# "TypeError: join() argument must be str... not 'NoneType'" cru, vindo lá de
# dentro do carregador — que não diz nada sobre o que fazer.
if not jornadas and fallout_core.usa_drive(_base_dir()):
    if _drive is not None:
        st.warning(
            f"A pasta do Google Drive **{_drive.pasta}** não tem nenhuma jornada "
            f"({_drive.total} arquivos sincronizados). Ela precisa ter os dois "
            f"relatórios do Octane (`{fallout_core.ARQ_DFT}` e `{fallout_core.ARQ_US}`) "
            "mais uma pasta por jornada."
        )
    if st.button("Atualizar dados", use_container_width=False):
        _sincronizar_drive.clear()
        _limpar_caches()
        st.rerun()
    st.stop()
if not jornadas:
    _pasta_dados = os.path.join(_base_dir(), fallout_core.PASTA_DADOS)
    st.warning(
        f"Nenhuma jornada encontrada em `{_pasta_dados}`.\n\n"
        "Essa pasta vem vazia de propósito: os dados não são distribuídos "
        "junto com o programa. Coloque nela os dois relatórios do Octane "
        f"(`{fallout_core.ARQ_DFT}` e `{fallout_core.ARQ_US}`) mais uma pasta "
        "por jornada, e clique no botão abaixo.\n\n"
        "O **LEIA-ME.txt**, ao lado do atalho, traz a estrutura completa."
    )
    if st.button("Atualizar dados", use_container_width=False):
        _limpar_caches()
        st.rerun()
    st.stop()

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
# Os meses do resumo, não os do bruto: os meses fechados vêm do histórico
# manual e não têm mais linha a linha, mas continuam no gráfico e nos KPIs.
meses_disp = sorted(resumo.index.tolist())
_meses_com_bruto = set(df["Mes"].unique().tolist())
mes_labels = {m: fallout_core.MESES_LABEL[m].capitalize() for m in meses_disp}

# O app abre no mês do relatório, pela mesma regra dos slides: o mais recente
# com dados (ver fallout_core.mes_do_relatorio).
_mes_rel = fallout_core.mes_do_relatorio(df, resumo)
_mes_padrao = (meses_disp.index(_mes_rel) if _mes_rel in meses_disp
               else len(meses_disp) - 1)

with col_sel:
    mes_escolhido = st.selectbox(
        "Mês de análise",
        options=meses_disp,
        format_func=lambda m: mes_labels[m],
        index=_mes_padrao,
    )

# A planilha de acompanhamento traz a taxa de fallout mas não o volume, então
# um mês fechado pode ter Pct e não ter Sucessos/Falhas. `mes_tem_volume` guarda
# essa diferença; quem detalha pedido a pedido usa `MES_SEM_BRUTO`, logo abaixo.
_linha_mes = resumo.loc[mes_escolhido] if mes_escolhido in resumo.index else None
_total_bruto = float(_linha_mes["Total"]) if _linha_mes is not None else float("nan")
mes_tem_volume = _total_bruto == _total_bruto          # False quando NaN
# Sem volume digitado, o denominador das abas que contam pedido a pedido é o
# total da extração (ver fallout_core.total_do_mes); 0 quando nem ela tem o mês.
_total_denominador = fallout_core.total_do_mes(resumo, mes_escolhido)
total_mes = int(_total_denominador) if _total_denominador == _total_denominador else 0
fallout_pct = float(_linha_mes["Pct"]) if _linha_mes is not None else 0.0

# Mês fechado: o volume vem do histórico digitado, mas as abas que abrem
# pedido a pedido não têm o que mostrar. Sem este aviso elas apareceriam
# zeradas, contradizendo o fallout rate exibido logo acima.
MES_SEM_BRUTO = mes_escolhido not in _meses_com_bruto
AVISO_SEM_BRUTO = (
    f"**{mes_labels[mes_escolhido]}** é um mês fechado: os totais vêm do "
    f"histórico manual (`extracoes/historico.xlsx`) e não há mais o detalhe "
    f"pedido a pedido. Escolha um mês recente para explorar este conteúdo."
)

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
    if dif != dif:                       # NaN: algum dos meses não tem volume
        return "", False
    return _num_br(abs(dif), casas) + sufixo, dif > 0


_serie_pct = [resumo.loc[m, "Pct"] for m in _meses_ord]
_d_pct, _up_pct = _delta("Pct", 2, "pp")
_d_tot, _up_tot = _delta("Total")
_d_suc, _up_suc = _delta("Sucessos")
_d_fal, _up_fal = _delta("Falhas")

_tem_mes = mes_escolhido in resumo.index
_v_suc = _num_br(int(resumo.loc[mes_escolhido, "Sucessos"])) if (_tem_mes and mes_tem_volume) else "—"
_v_fal = _num_br(int(resumo.loc[mes_escolhido, "Falhas"]))   if (_tem_mes and mes_tem_volume) else "—"

st.markdown(
    "<div class='kpi-row'>"
    + _kpi_tile("Fallout Rate", _num_br(fallout_pct, 2) + "%", _d_pct, _up_pct,
                _serie_pct, destaque=fallout_pct >= 1, rodape="Meta: abaixo de 1,00%")
    + _kpi_tile("Total de pedidos no mês",
                _num_br(total_mes) if mes_tem_volume else "—", _d_tot, _up_tot,
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
    col_a1, col_a2, col_a3, col_a4, _a_vazio = st.columns([1.5, 2.1, 2.4, 2.0, 2.0])
    with col_a1:
        if st.button("Atualizar dados", use_container_width=True):
            _sincronizar_drive.clear()
            _limpar_caches()
            st.success("Cache limpo. Recarregando dados mais recentes...")
            st.rerun()
    with col_a2:
        if st.button("Baixar PowerPoint ↓", use_container_width=True):
            from gerar_slide3col import gerar_pptx_completo
            _barra = _barra_progresso("Gerando PowerPoint")
            with st.spinner("Gerando PowerPoint completo (isso lê todos os dados, pode levar um tempo)..."):
                pptx_bytes = gerar_pptx_completo(_data_dir, progresso=_barra).getvalue()
            _barra.limpar()
            _PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            # Data e hora no nome para os relatórios não se sobrescreverem na
            # pasta de downloads; o formato ordena junto com a ordem cronológica.
            _nome = f"RelatorioExecutivo_{datetime.now():%Y-%m-%d_%H%M}.pptx"
            disparar_download(pptx_bytes, _nome, _PPTX_MIME)
            st.success(f"PowerPoint gerado: **{_nome}** — o download deve começar automaticamente.")

    with col_a3:
        if st.button("Baixar Ofensores (Excel) ↓", use_container_width=True):
            import gerar_ofensores_xlsx
            _barra_x = _barra_progresso("Gerando planilha de ofensores")
            with st.spinner("Gerando planilha de ofensores (uma aba por jornada, "
                            "isso lê todos os dados)..."):
                _xlsx, _feitas, _avisos = gerar_ofensores_xlsx.gerar(
                    _data_dir, listar_jornadas(), progresso=_barra_x)
            _barra_x.limpar()
            _XLSX_MIME = ("application/vnd.openxmlformats-officedocument."
                          "spreadsheetml.sheet")
            _nome_x = f"Ofensores_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
            disparar_download(_xlsx, _nome_x, _XLSX_MIME)
            st.success(f"Planilha gerada: **{_nome_x}** — {len(_feitas)} "
                       "jornadas na dinâmica, o download deve começar "
                       "automaticamente.")
            for _aviso in _avisos:
                st.warning(f"Jornada fora da planilha — {_aviso}")

    with col_a4:
        if st.button("Gravar fotografia", use_container_width=True):
            _barra_f = _barra_progresso("Gravando a fotografia")
            with st.spinner("Gravando a fotografia desta extração (isso lê todos "
                            "os dados, pode levar um tempo)..."):
                _foto = fotografia.gerar(_data_dir, progresso=_barra_f)
                fotografia.salvar(_foto, _data_dir)
            _barra_f.limpar()
            # A fotografia lê os arquivos de agora; a tela, o que estava em
            # cache. Recarregar depois de gravar deixa as duas com os mesmos
            # números — senão a tela seguia com a extração anterior até o
            # "Atualizar dados". O aviso atravessa o rerun pelo session_state.
            st.session_state["_aviso_fotografia"] = (_foto.id, dict(_foto.meta["erros"]))
            st.cache_data.clear()
            fallout_core.limpar_cache_corte()
            fallout_core.limpar_cache_historico()
            st.rerun()
        _aviso_foto = st.session_state.pop("_aviso_fotografia", None)
        if _aviso_foto:
            _id_foto, _erros_foto = _aviso_foto
            st.success("Fotografia gravada: extração com dados até "
                       f"**{datetime.fromisoformat(_id_foto):%d/%m/%Y}**.")
            for _visao, _erro in _erros_foto.items():
                st.warning("Visão fora da fotografia — "
                           f"{fallout_core.nome_exibicao(_visao)}: {_erro}")

    st.markdown("<hr>", unsafe_allow_html=True)   # separa faixa 1 da faixa 2

# Com as abas lazy, os widgets de uma aba fechada não são desenhados, e o
# Streamlit descarta o estado de widget que não aparece na execução: o filtro
# digitado sumia a cada troca de aba. Reatribuir a chave antes de desenhar é o
# jeito documentado de segurar o valor. Ficam de fora as datas, cujo padrão
# acompanha o mês escolhido no topo — fixá-las mudaria esse comportamento.
_FILTROS_DAS_ABAS = ("filtro_ofensores", "cor_motivo", "cor_busca", "erros_agrupar",
                     "erros_busca", "erros_filtro_dft", "def_busca", "def_min_occ",
                     "cmp_atual", "cmp_base", "cmp_min_occ", "cmp_fator",
                     "cmp_busca_def", "cmp_mostrar_texto", "cmp_dist_completa",
                     "cmp_def_todos", "cmp_err_detalhe", "par_jornadas", "par_busca",
                     "dist_cat", "dist_busca", "dia_janela", "prev_foto")
for _chave in _FILTROS_DAS_ABAS:
    if _chave in st.session_state:
        st.session_state[_chave] = st.session_state[_chave]

# `on_change="rerun"` faz as abas guardarem qual está aberta, e cada uma só
# executa quando é a aberta (ver `_render_aba`). Sem isso o Streamlit roda
# todas a cada clique em qualquer widget — a Análise de Erros sozinha custava
# ~5 s por interação no Base Residencial, mesmo com outra aba na tela.
(aba_distribuicao, aba_fallout_dia, aba_projecao, aba_previsto, aba_ofensores,
 aba_correcao, aba_erros, aba_defeitos, aba_comparacao, aba_parametros) = st.tabs(
    ["Distribuição Fallout", "Fallout por Dia", "Projeção", "Previsto x Realizado",
     "Ofensores", "Precisa de Correção", "Análise de Erros", "Erros por Defeito",
     "Comparação", "Parâmetros"],
    key="aba_ativa", on_change="rerun")


def _render_aba(aba, corpo):
    """Desenha a aba só quando ela é a selecionada."""
    with aba:
        if aba.open:
            corpo()


def _filtrar_texto(tabela, termo):
    """
    Linhas em que alguma coluna contém `termo`, sem diferenciar maiúscula.

    Uma busca por coluna, e não `apply(axis=1)`: aquele montava uma Series por
    linha e levava ~5,7 s em 7 mil pedidos, contra ~0,2 s assim. O termo é
    literal — parêntese e ponto, comuns nas mensagens de erro, não são regex.
    """
    mask = pd.Series(False, index=tabela.index)
    for col in tabela.columns:
        mask |= tabela[col].astype(str).str.contains(
            termo, case=False, regex=False, na=False)
    return tabela[mask]


# Rótulo sem defeito -> nome da categoria que ele tem na Distribuição, como já
# era feito com "Enviado e-mail - Outros Times". Rótulo que não está aqui
# aparece como veio.
_ETIQUETAS_SEM_DEFEITO = {
    "enviado e-mail": "Em Avaliação por MOPs",
    "erro de processo": "Falha no Processo Usuário",
    fallout_core.ROTULO_INCIDENTE: fallout_core.CAT_INCIDENTE,
}


def _dfts_por_erro(df_err):
    """
    Coluna "DFTs" da Análise de Erros: tudo o que está associado a cada
    mensagem de erro, como (ErrorHandled__c, DFTs). Três partes, nesta ordem:

      1. defeitos do Octane, pelo número — "DFT · entrega · status";
      2. defeitos identificados por rótulo, que o Octane não numera mas o
         relatório trata como defeito (`TemDefeito`): BRJ-* do Base
         Residencial, chamados INC-*/PDST-* do Jira, GML-*, "Paliativo RPA" e
         "NFCOM" do PME — mesmo formato, com o rótulo no lugar de "DFT";
      3. rótulos que não são defeito — Falha Pontual, Outros Times, MOPs,
         Erro de Processo, Incidente e os demais —, só o nome, no fim.

    Antes só a parte 1 e duas etiquetas da 3 apareciam: no Base Residencial,
    44.895 dos 46.035 pedidos ficavam com a coluna vazia.

    Montado por coluna, e não num laço linha a linha: o laço levava de 1 a 5 s
    no Base Residencial a cada interação com a aba. De quebra, um filtro sem
    resultado devolve a tabela vazia, em vez de derrubar a aba como o
    `groupby.apply`.
    """
    erro = df_err["ErrorHandled__c"]
    # `x or ""` como no laço original: NaN é truthy e vira o texto "nan".
    orig_txt = df_err["DefectNumber_orig"].map(lambda v: str(v or "")).str.strip()
    orig = orig_txt.str.lower()
    # "0" é a marca de "sem defeito" do relatório de RPA.
    vazio = orig.isin(["", "0", "nan", "none"])
    outros = orig == "enviado e-mail - outros times"
    dft = df_err["DefectNumber__c"]
    dft_num = pd.to_numeric(dft, errors="coerce")
    numerado = ~outros & dft.notna() & (dft_num != -1)
    pontual = numerado & (dft_num == 999999)
    com_dft = numerado & (dft_num != 999999)
    por_rotulo = (df_err["TemDefeito"].fillna(False).astype(bool)
                  & ~com_dft & ~pontual & ~outros & ~vazio)
    etiqueta = ~(com_dft | pontual | outros | por_rotulo | vazio)

    def _fmt_ms(v):
        ms = pd.to_datetime(v, errors="coerce")
        return ms.strftime("%d/%m/%Y") if not pd.isna(ms) else "s/ data"

    def _fmt_fase(v):
        # pd.isna antes do `or`: NaN é truthy e saía na tela como "nan" — raro
        # no DFT do Octane, comum nos defeitos por rótulo (GML-*, NFCOM).
        return "s/ status" if pd.isna(v) else (str(v).strip() or "s/ status")

    def _por_valor(serie, fn):
        # Formata cada valor distinto uma vez, com a mesma conversão escalar do
        # laço: a coluna pode vir como objeto com tipos misturados, e a conversão
        # vetorizada inferiria o formato de outro jeito.
        codigos, unicos = pd.factorize(serie, use_na_sentinel=True)
        saida = np.array([fn(v) for v in unicos] + [""], dtype=object)[codigos]
        ausentes = np.flatnonzero(codigos == -1)
        if len(ausentes):
            # Os ausentes vão um a um: None e NaN dão textos diferentes.
            saida[ausentes] = [fn(v) for v in serie.iloc[ausentes]]
        return pd.Series(saida, index=serie.index, dtype=object)

    defeito = com_dft | por_rotulo
    ident = orig_txt.where(por_rotulo, "")
    ident[com_dft] = dft_num[com_dft].astype("int64").astype(str)
    prefixo = pd.Series("", index=df_err.index, dtype=object).mask(com_dft, "DFT")
    pares = pd.DataFrame({
        "erro": erro[defeito],
        "rotulo": por_rotulo[defeito],          # o número do Octane vem antes
        "id": ident[defeito],
        "texto": (prefixo[defeito] + ident[defeito]
                  + " · " + _por_valor(df_err.loc[defeito, "DFT_BugfixMilestone"], _fmt_ms)
                  + " · " + _por_valor(df_err.loc[defeito, "DFT_Phase"], _fmt_fase)),
    })
    # No laço, a última linha de cada DFT sobrescrevia as anteriores no dict.
    pares = (pares.drop_duplicates(["erro", "id"], keep="last")
             .sort_values(["erro", "rotulo", "id"], kind="stable"))
    textos = pares.groupby("erro", sort=False)["texto"].agg(list)

    nomes = orig[etiqueta].map(_ETIQUETAS_SEM_DEFEITO).fillna(orig_txt[etiqueta])
    etiquetas = (pd.DataFrame({"erro": erro[etiqueta], "nome": nomes})
                 .drop_duplicates()
                 .sort_values(["erro", "nome"], kind="stable")
                 .groupby("erro", sort=False)["nome"].agg(list))

    tem_pontual = pontual.groupby(erro).any()
    tem_outros = outros.groupby(erro).any()
    grupos = erro.drop_duplicates().sort_values(kind="stable")
    linhas = []
    for e in grupos:
        partes = list(textos.get(e, []))
        if tem_pontual.get(e, False):
            partes.append("Falha Pontual")
        if tem_outros.get(e, False):
            partes.append("Em avaliação - Outros Times")
        partes.extend(etiquetas.get(e, []))
        linhas.append(" | ".join(partes))
    return pd.DataFrame({"ErrorHandled__c": grupos.to_numpy(), "DFTs": linhas})


# Linhas por tabela aberta: a lista é HTML puro, e uma categoria como Falta
# Associar pode ter centenas de mensagens. O resto vira uma linha "… e mais N";
# a lista completa está nos pedidos, logo abaixo.
LIMITE_LINHAS_DIST = 60


def _pct_dist(valor):
    """% com duas casas; o que arredonda para zero sem ser zero vira "<0,01%"."""
    return "<0,01%" if 0 < valor < 0.005 else f"{_num_br(valor, 2)}%"


def _tabela_dist(tabela, por_mensagem):
    """A tabela que abre sob a linha da categoria: por defeito ou por mensagem."""
    from html import escape as _esc
    if por_mensagem:
        cab = ("<th>Mensagem de erro</th><th class='num' style='width:90px'>Pedidos</th>"
               "<th class='num' style='width:80px'>%</th>")
        linhas = [f"<td>{_esc(str(r['Mensagem']))}</td>"
                  f"<td class='num'>{_num_br(r['Pedidos'])}</td>"
                  f"<td class='num'>{_pct_dist(r['%'])}</td>"
                  for _, r in tabela.head(LIMITE_LINHAS_DIST).iterrows()]
        colunas = 3
    else:
        cab = ("<th style='width:120px'>Defeito/US</th><th>Nome</th>"
               "<th style='width:200px'>Equipe</th><th style='width:100px'>Entrega</th>"
               "<th class='num' style='width:90px'>Pedidos</th>"
               "<th class='num' style='width:80px'>%</th>")
        _t = lambda v: "" if pd.isna(v) else _esc(str(v))
        linhas = [f"<td><b>{_t(r['Defeito/US'])}</b></td><td>{_t(r['Nome'])}</td>"
                  f"<td>{_t(r['Equipe'])}</td>"
                  f"<td>{'' if pd.isna(r['Entrega']) else r['Entrega'].strftime('%d/%m/%Y')}</td>"
                  f"<td class='num'>{_num_br(r['Pedidos'])}</td>"
                  f"<td class='num'>{_pct_dist(r['%'])}</td>"
                  for _, r in tabela.head(LIMITE_LINHAS_DIST).iterrows()]
        colunas = 6
    resto = tabela.iloc[LIMITE_LINHAS_DIST:]
    if len(resto):
        item = "mensagem" if por_mensagem else "defeito"
        linhas.append(f"<td colspan='{colunas}' class='fraco'>… e mais {_num_br(len(resto))} "
                      f"{item}{'' if len(resto) == 1 else 's'} "
                      f"({_num_br(int(resto['Pedidos'].sum()))} pedidos, "
                      f"{_pct_dist(resto['%'].sum())})</td>")
    return (f"<table class='cmp-tab'><tr>{cab}</tr>"
            + "".join(f"<tr>{l}</tr>" for l in linhas) + "</table>")


def _linha_dist(nome, pedidos, detalhe=None, sub=0):
    """
    Uma linha da Distribuição; com `detalhe`, ela abre e mostra a tabela.
    `sub` é o nível do subgrupo: 1 (US/DFT, Legado) ou 2 (c/ e s/ data). O US
    e o DFT do Em Tratamento só agrupam, e saem sem os números.
    """
    from html import escape as _esc
    pct = len(pedidos) / total_mes * 100 if total_mes else 0
    numeros = ("" if fallout_core.so_rotulo(nome, sub) else
               f"<span class='dist-num'>{_pct_dist(pct)}"
               f"<span class='qtd'>{_num_br(len(pedidos))} pedidos</span></span>")
    if detalhe is None:
        return f"<div class='dist-cat'>{_esc(nome)}{numeros}</div>"
    classes = "erro-det" + (" dist-sub" if sub else "") + (" dist-sub2" if sub == 2 else "")
    return (f"<details class='{classes}'>"
            f"<summary>{_esc(nome)}{numeros}</summary>"
            f"<div class='erro-corpo'>{detalhe}</div></details>")


def _aba_distribuicao():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>"
        f"Distribuição Fallout · {mes_labels[mes_escolhido]}</div>",
        unsafe_allow_html=True,
    )

    # Streamlit monta todas as abas numa passada só, então aqui não cabe
    # st.stop(): ele derrubaria as outras abas junto.
    if MES_SEM_BRUTO:
        st.info(AVISO_SEM_BRUTO)
        return

    # ── A Distribuição em linhas, como no slide detalhado ────────────────────
    # Em Tratamento no topo, com US e DFT logo abaixo e, dentro de cada um, o
    # c/ data e o s/ data; as demais categorias da maior para a menor. Cada
    # linha abre o que a preenche: os defeitos, ou as mensagens de erro nas
    # categorias sem defeito (ver distribuicao.py).
    em_trat = cats[fallout_core.CAT_TRATAMENTO]
    subgrupo = fallout_core.subgrupo_tratamento(em_trat)
    tipo = fallout_core.tipo_do_subgrupo(subgrupo)
    grupos = {}                     # rótulo da lista de pedidos -> pedidos

    def _linha_defeitos(rotulo, parte, nivel):
        return _linha_dist(rotulo, parte,
                           _tabela_dist(distribuicao.por_defeito(parte, total_mes), False),
                           sub=nivel)

    html_linhas = []
    if len(em_trat):
        html_linhas.append(_linha_dist(fallout_core.CAT_TRATAMENTO, em_trat))
        grupos[fallout_core.CAT_TRATAMENTO] = em_trat
        # Do maior para o menor em cada nível, como nos slides.
        for nome_tipo, subs in sorted(fallout_core.SUBGRUPOS_TRATAMENTO.items(),
                                      key=lambda par: -(tipo == par[0]).sum()):
            do_tipo = em_trat[tipo == nome_tipo]
            if not len(do_tipo):
                continue
            html_linhas.append(_linha_defeitos(nome_tipo, do_tipo, 1))
            grupos[f"{fallout_core.CAT_TRATAMENTO} · {nome_tipo}"] = do_tipo
            for nome in sorted(subs, key=lambda s: -(subgrupo == s).sum()):
                parte = em_trat[subgrupo == nome]
                if len(parte):
                    html_linhas.append(_linha_defeitos(fallout_core.rotulo_dist(nome, 2),
                                                       parte, 2))
                    grupos[f"{fallout_core.CAT_TRATAMENTO} · {nome}"] = parte
    demais = sorted(((n, p) for n, p in cats.items()
                     if n != fallout_core.CAT_TRATAMENTO and len(p)),
                    key=lambda par: -len(par[1]))
    for nome, pedidos in demais:
        # Legado Aberto (Base Residencial) também abre em subgrupos, como o Em
        # Tratamento: o "Feito manualmente" pelo defeito, a "Funcionalidade
        # idêntica" pela mensagem de erro, que é o que a define.
        if nome == fallout_core.CAT_LEGADO:
            html_linhas.append(_linha_dist(nome, pedidos))
            grupos[nome] = pedidos
            sub_legado = fallout_core.subgrupo_legado(pedidos)
            for sub in (fallout_core.SUB_LEGADO_MANUAL, fallout_core.SUB_LEGADO_FUNCIONALIDADE):
                parte = pedidos[sub_legado == sub]
                if not len(parte):
                    continue
                por_msg = sub == fallout_core.SUB_LEGADO_FUNCIONALIDADE
                tabela = (distribuicao.por_mensagem(parte, total_mes) if por_msg
                          else distribuicao.por_defeito(parte, total_mes))
                html_linhas.append(_linha_dist(sub, parte, _tabela_dist(tabela, por_msg),
                                               sub=1))
                grupos[f"{nome} · {sub}"] = parte
            continue
        por_msg = nome in distribuicao.POR_MENSAGEM
        tabela = (distribuicao.por_mensagem(pedidos, total_mes) if por_msg
                  else distribuicao.por_defeito(pedidos, total_mes))
        html_linhas.append(_linha_dist(nome, pedidos, _tabela_dist(tabela, por_msg)))
        grupos[nome] = pedidos

    st.markdown(
        f"<div style='font:600 14px {FONTE};margin:0 0 10px'>"
        f"{_num_br(len(df_mes))} falhas de {_num_br(total_mes)} pedidos · "
        f"fallout {_pct_dist(len(df_mes) / total_mes * 100 if total_mes else 0)}</div>"
        + "".join(html_linhas),
        unsafe_allow_html=True,
    )
    st.caption("Clique numa linha para ver o que a preenche: os defeitos (Defeito/US, "
               "nome, equipe, data de entrega) ou, nas categorias sem defeito, as "
               "mensagens de erro. O % é sobre o total de pedidos do mês. DFT com US "
               "de Melhoria aparece pelo número da US.")

    # ── Pedidos de uma categoria ─────────────────────────────────────────────
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:22px 0 6px'>Pedidos de uma categoria</div>",
        unsafe_allow_html=True,
    )
    # A escolha fica guardada entre execuções; na troca de jornada ou de mês a
    # categoria pode não existir mais, e a caixa não aceita valor fora da lista.
    if st.session_state.get("dist_cat") not in grupos:
        st.session_state["dist_cat"] = None
    escolha = st.selectbox("Categoria", list(grupos), index=None,
                           placeholder="Escolha uma categoria para ver os pedidos",
                           key="dist_cat", label_visibility="collapsed")
    if escolha is None or escolha not in grupos:
        return
    df_cat = grupos[escolha]
    cols_disp = [c for c in COLS_EXIBIR if c in df_cat.columns]
    df_show = df_cat[cols_disp].copy().reset_index(drop=True)
    if "DefectNumber_orig" in df_show.columns:
        df_show = df_show.rename(columns={"DefectNumber_orig": "DefectNumber__c"})
        df_show["DefectNumber__c"] = df_show["DefectNumber__c"].replace(
            {"nan": "", "-1": "", "999999": "999999 (Pontual)"})

    filtro = st.text_input("Filtrar por número de pedido, DFT ou qualquer campo", key="dist_busca")
    if filtro:
        df_show = _filtrar_texto(df_show, filtro)
    st.markdown(f"<div style='font:600 13px {FONTE};margin:4px 0 8px'>{escolha} — "
                f"{_num_br(len(df_show))} pedidos</div>", unsafe_allow_html=True)
    st.dataframe(df_show, use_container_width=True, height=400)

    nome_arquivo = (f"fallout_{escolha[:30].replace('/', '-').replace('·', '-').strip()}_"
                    f"{mes_labels[mes_escolhido]}.xlsx")
    # O Excel só é montado no clique: gerado aqui, custava ~5 s em todo
    # rerun com a lista aberta. O Streamlit chama a função depois que o
    # script terminou, por isso a tabela vai fixada no argumento padrão.
    st.download_button(
        label="⬇ Baixar Excel",
        data=lambda d=df_show: to_excel_bytes(d),
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


_render_aba(aba_distribuicao, _aba_distribuicao)


# ── Aba Fallout por Dia ───────────────────────────────────────────────────────
# O mesmo gráfico do slide "Evolução do Fallout por dia" (gerar_pptx), para a
# jornada escolhida no topo, com a tabela dos números logo abaixo.
DIAS_OPCOES = (7, 15, 30)


@st.cache_data(show_spinner="Calculando o fallout por dia...")
def carregar_fallout_dia(jornada, dias):
    """Tabela de `fallout_core.fallout_por_dia`; o df vem do cache da jornada."""
    df_j, _resumo_j = carregar_dados(jornada)
    return fallout_core.fallout_por_dia(df_j, _base_dir(), jornada, dias=dias)


def _tem_sucesso_diario(jornada):
    """
    Só as jornadas de extração trazem os sucessos separados por dia; o Base
    Residencial e a PME Móvel trazem sucesso e falha no mesmo arquivo, e lá
    não há o total de pedidos do dia para o percentual.
    """
    raiz = fallout_core.raiz_dados(_base_dir())
    return all(fallout_core.formato_jornada(os.path.join(raiz, j)) == "extracao"
               for j in fallout_core.resolver_jornadas(jornada, _base_dir()))


def _aba_fallout_dia():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Fallout por Dia · {JORNADA_LABEL}</div>",
        unsafe_allow_html=True,
    )
    if not _tem_sucesso_diario(jornada_escolhida):
        st.info(f"**{JORNADA_LABEL}** não traz os sucessos separados por dia, então não "
                "dá para calcular o fallout de cada dia. A visão diária existe para "
                "Consolidado, Base + Cross Sell e as jornadas de extração (Prospect, "
                "Base Móvel, Cross Sell e PME Fixa).")
        return

    # Caixa de seleção como os filtros do topo: o controle segmentado do
    # Streamlit não é coberto pelo tema do app e saía com os botões pretos.
    col_periodo, _ = st.columns([1.4, 4])
    with col_periodo:
        dias = st.selectbox("Período", DIAS_OPCOES, index=DIAS_OPCOES.index(15),
                            key="dia_janela", format_func=lambda n: f"Últimos {n} dias")
    tabela = carregar_fallout_dia(jornada_escolhida, dias)

    import gerar_pptx               # o mesmo desenho do slide, e só quando a aba abre
    validos = tabela[tabela["Pedidos"] > 0]
    if not len(validos):
        st.info(f"Os arquivos de sucesso de **{JORNADA_LABEL}** não trazem a quebra por "
                f"dia no período (até {tabela.index[-1]:%d/%m}) — só o total do mês —, "
                "então não há o total de pedidos de cada dia. Use a extração de "
                "sucessos com uma linha por dia.")
        return
    taxa = validos["Falhas"] / validos["Pedidos"] * 100
    pior = taxa.idxmax()
    st.markdown(
        f"<div style='font:600 14px {FONTE};margin:0 0 6px'>"
        f"{_num_br(int(validos['Falhas'].sum()))} falhas de "
        f"{_num_br(int(validos['Pedidos'].sum()))} pedidos no período · fallout "
        f"{_num_br(validos['Falhas'].sum() / validos['Pedidos'].sum() * 100, 2)}% · "
        f"pior dia {pior:%d/%m} ({_num_br(taxa[pior], 2)}%)</div>",
        unsafe_allow_html=True,
    )
    st.image(gerar_pptx._grafico_diario(tabela, 12.93, 6.8), use_container_width=True)

    # A tabela é a leitura exata do gráfico: pedidos, falhas e o % de cada
    # categoria por dia, com os mesmos nomes da legenda.
    categorias = [c for c in tabela.columns if c not in ("Falhas", "Sucessos", "Pedidos")]
    exibir = pd.DataFrame({
        "Dia": [d.strftime("%d/%m/%Y") for d in tabela.index],
        "Pedidos": tabela["Pedidos"].astype("Int64"),
        "Falhas": tabela["Falhas"].astype("Int64"),
        "Fallout": (tabela["Falhas"] / tabela["Pedidos"] * 100).where(tabela["Pedidos"] > 0),
    })
    for c in categorias:
        exibir[gerar_pptx.ROTULOS_DIARIO.get(c, c)] = (
            (tabela[c] / tabela["Pedidos"] * 100).where(tabela["Pedidos"] > 0).to_numpy())
    pct = st.column_config.NumberColumn(format="%.2f%%")
    st.dataframe(exibir, hide_index=True, use_container_width=True,
                 column_config={c: pct for c in exibir.columns[3:]},
                 height=min(38 + 35 * len(exibir), 560))
    st.caption("Um dia por linha, até o último dia completo; a aba não segue o mês "
               "selecionado no topo. Pedidos = sucessos + falhas do dia, e cada "
               "pedido está na categoria que tem hoje, a mesma da Distribuição. Dia "
               "sem sucessos na extração fica em branco: o total estaria incompleto.")


_render_aba(aba_fallout_dia, _aba_fallout_dia)

# Cor da meta no gráfico de projeção — a mesma do slide, para a aba e o
# PowerPoint não contarem a mesma história com aparências diferentes.
COR_META = "#1565C0"


def _fig_projecao(p):
    """Redesenha o gráfico do slide a partir das séries que o pipeline devolveu."""
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    import matplotlib.ticker as mticker

    fig = Figure(figsize=(11, 4.0))
    FigureCanvasAgg(fig)
    fig.patch.set_facecolor("white")
    ax = fig.add_subplot(111)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_facecolor("white")

    import pipeline

    # As mesmas funções de desenho dos slides: real contínuo até o mês passado
    # e pontilhado até o atual; uma linha verde só, Planejado no passado e
    # projeção no futuro. O real vem primeiro porque o eixo é categórico e a
    # ordem dos meses no eixo é a ordem em que eles aparecem nos plots.
    labels = p["labels_hist"]
    pipeline.desenhar_real(ax, labels, p["vals_hist"], fechado=p.get("mes_fechado", False))
    for m, v in zip(labels, p["vals_hist"]):
        ax.annotate(f"{_num_br(v, 2)}%", xy=(m, v), xytext=(0, 8),
                    textcoords="offset points", ha="center", fontsize=8,
                    **pipeline.ROTULO_SOBRE_LINHA)

    pipeline.desenhar_verde(ax, labels, p["plan_hist"],
                            p["proj_labels"], p["proj_vals"])
    for lbl, v in zip(p["proj_labels"][1:], p["proj_vals"][1:]):
        ax.annotate(f"{_num_br(v, 2)}%", xy=(lbl, v), xytext=(0, 8),
                    textcoords="offset points", ha="center", fontsize=8,
                    color=fallout_core.COR_VERDE, **pipeline.ROTULO_SOBRE_LINHA)
    # Base Residencial: a linha verde-escura do slide, sem o Legado Aberto.
    pipeline.desenhar_sem_legado(ax, labels, p.get("plan_sem_legado_hist"),
                                 p["proj_labels"], p.get("proj_sem_legado"), fontsize=8)

    ax.axhline(y=1, color=COR_META, linewidth=1.5, label="Meta (1%)")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    ax.set_ylim(bottom=0)
    # Mesma régua do slide: uma casa decimal quando a escala é curta, senão os
    # rótulos se repetem ("2%, 2%, 2%") em ticks de meio ponto.
    _faixa = ax.get_ylim()[1] - ax.get_ylim()[0]
    _dec = 0 if _faixa >= 5 else 1
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{v:.{_dec}f}%".replace(".", ",")))
    ax.tick_params(labelsize=9)
    # Embaixo do gráfico, numa linha só: com a verde-escura do Base Residencial
    # são cinco itens, e no canto de cima a legenda quebrava em duas linhas por
    # cima dos rótulos dos pontos.
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.1), frameon=False,
              fontsize=8, ncol=5)
    fig.tight_layout()
    return fig


def _aba_projecao():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Projeção · {JORNADA_LABEL}</div>",
        unsafe_allow_html=True,
    )

    _p = carregar_projecao(jornada_escolhida)
    _tem_proj = len(_p["proj_labels"]) > 1

    _base = _p["fallout_pct"]
    _ultimo = _p["proj_vals"][-1] if _tem_proj else _base
    _queda = _base - _ultimo
    _cabecalho = [
        f"Fallout de <b>{_p['mes_atual']}</b>: {_num_br(_base, 2)}%",
        f"{_num_br(_p['total_mes'])} pedidos no mês",
    ]
    if _tem_proj:
        _cabecalho.append(
            f"projeção para <b>{_p['proj_labels'][-1]}</b>: "
            f"<b>{_num_br(_ultimo, 2)}%</b> ({_num_br(_queda, 2)} pp abaixo)")
    _sem_legado = _p.get("proj_sem_legado") or []
    if _sem_legado:
        _cabecalho.append(f"sem Legado Aberto: <b>{_num_br(_sem_legado[-1], 2)}%</b>")
    st.markdown(
        f"<div style='font:600 13px {FONTE};margin:2px 0 14px'>"
        + " · ".join(_cabecalho) + "</div>",
        unsafe_allow_html=True,
    )

    st.pyplot(_fig_projecao(_p), use_container_width=True)
    if _sem_legado:
        st.caption(
            f"**Projeção s/ Legado Aberto** (verde-escura): cada ponto da projeção menos "
            f"{_num_br(_p['proj_vals'][0] - _sem_legado[0], 2)} pp, o Legado Aberto de "
            f"{_p['mes_atual']} (Feito manualmente no Legado + Funcionalidade idêntica ao "
            "legado). Nos meses passados, o valor digitado na coluna PlanejadoSemLegado "
            "do histórico.")

    if not _tem_proj:
        st.info(
            "Sem projeção para esta jornada: não há defeito com data de entrega "
            "futura, nem nada em **Resolvido** ou **Tratado - Em avaliação de "
            "eficácia** para sair da conta. A curva só desce quando existe uma "
            "dessas duas coisas."
        )
    else:
        # ── Como a conta é feita ──────────────────────────────────────────────
        st.markdown(
            f"<div style='font:600 13px {FONTE};margin:18px 0 6px'>"
            f"Como a conta é feita</div>"
            f"<div style='font-size:12px;color:{COR_MUTED_2};margin-bottom:10px'>"
            f"Parte do fallout realizado de {_p['mes_atual']} e desconta o que "
            f"não deve se repetir. Os mesmos números do gráfico do slide.</div>",
            unsafe_allow_html=True,
        )

        _linhas_conta = []
        for _d in _p["detalhe"]:
            _linhas_conta.append(
                f"<div style='display:flex;align-items:baseline;gap:10px;padding:8px 0;"
                f"border-bottom:1px solid {COR_DIVIDER}'>"
                f"<span style='font:600 13px {FONTE};min-width:62px'>{_d['mes']}</span>"
                f"<span style='font-size:12px;color:{COR_MUTED_2};flex:1'>"
                f"{_num_br(_base, 2)}% "
                f"<span style='color:{COR_MUTED}'>(fallout de {_p['mes_atual']})</span>"
                f" − {_num_br(_d['acumulado'], 2)} pp "
                f"<span style='color:{COR_MUTED}'>já corrigido e entregas anteriores</span>"
                f" − {_num_br(_d['entrega_mes'], 2)} pp "
                f"<span style='color:{COR_MUTED}'>entregas do mês (pro-rata)</span>"
                f"</span>"
                f"<span style='font:800 14px {FONTE};color:{COR_ACC_700};"
                f"min-width:66px;text-align:right'>{_num_br(_d['valor'], 2)}%</span>"
                f"</div>"
            )
        st.markdown("".join(_linhas_conta), unsafe_allow_html=True)

        st.markdown(
            f"<div style='font-size:12px;color:{COR_MUTED_2};margin:10px 0 0'>"
            f"Do total, <b>{_num_br(_p['pct_nao_recorrente'], 2)} pp</b> saem já no "
            f"primeiro mês: são os pedidos em <b>Resolvido</b> e <b>Tratado - Em "
            f"avaliação de eficácia</b>, cuja correção já está em produção, mais "
            f"os de <b>Incidente</b>, que são falha pontual e não se repetem.</div>",
            unsafe_allow_html=True,
        )

    # ── Planejamento de Redução ───────────────────────────────────────────────
    if _p["reducao"]:
        _pleno = sum(r["pct_plena"] for r in _p["reducao"])
        st.markdown(
            f"<div style='font:600 13px {FONTE};margin:22px 0 6px'>"
            f"Planejamento de Redução ({_num_br(_pleno, 2)}% pleno)</div>",
            unsafe_allow_html=True,
        )
        _df_red = pd.DataFrame([
            {"Entrega": r["data"],
             "Defeitos / US": ", ".join(r["dfts"]),
             "Pedidos": r["n"],
             "% no mês da entrega": r["pct_mes"],
             "% pleno": r["pct_plena"]}
            for r in _p["reducao"]
        ])
        st.dataframe(
            _df_red, use_container_width=True, hide_index=True,
            column_config={
                "Defeitos / US": st.column_config.TextColumn(width="large"),
                "% no mês da entrega": st.column_config.NumberColumn(format="%.2f%%"),
                "% pleno": st.column_config.NumberColumn(format="%.2f%%"),
            },
        )
        st.caption(
            "**% pleno** é quanto o defeito representa no fallout do mês; "
            "**% no mês da entrega** é essa fatia reduzida pelos dias que sobram "
            "depois da data — um defeito entregue no dia 17 só evita falhas nos "
            "dias restantes. A partir do mês seguinte ele conta cheio."
        )

    st.caption(
        f"A projeção é sempre do mês do relatório (**{_p['mes_atual']}**), não do "
        f"mês escolhido no seletor acima — é o mesmo recorte do slide. Ela assume "
        f"volume constante e nenhum defeito novo, e só enxerga defeito com data de "
        f"entrega marcada: o que está em tratamento sem data continua na conta. "
        f"A curva vai até o mês seguinte à última entrega planejada. Por que um mês "
        f"já em andamento não está batendo com a previsão: aba **Previsto x Realizado**."
    )


_render_aba(aba_projecao, _aba_projecao)


def _aba_ofensores():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Ofensores · {JORNADA_LABEL}</div>",
        unsafe_allow_html=True,
    )

    df_ofe, _meses_ofe, _sucesso_ofe = carregar_ofensores(jornada_escolhida)

    if df_ofe.empty:
        st.info("Nenhum defeito em aberto com ocorrência no mês mais recente "
                "desta jornada.")
    else:
        # A tabela é sempre dos dois meses mais recentes, como no slide — não
        # acompanha o seletor de mês do topo. Sem dizer isso, quem escolhe um
        # mês antigo lá em cima acha que a tabela está errada.
        _periodo = " e ".join(_meses_ofe)
        _suc_txt = " · ".join(f"{lbl} {_num_br(pct, 2)}%"
                              for lbl, pct in _sucesso_ofe.items())
        st.markdown(
            f"<div style='font:600 12px {FONTE};color:{COR_MUTED};margin:0 0 12px'>"
            f"{len(df_ofe)} defeitos/US em aberto · {_periodo} "
            f"(a tabela acompanha o slide, não o mês selecionado acima) · "
            f"sucesso da jornada: {_suc_txt}</div>",
            unsafe_allow_html=True,
        )

        filtro_ofe = st.text_input(
            "Filtrar por DFT, nome do defeito, time ou status",
            key="filtro_ofensores")
        _total_ofe = len(df_ofe)
        if filtro_ofe:
            df_ofe = _filtrar_texto(df_ofe, filtro_ofe)
            # O cabeçalho acima conta a jornada inteira; sem isto a tabela
            # filtrada parece estar contradizendo ele.
            st.caption(f"Mostrando {len(df_ofe)} de {_total_ofe}.")

        # O % fica numérico (e não "0,15%") para a coluna continuar ordenável —
        # é o que serve para achar o maior ofensor.
        _cfg = {f"{lbl} (%)": st.column_config.NumberColumn(format="%.2f%%")
                for lbl in _meses_ofe}
        st.dataframe(df_ofe, use_container_width=True, height=520,
                     hide_index=True, column_config=_cfg)

        # A marca vem do slide, onde não há espaço para legenda; aqui há.
        st.caption("No número do defeito: `*` = correção prevista para o 3º "
                   "trimestre, `**` = 4º trimestre.")

        st.download_button(
            label="⬇ Baixar Excel",
            data=lambda d=df_ofe: to_excel_bytes(d),
            file_name=f"ofensores_{JORNADA_LABEL.replace(' ', '')}_{'-'.join(_meses_ofe)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


_render_aba(aba_ofensores, _aba_ofensores)

# ── Aba "Precisa de Correção" ─────────────────────────────────────────────────
# Três sintomas de que o defeito associado ao pedido não vai resolvê-lo sozinho.
MOT_ENCERRADO = "DFT encerrado"
MOT_POS_MS    = "Criado pós-entrega"
MOT_AMBOS     = "DFT encerrado + criado pós-entrega"
MOT_INEFICAZ  = "Correção não resolveu"      # fallout_core.DFTS_CORRECAO_INEFICAZ
# Um DFT pode ter centenas de pedidos; a lista dentro do expander é HTML puro,
# e despejar tudo trava a renderização. O Excel ao lado sai completo.
LIMITE_PEDIDOS = 40


def _aba_correcao():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Precisa de Correção · {JORNADA_LABEL} · "
        f"{mes_labels[mes_escolhido]}</div>",
        unsafe_allow_html=True,
    )

    if MES_SEM_BRUTO:
        st.info(AVISO_SEM_BRUTO)
    else:
        df_cor = df_mes[df_mes["TemDefeito"]].copy()
        _dia_ped = fallout_core.dia_local(df_cor["CreatedDate"])
        _dia_ms  = fallout_core.dia_local(df_cor["DFT_BugfixMilestone"])

        _fase_cor = df_cor["DFT_Phase"].fillna("").str.strip()
        _cor_enc  = _fase_cor.isin(fallout_core.FASES_ENCERRADAS)
        # A mesma regra que manda esses pedidos para "Em Avaliação por MOPs" na
        # Distribuição: aba e categoria não podem divergir.
        _cor_mops, _cor_pos, _cor_inef = fallout_core.precisa_correcao(df_cor)

        # Motivo é exclusivo: o pós-entrega, mais específico, vence o
        # "Correção não resolveu" — que fica com os pedidos criados até a
        # entrega. O filtro desse motivo, mais abaixo, pega o DFT inteiro.
        df_cor["Motivo"] = MOT_ENCERRADO
        df_cor.loc[_cor_inef & ~_cor_mops, "Motivo"] = MOT_INEFICAZ
        df_cor.loc[_cor_pos, "Motivo"] = MOT_POS_MS
        df_cor.loc[_cor_enc & _cor_pos, "Motivo"] = MOT_AMBOS
        df_cor["_ineficaz"] = _cor_inef
        df_cor["Dias após entrega"] = (
            (_dia_ped - _dia_ms).dt.days.where(_cor_pos).astype("Int64"))
        # Quem entra é `_cor_mops`, e não `_cor_enc`: o DFT "Fechado" teve
        # entrega, então o pedido criado antes dela já foi resolvido por ela e
        # seria ruído aqui — no Cross Sell isso respondia por 398 dos 979 da
        # lista. O "Fechado" volta pelo `_cor_pos`, que é justamente o caso em
        # que a correção não pegou.
        df_cor = df_cor[_cor_mops | _cor_pos | _cor_inef]

        col_mot, col_busca = st.columns([2.4, 3])
        with col_mot:
            # Os filtros incluem quem bate em mais de um critério: esse pedido
            # é caso dos dois lados, não um grupo à parte a ser garimpado.
            filtro_mot = st.radio("Motivo", ["Todos", MOT_ENCERRADO, MOT_POS_MS,
                                             MOT_INEFICAZ],
                                  horizontal=True, key="cor_motivo")
        with col_busca:
            busca_cor = st.text_input(
                "Buscar por pedido, DFT, nome do defeito, time ou erro",
                key="cor_busca")

        if filtro_mot == MOT_ENCERRADO:
            df_cor = df_cor[df_cor["Motivo"].isin([MOT_ENCERRADO, MOT_AMBOS])]
        elif filtro_mot == MOT_POS_MS:
            df_cor = df_cor[df_cor["Motivo"].isin([MOT_POS_MS, MOT_AMBOS])]
        elif filtro_mot == MOT_INEFICAZ:
            df_cor = df_cor[df_cor["_ineficaz"]]

        if busca_cor:
            _mask_cor = pd.Series(False, index=df_cor.index)
            for _c in ("OrderNumber", "DefectNumber_orig", "DFT_Name",
                       "DFT_Team", "ErrorHandled__c"):
                if _c in df_cor.columns:
                    _mask_cor |= df_cor[_c].astype(str).str.contains(
                        busca_cor, case=False, na=False)
            df_cor = df_cor[_mask_cor]

        if df_cor.empty:
            st.info("Nenhum pedido do mês está em DFT encerrado, foi criado "
                    "depois da data de entrega do defeito ou está em DFT cuja "
                    "correção não resolveu.")
        else:
            from html import escape as _esc

            # Agrupa pelo DefectKey e não pelo bloco inteiro de colunas do
            # Octane: se o mesmo defeito vier com o nome grafado de dois jeitos,
            # o agrupamento por colunas o quebraria em dois expanders.
            grupos_cor = sorted(df_cor.groupby("DefectKey", sort=False),
                                key=lambda kv: len(kv[1]), reverse=True)
            _n_ped = len(df_cor)
            _pct_cor = _n_ped / total_mes * 100 if total_mes > 0 else 0
            # As parcelas são exclusivas entre si e somam o total; a de valor
            # zero sai da linha em vez de virar um "0 pedidos".
            _quebra = [
                (int((df_cor["Motivo"] == MOT_ENCERRADO).sum()),
                 "em DFT cancelado ou rejeitado"),
                (int((df_cor["Motivo"] == MOT_POS_MS).sum()),
                 "criados depois da entrega do defeito"),
                (int((df_cor["Motivo"] == MOT_AMBOS).sum()),
                 "nos dois casos"),
                (int((df_cor["Motivo"] == MOT_INEFICAZ).sum()),
                 "em DFT corrigido cuja correção não resolveu"),
            ]
            st.markdown(
                f"<div style='font:600 13px {FONTE};margin:4px 0 4px'>"
                f"{_num_br(_n_ped)} pedidos ({_num_br(_pct_cor, 2)}% do mês) "
                f"em {_num_br(len(grupos_cor))} DFTs</div>"
                f"<div style='font-size:12px;color:{COR_MUTED_2};margin-bottom:12px'>"
                + " · ".join(f"{_num_br(n)} {rot}" for n, rot in _quebra if n)
                + "</div>",
                unsafe_allow_html=True,
            )

            _cols_cor = ["OrderNumber", "CreatedDate", "DefectNumber_orig",
                         "DFT_Name", "DFT_Phase", "DFT_BugfixMilestone",
                         "DFT_Team", "DFT_Type", "Motivo", "Dias após entrega",
                         "ErrorHandled__c", "State", "Channel", "Segment"]
            df_exp_cor = (df_cor[[c for c in _cols_cor if c in df_cor.columns]]
                          .sort_values(["DefectNumber_orig", "CreatedDate"])
                          .rename(columns={"DefectNumber_orig": "Defeito / US"})
                          .reset_index(drop=True))
            st.download_button(
                label="⬇ Baixar Excel",
                data=lambda d=df_exp_cor: to_excel_bytes(d),
                file_name=(f"precisa_correcao_{JORNADA_LABEL.replace(' ', '')}_"
                           f"{mes_labels[mes_escolhido]}.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # `x or ""` não serve: NaN é truthy, e o campo vazio do Octane
            # viraria a string "nan" no cabeçalho do expander.
            _txt_cor = lambda v: "" if pd.isna(v) else str(v).strip()

            for _chave, pedidos in grupos_cor:
                r0    = pedidos.iloc[0]
                orig  = _txt_cor(r0["DefectNumber_orig"]) or _txt_cor(_chave)
                # "DFT " só na frente de número, como na aba de defeitos: as
                # jornadas de catálogo próprio identificam por rótulo (BRJ-218).
                dft_label = (orig if orig.upper().startswith("DFT") or not orig.isdigit()
                             else "DFT " + orig)
                phase = _txt_cor(r0["DFT_Phase"])
                nome  = _txt_cor(r0["DFT_Name"])
                team  = _txt_cor(r0["DFT_Team"])
                ms    = pd.to_datetime(r0["DFT_BugfixMilestone"], errors="coerce")
                ms_s  = ms.strftime("%d/%m/%Y") if not pd.isna(ms) else ""
                n_enc = int((pedidos["Motivo"] != MOT_POS_MS).sum())
                n_pos = int((pedidos["Motivo"] != MOT_ENCERRADO).sum())

                chips_cor = []
                if n_enc:
                    chips_cor.append(f"<span class='chip'>{n_enc} · {MOT_ENCERRADO}</span>")
                if n_pos:
                    chips_cor.append(f"<span class='chip'>{n_pos} · {MOT_POS_MS}</span>")

                linhas_ped = []
                for _, p in pedidos.sort_values("CreatedDate").head(LIMITE_PEDIDOS).iterrows():
                    criado = p["CreatedDate"]
                    dia_s  = (criado.tz_convert("America/Sao_Paulo").strftime("%d/%m")
                              if pd.notna(criado) else "")
                    dias   = p["Dias após entrega"]
                    selo   = (f"<span style='font:600 11px {FONTE};color:{COR_ACC_700}'>"
                              f"+{int(dias)}d</span>") if pd.notna(dias) else ""
                    erro_p = _txt_cor(p.get("ErrorHandled__c"))
                    erro_p = (erro_p[:90] + "…") if len(erro_p) > 90 else erro_p
                    linhas_ped.append(
                        f"<div style='display:flex;align-items:baseline;gap:12px;padding:7px 0;"
                        f"border-bottom:1px solid {COR_DIVIDER}'>"
                        f"<span style='font:600 13px {FONTE};min-width:96px'>"
                        f"{_esc(_txt_cor(p['OrderNumber']))}</span>"
                        f"<span style='font-size:12px;color:{COR_MUTED};min-width:44px'>{dia_s}</span>"
                        f"<span style='min-width:40px'>{selo}</span>"
                        f"<span style='font-size:12px;color:{COR_MUTED_2}'>{_esc(erro_p)}</span>"
                        f"</div>"
                    )
                if len(pedidos) > LIMITE_PEDIDOS:
                    linhas_ped.append(
                        f"<div style='font-size:12px;color:{COR_MUTED};padding:8px 0'>"
                        f"… e mais {_num_br(len(pedidos) - LIMITE_PEDIDOS)} pedidos "
                        f"— a lista completa está no Excel.</div>"
                    )

                partes_cab = [p for p in [phase, ("entrega " + ms_s) if ms_s else "", team] if p]
                cabecalho_cor = (f"{dft_label}  —  {len(pedidos)} pedidos"
                                 + (f" · {' · '.join(partes_cab)}" if partes_cab else ""))
                # Um markdown só por DFT, e <details> nativo em vez de
                # st.expander, pelo mesmo motivo da aba "Erros por Defeito".
                st.markdown(
                    "<details class='erro-det'>"
                    f"<summary>{_esc(cabecalho_cor)}</summary>"
                    "<div class='erro-corpo'>"
                    + (f"<div class='raw-msg'>{_esc(nome)}</div>" if nome else "")
                    + f"<div>{''.join(chips_cor)}</div>"
                    + f"{''.join(linhas_ped)}"
                    "</div></details>",
                    unsafe_allow_html=True,
                )

        st.caption(
            "Entram os pedidos do mês selecionado com defeito associado que "
            "estão em DFT **Cancelado ou Rejeitado** — nenhuma correção vai "
            "chegar, em qualquer data —, e os criados **depois** da Bugfix "
            "Milestone do DFT, onde a correção já era para estar em produção. "
            "`+Nd` é a distância entre a criação do pedido e a data de entrega; "
            "o pedido sem `+Nd` está aqui porque o defeito foi cancelado ou "
            "rejeitado, ou porque a correção dele não resolveu. O pedido criado "
            "**antes** da entrega de um DFT **Fechado** fica de fora: a correção "
            "veio depois da falha, e não há o que corrigir. \"Corrigido\" também "
            "não entra sozinho, pelo mesmo motivo — a exceção é o DFT marcado "
            "como **correção não resolveu**, que entra com todos os pedidos. "
            "DFT com **US de Melhoria em andamento** não entra pela data: ela é "
            "a da US, que ainda não foi entregue. "
            "Todo pedido listado aqui conta como **Em Avaliação por MOPs** na "
            "Distribuição."
        )


_render_aba(aba_correcao, _aba_correcao)


def _aba_erros():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Análise de Erros · {JORNADA_LABEL}</div>",
        unsafe_allow_html=True,
    )

    if MES_SEM_BRUTO:
        st.caption(
            f"O período abaixo cobre só os meses com extração. "
            f"{mes_labels[mes_escolhido]} é um mês fechado e não tem mais "
            f"detalhe pedido a pedido."
        )

    # Usa todo o histórico (não só o mês selecionado); o período é escolhido abaixo
    df_err = df.copy()
    df_err["ErrorHandled__c"] = df_err["ErrorHandled__c"].fillna("(sem mensagem de erro)")
    _orig_lower = df_err["DefectNumber_orig"].astype(str).str.strip().str.lower()
    df_err["_outros_times"] = _orig_lower == "enviado e-mail - outros times"
    df_err["_pontual"]      = df_err["DefectNumber__c"] == 999999
    # Defeito com número do Octane ou identificado por rótulo (BRJ-*, INC-*,
    # PDST-*, "Paliativo RPA"…), o mesmo critério da coluna DFTs. Só pelo
    # número, o "Com DFT" do Base Residencial não trazia nenhum dos 40.061
    # pedidos com BRJ-*.
    df_err["tem_dft"] = (
        (df_err["DefectNumber__c"].notna() &
         (df_err["DefectNumber__c"] != -1) &
         (df_err["DefectNumber__c"] != 999999)) |
        df_err["TemDefeito"].fillna(False).astype(bool)
    ) & ~df_err["_outros_times"]
    df_err["Data"] = df_err["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.date

    # ── Filtro de datas (todo o histórico disponível) ──────────────────────────
    datas_disp = sorted(df_err["Data"].dropna().unique())
    if datas_disp:
        # padrão = mês selecionado no topo, mas dá para ampliar até o histórico todo
        # `mes_escolhido` é a chave AAAAMM; comparar com `d.month` (1..12) nunca
        # casava, e a aba abria com o histórico inteiro em vez do mês do topo.
        _dm = [d for d in datas_disp
               if fallout_core.chave_mes(d.year, d.month) == mes_escolhido]
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
            agrupar = st.selectbox("Colunas por", ["Dia", "Mês"], key="erros_agrupar")
    else:
        col_f1, col_f2 = st.columns([3, 1.5])
        dt_ini, dt_fim, agrupar = None, None, "Dia"

    # ── Filtros de texto e DFT ─────────────────────────────────────────────────
    with col_f1:
        busca = st.text_input("Buscar na mensagem de erro", key="erros_busca")
    with col_f2:
        filtro_dft = st.selectbox(
            "Filtrar por DFT",
            ["Todos", "Com DFT", "Sem DFT", "Falha Pontual", "Outros Times"],
            key="erros_filtro_dft",
        )

    if dt_ini and dt_fim:
        df_err = df_err[(df_err["Data"] >= dt_ini) & (df_err["Data"] <= dt_fim)]

    # Denominador do "% do Total": soma dos meses cobertos pelo período escolhido.
    # A chave do resumo é AAAAMM — com `d.month` (1..12) nenhum mês casava, a
    # soma dava zero, o denominador caía no fallback 1 e a coluna "% do Total"
    # saía multiplicada por 100 (uma ocorrência aparecia como 100%).
    _meses_sel = sorted({fallout_core.chave_mes(d.year, d.month)
                         for d in df_err["Data"].dropna().unique()})
    # Mês fechado que veio da planilha de acompanhamento tem taxa mas não
    # volume: o Total é NaN e entraria na soma como tal.
    _totais_sel = [resumo.loc[m, "Total"] for m in _meses_sel if m in resumo.index]
    _total_per = int(sum(t for t in _totais_sel if pd.notna(t))) or 1

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
    pivot_erros = pivot_erros.merge(_dfts_por_erro(df_err), on="ErrorHandled__c", how="left")
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
    # Na tela, o dia sai como "dd/mm" (com o ano só se o período cruza o ano),
    # numa coluna estreita: a grade do Streamlit não deixa o cabeçalho na
    # vertical, e com "dd/mm/aaaa" cada coluna de dia ocupava o dobro. O Excel
    # continua com a data inteira.
    if agrupar == "Dia" and day_cols_raw:
        _um_ano = day_cols_raw[0].year == day_cols_raw[-1].year
        rotulo_tela = {rename_map[c]: c.strftime("%d/%m" if _um_ano else "%d/%m/%y")
                       for c in day_cols_raw}
    else:
        rotulo_tela = {d: d for d in day_col_strs}
    dias_tela = [rotulo_tela[d] for d in day_col_strs]
    col_config = {
        # Fixa à esquerda: com uma coluna por dia a tabela rola para o lado, e
        # a mensagem sumia de vista junto com a referência dos números.
        "Mensagem de Erro": st.column_config.TextColumn(width="large", pinned=True),
        "Total Geral":      st.column_config.NumberColumn(width="small"),
        "% do Total":       st.column_config.NumberColumn(format="%.2f%%", width="small"),
        "DFTs":             st.column_config.TextColumn(width="large"),
        **{d: st.column_config.NumberColumn(width=46 if agrupar == "Dia" else "small",
                                            alignment="center")
           for d in dias_tela},
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
    # `size` antes do max: um filtro sem resultado deixa a tabela vazia, e o max
    # de um array vazio levanta exceção.
    _vmax     = int(_day_vals.max()) if _day_vals.size and _day_vals.max() > 0 else 1

    # A cor depende só do valor, e a tabela repete poucos valores em milhares de
    # células: cada um passa pelo colormap uma vez.
    @functools.lru_cache(maxsize=None)
    def _color_cell(val):
        """Fundo da célula e tinta legível sobre ele (branca nos tons escuros)."""
        if val == 0 or pd.isna(val):
            return ""
        norm = (val - _vmin) / (_vmax - _vmin) if _vmax > _vmin else 1.0
        r, g, b, _a = _cmap(norm)
        # Cor explícita nos dois casos: a grade do Streamlit não entende "inherit"
        # e deixava o número invisível nos tons claros.
        tinta = "#ffffff" if 0.299 * r + 0.587 * g + 0.114 * b < 0.5 else "#201e1d"
        return f"background-color: {_mc.to_hex((r, g, b))}; color: {tinta}"

    tela = pivot_erros[col_order].rename(columns=rotulo_tela)
    styled = tela.style.map(_color_cell, subset=dias_tela)
    # Cabe ~21 linhas (eram 13, com 500 px); com poucas mensagens encolhe até o
    # conteúdo, em vez de sobrar faixa vazia. A rolagem continua dentro da
    # tabela, para o cabeçalho e a coluna fixa seguirem à vista.
    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        height=min(38 + 35 * len(pivot_erros), 800),
        column_config=col_config,
    )

    st.download_button(
        label="⬇ Baixar Excel",
        data=lambda d=pivot_erros[col_order]: to_excel_bytes(d),
        file_name=(f"erros_{jornada_escolhida}_"
                   f"{dt_ini.strftime('%Y%m%d')}-{dt_fim.strftime('%Y%m%d')}.xlsx"
                   if dt_ini else f"erros_{jornada_escolhida}.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


_render_aba(aba_erros, _aba_erros)


def _aba_defeitos():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Erros por Defeito · {JORNADA_LABEL} · "
        f"{mes_labels[mes_escolhido]}</div>",
        unsafe_allow_html=True,
    )

    if MES_SEM_BRUTO:
        st.info(AVISO_SEM_BRUTO)
    else:
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
            busca_def = st.text_input("Buscar por erro ou defeito", key="def_busca")
        with col_f2:
            min_occ = st.number_input("Mín. ocorrências por erro", min_value=1, step=1,
                                      key="def_min_occ")

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

        # Um groupby só, em vez de filtrar `grp_def` inteiro a cada erro. Cada
        # grupo mantém as linhas na ordem original, como o filtro fazia.
        grupos_def = dict(tuple(grp_def.groupby("ErrorHandled__c", sort=False)))
        for erro in erros_unicos:
            dfts       = grupos_def[erro].sort_values("Qtd", ascending=False)
            total_erro = int(dfts["Qtd"].sum())
            pct_erro   = total_erro / total_mes * 100 if total_mes > 0 else 0
            n_dfts     = len(dfts)

            titulo = (erro[:110] + "…") if len(erro) > 110 else erro
            sufixo = f" · {n_dfts} DFTs" if n_dfts > 1 else ""
            chips, linhas_dft = [], []
            for _, r in dfts.iterrows():
                # `x or ""` não serve aqui: NaN é truthy, então o campo vazio
                # do Octane virava a string "nan" e a linha saía "nan · nan · nan".
                _txt  = lambda v: "" if pd.isna(v) else str(v).strip()
                orig  = _txt(r["DefectNumber_orig"])
                qtd   = int(r["Qtd"])
                nome  = _txt(r["DFT_Name"])
                phase = _txt(r["DFT_Phase"])
                team  = _txt(r["DFT_Team"])
                ms    = pd.to_datetime(r["DFT_BugfixMilestone"], errors="coerce")
                ms_s  = ms.strftime("%d/%m/%Y") if not pd.isna(ms) else ""

                if orig in ("", "nan", "-1"):
                    dft_label = "(sem DFT)"
                    info      = "sem defeito associado"
                elif orig == "999999":
                    dft_label = "Pontual"
                    info      = ""
                else:
                    # "DFT " só na frente de número. As jornadas com catálogo
                    # próprio identificam o defeito por rótulo ("Erro de
                    # Processo"), e prefixar virava "DFT Erro de Processo".
                    # Alguns registros já vêm com o prefixo ("DFT 232143").
                    if orig.upper().startswith("DFT") or not orig.isdigit():
                        dft_label = orig
                    else:
                        dft_label = "DFT " + orig
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

            # Um markdown só por erro: são dezenas de blocos, e cada chamada
            # do Streamlit cria mais um nível de wrapper no DOM.
            cabecalho = (f"{titulo}  —  {total_erro} ocorrências "
                         f"({_num_br(pct_erro, 2)}%){sufixo}")
            st.markdown(
                "<details class='erro-det'>"
                f"<summary>{_esc(cabecalho)}</summary>"
                "<div class='erro-corpo'>"
                f"<div class='raw-msg'>{_esc(str(erro))}</div>"
                f"<div>{''.join(chips)}</div>"
                f"{''.join(linhas_dft)}"
                "</div></details>",
                unsafe_allow_html=True,
            )


_render_aba(aba_defeitos, _aba_defeitos)


# ── Aba "Comparação" ──────────────────────────────────────────────────────────
# Compara duas fotografias gravadas (ver fotografia.py e comparacao.py). A
# gravação é manual, pelo botão "Gravar fotografia" do topo; esta aba só lê o
# que já foi gravado — nada aqui recalcula a extração.

def _titulo_secao(texto, margem="18px 0 6px"):
    st.markdown(f"<div style='font:600 13px {FONTE};margin:{margem}'>{texto}</div>",
                unsafe_allow_html=True)


# As tabelas da aba são HTML, e não st.dataframe: a grade do Streamlit não tem
# etiqueta, barra nem texto em duas linhas, e é isso que deixa a tabela lida de
# relance. A tabela completa, ordenável, fica atrás de uma chave em cada seção.
# HTML escapado, e não markdown: as mensagens de erro trazem "|", colchetes e
# "$", que o markdown leria como tabela, link e fórmula.
def _esc(texto):
    from html import escape
    return escape(str(texto)).replace("$", "&#36;")


def _tabela_html(colunas, linhas):
    """
    `colunas`: (título, largura em %, alinhada à direita?). `linhas`: listas
    de células já em HTML, ou (classe da linha, células).
    """
    cab = "".join(
        f"<th class='num' style='width:{larg}%'>{_esc(tit)}</th>" if dir_
        else f"<th style='width:{larg}%'>{_esc(tit)}</th>"
        for tit, larg, dir_ in colunas)
    corpo = []
    for linha in linhas:
        classe, celulas = linha if isinstance(linha, tuple) else ("", linha)
        tds = "".join(f"<td class='num'>{cel}</td>" if dir_ else f"<td>{cel}</td>"
                      for cel, (_t, _l, dir_) in zip(celulas, colunas))
        corpo.append((f"<tr class='{classe}'>" if classe else "<tr>") + tds + "</tr>")
    return (f"<table class='cmp-tab'><thead><tr>{cab}</tr></thead>"
            f"<tbody>{''.join(corpo)}</tbody></table>")


def _pp_html(valor, maior=None):
    """
    Variação em pp: vermelha quando o fallout sobe, verde quando cai, com uma
    barra do tamanho relativo a `maior` (a maior variação da tabela).
    """
    if valor is None or valor != valor:
        return "<span class='fraco'>—</span>"
    if abs(valor) < 0.005:
        return "<span class='fraco'>0,00 pp</span>"
    tom = "pior" if valor > 0 else "melhor"
    texto = ("+" if valor > 0 else "−") + _num_br(abs(valor), 2) + " pp"
    barra = (f"<span class='barra {tom}' style='width:{max(2, round(abs(valor) / maior * 56))}px'>"
             f"</span>" if maior else "")
    return f"<span class='{tom}'>{texto}</span>{barra}"


def _pedidos_html(agora, variacao):
    """'272 (+24)': o total e, em cor, quanto mudou."""
    if not variacao:
        return _num_br(agora)
    tom = "pior" if variacao > 0 else "melhor"
    return (f"{_num_br(agora)} <span class='{tom}'>"
            f"({'+' if variacao > 0 else '−'}{_num_br(abs(variacao))})</span>")


def _sinal_html(sinal, tom):
    return f"<span class='sinal {tom}'>{_esc(sinal)}</span>"


def _partes_do_texto(texto):
    """(título, tópicos, ressalva) do texto da comparação (ver `comparacao._textos`)."""
    if not texto:
        return "", [], ""
    titulo, resto = texto[0], texto[1:]
    ressalva = next((l for l in resto if l.startswith("Ressalva: ")), "")
    return titulo, [l for l in resto if l != ressalva], ressalva


def _resumo_html(texto):
    """O texto da comparação: o título em destaque, os tópicos e a ressalva apagada."""
    titulo, topicos, ressalva = _partes_do_texto(texto)
    itens = "".join(
        f"<li><b>A fazer:</b> {_esc(l[len('A fazer: '):])}</li>" if l.startswith("A fazer: ")
        else f"<li>{_esc(l)}</li>"
        for l in topicos)
    return (f"<div class='cmp-resumo'><div class='titulo'>{_esc(titulo)}</div>"
            f"<ul>{itens}</ul>"
            + (f"<div class='ressalva'>{_esc(ressalva)}</div>" if ressalva else "")
            + "</div>")


def _nota(html):
    st.markdown(f"<div class='cmp-nota'>{html}</div>", unsafe_allow_html=True)


def _fig_comparacao(curvas):
    """As duas projeções no mesmo gráfico: a anterior em cinza, a atual com as cores do slide."""
    import comparacao
    import matplotlib.ticker as mticker
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure

    antes, agora = curvas["antes"], curvas["agora"]
    rotulos = sorted({r for curva in (antes, agora)
                      for r in curva["labels_hist"] + curva["proj_labels"]},
                     key=comparacao._chave_do_rotulo)
    posicao = {r: i for i, r in enumerate(rotulos)}

    fig = Figure(figsize=(11, 3.8))
    FigureCanvasAgg(fig)
    fig.patch.set_facecolor("white")
    ax = fig.add_subplot(111)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_facecolor("white")

    def _pontos(rotulos_curva, valores):
        pares = [(posicao[r], v) for r, v in zip(rotulos_curva, valores)
                 if v is not None and v == v]
        return [x for x, _ in pares], [y for _, y in pares]

    maior = 1.0
    cinza = "#9E9E9E"
    # A anterior primeiro, para a atual ficar por cima onde as duas coincidem
    # (os meses fechados, que vêm do histórico, são iguais nas duas).
    # O rótulo da anterior vai à direita e abaixo do ponto, e o da atual acima:
    # centrado embaixo, ele caía sobre a linha da meta quando a projeção chega
    # perto de 1%.
    for curva, cor_real, cor_proj, espessura, nome, deslocamento in (
            (antes, cinza, cinza, 1.2, "anterior", (16, -12)),
            (agora, "black", fallout_core.COR_VERDE, 1.8, "atual", (0, 8))):
        xs, ys = _pontos(curva["labels_hist"], curva["vals_hist"])
        ax.plot(xs, ys, color=cor_real, linewidth=espessura, marker="o", markersize=4,
                label=f"Realizado ({nome})")
        maior = max([maior] + ys)
        xs, ys = _pontos(curva["proj_labels"], curva["proj_vals"])
        if len(xs) > 1:
            ax.plot(xs, ys, color=cor_proj, linewidth=espessura, linestyle="--", marker="o",
                    markersize=4, label=f"Projeção ({nome})")
            for x, y in zip(xs[1:], ys[1:]):
                ax.annotate(f"{_num_br(y, 2)}%", xy=(x, y), xytext=deslocamento,
                            textcoords="offset points", ha="center", fontsize=8,
                            color=cor_proj)
            maior = max([maior] + ys)

    ax.axhline(y=1, color=COR_META, linewidth=1.2, label="Meta (1%)")
    ax.set_xticks(range(len(rotulos)))
    ax.set_xticklabels(rotulos, fontsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    # Folga no alto para a legenda não cobrir os pontos.
    ax.set_ylim(0, maior * 1.3)
    _faixa = ax.get_ylim()[1]
    _dec = 0 if _faixa >= 5 else 1
    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{v:.{_dec}f}%".replace(".", ",")))
    ax.tick_params(labelsize=9)
    ax.legend(loc="upper right", frameon=False, fontsize=8, ncol=3)
    fig.tight_layout()
    return fig


def _detalhe_erros(c):
    """As três tabelas completas dos erros: novos, em alta e que sumiram."""
    import comparacao

    formatos = {
        "Desde": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Última ocorrência": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Peso no mês (pp)": st.column_config.NumberColumn(format="%.3f"),
        "Sem defeito associado (%)": st.column_config.NumberColumn(format="%.0f%%"),
        "Média diária antes": st.column_config.NumberColumn(format="%.1f"),
        "Média diária agora": st.column_config.NumberColumn(format="%.1f"),
        "Vezes": st.column_config.NumberColumn(format="%.1f×"),
        "Acima do esperado": st.column_config.NumberColumn(format="%.0f"),
        "Erro": st.column_config.TextColumn(width="large"),
        "Mensagem de erro": st.column_config.TextColumn("Mensagem completa", width="large"),
    }
    for titulo, tab_erros in (("Erros novos", c.erros_novos),
                              ("Erros em alta", c.erros_em_alta),
                              ("Erros que sumiram", c.erros_sumidos)):
        _titulo_secao(f"{titulo} ({len(tab_erros)})")
        if not len(tab_erros):
            st.caption("Nenhum.")
            continue
        mostrar = tab_erros.copy()
        mostrar.insert(0, "Erro", mostrar["Mensagem de erro"].map(comparacao.resumir_mensagem))
        colunas = ["Erro"] + [col for col in mostrar.columns
                              if col not in ("Erro", "Mensagem de erro")] + ["Mensagem de erro"]
        st.dataframe(mostrar[colunas], hide_index=True, use_container_width=True,
                     height=min(38 + 35 * len(mostrar), 420), column_config=formatos)


def _aba_comparacao():
    import comparacao

    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Comparação · {JORNADA_LABEL}</div>",
        unsafe_allow_html=True,
    )

    fotos = fotografia.listar(_data_dir)
    if len(fotos) < 2:
        st.info("A comparação precisa de duas fotografias, e "
                + ("ainda não há nenhuma gravada." if not fotos else "só há uma gravada.")
                + " Grave uma a cada extração nova, no botão **Gravar fotografia**, no topo.")
        return

    por_id = {f["id"]: f for f in fotos}
    ids = list(por_id)

    def _rotulo_foto(id_foto):
        return f"Dados até {por_id[id_foto]['corte']:%d/%m/%Y}"

    # Os padrões vão para o estado antes de o widget ser desenhado, e não por
    # `index=`/`value=`: com o valor preservado entre abas (_FILTROS_DAS_ABAS),
    # os dois juntos fazem o Streamlit reclamar. A atual padrão é a fotografia
    # da extração que está na tela; sem ela, a mais recente.
    corte_tela = fallout_core.data_corte(_data_dir)
    da_tela = corte_tela.isoformat() if corte_tela else None
    if st.session_state.get("cmp_atual") not in ids:
        st.session_state["cmp_atual"] = da_tela if da_tela in ids else ids[0]
    st.session_state.setdefault("cmp_min_occ", comparacao.Limites.min_ocorrencias)
    st.session_state.setdefault("cmp_fator", comparacao.Limites.fator_alta)

    col_atual, col_base, col_occ, col_fator = st.columns([1.6, 1.6, 1.5, 1.5])
    with col_atual:
        atual_id = st.selectbox("Fotografia atual", ids, format_func=_rotulo_foto,
                                key="cmp_atual")
    anteriores = [i for i in ids if por_id[i]["corte"] < por_id[atual_id]["corte"]]
    if not anteriores:
        st.info(f"Não há fotografia anterior à de {por_id[atual_id]['corte']:%d/%m/%Y} "
                f"para comparar.")
        return
    if st.session_state.get("cmp_base") not in anteriores:
        st.session_state["cmp_base"] = comparacao.escolher_base(fotos, atual_id)
    with col_base:
        base_id = st.selectbox("Comparar com", anteriores, format_func=_rotulo_foto,
                               key="cmp_base")
    # Os limites ficam à vista, e não num st.expander: nesta máquina o expander
    # não abre (ver o CSS de "Erros por Defeito") — a seta gira e nada aparece.
    with col_occ:
        min_occ = st.number_input("Erro novo: mín. ocorrências", min_value=1, step=1,
                                  key="cmp_min_occ")
    with col_fator:
        fator = st.number_input("Em alta: vezes a média", min_value=1.0, step=0.5,
                                key="cmp_fator")

    if da_tela and da_tela not in ids:
        st.caption(f"A extração na tela (dados até {corte_tela:%d/%m/%Y}) ainda não tem "
                   "fotografia: grave-a em **Gravar fotografia**, no topo, para ela entrar "
                   "na comparação.")

    chave = (base_id, por_id[base_id]["gravada_em"], atual_id, por_id[atual_id]["gravada_em"])
    limites = (int(min_occ), float(fator))
    try:
        # O mês é o do seletor do topo, como na Distribuição. Se as duas
        # fotografias não o têm, a comparação volta para o mês do relatório e
        # avisa.
        c = carregar_comparacao(*chave, jornada_escolhida, int(mes_escolhido), *limites)
    except ValueError as e:
        st.info(str(e))
        return
    ctx = c.contexto

    # ── Cartões ───────────────────────────────────────────────────────────────
    info_mes = ctx["mes"]
    mes_rot = fallout_core.MESES_LABEL[c.mes].capitalize() if c.mes else "—"
    cartoes = []
    def _variacao(dif):
        """Texto da seta do cartão; sem seta quando não mudou (a seta de 0,00 engana)."""
        return "" if dif is None or abs(dif) < 0.005 else _num_br(abs(dif), 2) + "pp"

    # Relatório que virou de mês: primeiro o fallout que cada fotografia
    # reportava (set-26 3,24% → out-26 2,53%); o cartão seguinte é o do mês que
    # as duas têm, onde está a explicação da variação.
    rel = ctx.get("relatorio", {})
    if rel.get("virou") and rel["taxa_agora"] is not None:
        dif = None if rel["taxa_antes"] is None else rel["taxa_agora"] - rel["taxa_antes"]
        mes_ant = fallout_core.MESES_LABEL[rel["mes_antes"]]
        cartoes.append(_kpi_tile(
            f"Fallout {fallout_core.MESES_LABEL[rel['mes_agora']].capitalize()}",
            _num_br(rel["taxa_agora"], 2) + "%", _variacao(dif), (dif or 0) > 0, [],
            destaque=(dif or 0) >= 0.005,
            rodape=f"antes: {mes_ant} " + ("—" if rel["taxa_antes"] is None
                                         else _num_br(rel["taxa_antes"], 2) + "%")
                   + f" · {rel['dias_agora']} dias de dados"))
    ta, tb = info_mes.get("taxa_antes"), info_mes.get("taxa_agora")
    if tb is not None:
        dif = None if ta is None else tb - ta
        cartoes.append(_kpi_tile(
            f"Fallout {mes_rot}", _num_br(tb, 2) + "%", _variacao(dif), (dif or 0) > 0, [],
            destaque=(dif or 0) >= 0.005,
            rodape="antes: " + ("—" if ta is None else _num_br(ta, 2) + "%")))
    proj_meses = c.projecao["meses"]
    if len(proj_meses):
        p = proj_meses.iloc[0]
        dif = p["Variação (pp)"]
        cartoes.append(_kpi_tile(
            f"Projeção {p['Mês'].capitalize()}", _num_br(p["Agora"], 2) + "%",
            _variacao(dif), dif > 0, [], destaque=dif >= 0.005,
            rodape=f"antes: {_num_br(p['Antes'], 2)}%"))
    n_novos, n_alta = len(c.erros_novos), len(c.erros_em_alta)
    n_pos = (int((c.defeitos["Pós-entrega na semana"] >= c.limites.min_pos_entrega).sum())
             if len(c.defeitos) else 0)
    cartoes.append(_kpi_tile("Erros novos", _num_br(n_novos), "", False, [],
                             destaque=n_novos > 0,
                             rodape=f"com {c.limites.min_ocorrencias}+ ocorrências"))
    cartoes.append(_kpi_tile("Erros em alta", _num_br(n_alta), "", False, [],
                             destaque=n_alta > 0,
                             rodape=f"{_num_br(c.limites.fator_alta, 1)}× a média anterior"))
    cartoes.append(_kpi_tile("Pedidos após a entrega", _num_br(n_pos), "", False, [],
                             destaque=n_pos > 0, rodape="defeitos com pedido novo depois da entrega"))
    st.markdown("<div class='kpi-row'>" + "".join(cartoes) + "</div>", unsafe_allow_html=True)

    # ── O que mudou ───────────────────────────────────────────────────────────
    _titulo_secao(f"O que mudou · dados até {ctx['corte_antes']:%d/%m} → "
                  f"dados até {ctx['corte_agora']:%d/%m}", "22px 0 4px")
    st.markdown(_resumo_html(c.texto), unsafe_allow_html=True)

    # O texto para e-mail vai numa caixa de texto, aberta por um toggle: o
    # st.expander não abre nesta máquina, e o botão de copiar do st.code depende
    # da área de transferência do navegador, que a janela própria do app pode
    # bloquear. Na caixa, clicar dentro, Ctrl+A e Ctrl+C funciona em qualquer uma.
    col_texto, col_excel, _col_vazia = st.columns([2.2, 2.6, 4])
    with col_texto:
        mostrar_texto = st.toggle("Mostrar texto para e-mail", key="cmp_mostrar_texto")
    with col_excel:
        st.download_button(
            label="⬇ Baixar Excel da comparação",
            data=lambda c=c: comparacao.planilha(c),
            file_name=(f"comparacao_{JORNADA_LABEL.replace(' ', '')}_{base_id}_a_{atual_id}.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if mostrar_texto:
        titulo, topicos, ressalva = _partes_do_texto(c.texto)
        texto_email = (titulo + "\n\n" + "\n".join(f"• {l}" for l in topicos)
                       + (f"\n\n{ressalva}" if ressalva else ""))
        st.text_area("Clique dentro, Ctrl+A e Ctrl+C para copiar", texto_email,
                     height=min(110 + 38 * len(c.texto), 480))
    falhas = [v for v in c.verificacoes if not v["OK"]]
    if falhas:
        st.warning("Algumas conferências não fecharam — os números abaixo podem não somar:\n\n"
                   + "\n".join(f"- {v['Verificação']}: {v['Detalhe']}" for v in falhas))
    else:
        st.caption(f"Conferências: {len(c.verificacoes)} de {len(c.verificacoes)} certas — "
                   "em cada variação, as partes somam exatamente o total.")

    t_dist, t_proj, t_def, t_err, t_dados = st.tabs(
        ["Distribuição", "Projeção", "Defeitos", "Erros", "Qualidade dos dados"])

    # ── Distribuição ──────────────────────────────────────────────────────────
    with t_dist:
        # Relatório que virou de mês: a Distribuição que cada fotografia
        # reportava, a mesma da aba Distribuição. A tabela seguinte é a de um
        # mês que as duas têm, e por isso pode ser de outro mês que a da aba.
        if rel.get("virou") and len(rel["distribuicao"]):
            rot_a = fallout_core.MESES_LABEL[rel["mes_antes"]]
            rot_b = fallout_core.MESES_LABEL[rel["mes_agora"]]
            _titulo_secao(f"Distribuição do relatório · {rot_a} → {rot_b}", "10px 0 6px")
            dr = rel["distribuicao"]
            maior = dr["Diferença (pp)"].abs().max()
            st.markdown(_tabela_html(
                [("Categoria", 44, False), (f"{rot_a} (anterior)", 16, True),
                 (f"{rot_b} (atual)", 16, True), ("Diferença", 24, True)],
                [["&nbsp;" * 6 * int(l.get("Nível", int(l["Subgrupo"])))
                  + _esc(fallout_core.rotulo_dist(l["Categoria"],
                                                  l.get("Nível", int(l["Subgrupo"]))))]
                 + (["", "", ""] if fallout_core.so_rotulo(
                     l["Categoria"], l.get("Nível", int(l["Subgrupo"]))) else
                    [_num_br(l["% antes"], 2) + "%", _num_br(l["% agora"], 2) + "%",
                     _pp_html(l["Diferença (pp)"], maior)])
                 for _, l in dr.iterrows()]
                + [("total", ["Fallout do relatório",
                              "—" if rel["taxa_antes"] is None else _num_br(rel["taxa_antes"], 2) + "%",
                              _num_br(rel["taxa_agora"], 2) + "%",
                              _pp_html(None if rel["taxa_antes"] is None
                                       else rel["taxa_agora"] - rel["taxa_antes"])])]),
                unsafe_allow_html=True)
            _nota(f"Meses diferentes: {rot_b} tem {rel['dias_agora']} dias de dados e "
                  "outros pedidos que " + rot_a + ", então não há como dizer o que "
                  "migrou — a explicação por pedido fica na tabela abaixo, de "
                  + (fallout_core.MESES_LABEL[c.mes] if c.mes else "—") + ".")
            if c.mes is not None:
                _titulo_secao(f"Mesmo mês nas duas fotografias · "
                              f"{fallout_core.MESES_LABEL[c.mes]}", "18px 0 6px")
        if c.mes is None:
            st.info("As duas fotografias não têm nenhum mês em comum com volume e pedidos "
                    "desta jornada: a distribuição não é comparável.")
        else:
            dist = c.distribuicao
            tem_pp = "Variação (pp)" in dist.columns and len(dist)
            if tem_pp:
                # Da maior variação para a menor; a que não chega a 0,01 pp
                # sai da vista (continua na conta completa).
                mudou = dist[dist["Variação (pp)"].abs() >= 0.01]
                mudou = mudou.reindex(mudou["Variação (pp)"].abs()
                                      .sort_values(ascending=False).index)
                maior = mudou["Variação (pp)"].abs().max() if len(mudou) else None
                linhas = [[_esc(l["Categoria"]), _num_br(l["% agora"], 2) + "%",
                           _pp_html(l["Variação (pp)"], maior),
                           f"<span class='fraco'>{_esc(l['O que explica'])}</span>"]
                          for _, l in mudou.iterrows()]
                if tb is not None:
                    linhas.append(("total", ["Fallout do mês", _num_br(tb, 2) + "%",
                                             _pp_html(None if ta is None else tb - ta), ""]))
                st.markdown(_tabela_html(
                    [("Categoria", 29, False), ("% agora", 10, True),
                     ("Variação", 17, True), ("O que explica", 44, False)], linhas),
                    unsafe_allow_html=True)
                ocultas = len(dist) - len(mudou)
                partes = []
                if info_mes.get("total_antes") and info_mes.get("total_agora"):
                    partes.append(f"% sobre todos os pedidos do mês, sucesso incluído "
                                  f"({_num_br(info_mes['total_antes'])} → "
                                  f"{_num_br(info_mes['total_agora'])})")
                if ocultas:
                    partes.append(f"{ocultas} {'categoria' if ocultas == 1 else 'categorias'} "
                                  f"com variação abaixo de 0,01 pp "
                                  f"{'fica' if ocultas == 1 else 'ficam'} de fora")
                _nota("".join(_esc(p) + ". " for p in partes)
                      + "<span class='pior'>Vermelho</span> aumenta o fallout; "
                        "<span class='melhor'>verde</span> reduz.")

            st.session_state.setdefault("cmp_dist_completa", not tem_pp)
            if st.toggle("Mostrar a conta completa", key="cmp_dist_completa"):
                colunas = ["Categoria", "% antes", "% agora", "Variação (pp)",
                           "Período novo (pp)", "Reclassificação (pp)", "Removidos (pp)",
                           "Pedidos antes", "Pedidos agora", "Novos", "Entraram", "Saíram",
                           "Removidos"]
                st.dataframe(
                    dist[[col for col in colunas if col in dist.columns]],
                    hide_index=True, use_container_width=True,
                    column_config={
                        "Categoria": st.column_config.TextColumn(width="large"),
                        "% antes": st.column_config.NumberColumn(format="%.2f%%"),
                        "% agora": st.column_config.NumberColumn(format="%.2f%%"),
                        **{col: st.column_config.NumberColumn(format="%+.2f")
                           for col in ("Variação (pp)", "Período novo (pp)",
                                       "Reclassificação (pp)", "Removidos (pp)")},
                    },
                )
                st.caption(
                    "A variação de cada categoria se divide em três partes que somam "
                    "exatamente o total. **Período novo**: os pedidos dos dias novos, já "
                    "descontado o crescimento do mês — positivo quando os dias novos tiveram, "
                    "proporcionalmente, mais pedidos da categoria que o período que já "
                    "existia. **Reclassificação**: pedidos que já existiam e mudaram de "
                    "categoria. **Removidos**: pedidos que a extração nova não traz mais. No "
                    "Excel, o período novo aparece separado em pedidos novos e efeito do volume.")

    # ── Projeção ──────────────────────────────────────────────────────────────
    with t_proj:
        st.pyplot(_fig_comparacao(c.projecao["curvas"]), use_container_width=True)
        for nota in c.projecao["notas"]:
            st.caption(nota)
        if len(proj_meses):
            st.dataframe(
                proj_meses, hide_index=True, use_container_width=True,
                column_config={
                    "Antes": st.column_config.NumberColumn(format="%.2f%%"),
                    "Agora": st.column_config.NumberColumn(format="%.2f%%"),
                    **{col: st.column_config.NumberColumn(format="%+.2f")
                       for col in proj_meses.columns if col.endswith("(pp)")},
                })
            _titulo_secao("O que mudou em cada mês projetado")
            st.dataframe(
                c.projecao["itens"], hide_index=True, use_container_width=True,
                column_config={
                    "Detalhe": st.column_config.TextColumn(width="large"),
                    "Efeito (pp)": st.column_config.NumberColumn(format="%+.3f"),
                })
            st.caption(
                "**Efeito** é quanto cada item mudou a projeção do mês — positivo quando ela "
                "subiu. **Ponto de partida** é a variação do fallout do mês do relatório; "
                "**Já corrigido** é o que sai da conta no primeiro mês projetado (Resolvido, "
                "Tratado e Incidente); cada defeito do Planejamento de Redução entra pelo peso "
                "que tira do mês. A soma dos efeitos é a variação.")

    # ── Defeitos ──────────────────────────────────────────────────────────────
    with t_def:
        defeitos = c.defeitos
        if not len(defeitos):
            st.info("Nenhum defeito com pedido no mês comparado.")
        else:
            # Do sinal mais grave ao mais leve; dentro dele, do maior volume.
            # O sinal bom (saiu dos ofensores) não pede ação: vai numa linha só.
            ordem = list(comparacao.SINAIS_DEFEITO)
            acao = defeitos[defeitos["Sinal"].map(comparacao.SINAIS_DEFEITO).isin(["pior", "atencao"])]
            acao = (acao.assign(_ordem=acao["Sinal"].map(ordem.index))
                    .sort_values(["_ordem", "Pedidos agora"], ascending=[True, False], kind="stable"))
            saiu = (defeitos[defeitos["Sinal"] == "Saiu dos ofensores"]
                    .sort_values("No relatório antes", ascending=False))
            _titulo_secao(f"Pedem ação · {len(acao)} de {len(defeitos)} defeitos", "10px 0 8px")
            if len(acao):
                mostrar = acao.head(15)
                linhas = []
                for _, l in mostrar.iterrows():
                    sub = " · ".join(p for p in (l["Detalhe do sinal"], l["Time"]) if p)
                    nome = comparacao.nome_curto(l["Nome"]) or "—"
                    linhas.append([
                        _sinal_html(l["Sinal"], comparacao.SINAIS_DEFEITO[l["Sinal"]]),
                        _esc(l["Defeito"]),
                        f"{_esc(nome)}<span class='sub'>{_esc(sub)}</span>",
                        _pedidos_html(l["Pedidos agora"], l["Variação"]),
                    ])
                st.markdown(_tabela_html(
                    [("Sinal", 17, False), ("Defeito", 12, False), ("Nome", 55, False),
                     (f"Pedidos em {mes_rot}", 16, True)], linhas), unsafe_allow_html=True)
                if len(acao) > len(mostrar):
                    _nota(f"E mais {len(acao) - len(mostrar)} na tabela completa, abaixo.")
            else:
                st.caption("Nenhum defeito pede ação nesta comparação.")
            if len(saiu):
                nomes = [f"{l['Defeito']}" + (f" ({l['Detalhe do sinal']})" if l["Detalhe do sinal"] else "")
                         for _, l in saiu.head(4).iterrows()]
                if len(saiu) > 4:
                    nomes.append(f"mais {len(saiu) - 4}")
                _nota(f"<span class='melhor'>Saíram dos ofensores:</span> "
                      f"{_esc(comparacao._lista(nomes))}.")

            if st.toggle("Ver todos os defeitos", key="cmp_def_todos"):
                busca = st.text_input("Buscar por defeito, nome, time ou observação",
                                      key="cmp_busca_def")
                todos = _filtrar_texto(defeitos, busca) if busca else defeitos
                colunas = ["Defeito", "Nome", "Time", "Tipo", "Sinal", "Fase antes", "Fase agora",
                           "Entrega antes", "Entrega agora", "Pedidos antes", "Pedidos agora",
                           "Variação", "Pós-entrega na semana", "No relatório agora",
                           "Observações"]
                st.dataframe(
                    todos[colunas], hide_index=True, use_container_width=True,
                    height=min(38 + 35 * len(todos), 560),
                    column_config={
                        "Nome": st.column_config.TextColumn(width="large"),
                        "Observações": st.column_config.TextColumn(width="large"),
                        "Entrega antes": st.column_config.DateColumn(format="DD/MM/YYYY"),
                        "Entrega agora": st.column_config.DateColumn(format="DD/MM/YYYY"),
                    })
                st.caption(
                    "**Pós-entrega na semana**: pedidos criados no período novo depois da data "
                    "de entrega do defeito — com fase de corrigido, a correção não pegou; sem, "
                    "a data passou e o erro continua. **Pedidos antes/agora** são do mês "
                    f"comparado ({mes_rot}); **No relatório agora**, da tabela de ofensores do "
                    "mês do relatório atual.")

    # ── Erros ─────────────────────────────────────────────────────────────────
    with t_err:
        periodo = ctx["periodo_erros"]
        if periodo.get("nota"):
            st.info(periodo["nota"])
        else:
            atencao = comparacao.erros_em_atencao(c)
            _titulo_secao(f"Pedem atenção · {periodo['inicio']:%d/%m} a {periodo['fim']:%d/%m}",
                          "10px 0 8px")
            if len(atencao):
                tons = {"Novo": "pior", "Voltou": "pior", "Em alta": "atencao"}
                linhas = [[_sinal_html(l["Sinal"], tons[l["Sinal"]]),
                           _esc(comparacao.resumir_mensagem(l["Mensagem de erro"], 110)),
                           _num_br(l["Pedidos"]),
                           f"<span class='fraco'>{_esc(l['Comparado a antes'])}</span>",
                           _esc(l["Classificado como"])]
                          for _, l in atencao.iterrows()]
                st.markdown(_tabela_html(
                    [("Sinal", 10, False), ("Erro", 46, False), ("Pedidos", 9, True),
                     ("Comparado a antes", 19, False), ("Classificado como", 16, False)],
                    linhas), unsafe_allow_html=True)
            else:
                st.caption("Nenhum erro novo ou em alta nesta comparação.")
            sumidos = c.erros_sumidos
            if len(sumidos):
                maior = sumidos.sort_values("Na referência", ascending=False).iloc[0]
                janela = comparacao._janela(periodo["dias_ref"])
                _nota(f"<span class='melhor'>{_num_br(len(sumidos))} "
                      f"{'erro sumiu' if len(sumidos) == 1 else 'erros sumiram'}</span> no "
                      f"período{':' if len(sumidos) == 1 else ' · o maior:'} "
                      f"{_esc(comparacao.resumir_mensagem(maior['Mensagem de erro'], 80))} "
                      f"({_num_br(maior['Na referência'])} {janela} antes).")
            _nota(_esc(
                f"Período novo: {periodo['inicio']:%d/%m} a {periodo['fim']:%d/%m} "
                f"({periodo['dias']} dias), comparado aos {periodo['dias_ref']} dias anteriores. "
                f"Novo: não aparecia na fotografia anterior e tem {c.limites.min_ocorrencias}+ "
                f"ocorrências. Em alta: média diária {_num_br(c.limites.fator_alta, 1)}× a "
                f"dos dias anteriores. Voltou: não apareceu nos dias anteriores."))

            if st.toggle("Ver detalhes dos erros", key="cmp_err_detalhe"):
                _detalhe_erros(c)

    # ── Qualidade dos dados ───────────────────────────────────────────────────
    with t_dados:
        # Os avisos por extenso: no texto de cima eles viram uma ressalva curta.
        if ctx["avisos"]:
            _titulo_secao("Para ler esta comparação", "6px 0 4px")
            st.markdown("<div class='cmp-nota'><ul style='margin:0;padding-left:18px'>"
                        + "".join(f"<li>{_esc(a)}</li>" for a in ctx["avisos"])
                        + "</ul></div>", unsafe_allow_html=True)
        _titulo_secao("Jornadas desta visão", "6px 0 6px")
        st.dataframe(ctx["jornadas"], hide_index=True)
        _titulo_secao("Falhas por mês nas duas fotografias")
        qualidade = ctx["qualidade"].drop(columns="Mes")
        qualidade["Jornada"] = qualidade["Jornada"].map(fallout_core.nome_exibicao)
        st.dataframe(qualidade, hide_index=True, use_container_width=True,
                     height=min(38 + 35 * len(qualidade), 420))
        st.caption("**Suspeito**: o mês perdeu falhas de uma fotografia para a outra (5% ou mais, "
                   f"com pelo menos {c.limites.min_ocorrencias}). **Queda forte**: perdeu 30% ou "
                   "mais — quase sempre é o arquivo novo.")


_render_aba(aba_comparacao, _aba_comparacao)


# ── Aba Previsto x Realizado ──────────────────────────────────────────────────
# Por que o mês não ficou na projeção feita no mês anterior (ver previsao.py).
# A previsão sai de uma fotografia gravada; o realizado, dos dados na tela. Fica
# aqui embaixo, e não junto da Projeção, porque usa os desenhos de tabela da
# Comparação, definidos logo acima.
@st.cache_data(show_spinner="Comparando o previsto com o realizado...")
def carregar_previsto(foto_id, _gravada_em, jornada, mes):
    import previsao
    df_j, resumo_j = carregar_dados(jornada)
    return previsao.explicar(_base_dir(), foto_id, jornada, df_j, resumo_j, mes)


def _aba_previsto():
    import previsao

    mes_rot = mes_labels[mes_escolhido]
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Previsto x Realizado · {JORNADA_LABEL} · "
        f"{mes_rot}</div>",
        unsafe_allow_html=True,
    )
    if MES_SEM_BRUTO:
        st.info(AVISO_SEM_BRUTO)
        return
    ids = previsao.fotografias_com_previsao(_base_dir(), jornada_escolhida, mes_escolhido)
    if not ids:
        st.info(f"Nenhuma fotografia gravada antes de {mes_rot} projetou esse mês para "
                f"{JORNADA_LABEL}. A previsão sai da fotografia do mês anterior: grave uma "
                "em **Gravar fotografia**, no topo, toda semana.")
        return
    fotos = {f["id"]: f for f in fotografia.listar(_base_dir())}
    if st.session_state.get("prev_foto") not in ids:
        st.session_state["prev_foto"] = ids[0]
    col_f, _vazio = st.columns([1.6, 4])
    with col_f:
        foto_id = st.selectbox("Previsão da fotografia", ids, key="prev_foto",
                               format_func=lambda i: f"Dados até {fotos[i]['corte']:%d/%m/%Y}")
    try:
        r = carregar_previsto(foto_id, fotos[foto_id]["gravada_em"], jornada_escolhida,
                              int(mes_escolhido))
    except ValueError as e:
        st.info(str(e))
        return

    dif = r.realizado - r.previsto
    cartoes = [
        _kpi_tile(f"Previsto para {mes_rot}", _num_br(r.previsto, 2) + "%", "", False, [],
                  destaque=False,
                  rodape=f"fotografia de {fotos[foto_id]['corte']:%d/%m}, a partir de "
                         f"{fallout_core.MESES_LABEL[r.mes_base]}"),
        _kpi_tile(f"Realizado em {mes_rot}", _num_br(r.realizado, 2) + "%",
                  ("" if abs(dif) < 0.005 else _num_br(abs(dif), 2) + "pp"), dif > 0, [],
                  destaque=dif >= 0.005, rodape=f"{r.dias} dias de dados"),
    ]
    _plan = (resumo.loc[mes_escolhido, "PctPlanejado"]
             if "PctPlanejado" in resumo.columns and mes_escolhido in resumo.index else None)
    if _plan is not None and _plan == _plan and abs(_plan - r.previsto) >= 0.005:
        cartoes.append(_kpi_tile("Planejado no histórico", _num_br(_plan, 2) + "%", "", False,
                                 [], destaque=False, rodape="a meta digitada para o mês"))
    st.markdown("<div class='kpi-row'>" + "".join(cartoes) + "</div>", unsafe_allow_html=True)
    st.markdown(_resumo_html(r.texto), unsafe_allow_html=True)
    if not r.conferencia["ok"]:
        st.warning("A soma das categorias não fechou com a previsão ou com o realizado: "
                   f"{r.conferencia}")

    # ── Por categoria ─────────────────────────────────────────────────────────
    _titulo_secao("Por categoria", "18px 0 6px")
    tab = r.categorias.reindex(r.categorias["Diferença (pp)"].abs()
                               .sort_values(ascending=False).index)
    maior = tab["Diferença (pp)"].abs().max() or None
    linhas = [[_esc(l["Categoria"]), _num_br(l["Previsto"], 2) + "%",
               _num_br(l["Realizado"], 2) + "%", _pp_html(l["Diferença (pp)"], maior),
               f"<span class='fraco'>{_esc(l['O que explica'] or l['O que a previsão supôs'])}</span>"]
              for _, l in tab.iterrows()]
    linhas.append(("total", ["Fallout do mês", _num_br(r.previsto, 2) + "%",
                             _num_br(r.realizado, 2) + "%", _pp_html(dif), ""]))
    st.markdown(_tabela_html(
        [("Categoria", 24, False), ("Previsto", 9, True), ("Realizado", 9, True),
         ("Diferença", 15, True), ("O que explica", 43, False)], linhas),
        unsafe_allow_html=True)
    _nota("A previsão supôs, para cada categoria: <b>Resolvido, Tratado e Incidente</b> em "
          "zero (não se repetem); <b>Em Tratamento</b> com o % do mês anterior menos o que as "
          "entregas planejadas tiram; e as <b>demais</b> iguais ao mês anterior. "
          "<span class='pior'>Vermelho</span> está acima do previsto; "
          "<span class='melhor'>verde</span>, abaixo."
          + (f" Com {r.dias} dias de dados, o mês ainda vai mudar." if r.dias < 15 else ""))

    # ── Em Tratamento, por defeito ────────────────────────────────────────────
    d = r.defeitos
    if len(d):
        _titulo_secao(f"Em Tratamento, por defeito", "22px 0 6px")
        relevantes = d[d["Diferença (pp)"].abs() >= 0.005].head(15)
        maior_d = relevantes["Diferença (pp)"].abs().max() if len(relevantes) else None
        st.markdown(_tabela_html(
            [("Defeito", 16, False), ("Previsto", 9, True), ("Realizado", 9, True),
             ("Diferença", 15, True), ("O que aconteceu", 51, False)],
            [[f"<b>{_esc(l['Defeito'])}</b><br><span class='fraco'>"
              f"{_esc(_nome_curto(l['Nome']))}</span>",
              _num_br(l["Previsto"], 2) + "%", _num_br(l["Realizado"], 2) + "%",
              _pp_html(l["Diferença (pp)"], maior_d), _esc(l["O que aconteceu"])]
             for _, l in relevantes.iterrows()]),
            unsafe_allow_html=True)
        _nota("Previsto = o % do defeito no mês anterior menos o que a entrega dele tira deste "
              "mês. Entrega e fase são as do Octane na fotografia da previsão e agora.")

    # ── Categorias sem defeito, por mensagem ──────────────────────────────────
    for cat, t in r.mensagens.items():
        t = t[t["Diferença (pp)"].abs() >= 0.005].head(8)
        if not len(t):
            continue
        por_msg = cat in previsao.POR_MENSAGEM
        _titulo_secao(f"{cat}, por {'mensagem de erro' if por_msg else 'defeito'}", "22px 0 6px")
        maior_m = t["Diferença (pp)"].abs().max()
        st.markdown(_tabela_html(
            [("Mensagem de erro" if por_msg else "Defeito", 58, False),
             ("Previsto", 9, True), ("Realizado", 9, True),
             ("Diferença", 24, True)],
            [[_esc(l["Mensagem"]), _num_br(l["Previsto"], 2) + "%",
              _num_br(l["Realizado"], 2) + "%", _pp_html(l["Diferença (pp)"], maior_m)]
             for _, l in t.iterrows()]),
            unsafe_allow_html=True)


def _nome_curto(texto, limite=90):
    """Nome do defeito encurtado para a tabela; vazio quando não há."""
    if texto is None or texto != texto:
        return ""
    texto = str(texto)
    return texto if len(texto) <= limite else texto[:limite - 1].rstrip() + "…"


_render_aba(aba_previsto, _aba_previsto)


# ── Aba Parâmetros ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cruzando os parâmetros com o Octane...")
def carregar_parametros(caminho, _mtime):
    """A data de modificação só compõe a chave: arquivo novo, leitura nova."""
    return parametros.avaliar(caminho, _base_dir())


def _aba_parametros():
    st.markdown(
        f"<div style='font:600 10px {FONTE};letter-spacing:.1em;text-transform:uppercase;"
        f"color:{COR_MUTED};margin:6px 0 14px'>Parâmetros · o que cancelar</div>",
        unsafe_allow_html=True,
    )
    caminho = parametros.localizar_relatorio(_base_dir())
    if caminho is None:
        st.info("Nenhum relatório de parâmetros encontrado. Salve o export em "
                "`extracoes/` com **Parametros** ou **Parameters** no nome "
                "(ex.: `RelatorioParametros.csv`) e clique em **Atualizar dados**.")
        return
    try:
        df_par, inativos = carregar_parametros(caminho, os.path.getmtime(caminho))
    except (ValueError, FileNotFoundError) as e:
        st.error(str(e))
        return

    modificado = datetime.fromtimestamp(os.path.getmtime(caminho))
    st.caption(f"Relatório: `{os.path.basename(caminho)}` · salvo em "
               f"{modificado:%d/%m/%Y %H:%M} · a aba não segue a jornada nem o mês "
               f"selecionados no topo.")

    # A jornada aqui é a do parâmetro (PF, Base, XSELL…), não a pasta do app.
    col_j, col_b = st.columns([2.4, 3])
    with col_j:
        jornadas_par = sorted(df_par["Jornada"].unique())
        escolhidas = st.multiselect("Jornada do parâmetro", jornadas_par,
                                    placeholder="Todas", key="par_jornadas")
    with col_b:
        busca_par = st.text_input("Buscar por DFT ou parâmetro", key="par_busca")
    if escolhidas:
        df_par = df_par[df_par["Jornada"].isin(escolhidas)]
    if busca_par:
        df_par = _filtrar_texto(df_par, busca_par)

    contagem = df_par["Status"].value_counts()
    cancelar = df_par[df_par["Status"] == parametros.STATUS_CANCELAR]
    rotulo_total = "parâmetros" if inativos is None else "ativos"
    chips = [f"<span class='chip'>{_num_br(len(df_par))} {rotulo_total}</span>"]
    chips += [f"<span class='chip'>{_num_br(contagem[s])} · {s}</span>"
              for s in parametros.ORDEM_STATUS if contagem.get(s, 0)]
    st.markdown(f"<div style='margin:4px 0 12px'>{''.join(chips)}</div>",
                unsafe_allow_html=True)

    colunas = ["Jornada", "DFT", "Parâmetro", "Fase no Octane", "Encerrado em",
               "Nome do DFT", "Etapa", "Erro", "Ofensor", "Criado em", "Id"]
    colunas = [c for c in colunas if df_par[c].notna().any()]
    formato = {"Encerrado em": st.column_config.DatetimeColumn(format="DD/MM/YYYY"),
               "Criado em": st.column_config.DatetimeColumn(format="DD/MM/YYYY")}

    st.markdown(
        f"<div style='font:600 13px {FONTE};margin:6px 0 8px'>"
        f"{_num_br(len(cancelar))} parâmetros a cancelar · DFT "
        f"{' ou '.join(parametros.FASES_CANCELAR)} no Octane</div>",
        unsafe_allow_html=True,
    )
    if len(cancelar):
        st.dataframe(cancelar[colunas], hide_index=True, use_container_width=True,
                     column_config=formato,
                     height=min(38 + 35 * len(cancelar), 560))
    else:
        st.success("Nenhum parâmetro aponta para DFT fechado ou cancelado.")

    demais = df_par[df_par["Status"] != parametros.STATUS_CANCELAR]
    if len(demais):
        with st.expander(f"Demais parâmetros ({_num_br(len(demais))})"):
            st.dataframe(demais[["Status"] + colunas], hide_index=True,
                         use_container_width=True, column_config=formato)

    st.caption(
        ("Entram todos os parâmetros do relatório. " if inativos is None else
         "Entram só os parâmetros ativos"
         + (f" — {_num_br(inativos)} inativos do relatório ficaram de fora" if inativos else "")
         + ". ")
        + "**Cancelar**: o DFT está Fechado ou Cancelado no Octane — vale a fase "
        "do próprio DFT, mesmo quando ele aponta para uma US de Melhoria. "
        "**Manter**: o DFT está em qualquer outra fase — inclusive **Rejeitado**, "
        "que aqui não pede cancelamento. **User Story**: o número é de uma US, e "
        "não de um DFT. **Não encontrado**: "
        "número que não está no `RelatorioDFTOctane.xlsx`. **Sem DFT**: 999999, "
        "PDST-, INC- ou texto no lugar do número."
    )


_render_aba(aba_parametros, _aba_parametros)
