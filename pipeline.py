import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# ── 1. Leitura das bases ────────────────────────────────────────────────────

import glob, os, sys

# Jornada: pode ser passada como argumento ou usa "Base Móvel" como padrão
# Uso: python pipeline.py "Base Móvel"   ou   python pipeline.py "Cross Sell"
# Uso: python pipeline.py "Base Móvel" [mes_numero]
# Ex:  python pipeline.py "Base Móvel" 5   → gera dashboard do mês 5 (maio)
JORNADA = sys.argv[1] if len(sys.argv) > 1 else "Base Móvel"

# "Base + Cross Sell" combina as duas jornadas
JORNADAS_COMBO = ["Base Móvel", "Cross Sell"] if JORNADA == "Base + Cross Sell" else [JORNADA]
print(f"Jornada: {JORNADA}  |  Pastas: {', '.join(JORNADAS_COMBO)}")

# Convenção de nome: extrações/<Jornada>/falhas/RelatorioDeFalhas_jan26.csv
#                   extrações/<Jornada>/sucessos/RelatorioDeSucessos_jan26.csv

ABREV_MES = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,
             "jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}

def mes_do_arquivo(path):
    nome = os.path.splitext(os.path.basename(path))[0].lower()
    for abrev, num in ABREV_MES.items():
        if abrev in nome:
            return num
    return None

# ── Falhas: lê todos os CSVs das pastas e concatena ─────────────────────
_falhas_parts = []
for _j in JORNADAS_COMBO:
    _arqs = sorted(glob.glob(os.path.join("extrações", _j, "falhas/*.csv")))
    if not _arqs:
        raise FileNotFoundError(f"Nenhum arquivo encontrado em extrações/{_j}/falhas/")
    _falhas_parts.extend(_arqs)
    print(f"  [{_j}] falhas: {len(_arqs)} arquivo(s)")

df_falhas = pd.concat([pd.read_csv(f) for f in _falhas_parts], ignore_index=True)
df_falhas["CreatedDate"] = pd.to_datetime(df_falhas["CreatedDate"], utc=True, errors="coerce")
print(f"Falhas carregadas: {len(_falhas_parts)} arquivo(s), {len(df_falhas):,} linhas")

# ── Octane: sempre o arquivo mais recente (snapshot atual) ───────────────
df_dft = pd.read_excel("RelatorioDFTOctane.xlsx")
df_us  = pd.read_excel("RelatorioUSOctane.xlsx")

# ── Sucessos: lê todos os CSVs das pastas, deriva mês do nome do arquivo ─
partes_suc = []
_suc_count = 0
for _j in JORNADAS_COMBO:
    _arqs_suc = sorted(glob.glob(os.path.join("extrações", _j, "sucessos/*.csv")))
    if not _arqs_suc:
        raise FileNotFoundError(f"Nenhum arquivo encontrado em extrações/{_j}/sucessos/")
    for f in _arqs_suc:
        mes = mes_do_arquivo(f)
        if mes is None:
            print(f"  AVISO: não foi possível determinar o mês de '{f}', ignorando.")
            continue
        df_tmp = pd.read_csv(f).rename(columns={"expr0": "Sucessos"})
        df_tmp["Mes"] = mes
        partes_suc.append(df_tmp[["Mes", "Sucessos"]])
    _suc_count += len(_arqs_suc)
    print(f"  [{_j}] sucessos: {len(_arqs_suc)} arquivo(s)")

sucessos_mes = (
    pd.concat(partes_suc, ignore_index=True)
    .groupby("Mes")["Sucessos"].sum()
)
print(f"Sucessos carregados: {_suc_count} arquivo(s)")

# ── 2. Renomear colunas longas das falhas ───────────────────────────────────

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

# ── 3. Preparar Octane: unir DFT + US e selecionar colunas relevantes ──────

colunas_base = ["ID", "Name", "Phase", "Bugfix Milestone", "Team", "Type"]
COL_US_MELHORIA = "US de Melhoria"

# DFTs: se tiver "US de Melhoria" preenchida, substituir campos pelo da US vinculada
df_dft_prep = df_dft[colunas_base + ([COL_US_MELHORIA] if COL_US_MELHORIA in df_dft.columns else [])].copy()

if COL_US_MELHORIA in df_dft_prep.columns:
    def _norm_id(v):
        try:    return str(int(float(str(v).strip())))
        except: return str(v).strip()

    df_us_idx = df_us.set_index("ID")[["Name", "Phase", "Bugfix Milestone", "Team", "Type"]].copy()
    df_us_idx.index = [_norm_id(v) for v in df_us_idx.index]

    df_dft_prep[COL_US_MELHORIA] = df_dft_prep[COL_US_MELHORIA].apply(_norm_id)
    _has_us = df_dft_prep[COL_US_MELHORIA].notna() & ~df_dft_prep[COL_US_MELHORIA].isin(["", "nan", "None"])
    print(f"\nDFTs com US de Melhoria: {_has_us.sum()}")
    for i, row in df_dft_prep[_has_us].iterrows():
        us_id = row[COL_US_MELHORIA]
        found = us_id in df_us_idx.index
        print(f"  DFT {row['ID']} → US {us_id} {'✓' if found else '✗ NÃO ENCONTRADA'}")
        if found:
            us = df_us_idx.loc[us_id]
            for col in ["Name", "Phase", "Bugfix Milestone", "Team"]:
                if pd.notna(us[col]):
                    df_dft_prep.at[i, col] = us[col]
            df_dft_prep.at[i, "Type"] = "User Story"

df_octane = (
    pd.concat([df_dft_prep[colunas_base], df_us[colunas_base]], ignore_index=True)
    .drop_duplicates(subset="ID")
    .rename(columns={
        "ID":               "DefectNumber__c",
        "Name":             "DFT_Name",
        "Phase":            "DFT_Phase",
        "Bugfix Milestone": "DFT_BugfixMilestone",
        "Team":             "DFT_Team",
        "Type":             "DFT_Type",
    })
)

# DefectNumber__c na base de falhas pode vir como float (ex: 182366.0) — normaliza para int string
df_falhas["DefectNumber_orig"] = df_falhas["DefectNumber__c"].astype(str).str.strip()
df_falhas["DefectNumber__c"] = (
    pd.to_numeric(df_falhas["DefectNumber__c"], errors="coerce")
    .fillna(-1)
    .astype(int)
)
df_octane["DefectNumber__c"] = (
    pd.to_numeric(df_octane["DefectNumber__c"], errors="coerce")
    .fillna(-1)
    .astype(int)
)

# ── 4. Join: falhas ← Octane (left join, não-batedores ficam com NaN) ───────

df = df_falhas.merge(df_octane, on="DefectNumber__c", how="left")

# Desduplicar por OrderNumber: mantém a linha com DFT real (> 0) se existir
df = (
    df.sort_values("DefectNumber__c", ascending=False)   # DFT real primeiro
      .drop_duplicates(subset=["OrderNumber"])
      .reset_index(drop=True)
)

# ── 5. Diagnóstico rápido ───────────────────────────────────────────────────

total         = len(df)
com_octane    = df["DFT_Phase"].notna().sum()
sem_octane    = df["DFT_Phase"].isna().sum()

print(f"Total de registros de falha : {total:,}")
print(f"Com match no Octane         : {com_octane:,} ({com_octane/total:.1%})")
print(f"Sem match no Octane         : {sem_octane:,} ({sem_octane/total:.1%})")
print()
print("Primeiras linhas do DataFrame unificado:")
print(df[["CreatedDate", "OrderNumber", "DefectNumber__c", "DFT_Phase", "DFT_BugfixMilestone"]].head(5).to_string())

# ── 6. Gráfico: % de falha por mês ─────────────────────────────────────────

MESES_PT = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
            7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}

# Falhas por mês (contagem de linhas = ocorrências de falha)
df["Mes"] = df["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month
falhas_mes = df.groupby("Mes").size().rename("Falhas")

# sucessos_mes já carregado na seção 1

# Consolidar e calcular %
resumo = pd.DataFrame({"Falhas": falhas_mes, "Sucessos": sucessos_mes}).fillna(0)
resumo["Total"]    = resumo["Falhas"] + resumo["Sucessos"]
resumo["Pct"]      = resumo["Falhas"] / resumo["Total"] * 100
resumo["MesLabel"] = resumo.index.map(MESES_PT)

# Plot
fig, ax = plt.subplots(figsize=(10, 5))

ax.plot(resumo["MesLabel"], resumo["Pct"],
        color="black", linewidth=2, marker="o", markersize=6, label="% Falha")

ax.axhline(y=1, color="#1565C0", linewidth=1.5, linestyle="--", label="Meta (1%)")

# Rótulos sobre cada ponto
for mes, row in resumo.iterrows():
    ax.annotate(
        f"{row['Pct']:.1f}%",
        xy=(row["MesLabel"], row["Pct"]),
        xytext=(0, 10),
        textcoords="offset points",
        ha="center", fontsize=9, color="black"
    )

ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.1f}%"))
ax.set_xlabel("")
ax.set_ylabel("% de Falha", fontsize=10)
ax.set_title("Taxa de Falha por Mês", fontsize=13, fontweight="bold")
ax.legend(fontsize=9)
ax.grid(axis="y", linestyle="--", alpha=0.4)
ax.spines[["top", "right"]].set_visible(False)

plt.tight_layout()
plt.close(fig)

# ── 7. Tabela: Volume de Pedidos por mês ────────────────────────────────────

MESES_LABEL = {1: "jan-26", 2: "fev-26", 3: "mar-26", 4: "abr-26",
               5: "mai-26", 6: "jun-26", 7: "jul-26", 8: "ago-26",
               9: "set-26", 10: "out-26", 11: "nov-26", 12: "dez-26"}

meses_disponiveis = sorted(resumo.index.tolist())
col_labels = [MESES_LABEL[m] for m in meses_disponiveis]

row_sucessos   = [f"{int(resumo.loc[m, 'Sucessos']):,}".replace(",", ".") for m in meses_disponiveis]
row_falhas     = [f"{int(resumo.loc[m, 'Falhas']):,}".replace(",", ".")   for m in meses_disponiveis]
row_fallout    = [f"{resumo.loc[m, 'Pct']:.2f}%".replace(".", ",")        for m in meses_disponiveis]

n_cols = len(meses_disponiveis)
fig_w  = max(6, 1.5 * n_cols + 2)

fig, ax = plt.subplots(figsize=(fig_w, 1.8))
ax.set_axis_off()

# Cabeçalho duplo: título geral + labels de mês
COR_HEADER  = "#8B0000"   # vermelho escuro
COR_BRANCO  = "#FFFFFF"
COR_CINZA   = "#F2F2F2"
COR_TEXTO   = "#222222"

# Estrutura da tabela: coluna 0 = label da linha, colunas 1..n = meses
n_rows = 3  # Vendas com Sucesso / Falha / Fallout Rate
n_tcols = n_cols + 1

cell_h = 0.28
header_h = 0.22
label_h  = 0.22
total_h  = header_h + label_h + n_rows * cell_h + 0.05

fig.set_size_inches(fig_w, total_h + 0.3)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, n_tcols)
ax.set_ylim(0, total_h)
ax.set_axis_off()

y_top = total_h

col_w = [2.2] + [1.0] * n_cols  # largura relativa de cada coluna
x_pos = [sum(col_w[:i]) for i in range(n_tcols + 1)]  # posições x acumuladas
total_w = x_pos[-1]

# Reescala para n_tcols unidades
scale = n_tcols / total_w
x_pos = [x * scale for x in x_pos]
col_w = [w * scale for w in col_w]

def draw_cell(ax, x, y, w, h, text, bg, fg="#222222", bold=False, fontsize=8, align="center"):
    ax.add_patch(plt.Rectangle((x, y - h), w, h, facecolor=bg, edgecolor="white", linewidth=0.5))
    weight = "bold" if bold else "normal"
    ha = "left" if align == "left" else "center"
    xtext = x + 0.08 if align == "left" else x + w / 2
    ax.text(xtext, y - h / 2, text, ha=ha, va="center",
            fontsize=fontsize, color=fg, fontweight=weight)

# Linha 0: "Volume de Pedidos" abrangendo colunas de mês
y = y_top
draw_cell(ax, x_pos[0], y, col_w[0], header_h, "", COR_HEADER, COR_BRANCO, bold=True, fontsize=9)
draw_cell(ax, x_pos[1], y, sum(col_w[1:]), header_h, "Volume de Pedidos",
          COR_HEADER, COR_BRANCO, bold=True, fontsize=9)

# Linha 1: labels dos meses
y -= header_h
draw_cell(ax, x_pos[0], y, col_w[0], label_h, "", COR_HEADER, COR_BRANCO, fontsize=8)
for i, label in enumerate(col_labels):
    draw_cell(ax, x_pos[i + 1], y, col_w[i + 1], label_h, label,
              COR_HEADER, COR_BRANCO, bold=True, fontsize=8)

# Linhas de dados
rows = [
    ("Vendas com Sucesso",    row_sucessos),
    ("Falha (Análise Técnica)", row_falhas),
    ("Fallout Rate",           row_fallout),
]
for r_idx, (label, valores) in enumerate(rows):
    y -= cell_h
    bg = COR_BRANCO if r_idx % 2 == 0 else COR_CINZA
    draw_cell(ax, x_pos[0], y, col_w[0], cell_h, label, bg, COR_TEXTO,
              fontsize=8, align="left")
    for i, val in enumerate(valores):
        draw_cell(ax, x_pos[i + 1], y, col_w[i + 1], cell_h, val, bg, COR_TEXTO, fontsize=8)

plt.close(fig)

# ── 8. Distribuição do Fallout – mês atual ──────────────────────────────────

hoje          = pd.Timestamp.now(tz="UTC").normalize()
_meses_com_dados = df["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month.value_counts()
_meses_com_dados = _meses_com_dados[_meses_com_dados >= 10]  # ignora meses com < 10 registros
mes_atual     = int(_meses_com_dados.index.max()) if len(_meses_com_dados) else hoje.tz_convert("America/Sao_Paulo").month
quinze_dias   = hoje - pd.Timedelta(days=15)

df_mes        = df[df["Mes"] == mes_atual].copy()
total_mes     = resumo.loc[mes_atual, "Total"]
fallout_pct   = resumo.loc[mes_atual, "Pct"]

def pct(n): return n / total_mes * 100

milestone_dt  = pd.to_datetime(df_mes["DFT_BugfixMilestone"], utc=True, errors="coerce")

FASES_CORRIGIDO = {"Corrigido", "Fechado"}
FASES_MOPS      = {"Cancelado", "Rejeitado"}

_corrigido     = df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO)
_tem_milestone = df_mes["DFT_BugfixMilestone"].notna()
_outros_mask   = df_mes["DefectNumber_orig"].str.strip().str.lower() == "enviado e-mail - outros times"

# Falha Pontual
pontual = df_mes[df_mes["DefectNumber__c"] == 999999]

# Em avaliação - Outros times: DefectNumber_orig é "Enviado e-mail - Outros Times"
outros_times = df_mes[_outros_mask]

# Em Avaliação por MOPs: DFT real com phase Cancelado ou Rejeitado
mops = df_mes[
    (df_mes["DefectNumber__c"] > 0) &
    (df_mes["DefectNumber__c"] != 999999) &
    df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS)
]

# Resolvido: fase corrigido/fechado + milestone há mais de 15 dias
resolvido = df_mes[
    _corrigido & _tem_milestone &
    (milestone_dt < quinze_dias)
]

# Em avaliação de eficácia: fase corrigido/fechado + milestone nos últimos 15 dias
tratado = df_mes[
    _corrigido & _tem_milestone &
    (milestone_dt >= quinze_dias) &
    (milestone_dt <= hoje)
]

_ja_encerrado = _corrigido & _tem_milestone & (milestone_dt <= hoje)  # futuro = ainda em tratamento

# Em Tratamento: DFT real, não pontual, não mops, não encerrado
em_trat = df_mes[
    df_mes["DefectNumber__c"].notna() &
    (df_mes["DefectNumber__c"] != 999999) &
    (df_mes["DefectNumber__c"] != -1) &
    ~df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS) &
    ~_ja_encerrado
]

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
    _in_mes    = (_rows_dft["Mes"] == mes_atual).any()
    _in_mops   = _rows_dft["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS).any()
    _in_enc    = (_rows_dft["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO) & _rows_dft["DFT_BugfixMilestone"].notna()).any()
    print(f"  Tem registros no mês atual : {_in_mes}")
    print(f"  Phase é MOPs (Cancelado/Rejeitado) : {_in_mops}")
    print(f"  Está encerrado (Corrigido+Milestone): {_in_enc}")
    print(f"  Aparece em em_trat   : {(em_trat['DefectNumber__c'] == _DFT_DEBUG).any()}")
else:
    print(f"\n[DEBUG DFT {_DFT_DEBUG}] NÃO encontrado em nenhuma linha do df (não está no CSV de falhas do Cross Sell)")
# ─────────────────────────────────────────────────────────────────────────────

planejado    = em_trat[em_trat["DFT_BugfixMilestone"].notna()]
us_sem_data  = em_trat[em_trat["DFT_BugfixMilestone"].isna() & (em_trat["DFT_Type"].fillna("").str.strip() == "User Story")]
dft_sem_data = em_trat[em_trat["DFT_BugfixMilestone"].isna() & (em_trat["DFT_Type"].fillna("").str.strip() != "User Story")]

# Falta associar: sem DFT real, e não é "outros times"
falta_assoc = df_mes[
    (df_mes["DefectNumber__c"].isna() | (df_mes["DefectNumber__c"] == -1)) &
    ~_outros_mask
]

# Planejamento de Redução: DFTs com milestone futuro (pro-rate por dia de entrega)
import calendar as _cal

_ms_future = pd.to_datetime(em_trat["DFT_BugfixMilestone"], utc=True, errors="coerce")
plan_futuro = em_trat[_ms_future > hoje].copy()
plan_futuro["MilestoneDate"] = _ms_future[plan_futuro.index].dt.tz_convert("America/Sao_Paulo").dt.normalize()

# Por DFT: contagem de pedidos e % bruta
dft_counts = (
    plan_futuro.groupby(["MilestoneDate", "DefectNumber__c"])
    .size().reset_index(name="n")
)
dft_counts["Pct_bruta"] = dft_counts["n"] / total_mes * 100

# Pro-rate: fração do mês restante após o milestone
def _prorate(ms_date):
    """Fração de dias restantes no mês após a data de entrega."""
    dias_no_mes = _cal.monthrange(ms_date.year, ms_date.month)[1]
    dias_restantes = dias_no_mes - ms_date.day
    return max(dias_restantes, 0) / dias_no_mes

dft_counts["ProRate"]   = dft_counts["MilestoneDate"].apply(
    lambda d: _prorate(d.to_pydatetime())
)
dft_counts["Pct_mes"]   = dft_counts["Pct_bruta"] * dft_counts["ProRate"]   # redução no mês de entrega
dft_counts["Pct_plena"] = dft_counts["Pct_bruta"]                            # redução nos meses seguintes

reducao = (
    dft_counts.groupby("MilestoneDate")
    .agg(
        DFTs      =("DefectNumber__c", lambda x: "\n".join(f"DFT{int(v)}" for v in sorted(x.unique()))),
        n         =("n", "sum"),
        Pct_mes   =("Pct_mes",   "sum"),
        Pct_plena =("Pct_plena", "sum"),
    )
    .reset_index()
    .sort_values("MilestoneDate")
)
# Pct exibida na tabela = redução plena (meses seguintes ao milestone)
reducao["Pct"] = reducao["Pct_plena"]

# ── Verificação de cobertura ──────────────────────────────────────────────────
soma = len(em_trat) + len(resolvido) + len(pontual) + len(falta_assoc) + len(tratado) + len(mops) + len(outros_times)
print(f"\nDistribuição Fallout {MESES_LABEL[mes_atual]} ({fallout_pct:.2f}%)")
print(f"  Em Tratamento          {pct(len(em_trat)):.2f}%  (Planej:{pct(len(planejado)):.2f}% | US s/data:{pct(len(us_sem_data)):.2f}% | DFT s/data:{pct(len(dft_sem_data)):.2f}%)")
print(f"  Resolvido              {pct(len(resolvido)):.2f}%")
print(f"  Falha Pontual          {pct(len(pontual)):.2f}%")
print(f"  Falta associar         {pct(len(falta_assoc)):.2f}%")
print(f"  Em avaliação eficácia  {pct(len(tratado)):.2f}%")
print(f"  Em Avaliação MOPs      {pct(len(mops)):.2f}%")
print(f"  Outros times           {pct(len(outros_times)):.2f}%")
print(f"  Total categorizado: {soma}/{len(df_mes)}")

# ── Tabela visual de distribuição ────────────────────────────────────────────

COR_HDR_D = "#C0392B"
COR_W     = "#FFFFFF"
COR_G     = "#F2F2F2"
COR_FG    = "#333333"

dist_rows = [
    # (label, valor, indent, is_header)
    (f"Distribuição Fallout ({fallout_pct:.2f}%)", "",                                   0, True),
    ("Em Tratamento/Avaliação pela Squad",          f"{pct(len(em_trat)):.2f}%",         0, False),
    ("Planejado",                                   f"{pct(len(planejado)):.2f}%",       1, False),
    ("US s/ data",                                  f"{pct(len(us_sem_data)):.2f}%",     2, False),
    ("DFT s/ data",                                 f"{pct(len(dft_sem_data)):.2f}%",    2, False),
    ("Resolvido",                                   f"{pct(len(resolvido)):.2f}%",       0, False),
    ("Falha Pontual",                               f"{pct(len(pontual)):.2f}%",         0, False),
    ("Falta associar problema ao Defeito/US",       f"{pct(len(falta_assoc)):.2f}%",     0, False),
    ("Em avaliação de eficácia",                    f"{pct(len(tratado)):.2f}%",         0, False),
    ("Em Avaliação por MOPs",                       f"{pct(len(mops)):.2f}%",            0, False),
    ("Em avaliação - Outros times",                 f"{pct(len(outros_times)):.2f}%",    0, False),
    (f"Planejamento Redução ({reducao['Pct_plena'].sum():.2f}% pleno)", "",                        0, True),
]
for _, r in reducao.iterrows():
    dist_rows.append((r["MilestoneDate"].strftime("%d/%m/%Y"), "", 1, False))
    for dft_id in r["DFTs"].split("\n"):
        dist_rows.append((dft_id.strip(), f"{r['Pct'] / (r['DFTs'].count(chr(10)) + 1):.2f}%", 2, False))

cell_h_d  = 0.30
fig_hd    = len(dist_rows) * cell_h_d + 0.2
fig_d, _  = plt.subplots()
ax_d      = fig_d.add_axes([0, 0, 1, 1])
fig_d.set_size_inches(7, fig_hd + 0.3)
ax_d.set_xlim(0, 7); ax_d.set_ylim(0, fig_hd + 0.3); ax_d.set_axis_off()

LABEL_W = 5.5; VAL_W = 1.5; INDENT = 0.35
yd = fig_hd + 0.15
toggle = False
for label, valor, indent, is_header in dist_rows:
    yd -= cell_h_d
    if is_header:
        bg = COR_HDR_D; fg = "#FFFFFF"; bold = True; fs = 8.5
    else:
        bg = COR_W if not toggle else COR_G; fg = COR_FG; bold = False; fs = 8
        toggle = not toggle
    ax_d.add_patch(plt.Rectangle((0, yd), LABEL_W + VAL_W, cell_h_d,
                                 facecolor=bg, edgecolor="#DDDDDD", linewidth=0.4))
    ax_d.text(0.12 + indent * INDENT, yd + cell_h_d / 2, label,
              ha="left", va="center", fontsize=fs, color=fg,
              fontweight="bold" if bold else "normal")
    if valor:
        ax_d.text(LABEL_W + VAL_W - 0.12, yd + cell_h_d / 2, valor,
                  ha="right", va="center", fontsize=fs, color=fg)

plt.close(fig_d)

# ── 9. Tabela detalhada por DFT/US (mês atual + mês anterior) ───────────────

mes_ant = mes_atual - 1 if mes_atual > 1 else 12

meses_detalhe = [mes_ant, mes_atual]
labels_detalhe = [MESES_LABEL[m][:3].capitalize() for m in meses_detalhe]  # ex: ["Mai", "Jun"]

# Base: falhas "Em Tratamento" (DFT real, não 999999) dos dois meses
# Exclui registros Resolvidos/Tratados: Phase == Fechado E Bugfix Milestone preenchido E Milestone < CreatedDate
df_dois_meses = df[df["Mes"].isin(meses_detalhe)].copy()

_ms_dois = pd.to_datetime(df_dois_meses["DFT_BugfixMilestone"], utc=True, errors="coerce")
_resolvido = (
    df_dois_meses["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO) &
    df_dois_meses["DFT_BugfixMilestone"].notna() &
    (_ms_dois <= hoje)
)

df_dois_meses = df_dois_meses[
    df_dois_meses["DefectNumber__c"].notna() &
    (df_dois_meses["DefectNumber__c"] != 999999) &
    (df_dois_meses["DefectNumber__c"] != -1) &
    ~_resolvido
]

# Total de pedidos por mês (para calcular %)
totais = {m: resumo.loc[m, "Total"] for m in meses_detalhe if m in resumo.index}

# Contagem de falhas por DFT por mês
pivot = (
    df_dois_meses.groupby(["DefectNumber__c", df_dois_meses["Mes"]])
    .size()
    .unstack(fill_value=0)
    .rename(columns={m: MESES_LABEL[m][:3].capitalize() for m in meses_detalhe})
)
for col, m in zip(labels_detalhe, meses_detalhe):
    if col not in pivot.columns:
        pivot[col] = 0
    pivot[f"pct_{col}"] = pivot[col] / totais.get(m, 1) * 100

# Ordenar pelo % do mês atual decrescente e filtrar apenas quem tem ocorrência no mês atual
col_atual = labels_detalhe[-1]
pivot = (
    pivot[pivot[col_atual] > 0]
    .sort_values(f"pct_{col_atual}", ascending=False)
    .reset_index()
)

# Enriquecer com dados do Octane
octane_lookup = df_octane.set_index("DefectNumber__c")[
    ["DFT_Name", "DFT_Team", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Type"]
]
pivot = pivot.join(octane_lookup, on="DefectNumber__c", how="left")

# Formatar Data Prod (Bugfix Milestone)
pivot["DataProd"] = pd.to_datetime(pivot["DFT_BugfixMilestone"], errors="coerce").dt.strftime("%d/%m/%Y")
pivot["DataProd"] = pivot["DataProd"].fillna("")

# Tipo: abreviar
pivot["Tipo"] = pivot["DFT_Type"].fillna("").str.strip()
pivot["Tipo"] = pivot["Tipo"].replace({"User Story": "US", "Defect": "Defeito"})

# ID formatado
pivot["ID_fmt"] = pivot["DefectNumber__c"].apply(
    lambda x: str(int(x)) if pd.notna(x) else ""
)

# ── Construir linhas da tabela ───────────────────────────────────────────────

# Cabeçalho
COLS = ["Classificação", "Agrupamento", "Tipo", "Defeito/US",
        "Data Prod\nnão resolvidos", "Nome Defeito", "Time",
        "Status\nDefeito/US"] + labels_detalhe

col_widths = [1.5, 2.8, 1.5, 1.5, 1.8, 5.5, 2.5, 2.5] + [1.0] * len(labels_detalhe)
total_w3 = sum(col_widths)

COR_H3    = "#2C2C2C"
COR_SUB   = "#E8E8E8"
COR_B3    = "#FFFFFF"
COR_C3    = "#F5F5F5"
COR_FG    = "#222222"

cell_h3  = 0.38
header_h3 = 0.38
n_data_rows = len(pivot) + 2  # +2: linha Sucesso e linha de grupo "Em Tratamento"
fig_h3   = header_h3 + n_data_rows * cell_h3 + 0.2
fig3     = plt.figure(figsize=(total_w3 * 0.7, fig_h3 * 1.1))
ax3      = fig3.add_axes([0, 0, 1, 1])
ax3.set_xlim(0, total_w3)
ax3.set_ylim(0, fig_h3)
ax3.set_axis_off()

x_starts = [sum(col_widths[:i]) for i in range(len(col_widths) + 1)]

def cell3(ax, col_idx, y, h, text, bg, fg=COR_FG, bold=False, fs=7, align="center", span=1):
    x = x_starts[col_idx]
    w = sum(col_widths[col_idx:col_idx + span])
    ax.add_patch(plt.Rectangle((x, y), w, h, facecolor=bg, edgecolor="#CCCCCC", linewidth=0.3))
    ha = "left" if align == "left" else "center"
    xt = x + 0.08 if align == "left" else x + w / 2
    ax.text(xt, y + h / 2, str(text), ha=ha, va="center", fontsize=fs,
            color=fg, fontweight="bold" if bold else "normal", wrap=True,
            multialignment=ha)

y3 = fig_h3 - header_h3

# Linha de cabeçalho
for ci, col in enumerate(COLS):
    cell3(ax3, ci, y3, header_h3, col, COR_H3, fg="#FFFFFF", bold=True, fs=7)

# Linha Sucesso
y3 -= cell_h3
cell3(ax3, 0, y3, cell_h3, "Sucesso", COR_B3, bold=False, fs=7)
for ci in range(1, len(COLS) - len(labels_detalhe)):
    cell3(ax3, ci, y3, cell_h3, "", COR_B3)
for li, col in enumerate(labels_detalhe):
    m = meses_detalhe[li]
    suc_pct = resumo.loc[m, "Sucessos"] / totais[m] * 100 if m in totais else 0
    cell3(ax3, len(COLS) - len(labels_detalhe) + li, y3, cell_h3,
          f"{suc_pct:.2f}%", COR_B3, fs=7)

# Linha de grupo "Em Tratamento"
y3 -= cell_h3
cell3(ax3, 0, y3, cell_h3, "Falha", COR_C3, bold=False, fs=7)
cell3(ax3, 1, y3, cell_h3, "Em Tratamento/\nAvaliação pela Squad", COR_C3, bold=False, fs=6.5)
for ci in range(2, len(COLS)):
    cell3(ax3, ci, y3, cell_h3, "", COR_C3)

# Linhas de DFT
for ri, row in pivot.iterrows():
    y3 -= cell_h3
    bg = COR_B3 if ri % 2 == 0 else COR_C3
    cell3(ax3, 0, y3, cell_h3, "", bg)
    cell3(ax3, 1, y3, cell_h3, "", bg)
    cell3(ax3, 2, y3, cell_h3, row["Tipo"], bg, fs=7)
    cell3(ax3, 3, y3, cell_h3, row["ID_fmt"], bg, fs=7)
    cell3(ax3, 4, y3, cell_h3, row["DataProd"], bg, fs=7)
    nome = str(row["DFT_Name"])[:80] if pd.notna(row["DFT_Name"]) else ""
    cell3(ax3, 5, y3, cell_h3, nome, bg, fs=6, align="left")
    cell3(ax3, 6, y3, cell_h3, str(row["DFT_Team"])[:30] if pd.notna(row["DFT_Team"]) else "", bg, fs=6.5)
    cell3(ax3, 7, y3, cell_h3, str(row["DFT_Phase"])[:25] if pd.notna(row["DFT_Phase"]) else "", bg, fs=6.5)
    for li, lbl in enumerate(labels_detalhe):
        val = f"{row[f'pct_{lbl}']:.2f}%"
        cell3(ax3, 8 + li, y3, cell_h3, val, bg, fs=7)

plt.close(fig3)

# ── 10. Tabela Em Tratamento – estilo com células mescladas ─────────────────

# Meses exibidos: anterior + atual
mes_ant2      = mes_atual - 1 if mes_atual > 1 else 12
meses_tab     = [mes_ant2, mes_atual]
labels_tab    = [MESES_LABEL[m][:3].capitalize() for m in meses_tab]   # ["Mai", "Jun"]
totais_tab    = {m: resumo.loc[m, "Total"] for m in meses_tab if m in resumo.index}

def pct_m(n, m):
    return n / totais_tab.get(m, 1) * 100

# Falhas Em Tratamento dos dois meses (mesma regra da seção 8: exclui Corrigido/Fechado com milestone passado)
_ms_trat2 = pd.to_datetime(df["DFT_BugfixMilestone"], utc=True, errors="coerce")
_fech2 = (
    df["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO) &
    df["DFT_BugfixMilestone"].notna() &
    (_ms_trat2 <= hoje)
)

df_trat2 = df[
    df["Mes"].isin(meses_tab) &
    df["DefectNumber__c"].notna() &
    (df["DefectNumber__c"] != 999999) &
    (df["DefectNumber__c"] != -1) &
    ~_fech2
].copy()
df_trat2["Mes"] = df_trat2["Mes"]

# Por DFT por mês
_idx_cols = ["DefectNumber__c", "DFT_Name", "DFT_Phase",
             "DFT_BugfixMilestone", "DFT_Type", "DFT_Team"]
pivot2 = (
    df_trat2.groupby(_idx_cols + ["Mes"], dropna=False)
    .size()
    .reset_index(name="n")
    .set_index(_idx_cols + ["Mes"])["n"]
    .unstack("Mes", fill_value=0)
    .reset_index()
)
for m in meses_tab:
    if m not in pivot2.columns:
        pivot2[m] = 0

# Diagnóstico: verificar DFTs em em_trat que somem no pivot
_dft_em_trat = set(em_trat["DefectNumber__c"].unique())
_dft_pivot_antes = set(pivot2["DefectNumber__c"].unique()) if "DefectNumber__c" in pivot2.columns else set()
_dft_sumidos = _dft_em_trat - _dft_pivot_antes
if _dft_sumidos:
    print(f"  [DIAG] DFTs em em_trat mas fora do pivot (antes filtro mes_atual): {sorted(_dft_sumidos)}")
_dft_sem_mes_atual = set(pivot2[pivot2[mes_atual] == 0]["DefectNumber__c"].unique()) if mes_atual in pivot2.columns else set()
_dft_filtrados = _dft_em_trat & _dft_sem_mes_atual
if _dft_filtrados:
    print(f"  [DIAG] DFTs removidos por nao ter ocorrencia em mes_atual ({mes_atual}): {sorted(_dft_filtrados)}")

# Ordenar pelo mês atual decrescente, mostrar só quem tem ocorrência no mês atual
pivot2 = pivot2[pivot2[mes_atual] > 0].sort_values(mes_atual, ascending=False)

# Notação de trimestre no ID (baseada no Bugfix Milestone)
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
    raw = row["DefectNumber__c"]
    try:
        base = str(int(raw))
    except Exception:
        base = str(raw)
    suffix = quarter_suffix(row["DFT_BugfixMilestone"])
    return base + suffix

def fmt_milestone(ms):
    dt = pd.to_datetime(ms, errors="coerce")
    return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else ""

# ── Layout ───────────────────────────────────────────────────────────────────

# Colunas: Classificação | Agrupamento | Tipo | Defeito/US | Data Prod | Nome | Time | Status | Mes...
COL_NAMES = ["Classificação", "Agrupamento", "Tipo", "Defeito / US",
             "Data Prod\nnão resolvidos", "Nome Defeito", "Time",
             "Status\nDefeito/US"] + labels_tab
COL_W     = [1.5, 2.8, 1.3, 1.5, 1.8, 6.2, 2.6, 2.2] + [1.3] * len(meses_tab)
TOTAL_W   = sum(COL_W)
X         = [sum(COL_W[:i]) for i in range(len(COL_W) + 1)]

COR_HDR  = "#C0392B"
COR_W2   = "#FFFFFF"
COR_G2   = "#F2F2F2"
COR_FG2  = "#222222"

CELL_H   = 0.32
HDR_H    = 0.38
n_data   = 1 + 1 + len(pivot2)          # Sucesso + Falha-grupo + DFTs
fig_h    = HDR_H + n_data * CELL_H + 0.45   # +0.45 p/ rodapé
fig5     = plt.figure(figsize=(TOTAL_W * 0.68, fig_h))
ax5      = fig5.add_axes([0, 0, 1, 1])
ax5.set_xlim(0, TOTAL_W)
ax5.set_ylim(0, fig_h)
ax5.set_axis_off()

def cell5(col, y, h, text, bg, fg=COR_FG2, bold=False, fs=7,
          align="center", span=1, italic=False):
    x = X[col]
    w = sum(COL_W[col:col + span])
    ax5.add_patch(plt.Rectangle((x, y), w, h,
                                facecolor=bg, edgecolor="#BBBBBB", linewidth=0.3))
    ha  = "left" if align == "left" else "center"
    xt  = x + 0.1 if align == "left" else x + w / 2
    style = "italic" if italic else "normal"
    ax5.text(xt, y + h / 2, str(text), ha=ha, va="center",
             fontsize=fs, color=fg, fontweight="bold" if bold else "normal",
             fontstyle=style, multialignment=ha, wrap=True, clip_on=True)

# ── Cabeçalho ────────────────────────────────────────────────────────────────
y5 = fig_h - HDR_H
for ci, col in enumerate(COL_NAMES):
    cell5(ci, y5, HDR_H, col, COR_HDR, fg="#FFFFFF", bold=True, fs=7)

# ── Linha Sucesso ─────────────────────────────────────────────────────────────
y5 -= CELL_H
cell5(0, y5, CELL_H, "Sucesso", COR_W2, fs=7)
for ci in range(1, len(COL_NAMES) - len(meses_tab)):
    cell5(ci, y5, CELL_H, "", COR_W2)
for li, m in enumerate(meses_tab):
    suc_n   = resumo.loc[m, "Sucessos"] if m in resumo.index else 0
    suc_pct = pct_m(suc_n, m)
    cell5(len(COL_NAMES) - len(meses_tab) + li, y5, CELL_H,
          f"{suc_pct:.2f}%", COR_W2, fs=7)

# ── Linha de grupo Falha / Em Tratamento ─────────────────────────────────────
y5 -= CELL_H
cell5(0, y5, CELL_H, "Falha", COR_G2, fs=7)
cell5(1, y5, CELL_H, "Em Tratamento/Avaliação\npela Squad", COR_G2, fs=6.5)
for ci in range(2, len(COL_NAMES)):
    cell5(ci, y5, CELL_H, "", COR_G2)

# ── Linhas de DFT/US ─────────────────────────────────────────────────────────
has_q3 = has_q4 = False
pivot2_reset_5 = pivot2.reset_index(drop=True)
for ri in range(len(pivot2_reset_5)):
    row = pivot2_reset_5.iloc[ri]
    y5 -= CELL_H
    bg = COR_W2 if ri % 2 == 0 else COR_G2

    tipo     = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
    tipo_fmt = "US" if tipo == "User Story" else "Defeito"
    id_str   = fmt_id({"DefectNumber__c": row["DefectNumber__c"],
                        "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]})
    ms_fmt   = fmt_milestone(row["DFT_BugfixMilestone"])
    nome     = str(row["DFT_Name"])[:70] if pd.notna(row["DFT_Name"]) else ""
    team     = str(row["DFT_Team"])[:30] if pd.notna(row["DFT_Team"]) else ""
    phase    = str(row["DFT_Phase"])[:30] if pd.notna(row["DFT_Phase"]) else ""

    if "*" in id_str and "**" not in id_str: has_q3 = True
    if "**" in id_str: has_q4 = True

    cell5(0, y5, CELL_H, "", bg)
    cell5(1, y5, CELL_H, "", bg)
    cell5(2, y5, CELL_H, tipo_fmt, bg, fs=7)
    cell5(3, y5, CELL_H, id_str, bg, fs=7)
    cell5(4, y5, CELL_H, ms_fmt, bg, fs=7)
    cell5(5, y5, CELL_H, nome, bg, fs=5.5, align="left")
    cell5(6, y5, CELL_H, team, bg, fs=6.5)
    cell5(7, y5, CELL_H, phase, bg, fs=6.5)
    for li, m in enumerate(meses_tab):
        val = int(row.loc[m]) if m in pivot2_reset_5.columns else 0
        pct_val = pct_m(val, m)
        cell5(8 + li, y5, CELL_H, f"{pct_val:.2f}%", bg, fs=7)

# ── Rodapé ───────────────────────────────────────────────────────────────────
rodape_parts = []
if has_q3: rodape_parts.append("* Q3")
if has_q4: rodape_parts.append("** Q4")
rodape = "  //  ".join(rodape_parts)
if rodape:
    ax5.text(0.15, y5 - 0.20, rodape, ha="left", va="top",
             fontsize=7, color="#444444", fontstyle="italic")

plt.close(fig5)

# ── 11. Visão Consolidada (dashboard) ───────────────────────────────────────

COR_RED   = "#C0392B"
COR_W     = "#FFFFFF"
COR_G     = "#F0F0F0"
COR_BLUE  = "#1565C0"
COR_GRAY  = "#B0B0B0"
COR_FG    = "#222222"

fig = plt.figure(figsize=(16, 9))
gs  = fig.add_gridspec(3, 2, height_ratios=[0.09, 1, 1.1], width_ratios=[1, 1.2])

ax_title = fig.add_subplot(gs[0, :])
ax_left  = fig.add_subplot(gs[1, 0])
ax_chart = fig.add_subplot(gs[1, 1])
ax_table = fig.add_subplot(gs[2, :])

for a in (ax_title, ax_left, ax_chart, ax_table):
    a.set_axis_off()

# ── Faixa de título ──────────────────────────────────────────────────────────
ax_title.set_xlim(0, 1)
ax_title.set_ylim(0, 1)
ax_title.add_patch(plt.Rectangle((0, 0), 1, 1, facecolor=COR_RED, transform=ax_title.transAxes, clip_on=False))
ax_title.text(0.5, 0.5, f"{JORNADA} – Realização e Projeção de Fallout",
              transform=ax_title.transAxes, ha="center", va="center",
              fontsize=16, fontweight="bold", color="white")

# ── 11a. Bloco esquerdo: Volume + Distribuição + Planejamento ──────────────

meses_vol = sorted(resumo.index.tolist())[-3:] if len(resumo) >= 3 else sorted(resumo.index.tolist())
labels_vol = [MESES_LABEL[m].capitalize() for m in meses_vol]

n_planos = len(reducao)
n_dft_lines = sum(r["DFTs"].count("\n") + 1 for _, r in reducao.iterrows())

left_rows_n = (
    2 +                       # cabeçalho Volume Pedidos + linha de meses
    3 +                       # Vendas/Falha/Fallout
    1 +                       # cabeçalho Distribuição Fallout
    1 + 3 + 1 + 4 +           # Em Trat + 3 sub + Resolvido + (Pontual,Falta,Tratado,MOPs)
    1 +                       # cabeçalho Planejamento Redução
    n_planos + n_dft_lines    # linhas de data + DFTs
)
LW, VW = 5.5, 2.0
left_w = LW + VW
row_h  = 0.72
left_h = left_rows_n * row_h
ax_left.set_xlim(0, left_w)
ax_left.set_ylim(0, left_h)

def cellL(x0, y0, w, h, text, bg, fg=COR_FG, bold=False, fs=6.5, align="center"):
    ax_left.add_patch(plt.Rectangle((x0, y0), w, h, facecolor=bg, edgecolor="white", linewidth=0.6))
    ha = "left" if align == "left" else "center"
    xt = x0 + 0.1 if align == "left" else x0 + w / 2
    ax_left.text(xt, y0 + h / 2, text, ha=ha, va="center", fontsize=fs,
                 color=fg, fontweight="bold" if bold else "normal")

y = left_h

# Colunas do bloco Volume: label larga + N colunas de mês iguais
n_vol     = len(meses_vol)
VOL_LBL_W = left_w * 0.46          # coluna do label (~46% da largura)
VOL_COL_W = (left_w - VOL_LBL_W) / n_vol  # colunas de mês iguais

# Coluna de % da Distribuição
DIST_LBL_W = left_w * 0.74         # label ocupa 74%
DIST_VAL_W = left_w - DIST_LBL_W   # % ocupa 26%

# ── Volume de Pedidos ────────────────────────────────────────────────────
y -= row_h
cellL(0, y, VOL_LBL_W, row_h, "", COR_RED)
cellL(VOL_LBL_W, y, left_w - VOL_LBL_W, row_h, "Volume de Pedidos",
      COR_RED, COR_W, bold=True, fs=7.5)

y -= row_h
cellL(0, y, VOL_LBL_W, row_h, "", COR_RED)
for i, lbl in enumerate(labels_vol):
    cellL(VOL_LBL_W + i * VOL_COL_W, y, VOL_COL_W, row_h,
          lbl, COR_RED, COR_W, bold=True, fs=7)

vol_data = [
    ("Vendas com Sucesso", "Sucessos", lambda v: f"{int(v):,}".replace(",", ".")),
    ("Falha (Análise Técnica)", "Falhas", lambda v: f"{int(v):,}".replace(",", ".")),
    ("Fallout Rate",  "Pct",     lambda v: f"{v:.2f}%".replace(".", ",")),
]
for ri, (label, key, fmt) in enumerate(vol_data):
    y -= row_h
    bg = COR_W if ri % 2 == 0 else COR_G
    cellL(0, y, VOL_LBL_W, row_h, label, bg, COR_FG, fs=8, align="left")
    for i, m in enumerate(meses_vol):
        cellL(VOL_LBL_W + i * VOL_COL_W, y, VOL_COL_W, row_h,
              fmt(resumo.loc[m, key]), bg, COR_FG, fs=8)

# ── Distribuição Fallout ─────────────────────────────────────────────────
y -= row_h
cellL(0, y, left_w, row_h,
      f"Distribuição Fallout ({fallout_pct:.2f}%)", COR_RED, COR_W, bold=True, fs=7.5)

# (label, valor, é_subitem, negrito_label)
dist_items = [
    ("Em Tratamento/Avaliação pela Squad",    pct(len(em_trat)),       False, True),
    ("Planejado",                             pct(len(planejado)),     True,  False),
    ("US s/ data",                            pct(len(us_sem_data)),   True,  False),
    ("DFT s/ data",                           pct(len(dft_sem_data)),  True,  False),
    ("Resolvido",                             pct(len(resolvido)),     False, False),
    ("Falha Pontual",                         pct(len(pontual)),       False, False),
    ("Falta associar problema ao Defeito/US", pct(len(falta_assoc)),   False, False),
    ("Em avaliação de eficácia",              pct(len(tratado)),       False, False),
    ("Em Avaliação por MOPs",                 pct(len(mops)),          False, False),
    ("Em avaliação - Outros times",           pct(len(outros_times)),  False, False),
]
toggle = False
for label, val, subitem, bold_label in dist_items:
    y -= row_h
    bg = COR_W if not toggle else COR_G
    toggle = not toggle
    val_str = f"{val:.2f}%".replace(".", ",")
    if subitem:
        # Fundo cobre toda a largura do label
        ax_left.add_patch(plt.Rectangle((0, y), DIST_LBL_W, row_h,
                                        facecolor=bg, edgecolor="white", linewidth=0.6))
        # Texto do sub-item alinhado à direita, próximo à coluna de %
        ax_left.text(DIST_LBL_W - 0.15, y + row_h / 2, label,
                     ha="right", va="center", fontsize=6.5, color=COR_FG)
        cellL(DIST_LBL_W, y, DIST_VAL_W, row_h, val_str, bg, COR_FG, fs=6.5)
    else:
        cellL(0, y, DIST_LBL_W, row_h, label, bg, COR_FG,
              bold=bold_label, fs=6.5, align="left")
        cellL(DIST_LBL_W, y, DIST_VAL_W, row_h, val_str, bg, COR_FG,
              bold=bold_label, fs=6.5)

# ── Planejamento Redução ─────────────────────────────────────────────────
y -= row_h
cellL(0, y, left_w, row_h,
      f"Planejamento Redução ({reducao['Pct_plena'].sum():.2f}% pleno)", COR_RED, COR_W, bold=True, fs=7.5)

# Larguras: data | DFT ID | %
PLAN_DATE_W = left_w * 0.28
PLAN_ID_W   = left_w * 0.46
PLAN_VAL_W  = left_w - PLAN_DATE_W - PLAN_ID_W

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
            cellL(PLAN_DATE_W + PLAN_ID_W, y, PLAN_VAL_W, row_h,
                  f"{r['Pct']:.2f}%".replace(".", ","), bg, COR_FG, bold=True, fs=7)
        else:
            cellL(PLAN_DATE_W + PLAN_ID_W, y, PLAN_VAL_W, row_h, "", bg)

# ── 11b. Gráfico de linha com projeção ──────────────────────────────────────

ax_chart.set_axis_on()
ax_chart.spines[["top", "right"]].set_visible(False)

meses_hist  = sorted(resumo.index.tolist())
vals_hist   = [resumo.loc[m, "Pct"] for m in meses_hist]
labels_hist = [MESES_LABEL[m] for m in meses_hist]

ax_chart.plot(labels_hist, vals_hist, color="black", linewidth=1.5,
              marker="o", markersize=4, label="% Falha Real")
for m, v in zip(labels_hist, vals_hist):
    ax_chart.annotate(f"{v:.2f}%".replace(".", ","), xy=(m, v), xytext=(0, 7),
                       textcoords="offset points", ha="center", fontsize=7)

# ── Linha de expectativa mês-a-mês ───────────────────────────────────────────
# Para cada mês M: mostra o que foi projetado para M com base no real de M-1
# Fórmula: expected[M] = actual[M-1] - Pct_mes(DFTs com milestone em M,
#           pesados pela distribuição de pedidos do mês M-1)

# Todos os pedidos com DFT e milestone preenchidos (qualquer mês)
_all_dft_ms = df[
    df["DefectNumber__c"].notna() &
    (df["DefectNumber__c"] > 0) &
    (df["DefectNumber__c"] != 999999) &
    df["DFT_BugfixMilestone"].notna()
].copy()
_ms_parsed = pd.to_datetime(_all_dft_ms["DFT_BugfixMilestone"], utc=True, errors="coerce")
_all_dft_ms = _all_dft_ms[_ms_parsed.notna()].copy()
_all_dft_ms["_MS_brt"] = _ms_parsed[_all_dft_ms.index].dt.tz_convert("America/Sao_Paulo").dt.normalize()
_all_dft_ms["_MS_mes"] = _all_dft_ms["_MS_brt"].dt.month

# Milestone por DFT (único por DFT)
_dft_to_ms = (
    _all_dft_ms[["DefectNumber__c", "_MS_brt", "_MS_mes"]]
    .drop_duplicates("DefectNumber__c")
    .set_index("DefectNumber__c")
)

# Para cada par (M-1, M) nos meses históricos, calcula a redução esperada para M
# usando a distribuição de pedidos por DFT no mês M-1
expect_labels = []
expect_vals   = []

for i in range(1, len(meses_hist)):
    m_ant = meses_hist[i - 1]   # mês base (de onde vem o fallout)
    m     = meses_hist[i]       # mês projetado

    # Pedidos por DFT no mês anterior
    _n_ant = (
        df[df["Mes"] == m_ant]
        .groupby("DefectNumber__c").size()
    )
    _total_ant_n = int(resumo.loc[m_ant, "Total"]) if m_ant in resumo.index else max(_n_ant.sum(), 1)

    # Redução esperada: DFTs cujo milestone cai em M
    red = 0.0
    for dft_id, row_ms in _dft_to_ms.iterrows():
        if row_ms["_MS_mes"] != m:
            continue
        n = int(_n_ant.get(dft_id, 0))
        if n == 0:
            continue
        pct_bruta = n / _total_ant_n * 100
        pr = _prorate(row_ms["_MS_brt"].to_pydatetime())
        red += pct_bruta * pr

    expect_labels.append(MESES_LABEL[m])
    expect_vals.append(round(max(vals_hist[i - 1] - red, 0), 2))


COR_EXPECT = "#E67E22"   # laranja

if len(expect_labels) > 0:
    # alinha com posições do eixo X (labels_hist + proj_labels[1:])
    ax_chart.plot(expect_labels, expect_vals, color=COR_EXPECT, linewidth=1.5,
                  linestyle="--", marker="^", markersize=5, markerfacecolor="white",
                  label="Expectativa mês anterior")
    for lbl, v in zip(expect_labels, expect_vals):
        ax_chart.annotate(f"{v:.2f}%".replace(".", ","), xy=(lbl, v), xytext=(0, -13),
                           textcoords="offset points", ha="center", fontsize=6, color=COR_EXPECT)

# ── Projeção futura ───────────────────────────────────────────────────────────
proj_months_num = sorted(set(reducao["MilestoneDate"].dt.month)) if len(reducao) else []
proj_months_num = [m for m in proj_months_num if m > mes_atual]
# +2: inclui todos os meses até o último milestone e mais um mês após (efeito pleno)
_last_ms_mes = proj_months_num[-1] if proj_months_num else mes_atual
todos_meses_proj = list(range(mes_atual + 1, min(_last_ms_mes + 2, 13))) if proj_months_num else []

proj_labels   = [labels_hist[-1]]
proj_vals     = [vals_hist[-1]]
cur_val       = vals_hist[-1]
# Milestones do mês atual já são plenos em agosto em diante
pct_acumulada = reducao.loc[reducao["MilestoneDate"].dt.month == mes_atual, "Pct_plena"].sum()

for m in todos_meses_proj:
    red_mes   = reducao.loc[reducao["MilestoneDate"].dt.month == m, "Pct_mes"].sum()
    red_plena = reducao.loc[reducao["MilestoneDate"].dt.month == m, "Pct_plena"].sum()
    cur_val   = max(vals_hist[-1] - pct_acumulada - red_mes, 0)
    proj_labels.append(MESES_LABEL[m])
    proj_vals.append(cur_val)
    pct_acumulada += red_plena

# ── Desenha linhas ────────────────────────────────────────────────────────────
if len(proj_labels) > 1:
    ax_chart.plot(proj_labels, proj_vals, color=COR_GRAY, linewidth=2,
                  linestyle="--", marker="o", markersize=6, markerfacecolor="white",
                  label="Projeção futura")
    ax_chart.annotate(f"{proj_vals[-1]:.2f}%".replace(".", ","),
                       xy=(proj_labels[-1], proj_vals[-1]), xytext=(8, 0),
                       textcoords="offset points", ha="left", va="center", fontsize=8)

# Linha de meta
todas_labels = labels_hist + proj_labels[1:]
ax_chart.axhline(y=1, color=COR_BLUE, linewidth=1.5, label="Meta (1%)")

ax_chart.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.0f}%"))
ax_chart.grid(axis="y", linestyle="--", alpha=0.3)
ax_chart.legend(fontsize=7, loc="upper right")
ax_chart.set_xticks(range(len(todas_labels)))
ax_chart.set_xticklabels(todas_labels, rotation=30, ha="right", fontsize=7)

data_corte = hoje.tz_convert("America/Sao_Paulo")
ax_chart.text(0.99, 1.04, f"Data de Corte: {data_corte.strftime('%d/%b')}",
              transform=ax_chart.transAxes, ha="right", fontsize=7.5, fontstyle="italic")

# ── 11c. Tabela detalhada (reaproveitando pivot2) ───────────────────────────

ax_table.set_xlim(0, sum(COL_W))
ax_table.set_ylim(0, fig_h)

y_t = fig_h - HDR_H
for ci, col in enumerate(COL_NAMES):
    x = X[ci]; w = sum(COL_W[ci:ci+1])
    ax_table.add_patch(plt.Rectangle((x, y_t), w, HDR_H, facecolor=COR_HDR, edgecolor="#FFFFFF", linewidth=0.5))
    ax_table.text(x + w/2, y_t + HDR_H/2, col, ha="center", va="center", fontsize=6.5,
                  color="#FFFFFF", fontweight="bold", multialignment="center")

def tcell(ci, y, h, text, bg, fg=COR_FG2, bold=False, fs=6, align="center", span=1):
    x = X[ci]; w = sum(COL_W[ci:ci + span])
    ax_table.add_patch(plt.Rectangle((x, y), w, h, facecolor=bg, edgecolor="#CCCCCC", linewidth=0.3))
    if not text: return
    ha = "left" if align == "left" else "center"
    xt = x + 0.1 if align == "left" else x + w / 2
    ax_table.text(xt, y + h / 2, str(text), ha=ha, va="center", fontsize=fs,
                  color=fg, fontweight="bold" if bold else "normal",
                  multialignment=ha, clip_on=True)

# ── Linha Sucesso ────────────────────────────────────────────────────────
y_t -= CELL_H
for ci in range(len(COL_NAMES) - len(meses_tab)):
    tcell(ci, y_t, CELL_H, "Sucesso" if ci == 0 else "", COR_W2, fs=7)
for li, m in enumerate(meses_tab):
    suc_pct = pct_m(resumo.loc[m, "Sucessos"] if m in resumo.index else 0, m)
    tcell(len(COL_NAMES) - len(meses_tab) + li, y_t, CELL_H,
          f"{suc_pct:.2f}%".replace(".", ","), COR_W2, fs=7)

# ── Linhas de DFT: "Falha / Em Tratamento" mesclado na primeira linha ───
pivot2_reset = pivot2.reset_index(drop=True)
for ri in range(len(pivot2_reset)):
    row = pivot2_reset.iloc[ri]
    y_t -= CELL_H
    bg = COR_W2 if ri % 2 == 0 else COR_G2

    tipo     = str(row["DFT_Type"]).strip() if pd.notna(row["DFT_Type"]) else ""
    tipo_fmt = "US" if tipo == "User Story" else "Defeito"
    id_str   = fmt_id({"DefectNumber__c": row["DefectNumber__c"],
                        "DFT_BugfixMilestone": row["DFT_BugfixMilestone"]})
    ms_fmt   = fmt_milestone(row["DFT_BugfixMilestone"])
    nome     = str(row["DFT_Name"])[:80] if pd.notna(row["DFT_Name"]) else ""
    team     = str(row["DFT_Team"])[:35] if pd.notna(row["DFT_Team"]) else ""
    phase    = str(row["DFT_Phase"])[:25] if pd.notna(row["DFT_Phase"]) else ""
    tem_data = bool(ms_fmt)
    team_bold = team == team.upper() and len(team) > 3  # em maiúsculas = negrito

    # col 0: Classificação — só na primeira linha
    tcell(0, y_t, CELL_H, "Falha" if ri == 0 else "", bg, fs=7)
    # col 1: Agrupamento — só na primeira linha
    tcell(1, y_t, CELL_H,
          "Em Tratamento/Avaliação\npela Squad" if ri == 0 else "",
          bg, fs=6.5)
    # col 2: Tipo
    tcell(2, y_t, CELL_H, tipo_fmt, bg, fs=7)
    # col 3: Defeito/US
    tcell(3, y_t, CELL_H, id_str, bg, fs=7)
    # col 4: Data Prod
    tcell(4, y_t, CELL_H, ms_fmt, bg, bold=tem_data, fs=7)
    # col 5: Nome Defeito
    tcell(5, y_t, CELL_H, nome, bg, fs=5.5, align="left")
    # col 6: Time
    tcell(6, y_t, CELL_H, team, bg, bold=team_bold, fs=6.5)
    # col 7: Status
    tcell(7, y_t, CELL_H, phase, bg, fs=6.5)
    # col 8+: % meses
    for li, m in enumerate(meses_tab):
        val = int(row.loc[m]) if m in pivot2_reset.columns else 0
        tcell(8 + li, y_t, CELL_H,
              f"{pct_m(val, m):.2f}%".replace(".", ","), bg, fs=7)

fig.subplots_adjust(left=0.01, right=0.99, top=0.97, bottom=0.01, hspace=0.18, wspace=0.12)
nome_arquivo = f"dashboard_{JORNADA.replace(' + ', '_').replace(' ', '_')}.png"
plt.savefig(nome_arquivo, dpi=150, bbox_inches="tight", pad_inches=0.1)
plt.close()
print(f"Dashboard consolidado salvo em {nome_arquivo}")
