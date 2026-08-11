# -*- coding: utf-8 -*-
"""
Núcleo de dados compartilhado entre app.py (dashboard interativo) e pipeline.py
(geração de relatório/PNG/PPTX): leitura das bases, join com Octane e
categorização do fallout por mês.

Este módulo não depende de Streamlit nem de estado global — todas as funções
recebem `base_dir` explicitamente e retornam dados, para poder ser chamado
com segurança a partir de múltiplas sessões/threads simultâneas.
"""
import glob
import os

import pandas as pd

ABREV_MES = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
             "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
MESES_PT = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
            7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
MESES_LABEL = {m: f"{abrev}-26" for abrev, m in ABREV_MES.items()}

FASES_CORRIGIDO = {"Corrigido", "Fechado"}
FASES_MOPS = {"Cancelado", "Rejeitado"}

COLUNAS_OCTANE_BASE = ["ID", "Name", "Phase", "Bugfix Milestone", "Team", "Type"]
COL_US_MELHORIA = "US de Melhoria"

RENAME_FALHAS = {
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.OrderNumber": "OrderNumber",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Channel__c": "Channel",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Segment__c": "Segment",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.Status": "OrderStatus",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.SubStatus__c": "SubStatus",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.BiometryStatus__c": "BiometryStatus",
    "vlocity_cmt__State__c": "State",
    "vlocity_cmt__OrchestrationPlanId__r.vlocity_cmt__OrderId__r.vlocity_cmt__Reason__c": "Reason",
}


def mes_do_arquivo(path):
    """Deduz o número do mês (1-12) a partir do nome do arquivo (ex.: '...jan26.csv')."""
    nome = os.path.splitext(os.path.basename(path))[0].lower()
    for abrev, num in ABREV_MES.items():
        if abrev in nome:
            return num
    return None


def jornadas_disponiveis(base_dir):
    """Pastas de jornada existentes em `base_dir/extrações` (as que têm subpasta 'falhas')."""
    raiz = os.path.join(base_dir, "extrações")
    if not os.path.isdir(raiz):
        return []
    return sorted(
        d for d in os.listdir(raiz)
        if os.path.isdir(os.path.join(raiz, d, "falhas"))
    )


def resolver_jornadas(jornada, base_dir):
    """Traduz o nome escolhido nas pastas que devem ser lidas."""
    if jornada == "Consolidado":
        return jornadas_disponiveis(base_dir)
    if jornada == "Base + Cross Sell":
        return ["Base Móvel", "Cross Sell"]
    return [jornada]


def _norm_id(v):
    """Normaliza um ID que pode vir como float ('182366.0') para string inteira."""
    try:
        return str(int(float(str(v).strip())))
    except (TypeError, ValueError):
        return str(v).strip()


def carregar_base(base_dir, jornada):
    """
    Lê falhas + sucessos (CSV) e Octane (Excel) para a jornada informada,
    faz o join e devolve (df, df_octane, resumo, jornadas_combo).

    `jornada` pode ser um nome de pasta única, "Base + Cross Sell" ou
    "Consolidado" — a resolução para as pastas reais é feita aqui.
    """
    jornadas_combo = resolver_jornadas(jornada, base_dir)
    if not jornadas_combo:
        raise FileNotFoundError("Nenhuma pasta de jornada encontrada em extrações/")

    # ── Falhas ───────────────────────────────────────────────────────────
    arquivos_falhas = []
    for j in jornadas_combo:
        arqs = sorted(glob.glob(os.path.join(base_dir, "extrações", j, "falhas", "*.csv")))
        if not arqs:
            raise FileNotFoundError(f"Nenhum arquivo encontrado em extrações/{j}/falhas/")
        arquivos_falhas += arqs

    df_falhas = pd.concat([pd.read_csv(f) for f in arquivos_falhas], ignore_index=True)
    df_falhas["CreatedDate"] = pd.to_datetime(df_falhas["CreatedDate"], utc=True, errors="coerce")
    df_falhas = df_falhas.rename(columns=RENAME_FALHAS)

    # O export do CRM às vezes traz o número do defeito com espaços não-quebráveis
    # no fim (ex.: "233538\xa0\xa0"). Sem limpar, pd.to_numeric devolve NaN e o
    # pedido acaba contado como "Falta Associar ao Defeito/US" mesmo tendo DFT.
    df_falhas["DefectNumber_orig"] = (
        df_falhas["DefectNumber__c"].astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )
    # Alguns registros vêm com o prefixo do próprio sistema ("DFT 232143"). Só o
    # prefixo DFT é removido: INC-, DDP-, PDST- e OS referenciam outros sistemas,
    # não existem no Octane e devem continuar sem defeito associado.
    df_falhas["DefectNumber__c"] = (
        pd.to_numeric(
            df_falhas["DefectNumber_orig"].str.replace(
                r"^DFT\s*(?=\d)", "", regex=True, case=False),
            errors="coerce",
        )
        .fillna(-1).astype(int)
    )

    # ── Octane: DFTs + US, com resolução de "US de Melhoria" ───────────────
    df_dft = pd.read_excel(os.path.join(base_dir, "RelatorioDFTOctane.xlsx"))
    df_us = pd.read_excel(os.path.join(base_dir, "RelatorioUSOctane.xlsx"))

    df_dft_prep = df_dft[
        COLUNAS_OCTANE_BASE + ([COL_US_MELHORIA] if COL_US_MELHORIA in df_dft.columns else [])
    ].copy()

    if COL_US_MELHORIA in df_dft_prep.columns:
        df_us_idx = df_us.set_index("ID")[
            ["Name", "Phase", "Bugfix Milestone", "Team", "Type"]
        ].copy()
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
        pd.concat([df_dft_prep[COLUNAS_OCTANE_BASE], df_us[COLUNAS_OCTANE_BASE]], ignore_index=True)
        .drop_duplicates(subset="ID")
        .rename(columns={
            "ID": "DefectNumber__c", "Name": "DFT_Name", "Phase": "DFT_Phase",
            "Bugfix Milestone": "DFT_BugfixMilestone", "Team": "DFT_Team", "Type": "DFT_Type",
        })
    )
    df_octane["DefectNumber__c"] = (
        pd.to_numeric(df_octane["DefectNumber__c"], errors="coerce")
        .fillna(-1).astype(int)
    )

    # ── Join falhas ← Octane ────────────────────────────────────────────
    df = df_falhas.merge(df_octane, on="DefectNumber__c", how="left")

    # Desduplicar por OrderNumber mantendo a linha mais informativa:
    #   1º) DFT real (número > 0);
    #   2º) no empate entre linhas sem DFT (todas viram -1), vence a que tem
    #       rótulo explícito ("Enviado e-mail - Outros Times", "Erro de
    #       Processo", "INC-…") em vez da linha em branco.
    # `kind="stable"` evita que o desempate mude de execução para execução — o
    # quicksort padrão do pandas reordenava empates de forma arbitrária.
    _sem_rotulo = df["DefectNumber_orig"].isin(["", "nan", "None", "-1"])
    df = (
        df.assign(_tem_rotulo=(~_sem_rotulo).astype(int))
          .sort_values(["DefectNumber__c", "_tem_rotulo"],
                       ascending=[False, False], kind="stable")
          .drop_duplicates(subset=["OrderNumber"])
          .drop(columns="_tem_rotulo")
          .reset_index(drop=True)
    )
    df["Mes"] = df["CreatedDate"].dt.tz_convert("America/Sao_Paulo").dt.month

    # ── Sucessos por mês (mesmo denominador do fallout rate) ────────────
    arquivos_suc = []
    for j in jornadas_combo:
        arqs_suc = sorted(glob.glob(os.path.join(base_dir, "extrações", j, "sucessos", "*.csv")))
        if not arqs_suc:
            raise FileNotFoundError(f"Nenhum arquivo encontrado em extrações/{j}/sucessos/")
        arquivos_suc += arqs_suc

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
    resumo["Pct"] = resumo["Falhas"] / resumo["Total"] * 100

    return df, df_octane, resumo, jornadas_combo


def categorizar(df, mes, hoje=None):
    """
    Divide os registros de falha de um mês nas categorias de negócio do fallout.
    Retorna (df_mes, cats, total), onde `cats` é um dict {nome_categoria: DataFrame}.

    `hoje` pode ser fixado (ex.: para reprodutibilidade em testes); por padrão
    usa a data/hora atual normalizada para meia-noite UTC.
    """
    if hoje is None:
        hoje = pd.Timestamp.now(tz="UTC").normalize()
    quinze_dias = hoje - pd.Timedelta(days=15)
    df_mes = df[df["Mes"] == mes].copy()

    milestone_dt = pd.to_datetime(df_mes["DFT_BugfixMilestone"], utc=True, errors="coerce")
    _corrigido = df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_CORRIGIDO)
    _tem_ms = df_mes["DFT_BugfixMilestone"].notna()
    _encerrado = _corrigido & _tem_ms & (milestone_dt <= hoje)
    _outros_mask = df_mes["DefectNumber_orig"].str.strip().str.lower() == "enviado e-mail - outros times"
    _erro_proc_mask = df_mes["DefectNumber_orig"].str.strip().str.lower() == "erro de processo"

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
            ~_outros_mask & ~_erro_proc_mask
        ],
        "Tratado - Em avaliação de eficácia": df_mes[
            _corrigido & _tem_ms & (milestone_dt >= quinze_dias) & (milestone_dt <= hoje)
        ],
        "Em Avaliação por MOPs": df_mes[
            (df_mes["DefectNumber__c"] > 0) &
            (df_mes["DefectNumber__c"] != 999999) &
            df_mes["DFT_Phase"].fillna("").str.strip().isin(FASES_MOPS)
        ],
        "Em avaliação - Outros times": df_mes[_outros_mask],
        "Falha no Processo Usuário": df_mes[_erro_proc_mask],
    }
    total = len(df_mes)
    return df_mes, cats, total
