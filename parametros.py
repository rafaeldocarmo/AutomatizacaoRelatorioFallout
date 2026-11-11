"""
Aba "Parâmetros": quais parâmetros (Parameters__c) já podem ser cancelados.

Cada parâmetro trata um erro de orquestração e aponta para um DFT e uma
jornada. Quando o DFT está Fechado ou Cancelado no Octane, o parâmetro perdeu o
motivo de existir: ou a correção já foi para produção, ou nenhuma vai chegar.

O relatório é o export do Salesforce, salvo em `extracoes/` (junto das
planilhas do Octane) com "Parametros" ou "Parameters" no nome — vale o mais
recente. Só a coluna do DFT é obrigatória: o RelatorioParametros traz Id,
Name, DFT e jornada, e o export completo de `Parameters__c` traz também ativo,
etapa, erro e ofensor — o que vier aparece na aba.
"""
import functools
import glob
import os
import re

import pandas as pd

import fallout_core

PADROES_ARQUIVO = ("*Parametros*", "*Parâmetros*", "*Parameters*")
EXTENSOES = (".csv", ".xlsx", ".xls")

# Fases do DFT no Octane que pedem o cancelamento do parâmetro. "Rejeitado"
# fica de fora de propósito, embora o resto do app o trate junto com
# "Cancelado" (ver fallout_core.FASES_MOPS): o parâmetro de DFT rejeitado segue.
FASES_CANCELAR = ("Fechado", "Cancelado")

STATUS_CANCELAR = "Cancelar"
STATUS_ABERTO = "Manter"
STATUS_US = "User Story (fora da regra)"
STATUS_NAO_ACHADO = "DFT não encontrado no Octane"
STATUS_SEM_DFT = "Sem DFT"
ORDEM_STATUS = (STATUS_CANCELAR, STATUS_ABERTO, STATUS_US, STATUS_NAO_ACHADO,
                STATUS_SEM_DFT)

# Coluna canônica -> nomes aceitos, em minúsculas. Comparados com o nome
# inteiro: "Name" é o nome do parâmetro, e "OrchestrationItemName__c" é a etapa
# da orquestração — sem o prefixo os dois virariam a mesma coisa.
COLUNAS = {
    "Id": ("id",),
    "Parâmetro": ("name", "parametro", "parâmetro"),
    "Ativo": ("active__c", "ativo"),
    "DFT": ("orchestrationitemdefectnumber__c", "defectnumber__c", "dft", "defeito"),
    "Jornada": ("orchestrationitemjourney__c", "journey__c", "jornada"),
    "Etapa": ("orchestrationitemname__c", "etapa"),
    "Erro": ("orchestrationitemerrorhandled__c", "errorhandled__c", "erro"),
    "Ofensor": ("orchestrationitemoffendername__c", "offendername__c", "ofensor"),
    "Criado em": ("createddate", "criado em"),
}


def localizar_relatorio(base_dir):
    """O relatório de parâmetros mais recente, ou None se não houver nenhum."""
    raiz = fallout_core.raiz_dados(base_dir)
    achados = set()
    for pasta in (raiz, base_dir):
        for padrao in PADROES_ARQUIVO:
            achados.update(
                c for c in glob.glob(os.path.join(pasta, padrao))
                if c.lower().endswith(EXTENSOES)
                and not os.path.basename(c).startswith(("~$", ".~")))
    return max(achados, key=os.path.getmtime) if achados else None


def _ler_bruto(caminho):
    if caminho.lower().endswith(".csv"):
        # O Data Loader grava em UTF-8 com BOM; o latin-1 cobre o CSV salvo
        # pelo Excel.
        try:
            return pd.read_csv(caminho, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(caminho, dtype=str, encoding="latin-1")
    return pd.read_excel(caminho, dtype=str)


def _padronizar(bruto):
    """Renomeia para as colunas canônicas; as que faltam entram vazias."""
    por_nome = {str(c).strip().lower(): c for c in bruto.columns}
    df = pd.DataFrame(index=bruto.index)
    for canon, aceitos in COLUNAS.items():
        origem = next((por_nome[a] for a in aceitos if a in por_nome), None)
        df[canon] = bruto[origem].str.strip() if origem else pd.NA
    if df["DFT"].isna().all():
        raise ValueError("O relatório de parâmetros não tem a coluna do DFT "
                         "(ex.: OrchestrationItemDefectNumber__c).")
    return df


def _id_dft(valor):
    """Número do DFT, ou None quando o campo não traz um (999999, PDST-, texto)."""
    if pd.isna(valor):
        return None
    txt = re.sub(r"^DFT\s*", "", str(valor).strip(), flags=re.I)
    txt = txt[:-2] if txt.endswith(".0") else txt
    if not txt.isdigit() or txt == "999999":
        return None
    return int(txt)


@functools.lru_cache(maxsize=2)
def _octane_bruto(caminho_dft, _mtime_dft, caminho_us, _mtime_us):
    """
    Fase do próprio DFT, e não a da US de Melhoria que o catálogo do
    `fallout_core` põe no lugar: a pergunta aqui é se o defeito fechou.
    """
    dft = pd.read_excel(caminho_dft)
    dft = pd.DataFrame({
        "ID": pd.to_numeric(dft["ID"], errors="coerce"),
        "Nome do DFT": dft["Name"],
        "Fase no Octane": dft["Phase"].astype(str).str.strip(),
        "Encerrado em": pd.to_datetime(dft.get("Done date"), errors="coerce"),
        "Time": dft.get("Team"),
    }).dropna(subset=["ID"]).drop_duplicates("ID").astype({"ID": int}).set_index("ID")
    us = pd.read_excel(caminho_us)
    ids_us = set(pd.to_numeric(us["ID"], errors="coerce").dropna().astype(int))
    return dft, frozenset(ids_us)


def _octane(base_dir):
    raiz = fallout_core.raiz_dados(base_dir)
    c_dft = fallout_core._caminho_octane(raiz, base_dir, fallout_core.ARQ_DFT)
    c_us = fallout_core._caminho_octane(raiz, base_dir, fallout_core.ARQ_US)
    return _octane_bruto(c_dft, os.path.getmtime(c_dft), c_us, os.path.getmtime(c_us))


def avaliar(caminho, base_dir):
    """
    Todos os parâmetros do relatório, com o DFT cruzado contra o Octane e a
    coluna "Status". Devolve (tabela, número de inativos que ficaram de fora);
    sem a coluna de ativo no relatório, todos contam e o número vem None.
    """
    df = _padronizar(_ler_bruto(caminho))
    if df["Ativo"].notna().any():
        ativo = df["Ativo"].str.lower().isin(["true", "1", "sim", "verdadeiro"])
        inativos = int((~ativo).sum())
        df = df[ativo]
    else:
        inativos = None
    df = df.drop(columns="Ativo").copy()

    dft, ids_us = _octane(base_dir)
    df["_id"] = df["DFT"].map(_id_dft)
    df = df.join(dft, on="_id")
    df["Criado em"] = pd.to_datetime(df["Criado em"], errors="coerce", dayfirst=True)

    tem_id = df["_id"].notna()
    no_dft = df["Fase no Octane"].notna()
    df["Status"] = STATUS_SEM_DFT
    df.loc[tem_id, "Status"] = STATUS_NAO_ACHADO
    df.loc[tem_id & ~no_dft & df["_id"].isin(ids_us), "Status"] = STATUS_US
    df.loc[no_dft, "Status"] = STATUS_ABERTO
    df.loc[df["Fase no Octane"].isin(FASES_CANCELAR), "Status"] = STATUS_CANCELAR

    df["Status"] = pd.Categorical(df["Status"], ORDEM_STATUS, ordered=True)
    df["Jornada"] = df["Jornada"].fillna("(sem jornada)")
    df = df.sort_values(["Status", "Jornada", "_id"], kind="stable").drop(columns="_id")
    return df.reset_index(drop=True), inativos
