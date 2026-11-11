# -*- coding: utf-8 -*-
"""
O que preenche cada categoria da Distribuição: os pedidos agrupados pelo
defeito ou pela mensagem de erro.

Uma regra só para o slide "Distribuição do Fallout detalhada"
(gerar_slide3col.py) e a aba Distribuição do app mostrarem a mesma coisa.
"""
import pandas as pd

import fallout_core

# Categorias sem defeito associado: o que as explica é a mensagem de erro. As
# demais — os subgrupos de Em Tratamento, Tratado, Resolvido, MOPs, BSIM e
# Legado Aberto — são explicadas pelo defeito.
POR_MENSAGEM = {
    fallout_core.CAT_INCIDENTE, "Falha Pontual", "Falta Associar ao Defeito/US",
    "Em avaliação - Outros times", "Falha no Processo Usuário",
    # Definida pelo próprio "Erro agrupado", que no Base Residencial é a mensagem.
    # Hoje é subgrupo do Legado Aberto (a aba Distribuição a abre por mensagem);
    # fica aqui pelas fotografias de antes de 15/10, em que era categoria.
    fallout_core.CAT_ERRO_LEGADO,
    # Base Residencial: sem defeito, definidas pelo número reservado e pelo
    # "Erro agrupado".
    fallout_core.CAT_SCRIPT_RPA, fallout_core.CAT_AGENDAMENTO,
}
# Os subgrupos do Em Tratamento (ver fallout_core.SUBGRUPOS_TRATAMENTO).
SUBGRUPOS = tuple(s for subs in fallout_core.SUBGRUPOS_TRATAMENTO.values() for s in subs)
LIMITE_MENSAGEM = 110
COLUNAS_DEFEITO = ["Defeito/US", "Nome", "Equipe", "Entrega", "Pedidos", "%"]
COLUNAS_MENSAGEM = ["Mensagem", "Pedidos", "%"]


def por_defeito(pedidos, total):
    """
    Um defeito por linha, do maior para o menor, com as colunas de
    COLUNAS_DEFEITO; o "%" é sobre os `total` pedidos do mês. O DFT com US de
    Melhoria aparece pelo número da US, como no Top Ofensores, e DFTs que
    apontam para a mesma US viram uma linha.
    """
    if not len(pedidos):
        return pd.DataFrame(columns=COLUNAS_DEFEITO)
    us = (pedidos["DFT_US"] if "DFT_US" in pedidos.columns
          else pd.Series(pd.NA, index=pedidos.index))
    grupos = (pedidos
              .assign(_chave=us.where(us.notna(), pedidos["DefectKey"]).astype(str),
                      _tipo=pedidos["DFT_Type"].where(us.isna(), "User Story"))
              .groupby("_chave", sort=False)
              .agg(Pedidos=("_chave", "size"), Nome=("DFT_Name", "first"),
                   Equipe=("DFT_Team", "first"), Entrega=("DFT_BugfixMilestone", "first"),
                   _tipo=("_tipo", "first"))
              .sort_values("Pedidos", ascending=False, kind="stable")
              .reset_index())
    # "DFT"/"US" só na frente de número: BRJ-*, INC-* e os rótulos ficam como vieram.
    grupos["Defeito/US"] = [
        (("US " if tipo == "User Story" else "DFT ") + chave) if chave.isdigit() else chave
        for chave, tipo in zip(grupos["_chave"], grupos["_tipo"])]
    grupos["Entrega"] = pd.to_datetime(grupos["Entrega"], errors="coerce")
    grupos["%"] = grupos["Pedidos"] / total * 100
    return grupos[COLUNAS_DEFEITO]


def por_mensagem(pedidos, total, limite=LIMITE_MENSAGEM):
    """
    Uma mensagem de erro por linha, da mais frequente para a menos, com as
    colunas de COLUNAS_MENSAGEM. Agrupa pela mensagem resumida, que é a que
    aparece na tela: duas variações que só mudam no trecho cortado viravam duas
    linhas iguais.
    """
    import comparacao                  # só pelo resumo da mensagem de erro

    if not len(pedidos) or "ErrorHandled__c" not in pedidos.columns:
        return pd.DataFrame(columns=COLUNAS_MENSAGEM)
    bruta = pedidos["ErrorHandled__c"].fillna("(sem mensagem de erro)").astype(str)
    resumo = {m: comparacao.resumir_mensagem(m, limite) for m in bruta.unique()}
    contagem = bruta.map(resumo).value_counts()
    return pd.DataFrame({"Mensagem": contagem.index, "Pedidos": contagem.values,
                         "%": contagem.values / total * 100})
