# -*- coding: utf-8 -*-
"""
Converte a planilha de acompanhamento (blocos por jornada, lado a lado) para
`extracoes/historico.xlsx`, que é o formato que o sistema lê.

A planilha de origem é a fonte oficial dos meses fechados: as extrações antigas
não reproduzem aqueles números — o Octane continua mexendo em defeito velho e o
mês fechado mudava sozinho de uma semana para a outra.

As colunas Sucessos e Falhas não existem na origem: são opcionais e digitadas
direto no historico.xlsx. Reconverter a planilha preserva o que já estiver lá.

Da origem interessam duas colunas por jornada:
    "Realizado"                     -> Pct   (a linha preta do gráfico)
    "Novo Planejamento (Pro-Rata)"  -> ProRata (a linha laranja)
"Meta" fica de fora: a linha de 1% já é desenhada pelo próprio gráfico.

Em Base Residencial essas colunas foram renomeadas sem mudar de posição no
bloco: o planejamento virou "Planejamento s/ Legado Aberto" e o pro-rata virou
"Planejamento (agendados)". Os dois nomes de cada par são aceitos.

Uso:  python converter_historico.py "C:/caminho/planilha.csv" --ate 2026-06
"""
import argparse
import csv
import io
import os
import sys

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import pandas as pd

import fallout_core

RAIZ = os.path.dirname(os.path.abspath(__file__))

# Nome na planilha -> nome usado pelo sistema (pasta da jornada ou combinação).
# As jornadas de Assetização existem na planilha mas não neste relatório, e por
# isso ficam de fora em vez de virar linhas órfãs.
NOMES = {
    "Consolidado": "Consolidado",
    "Prospect": "Prospect",
    "Base Móvel": "Base Móvel",
    "Base Residencial": "NBA",
    "PME Móvel": "PME",
    "Cross Cell": "Cross Sell",
}
COL_REALIZADO = "Realizado"
COLS_PRORATA = ("Novo Planejamento (Pro-Rata)", "Planejamento (agendados)")
# A coluna do planejamento muda de nome conforme o bloco, e alguns preenchem
# uma enquanto outros preenchem outra: no Base Móvel de jul/26 há "Planejado" e
# o Pro-Rata está vazio; no Cross Cell é o contrário. A ordem aqui é a de
# prioridade — vale a primeira que estiver preenchida na linha.
COLS_PLANEJADO = ("Planejado", "Planejamento s/ LA", "Planejamento s/ Legado Aberto",
                  "Novo Planejamento (Pro-Rata)")

# A planilha não tem bloco de "Base + Cross Sell", que é a terceira coluna do
# slide consolidado. O acompanhamento usa o bloco de Base Móvel para essa
# coluna, então as linhas são replicadas com esse nome.
JORNADA_ESPELHADA = ("Base Móvel", "Base + Cross Sell")


def _numero(texto):
    """'9,51%' -> 9.51 ; célula vazia -> None."""
    limpo = (texto or "").strip().replace("%", "").replace(".", "").replace(",", ".")
    if not limpo:
        return None
    try:
        return float(limpo)
    except ValueError:
        return None


def _mes_ano(texto):
    """'ago/24' -> (2024, 8). Devolve None no que não for mês."""
    partes = (texto or "").strip().lower().split("/")
    if len(partes) != 2 or partes[0] not in fallout_core.ABREV_MES:
        return None
    try:
        ano = int(partes[1])
    except ValueError:
        return None
    return (2000 + ano if ano < 100 else ano), fallout_core.ABREV_MES[partes[0]]


def ler_planilha(caminho):
    """Devolve as linhas normalizadas da planilha de acompanhamento."""
    with io.open(caminho, encoding="cp1252", newline="") as fh:
        linhas = list(csv.reader(fh, delimiter=";"))
    if len(linhas) < 5:
        sys.exit(f"{caminho}: formato inesperado (menos de 5 linhas)")

    nomes, cabecalho = linhas[1], linhas[3]
    # Cada bloco começa na coluna cujo cabeçalho é "Mês"; o nome da jornada
    # está na linha de cima, na primeira coluna preenchida do bloco.
    inicios = [i for i, c in enumerate(cabecalho) if c.strip().lower() == "mês"]
    if not inicios:
        sys.exit(f"{caminho}: nenhuma coluna 'Mês' encontrada na linha 4")

    registros, ignoradas = [], []
    for k, ini in enumerate(inicios):
        fim = inicios[k + 1] if k + 1 < len(inicios) else len(cabecalho)
        nome = next((nomes[i].strip() for i in range(ini, min(fim, len(nomes)))
                     if nomes[i].strip()), "")
        if nome not in NOMES:
            if nome:
                ignoradas.append(nome)
            continue
        indices = {cabecalho[i].strip(): i for i in range(ini, fim) if cabecalho[i].strip()}
        for linha in linhas[4:]:
            if len(linha) <= ini:
                continue
            quando = _mes_ano(linha[ini])
            if quando is None:
                continue
            ano, mes = quando
            def _col(nome_col):
                i = indices.get(nome_col)
                return _numero(linha[i]) if i is not None and i < len(linha) else None
            planejado = next((v for v in (_col(c) for c in COLS_PLANEJADO)
                              if v is not None), None)
            prorata = next((v for v in (_col(c) for c in COLS_PRORATA)
                            if v is not None), None)
            registros.append({
                "Jornada": NOMES[nome], "Ano": ano, "Mes": mes,
                "Realizado": _col(COL_REALIZADO), "ProRata": prorata,
                "Planejado": planejado, "Sucessos": None, "Falhas": None,
            })

    espelho = [dict(r, Jornada=JORNADA_ESPELHADA[1])
               for r in registros if r["Jornada"] == JORNADA_ESPELHADA[0]]
    return registros + espelho, sorted(set(ignoradas))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("planilha", help="CSV exportado da planilha de acompanhamento")
    ap.add_argument("--ate", required=True,
                    help="último mês manual, AAAA-MM (ex.: 2026-06)")
    args = ap.parse_args()

    try:
        ano_corte, mes_corte = (int(p) for p in args.ate.split("-"))
    except ValueError:
        sys.exit(f"--ate inválido: {args.ate!r}. Use AAAA-MM, ex.: 2026-06")
    corte = ano_corte * 100 + mes_corte

    registros, ignoradas = ler_planilha(args.planilha)
    if ignoradas:
        print(f"  jornadas da planilha que não existem neste relatório, ignoradas:")
        for n in ignoradas:
            print(f"    {n}")

    # Os meses a partir do corte continuam vindo da extração: deles fica só o
    # "Planejado", que é uma meta e não tem como ser calculado do bruto. Manter
    # Realizado/ProRata aqui faria o valor digitado passar por cima do que a
    # extração apura para o mês corrente.
    for r in registros:
        if r["Ano"] * 100 + r["Mes"] > corte:
            r["Realizado"] = r["ProRata"] = None
    registros = [r for r in registros if any(
        r[c] is not None for c in ("Realizado", "ProRata", "Planejado"))]

    tabela = pd.DataFrame(registros, columns=fallout_core.COLS_HISTORICO)
    tabela = tabela.sort_values(["Jornada", "Ano", "Mes"]).reset_index(drop=True)

    destino = os.path.join(fallout_core.raiz_dados(RAIZ), fallout_core.ARQ_HISTORICO)

    # Sucessos e Falhas não vêm da planilha de acompanhamento: são digitados
    # direto no historico.xlsx e por isso sobreviveriam apenas até a próxima
    # conversão. Aqui eles são lidos do arquivo atual e recolocados, para
    # reconverter a planilha não apagar o que foi preenchido à mão.
    preservados = 0
    if os.path.exists(destino):
        try:
            atual = pd.read_excel(destino)
        except Exception as e:                                  # noqa: BLE001
            sys.exit(f"não consegui ler {destino} para preservar o volume: {e}")
        if {"Sucessos", "Falhas"} <= set(atual.columns):
            volumes = {}
            for _, linha in atual.iterrows():
                suc = pd.to_numeric(linha.get("Sucessos"), errors="coerce")
                fal = pd.to_numeric(linha.get("Falhas"), errors="coerce")
                if pd.isna(suc) and pd.isna(fal):
                    continue
                chave = (str(linha["Jornada"]).strip(),
                         pd.to_numeric(linha["Ano"], errors="coerce"),
                         pd.to_numeric(linha["Mes"], errors="coerce"))
                volumes[chave] = (suc, fal)
            for i, linha in tabela.iterrows():
                achado = volumes.get((linha["Jornada"], linha["Ano"], linha["Mes"]))
                if achado is None:
                    continue
                tabela.at[i, "Sucessos"], tabela.at[i, "Falhas"] = achado
                preservados += 1
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        tabela.to_excel(writer, index=False, sheet_name="historico")

    print(f"\n✓ {destino}")
    print(f"  {len(tabela)} linhas (meses fechados até {args.ate}; "
          f"depois disso, só a meta 'Planejado')")
    if preservados:
        print(f"  {preservados} linhas com volume preenchido à mão foram preservadas")
    for jornada, grupo in tabela.groupby("Jornada"):
        ini, fim = grupo.iloc[0], grupo.iloc[-1]
        com_prorata = grupo["ProRata"].notna().sum()
        print(f"    {fallout_core.nome_exibicao(jornada):18} "
              f"{fallout_core.MESES_PT[ini['Mes']]}/{str(ini['Ano'])[2:]} a "
              f"{fallout_core.MESES_PT[fim['Mes']]}/{str(fim['Ano'])[2:]}  "
              f"({len(grupo)} meses, {com_prorata} com pro-rata)")


if __name__ == "__main__":
    main()
