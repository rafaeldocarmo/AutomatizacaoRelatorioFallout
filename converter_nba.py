# -*- coding: utf-8 -*-
"""
Converte as planilhas do NBA (XLSX) para pares de CSV.

Ler XLSX com openpyxl é lento: os três meses levam cerca de 40 s, e a faixa de
KPIs do app carrega todas as jornadas. Em CSV a mesma leitura fica em poucos
segundos. Cada planilha vira dois arquivos, um por aba:

    jun26.xlsx  ->  jun26__analitico.csv
                    jun26__defeitos.csv

O fallout_core prefere os CSVs quando existem e continua lendo direto qualquer
XLSX ainda não convertido, então dá para largar um arquivo novo na pasta e rodar
este script depois.

Uso:  python converter_nba.py            (converte extracoes/NBA)
      python converter_nba.py <pasta>
"""
import glob
import os
import sys
import time

import pandas as pd

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import fallout_core as fc


def converter(pasta):
    arquivos = sorted(glob.glob(os.path.join(pasta, "*.xlsx")))
    if not arquivos:
        print(f"Nenhum XLSX em {pasta}")
        return []

    gerados = []
    for arq in arquivos:
        base = os.path.splitext(arq)[0]
        nome = os.path.basename(arq)
        print(f"→ {nome}")
        t = time.time()
        for aba, sufixo in ((fc.NBA_ABA_ANALITICO, fc.NBA_SUF_ANALITICO),
                            (fc.NBA_ABA_DEFEITOS, fc.NBA_SUF_DEFEITOS)):
            df = pd.read_excel(arq, sheet_name=aba, dtype=str)
            destino = base + sufixo
            # utf-8-sig para o Excel abrir os acentos direito num duplo clique.
            df.to_csv(destino, index=False, encoding="utf-8-sig")
            gerados.append(destino)
            print(f"   {aba:10s} {len(df):>7,} linhas  ->  {os.path.basename(destino)}")
        print(f"   ({time.time() - t:.1f}s)")
    return gerados


def main():
    raiz = fc.raiz_dados(os.path.dirname(os.path.abspath(__file__)))
    pasta = sys.argv[1] if len(sys.argv) > 1 else os.path.join(raiz, "NBA")
    print(f"pasta: {pasta}\n")

    gerados = converter(pasta)
    if not gerados:
        return

    print(f"\n✓ {len(gerados)} arquivos gerados")
    print("  Os XLSX originais continuam na pasta e passam a ser ignorados;")
    print("  apague-os quando quiser, depois de conferir os números.")


if __name__ == "__main__":
    main()
