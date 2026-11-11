# -*- coding: utf-8 -*-
"""
Converte as planilhas do NBA (XLSX) para pares de CSV.

Ler XLSX com openpyxl é lento: os três meses levam cerca de 40 s, e a faixa de
KPIs do app carrega todas as jornadas. Em CSV a mesma leitura fica em poucos
segundos. Cada planilha vira dois arquivos, um por aba:

    jun26.xlsx  ->  jun26__analitico.csv
                    jun26__defeitos.csv

O app já faz isso sozinho: ao ler um XLSX do NBA, o fallout_core grava o par ao
lado dele e das vezes seguintes lê os CSVs. Este script serve para pagar essa
conversão antes, em vez de na primeira abertura do app. As regras do par são as
do fallout_core (`converter_nba`), inclusive a data de modificação herdada do
XLSX; aqui, diferente do app, são convertidas todas as planilhas da pasta.

Uma planilha já convertida é pulada.

Uso:  python converter_nba.py              (converte extracoes/NBA)
      python converter_nba.py <pasta>
      python converter_nba.py --forcar     (refaz mesmo as já convertidas)
"""
import os
import sys
import time

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import fallout_core as fc


def converter(pasta, forcar=False):
    arquivos = fc._arquivos(pasta, "*.xlsx")
    if not arquivos:
        print(f"Nenhum XLSX em {pasta}")
        return []

    gerados = []
    for arq in arquivos:
        nome = os.path.basename(arq)
        destinos = fc._espelhos_nba(arq)
        if not forcar and fc._espelho_em_dia(arq, *destinos):
            print(f"· {nome}: já convertido")
            continue
        print(f"→ {nome}")
        t = time.time()
        if fc.converter_nba(arq):
            gerados += destinos
            print(f"   -> {' + '.join(os.path.basename(d) for d in destinos)}  "
                  f"({time.time() - t:.1f}s)")
    return gerados


def main():
    args = [a for a in sys.argv[1:] if a != "--forcar"]
    forcar = "--forcar" in sys.argv[1:]
    raiz = fc.raiz_dados(os.path.dirname(os.path.abspath(__file__)))
    pasta = args[0] if args else os.path.join(raiz, "NBA")
    print(f"pasta: {pasta}\n")

    gerados = converter(pasta, forcar=forcar)
    if not gerados:
        return

    print(f"\n✓ {len(gerados)} arquivos gerados")
    print("  Os XLSX originais continuam na pasta e passam a ser ignorados;")
    print("  apague-os quando quiser, depois de conferir os números.")


if __name__ == "__main__":
    main()
