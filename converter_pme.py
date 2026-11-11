# -*- coding: utf-8 -*-
"""
Converte os exports da PME (XLSX, aba "Base") para CSV.

O app já faz isso sozinho: ao ler um XLSX da PME, o fallout_core grava o CSV ao
lado dele e das vezes seguintes lê o CSV. Este script serve para pagar essa
conversão antes — logo depois de largar um export novo na pasta —, em vez de
na primeira abertura do app, que é quando ela aconteceria.

    Radar Ofensores RPA V2.0 1 (32).xlsx  ->  Radar Ofensores RPA V2.0 1 (32)__base.csv

As regras do CSV são as do fallout_core (`converter_rpa`), para o arquivo sair
igual ao que o app gravaria: separado por ';', em UTF-8 e com a data de
modificação do XLSX. Aqui, diferente do app, são convertidos todos os exports da
pasta, e não só o mais recente.

Um XLSX já convertido é pulado.

Uso:  python converter_pme.py              (converte extracoes/PME)
      python converter_pme.py <pasta>
      python converter_pme.py --forcar     (refaz mesmo os já convertidos)
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
        destino = fc._espelho_rpa(arq)
        # Antes de olhar as abas: listá-las já abre o XLSX, que leva segundos.
        if not forcar and fc._espelho_em_dia(arq, destino):
            print(f"· {nome}: já convertido")
            continue
        if fc.RPA_ABA not in fc._abas(arq):
            print(f"· {nome}: sem a aba {fc.RPA_ABA!r}, ignorado")
            continue
        print(f"→ {nome}")
        t = time.time()
        if fc.converter_rpa(arq):
            gerados.append(destino)
            print(f"   -> {os.path.basename(destino)}  ({time.time() - t:.1f}s)")
    return gerados


def main():
    args = [a for a in sys.argv[1:] if a != "--forcar"]
    forcar = "--forcar" in sys.argv[1:]
    raiz = fc.raiz_dados(os.path.dirname(os.path.abspath(__file__)))
    pasta = args[0] if args else os.path.join(raiz, "PME")
    print(f"pasta: {pasta}\n")

    gerados = converter(pasta, forcar=forcar)
    if not gerados:
        return

    print(f"\n✓ {len(gerados)} arquivo(s) gerado(s)")
    print("  Os XLSX originais continuam na pasta e passam a ser ignorados;")
    print("  apague-os quando quiser, depois de conferir os números.")
    print("  Rode 'Atualizar dados' no app, ou reinicie, para o cache soltar.")


if __name__ == "__main__":
    main()
