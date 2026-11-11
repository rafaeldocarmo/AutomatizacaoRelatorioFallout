# -*- coding: utf-8 -*-
"""
Põe as extrações novas no formato que o sistema lê e as coloca nas pastas.

O portal exporta em dois formatos que não são o que a extensão diz:

  falhas   ".xls" que na verdade é uma tabela HTML em ISO-8859-1
  sucessos ".xlsx" de verdade, com o cabeçalho do relatório e a quebra diária

Os dois viram CSV em `extracoes/<jornada>/`, mantendo o nome que já existe lá —
`mes_do_arquivo` deduz o mês pela abreviação e pelos dois dígitos do ano.

A data das falhas é convertida de dd/mm/aaaa para aaaa-mm-dd. Não é preferência
de formato: o leitor usa `pd.to_datetime` sem `dayfirst`, então "01/08/2026"
seria lido como 8 de janeiro. Todo dia até o 12 do mês entraria no mês errado,
sem erro nenhum aparecendo.

Uso:  python converter_extracao.py --mes ago26
      python converter_extracao.py --mes ago26 --so Base
"""
import argparse
import csv
import io
import os
import re
import shutil
import sys

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import pandas as pd

import fallout_core

RAIZ = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")

# jornada -> (padrão do arquivo de falhas, padrão do de sucessos,
#             nome do CSV de falhas, nome do CSV de sucessos)
# Os nomes de destino seguem o que cada pasta já usa — eles diferem entre as
# jornadas por herança, e renomear agora só criaria arquivo duplicado.
JORNADAS = {
    "Base Móvel": ("Falhas_Base", "Pedidos Sucesso  Base",
                   "RelatorioDeFalhas_{mes}.csv", "RelatorioDeSucessos_{mes}.csv"),
    "Cross Sell": ("Falhas_Cross Sell", "Pedidos em Sucesso_Cross Sell",
                   "{mes}.csv", "sucesso_{mes}.csv"),
    "Prospect": ("Falhas_PF", "Pedidos em Sucesso Prospepect PF",
                 "{mes}.csv", "{mes}.csv"),
}

COL_DATA = "Data de Criação"


def _achar(padrao):
    """Arquivo mais recente do Downloads que começa com `padrao`."""
    achados = [f for f in os.listdir(DOWNLOADS)
               if f.lower().startswith(padrao.lower())]
    if not achados:
        return None
    achados.sort(key=lambda f: os.path.getmtime(os.path.join(DOWNLOADS, f)),
                 reverse=True)
    return os.path.join(DOWNLOADS, achados[0])


def _sem_tags(html):
    return re.sub(r"<[^>]+>", "", html).replace("&amp;", "&").replace(
        "&lt;", "<").replace("&gt;", ">").replace("&nbsp;", " ").strip()


def _iso(data_br):
    """'25/08/2026' -> '2026-08-25'. Devolve o texto original se não casar."""
    achado = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", data_br.strip())
    if not achado:
        return data_br
    dia, mes, ano = achado.groups()
    return f"{ano}-{mes}-{dia}"


def converter_falhas(origem, destino):
    """Tabela HTML do portal -> CSV, com a data normalizada para ISO."""
    html = io.open(origem, encoding="ISO-8859-1").read()
    linhas_html = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S)
    if not linhas_html:
        raise ValueError(f"{origem}: nenhuma linha <tr> encontrada")

    cabecalho = [_sem_tags(c) for c in
                 re.findall(r"<th[^>]*>(.*?)</th>", linhas_html[0], re.S)]
    if COL_DATA not in cabecalho:
        raise ValueError(f"{origem}: sem a coluna {COL_DATA!r}. Veio: {cabecalho}")
    i_data = cabecalho.index(COL_DATA)

    linhas, convertidas = [], 0
    for bruto in linhas_html[1:]:
        celulas = [_sem_tags(c) for c in
                   re.findall(r"<td[^>]*>(.*?)</td>", bruto, re.S)]
        if not celulas:
            continue
        # A tabela pode vir com menos células que colunas na última linha.
        celulas += [""] * (len(cabecalho) - len(celulas))
        antes = celulas[i_data]
        celulas[i_data] = _iso(antes)
        convertidas += celulas[i_data] != antes
        linhas.append(celulas[:len(cabecalho)])

    with io.open(destino, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(cabecalho)
        w.writerows(linhas)
    return len(linhas), convertidas


def converter_sucessos(origem, destino):
    """XLSX do portal -> CSV, célula a célula (o leitor ignora as vazias)."""
    tabela = pd.read_excel(origem, header=None, dtype=str)
    dias = 0
    with io.open(destino, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        for _, linha in tabela.iterrows():
            celulas = ["" if pd.isna(v) else str(v).strip() for v in linha]
            cheias = [c for c in celulas if c]
            if len(cheias) == 2 and fallout_core.SUC_DATA_DIA.match(cheias[0]):
                dias += 1
            w.writerow(celulas)
    return len(tabela), dias


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mes", required=True,
                    help="mês do destino, no padrão dos arquivos (ex.: ago26)")
    ap.add_argument("--so", help="converte só a jornada cujo nome contenha isto")
    args = ap.parse_args()

    if fallout_core.mes_do_arquivo(f"x{args.mes}.csv") is None:
        sys.exit(f"--mes inválido: {args.mes!r}. Use abreviação+ano, ex.: ago26")

    raiz = fallout_core.raiz_dados(RAIZ)
    for jornada, (pad_f, pad_s, nome_f, nome_s) in JORNADAS.items():
        if args.so and args.so.lower() not in jornada.lower():
            continue
        print(f"\n{fallout_core.nome_exibicao(jornada)}")
        for padrao, subpasta, nome, funcao in (
            (pad_f, "falhas", nome_f, converter_falhas),
            (pad_s, "sucessos", nome_s, converter_sucessos),
        ):
            origem = _achar(padrao)
            if origem is None:
                print(f"  [pulado] nada em Downloads começando com {padrao!r}")
                continue
            destino = os.path.join(raiz, jornada, subpasta, nome.format(mes=args.mes))
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            total, extra = funcao(origem, destino)
            rotulo = "datas convertidas" if subpasta == "falhas" else "dias"
            print(f"  {subpasta:9} {os.path.basename(origem)[:44]:44} -> "
                  f"{os.path.basename(destino)}  ({total} linhas, {extra} {rotulo})")

    print("\nRode 'Atualizar dados' no app, ou reinicie, para o cache soltar.")


if __name__ == "__main__":
    main()
