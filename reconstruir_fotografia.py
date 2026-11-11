# -*- coding: utf-8 -*-
"""
Reconstrói a fotografia de uma extração passada com os arquivos que continuam
em `extracoes/`.

As pastas de jornada acumulam exports: o da semana passada fica ao lado do
novo, e o leitor escolhe o mais recente. Filtrando pela data de modificação,
dá para remontar a pasta como ela estava antes de uma atualização e
fotografar aquele estado — é assim que a primeira comparação sai sem esperar
uma semana pela próxima extração.

A pasta remontada é temporária, feita de links para os arquivos originais
(cópia, se o link não der), e nada em `extracoes/` é alterado além da
fotografia gravada no fim.

Limites, que ficam registrados na própria fotografia:
  * as planilhas da raiz (Octane, Jira, histórico) são uma só, substituída a
    cada atualização: se a atual é posterior à data pedida, ela entra assim
    mesmo, com aviso;
  * do histórico entram só os meses anteriores ao da data pedida — os que já
    estavam fechados. Com ele inteiro, a fotografia de 10/08 mostraria agosto
    com a taxa oficial do mês fechado, quando agosto tinha dez dias. A meta
    "Planejado" dos meses seguintes continua;
  * arquivo de jornada sobrescrito com o mesmo nome não volta — só o que ainda
    existe entra.

Ferramenta de quem mantém os dados, fora do pacote portátil, como os
converter_*.py.

Uso:  python reconstruir_fotografia.py --antes-de 2026-09-17 [--hoje 2026-09-11]
                                       [--pasta dist/ExploradorDeFallout] [--sobrescrever]

  --antes-de     entram só os arquivos de jornada modificados antes deste
                 momento: um dia (2026-09-17) ou dia e hora (2026-09-08T16:50),
                 para separar dois lotes que chegaram no mesmo dia
  --pasta        a pasta do app cujos dados remontar e onde gravar a fotografia
                 (a que tem a "extracoes"); padrão: a deste script. Serve para
                 fotografar os dados de uma cópia portátil
  --hoje         data de referência das regras; padrão: o dia do arquivo mais
                 novo que entrou, que é quando aquela extração ficou completa
  --sobrescrever substitui uma fotografia que já exista com a mesma data de
                 corte (sem ele, a reconstrução para — ela nunca passa por cima
                 de uma fotografia tirada na época sem você pedir)
"""
import sys

# No Windows o console usa cp1252 e quebra nos prints com acento e "→".
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import argparse
import datetime as _dt
import os
import shutil
import tempfile

import fallout_core
import fotografia

# As planilhas da raiz que o app lê. O resto da raiz (o historico.xlsx.bak,
# por exemplo) não é dado.
ARQUIVOS_DA_RAIZ = (fallout_core.ARQ_DFT, fallout_core.ARQ_US,
                    fallout_core.ARQ_JIRA, fallout_core.ARQ_HISTORICO)


def _ligar(origem, destino):
    """Link para o arquivo original; cópia quando o link não der (outro disco)."""
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    try:
        os.link(origem, destino)
    except OSError:
        shutil.copy2(origem, destino)       # copy2 preserva a data de modificação


def _data(caminho):
    return _dt.datetime.fromtimestamp(os.path.getmtime(caminho))


# Colunas do histórico que dizem que o mês fechou. "Planejado" fica de fora: é
# meta, e existe para mês que ainda nem começou.
_COLS_FECHAMENTO = ("Realizado", "ProRata", "Sucessos", "Falhas")


def _historico_ate(origem, destino, limite):
    """
    Grava em `destino` o histórico de `origem` sem o fechamento dos meses a
    partir do de `limite`: nesses, a fotografia usa a extração, como o app
    fazia na época. Devolve quantas linhas perderam o fechamento.
    """
    import pandas as pd

    tabela = pd.read_excel(origem)
    chave = (pd.to_numeric(tabela["Ano"], errors="coerce") * 100
             + pd.to_numeric(tabela["Mes"], errors="coerce"))
    colunas = [c for c in _COLS_FECHAMENTO if c in tabela.columns]
    abertos = (chave >= limite.year * 100 + limite.month) & tabela[colunas].notna().any(axis=1)
    tabela.loc[abertos, colunas] = None
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    tabela.to_excel(destino, index=False)
    return int(abertos.sum())


def montar_pasta(base_dir, destino, limite):
    """
    Remonta em `destino` a pasta de dados como ela estava antes de `limite`.

    Devolve (arquivos que entraram, avisos, data do arquivo de jornada mais
    novo). Os leitores escolhem o export pela data de modificação, e os links
    e o `copy2` a preservam — então a escolha na pasta remontada é a mesma
    que o app fez na época.
    """
    raiz = fallout_core.raiz_dados(base_dir)
    entraram, avisos, mais_novo = 0, [], None

    for nome in ARQUIVOS_DA_RAIZ:
        # O Octane e o Jira também podem estar na raiz do projeto (ver
        # `fallout_core._caminho_octane`); o histórico, só junto das jornadas.
        achado = next((os.path.join(p, nome) for p in (raiz, base_dir)
                       if os.path.isfile(os.path.join(p, nome))), None)
        if achado is None:
            continue
        if nome == fallout_core.ARQ_HISTORICO:
            removidas = _historico_ate(achado, os.path.join(destino, nome), limite)
            if removidas:
                avisos.append(
                    f"{nome}: entraram só os meses fechados antes de "
                    f"{fallout_core.MESES_LABEL[limite.year * 100 + limite.month]}; "
                    f"desse mês em diante ({removidas} linhas) a taxa vem da extração, "
                    f"como na época.")
        else:
            if _data(achado) >= limite:
                avisos.append(f"{nome} é de {_data(achado):%d/%m/%Y}, posterior à data "
                              f"pedida, e entrou assim mesmo: não há versão anterior.")
            _ligar(achado, os.path.join(destino, nome))
        entraram += 1

    for j in fallout_core.jornadas_disponiveis(base_dir):
        for pasta, _subdirs, nomes in os.walk(os.path.join(raiz, j)):
            for nome in nomes:
                origem = os.path.join(pasta, nome)
                if nome.startswith(("~$", ".~")) or nome.endswith(".parcial"):
                    continue
                quando = _data(origem)
                if quando >= limite:
                    continue
                _ligar(origem, os.path.join(destino, os.path.relpath(origem, raiz)))
                entraram += 1
                mais_novo = quando if mais_novo is None else max(mais_novo, quando)
    return entraram, avisos, mais_novo


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Reconstrói a fotografia de uma extração passada.")
    parser.add_argument("--antes-de", required=True,
                        help="AAAA-MM-DD ou AAAA-MM-DDTHH:MM: entram só os arquivos de "
                             "jornada modificados antes deste momento")
    parser.add_argument("--hoje", help="data de referência das regras (AAAA-MM-DD)")
    parser.add_argument("--pasta", help="pasta do app com a extracoes a remontar e onde "
                                        "gravar; padrão: a deste script")
    parser.add_argument("--sobrescrever", action="store_true",
                        help="substitui a fotografia que já exista com a mesma data de corte")
    args = parser.parse_args(argv)

    base_dir = os.path.abspath(args.pasta or os.path.dirname(os.path.abspath(__file__)))
    if not os.path.isdir(os.path.join(base_dir, fallout_core.PASTA_DADOS)):
        print(f"Não há pasta {fallout_core.PASTA_DADOS} em {base_dir}.")
        return 1
    # Só a data vira meia-noite; com hora, o corte é naquele minuto.
    limite = _dt.datetime.fromisoformat(args.antes_de)
    quando = f"{limite:%d/%m/%Y}" + ("" if limite.time() == _dt.time() else f" {limite:%H:%M}")

    temp = tempfile.mkdtemp(prefix="fotografia-")
    try:
        pasta_dados = os.path.join(temp, fallout_core.PASTA_DADOS)
        entraram, avisos, mais_novo = montar_pasta(base_dir, pasta_dados, limite)
        if mais_novo is None:
            print(f"Nenhum arquivo de jornada anterior a {quando}.")
            return 1
        hoje = args.hoje or mais_novo.date().isoformat()
        print(f"Pasta remontada com {entraram} arquivos anteriores a "
              f"{quando} (o mais novo é de {mais_novo:%d/%m/%Y %H:%M}).")

        foto = fotografia.gerar(
            temp, hoje=hoje,
            origem=f"reconstruída com os arquivos anteriores a {quando}",
            progresso=lambda feitos, total, rotulo:
                print(f"  [{feitos + 1}/{total}] {rotulo}"))
        foto.meta["avisos"] += avisos

        ja_existe = any(f["id"] == foto.id for f in fotografia.listar(base_dir))
        if ja_existe and not args.sobrescrever:
            print(f"\nJá existe uma fotografia com dados até "
                  f"{_dt.date.fromisoformat(foto.id):%d/%m/%Y}. Nada foi gravado; "
                  f"use --sobrescrever para substituí-la pela reconstrução.")
            return 1
        destino = fotografia.salvar(foto, base_dir)
        fotografia.imprimir_resumo(foto, destino)
        return 0
    finally:
        shutil.rmtree(temp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
