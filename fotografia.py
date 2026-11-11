# -*- coding: utf-8 -*-
"""
Fotografias das extrações: o resultado processado de uma extração, gravado em
disco para ser comparado com o de outra semana.

O app recalcula tudo a partir do que está em `extracoes/`, e o Octane é um
arquivo só, substituído a cada atualização — então a extração da semana
passada deixa de existir assim que a nova chega. A fotografia guarda o
**resultado**, e não os arquivos brutos, porque é ele que foi reportado: as
categorias dependem do Octane da época, do dia em que foram calculadas (os 15
dias da avaliação de eficácia) e das regras daquele momento, e nada disso se
refaz a partir dos arquivos antigos.

Cada fotografia é uma pasta em `extracoes/_fotografias/`, com o nome da data de
corte (ver `fallout_core.data_corte`):

    meta.json          quando e de onde: corte, "hoje", arquivos de entrada,
                       assinatura das regras e os números de cada relatório
    pedidos.parquet    pedido a pedido, de cada visão, com a categoria
    resumo.parquet     falhas, sucessos e taxa por mês, de cada visão
    reducao.parquet    o Planejamento de Redução da projeção, defeito a defeito
    ofensores.parquet  a tabela de ofensores do slide, um mês por linha

Fica dentro de `extracoes/` de propósito: tem dado pedido a pedido e segue a
mesma regra das extrações — não viaja no pacote portátil e não vai para o git.
A descoberta de jornadas pula a pasta (ver `fallout_core.PREFIXO_INTERNO`).

A gravação é manual: o botão "Gravar fotografia" do app ou a linha de comando.

    python fotografia.py                     grava a fotografia dos dados atuais
    python fotografia.py --hoje 2026-09-11   idem, com outra data de referência
    python fotografia.py --listar            lista as fotografias gravadas
"""
import sys

# No Windows o console usa cp1252 e quebra nos prints com acento e "→".
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import contextlib
import dataclasses
import datetime as _dt
import hashlib
import io
import json
import math
import os
import re
import shutil
import threading

import pandas as pd

import fallout_core

PASTA_FOTOS = fallout_core.PREFIXO_INTERNO + "fotografias"
ARQ_META = "meta.json"
TABELAS = ("pedidos", "resumo", "reducao", "ofensores")
# Sobe quando o conteúdo mudar de um jeito que um leitor antigo não entenda.
VERSAO_FORMATO = 1
RE_ID = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# Colunas do pedido a pedido: as que o app mostra (COLS_EXIBIR) mais tudo o que
# `categorizar` lê. Com estas, uma fotografia antiga pode ser recategorizada
# pela regra nova sem voltar aos arquivos — é o que separa, na comparação, a
# mudança de dado da mudança de regra.
COLS_PEDIDO = [
    "Jornada", "OrderNumber", "CreatedDate", "Mes", "ErrorHandled__c",
    "DefectKey", "DefectNumber_orig", "DefectNumber__c", "TemDefeito", "LegadoAberto",
    "ErroLegado", "DFT_Name", "DFT_Phase", "DFT_BugfixMilestone", "DFT_Team", "DFT_Type", "DFT_US",
    "Channel", "Segment", "State", "SubStatus", "Reason",
]
# Colunas da tabela de ofensores (pivot2 do pipeline) que descrevem o defeito.
COLS_OFENSOR = ["DefectKey", "DFT_Name", "DFT_Phase", "DFT_BugfixMilestone",
                "DFT_Type", "DFT_Team"]
# A data de entrega chega misturada — data do Octane, texto de uma planilha,
# vazio — e no Parquet a coluna precisa ter um tipo só.
COLS_DATA = ("DFT_BugfixMilestone",)
# A categoria cujos pedidos ganham subgrupo (US c/ data, US s/ data, DFT c/ data e
# DFT s/ data; até 11/11/2026, Planejado / US s/ data / DFT s/ data).
CAT_TRATAMENTO = fallout_core.CAT_TRATAMENTO


@dataclasses.dataclass
class Fotografia:
    """Uma extração fotografada: o meta.json e as quatro tabelas."""
    meta: dict
    pedidos: pd.DataFrame
    resumo: pd.DataFrame
    reducao: pd.DataFrame
    ofensores: pd.DataFrame

    @property
    def id(self):
        """A data de corte em AAAA-MM-DD, que é também o nome da pasta."""
        return self.meta["corte"]


def pasta_fotografias(base_dir):
    """
    `extracoes/_fotografias` do projeto.

    Direto em `extracoes/`, e não na raiz das jornadas (`raiz_dados`), que pode
    descer um nível quando as jornadas vêm agrupadas numa subpasta: a pasta das
    fotografias não pode mudar de lugar com isso.

    A exceção é a pasta de dados inteira copiada para dentro de outra
    (extracoes\\extracoes\\…, como em 08/10): as fotografias vieram junto, e
    sem isto o app não achava nenhuma — e gravava a próxima em outro lugar.
    """
    padrao = os.path.join(base_dir, fallout_core.PASTA_DADOS, PASTA_FOTOS)
    if not os.path.isdir(padrao):
        aninhada = os.path.join(fallout_core.raiz_dados(base_dir), PASTA_FOTOS)
        if os.path.isdir(aninhada):
            return aninhada
    return padrao


def assinatura_regras():
    """
    Doze caracteres que mudam quando o código das regras muda: o SHA-1 do
    fallout_core.py com o pipeline.py.

    Na comparação, assinaturas diferentes avisam que parte da diferença entre
    as duas fotografias pode ser da regra, e não dos dados — o Legado Aberto,
    por exemplo, mudou categorias sem nenhum arquivo novo. As quebras de linha
    são normalizadas: o mesmo código copiado de outra máquina, com CRLF no
    lugar de LF, não pode parecer outra regra.
    """
    import pipeline

    resumo = hashlib.sha1()
    for modulo in (fallout_core, pipeline):
        with open(modulo.__file__, "rb") as f:
            resumo.update(f.read().replace(b"\r\n", b"\n"))
    return resumo.hexdigest()[:12]


def _num(valor):
    """Número para o JSON: float, ou None no ausente (o JSON não tem NaN)."""
    if valor is None:
        return None
    valor = float(valor)
    return None if math.isnan(valor) else valor


def _categorias(df, hoje):
    """Categoria e subgrupo de cada pedido (ver `fallout_core.categorias_por_pedido`)."""
    return fallout_core.categorias_por_pedido(df, hoje)


def _ofensores_longo(ns):
    """
    A tabela de ofensores do slide (pivot2), com um mês por linha em vez de um
    mês por coluna — as colunas do pivot são os meses do relatório, que mudam
    de uma fotografia para a outra. `Ordem` guarda a ordem do slide.
    """
    meses = list(ns["meses_tab"])
    # `reindex`, e não seleção direta: sem nenhum ofensor o pivot pode sair sem
    # as colunas do defeito.
    piv = ns["pivot2"].reset_index(drop=True).reindex(columns=COLS_OFENSOR + meses)
    piv[meses] = piv[meses].fillna(0)
    piv.insert(0, "Ordem", range(len(piv)))
    # As colunas do pivot saem do `unstack("Mes")` e carregam o nome "Mes",
    # que o `melt` tentaria usar para a coluna nova.
    piv.columns.name = None
    return piv.melt(id_vars=["Ordem"] + COLS_OFENSOR, value_vars=meses,
                    var_name="Mes", value_name="Pedidos")


def _relatorio(ns, jornadas_combo, diagnostico):
    """Os números do relatório de uma visão, em tipos que o JSON aceita."""
    return {
        "jornadas": list(jornadas_combo),
        "mes_atual": int(ns["mes_atual"]),
        "fallout_pct": _num(ns["fallout_pct"]),
        "total_mes": _num(ns["total_mes"]),
        "meses_tab": [int(m) for m in ns["meses_tab"]],
        # O subitem é o nível (0, 1 ou 2); nas fotografias antigas é booleano,
        # e True vale 1.
        "dist_items": [[rotulo, _num(valor), int(subitem), bool(negrito)]
                       for rotulo, valor, subitem, negrito in ns["dist_items"]],
        "pct_nao_recorrente": _num(ns["pct_nao_recorrente"]),
        "labels_hist": list(ns["labels_hist"]),
        "vals_hist": [_num(v) for v in ns["vals_hist"]],
        "plan_hist": [_num(v) for v in ns["plan_hist"]],
        "proj_labels": list(ns["proj_labels"]),
        "proj_vals": [_num(v) for v in ns["proj_vals"]],
        "proj_detalhe": [{k: (v if k == "mes" else _num(v)) for k, v in d.items()}
                         for d in ns["proj_detalhe"]],
        "diagnostico": diagnostico,
    }


def _com_visao(tabela, visao):
    """A tabela de uma visão com a coluna `Visao` na frente."""
    tabela = tabela.copy()
    tabela.insert(0, "Visao", visao)
    return tabela


def _concatenar(partes):
    """Empilha as tabelas das visões, sem as vazias (que só confundem os tipos)."""
    cheias = [p for p in partes if len(p)]
    return pd.concat(cheias or partes[:1], ignore_index=True)


def _para_parquet(tabela):
    """
    Tipos que o Parquet aceita. Uma coluna não pode misturar texto, número e
    data, e é o que acontece nas colunas que juntam jornadas de formatos
    diferentes: texto vira `string`, a data de entrega vira data e o mês,
    inteiro.
    """
    tabela = tabela.copy()
    for col in tabela.columns:
        serie = tabela[col]
        if col in COLS_DATA and not pd.api.types.is_datetime64_any_dtype(serie):
            try:
                tabela[col] = pd.to_datetime(serie, errors="coerce")
            except (TypeError, ValueError):
                # Datas com e sem fuso na mesma coluna não se convertem juntas.
                # É data de calendário: vale o dia no fuso de Brasília.
                tabela[col] = fallout_core.dia_local(
                    pd.to_datetime(serie, errors="coerce", utc=True))
        elif serie.dtype == object:
            tabela[col] = serie.astype("string")
    if "Mes" in tabela.columns:
        tabela["Mes"] = tabela["Mes"].astype("Int64")
    return tabela


def _json_padrao(obj):
    """O que o `json` não serializa sozinho: escalares do numpy e datas."""
    if hasattr(obj, "item"):
        return obj.item()
    if isinstance(obj, (_dt.date, _dt.datetime)):
        return obj.isoformat()
    raise TypeError(f"{type(obj).__name__} não cabe no meta.json")


def _fotografar_visao(base_dir, visao, hoje):
    """As quatro tabelas e o relatório de uma visão (ver `gerar`)."""
    # Aqui, e não no topo: o app importa este módulo ao abrir, só para listar
    # as fotografias, e o pipeline traz o matplotlib junto.
    import pipeline

    # O diagnóstico que o pipeline imprime não interessa aqui, como no Excel
    # de ofensores.
    with contextlib.redirect_stdout(io.StringIO()):
        df, _octane, resumo, combo = fallout_core.carregar_base(base_dir, visao, hoje=hoje)
        ns = pipeline.calcular_relatorio(base_dir, visao, base=(df, resumo), hoje=hoje)

    categoria, subgrupo, sobrepostos = _categorias(df, hoje)
    pedidos = df.reindex(columns=COLS_PEDIDO)
    pedidos["Categoria"] = categoria
    pedidos["Subgrupo"] = subgrupo

    # A fonte de cada mês: o histórico digitado vence a extração nos meses
    # fechados. Quando um mês passa de uma para a outra entre duas fotografias,
    # a variação é troca de fonte, e não do fallout.
    fechados = fallout_core.meses_do_historico(
        fallout_core.historico_manual(base_dir, visao, combo))
    resumo = resumo.rename_axis("Mes").reset_index()
    resumo["Fonte"] = ["histórico" if m in fechados else "extração" for m in resumo["Mes"]]

    diagnostico = {
        "pedidos": int(len(df)),
        "sobrepostos": int(sobrepostos),
        "sem_categoria": int((categoria.isna() & df["Mes"].notna()).sum()),
    }
    return {
        "pedidos": _com_visao(pedidos, visao),
        "resumo": _com_visao(resumo, visao),
        "reducao": _com_visao(ns["reducao_defeitos"], visao),
        "ofensores": _com_visao(_ofensores_longo(ns), visao),
        "relatorio": _relatorio(ns, combo, diagnostico),
    }


def gerar(base_dir, hoje=None, origem="app", progresso=None):
    """
    Calcula a fotografia dos dados que estão agora em `extracoes/`, sem gravar.

    `hoje` é a data de referência das regras (ver `fallout_core.hoje_utc`);
    sem ela, é a data atual. `origem` fica no meta.json, para a comparação
    dizer se a fotografia foi tirada no app, pela linha de comando ou
    reconstruída. `progresso(feitos, total, rotulo)`, quando passado, é chamado
    antes de cada visão, como nos geradores de PowerPoint e Excel.

    Uma visão que falhar não derruba as outras: ela fica de fora e o motivo vai
    para `meta["erros"]`, como no Excel de ofensores.
    """
    corte = fallout_core.data_corte(base_dir, hoje)
    if corte is None:
        raise ValueError("Nenhuma data legível nas extrações: não há o que fotografar.")
    hoje = fallout_core.hoje_utc(hoje)
    visoes = fallout_core.visoes_disponiveis(base_dir)

    pedidos, resumos, reducoes, ofensores = [], [], [], []
    relatorios, erros = {}, {}
    for i, visao in enumerate(visoes):
        if progresso:
            progresso(i, len(visoes), fallout_core.nome_exibicao(visao))
        try:
            partes = _fotografar_visao(base_dir, visao, hoje)
        except Exception as e:                              # noqa: BLE001
            erros[visao] = f"{type(e).__name__}: {e}"
            continue
        pedidos.append(partes["pedidos"])
        resumos.append(partes["resumo"])
        reducoes.append(partes["reducao"])
        ofensores.append(partes["ofensores"])
        relatorios[visao] = partes["relatorio"]

    if not relatorios:
        raise RuntimeError("Nenhuma visão pôde ser fotografada. "
                           + " | ".join(f"{v}: {e}" for v, e in erros.items()))

    meta = {
        "versao_formato": VERSAO_FORMATO,
        "corte": corte.isoformat(),
        "hoje": hoje.date().isoformat(),
        "gerada_em": _dt.datetime.now().isoformat(timespec="seconds"),
        "origem": origem,
        "regras": assinatura_regras(),
        "visoes": list(relatorios),
        "erros": erros,
        "avisos": [],
        "fontes": fallout_core.arquivos_de_entrada(base_dir),
        "relatorios": relatorios,
    }
    return Fotografia(
        meta=meta,
        pedidos=_para_parquet(_concatenar(pedidos)),
        resumo=_para_parquet(_concatenar(resumos)),
        reducao=_para_parquet(_concatenar(reducoes)),
        ofensores=_para_parquet(_concatenar(ofensores)),
    )


def salvar(foto, base_dir):
    """
    Grava a fotografia em `extracoes/_fotografias/<corte>/` e devolve a pasta.

    Uma por data de corte: gravar de novo a mesma extração substitui a anterior
    — é como se refaz a fotografia depois de trocar o Octane, por exemplo. A
    pasta nova é montada ao lado e trocada no fim, como em
    `fallout_core._gravar_espelho`: uma gravação interrompida, ou duas sessões
    gravando juntas, não deixam fotografia pela metade.
    """
    pai = pasta_fotografias(base_dir)
    os.makedirs(pai, exist_ok=True)
    destino = os.path.join(pai, foto.id)
    sufixo = f"{os.getpid()}-{threading.get_ident()}"
    parcial = f"{destino}.{sufixo}.parcial"
    antiga = f"{destino}.{sufixo}.antiga"
    try:
        os.makedirs(parcial)
        for nome in TABELAS:
            getattr(foto, nome).to_parquet(os.path.join(parcial, f"{nome}.parquet"),
                                           index=False, compression="zstd")
        with open(os.path.join(parcial, ARQ_META), "w", encoding="utf-8") as f:
            json.dump(foto.meta, f, ensure_ascii=False, indent=1,
                      allow_nan=False, default=_json_padrao)

        # No Windows não dá para renomear uma pasta por cima de outra: a
        # antiga sai do caminho primeiro e volta se a troca falhar.
        tinha_antiga = os.path.isdir(destino)
        if tinha_antiga:
            os.replace(destino, antiga)
        try:
            os.replace(parcial, destino)
        except OSError:
            if tinha_antiga:
                os.replace(antiga, destino)
            raise
    finally:
        shutil.rmtree(parcial, ignore_errors=True)
        shutil.rmtree(antiga, ignore_errors=True)
    return destino


def listar(base_dir):
    """
    As fotografias gravadas, da mais recente para a mais antiga, como
    [{"id", "corte", "gravada_em"}].

    Só o nome da pasta e a data do meta.json, sem abrir nada: o app mostra
    isso a cada interação. Pastas que não são de fotografia (a `.parcial` de
    uma gravação em andamento, por exemplo) ficam de fora.
    """
    pasta = pasta_fotografias(base_dir)
    if not os.path.isdir(pasta):
        return []
    fotos = []
    for nome in sorted(os.listdir(pasta), reverse=True):
        meta = os.path.join(pasta, nome, ARQ_META)
        if not (RE_ID.match(nome) and os.path.isfile(meta)):
            continue
        try:
            corte = _dt.date.fromisoformat(nome)
        except ValueError:                  # "2026-13-01": parece, mas não é data
            continue
        fotos.append({
            "id": nome,
            "corte": corte,
            "gravada_em": _dt.datetime.fromtimestamp(os.path.getmtime(meta)),
        })
    return fotos


def ler_meta(base_dir, id_foto):
    """O meta.json de uma fotografia."""
    with open(os.path.join(pasta_fotografias(base_dir), id_foto, ARQ_META),
              encoding="utf-8") as f:
        meta = json.load(f)
    if meta.get("versao_formato", 0) > VERSAO_FORMATO:
        raise ValueError(
            f"A fotografia {id_foto} foi gravada por uma versão mais nova do app "
            f"(formato {meta['versao_formato']}); atualize o programa para abri-la.")
    return meta


def carregar(base_dir, id_foto):
    """Uma fotografia gravada, inteira."""
    pasta = os.path.join(pasta_fotografias(base_dir), id_foto)
    tabelas = {nome: pd.read_parquet(os.path.join(pasta, f"{nome}.parquet"))
               for nome in TABELAS}
    return Fotografia(meta=ler_meta(base_dir, id_foto), **tabelas)


def _tamanho(pasta):
    return sum(os.path.getsize(os.path.join(pasta, n)) for n in os.listdir(pasta))


def main(argv=None):
    import argparse

    parser = argparse.ArgumentParser(
        description="Grava ou lista as fotografias das extrações.")
    parser.add_argument("--hoje", help="data de referência das regras (AAAA-MM-DD); "
                                       "padrão: a data atual")
    parser.add_argument("--listar", action="store_true",
                        help="só lista as fotografias gravadas")
    args = parser.parse_args(argv)
    base_dir = os.path.dirname(os.path.abspath(__file__))

    if args.listar:
        fotos = listar(base_dir)
        if not fotos:
            print("Nenhuma fotografia gravada.")
        for f in fotos:
            meta = ler_meta(base_dir, f["id"])
            print(f"  dados até {f['corte']:%d/%m/%Y} · gravada em "
                  f"{f['gravada_em']:%d/%m/%Y %H:%M} · {meta['origem']}")
        return 0

    foto = gerar(base_dir, hoje=args.hoje, origem="linha de comando",
                 progresso=lambda feitos, total, rotulo:
                     print(f"  [{feitos + 1}/{total}] {rotulo}"))
    destino = salvar(foto, base_dir)
    imprimir_resumo(foto, destino)
    return 0


def imprimir_resumo(foto, destino):
    """O que foi gravado, visão a visão — para a linha de comando."""
    meta = foto.meta
    print(f"\nFotografia gravada em {destino}")
    print(f"  dados até {_dt.date.fromisoformat(meta['corte']):%d/%m/%Y} · "
          f"hoje = {_dt.date.fromisoformat(meta['hoje']):%d/%m/%Y} · "
          f"{_tamanho(destino) / 1024 / 1024:.1f} MB · regras {meta['regras']}")
    for visao, rel in meta["relatorios"].items():
        diag = rel["diagnostico"]
        alerta = ""
        if diag["sobrepostos"] or diag["sem_categoria"]:
            alerta = (f"  ⚠ {diag['sobrepostos']} em duas categorias e "
                      f"{diag['sem_categoria']} sem categoria")
        taxa = fallout_core.fmt_pct(rel["fallout_pct"]) or "—"
        print(f"  {fallout_core.nome_exibicao(visao):18s} "
              f"{fallout_core.MESES_LABEL[rel['mes_atual']]}: {taxa:>7s} · "
              f"{fallout_core.fmt_inteiro(diag['pedidos']):>6s} pedidos{alerta}")
    for visao, erro in meta["erros"].items():
        print(f"  {visao}: FORA DA FOTOGRAFIA — {erro}")
    for aviso in meta["avisos"]:
        print(f"  aviso: {aviso}")


if __name__ == "__main__":
    sys.exit(main())
