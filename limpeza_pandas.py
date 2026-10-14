#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Camada pandas: DataFrame entra, DataFrame limpo sai.

E a interface que o app web e qualquer notebook usam. Internamente delega as
regras de negocio a corrigir_planilha.py e as verificacoes a validacoes.py.

    import limpeza_pandas as lp

    df = lp.ler("data/planilha_vendas.xlsx")
    r  = lp.limpar(df)

    r.df                      # DataFrame limpo, com tipos corretos
    r.correcoes               # DataFrame do que mudou
    r.pendencias              # DataFrame do que exige decisao humana
    r.laudo                   # DataFrame das 16 verificacoes
    lp.resumo_por(r.df, "Regiao")
    lp.serie_mensal(r.df)

O pandas cuida de I/O, tipagem, agregacao e analise. As regras seguem linha a
linha em corrigir_planilha.py porque o log de correcao e por celula, e
vetorizar custaria essa rastreabilidade.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

import corrigir_planilha as cp
import planilha_io as pio
import validacoes as vl

COLUNAS = pio.COLUNAS_SAIDA
GANHO = "Fechado Ganho"
PERDIDO = "Fechado Perdido"


# ===========================================================================
# LEITURA
# ===========================================================================

def ler(origem, aba: str = None, codec: str = None, separador: str = None,
        nome: str = None) -> pd.DataFrame:
    """Le um CSV ou XLSX (caminho, bytes ou file-like) como DataFrame bruto.

    O conteudo vem como o arquivo entregou, sem nenhuma correcao: quem limpa e
    `limpar`. Colunas e tipos ficam exatamente como na origem, para que toda
    diferenca apareca no log.
    """
    if isinstance(origem, (str, Path)):
        caminho = Path(origem)
        nome, dados = caminho.name, caminho.read_bytes()
    elif isinstance(origem, bytes):
        nome, dados = nome or "arquivo.csv", origem
    else:                                      # file-like (upload do Streamlit)
        nome = nome or getattr(origem, "name", "arquivo.csv")
        dados = origem.getvalue() if hasattr(origem, "getvalue") else origem.read()

    registros, colunas, _ = pio.ler_arquivo(nome, dados, codec, separador, aba)
    return pd.DataFrame(registros, columns=colunas)


# ===========================================================================
# LIMPEZA
# ===========================================================================

@dataclass
class Resultado:
    """Tudo que a limpeza produziu, ja em DataFrame."""
    df: pd.DataFrame
    correcoes: pd.DataFrame
    pendencias: pd.DataFrame
    removidos: pd.DataFrame
    laudo: pd.DataFrame
    alertas: list = field(default_factory=list)
    auditoria: object = None          # a Auditoria crua, para quem precisar
    retrato: object = None            # o retrato do arquivo antes da limpeza
    mapa: dict = field(default_factory=dict)
    pedidos: list = field(default_factory=list)   # objetos Pedido, para o escritor de Excel

    @property
    def linhas_lidas(self) -> int:
        return self.retrato.linhas if self.retrato else len(self.df)

    @property
    def total_melhorias(self) -> int:
        return len(self.correcoes)

    @property
    def total_pendencias(self) -> int:
        return len(self.pendencias)

    @property
    def total_problemas(self) -> int:
        return self.total_melhorias + self.total_pendencias

    @property
    def linhas_a_revisar(self) -> int:
        return int((self.df["Situacao"] == "Revisar").sum()) if len(self.df) else 0

    @property
    def aprovado(self) -> bool:
        return bool(self.laudo.empty) or bool(self.laudo["ok"].all())


def para_dataframe(pedidos: list) -> pd.DataFrame:
    """Pedidos limpos em DataFrame, com datetime64 e float.

    A formatacao para dd/mm/aaaa e virgula decimal fica para a exportacao, para
    nao atrapalhar agregacao e ordenacao.
    """
    df = pd.DataFrame([{
        "ID Pedido": p.id,
        "Nome do Cliente": p.cliente,
        "Regiao": p.regiao,
        "Representante": p.representante,
        "Data do Pedido": p.data,
        "Receita": p.receita,
        "Produto": p.produto,
        "Status": p.status,
        "Observacoes": p.observacoes,
        "Situacao": "Revisar" if p.pendencias else "OK",
        "Pendencias": "; ".join(p.pendencias) if p.pendencias else None,
    } for p in pedidos], columns=COLUNAS)

    if df.empty:
        return df
    df["Data do Pedido"] = pd.to_datetime(df["Data do Pedido"], errors="coerce")
    df["Receita"] = pd.to_numeric(df["Receita"], errors="coerce")
    return df


def _correcoes_df(aud) -> pd.DataFrame:
    return pd.DataFrame([{
        "ID Pedido": c.id_pedido,
        "Coluna": c.coluna,
        "Antes": "(vazio)" if c.antes is None else str(c.antes),
        "Depois": "(vazio)" if c.depois is None else str(c.depois),
        "Regra": c.regra,
    } for c in aud.correcoes], columns=["ID Pedido", "Coluna", "Antes", "Depois", "Regra"])


def _pendencias_df(aud) -> pd.DataFrame:
    return pd.DataFrame([{
        "ID Pedido": p.id_pedido, "Tipo": p.tipo,
        "Detalhe": p.detalhe, "Acao sugerida": p.acao,
    } for p in aud.pendencias], columns=["ID Pedido", "Tipo", "Detalhe", "Acao sugerida"])


def _removidos_df(aud) -> pd.DataFrame:
    return pd.DataFrame([{
        "ID Pedido": p.id, "Nome do Cliente": p.cliente, "Regiao": p.regiao,
        "Representante": p.representante, "Data do Pedido": p.data,
        "Receita": p.receita, "Produto": p.produto, "Status": p.status,
        "Motivo da remocao": motivo,
    } for p, motivo in aud.removidos],
        columns=["ID Pedido", "Nome do Cliente", "Regiao", "Representante",
                 "Data do Pedido", "Receita", "Produto", "Status", "Motivo da remocao"])


def _laudo_df(laudo: list) -> pd.DataFrame:
    return pd.DataFrame([{
        "Item": r.item, "Verificacao": r.titulo, "ok": r.ok,
        "Situacao": "Aprovado" if r.ok else ("Alerta" if r.severidade == "alerta" else "Falha"),
        "Detalhe": r.detalhe,
    } for r in laudo], columns=["Item", "Verificacao", "ok", "Situacao", "Detalhe"])


def limpar(df: pd.DataFrame, mapa: dict = None,
           imputar_receita: bool = False) -> Resultado:
    """Aplica a limpeza a um DataFrame bruto e devolve tudo em DataFrame.

    `mapa` associa cada campo canonico ao nome da coluna no DataFrame. Quando
    omitido, os nomes sao detectados automaticamente.
    """
    mapa = mapa or cp.detectar_colunas(list(df.columns))
    registros = pio.registros_de(df)

    pedidos = cp.pedidos_de_registros(registros, mapa)
    retrato = vl.Retrato.de(pedidos)
    limpos, aud = cp.processar(pedidos, imputar_receita=imputar_receita)
    laudo = vl.validar(retrato, limpos, aud)

    return Resultado(
        df=para_dataframe(limpos),
        correcoes=_correcoes_df(aud),
        pendencias=_pendencias_df(aud),
        removidos=_removidos_df(aud),
        laudo=_laudo_df(laudo),
        alertas=list(aud.alertas),
        auditoria=aud,
        retrato=retrato,
        mapa=mapa,
        pedidos=limpos,
    )


def limpar_arquivo(origem, aba: str = None, codec: str = None, separador: str = None,
                   imputar_receita: bool = False) -> Resultado:
    """Atalho: le o arquivo e limpa em uma chamada."""
    return limpar(ler(origem, aba, codec, separador), imputar_receita=imputar_receita)


# ===========================================================================
# DIAGNOSTICO - a mesma medicao aplicada antes e depois
# ===========================================================================

ROTULO_SAIDA = {
    "id": "ID Pedido", "cliente": "Nome do Cliente", "regiao": "Regiao",
    "representante": "Representante", "data": "Data do Pedido", "receita": "Receita",
    "produto": "Produto", "status": "Status", "observacoes": "Observacoes",
}


def _series_canonicas(df: pd.DataFrame, mapa: dict = None) -> dict:
    """{campo canonico: Series}, funcionando tanto no bruto quanto no limpo."""
    if mapa is None:
        mapa = {campo: rotulo for campo, rotulo in ROTULO_SAIDA.items()
                if rotulo in df.columns}
        if not mapa:
            mapa = cp.detectar_colunas(list(df.columns))
    return {campo: df[coluna] for campo, coluna in mapa.items()
            if coluna and coluna in df.columns}


def _vazio(valor) -> bool:
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(valor, str) and not valor.strip()


def _forma_da_data(valor) -> str:
    """Assinatura do formato de uma data, para contar quantos convivem."""
    if _vazio(valor):
        return ""
    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return "tipo data"
    return re.sub(r"\d", "9", str(valor).strip())


def diagnosticar(df: pd.DataFrame, mapa: dict = None) -> dict:
    """Mede os defeitos de um DataFrame, bruto ou limpo.

    A mesma funcao roda nos dois lados do comparativo: medir cada lado de um
    jeito nao provaria nada.
    """
    colunas = _series_canonicas(df, mapa)
    if df.empty or not colunas:
        return {"linhas": len(df), "nulos": 0, "receita_texto": 0, "status_fora": 0,
                "caixa_inconsistente": 0, "formatos_data": 0, "duplicatas": 0,
                "detalhes": {}}

    # Observacoes fica de fora: a limpeza a esvazia de proposito ao tirar as
    # marcacoes de auditoria, e conta-la faria o indicador piorar por isso.
    nulos = int(sum(serie.map(_vazio).sum()
                    for campo, serie in colunas.items() if campo != "observacoes"))

    receita = colunas.get("receita")
    receita_texto = int(receita.map(
        lambda v: isinstance(v, str) and bool(re.search(r"R\$|[A-Za-z]", v))
    ).sum()) if receita is not None else 0

    status = colunas.get("status")
    validos = set(cp.STATUS_CANONICO.values())
    status_fora = int(status.map(
        lambda v: not _vazio(v) and str(v).strip() not in validos
    ).sum()) if status is not None else 0

    caixa = 0
    detalhe_caixa = {}
    for campo in ("cliente", "regiao", "representante", "status"):
        serie = colunas.get(campo)
        if serie is None:
            continue
        fora = serie.map(lambda v: not _vazio(v) and str(v) != cp.capitalizar(v))
        caixa += int(fora.sum())
        if int(fora.sum()):
            detalhe_caixa[ROTULO_SAIDA[campo]] = int(fora.sum())

    data = colunas.get("data")
    formas = set()
    if data is not None:
        formas = {f for f in data.map(_forma_da_data) if f}

    duplicatas = 0
    if {"cliente", "produto", "receita"} <= set(colunas):
        chave = pd.DataFrame({
            "cliente": colunas["cliente"].map(cp.chave),
            "produto": colunas["produto"].map(cp.chave),
            "receita": colunas["receita"].map(lambda v: cp.ler_receita(v)),
        })
        marcadas = chave.dropna(subset=["receita"]).duplicated(keep=False)
        duplicatas = int(marcadas.sum())

    return {
        "linhas": len(df),
        "nulos": nulos,
        "receita_texto": receita_texto,
        "status_fora": status_fora,
        "caixa_inconsistente": caixa,
        "formatos_data": len(formas),
        "duplicatas": duplicatas,
        "detalhes": {"formatos de data": sorted(formas), "caixa por coluna": detalhe_caixa},
    }


# So defeitos: contagem de linhas vai no resumo, como entrada e saida.
ROTULOS_DIAGNOSTICO = [
    ("nulos", "Nulos em campos obrigatorios"),
    ("receita_texto", "Receita como texto"),
    ("status_fora", "Status fora do dominio"),
    ("caixa_inconsistente", "Caixa inconsistente"),
    ("formatos_data", "Formatos de data distintos"),
    ("duplicatas", "Linhas em duplicidade"),
]

# Nao devem zerar: suspeita espera decisao humana, e um formato de data e o certo.
NAO_ZERAM = {"duplicatas", "formatos_data"}


def _situacao(chave: str, antes: int, depois: int) -> str:
    if antes == 0 and depois == 0:
        return "Sem defeito"
    if depois == 0:
        return "Resolvido"
    if chave in NAO_ZERAM and depois <= antes:
        return "Esperado" if depois < antes or chave == "formatos_data" else "Em aberto"
    if depois < antes:
        return "Reduzido"
    return "Em aberto"


def comparativo(antes: dict, depois: dict) -> pd.DataFrame:
    """Tabela antes-depois dos indicadores de defeito, com a situacao de cada um."""
    linhas = []
    for chave, rotulo in ROTULOS_DIAGNOSTICO:
        a, d = antes.get(chave, 0), depois.get(chave, 0)
        linhas.append({"Indicador": rotulo, "Antes": a, "Depois": d,
                       "Variacao": d - a, "Situacao": _situacao(chave, a, d)})
    return pd.DataFrame(linhas)


def correcoes_por_regra(resultado: "Resultado") -> pd.DataFrame:
    """Correcoes agrupadas pela regra que as motivou, nao pela coluna."""
    if resultado.correcoes.empty:
        return pd.DataFrame(columns=["Correcao", "Quantidade", "Linhas"])
    return (resultado.correcoes
            .groupby("Regra")
            .agg(Quantidade=("Regra", "size"), Linhas=("ID Pedido", "nunique"))
            .sort_values("Quantidade", ascending=False)
            .reset_index()
            .rename(columns={"Regra": "Correcao"}))


# ===========================================================================
# ANALISE
# ===========================================================================

def _ganhos(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["Status"] == GANHO]


def metricas(df: pd.DataFrame) -> dict:
    """Indicadores de fechamento do arquivo limpo."""
    if df.empty:
        return {"pedidos": 0, "ganhos": 0, "ganhos_com_receita": 0, "perdidos": 0,
                "receita_ganha": 0.0, "receita_perdida": 0.0, "ticket_medio": 0.0,
                "taxa_de_ganho": 0.0, "sem_receita": 0, "a_revisar": 0}

    ganhos = _ganhos(df)
    perdidos = df[df["Status"] == PERDIDO]
    decididos = len(ganhos) + len(perdidos)

    # O .mean() ignora nulo, e e a conta certa: pedido sem valor no denominador
    # trataria receita ausente como zero. O campo abaixo deixa isso visivel.
    com_receita = int(ganhos["Receita"].notna().sum())

    return {
        "pedidos": len(df),
        "ganhos": len(ganhos),
        "ganhos_com_receita": com_receita,
        "perdidos": len(perdidos),
        "receita_ganha": float(ganhos["Receita"].sum()),
        "receita_perdida": float(perdidos["Receita"].sum()),
        "ticket_medio": float(ganhos["Receita"].mean()) if com_receita else 0.0,
        "taxa_de_ganho": len(ganhos) / decididos if decididos else 0.0,
        "sem_receita": int(df["Receita"].isna().sum()),
        "a_revisar": int((df["Situacao"] == "Revisar").sum()),
    }


def resumo_por(df: pd.DataFrame, dimensao: str) -> pd.DataFrame:
    """Receita e desempenho agregados por regiao, representante ou produto."""
    if df.empty or dimensao not in df.columns:
        return pd.DataFrame()

    por_dimensao = df.groupby(dimensao, dropna=False)
    tabela = por_dimensao.agg(
        Pedidos=("ID Pedido", "count"),
        Ganhos=("Status", lambda s: int((s == GANHO).sum())),
        Perdidos=("Status", lambda s: int((s == PERDIDO).sum())),
    )
    ganhos = _ganhos(df).groupby(dimensao, dropna=False)["Receita"]
    tabela["Receita ganha"] = ganhos.sum()
    tabela["Ticket medio"] = ganhos.mean()
    tabela["A revisar"] = por_dimensao["Situacao"].apply(lambda s: int((s == "Revisar").sum()))

    decididos = tabela["Ganhos"] + tabela["Perdidos"]
    tabela["Taxa de ganho"] = (tabela["Ganhos"] / decididos.where(decididos > 0)).fillna(0)

    tabela = tabela.fillna({"Receita ganha": 0.0, "Ticket medio": 0.0})
    return tabela.sort_values("Receita ganha", ascending=False).reset_index()


def serie_mensal(df: pd.DataFrame) -> pd.DataFrame:
    """Receita ganha e pedidos por mes, com os meses sem venda preenchidos."""
    ganhos = _ganhos(df).dropna(subset=["Data do Pedido"])
    if ganhos.empty:
        return pd.DataFrame(columns=["Mes", "Receita ganha", "Pedidos"])

    por_mes = (ganhos.set_index("Data do Pedido")
                     .resample("MS")
                     .agg(**{"Receita ganha": ("Receita", "sum"),
                             "Pedidos": ("ID Pedido", "count")})
                     .reset_index()
                     .rename(columns={"Data do Pedido": "Mes"}))
    por_mes["Mes"] = por_mes["Mes"].dt.strftime("%Y-%m")
    return por_mes


def qualidade_por_coluna(resultado: Resultado) -> pd.DataFrame:
    """Quantas correcoes cada coluna recebeu e quantas linhas foram afetadas."""
    if resultado.correcoes.empty:
        return pd.DataFrame(columns=["Coluna", "Correcoes", "Linhas afetadas"])
    return (resultado.correcoes
            .groupby("Coluna")
            .agg(Correcoes=("Coluna", "size"),
                 **{"Linhas afetadas": ("ID Pedido", "nunique")})
            .sort_values("Correcoes", ascending=False)
            .reset_index())


def pendencias_por_tipo(resultado: Resultado) -> pd.DataFrame:
    if resultado.pendencias.empty:
        return pd.DataFrame(columns=["Tipo", "Ocorrencias", "Pedidos"])
    return (resultado.pendencias
            .groupby("Tipo")
            .agg(Ocorrencias=("Tipo", "size"), Pedidos=("ID Pedido", "nunique"))
            .sort_values("Ocorrencias", ascending=False)
            .reset_index())


# ===========================================================================
# EXPORTACAO
# ===========================================================================

def para_csv(resultado: Resultado, perfil: str = "Excel brasileiro") -> bytes:
    separador, decimal, codec = pio.PERFIS_CSV[perfil]
    df = resultado.df.copy()
    if not df.empty:
        df["Data do Pedido"] = df["Data do Pedido"].dt.strftime("%d/%m/%Y")
    return df.to_csv(index=False, sep=separador, decimal=decimal,
                     float_format="%.2f", encoding=codec).encode(codec, errors="replace")


def para_excel(resultado: Resultado, origem: str = "origem") -> bytes:
    """A pasta de trabalho completa, com as cinco abas de corrigir_planilha.py."""
    return pio.para_excel(resultado.pedidos, resultado.auditoria,
                          origem, resultado.linhas_lidas)
