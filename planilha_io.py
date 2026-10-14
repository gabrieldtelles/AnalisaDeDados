#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Leitura e escrita de planilhas para o app web.

Fica separado de app.py de proposito: importar o app executa a pagina inteira do
Streamlit, o que tornaria estas funcoes impossiveis de testar isoladamente. Aqui
nao ha nenhuma chamada de interface - so entrada e saida de dados.
"""

from __future__ import annotations

import csv
import io
import tempfile
from datetime import date, datetime
from pathlib import Path

import pandas as pd

import corrigir_planilha as cp

COLUNAS_SAIDA = ["ID Pedido", "Nome do Cliente", "Regiao", "Representante",
                 "Data do Pedido", "Receita", "Produto", "Status",
                 "Observacoes", "Situacao", "Pendencias"]

CAMPOS = ["id", "cliente", "regiao", "representante", "data",
          "receita", "produto", "status", "observacoes"]

ROTULOS = {"id": "ID do pedido", "cliente": "Nome do cliente", "regiao": "Regiao",
           "representante": "Representante", "data": "Data do pedido",
           "receita": "Receita", "produto": "Produto", "status": "Status",
           "observacoes": "Observacoes"}

OBRIGATORIAS = {"id", "cliente", "data", "receita"}

PERFIS_CSV = {
    # rotulo: (separador, separador decimal, codificacao)
    "Excel brasileiro": (";", ",", "utf-8-sig"),
    "Padrao internacional": (",", ".", "utf-8"),
}


# ===========================================================================
# LEITURA
# ===========================================================================

def detectar_codificacao(dados: bytes) -> str:
    """CSV exportado do Excel brasileiro costuma vir em cp1252, nao em utf-8."""
    for codec in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            dados.decode(codec)
            return codec
        except UnicodeDecodeError:
            continue
    return "latin-1"


def detectar_separador(texto: str) -> str:
    """O Excel em portugues grava CSV com ponto e virgula, nao com virgula."""
    linhas = texto.splitlines()[:20]
    amostra = "\n".join(linhas)
    try:
        return csv.Sniffer().sniff(amostra, delimiters=";,\t|").delimiter
    except csv.Error:
        primeira = linhas[0] if linhas else ""
        return max(";,\t|", key=primeira.count) if primeira else ","


def registros_de(df: pd.DataFrame) -> list:
    """DataFrame para lista de dicionarios, com NaN virando None.

    O motor de limpeza trabalha com None para celula vazia. Deixar NaN passar
    faria `if valor is None` falhar silenciosamente, porque NaN nao e None.
    """
    return [{coluna: (None if pd.isna(valor) else valor) for coluna, valor in linha.items()}
            for linha in df.to_dict(orient="records")]


def ler_csv(dados: bytes, codec: str = None, separador: str = None):
    """Devolve (registros, cabecalho, codificacao usada, separador usado).

    A leitura e feita pelo pandas. O texto vem todo como string e a celula vazia
    como "", porque quem decide o tipo de cada coluna e corrigir_planilha.py:
    deixar o pandas inferir faria "1001" virar inteiro e "4320,00" virar texto
    sem que a conversao aparecesse no log de correcoes.
    """
    codec = codec or detectar_codificacao(dados)
    separador = separador or detectar_separador(dados.decode(codec, errors="replace"))
    df = pd.read_csv(io.BytesIO(dados), sep=separador, encoding=codec,
                     dtype=str, keep_default_na=False, na_values=[],
                     encoding_errors="replace")
    df.columns = [str(c).strip() for c in df.columns]
    return registros_de(df), list(df.columns), codec, separador


def abas_de(dados: bytes) -> list:
    import openpyxl
    return openpyxl.load_workbook(io.BytesIO(dados), read_only=True).sheetnames


def ler_xlsx(dados: bytes, aba: str = None):
    """Devolve (registros, cabecalho, abas disponiveis).

    Diferente do CSV, aqui os tipos nativos do Excel sao preservados: data vem
    como Timestamp e numero como numero. O motor sabe lidar com os dois.
    """
    abas = abas_de(dados)
    escolhida = aba if aba in abas else abas[0]
    df = pd.read_excel(io.BytesIO(dados), sheet_name=escolhida)
    df.columns = [str(c).strip() for c in df.columns]
    return registros_de(df), list(df.columns), abas


def ler_arquivo(nome: str, dados: bytes, codec=None, separador=None, aba=None):
    """Entrada unica para CSV e XLSX. Devolve (registros, cabecalho, descricao)."""
    if nome.lower().endswith(".csv"):
        registros, cabecalho, codec, separador = ler_csv(dados, codec, separador)
        descricao = f"codificacao `{codec}` · separador `{separador}`"
    else:
        registros, cabecalho, _ = ler_xlsx(dados, aba)
        descricao = f"aba `{aba}`" if aba else "primeira aba"
    return registros, cabecalho, descricao


# ===========================================================================
# ESCRITA
# ===========================================================================

def como_linhas(pedidos: list) -> list:
    """Os pedidos limpos no formato de saida, com a data ja em dd/mm/aaaa."""
    return [{
        "ID Pedido": p.id,
        "Nome do Cliente": p.cliente,
        "Regiao": p.regiao,
        "Representante": p.representante,
        "Data do Pedido": p.data.strftime("%d/%m/%Y")
                          if isinstance(p.data, (date, datetime)) else p.data,
        "Receita": p.receita,
        "Produto": p.produto,
        "Status": p.status,
        "Observacoes": p.observacoes,
        "Situacao": "Revisar" if p.pendencias else "OK",
        "Pendencias": "; ".join(p.pendencias) if p.pendencias else None,
    } for p in pedidos]


def para_csv(pedidos: list, separador: str = ";", decimal: str = ",",
             codec: str = "utf-8-sig") -> bytes:
    """CSV da planilha limpa, no perfil escolhido pelo usuario."""
    buffer = io.StringIO()
    escritor = csv.DictWriter(buffer, fieldnames=COLUNAS_SAIDA,
                              delimiter=separador, lineterminator="\r\n")
    escritor.writeheader()
    for linha in como_linhas(pedidos):
        pronta = dict(linha)
        valor = pronta["Receita"]
        if isinstance(valor, (int, float)):
            pronta["Receita"] = f"{valor:.2f}".replace(".", decimal)
        escritor.writerow({k: ("" if v is None else v) for k, v in pronta.items()})
    return buffer.getvalue().encode(codec, errors="replace")


def para_excel(pedidos: list, aud, origem: str, lidos: int) -> bytes:
    """A pasta de trabalho completa, com as cinco abas de corrigir_planilha.py."""
    with tempfile.TemporaryDirectory() as pasta:
        destino = Path(pasta) / "limpa.xlsx"
        cp.gravar(destino, pedidos, aud, Path(origem), lidos)
        return destino.read_bytes()
