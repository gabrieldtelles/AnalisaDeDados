#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Corrige as inconsistencias diagnosticadas em planilha_vendas.xlsx e grava uma
planilha nova, sem alterar o arquivo original.

O script e dirigido por regras declaradas no bloco de configuracao abaixo: nao
ha nenhum ID de pedido embutido na logica. Toda alteracao feita no dado entra na
aba "log_correcoes", com o valor antes, o valor depois e a regra que a motivou.

O que e corrigido automaticamente
---------------------------------
  1. Receita gravada como texto ("R$ 4320.00")  -> numero + formato de moeda
  2. Formato monetario hibrido (R$ com ponto)   -> formato pt-BR na celula
  3. Caixa irregular em texto                   -> capitalizacao canonica
  4. Regiao com grafias multiplas               -> dominio fechado de regioes
  5. Representante com grafias multiplas        -> forma canonica por nome
  6. Duplicata exata                            -> movida para a aba "removidos"
  7. Status ausente gravado em Observacoes      -> devolvido a coluna Status
  8. Status em outro idioma ("closed won")      -> dominio em portugues
  9. Regiao ausente                             -> imputada pelo territorio do rep.
 10. Representante ausente                      -> imputado pela regiao
 11. ID do pedido como decimal (1001.0)         -> inteiro
 12. Observacoes como campo multiuso            -> nota de negocio x marcacao de QA

O que NAO e corrigido automaticamente (vai para a aba "revisar")
----------------------------------------------------------------
 13. Duplicatas suspeitas    - exigem confronto com contrato ou nota fiscal
 14. Receita ausente         - o valor nao existe no arquivo, precisa da origem
 15. ID fora de ordem        - nao da para saber a intencao da numeracao
 16. Representante em duas regioes - pode ser acumulo legitimo de territorio

Uso
---
    python3 corrigir_planilha.py
    python3 corrigir_planilha.py --entrada data/planilha_vendas.xlsx \
                                 --saida   data/planilha_vendas_corrigida.xlsx
"""

from __future__ import annotations

import argparse
import re
import statistics
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    sys.exit("Dependencia ausente. Instale com:  python3 -m pip install openpyxl")


# ===========================================================================
# CONFIGURACAO
# ===========================================================================

ABA_ORIGEM = "pedidos_vendas"

COLUNAS_ORIGEM = [
    "id", "cliente", "regiao", "representante",
    "data", "receita", "produto", "status", "observacoes",
]

SINONIMOS_DE_COLUNA = {
    "id": ["id pedido", "id", "pedido", "numero do pedido", "num pedido", "n pedido",
           "codigo", "codigo do pedido", "order id", "order"],
    "cliente": ["nome do cliente", "cliente", "razao social", "nome", "customer",
                "customer name", "conta"],
    "regiao": ["regiao", "region", "territorio", "uf", "estado", "praca"],
    "representante": ["representante", "vendedor", "rep", "consultor", "responsavel",
                      "executivo", "sales rep", "owner"],
    "data": ["data do pedido", "data", "data da venda", "data emissao", "emissao",
             "date", "order date"],
    "receita": ["receita", "valor", "valor total", "total", "faturamento", "preco",
                "revenue", "amount", "value"],
    "produto": ["produto", "item", "servico", "product", "sku", "descricao"],
    "status": ["status", "situacao", "etapa", "estagio", "stage", "fase"],
    "observacoes": ["observacoes", "observacao", "obs", "comentarios", "notas",
                    "notes", "comment"],
}

# Chave: forma normalizada. Valor: grafia que vai para a planilha.
REGIOES_CANONICAS = {
    "norte": "Norte",
    "nordeste": "Nordeste",
    "centrooeste": "Centro-Oeste",
    "sudeste": "Sudeste",
    "sul": "Sul",
}

STATUS_CANONICO = {
    "fechado ganho": "Fechado Ganho",
    "fechado perdido": "Fechado Perdido",
    "ganho": "Fechado Ganho",
    "perdido": "Fechado Perdido",
    "closed won": "Fechado Ganho",
    "closed lost": "Fechado Perdido",
    "won": "Fechado Ganho",
    "lost": "Fechado Perdido",
}

# Marcacoes de auditoria: saem de Observacoes e viram pendencia estruturada.
MARCACOES_DE_AUDITORIA = {
    "duplicado": "Marcado como duplicado pelo revisor",
    "possivel duplicado": "Marcado como possivel duplicado pelo revisor",
    "receita ausente": "Receita ausente",
    "sem representante": "Sem representante",
    "revisar": "Marcado para revisao pelo revisor",
}

PARTICULAS = {"de", "da", "do", "das", "dos", "e", "di", "du", "del", "la", "van", "von"}

FORMATO_MOEDA = "R$ #,##0.00"
FORMATO_DATA = "DD/MM/YYYY"


# ===========================================================================
# NORMALIZACAO DE TEXTO
# ===========================================================================

def sem_acento(texto: str) -> str:
    decomposto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in decomposto if not unicodedata.combining(c))


def chave(valor) -> str:
    """Forma comparavel de um texto: sem acento, minuscula, espacos colapsados."""
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", sem_acento(str(valor)).strip().lower())


def chave_regiao(valor) -> str:
    """Como `chave`, mas ignorando hifen e espaco, para casar 'Centro-Oeste',
    'centro oeste' e 'CENTROOESTE' na mesma entrada do dominio."""
    return re.sub(r"[\s\-]", "", chave(valor))


def capitalizar(valor) -> str | None:
    """Capitalizacao de nome proprio em portugues, respeitando hifen e particulas."""
    if valor is None:
        return None
    bruto = re.sub(r"\s+", " ", str(valor).strip())
    if not bruto:
        return None

    palavras = []
    for i, palavra in enumerate(bruto.split(" ")):
        if i > 0 and chave(palavra) in PARTICULAS:
            palavras.append(palavra.lower())
            continue
        palavras.append("-".join(p[:1].upper() + p[1:].lower() for p in palavra.split("-")))
    return " ".join(palavras)


def ler_receita(valor):
    """Converte 'R$ 4320.00' (ou 'R$ 4.320,00', ou um numero) em float.

    O arquivo de origem usa o simbolo pt-BR com separador decimal en-US. A funcao
    aceita os dois formatos decidindo pelo ultimo separador encontrado.
    """
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = re.sub(r"[^\d,.\-]", "", str(valor))
    if not texto:
        return None

    ultima_virgula = texto.rfind(",")
    ultimo_ponto = texto.rfind(".")
    if ultima_virgula > ultimo_ponto:          # 4.320,00  -> pt-BR
        texto = texto.replace(".", "").replace(",", ".")
    else:                                      # 4,320.00 ou 4320.00 -> en-US
        texto = texto.replace(",", "")
    try:
        return float(texto)
    except ValueError:
        return None


def ler_id(valor):
    """Normaliza o identificador do pedido para inteiro quando ele for numerico.

    A origem pode entregar o mesmo ID de tres formas: 1001 (inteiro), 1001.0
    (decimal, como o Excel costuma gravar) ou "1001" (texto, como todo CSV
    entrega). As tres precisam virar o mesmo inteiro, senao a deduplicacao e a
    checagem de cronologia comparam tipos diferentes e falham em silencio.

    Um ID genuinamente textual ("PED-001") e preservado como texto.
    """
    if valor is None:
        return None
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, int):
        return valor
    if isinstance(valor, float):
        return int(valor) if valor.is_integer() else valor

    texto = str(valor).strip()
    if not texto:
        return None
    try:
        numero = float(texto)
    except ValueError:
        return texto
    return int(numero) if numero.is_integer() else numero


def ler_data(valor):
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(str(valor).strip(), formato).date()
        except ValueError:
            continue
    return None


# ===========================================================================
# ESTRUTURAS
# ===========================================================================

@dataclass
class Correcao:
    id_pedido: object
    coluna: str
    antes: object
    depois: object
    regra: str


@dataclass
class Pendencia:
    id_pedido: object
    tipo: str
    detalhe: str
    acao: str


@dataclass
class Pedido:
    linha_origem: int
    id: object = None
    cliente: str | None = None
    regiao: str | None = None
    representante: str | None = None
    data: date | None = None
    receita: float | None = None
    produto: str | None = None
    status: str | None = None
    observacoes: str | None = None
    pendencias: list[str] = field(default_factory=list)


class Auditoria:
    """Acumula tudo que o script mudou e tudo que deixou para decisao humana."""

    def __init__(self) -> None:
        self.correcoes: list[Correcao] = []
        self.pendencias: list[Pendencia] = []
        self.removidos: list[tuple[Pedido, str]] = []
        self.alertas: list[str] = []

    def corrigir(self, pedido: Pedido, coluna: str, antes, depois, regra: str) -> None:
        # O tipo entra na comparacao: 1001.0 == 1001 e verdadeiro em Python, e sem
        # isso a conversao de decimal para inteiro sumiria do log.
        if antes == depois and type(antes) is type(depois):
            return
        self.correcoes.append(Correcao(pedido.id, coluna, antes, depois, regra))

    def pendente(self, pedido: Pedido, tipo: str, detalhe: str, acao: str) -> None:
        self.pendencias.append(Pendencia(pedido.id, tipo, detalhe, acao))
        if tipo not in pedido.pendencias:
            pedido.pendencias.append(tipo)


# ===========================================================================
# LEITURA
# ===========================================================================

def detectar_colunas(cabecalho: list) -> dict:
    """Associa cada coluna canonica ao titulo correspondente no arquivo enviado.

    Devolve {chave_canonica: titulo_encontrado_ou_None}. A busca e feita em duas
    passadas: primeiro o nome exato (normalizado), depois correspondencia parcial,
    para que 'Valor Total (R$)' ainda encontre a coluna de receita. Um titulo ja
    usado nao e reaproveitado por outra coluna.
    """
    titulos = [t for t in cabecalho if t is not None and str(t).strip()]
    disponiveis = {chave(t): t for t in titulos}
    mapa = {}
    usados = set()

    for canonica, sinonimos in SINONIMOS_DE_COLUNA.items():
        achado = None
        for sinonimo in sinonimos:
            titulo = disponiveis.get(sinonimo)
            if titulo is not None and titulo not in usados:
                achado = titulo
                break
        mapa[canonica] = achado
        if achado:
            usados.add(achado)

    for canonica, sinonimos in SINONIMOS_DE_COLUNA.items():
        if mapa[canonica]:
            continue
        for k, titulo in disponiveis.items():
            if titulo in usados:
                continue
            if any(s in k or k in s for s in sinonimos):
                mapa[canonica] = titulo
                usados.add(titulo)
                break
    return mapa


def pedidos_de_registros(registros: list, mapa: dict = None, inicio: int = 2) -> list[Pedido]:
    """Monta a lista de Pedido a partir de dicionarios de linha.

    `mapa` traduz a chave canonica para o titulo usado no arquivo; quando omitido,
    assume que os registros ja vem com as chaves canonicas. String vazia e tratada
    como ausencia, que e como um CSV representa celula em branco.
    """
    mapa = mapa or {c: c for c in COLUNAS_ORIGEM}
    pedidos: list[Pedido] = []

    for numero, bruto in enumerate(registros, start=inicio):
        campos = {}
        for canonica in COLUNAS_ORIGEM:
            titulo = mapa.get(canonica)
            valor = bruto.get(titulo) if titulo else None
            if isinstance(valor, str) and not valor.strip():
                valor = None
            campos[canonica] = valor
        if all(v is None for v in campos.values()):
            continue
        pedidos.append(Pedido(linha_origem=numero, **campos))
    return pedidos


def carregar(caminho: Path, aba: str) -> list[Pedido]:
    livro = openpyxl.load_workbook(caminho, data_only=True)
    if aba not in livro.sheetnames:
        sys.exit(f"Aba '{aba}' nao encontrada em {caminho.name}. Abas: {livro.sheetnames}")

    planilha = livro[aba]
    registros = [
        dict(zip(COLUNAS_ORIGEM, list(valores) + [None] * len(COLUNAS_ORIGEM)))
        for valores in planilha.iter_rows(min_row=2, values_only=True)
    ]
    return pedidos_de_registros(registros)


# ===========================================================================
# ETAPAS DE CORRECAO
# ===========================================================================

def etapa_tipos(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Itens 1, 2 e 11: ID inteiro, data como data, receita como numero."""
    for p in pedidos:
        # Atribui antes de registrar, para o log inteiro usar o ID ja canonico.
        bruto = p.id
        p.id = ler_id(p.id)
        aud.corrigir(p, "ID Pedido", bruto, p.id, "ID normalizado para inteiro")

        data = ler_data(p.data)
        if data != p.data:
            aud.corrigir(p, "Data do Pedido", p.data, data, "Data normalizada para tipo data")
        p.data = data

        receita = ler_receita(p.receita)
        if isinstance(p.receita, str):
            aud.corrigir(p, "Receita", p.receita, receita,
                         "Receita gravada como texto convertida para numero")
        p.receita = receita


def etapa_texto(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Itens 3, 4, 5 e 8: caixa, dominio de regiao, dominio de status."""
    for p in pedidos:
        for campo, rotulo in (("cliente", "Nome do Cliente"),
                              ("representante", "Representante"),
                              ("produto", "Produto")):
            antes = getattr(p, campo)
            depois = capitalizar(antes)
            if antes is not None and antes != depois:
                aud.corrigir(p, rotulo, antes, depois, "Capitalizacao padronizada")
            setattr(p, campo, depois)

        if p.regiao is not None:
            antes = p.regiao
            depois = REGIOES_CANONICAS.get(chave_regiao(antes))
            if depois is None:
                depois = capitalizar(antes)
                aud.pendente(p, "Regiao fora do dominio", f"Valor '{antes}' nao esta na lista de regioes",
                             "Conferir a regiao correta e incluir no dominio")
            elif antes != depois:
                aud.corrigir(p, "Regiao", antes, depois, "Regiao padronizada pelo dominio")
            p.regiao = depois

        if p.status is not None:
            antes = p.status
            depois = STATUS_CANONICO.get(chave(antes))
            if depois is None:
                depois = capitalizar(antes)
                aud.pendente(p, "Status fora do dominio", f"Valor '{antes}' nao esta na lista de status",
                             "Conferir o status correto e incluir no dominio")
            elif antes != depois:
                traduzido = chave(antes) in {"closed won", "closed lost", "won", "lost"}
                aud.corrigir(p, "Status", antes, depois,
                             "Status traduzido do ingles e padronizado" if traduzido
                             else "Status padronizado pelo dominio")
            p.status = depois


def etapa_observacoes(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Itens 7 e 12: separar nota de negocio, marcacao de QA e status fora de lugar."""
    for p in pedidos:
        if p.observacoes is None or not str(p.observacoes).strip():
            p.observacoes = None
            continue

        bruto = str(p.observacoes).strip()
        restantes: list[str] = []

        for trecho in [t.strip() for t in re.split(r"[;|]", bruto) if t.strip()]:
            k = chave(trecho)

            if k in STATUS_CANONICO:
                valor = STATUS_CANONICO[k]
                if p.status is None:
                    p.status = valor
                    aud.corrigir(p, "Status", None, valor,
                                 "Status recuperado da coluna Observacoes")
                    aud.corrigir(p, "Observacoes", trecho, None,
                                 "Status removido de Observacoes apos ir para a coluna propria")
                elif p.status != valor:
                    aud.pendente(p, "Status divergente",
                                 f"Coluna Status diz '{p.status}' e Observacoes diz '{trecho}'",
                                 "Definir qual dos dois vale e corrigir a origem")
                continue

            if k in MARCACOES_DE_AUDITORIA:
                aud.corrigir(p, "Observacoes", trecho, None,
                             "Marcacao de auditoria movida para a coluna de pendencias")
                continue

            restantes.append(trecho)

        p.observacoes = "; ".join(restantes) if restantes else None


def etapa_imputacao(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Itens 9 e 10: preencher regiao e representante quando o territorio for
    inequivoco. Se houver mais de um candidato, vira pendencia em vez de chute."""
    rep_para_regioes: dict[str, set[str]] = defaultdict(set)
    regiao_para_reps: dict[str, set[str]] = defaultdict(set)
    for p in pedidos:
        if p.representante and p.regiao:
            rep_para_regioes[p.representante].add(p.regiao)
            regiao_para_reps[p.regiao].add(p.representante)

    for rep, regioes in sorted(rep_para_regioes.items()):
        if len(regioes) > 1:
            aud.alertas.append(
                f"{rep} aparece em mais de uma regiao ({', '.join(sorted(regioes))}). "
                "Confirmar se o acumulo de territorio e intencional."
            )

    for p in pedidos:
        if p.regiao is None and p.representante:
            candidatas = rep_para_regioes.get(p.representante, set())
            if len(candidatas) == 1:
                valor = next(iter(candidatas))
                p.regiao = valor
                aud.corrigir(p, "Regiao", None, valor,
                             "Regiao imputada pelo territorio do representante")
            else:
                aud.pendente(p, "Regiao ausente",
                             f"Representante '{p.representante}' atua em {len(candidatas)} regioes",
                             "Informar a regiao manualmente")

        if p.representante is None and p.regiao:
            candidatos = regiao_para_reps.get(p.regiao, set())
            if len(candidatos) == 1:
                valor = next(iter(candidatos))
                p.representante = valor
                aud.corrigir(p, "Representante", None, valor,
                             "Representante imputado pela regiao atendida")
            else:
                aud.pendente(p, "Representante ausente",
                             f"Regiao '{p.regiao}' e atendida por {len(candidatos)} representantes",
                             "Informar o representante manualmente")

        if p.regiao is None and p.representante is None:
            aud.pendente(p, "Territorio ausente", "Regiao e representante vazios",
                         "Informar os dois manualmente")


def etapa_campos_obrigatorios(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Item 14: receita ausente nao pode ser inventada - vira pendencia."""
    for p in pedidos:
        if p.receita is None:
            aud.pendente(p, "Receita ausente",
                         f"Pedido com status '{p.status or 'indefinido'}' e sem valor",
                         "Buscar o valor no contrato ou na nota fiscal")
        if p.status is None:
            aud.pendente(p, "Status ausente", "Nenhum status na linha nem em Observacoes",
                         "Definir se o negocio foi ganho ou perdido")
        if p.data is None:
            aud.pendente(p, "Data ausente", "Data do pedido vazia ou ilegivel",
                         "Informar a data do pedido")


def etapa_duplicatas(pedidos: list[Pedido], aud: Auditoria) -> list[Pedido]:
    """Itens 6 e 13.

    Duplicata exata (mesmo cliente, produto, data e valor) e removida: mantem-se
    o menor ID e o restante vai para a aba "removidos".

    Duplicata suspeita (mesmo cliente, produto e valor em datas diferentes) NAO e
    removida - so sinalizada, porque pode ser recompra legitima.
    """
    def ordem(p: Pedido):
        return (0, p.id) if isinstance(p.id, int) else (1, str(p.id))

    exatas: dict[tuple, list[Pedido]] = defaultdict(list)
    for p in pedidos:
        exatas[(chave(p.cliente), chave(p.produto), p.data, p.receita)].append(p)

    descartados: set[int] = set()
    for grupo in exatas.values():
        if len(grupo) < 2:
            continue
        mantido, *copias = sorted(grupo, key=ordem)
        for copia in copias:
            motivo = (f"Duplicata exata do pedido {mantido.id}: mesmo cliente, produto, "
                      f"data e valor")
            aud.removidos.append((copia, motivo))
            # Regra generica no log; o motivo detalhado fica na aba removidos.
            aud.correcoes.append(Correcao(copia.id, "(linha inteira)", "mantida",
                                          "removida", "Duplicata exata removida"))
            descartados.add(id(copia))

    mantidos = [p for p in pedidos if id(p) not in descartados]

    suspeitas: dict[tuple, list[Pedido]] = defaultdict(list)
    for p in mantidos:
        suspeitas[(chave(p.cliente), chave(p.produto), p.receita)].append(p)

    for grupo in suspeitas.values():
        if len(grupo) < 2 or grupo[0].receita is None:
            continue
        if len({g.data for g in grupo}) < 2:
            continue
        for atual in grupo:
            outros = [str(g.id) for g in grupo if g is not atual]
            datas = sorted(g.data for g in grupo if g.data)
            intervalo = (datas[-1] - datas[0]).days if len(datas) > 1 else 0
            aud.pendente(
                atual, "Duplicata suspeita",
                f"Mesmo cliente, produto e valor do pedido {', '.join(outros)} "
                f"({intervalo} dias de intervalo)",
                "Confrontar com contrato ou nota fiscal e excluir se for repeticao",
            )
    return mantidos


def etapa_receita_imputada(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Opcional: preenche receita ausente com a mediana do mesmo produto.

    Desligada por padrao, e de proposito: o valor imputado nao e um dado, e uma
    estimativa, e entra somando no faturamento apurado. Quando ligada, cada linha
    preenchida continua aparecendo como pendencia - sob o tipo "Receita imputada"
    em vez de "Receita ausente" - para que ninguem feche um numero achando que
    aquele valor veio do contrato.

    Roda DEPOIS da deduplicacao: a receita entra na chave que identifica
    duplicata, e imputar antes criaria pares falsos entre pedidos que so se
    parecem por causa da estimativa.
    """
    por_produto: dict[str, list[float]] = defaultdict(list)
    for p in pedidos:
        if p.receita is not None and p.produto:
            por_produto[chave(p.produto)].append(p.receita)

    for p in pedidos:
        if p.receita is not None:
            continue
        valores = por_produto.get(chave(p.produto), [])
        if not valores:
            continue
        mediana = round(statistics.median(valores), 2)
        aud.corrigir(p, "Receita", None, mediana, "Receita imputada pela mediana do produto")
        p.receita = mediana
        aud.pendente(
            p, "Receita imputada",
            f"Valor estimado pela mediana de {len(valores)} pedido(s) de '{p.produto}'",
            "Conferir o valor real no contrato ou na nota fiscal e substituir",
        )


def etapa_cronologia(pedidos: list[Pedido], aud: Auditoria) -> None:
    """Item 15: ID sequencial que anda para tras no tempo."""
    numerados = [p for p in pedidos if isinstance(p.id, int) and p.data]
    numerados.sort(key=lambda p: p.id)
    for anterior, atual in zip(numerados, numerados[1:]):
        if atual.data < anterior.data:
            aud.pendente(
                atual, "ID fora de ordem cronologica",
                f"Data {atual.data:%d/%m/%Y} e anterior a do pedido {anterior.id} "
                f"({anterior.data:%d/%m/%Y})",
                "Confirmar se o ID e sequencial por data ou se a data esta errada",
            )


def processar(pedidos: list[Pedido],
              imputar_receita: bool = False) -> tuple[list[Pedido], Auditoria]:
    """Aplica as etapas de correcao na ordem em que uma depende da outra.

    Texto antes de Observacoes (para o status achado la cair no dominio certo),
    territorio antes da deduplicacao, cronologia por ultimo.
    """
    aud = Auditoria()
    etapa_tipos(pedidos, aud)
    etapa_texto(pedidos, aud)
    etapa_observacoes(pedidos, aud)
    etapa_imputacao(pedidos, aud)
    pedidos = etapa_duplicatas(pedidos, aud)
    if imputar_receita:
        etapa_receita_imputada(pedidos, aud)
    etapa_campos_obrigatorios(pedidos, aud)
    etapa_cronologia(pedidos, aud)
    return pedidos, aud


# ===========================================================================
# ESCRITA
# ===========================================================================

CABECALHO_FILL = PatternFill("solid", fgColor="1F3864")
CABECALHO_FONT = Font(color="FFFFFF", bold=True, size=10)
ALERTA_FILL = PatternFill("solid", fgColor="FCE4E4")
OK_FILL = PatternFill("solid", fgColor="E7F3E9")
BORDA = Border(bottom=Side(style="thin", color="D9D9D9"))


def escrever_cabecalho(planilha, titulos: list[str]) -> None:
    planilha.append(titulos)
    for celula in planilha[1]:
        celula.fill = CABECALHO_FILL
        celula.font = CABECALHO_FONT
        celula.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
    planilha.row_dimensions[1].height = 28
    planilha.freeze_panes = "A2"


def ajustar_larguras(planilha, larguras: dict[str, int]) -> None:
    for coluna, largura in larguras.items():
        planilha.column_dimensions[coluna].width = largura


def aba_pedidos(livro, pedidos: list[Pedido]) -> None:
    planilha = livro.create_sheet("pedidos_vendas")
    titulos = ["ID Pedido", "Nome do Cliente", "Regiao", "Representante", "Data do Pedido",
               "Receita", "Produto", "Status", "Observacoes", "Situacao", "Pendencias"]
    escrever_cabecalho(planilha, titulos)

    for p in pedidos:
        situacao = "Revisar" if p.pendencias else "OK"
        planilha.append([
            p.id, p.cliente, p.regiao, p.representante, p.data, p.receita,
            p.produto, p.status, p.observacoes, situacao,
            "; ".join(p.pendencias) if p.pendencias else None,
        ])
        linha = planilha.max_row
        planilha.cell(linha, 1).number_format = "0"
        planilha.cell(linha, 5).number_format = FORMATO_DATA
        planilha.cell(linha, 6).number_format = FORMATO_MOEDA
        celula_situacao = planilha.cell(linha, 10)
        celula_situacao.fill = ALERTA_FILL if p.pendencias else OK_FILL
        celula_situacao.alignment = Alignment(horizontal="center")
        planilha.cell(linha, 11).alignment = Alignment(wrap_text=True, vertical="top")
        for coluna in range(1, len(titulos) + 1):
            planilha.cell(linha, coluna).border = BORDA

    ajustar_larguras(planilha, {"A": 11, "B": 20, "C": 14, "D": 18, "E": 15, "F": 14,
                                "G": 22, "H": 17, "I": 22, "J": 11, "K": 46})
    planilha.auto_filter.ref = f"A1:{get_column_letter(len(titulos))}{planilha.max_row}"

    ultima = max(planilha.max_row + 200, 300)
    dominios = {
        "C": sorted({p.regiao for p in pedidos if p.regiao}),
        "D": sorted({p.representante for p in pedidos if p.representante}),
        "G": sorted({p.produto for p in pedidos if p.produto}),
        "H": sorted(set(STATUS_CANONICO.values())),
    }
    for coluna, valores in dominios.items():
        formula = '"' + ",".join(valores) + '"'
        if len(formula) > 255 or any("," in v for v in valores):
            continue
        validacao = DataValidation(type="list", formula1=formula, allow_blank=True,
                                   showErrorMessage=True)
        validacao.error = "Valor fora da lista permitida para esta coluna."
        validacao.errorTitle = "Valor invalido"
        planilha.add_data_validation(validacao)
        validacao.add(f"{coluna}2:{coluna}{ultima}")


def aba_revisar(livro, aud: Auditoria) -> None:
    planilha = livro.create_sheet("revisar")
    escrever_cabecalho(planilha, ["ID Pedido", "Tipo de pendencia", "Detalhe", "Acao sugerida"])
    for pend in sorted(aud.pendencias, key=lambda x: (x.tipo, str(x.id_pedido))):
        planilha.append([pend.id_pedido, pend.tipo, pend.detalhe, pend.acao])
        planilha.cell(planilha.max_row, 1).number_format = "0"
        for coluna in range(1, 5):
            celula = planilha.cell(planilha.max_row, coluna)
            celula.alignment = Alignment(wrap_text=True, vertical="top")
            celula.border = BORDA
    ajustar_larguras(planilha, {"A": 11, "B": 30, "C": 58, "D": 50})
    if planilha.max_row > 1:
        planilha.auto_filter.ref = f"A1:D{planilha.max_row}"


def aba_log(livro, aud: Auditoria) -> None:
    planilha = livro.create_sheet("log_correcoes")
    escrever_cabecalho(planilha, ["ID Pedido", "Coluna", "Valor anterior", "Valor corrigido", "Regra aplicada"])
    for c in aud.correcoes:
        planilha.append([
            c.id_pedido, c.coluna,
            "(vazio)" if c.antes is None else str(c.antes),
            "(vazio)" if c.depois is None else str(c.depois),
            c.regra,
        ])
        planilha.cell(planilha.max_row, 1).number_format = "0"
        for coluna in range(1, 6):
            celula = planilha.cell(planilha.max_row, coluna)
            celula.alignment = Alignment(wrap_text=True, vertical="top")
            celula.border = BORDA
    ajustar_larguras(planilha, {"A": 11, "B": 18, "C": 26, "D": 26, "E": 56})
    if planilha.max_row > 1:
        planilha.auto_filter.ref = f"A1:E{planilha.max_row}"


def aba_removidos(livro, aud: Auditoria) -> None:
    planilha = livro.create_sheet("removidos")
    escrever_cabecalho(planilha, ["ID Pedido", "Cliente", "Regiao", "Representante", "Data",
                                  "Receita", "Produto", "Status", "Motivo da remocao"])
    for p, motivo in aud.removidos:
        planilha.append([p.id, p.cliente, p.regiao, p.representante, p.data,
                         p.receita, p.produto, p.status, motivo])
        linha = planilha.max_row
        planilha.cell(linha, 1).number_format = "0"
        planilha.cell(linha, 5).number_format = FORMATO_DATA
        planilha.cell(linha, 6).number_format = FORMATO_MOEDA
        planilha.cell(linha, 9).alignment = Alignment(wrap_text=True, vertical="top")
    ajustar_larguras(planilha, {"A": 11, "B": 20, "C": 14, "D": 18, "E": 13,
                                "F": 14, "G": 22, "H": 17, "I": 52})


def aba_resumo(livro, pedidos: list[Pedido], aud: Auditoria, origem: Path, lidos: int) -> None:
    planilha = livro.create_sheet("resumo", 0)
    escrever_cabecalho(planilha, ["Indicador", "Valor"])

    ganhos = [p for p in pedidos if p.status == "Fechado Ganho" and p.receita is not None]
    perdidos = [p for p in pedidos if p.status == "Fechado Perdido" and p.receita is not None]
    a_revisar = [p for p in pedidos if p.pendencias]

    linhas = [
        ("Arquivo de origem", origem.name),
        ("Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Linhas lidas da origem", lidos),
        ("Linhas na planilha corrigida", len(pedidos)),
        ("Linhas removidas (duplicata exata)", len(aud.removidos)),
        ("Correcoes automaticas aplicadas", len(aud.correcoes)),
        ("Linhas que ainda exigem revisao", len(a_revisar)),
        ("Pendencias abertas", len(aud.pendencias)),
        ("Pedidos ganhos", len(ganhos)),
        ("Pedidos perdidos", len(perdidos)),
        ("Receita ganha (apos correcoes)", sum(p.receita for p in ganhos)),
        ("Receita perdida", sum(p.receita for p in perdidos)),
    ]
    for rotulo, valor in linhas:
        planilha.append([rotulo, valor])
        if isinstance(valor, float):
            planilha.cell(planilha.max_row, 2).number_format = FORMATO_MOEDA
        planilha.cell(planilha.max_row, 1).border = BORDA
        planilha.cell(planilha.max_row, 2).border = BORDA

    if aud.pendencias:
        planilha.append([])
        planilha.append(["Pendencias por tipo", ""])
        planilha.cell(planilha.max_row, 1).font = Font(bold=True)
        for tipo, quantidade in Counter(p.tipo for p in aud.pendencias).most_common():
            planilha.append([f"   {tipo}", quantidade])

    if aud.alertas:
        planilha.append([])
        planilha.append(["Alertas de consistencia", ""])
        planilha.cell(planilha.max_row, 1).font = Font(bold=True)
        for alerta in aud.alertas:
            planilha.append([f"   {alerta}", ""])
            planilha.cell(planilha.max_row, 1).alignment = Alignment(wrap_text=True)

    ajustar_larguras(planilha, {"A": 58, "B": 22})


def gravar(caminho: Path, pedidos: list[Pedido], aud: Auditoria, origem: Path, lidos: int) -> None:
    livro = openpyxl.Workbook()
    livro.remove(livro.active)
    aba_pedidos(livro, pedidos)
    aba_revisar(livro, aud)
    aba_log(livro, aud)
    aba_removidos(livro, aud)
    aba_resumo(livro, pedidos, aud, origem, lidos)
    livro.active = 0
    caminho.parent.mkdir(parents=True, exist_ok=True)
    livro.save(caminho)


# ===========================================================================
# RELATORIO NO TERMINAL
# ===========================================================================

def relatar(pedidos: list[Pedido], aud: Auditoria, lidos: int, saida: Path) -> None:
    moeda = lambda v: f"R$ {v:,.2f}".replace(",", "~").replace(".", ",").replace("~", ".")
    ganhos = [p for p in pedidos if p.status == "Fechado Ganho" and p.receita is not None]
    a_revisar = [p for p in pedidos if p.pendencias]

    print(f"\n  Linhas lidas .................. {lidos}")
    print(f"  Linhas gravadas ............... {len(pedidos)}")
    print(f"  Removidas (duplicata exata) ... {len(aud.removidos)}")
    print(f"  Correcoes automaticas ......... {len(aud.correcoes)}")

    if aud.correcoes:
        print("\n  Correcoes por coluna:")
        for coluna, n in Counter(c.coluna for c in aud.correcoes).most_common():
            print(f"    {coluna:<20} {n:>3}")

    print(f"\n  Linhas que exigem revisao ..... {len(a_revisar)}")
    if aud.pendencias:
        for tipo, n in Counter(p.tipo for p in aud.pendencias).most_common():
            print(f"    {tipo:<34} {n:>3}")

    if aud.alertas:
        print("\n  Alertas de consistencia:")
        for alerta in aud.alertas:
            print(f"    - {alerta}")

    print(f"\n  Receita ganha apos correcoes .. {moeda(sum(p.receita for p in ganhos))}")
    print(f"  Arquivo gerado ................ {saida}\n")


# ===========================================================================
# MAIN
# ===========================================================================

def main() -> int:
    raiz = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Corrige as inconsistencias de planilha_vendas.xlsx e gera uma planilha nova.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--entrada", type=Path, default=raiz / "data" / "planilha_vendas.xlsx",
                        help="planilha de origem (nao e modificada)")
    parser.add_argument("--saida", type=Path, default=raiz / "data" / "planilha_vendas_corrigida.xlsx",
                        help="planilha corrigida a ser gerada")
    parser.add_argument("--aba", default=ABA_ORIGEM, help=f"aba de origem (padrao: {ABA_ORIGEM})")
    parser.add_argument("--sobrescrever", action="store_true",
                        help="sobrescreve a planilha de saida se ela ja existir")
    parser.add_argument("--imputar-receita", action="store_true",
                        help="estima a receita ausente pela mediana do produto "
                             "(altera o faturamento apurado; fica sinalizado)")
    args = parser.parse_args()

    if not args.entrada.exists():
        sys.exit(f"Arquivo de entrada nao encontrado: {args.entrada}")
    if args.saida.exists() and not args.sobrescrever:
        sys.exit(f"{args.saida} ja existe. Use --sobrescrever para substituir.")
    if args.saida.resolve() == args.entrada.resolve():
        sys.exit("A saida nao pode ser o mesmo arquivo da entrada.")

    print(f"\n  Lendo {args.entrada} (aba '{args.aba}')...")
    pedidos = carregar(args.entrada, args.aba)
    lidos = len(pedidos)
    pedidos, aud = processar(pedidos, imputar_receita=args.imputar_receita)

    gravar(args.saida, pedidos, aud, args.entrada, lidos)
    relatar(pedidos, aud, lidos, args.saida)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
