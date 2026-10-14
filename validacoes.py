#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Verificacao em tempo de execucao do resultado da limpeza.

Sao as regras que valem para qualquer arquivo, nao so para a planilha de
referencia: recebem o que o usuario acabou de enviar e respondem se a limpeza
entregou o que prometeu naquele arquivo.

testar_correcoes.py responde outra pergunta - se o script regrediu - e para isso
usa valores ancorados na planilha de referencia.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime

import corrigir_planilha as cp


SEVERIDADE_ORDEM = {"erro": 0, "alerta": 1, "ok": 2}


@dataclass
class Resultado:
    item: str          # a que item do diagnostico a checagem corresponde
    titulo: str
    ok: bool
    detalhe: str
    severidade: str = "erro"   # "erro" quando falha de verdade, "alerta" quando e so aviso

    @property
    def icone(self) -> str:
        if self.ok:
            return "ok"
        return "alerta" if self.severidade == "alerta" else "erro"


@dataclass
class Retrato:
    """O que foi medido no arquivo ANTES da limpeza, para comparar depois."""
    linhas: int
    ids: set
    receita_bruta: float
    ids_normalizados: set
    ids_sem_receita: set

    @classmethod
    def de(cls, pedidos: list) -> "Retrato":
        # Normaliza o ID pela mesma funcao do pipeline: comparar '1001' com 1001
        # acusaria falsa perda de linha em todo arquivo CSV.
        return cls(
            linhas=len(pedidos),
            ids={cls._id(p.id) for p in pedidos},
            receita_bruta=sum(cp.ler_receita(p.receita) or 0 for p in pedidos),
            ids_normalizados={cls._id(p.id) for p in pedidos if cls._mudou(p.id)},
            ids_sem_receita={cls._id(p.id) for p in pedidos
                             if cp.ler_receita(p.receita) is None},
        )

    @staticmethod
    def _id(valor):
        return cp.ler_id(valor)

    @staticmethod
    def _mudou(valor) -> bool:
        """O ID precisou ser normalizado? Vale para 1001.0 e para '1001'."""
        novo = cp.ler_id(valor)
        return novo != valor or type(novo) is not type(valor)


# ===========================================================================
# CHECAGENS
# ===========================================================================

def _ids(pedidos: list) -> set:
    return {p.id for p in pedidos}


def validar(antes: Retrato, limpos: list, aud) -> list[Resultado]:
    """Roda todas as checagens e devolve o laudo, na ordem do diagnostico."""
    r: list[Resultado] = []
    add = r.append

    # ---- Item 1 e 2: receita numerica -------------------------------------
    textuais = [p.id for p in limpos if isinstance(p.receita, str)]
    add(Resultado(
        "1 e 2", "Receita convertida para numero",
        not textuais,
        "Toda receita preenchida e numerica e recebe formato de moeda na celula."
        if not textuais else f"Ainda ha receita como texto nos pedidos {sorted(textuais)[:8]}",
    ))

    # ---- Item 3: caixa padronizada ----------------------------------------
    colisoes = {}
    for campo in ("cliente", "regiao", "representante", "status"):
        grupos = defaultdict(set)
        for p in limpos:
            valor = getattr(p, campo)
            if valor:
                grupos[cp.chave(valor)].add(valor)
        conflitos = [sorted(v) for v in grupos.values() if len(v) > 1]
        if conflitos:
            colisoes[campo] = conflitos
    add(Resultado(
        "3", "Caixa de texto padronizada",
        not colisoes,
        "Nenhum valor aparece em duas grafias que so diferem na caixa."
        if not colisoes else f"Grafias concorrentes em: {colisoes}",
    ))

    # ---- Item 4: regiao no dominio ----------------------------------------
    validas = set(cp.REGIOES_CANONICAS.values())
    fora = sorted({p.regiao for p in limpos if p.regiao and p.regiao not in validas})
    add(Resultado(
        "4", "Regiao dentro do dominio",
        not fora,
        f"Todas as regioes pertencem ao dominio: {', '.join(sorted(validas))}."
        if not fora else f"Fora do dominio: {fora}. Inclua no dominio ou corrija a origem.",
        severidade="alerta",
    ))

    # ---- Item 5: representante canonico -----------------------------------
    grupos = defaultdict(set)
    for p in limpos:
        if p.representante:
            grupos[cp.chave(p.representante)].add(p.representante)
    multi = {k: sorted(v) for k, v in grupos.items() if len(v) > 1}
    add(Resultado(
        "5", "Representante com grafia unica",
        not multi,
        f"{len(grupos)} representante(s), cada um sob uma unica grafia."
        if not multi else f"Grafias concorrentes: {multi}",
    ))

    # ---- Item 6: duplicata exata removida ---------------------------------
    chaves = defaultdict(list)
    for p in limpos:
        chaves[(cp.chave(p.cliente), cp.chave(p.produto), p.data, p.receita)].append(p.id)
    restantes = [ids for ids in chaves.values() if len(ids) > 1]
    add(Resultado(
        "6", "Duplicata exata removida",
        not restantes,
        f"{len(aud.removidos)} linha(s) movida(s) para a aba de removidos, com o motivo registrado."
        if not restantes else f"Ainda ha duplicata exata na saida: {restantes}",
    ))

    # ---- Item 7 e 8: status --------------------------------------------
    status_validos = set(cp.STATUS_CANONICO.values())
    sem_status = [p.id for p in limpos if not p.status]
    fora_dominio = sorted({p.status for p in limpos if p.status and p.status not in status_validos})
    em_obs = [p.id for p in limpos if cp.chave(p.observacoes) in cp.STATUS_CANONICO]
    add(Resultado(
        "7 e 8", "Status na coluna certa e no idioma certo",
        not (fora_dominio or em_obs),
        f"Status restrito a {', '.join(sorted(status_validos))}."
        + (f" {len(sem_status)} pedido(s) seguem sem status e viraram pendencia." if sem_status else "")
        if not (fora_dominio or em_obs)
        else f"Fora do dominio: {fora_dominio}. Status em Observacoes: {em_obs}",
    ))

    # ---- Item 9 e 10: territorio -----------------------------------------
    sem_regiao = [p.id for p in limpos if not p.regiao]
    sem_rep = [p.id for p in limpos if not p.representante]
    pendentes_territorio = {
        pend.id_pedido for pend in aud.pendencias
        if pend.tipo in {"Regiao ausente", "Representante ausente", "Territorio ausente"}
    }
    nao_explicados = [i for i in sem_regiao + sem_rep if i not in pendentes_territorio]
    add(Resultado(
        "9 e 10", "Territorio preenchido ou declarado pendente",
        not nao_explicados,
        "Regiao e representante foram imputados onde o territorio era inequivoco; "
        "o que sobrou virou pendencia explicita."
        if not nao_explicados
        else f"Lacuna de territorio sem pendencia correspondente: {nao_explicados}",
    ))

    # ---- Item 11: ID inteiro ---------------------------------------------
    # Texto tambem conta: um CSV entrega "1001", e o tipo errado faz a
    # deduplicacao e a cronologia passarem batido.
    decimais = [p.id for p in limpos
                if isinstance(p.id, float)
                or (isinstance(p.id, str) and p.id.strip().replace(".", "", 1).isdigit())]
    registrados = {c.id_pedido for c in aud.correcoes if c.coluna == "ID Pedido"}
    nao_logados = antes.ids_normalizados - registrados - {
        Retrato._id(p.id) for p, _ in aud.removidos}
    add(Resultado(
        "11", "ID do pedido como inteiro",
        not decimais and not nao_logados,
        (f"{len(antes.ids_normalizados)} ID(s) normalizado(s) para inteiro, todos no log."
         if antes.ids_normalizados else "Os IDs ja chegaram como inteiro.")
        if not decimais and not nao_logados
        else f"IDs ainda nao normalizados: {decimais}. "
             f"Conversoes fora do log: {sorted(nao_logados, key=str)}",
    ))

    # ---- Item 12: observacoes limpas -------------------------------------
    sujas = []
    for p in limpos:
        if not p.observacoes:
            continue
        for trecho in str(p.observacoes).split(";"):
            if cp.chave(trecho) in cp.MARCACOES_DE_AUDITORIA:
                sujas.append((p.id, trecho.strip()))
    add(Resultado(
        "12", "Observacoes sem marcacao de auditoria",
        not sujas,
        "A coluna guarda nota de negocio; as marcacoes de controle viraram pendencia estruturada."
        if not sujas else f"Marcacoes remanescentes: {sujas}",
    ))

    # ---- Item 13: duplicata suspeita preservada e sinalizada --------------
    suspeitas = defaultdict(list)
    for p in limpos:
        if p.receita is not None:
            suspeitas[(cp.chave(p.cliente), cp.chave(p.produto), p.receita)].append(p)
    pares = [g for g in suspeitas.values() if len(g) > 1 and len({x.data for x in g}) > 1]
    sinalizados = {pend.id_pedido for pend in aud.pendencias if pend.tipo == "Duplicata suspeita"}
    faltando = [p.id for grupo in pares for p in grupo if p.id not in sinalizados]
    add(Resultado(
        "13", "Duplicata suspeita sinalizada, nunca removida em silencio",
        not faltando,
        f"{len(pares)} par(es) suspeito(s) preservado(s) e marcado(s) para conferencia humana."
        if not faltando else f"Suspeita nao sinalizada nos pedidos {faltando}",
    ))

    # ---- Item 14: receita ausente nao inventada ---------------------------
    # Quem chegou sem receita termina vazio e sinalizado, ou imputado e
    # sinalizado. O que nao pode e um valor entrar calado no faturamento.
    sinalizadas = {pend.id_pedido for pend in aud.pendencias
                   if pend.tipo in {"Receita ausente", "Receita imputada"}}
    presentes = {p.id for p in limpos}
    caladas = [i for i in antes.ids_sem_receita if i in presentes and i not in sinalizadas]
    imputadas = [pend.id_pedido for pend in aud.pendencias if pend.tipo == "Receita imputada"]
    ainda_vazias = [p.id for p in limpos if p.receita is None]
    add(Resultado(
        "14", "Receita ausente nao foi inventada em silencio",
        not caladas,
        (f"{len(ainda_vazias)} pedido(s) seguem sem receita e viraram pendencia."
         + (f" {len(imputadas)} receberam valor estimado, cada um sinalizado como imputado."
            if imputadas else ""))
        if not caladas
        else f"Receita preenchida sem sinalizacao nos pedidos {sorted(caladas, key=str)}",
    ))

    # ---- Item 15: cronologia ----------------------------------------------
    numerados = sorted((p for p in limpos if isinstance(p.id, int) and p.data), key=lambda p: p.id)
    quebras = {b.id for a, b in zip(numerados, numerados[1:]) if b.data < a.data}
    apontadas = {pend.id_pedido for pend in aud.pendencias
                 if pend.tipo == "ID fora de ordem cronologica"}
    add(Resultado(
        "15", "Quebra de cronologia apontada",
        quebras <= apontadas,
        (f"{len(quebras)} quebra(s) de ordem entre ID e data apontada(s) para conferencia."
         if quebras else "A numeracao dos pedidos acompanha a ordem das datas.")
        if quebras <= apontadas else f"Quebra nao apontada: {sorted(quebras - apontadas)}",
        severidade="alerta",
    ))

    # ---- Item 16: territorio acumulado ------------------------------------
    territorio = defaultdict(set)
    for p in limpos:
        if p.representante and p.regiao:
            territorio[p.representante].add(p.regiao)
    acumulam = {rep for rep, regioes in territorio.items() if len(regioes) > 1}
    cobertos = {rep for rep in acumulam if any(rep in a for a in aud.alertas)}
    add(Resultado(
        "16", "Representante em mais de uma regiao alertado",
        acumulam == cobertos,
        (f"{len(acumulam)} representante(s) acumulam territorio e geraram alerta."
         if acumulam else "Cada representante atua em uma unica regiao.")
        if acumulam == cobertos else f"Sem alerta: {sorted(acumulam - cobertos)}",
        severidade="alerta",
    ))

    # ---- Integridade: nada perdido, nada criado ---------------------------
    saida = _ids(limpos) | {Retrato._id(p.id) for p, _ in aud.removidos}
    perdidos = antes.ids - saida
    criados = saida - antes.ids
    add(Resultado(
        "Integridade", "Nenhuma linha perdida nem criada",
        not perdidos and not criados,
        f"As {antes.linhas} linhas lidas estao na planilha limpa ou na aba de removidos."
        if not perdidos and not criados
        else f"Perdidas: {sorted(perdidos)}. Criadas do nada: {sorted(criados)}",
    ))

    mantida = sum(p.receita or 0 for p in limpos)
    removida = sum(p.receita or 0 for p, _ in aud.removidos)
    # A receita imputada e dinheiro que a origem nao tinha: sai da conta para a
    # conservacao seguir medindo que nenhum centavo real apareceu ou sumiu.
    imputada = sum(
        float(c.depois) for c in aud.correcoes
        if c.coluna == "Receita" and c.depois is not None and "imputada" in c.regra.lower()
    )
    diferenca = abs((mantida + removida - imputada) - antes.receita_bruta)
    add(Resultado(
        "Integridade", "Receita conservada",
        diferenca < 0.01,
        ("Mantida mais removida reproduz exatamente o total lido da origem."
         + (f" Fora os R$ {imputada:,.2f} imputados, que a origem nao tinha."
            if imputada else ""))
        if diferenca < 0.01
        else f"Divergencia de {diferenca:.2f} entre a origem e a saida",
    ))

    pend_ids = {pend.id_pedido for pend in aud.pendencias}
    marcados = {p.id for p in limpos if p.pendencias}
    add(Resultado(
        "Integridade", "Toda pendencia marcada na planilha",
        pend_ids == marcados,
        f"{len(marcados)} linha(s) marcada(s) para revisao, exatamente as que tem pendencia."
        if pend_ids == marcados
        else f"Divergencia entre pendencias e marcacao: {pend_ids ^ marcados}",
    ))

    return r


def resumir(resultados: list[Resultado]) -> dict:
    return {
        "total": len(resultados),
        "aprovados": sum(1 for x in resultados if x.ok),
        "erros": sum(1 for x in resultados if not x.ok and x.severidade == "erro"),
        "alertas": sum(1 for x in resultados if not x.ok and x.severidade == "alerta"),
        "tudo_ok": all(x.ok for x in resultados),
    }
