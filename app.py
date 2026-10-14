#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
App web de limpeza automatizada de planilhas de vendas.

A tela segue uma ordem deliberada: o que o usuario precisa saber vem primeiro
(resumo e veredito), o que ele veio buscar vem em seguida (download), e todo o
detalhamento fica recolhido, a um clique de distancia.

Este arquivo e so a interface. O trabalho com dados fica em limpeza_pandas.py,
as regras em corrigir_planilha.py e as verificacoes em validacoes.py.

    streamlit run app.py
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

import corrigir_planilha as cp
import limpeza_pandas as lp
import planilha_io as pio

RAIZ = Path(__file__).resolve().parent
SUITE = RAIZ / "testar_correcoes.py"

NOMES_SEPARADOR = {";": "ponto e virgula", ",": "virgula",
                   "\t": "tabulacao", "|": "barra vertical"}

AZUL, AMBAR, VERDE = "#2563eb", "#d97706", "#059669"
MOEDA = st.column_config.NumberColumn(format="R$ %.2f")
DATA = st.column_config.DateColumn(format="DD/MM/YYYY")

CORES_SITUACAO = {"Resolvido": VERDE, "Reduzido": AZUL,
                  "Esperado": "#6b7280", "Sem defeito": "#6b7280", "Em aberto": AMBAR}


def reais(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "~").replace(".", ",").replace("~", ".")


# Em cache: sem isso, qualquer controle da tela reprocessaria o arquivo inteiro.

@st.cache_data(show_spinner=False)
def carregar(nome: str, dados: bytes, aba, codec, separador) -> pd.DataFrame:
    return lp.ler(dados, aba=aba, codec=codec, separador=separador, nome=nome)


@st.cache_data(show_spinner=False)
def processar(bruto: pd.DataFrame, mapa: dict, imputar: bool):
    r = lp.limpar(bruto, mapa, imputar_receita=imputar)
    return r, lp.diagnosticar(bruto, mapa), lp.diagnosticar(r.df)


@st.cache_data(show_spinner=False)
def rodar_suite() -> tuple[bool, str, str]:
    """Devolve (passou, resumo no estilo '48 passed in 0.78s', saida completa)."""
    if not SUITE.exists():
        return False, "", "testar_correcoes.py nao encontrado."
    proc = subprocess.run([sys.executable, str(SUITE), "-q"],
                          capture_output=True, text=True, cwd=RAIZ)
    saida = (proc.stdout + proc.stderr).strip()
    passou = proc.returncode == 0
    achou = re.search(r"Ran (\d+) tests? in ([\d.]+)s", saida)
    if achou:
        verbo = "passed" if passou else "failed"
        resumo = f"{achou.group(1)} {verbo} in {float(achou.group(2)):.2f}s"
    else:
        resumo = ""
    return passou, resumo, saida


st.set_page_config(page_title="Limpeza de Planilha de Vendas",
                   page_icon="🧹", layout="wide")

# ===========================================================================
# BARRA LATERAL
# ===========================================================================

with st.sidebar:
    st.header("Configuracoes")

    rodar_testes = st.toggle(
        "Rodar suite de testes ao final", value=True,
        help="Executa testar_correcoes.py, que valida o proprio script de limpeza "
             "contra a planilha de referencia. Independe do arquivo enviado.")

    imputar_receita = st.toggle(
        "Imputar receita ausente pela mediana do produto", value=False,
        help="Desligado por padrao: preencher receita que nao existe altera o "
             "faturamento apurado. Ligado, o valor entra marcado como imputado.")

    with st.expander("Leitura do arquivo"):
        auto = st.checkbox("Detectar codificacao e separador", value=True,
                           help="Desmarque se o arquivo abrir com acentuacao errada.")
        codec_manual = None if auto else st.selectbox(
            "Codificacao", ["utf-8-sig", "utf-8", "cp1252", "latin-1"])
        sep_manual = None if auto else st.selectbox(
            "Separador", [";", ",", "\t", "|"], format_func=NOMES_SEPARADOR.get)

    with st.expander("Formato do CSV de saida"):
        perfil = st.radio("Perfil", list(pio.PERFIS_CSV), label_visibility="collapsed",
                          help="O Excel em portugues espera ponto e virgula e virgula decimal.")
        separador, decimal, codec = pio.PERFIS_CSV[perfil]
        st.caption(f"Separador {NOMES_SEPARADOR[separador]}, decimal `{decimal}`, "
                   f"codificacao `{codec}`.")

    with st.expander("Colunas esperadas"):
        st.caption("Reconhecidas automaticamente, com variacoes de titulo.")
        for campo in pio.CAMPOS:
            marca = " *" if campo in pio.OBRIGATORIAS else ""
            st.markdown(f"- `{pio.ROTULOS[campo]}`{marca}")
        st.caption("* obrigatoria")

# ===========================================================================
# CABECALHO E ENVIO
# ===========================================================================

st.title("🧹 Limpeza Automatizada de Planilha de Vendas")
st.caption("Envie um CSV ou Excel de pedidos. O app diagnostica, corrige "
           "automaticamente e valida o resultado.")

enviado = st.file_uploader("Selecione a planilha de vendas", type=["csv", "xlsx"],
                           help="Uma linha por pedido, com os titulos na primeira linha.")

if not enviado:
    st.info("Faca upload de um CSV ou XLSX para comecar.")
    with st.expander("Quais titulos de coluna sao reconhecidos"):
        st.dataframe(
            pd.DataFrame([{"Coluna": pio.ROTULOS[c],
                           "Obrigatoria": "sim" if c in pio.OBRIGATORIAS else "nao",
                           "Titulos reconhecidos": ", ".join(cp.SINONIMOS_DE_COLUNA[c][:6])}
                          for c in pio.CAMPOS]),
            hide_index=True, width="stretch")
    st.stop()

dados = enviado.getvalue()
aba = None
if not enviado.name.lower().endswith(".csv"):
    try:
        abas = pio.abas_de(dados)
    except Exception as erro:
        st.error(f"Nao consegui abrir a planilha: {erro}")
        st.stop()
    aba = st.selectbox("Aba da planilha", abas) if len(abas) > 1 else abas[0]

try:
    bruto = carregar(enviado.name, dados, aba, codec_manual, sep_manual)
except Exception as erro:
    st.error(f"Nao consegui ler o arquivo: {erro}")
    st.stop()

if bruto.empty:
    st.warning("O arquivo nao tem nenhuma linha de dados abaixo do cabecalho.")
    st.stop()

# ===========================================================================
# ASSOCIACAO DE COLUNAS
# ===========================================================================

cabecalho = list(bruto.columns)
detectado = cp.detectar_colunas(cabecalho)
nao_detectadas = [pio.ROTULOS[c] for c in pio.OBRIGATORIAS if not detectado.get(c)]

rotulo = ("⚠️ Associacao de colunas — confira, nem tudo foi reconhecido"
          if nao_detectadas else "Associacao de colunas — tudo reconhecido")
with st.expander(rotulo, expanded=bool(nao_detectadas)):
    if nao_detectadas:
        st.warning("Nao reconheci: " + ", ".join(nao_detectadas) + ". Associe abaixo.")
    opcoes = ["(nenhuma)"] + cabecalho
    mapa = {}
    colunas_ui = st.columns(3)
    for i, campo in enumerate(pio.CAMPOS):
        atual = detectado.get(campo)
        with colunas_ui[i % 3]:
            escolha = st.selectbox(
                pio.ROTULOS[campo] + (" *" if campo in pio.OBRIGATORIAS else ""),
                opcoes, index=opcoes.index(atual) if atual in opcoes else 0,
                key=f"mapa_{campo}")
        mapa[campo] = None if escolha == "(nenhuma)" else escolha

faltando = [pio.ROTULOS[c] for c in pio.OBRIGATORIAS if not mapa.get(c)]
if faltando:
    st.error("Associe as colunas obrigatorias antes de continuar: " + ", ".join(faltando))
    st.stop()

# ===========================================================================
# RESUMO
# ===========================================================================

with st.spinner("Diagnosticando e corrigindo..."):
    r, antes, depois = processar(bruto, mapa, imputar_receita)

if r.df.empty:
    st.error("Nenhuma linha util encontrada depois da associacao de colunas.")
    st.stop()

falhas = int((r.laudo["Situacao"] == "Falha").sum())
alertas = int((r.laudo["Situacao"] == "Alerta").sum())
aprovados = int(r.laudo["ok"].sum())

st.divider()
st.subheader("Resumo")

m = st.columns(4)
m[0].metric("Registros (entrada)", r.linhas_lidas)
m[1].metric("Registros (saida)", len(r.df),
            delta=f"-{len(r.removidos)}" if len(r.removidos) else None)
m[2].metric("Erros corrigidos", r.total_melhorias)
m[3].metric("Pendencias para decidir", r.total_pendencias,
            delta=f"{r.linhas_a_revisar} linha(s)" if r.linhas_a_revisar else None,
            delta_color="inverse")

if falhas:
    st.error(f"Limpeza concluida, mas {falhas} verificacao(oes) falhou(aram). "
             f"Confira a secao Validacao antes de usar o arquivo.")
elif alertas:
    st.warning(f"Limpeza concluida — {aprovados} de {len(r.laudo)} verificacoes "
               f"aprovadas, {alertas} com alerta.")
else:
    st.success(f"Limpeza concluida e aprovada nas {len(r.laudo)} verificacoes. "
               f"{r.total_melhorias} correcao(oes) em {r.linhas_lidas} registro(s)"
               + (f", {r.total_pendencias} pendencia(s) para voce decidir."
                  if r.total_pendencias else "."))

# ===========================================================================
# DOWNLOAD
# ===========================================================================

base = Path(enviado.name).stem
d = st.columns(2)
d[0].download_button(
    "📄 Baixar CSV", data=lp.para_csv(r, perfil),
    file_name=f"{base}_limpo.csv", mime="text/csv", width="stretch",
    help=f"Separador {NOMES_SEPARADOR[separador]}, decimal '{decimal}', codificacao {codec}.")
d[1].download_button(
    "📊 Baixar Excel", data=lp.para_excel(r, enviado.name),
    file_name=f"{base}_limpo.xlsx", type="primary", width="stretch",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Cinco abas: resumo, dados limpos, pendencias, log de correcoes e removidos.")

# ===========================================================================
# ANTES E DEPOIS
# ===========================================================================

st.divider()
st.subheader("Antes e depois")

comparativo = lp.comparativo(antes, depois)
st.dataframe(
    comparativo, hide_index=True, width="stretch",
    column_config={
        "Indicador": st.column_config.TextColumn(width="medium"),
        "Antes": st.column_config.NumberColumn(width="small"),
        "Depois": st.column_config.NumberColumn(width="small"),
        "Variacao": st.column_config.NumberColumn(width="small", format="%+d"),
        "Situacao": st.column_config.TextColumn(width="small"),
    })
st.caption("**Esperado** marca o que nao deve zerar: duplicata apenas suspeita fica "
           "na planilha para sua decisao, e um unico formato de data ja e o estado correto.")

with st.expander("Detalhe por indicador"):
    esq, dir_ = st.columns(2)
    for coluna, titulo, medida in ((esq, "Antes", antes), (dir_, "Depois", depois)):
        with coluna:
            st.markdown(f"**{titulo}**")
            formatos = medida["detalhes"].get("formatos de data") or []
            st.markdown("Formatos de data: "
                        + (", ".join(f"`{f}`" for f in formatos) or "nenhum"))
            caixa = medida["detalhes"].get("caixa por coluna") or {}
            if caixa:
                st.dataframe(
                    pd.DataFrame([{"Coluna": k, "Celulas fora do padrao": v}
                                  for k, v in caixa.items()]),
                    hide_index=True, width="stretch")
            else:
                st.caption("Nenhuma celula com caixa fora do padrao.")

# ===========================================================================
# VALIDACAO
# ===========================================================================

st.divider()
st.subheader("🧪 Validacao automatizada")

v = st.columns(2)
with v[0]:
    st.markdown("**Este arquivo**")
    if falhas:
        st.error(f"{falhas} de {len(r.laudo)} verificacoes falharam.")
    elif alertas:
        st.warning(f"{aprovados} de {len(r.laudo)} verificacoes aprovadas, "
                   f"{alertas} com alerta.")
    else:
        st.success(f"Todas as {len(r.laudo)} verificacoes passaram.")
    st.caption("Confere, no arquivo que voce enviou, se a limpeza entregou o "
               "que prometeu.")

with v[1]:
    st.markdown("**Suite de regressao**")
    if not rodar_testes:
        st.info("Desligada nas configuracoes.")
    else:
        with st.spinner("Rodando a suite..."):
            passou, resumo, saida = rodar_suite()
        if passou:
            st.success(f"Todos os testes passaram. {resumo}")
        else:
            st.error(f"A suite falhou. {resumo}")
    st.caption("Valida o script de limpeza contra a planilha de referencia do "
               "projeto, nao contra o arquivo enviado.")

with st.expander(f"Detalhe das verificacoes ({aprovados}/{len(r.laudo)})"):
    st.dataframe(r.laudo[["Item", "Verificacao", "Situacao", "Detalhe"]],
                 hide_index=True, width="stretch", height=300)

if rodar_testes:
    with st.expander("Saida completa da suite"):
        st.code(saida or "(sem saida)", language="text")

# ===========================================================================
# DETALHAMENTO
# ===========================================================================

st.divider()

with st.expander(f"🔧 Correcoes aplicadas ({r.total_melhorias})"):
    if r.correcoes.empty:
        st.info("Nenhuma correcao foi necessaria.")
    else:
        por_regra = lp.correcoes_por_regra(r)
        esq, dir_ = st.columns([3, 2])
        with esq:
            st.dataframe(por_regra, hide_index=True, width="stretch")
        with dir_:
            st.bar_chart(por_regra.set_index("Correcao")["Quantidade"].sort_values(),
                         horizontal=True, color=AZUL, height=300)
        st.markdown("**Log completo — o que mudou em cada celula**")
        st.dataframe(r.correcoes, hide_index=True, width="stretch", height=300)

with st.expander(f"⚠️ Pendencias para voce decidir ({r.total_pendencias})"):
    if r.pendencias.empty:
        st.success("Nenhuma pendencia: o arquivo nao exige nenhuma decisao sua.")
    else:
        st.caption("O app nao resolve estes casos sozinho porque qualquer escolha "
                   "automatica poderia alterar o resultado do negocio.")
        esq, dir_ = st.columns([2, 3])
        with esq:
            st.dataframe(lp.pendencias_por_tipo(r), hide_index=True, width="stretch")
        with dir_:
            st.dataframe(r.pendencias, hide_index=True, width="stretch", height=260)
    for alerta in r.alertas:
        st.warning(alerta)

with st.expander(f"👀 Dados limpos ({len(r.df)})"):
    st.dataframe(r.df, hide_index=True, width="stretch", height=380,
                 column_config={"Receita": MOEDA, "Data do Pedido": DATA})
    if not r.removidos.empty:
        st.markdown(f"**🗑️ Duplicatas removidas ({len(r.removidos)})** — "
                    "preservadas aqui e na aba `removidos` do Excel.")
        st.dataframe(r.removidos, hide_index=True, width="stretch",
                     column_config={"Receita": MOEDA, "Data do Pedido": DATA})

with st.expander("📊 Analise da planilha limpa"):
    met = lp.metricas(r.df)
    k = st.columns(4)
    k[0].metric("Receita ganha", reais(met["receita_ganha"]))
    k[1].metric("Ticket medio", reais(met["ticket_medio"]),
                help=f"Media sobre os {met['ganhos_com_receita']} pedidos ganhos com "
                     f"receita registrada. Pedido sem valor nao entra como zero.")
    k[2].metric("Taxa de ganho", f"{met['taxa_de_ganho']:.0%}",
                help=f"{met['ganhos']} ganhos de {met['ganhos'] + met['perdidos']} decididos.")
    k[3].metric("Pedidos sem receita", met["sem_receita"])

    if r.linhas_a_revisar:
        st.caption(f"{r.linhas_a_revisar} linha(s) ainda tem pendencia aberta, "
                   f"entao estes numeros podem mudar depois da sua revisao.")

    mensal = lp.serie_mensal(r.df)
    if len(mensal) > 1:
        st.markdown("**Receita ganha por mes**")
        st.line_chart(mensal.set_index("Mes")["Receita ganha"], color=VERDE, height=220)

    dimensao = st.radio("Agrupar por", ["Regiao", "Representante", "Produto"],
                        horizontal=True, key="dimensao")
    resumo_dim = lp.resumo_por(r.df, dimensao)
    if not resumo_dim.empty:
        st.dataframe(resumo_dim, hide_index=True, width="stretch", column_config={
            "Receita ganha": MOEDA, "Ticket medio": MOEDA,
            "Taxa de ganho": st.column_config.ProgressColumn(
                format="%.0f%%", min_value=0, max_value=1)})
