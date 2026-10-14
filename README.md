# AnalisaDeDados

Limpeza automatizada de planilhas de vendas. Você envia um CSV ou Excel, o
sistema corrige o que dá para corrigir sozinho, separa o que exige decisão
humana e devolve a planilha limpa com um log de tudo que mudou.

## Instalação

```bash
pip install -r requirements.txt
```

Requer Python 3.9 ou superior.

## Usando

### App web

```bash
streamlit run app.py
```

Abre em `http://localhost:8501`. Envie o arquivo e baixe o resultado em CSV ou
Excel. A tela mostra o comparativo antes e depois, as correções aplicadas, as
pendências e o resultado das verificações.

### Linha de comando

```bash
python3 corrigir_planilha.py
python3 corrigir_planilha.py --entrada data/vendas.xlsx --saida data/limpa.xlsx
```

O arquivo de origem nunca é modificado.

### Como biblioteca

```python
import limpeza_pandas as lp

r = lp.limpar_arquivo("data/planilha_vendas.xlsx")

r.df          # DataFrame limpo
r.correcoes   # o que mudou em cada célula
r.pendencias  # o que exige decisão humana
r.laudo       # as verificações

lp.resumo_por(r.df, "Regiao")
lp.serie_mensal(r.df)
```

## O que é corrigido automaticamente

- Receita gravada como texto (`"R$ 4320.00"`) vira número com formato de moeda
- Caixa irregular em cliente, região, representante e status
- Região e representante com grafias concorrentes
- Status gravado na coluna errada ou em outro idioma (`closed won`)
- Região e representante ausentes, quando o território é inequívoco
- ID do pedido como decimal (`1001.0`) vira inteiro
- Duplicata exata, movida para uma aba separada com o motivo
- Observações usada como campo multiuso, separando nota de negócio de
  marcação de controle

## O que não é corrigido sozinho

Estes casos vão para a aba `revisar`, porque qualquer escolha automática
mudaria o resultado do negócio:

- **Duplicata suspeita** — mesmo cliente, produto e valor em datas diferentes.
  Pode ser recompra legítima.
- **Receita ausente** — o valor não existe no arquivo. Há uma opção para
  estimá-lo pela mediana do produto, desligada por padrão: ela altera o
  faturamento apurado.
- **ID fora de ordem cronológica** — não dá para saber a intenção da numeração.
- **Representante em mais de uma região** — pode ser acúmulo legítimo de
  território.

## A planilha gerada

Cinco abas: `resumo` com os indicadores, `pedidos_vendas` com os dados limpos,
`revisar` com as pendências, `log_correcoes` com o valor antes, o depois e a
regra de cada alteração, e `removidos` com as duplicatas e o motivo.

A aba principal traz listas suspensas em região, representante, produto e
status, para a caixa irregular não voltar na próxima digitação.

## Testes

```bash
./executar_testes.sh
```

Duas suítes, 92 testes, sem pytest:

- `testar_correcoes.py` — o motor de limpeza, com um teste por item do
  diagnóstico, mais integridade e interface de linha de comando
- `testar_app.py` — leitura de arquivo, detecção de colunas, camada pandas,
  exportação e a renderização da interface

## Organização

| Arquivo | Papel |
|---|---|
| `app.py` | Interface web, sem regra de negócio |
| `limpeza_pandas.py` | API pandas: DataFrame entra, DataFrame sai |
| `corrigir_planilha.py` | As regras de correção e a linha de comando |
| `planilha_io.py` | Leitura e escrita de CSV e Excel |
| `validacoes.py` | Verificações do resultado, rodadas a cada limpeza |

O pandas cuida de leitura, tipagem, agregação e análise. As regras seguem linha
a linha porque o log de correção é por célula, e vetorizar custaria essa
rastreabilidade.
