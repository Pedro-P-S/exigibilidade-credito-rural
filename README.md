# Exigibilidade do Crédito Rural

Calcula o **realizado** e a **projeção** dos saldos aplicados em crédito rural frente às exigibilidades do Manual de Crédito Rural (MCR, capítulo 6): Recursos Obrigatórios, Poupança Rural e LCA, com as subexigibilidades.

- Programa: [exigibilidade.py](exigibilidade.py), arquivo único, só biblioteca padrão do Python (o comando `estrutura` usa `openpyxl` para ler Excel).
- Regras: MCR Atualização nº 760 (22/10/2026).
- Prompt para usar com um assistente de IA só de chat: [docs/prompt-gpt.md](docs/prompt-gpt.md).
- Todos os números de exemplo e dos testes são inventados.

## O que faz

| Comando | O que faz |
|---|---|
| `autoteste` | Testa as regras com dados inventados. Grava `resumo_verificacao.txt`, que tem que terminar com `RESULTADO: TUDO OK`. |
| `exemplo` | Gera dados inventados em `exemplo/` e roda o fluxo inteiro. |
| `projetar` | Projeta o saldo de cada operação dia a dia até a liquidação (parcelas, juros, cenários) e tira o saldo médio mensal. Cada rodada vai para uma pasta própria em `saidas/rodadas/` e nunca é sobrescrita. |
| `calcular` | Junta realizado e projetado e calcula exigido, aplicado e sobra/falta por fonte e subexigibilidade, mês a mês (média acumulada), até junho e no ano-safra seguinte. |
| `comparar` | Compara duas rodadas operação a operação: o que mudou de saldo, de fonte/categoria, o que entrou e o que saiu. |
| `estrutura` | Lê uma projeção existente (Excel, código) e gera um `.md` só com a estrutura: abas, cabeçalhos e fórmulas, sem valores. |

## Passo a passo

1. Crie uma pasta, por exemplo `C:\exigibilidade`.
2. No GitHub, abra `exigibilidade.py`, clique em **Raw**, `Ctrl+A`, `Ctrl+C`. Cole no Bloco de Notas e salve como `exigibilidade.py` na pasta, em **UTF-8**.
3. Teste. No Spyder, abra o arquivo e aperte **F5**: aparece um menu, digite `1`. No terminal: `python exigibilidade.py autoteste`.
4. Rode o exemplo: menu `2`, ou `python exigibilidade.py exemplo`. Veja os arquivos em `exemplo\` e `saidas\`.
5. Gere os arquivos de entrada a partir das suas consultas, no formato abaixo, e rode:
   ```
   python exigibilidade.py projetar operacoes.csv parcelas.csv --data-base 2026-11-20 --ate 2028-06-30 --cenario cenario.csv --feriados feriados.csv
   python exigibilidade.py calcular captacoes.csv aplicacoes_realizado.csv saidas\rodadas\<rodada>\aplicacoes_projetadas.csv
   python exigibilidade.py comparar saidas\rodadas\<rodada_A>\saldo_operacoes.csv saidas\rodadas\<rodada_B>\saldo_operacoes.csv
   ```
   No Spyder: `runfile('exigibilidade.py', args='autoteste')`.

## Formato dos arquivos de entrada

CSV com `;` ou `,`, número com vírgula ou ponto decimal, data `AAAA-MM-DD` ou `DD/MM/AAAA`. Mês é `AAAA-MM`.

**Fonte e categoria** (o que cada saldo cumpre):

| fonte | categorias |
|---|---|
| `RO` (Recursos Obrigatórios, MCR 6-2) | `pronamp`, `pronamp_invest`, `pronaf`, `pronaf_ind_com`, `outras` |
| `POUP` (Poupança Rural, MCR 6-4) | `credito_rural`, `cpr` |
| `LCA` (MCR 6-7) | `credito_rural`, `faculdade` |
| `LIVRE` (não cumpre exigibilidade) | `livre` |

| Arquivo | Colunas | Uma linha por |
|---|---|---|
| `operacoes.csv` | `id`, `fonte`, `categoria`, `saldo` (na data-base), `taxa_aa` (% a.a.), `fator` (ponderador do MCR, vazio = 1) | operação |
| `parcelas.csv` | `id`, `data`, `valor` | parcela futura; a última liquida a operação |
| `cenario.csv` | `tipo`, `id`, `data`, `valor`, `fonte`, `categoria`, `pct`, `taxa_aa` | ajuste manual (abaixo) |
| `feriados.csv` | `data` | feriado (dia útil = segunda a sexta fora dos feriados) |
| `captacoes.csv` | `mes`, `dias_uteis`, `vsr_avista`, `vsr_poupanca`, `saldo_lca`, `tipo` (`R` realizado, `P` projetado) | mês, desde o início do período de cálculo |
| `aplicacoes_*.csv` | `mes`, `fonte`, `categoria`, `saldo_medio`, `fator`, `tipo` | mês, fonte e categoria |

**Cenários** (`cenario.csv`, coluna `tipo`):

| tipo | preencha | efeito |
|---|---|---|
| `liberacao` | `id`, `data`, `valor`, `fonte`, `categoria`, `taxa_aa` | soma ao saldo; `id` novo cria a operação (as parcelas dela vão no `parcelas.csv`) |
| `pct_pagamento` | `data`, `pct`, `fonte` e `categoria` opcionais | a partir da data, só `pct`% de cada parcela é pago; o resto fica no saldo |
| `liquidacao` | `id`, `data` | zera a operação na data |
| `reclassificacao` | `id`, `data`, `fonte`, `categoria` | a partir da data o saldo conta na nova fonte/categoria |

## Regras implementadas

- **Saldo médio mensal:** média do saldo dos dias úteis. No mês em que a operação liquida, ela entra com o último saldo do mês anterior.
- **Recursos Obrigatórios:** 31,5% de (média do VSR à vista de julho a junho − R$ 500 mi); cooperativas 13% em 2026/27 e 22% em 2027/28. Isenta se ≤ R$ 10 mi. Subexigibilidade Pronamp 50% (até 10% dela com investimento Pronamp) e Pronaf 35% (até 10% dela com industrialização/comercialização).
- **Poupança Rural:** 70% da média do VSR; no mínimo 95% em crédito rural e até 5% em CPR.
- **LCA:** 60% de (média do saldo de LCA de junho a maio − R$ 500 mi), isenta se ≤ R$ 10 mi; no mínimo 45% em crédito rural, até 55% em faculdades, e o excesso de Recursos Obrigatórios também conta.
- **Ponderadores** (ex. 1,37 do Pronaf custeio): vêm na coluna `fator`.

**Ainda não trata:** custo financeiro da deficiência (MCR 6-5), limites de normas transitórias (MCR 6-8), limite de 10% de pequenos e médios no Pronamp, juros pela regra exata de cada contrato (usa juros compostos por dia corrido).

## O que é público e o que não é

Só vão para o GitHub este README, o `exigibilidade.py` e o `docs/prompt-gpt.md`. O `.gitignore` bloqueia todo o resto (dados, saídas, exemplos, notas). Nada de dado real, nome de sistema, tabela ou pessoa aqui.
