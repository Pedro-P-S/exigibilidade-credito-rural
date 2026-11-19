# Exigibilidade do Crédito Rural

Calcula o **realizado** e a **projeção** das exigibilidades do crédito rural na forma do **MCR - Documento 6**: Anexo II (Recursos Obrigatórios, MCR 6-2), Anexo III (Poupança Rural, MCR 6-4) e Anexo IV (LCA, MCR 6-7). DIR, Pronaf, Pronamp e limites ficam dentro de cada anexo.

Gera a **posição de cada mês** (média acumulada de julho até o mês, como o Sisex) do ano-safra atual e do seguinte inteiro, e dois **HTML**:
- **`diretores.html`**: uma linha por exigibilidade (Obrigatórios Pronaf, Obrigatórios Pronamp, Obrigatórios Geral, Poupança, LCA) com o excesso (+) ou a deficiência (−) até o mês e o fechamento de junho em destaque. Meses passados com o valor oficial (Sisex: os próprios códigos de excesso e deficiência), meses futuros com a projeção partindo do oficial. Posição que não veio no arquivo `oficial` usa o calculado, e o relatório avisa.
- **`operacional.html`**: **Por exigibilidade** (um bloco por exigibilidade: exigibilidade, DIR, saldo médio do mês, aplicação média acumulada, limites, excesso e deficiência; Poupança e LCA com as duas contas, total e mínimo em crédito rural), **Detalhe por código** (linhas dos anexos com o código do Documento 6) e **Parâmetros e simulação** (liberação, % de reembolso, prazo e DIR, recalcula na hora, salva os parâmetros e gera o `diretores.html` com os parâmetros da tela). Onde há valor oficial aparecem calculado, oficial e diferença. Cores separam passado, mês da data-base e projeção.

Programa: [exigibilidade.py](exigibilidade.py), arquivo único, só biblioteca padrão do Python (Excel `.xlsx` precisa do `openpyxl`). Regras: MCR Atualização nº 760 (22/10/2026). Todos os números de exemplo e dos testes são inventados. Prompts para um assistente de IA só de chat: [docs/prompt-gpt.md](docs/prompt-gpt.md).

## Passo a passo

1. Crie uma pasta, por exemplo `C:\exigibilidade`.
2. No GitHub, abra `exigibilidade.py`, clique em **Raw**, `Ctrl+A`, `Ctrl+C`. Cole no Bloco de Notas e salve como `exigibilidade.py` na pasta, em **UTF-8**.
3. Teste: no Spyder, abra o arquivo, aperte **F5** e digite `1`. No terminal: `python exigibilidade.py autoteste`. O `resumo_verificacao.txt` tem que terminar com `RESULTADO: TUDO OK`.
4. Exemplo: menu `2` ou `python exigibilidade.py exemplo`. Abra o `operacional.html` em `saidas\rodadas\exemplo_...\`. No topo tem que aparecer `CHECK simulação = Python: OK`.
5. Dados reais: monte uma pasta de entradas (formatos abaixo, mesmo nome de arquivo do `exemplo\`) e rode:
   ```
   python exigibilidade.py rodar C:\exigibilidade\entradas --data-base 2026-11-20
   ```
   Cada rodada vai para `saidas\rodadas\<data_hora>\` e nunca é sobrescrita.
6. Comparar duas rodadas operação a operação (o que mudou de saldo, código, entrou ou saiu):
   ```
   python exigibilidade.py comparar saidas\rodadas\<rodada_A> saidas\rodadas\<rodada_B>
   ```

## Arquivos de entrada (pasta de entradas)

CSV, TXT (separador `;` `,` tab ou `|`) ou XLSX (1ª aba). Número com vírgula ou ponto decimal; data `AAAA-MM-DD` ou `DD/MM/AAAA`; mês `AAAA-MM`.

| Arquivo | Obrigatório | Colunas | Uma linha por |
|---|---|---|---|
| `operacoes` | sim | `id`, `saldo` (na data-base), `taxa_aa` (% a.a.) e as colunas de classificação abaixo | operação |
| `parcelas` | sim | `id`, `data`, `valor` | parcela futura; a última liquida |
| `liberacoes` | não | `id`, `data`, `valor`; para operação nova também `taxa_aa` e a classificação (`data_contratacao` vazia = data da liberação) | liberação futura que vem da base |
| `saldos_diarios` | não (sem ele os meses passados ficam zero) | `id`, `data`, `saldo` e as colunas de classificação daquele dia | operação e dia útil, de 1º de julho do ano-safra até o dia útil antes da data-base. É daqui que sai o realizado: saldo médio do mês = soma dos dias úteis ÷ dias úteis, dia útil sem linha = zero (Doc 6, Anexo I, 4.6-b); código do mês = regra na linha do último dia útil até o dia 15 (Anexo I, 12.2-b) |
| `oficial` | não | `posicao`, `codigo`, `valor` (o valor do código na posição, como está no Sisex) | posição e código |
| `vsr` | sim | `mes` (ou `data`), `vsr_avista`, `vsr_poupanca` | apuração do VSR (realizada ou projetada) |
| `lca` | sim | `data` (saldo diário) ou `mes` (saldo médio do mês projetado), `saldo` | dia ou mês |
| `dir` | não | `mes`, `modalidade`, `captacao`, `aplicacao` (saldo médio do mês) | mês e modalidade |
| `cenario` | não | `tipo` (`liquidacao` ou `reclassificacao`), `id`, `data`, `codigo` (novo código; vazio = sai da exigibilidade) | ajuste numa operação |
| `feriados` | não | `data` | feriado (dia útil = segunda a sexta fora dos feriados) |
| `parametros.json` | não | salvo pelo botão do HTML | parâmetros da gerência |

**Colunas de classificação** (o programa acha o código do Documento 6 sozinho):

| Coluna | Valores |
|---|---|
| `fonte` | `RO`, `POUP`, `LCA` ou `LIVRE` (não cumpre exigibilidade) |
| `programa` | `pronaf`, `pronamp`, `demais` |
| `finalidade` | `custeio`, `investimento`, `comercializacao`, `industrializacao` |
| `data_contratacao` | data do contrato |
| `recursos` | `controlados` ou `nao_controlados` (Poupança e LCA) |
| `beneficiario` | `produtor` (padrão), `cooperativa`, `cooperativa_cooperados` (para atender cooperados), `cooperativa_servicos` (prestadora de serviços) |
| `linha` | vazio na operação comum. Casos especiais do MCR: `abc`, `pca`, `peq_medio`, `proagro`, `renegociacao`, `renegociacao_2238`, `renegociacao_2238_tn`, `renegociacao_4028`, `renegociacao_4802`, `res_4801`, `res_4802`, `res_5330`, `res_5247`, `res_5247_i`, `res_5247_ii_iii`, `desclassificada`, `excluida_subvencao`, `outras_fontes`, `exig_adicional`, `cpr`, `cpr_5330`, `cra`, `cda_wa`, `cdca`, `fundo_garantidor`, `mcr_6_8_6`, `bndes`, `armazenagem`, `semiarido`, `bovinos`, `pecuario`, `cultivo_protegido`, `avicultura_suinocultura_piscicultura`, `desconto_leite`, `cotas_partes`, `invest_acima_300mil`, `invest_ate_200mil`, `invest_ate_300mil`, `mcr_10_4_2_a`, `mcr_10_4_2_b`, `mcr_6_2_17_b`, `pronaf_condicoes_6_2`, `res_4259` |
| `codigo` | opcional. Se vier preenchido (por exemplo, do de-para de origem que já existe), vale ele, e o `classificacao.csv` mostra quando a regra daria outro |

**DIR, modalidades:** `geral`, `pronaf`, `pronamp`, `poup`, `lca_cr`, `lca_livre`. A captação soma na exigibilidade e a aplicação conta no cumprimento.

## Saídas de cada rodada

- `diretores.html` e `operacional.html` (acima).
- `classificacao.csv`: código de cada operação, se a instituição informa ou o Sisex calcula, e o motivo (`regra`, `informado`, `informado (regra daria X)`, `sem código no Documento 6`, `livre`).
- `codigos.csv`: todos os códigos com valor, posição a posição: aplicações, ponderadores (seção 4), exigibilidade, limites, deficiência e excesso. Cada um marcado `informado` ou `calculado`, com `calculado`, `oficial` e `diferenca` nas posições que vieram no arquivo `oficial`. Código que só existe num dos lados aparece com o outro em zero.
- `exigibilidade.csv`, `aplicacoes.csv`, `saldo_operacoes.csv`.

O relatório do Sisex lista as operações de cada código calculado (Doc 6, Anexo I, 12.1-e e 7.4-i), então dá para descer da diferença até a operação. O valor oficial é a posição acumulada; o saldo médio de cada mês sai da diferença entre posições seguidas, pesada pelos dias úteis (Anexo I, 4.6-c/d), e alimenta a série dos diretores (oficial nos meses passados, projeção nos futuros).

## Regras implementadas

- **Código de cada operação:** tirado do título de cada código de aplicação dos anexos II, III e IV (337 códigos, tabela `REGRAS_TXT` no programa): fonte (anexo), programa (seção: 3-A Pronaf, 3-C Geral, 3-D Pronamp), finalidade, período de contratação, teto de taxa, recursos controlados ou não, beneficiário e linha especial. Se mais de uma regra serve, vale a mais específica (período mais curto, depois menor teto de taxa). O autoteste confere que cada código tem regra própria.
- **Informado ou calculado:** a partir de 2025/26 o Sisex calcula os códigos de aplicação com os dados do Sicor (Doc 6, Anexo I, 12.1); os de safras anteriores a instituição informa. A coluna `tipo` segue a descrição de cada código no MCR.
- **Ponderadores:** cada código da seção 4 ("Valor correspondente a N% da média do saldo informado no código X") vira fator 1 + N% do código X. Ex.: 3.1.13.49-9 (Pronaf custeio 2026/27 até 3% a.a.) tem 37% (4.1.34.19-2).
- **Limites:** lidos da descrição de cada código de limite: ind./com. Pronaf 10% (3.1.13.48-2), investimento Pronamp 10% (3.1.41.39-9), pequenos e médios 10% com o excesso indo para a Geral (3.1.40.34-5 e 3.1.30.78-8), ABC+/RenovAgro 3,6% (3.1.30.87-4), PCA 2,4% (3.1.30.89-8), Res. 5.330 40% (3.1.12.09-8, 3.1.40.38-3, 3.1.32.01-6), Res. 2.238 60% da própria (3.1.30.53-7), CPR na Poupança 5% da própria (3.2.30.10-7), Res. 5.247 (3.2.20.29-6, 3.3.20.17-2, 3.3.40.09-7), Res. 4.259 (3.2.10.54-3, 3.2.10.72-5).
- **Saldo médio do mês:** soma dos saldos dos dias úteis ÷ dias úteis do mês. Conta zero antes da 1ª liberação e depois da liquidação (Doc 6, Anexo I, 4.6-b).
- **Código do mês:** o que a operação tinha no **dia 15** vale para o mês inteiro (Doc 6, Anexo I, 12.2-b). No realizado, nos códigos que o Sisex calcula (a partir de 2025/26, Anexo I, 12.1-a), a regra roda sobre a linha do `saldos_diarios` do último dia útil até o dia 15. Operação que não existe no dia 15 (o 12.2-b não trata): contratada depois do dia 15 fica com a fonte da contratação; encerrada antes do dia 15 fica com a fonte do dia 15 do mês anterior; contratada e encerrada no mesmo mês fica fora, como o Sisex faz. Nos códigos informados, o saldo de cada dia vai para o código da fonte daquele dia. Decisão registrada em [ADR 0006](docs/adr/0006-codigo-na-contratacao-e-no-encerramento.md).
- **Posição do mês:** média acumulada de julho até o mês, ponderada pelos dias úteis (Doc 6, Anexo I, 4.6-c/d).
- **Anexo II (RO):** própria = 31,5% × (média do VSR à vista de jul a jun do ano anterior − R$ 500 mi), cooperativas 13% em 2026/27 e 22% em 2027/28, isenta se ≤ R$ 10 mi. Pronaf 35%, Pronamp 50%, Geral = o resto, cada uma + captação DIR. Deficiência e excesso separados (códigos 5.1.x); o excesso de Pronaf e Pronamp cobre a Geral.
- **Anexo III (Poupança):** 70% da média do VSR; no mínimo 95% em crédito rural; exigibilidade líquida = total − aplicação DIR-Poup (2.2.50.00-9).
- **Anexo IV (LCA):** 60% × (média acumulada das LCA de junho até o mês anterior à posição − R$ 500 mi), isenta se ≤ R$ 10 mi; 45% em crédito rural e 55% de faculdade. O excesso de RO entra só na posição de junho (MCR 6-7-9-h).

**Premissas da gerência (não são norma do MCR):** a liberação da gerência (grupos RO-Pronaf, RO-Pronamp, RO-Demais, Poupança e LCA, por mês, em reais) entra no meio do mês, sem ponderador, e é paga em parcela única depois do número de meses informado (padrão 12); o % de reembolso (padrão 80%; campo geral, por grupo e por mês, o mais específico vale) deixa o que não foi pago no saldo até o fim do horizonte; só valem a partir do mês da data-base. O `parametros.json` salvo é o padrão da próxima rodada. Juros compostos por dia corrido. No mês da data-base, os dias úteis anteriores a ela vêm do `saldos_diarios`.

**Ainda não trata:** custo financeiro da deficiência (MCR 6-5); juros pela regra exata de cada contrato; ponderador dos códigos limitados 3.2.10.54-3 e 3.2.10.72-5 (Res. 4.259/2013). Dois códigos de limite citam no MCR um código que não existe e foram lidos pelo vizinho: 3.1.32.01-6 cita 3.1.30.00-9 (usado 3.1.32.00-9) e 3.2.30.10-7 cita 3.2.20.23-1 (usado 3.2.30.23-1).

## O que é público e o que não é

Só vão para o GitHub este README, o `exigibilidade.py` e o `docs/prompt-gpt.md`. O `.gitignore` bloqueia todo o resto (entradas, saídas, exemplos, notas). Nada de dado real, nome de sistema, tabela ou pessoa aqui.
