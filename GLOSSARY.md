# Exigibilidade do Crédito Rural

Apuração e projeção do cumprimento das exigibilidades do crédito rural na forma do MCR cap. 6 e do MCR - Documento 6 (anexos II, III e IV). A fonte de verdade de toda regra é o MCR; o que não está nele é premissa gerencial.

## Exigibilidades

**Exigibilidade**:
Valor que a instituição tem de manter aplicado em crédito rural com recursos de uma fonte. Uma por anexo: Recursos Obrigatórios (Anexo II, MCR 6-2), Poupança Rural (Anexo III, MCR 6-4) e LCA (Anexo IV, MCR 6-7).
_Avoid_: obrigação, meta

**Fonte**:
De onde vem o recurso da operação: `RO` (Recursos Obrigatórios), `POUP` (Poupança Rural), `LCA` ou `LIVRE` (não cumpre exigibilidade). Cada fonte, menos `LIVRE`, corresponde a um anexo.
_Avoid_: origem (é um campo da base do trabalho, não a fonte)

**Exigibilidade própria**:
Parte da exigibilidade calculada sobre a base da instituição, antes da captação DIR (códigos 2.1.10.00-8, 2.2.10.00-1 e 2.3.10.00-4).

**Subexigibilidade**:
Parcela da exigibilidade com destino obrigatório: Pronaf (35%) e Pronamp (50%) no Anexo II, crédito rural (95%) no Anexo III e crédito rural (45%) no Anexo IV. A Geral é o que sobra no Anexo II.
_Avoid_: sublimite

**Exigibilidade Geral**:
O que resta da exigibilidade de Recursos Obrigatórios depois de Pronaf e Pronamp (código 2.1.00.40-3).

**Direcionamento**:
Nome que o MCR dá à exigibilidade da LCA (Anexo IV).

**Faculdade**:
Parte da exigibilidade que pode ser cumprida com aplicações que não são crédito rural direto: CPR na Poupança, e CPR, CRA, CDCA e outros títulos na LCA.

**DIR**:
Depósito Interfinanceiro Vinculado ao Crédito Rural. A captação soma na exigibilidade de quem capta e a aplicação conta no cumprimento de quem aplica; uma modalidade por subexigibilidade (`geral`, `pronaf`, `pronamp`, `poup`, `lca_cr`, `lca_livre`).

**Deficiência**:
Quanto faltou aplicar numa exigibilidade ou subexigibilidade na posição (códigos 5.x.x1).
_Avoid_: falta, déficit

**Excesso**:
Quanto foi aplicado além da exigibilidade ou subexigibilidade na posição (códigos 5.x.x2). O excesso de Recursos Obrigatórios ao final do período conta na LCA (MCR 6-7-9-h).
_Avoid_: sobra, superávit

## Bases e períodos

**Ano-safra**:
Período de julho a junho. É o período de cumprimento das exigibilidades.
_Avoid_: ano agrícola, exercício

**Período de cálculo**:
Janela da base da exigibilidade. Recursos Obrigatórios e Poupança: julho a junho do ano-safra anterior. LCA: do primeiro dia útil de junho até o último dia útil do mês anterior à posição (Anexo IV, 1.3.10.00-5).

**VSR**:
Valor Sujeito a Recolhimento: base dos recursos à vista (Recursos Obrigatórios) e da poupança rural (Poupança).

**Posição**:
A situação de um mês no Documento 6: média acumulada dos saldos de julho até o mês, ponderada pelos dias úteis (Anexo I, 4.6-c/d).

**Saldo médio do mês**:
Soma dos saldos diários dos dias úteis do mês dividida pelos dias úteis; zero antes da primeira liberação e depois da liquidação (Anexo I, 4.6-b).

**Data-base**:
Data da foto da carteira. Antes dela é realizado, a partir dela é projeção.

**Realizado**:
Valor de cada código nos meses antes da data-base, calculado pelo programa a partir da base de operações.

**Valor oficial**:
Valor de cada código informado ou calculado no Sisex. Fica ao lado do realizado, com a diferença, para validar um contra o outro.
_Avoid_: realizado (o realizado é o calculado)

**Projeção**:
Valor de cada código a partir da data-base, calculado operação a operação com parcelas, liberações e premissas gerenciais.

## Códigos do Documento 6

**Código de aplicação**:
Código dos anexos II, III ou IV em que a operação é informada, conforme fonte, programa, finalidade, período de contratação e demais condições do título do código.
_Avoid_: categoria, rubrica

**Código informado**:
Código cujo valor a instituição envia ao Sisex.

**Código calculado**:
Código que o Sisex preenche sozinho; a partir de 2025/26 inclui os códigos de aplicação, calculados com os dados do Sicor (Anexo I, 12.1).

**Código de limite**:
Código calculado que soma outros códigos de aplicação até um percentual de uma base (ex.: 3.1.41.39-9, investimento Pronamp até 10% da subexigibilidade Pronamp).

**Ponderador**:
Acréscimo de um percentual sobre a média de um código de aplicação, informado na seção 4 do anexo (ex.: 37% sobre 3.1.13.49-9 no código 4.1.34.19-2).
_Avoid_: fator de ponderação, peso

**Código do mês**:
Código em que a operação estava no dia 15; vale para o mês inteiro (Anexo I, 12.2-b).

## Operação

**Operação**:
Um financiamento de crédito rural, com saldo, parcelas e liberações.
_Avoid_: contrato, cédula

**Programa**:
`pronaf`, `pronamp` ou `demais`.

**Finalidade**:
`custeio`, `investimento`, `comercializacao` ou `industrializacao`.

**Recursos controlados**:
Operação com taxa e condições definidas pelo MCR para a fonte; o oposto é recursos não controlados (condições livres). Só separa códigos na Poupança e na LCA.

**Linha especial**:
Condição do MCR que leva a operação a um código próprio: renegociação, Proagro, desclassificação parcial, ABC, PCA, CPR, Res. 5.330 e outras.

**Classificação**:
Atribuição do código de aplicação a uma operação pelas regras tiradas do MCR.
_Avoid_: de-para

**Liberação**:
Entrada de dinheiro numa operação. Liberação da base: futura e já prevista no sistema. Liberação da gerência: valor por grupo e mês informado pela gerência. Liberação passada não entra, porque já está no saldo.

## Gerência

**Premissa gerencial**:
Suposição da gerência que não é norma do MCR: liberação da gerência, % de reembolso, prazo de retorno, juros da projeção.
_Avoid_: regra

**Grupo da gerência**:
Agrupamento em que a gerência informa liberação e % de reembolso: RO-Pronaf, RO-Pronamp, RO-Demais, Poupança e LCA.

**% de reembolso**:
Parte das parcelas previstas no mês que a gerência espera receber; o resto fica no saldo.

**Rodada**:
Uma execução do programa com uma data-base e um conjunto de entradas; nunca é sobrescrita.
