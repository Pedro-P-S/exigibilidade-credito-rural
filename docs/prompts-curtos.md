# Prompts curtos para o GPT do trabalho

Link deste arquivo para copiar: https://raw.githubusercontent.com/Pedro-P-S/exigibilidade-credito-rural/main/docs/prompts-curtos.md

Cole um bloco por mensagem, na ordem. Não cole dados nem valores, só nomes de tabelas e colunas.

## 0. Regras (1ª mensagem do chat)

```
Regras para todo este chat:
1. Responda só o que eu pedir. Sem introdução, sem explicação, sem resumo, sem sugestões extras.
2. Quando eu pedir SQL: devolva só um bloco SQL Oracle. SELECT ou WITH, só leitura, sem ponto e vírgula no fim, colunas renomeadas com AS exatamente como eu pedir.
3. Use só os parâmetros :data_base, :inicio_safra e :inicio_anterior (datas).
4. Quando eu pedir checagem: devolva só um SQL que retorne OK ou FALHOU (ou só um número).
5. Se faltar informação: no máximo 1 pergunta curta, já com a sua recomendação.
6. Nunca me peça dados, valores ou nomes de clientes.
Responda só "ok".
```

## 1. Tabelas (2ª mensagem)

```
Tabelas e colunas que existem (só nomes):
[cole aqui nome da tabela: coluna, coluna, ...]
Responda só "ok".
```

## 2. operacoes

```
SQL do arquivo operacoes: uma linha por operação com saldo em :data_base.
Colunas: id, saldo, taxa_aa (% a.a.), fonte (RO, POUP, LCA ou LIVRE), programa (pronaf, pronamp, demais), finalidade (custeio, investimento, comercializacao, industrializacao), data_contratacao, recursos (controlados, nao_controlados), beneficiario (produtor, cooperativa, cooperativa_cooperados, cooperativa_servicos), linha (vazio na operação comum).
Converta os valores do banco para esses textos com CASE. Finalidade no padrão do Bacen: 1 = comercializacao, 2 = custeio, 3 = investimento, 4 = industrializacao.
```

## 3. saldos_diarios

```
SQL do arquivo saldos_diarios: uma linha por operação e dia útil, de :inicio_safra até o dia útil antes de :data_base.
Colunas: id, data, saldo, e as mesmas colunas de classificação do operacoes, com o valor daquele dia (use o campo de fonte do histórico diário, não o de hoje).
Sem linha em sábado, domingo ou dia sem saldo.
```

## 4. parcelas

```
SQL do arquivo parcelas: parcelas em aberto com vencimento depois de :data_base.
Colunas: id, data (vencimento), valor.
```

## 5. vsr

```
SQL do arquivo vsr: uma linha por mês de apuração.
Colunas: mes (primeiro dia do mês), vsr_avista, vsr_poupanca.
```

## 6. lca

```
SQL do arquivo lca: saldo diário de LCA de :inicio_anterior até o dia útil antes de :data_base.
Colunas: data, saldo.
```

## 7. Opcionais (só se existir no banco)

```
SQL do arquivo liberacoes: liberações futuras depois de :data_base. Colunas: id, data, valor, taxa_aa e as colunas de classificação do operacoes.
```

```
SQL do arquivo feriados: feriados de :inicio_anterior em diante. Coluna: data.
```

## 8. Checagens (uma por mensagem, depois de cada SQL)

```
Checagem: o SQL do saldos_diarios tem linha duplicada (mesmo id e data)? Devolva SQL que retorne OK ou FALHOU.
```

```
Checagem: no último dia útil antes de :data_base, a soma do saldo do saldos_diarios difere mais de 1% da soma do operacoes? Devolva SQL que retorne OK ou FALHOU.
```

```
Checagem: alguma coluna de classificação saiu com valor fora da lista que eu dei? Devolva SQL que retorne OK ou FALHOU.
```
