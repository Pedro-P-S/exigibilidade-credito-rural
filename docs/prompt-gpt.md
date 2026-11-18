# Prompts para um assistente de IA só de chat

O assistente não lê arquivos nem roda código. Você cola texto e cola de volta só nomes de colunas e resultados OK/FALHOU. **Nunca cole valores, nomes de clientes, nomes de servidores ou senhas.**

---

## 1. Entender a projeção que já existe

Rode antes `python exigibilidade.py estrutura <pasta da projeção atual> --mascarar "<nomes a esconder>"`. Leia o `saidas\estrutura_projecao.md` e apague o que ainda parecer confidencial. Depois cole:

```
Você vai me ajudar a documentar uma projeção de exigibilidade do crédito rural (MCR cap. 6) que hoje é feita em planilha e scripts. Abaixo está só a estrutura (abas, cabeçalhos, fórmulas e código), sem valores.

Responda em português, curto, numerado:
1. Para cada aba ou script: o que entra, o que calcula e o que sai, em 1 a 3 linhas.
2. A cadeia de cálculo da exigibilidade e da projeção, na ordem.
3. Premissas fixas escondidas em fórmulas (percentuais, prazos, constantes).
4. Pontos frágeis: valores digitados no meio de fórmulas, referências que podem quebrar, somas que podem contar algo duas vezes.
5. Uma tabela de-para entre essa estrutura e estes arquivos de entrada:
   operacoes.csv (id; fonte; categoria; saldo; taxa_aa; fator)
   parcelas.csv (id; data; valor)
   captacoes.csv (mes; dias_uteis; vsr_avista; vsr_poupanca; saldo_lca; tipo)
   aplicacoes.csv (mes; fonte; categoria; saldo_medio; fator; tipo)
   fonte = RO, POUP, LCA ou LIVRE. Categorias: RO: pronamp, pronamp_invest, pronaf, pronaf_ind_com, outras; POUP: credito_rural, cpr; LCA: credito_rural, faculdade; LIVRE: livre.

Não peça valores. Se algo não der para saber pela estrutura, liste como pergunta com a resposta que você recomenda.

<ESTRUTURA>
(cole aqui o estrutura_projecao.md)
</ESTRUTURA>
```

## 2. Montar as consultas que geram os arquivos de entrada

```
Escreva consultas SQL (Oracle) que gerem estes CSVs, a partir das tabelas e colunas abaixo (só nomes e tipos, sem dados):

operacoes.csv: id; fonte; categoria; saldo; taxa_aa; fator  (uma linha por operação com saldo na data-base)
parcelas.csv: id; data; valor  (parcelas futuras)
aplicacoes_realizado.csv: mes; fonte; categoria; saldo_medio; fator; tipo='R'  (saldo médio dos dias úteis do mês; no mês em que a operação liquida, usar o último saldo do mês anterior)
captacoes.csv: mes; dias_uteis; vsr_avista; vsr_poupanca; saldo_lca; tipo

Regras: fonte = RO, POUP, LCA ou LIVRE; categoria conforme a lista (RO: pronamp, pronamp_invest, pronaf, pronaf_ind_com, outras; POUP: credito_rural, cpr; LCA: credito_rural, faculdade; LIVRE: livre). Use :data_base como parâmetro.
Para cada consulta, dê também uma checagem que imprima só OK/FALHOU (ex.: soma do saldo do CSV = soma da tabela de origem), nunca o valor.
Se faltar um mapeamento (por exemplo, código de fonte para fonte/categoria), pergunte em lista numerada com a sua recomendação.

<TABELAS>
(cole nomes de tabelas e colunas, com tipos)
</TABELAS>
```

## 3. Explicar a diferença entre duas rodadas

Rode `python exigibilidade.py comparar <rodada A> <rodada B>` e cole só o bloco "Impacto B - A por fonte/categoria" **se os valores puderem sair do trabalho**; se não, cole só os nomes das fontes/categorias e a contagem por motivo.

```
Comparei duas rodadas da projeção operação a operação. Os motivos são: mudou_saldo, mudou_classificacao, so_na_A, so_na_B.
Me diga, em ordem, o que verificar para explicar a diferença e que consulta rodar em cada passo. Comece pela fonte que mais mudou.
```
