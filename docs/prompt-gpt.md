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
   operacoes (id; saldo; taxa_aa; colunas de classificação)
   parcelas (id; data; valor)
   liberacoes (id; data; valor; e, para operação nova, taxa_aa e colunas de classificação)
   saldos_diarios (id; data; saldo; colunas de classificação)  uma linha por operação por dia útil desde 1º de julho
   vsr (mes; vsr_avista; vsr_poupanca)
   lca (data ou mes; saldo)
   dir (mes; modalidade; captacao; aplicacao)  modalidade = geral, pronaf, pronamp, poup, lca_cr, lca_livre
   Colunas de classificação (operacoes e liberacoes): fonte (RO, POUP, LCA, LIVRE); programa (pronaf, pronamp, demais); finalidade (custeio, investimento, comercializacao, industrializacao); data_contratacao; recursos (controlados, nao_controlados); beneficiario (produtor, cooperativa, cooperativa_cooperados, cooperativa_servicos); linha (vazio ou caso especial do MCR: abc, pca, peq_medio, proagro, renegociacao, res_5330, desclassificada, cpr, cra, cdca etc.); codigo (opcional: código do Documento 6, se já existir).

Não peça valores. Se algo não der para saber pela estrutura, liste como pergunta com a resposta que você recomenda.

<ESTRUTURA>
(cole aqui o estrutura_projecao.md)
</ESTRUTURA>
```

## 2. Montar as consultas que geram os arquivos de entrada

```
Escreva um SELECT (Oracle) para cada um destes arquivos, a partir das tabelas e colunas abaixo (só nomes e tipos, sem dados). Cada SELECT vai colado no consultas_banco.ini do programa, que roda a consulta e grava o CSV: só leitura, sem ponto e vírgula no fim, colunas renomeadas com AS para os nomes abaixo, parâmetros :data_base, :inicio_safra (1º de julho do ano-safra) e :inicio_anterior (1º de julho do ano-safra anterior):

operacoes.csv: id; fonte; programa; finalidade; data_contratacao; recursos; beneficiario; linha; codigo; taxa_aa; saldo  (uma linha por operação com saldo na data-base)
saldos_diarios.csv: id; data; saldo; fonte; programa; finalidade; data_contratacao; recursos; beneficiario; linha; codigo; taxa_aa  (uma linha por operação por dia útil com saldo, de 1º de julho do ano-safra da :data_base até o dia útil anterior à :data_base; as colunas de classificação com o valor que a operação tinha NAQUELE dia, não o de hoje)
parcelas.csv: id; data; valor  (parcelas futuras)
liberacoes.csv: id; data; valor; fonte; programa; finalidade; data_contratacao; recursos; beneficiario; linha; codigo; taxa_aa  (liberações futuras já previstas na base; colunas de classificação só para operação nova)

Regras do saldos_diarios: dia útil = segunda a sexta fora dos feriados; dia sem saldo não precisa de linha (o programa conta zero); o programa usa a linha do último dia útil até o dia 15 de cada mês para achar o código do mês, por isso a classificação tem de ser a do dia, vinda do histórico, nunca a atual repetida. Se a base só guarda a classificação atual, diga isso em vez de inventar.

Valores aceitos: fonte RO, POUP, LCA ou LIVRE; programa pronaf, pronamp ou demais; finalidade custeio, investimento, comercializacao ou industrializacao; recursos controlados ou nao_controlados (só Poupança e LCA); beneficiario produtor, cooperativa, cooperativa_cooperados ou cooperativa_servicos; linha vazia, ou o caso especial do MCR (abc, pca, peq_medio, proagro, renegociacao, res_5330, desclassificada, cpr, cra, cdca e os outros do README). 
Para cada consulta, dê também uma checagem que imprima só OK/FALHOU, nunca o valor. No mínimo:
- soma do saldo do CSV = soma da tabela de origem;
- saldos_diarios sem linha repetida para o mesmo id e data;
- saldos_diarios sem linha em sábado, domingo ou feriado;
- soma do saldo do último dia útil antes da :data_base no saldos_diarios até 1% diferente da soma de operacoes.csv (a diferença normal é de juros e movimentos de um dia);
- nenhuma linha com fonte, programa, finalidade, recursos, beneficiario ou linha fora dos valores aceitos;
- quantas operações mudaram de programa, finalidade ou linha ao longo do período (só a contagem).

Renegociações (linha):
- linha = res_5330 nas operações de composição de dívidas rurais da MP 1.376/2026 (Resoluções CMN 5.330 e 5.340; as duas vão para os mesmos códigos do Documento 6); cpr_5330 nas CPR adquiridas pelo art. 6º da Res. 5.330.
- linha = res_5247 (Poupança), res_5247_i ou res_5247_ii_iii (LCA) nas operações para liquidar dívidas de produtores atingidos por eventos adversos da MP 1.314/2025 (Res. CMN 5.247), conforme o inciso do art. 2º.
- Se não houver campo que identifique essas renegociações, diga isso e pergunte qual campo usar, com a sua recomendação.
Checagens extras (OK/FALHOU ou contagem):
- toda operação com res_5247 na Poupança: a fonte da operação liquidada era Poupança Rural Livre ou Poupança Rural Controlados - Subvenção Econômica em 22/09/2025 (use o histórico diário de fonte) e a operação nova está em Poupança Rural Livre em todos os dias (MCR 6-8-33); dê só a contagem das que não cumprem;
- toda operação com res_5247_i ou res_5247_ii_iii está em LCA - Taxa Livre em todos os dias (MCR 6-8-34); só a contagem das que não cumprem.
Se faltar um mapeamento (por exemplo, de origem ou linha de crédito para programa, finalidade ou linha), pergunte em lista numerada com a sua recomendação.

<TABELAS>
(cole nomes de tabelas e colunas, com tipos)
</TABELAS>
```

## 3. Explicar a diferença entre duas rodadas

Rode `python exigibilidade.py comparar <rodada A> <rodada B>` e cole só o bloco "Impacto B - A por fonte/código" **se os valores puderem sair do trabalho**; se não, cole só os códigos e a contagem por motivo.

```
Comparei duas rodadas da projeção operação a operação (comando comparar). Os motivos são: mudou_saldo, mudou_classificacao, so_na_A, so_na_B.
Me diga, em ordem, o que verificar para explicar a diferença e que consulta rodar em cada passo. Comece pela fonte que mais mudou.
```

## 4. Conferir a classificação contra o de-para que já existe

Rode a projeção com a coluna `codigo` preenchida pelo de-para de origem que já existe. O `classificacao.csv` marca `informado (regra daria X)` onde a regra do MCR discorda. Conte as linhas por par (código do de-para, código da regra), sem id nem valor, e cole:

```
Comparei a classificação de operações de crédito rural no código do MCR Documento 6 feita por um de-para antigo com a feita por regras tiradas do título de cada código. Abaixo, cada par (código do de-para; código da regra; quantidade de operações).
Para cada par, diga em uma linha qual coluna da operação provavelmente explica a diferença (programa, finalidade, data de contratação, taxa, recursos, beneficiário ou linha especial) e que consulta fazer para confirmar. Não peça valores.

<PARES>
(cole aqui)
</PARES>
```
