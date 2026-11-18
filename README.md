# Exigibilidade do Crédito Rural

Calcula o **realizado** e a **projeção** das exigibilidades do crédito rural na forma do **MCR - Documento 6**: Anexo II (Recursos Obrigatórios, MCR 6-2), Anexo III (Poupança Rural, MCR 6-4) e Anexo IV (LCA, MCR 6-7). DIR, Pronaf, Pronamp e limites ficam dentro de cada anexo.

Gera a **posição de cada mês** (média acumulada de julho até o mês, como o Sisex) do ano-safra atual e do seguinte inteiro, e um **HTML** com:
- **Visão geral** (diretores): sobra ou falta no fechamento de junho e mês a mês;
- **Detalhe dos anexos** (gestores): as linhas com os códigos do Documento 6;
- **Parâmetros e simulação** (gestores): liberação, % de reembolso, prazo e DIR por mês. Recalcula na hora e salva os parâmetros para a próxima rodada.

Programa: [exigibilidade.py](exigibilidade.py), arquivo único, só biblioteca padrão do Python (Excel `.xlsx` precisa do `openpyxl`). Regras: MCR Atualização nº 760 (22/10/2026). Todos os números de exemplo e dos testes são inventados. Prompts para um assistente de IA só de chat: [docs/prompt-gpt.md](docs/prompt-gpt.md).

## Passo a passo

1. Crie uma pasta, por exemplo `C:\exigibilidade`.
2. No GitHub, abra `exigibilidade.py`, clique em **Raw**, `Ctrl+A`, `Ctrl+C`. Cole no Bloco de Notas e salve como `exigibilidade.py` na pasta, em **UTF-8**.
3. Teste: no Spyder, abra o arquivo, aperte **F5** e digite `1`. No terminal: `python exigibilidade.py autoteste`. O `resumo_verificacao.txt` tem que terminar com `RESULTADO: TUDO OK`.
4. Exemplo: menu `2` ou `python exigibilidade.py exemplo`. Abra o `relatorio.html` em `saidas\rodadas\exemplo_...\`. No topo tem que aparecer `CHECK simulação = Python: OK`.
5. Dados reais: monte uma pasta de entradas (formatos abaixo, mesmo nome de arquivo do `exemplo\`) e rode:
   ```
   python exigibilidade.py rodar C:\exigibilidade\entradas --data-base 2026-11-20
   ```
   Cada rodada vai para `saidas\rodadas\<data_hora>\` e nunca é sobrescrita.
6. Comparar duas rodadas operação a operação (o que mudou de saldo, fonte/categoria, entrou ou saiu):
   ```
   python exigibilidade.py comparar saidas\rodadas\<rodada_A> saidas\rodadas\<rodada_B>
   ```

## Arquivos de entrada (pasta de entradas)

CSV, TXT (separador `;` `,` tab ou `|`) ou XLSX (1ª aba). Número com vírgula ou ponto decimal; data `AAAA-MM-DD` ou `DD/MM/AAAA`; mês `AAAA-MM`.

| Arquivo | Obrigatório | Colunas | Uma linha por |
|---|---|---|---|
| `operacoes` | sim | `id`, `fonte`, `categoria`, `saldo` (na data-base), `taxa_aa` (% a.a.), `fator` (ponderador, vazio = 1) | operação |
| `parcelas` | sim | `id`, `data`, `valor` | parcela futura; a última liquida |
| `liberacoes` | não | `id`, `data`, `valor`; para operação nova também `fonte`, `categoria`, `taxa_aa`, `fator` | liberação futura que vem da base |
| `realizado` | não | `mes`, `fonte`, `categoria`, `saldo_medio`, `fator` | mês antes da data-base, fonte e categoria |
| `vsr` | sim | `mes` (ou `data`), `vsr_avista`, `vsr_poupanca` | apuração do VSR (realizada ou projetada) |
| `lca` | sim | `data` (saldo diário) ou `mes` (saldo médio do mês projetado), `saldo` | dia ou mês |
| `dir` | não | `mes`, `modalidade`, `captacao`, `aplicacao` (saldo médio do mês) | mês e modalidade |
| `cenario` | não | `tipo` (`liquidacao` ou `reclassificacao`), `id`, `data`, `fonte`, `categoria` | ajuste numa operação |
| `feriados` | não | `data` | feriado (dia útil = segunda a sexta fora dos feriados) |
| `parametros.json` | não | salvo pelo botão do HTML | parâmetros da gerência |

**Fonte e categoria:**

| fonte | categorias | limite |
|---|---|---|
| `RO` | `pronaf`, `pronaf_ind_com` | ind./com. até 10% da sub Pronaf |
| | `pronamp`, `pronamp_invest`, `pronamp_peq_medio` | investimento e pequenos/médios até 10% da sub Pronamp, cada |
| | `demais`, `demais_abc`, `demais_pca` | ABC/RenovAgro até 3,6% e PCA até 2,4% da Geral |
| `POUP` | `credito_rural`, `cpr` | CPR até 5% |
| `LCA` | `credito_rural`, `faculdade` | faculdade até 55% |
| `LIVRE` | `livre` | não cumpre exigibilidade |

**DIR, modalidades:** `geral`, `pronaf`, `pronamp`, `poup`, `lca_cr`, `lca_livre`. A captação soma na exigibilidade e a aplicação conta no cumprimento.

## Regras implementadas

- **Saldo médio do mês:** soma dos saldos dos dias úteis ÷ dias úteis do mês. Conta zero antes da 1ª liberação e depois da liquidação (Doc 6, Anexo I, 4.6-b).
- **Fonte do mês:** a fonte e a categoria que a operação tinha no **dia 15** valem para o mês inteiro (Doc 6, Anexo I, 12.2-b).
- **Posição do mês:** média acumulada de julho até o mês, ponderada pelos dias úteis (Doc 6, Anexo I, 4.6-c/d).
- **Anexo II (RO):** própria = 31,5% × (média do VSR à vista de jul a jun do ano anterior − R$ 500 mi), cooperativas 13% em 2026/27 e 22% em 2027/28, isenta se ≤ R$ 10 mi. Pronaf 35%, Pronamp 50%, Geral = o resto, cada uma + captação DIR. Deficiência e excesso separados (códigos 5.1.x); o excesso de Pronaf e Pronamp cobre a Geral.
- **Anexo III (Poupança):** 70% da média do VSR; no mínimo 95% em crédito rural; CPR até 5%.
- **Anexo IV (LCA):** 60% × (média acumulada das LCA de junho até o mês anterior à posição − R$ 500 mi), isenta se ≤ R$ 10 mi; 45% em crédito rural e 55% de faculdade. O excesso de RO entra só na posição de junho (MCR 6-7-9-h).
- **Parâmetros da gerência** (grupos RO-Pronaf, RO-Pronamp, RO-Demais, Poupança e LCA, por mês, em R$ mi): a liberação entra no meio do mês e volta em parcela única depois do prazo; o % de reembolso deixa o que não foi pago no saldo até o fim do horizonte. Só valem a partir do mês da data-base.

**Ainda não trata:** custo financeiro da deficiência (MCR 6-5); renegociações da Res. 2.238 (limite de 60%) e outras normas transitórias (MCR 6-8); juros pela regra exata de cada contrato (usa juros compostos por dia corrido); o mês da data-base usa o saldo da data-base nos dias úteis anteriores a ela.

## O que é público e o que não é

Só vão para o GitHub este README, o `exigibilidade.py` e o `docs/prompt-gpt.md`. O `.gitignore` bloqueia todo o resto (entradas, saídas, exemplos, notas). Nada de dado real, nome de sistema, tabela ou pessoa aqui.
