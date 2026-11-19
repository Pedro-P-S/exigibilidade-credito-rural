# Padrões de código

## Norma

- Toda regra de cálculo cita o item do MCR ou o código do Documento 6 de onde saiu, no código e no README.
- O que não é norma (prazo, % de reembolso, juros da projeção, aproximações) é rotulado como **premissa gerencial** e nunca chamado de regra.
- Na dúvida entre o que a norma diz e o que se fazia antes, vale o texto do MCR vigente (versão no topo do programa).

## Sigilo (repositório público)

- Nada de dado real, valor, nome de cliente, pessoa, sistema, tabela ou servidor, nem nada de reuniões.
- Exemplos e testes só com números inventados.
- O `.gitignore` é lista de permissão: arquivo novo só entra no repositório se for liberado nele de propósito.
- Commits com o e-mail noreply do GitHub.

## Forma

- Um arquivo só (`exigibilidade.py`), só biblioteca padrão; `openpyxl` só para `.xlsx` (ADR 0002).
- Nomes em português, os mesmos termos do `GLOSSARY.md`.
- Mudou `pct_de`, `com_parametros`, `dir_final`, `acum`, `aplicar` ou `sisex`: mude a cópia em JavaScript no HTML também (ADR 0003).
- Atalho com limite conhecido leva o comentário `shortcut: <o limite>, <quando melhorar>`.

## Verificação

- Toda regra nova ganha uma checagem no `autoteste`, que imprime só OK/FALHOU; o autoteste tem de terminar em `RESULTADO: TUDO OK` antes de qualquer publicação.
- O autoteste garante que cada código de aplicação tem regra própria, sem empate.
- Prompts para o assistente de IA do trabalho nunca pedem valores; a resposta volta como OK/FALHOU, nomes de colunas ou contagens.
