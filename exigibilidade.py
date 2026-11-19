"""Exigibilidades do crédito rural (MCR cap. 6 e MCR - Documento 6, anexos II, III e IV).

Calcula a posição de cada mês (média acumulada de julho até o mês, como o Sisex) do ano-safra atual e do
seguinte, com o realizado, a projeção operação a operação e os parâmetros da gerência, e gera um HTML com a
visão dos diretores e o detalhe/simulação dos gestores.

Comandos:
  python exigibilidade.py autoteste                      testa as regras com dados inventados
  python exigibilidade.py exemplo                        gera entradas inventadas em exemplo/ e roda tudo
  python exigibilidade.py rodar PASTA_ENTRADAS --data-base AAAA-MM-DD [--tipo-if banco|cooperativa] [--rodada NOME]
  python exigibilidade.py comparar RODADA_A RODADA_B [--mes AAAA-MM]   (pastas ou saldo_operacoes.csv)
  python exigibilidade.py estrutura ARQUIVO_OU_PASTA [...] [--mascarar "NOME1,NOME2"]

Sem argumentos (F5 no Spyder) abre um menu. Saídas em saidas/ ao lado do script. Só biblioteca padrão
(Excel .xlsx precisa do openpyxl). Arquivos de entrada: README.md. Regras: MCR Atualização nº 760 (22/10/2026).
"""
import argparse
import calendar
import csv
import json
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

VERSAO = "0.4.0"
PASTA_SCRIPT = Path(__file__).resolve().parent
SAIDAS = PASTA_SCRIPT / "saidas"

DEDUCAO = 500_000_000  # MCR 6-2-2 e 6-7-4
ISENCAO = 10_000_000   # MCR 6-2-5 e 6-7-6

FONTES = ("RO", "POUP", "LCA", "LIVRE")  # LIVRE: não cumpre exigibilidade, fica fora
FONTE_DO_CODIGO = {"3.1": "RO", "3.2": "POUP", "3.3": "LCA"}  # anexos II, III e IV

# Códigos de limite do Documento 6: código -> (anexo, onde soma, base do limite, %, para onde vai o excesso, texto).
# Texto de cada um no MCR (descrição do código): os códigos que o compõem estão na coluna "cat" de REGRAS_TXT.
LIMITES = {
    "3.1.13.48-2": ("RO", "ap_pronaf", "sub_pronaf", 0.10, None, "Ind./com. Pronaf até 10% de 2.1.00.20-7"),
    "3.1.12.09-8": ("RO", "ap_pronaf", "sub_pronaf", 0.40, None, "Res. 5.330 Pronaf até 40% de 2.1.00.20-7"),
    "3.1.41.39-9": ("RO", "ap_pronamp", "sub_pronamp", 0.10, None, "Investimento Pronamp até 10% de 2.1.00.30-0"),
    "3.1.40.34-5": ("RO", "ap_pronamp", "sub_pronamp", 0.10, "ap_geral", "Peq./médios produtores até 10% de 2.1.00.30-0"),
    "3.1.40.38-3": ("RO", "ap_pronamp", "sub_pronamp", 0.40, None, "Res. 5.330 Pronamp até 40% de 2.1.00.30-0"),
    "3.1.30.87-4": ("RO", "ap_geral", "geral", 0.036, None, "ABC+/RenovAgro até 3,6% de 2.1.00.40-3"),
    "3.1.30.89-8": ("RO", "ap_geral", "geral", 0.024, None, "PCA até 2,4% de 2.1.00.40-3"),
    "3.1.32.01-6": ("RO", "ap_geral", "geral", 0.40, None, "Res. 5.330 até 40% de 2.1.00.40-3"),
    "3.1.30.53-7": ("RO", "ap_geral", "propria", 0.60, None, "Res. 2.238 até 60% de 2.1.10.00-8"),
    "3.2.10.54-3": ("POUP", "ap_cr", "exig_liquida", 0.075, None, "Res. 4.259 custeio até 7,5% de 2.2.50.00-9"),
    "3.2.10.72-5": ("POUP", "ap_cr", "exig_liquida", 0.015, None, "Res. 4.259 investimento até 1,5% de 2.2.50.00-9"),
    "3.2.20.29-6": ("POUP", "ap_cr", "exig_total", 0.10, None, "Res. 5.247 até 10% de 2.2.00.00-4"),
    "3.2.30.10-7": ("POUP", "ap_cpr", "propria", 0.05, None, "CPR até 5% de 2.2.10.00-1"),
    "3.3.20.17-2": ("LCA", "ap_cr", "sub", 0.20, None, "Res. 5.247 inc. I até 20% de 2.3.00.10-0"),
    "3.3.40.09-7": ("LCA", "ap_fac", "faculdade", 0.22, None, "Res. 5.247 inc. II/III até 22% de 2.3.00.20-3"),
}
# onde soma cada categoria sem limite
BALDE = {"RO": {"pronaf": "ap_pronaf", "pronamp": "ap_pronamp", "demais": "ap_geral"},
         "POUP": {"credito_rural": "ap_cr"}, "LCA": {"credito_rural": "ap_cr", "faculdade": "ap_fac"}}

# Regras operação -> código do Documento 6, tiradas do título de cada código de aplicação dos anexos II, III e IV
# (MCR Atualização nº 760). tipo: i = a instituição informa, c = o Sisex calcula (Doc 6, Anexo I, 12.1).
# cat: o limite que o código compõe (descrição do código de limite) ou a subexigibilidade da seção.
# de/ate: período de contratação; taxa_max/taxa_min: % a.a.; fator e ponderador: seção 4 do anexo
# ("Ponderação ... Valor correspondente a N% da média do saldo informado no código X" -> fator 1 + N%).
REGRAS_TXT = """\
3.1.13.08-0;i;pronaf;pronaf;custeio;2018-07-01;2019-06-30;;;;produtor;mcr_10_4_2_a;1.38;4.1.34.04-4;Operações de custeio de que trata o MCR 10-4-2-“a” - contratadas de 1/7/2018 a 30/6/2019.
3.1.13.09-7;i;pronaf;pronaf;custeio;2018-07-01;2019-06-30;;;;produtor;mcr_10_4_2_b;1.15;4.1.34.05-1;Operações de custeio de que trata o MCR 10-4-2-“b” - contratadas de 1/7/2018 a 30/6/2019.
3.1.13.15-2;i;pronaf;pronaf;;;2018-06-30;2.5;;;produtor;;1.38;4.1.34.07-5;Operações ao amparo do Pronaf com taxa até 2,5% a.a. (MCR 6-2-17-D) - contratadas até 30/6/2018.
3.1.13.16-9;i;pronaf;pronaf;;;2018-06-30;;2.5;;produtor;;1.15;4.1.34.08-2;Operações ao amparo do Pronaf com taxa superior a 2,5% a.a. (MCR 6-2-17-D) - contratadas até 30/6/2018.
3.1.13.17-6;i;pronaf;pronaf;custeio;2019-07-01;2020-06-30;;;;produtor;mcr_10_4_2_a;1.16;4.1.34.09-9;Operações de custeio de que trata o MCR 10-4-2-“a” - contratadas de 1/7/2019 a 30/6/2020.
3.1.13.18-3;i;pronaf;pronaf;custeio;2019-07-01;2020-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2019 a 30/6/2020.
3.1.13.21-7;i;pronaf;pronaf;investimento;2020-02-01;2020-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronaf - contratadas de 1/2/2020 a 30/6/2020.
3.1.13.24-8;i;pronaf;pronaf;custeio;2020-07-01;2021-06-30;;;;produtor;mcr_10_4_2_a;1.24;4.1.34.10-9;Operações de custeio de que trata o MCR 10-4-2-“a” - contratadas de 1/7/2020 a 30/6/2021.
3.1.13.25-5;i;pronaf;pronaf;custeio;2020-07-01;2021-06-30;;;;produtor;mcr_10_4_2_b;1.11;4.1.34.11-6;Operações de custeio de que trata o MCR 10-4-2-“b” - contratadas de 1/7/2020 a 30/6/2021.
3.1.13.27-9;i;pronaf;pronaf;investimento;2021-03-01;2021-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronaf - contratadas de 1/3/2021 a 30/6/2021.
3.1.13.29-3;i;pronaf;pronaf;custeio;2021-07-01;2022-06-30;3;;;produtor;;1.4;4.1.34.12-3;Operações de custeio contratadas de 1/7/2021 a 30/6/2022 com taxa de até 3% a.a. - Res. nº 4.928/2021.
3.1.13.30-3;i;pronaf;pronaf;custeio;2021-07-01;2022-06-30;4.5;;;produtor;;1.16;4.1.34.13-0;Operações de custeio contratadas de 1/7/2021 a 30/6/2022 com taxa de até 4,5% a.a. - Res. nº 4.928/2021.
3.1.13.31-0;i;pronaf;pronaf;custeio;2021-07-01;2022-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2021 a 30/6/2022.
3.1.13.33-4;i;pronaf;pronaf;custeio;2022-07-01;2023-06-30;5;;;produtor;;1.57;4.1.34.14-7;Operações de custeio contratadas de 1/7/2022 a 30/6/2023 com taxa de até 5% a.a. - Res. nº 5.028/2022.
3.1.13.34-1;i;pronaf;pronaf;custeio;2022-07-01;2023-06-30;6;;;produtor;;1.38;4.1.34.15-4;Operações de custeio contratadas de 1/7/2022 a 30/6/2023 com taxa de até 6% a.a. - Res. nº 5.028/2022.
3.1.13.35-8;i;pronaf;pronaf;custeio;2022-07-01;2023-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2022 a 30/6/2023.
3.1.13.36-5;i;3.1.13.48-2;pronaf;industrializacao;2022-07-01;2023-06-30;;;;produtor;;;;Operações de industrialização - contratadas de 1/7/2022 a 30/6/2023.
3.1.13.37-2;i;pronaf;pronaf;custeio;2023-07-01;2024-06-30;4;;;produtor;;1.26;4.1.34.16-1;Operações de custeio - contratadas de 1/7/2023 a 30/6/2024 com taxa de até 4% a.a. - Res. n° 5.087/2023.
3.1.13.38-9;i;pronaf;pronaf;custeio;2023-07-01;2024-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2023 a 30/6/2024.
3.1.13.39-6;i;3.1.13.48-2;pronaf;industrializacao;2023-07-01;2024-06-30;;;;produtor;;;;Operações de industrialização - contratadas de 1/7/2023 a 30/6/2024.
3.1.13.40-6;i;pronaf;pronaf;custeio;2024-07-01;2025-06-30;3;;;produtor;;1.37;4.1.34.17-8;Operações de custeio - contratadas de 1/7/2024 a 30/6/2025 com taxa de até 3% a.a. - Res. n° 5.170/2024.
3.1.13.41-3;i;pronaf;pronaf;custeio;2024-07-01;2025-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2024 a 30/6/2025.
3.1.13.42-0;i;3.1.13.48-2;pronaf;industrializacao;2024-07-01;2025-06-30;;;;produtor;;;;Operações de industrialização - contratadas de 1/7/2024 a 30/6/2025.
3.1.13.43-7;c;pronaf;pronaf;custeio;2025-07-01;2026-06-30;3;;;produtor;;1.37;4.1.34.18-5;Operações de custeio contratadas de 1/7/2025 a 30/6/2026 com taxa de até 3% a.a. - Res. n° 5.170/2024.
3.1.13.44-4;c;pronaf;pronaf;custeio;2025-07-01;2026-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2025 a 30/6/2026.
3.1.13.46-8;c;3.1.13.48-2;pronaf;industrializacao;2026-07-01;2027-06-30;;;;produtor;;;;Operações de industrialização - MCR 10-11 - contratadas de 1/7/2026 a 30/6/2027.
3.1.13.47-5;c;3.1.13.48-2;pronaf;comercializacao;2026-07-01;2027-06-30;;;;produtor;;;;Operações de comercialização - MCR 10-11 - contratadas de 1/7/2026 a 30/6/2027.
3.1.13.49-9;c;pronaf;pronaf;custeio;2026-07-01;2027-06-30;3;;;produtor;;1.37;4.1.34.19-2;Operações de custeio contratadas de 1/7/2026 a 30/6/2027 com taxa de até 3% a.a. - Res. nº 5.170/2024.
3.1.13.50-9;c;pronaf;pronaf;custeio;2026-07-01;2027-06-30;;;;produtor;;;;Operações de custeio - Demais operações - contratadas de 1/7/2026 a 30/6/2027.
3.1.10.51-9;i;pronaf;pronaf;;;;;;;produtor;proagro;;;Proagro - Ressarcimentos pendentes vinculados ao Pronaf (MCR 6-2).
3.1.10.52-6;i;pronaf;pronaf;;;;;;;produtor;excluida_subvencao;;;Financiamentos rurais excluídos da base da subvenção do Tesouro Nacional (TN), concedidos a beneficiários do Pronaf (MCR 6-2).
3.1.10.54-0;i;pronaf;pronaf;;;;;;;produtor;outras_fontes;;;Operações contratadas originalmente com recursos de outras fontes, com beneficiários do Pronaf (MCR 6-2).
3.1.11.48-4;i;pronaf;pronaf;;;;;;;produtor;renegociacao;;;Renegociação de dívidas de operações de crédito rural contratadas no âmbito do Pronaf - Outras.
3.1.11.50-1;i;pronaf;pronaf;custeio;;;;;;produtor;res_4801;;;Operações de custeio vinculadas ao Pronaf - Res. nº 4.801/2020.
3.1.11.97-2;i;pronaf;pronaf;custeio;;;;;;produtor;renegociacao_4802;;;Renegociação de operações de custeio vinculadas ao Pronaf - Res. nº 4.802/2020.
3.1.12.04-3;i;pronaf;pronaf;custeio;;;;;;produtor;res_4802;;;Operações de custeio vinculadas ao Pronaf - Res. nº 4.802/2020.
3.1.13.45-1;i;pronaf;pronaf;;;;;;;produtor;desclassificada;;;Operações desclassificadas parcialmente, contratadas no âmbito do Pronaf.
3.1.12.08-1;c;3.1.12.09-8;pronaf;;;;;;;produtor;res_5330;;;Operações de crédito rural, contratadas no âmbito do Pronaf - Resolução CMN nº 5.330/2026.
3.1.30.14-2;i;demais;demais;custeio;;;;;;produtor;avicultura_suinocultura_piscicultura;;;Custeio - Avicultura, suinocultura e piscicultura.
3.1.30.35-5;i;demais;demais;comercializacao;;;;;;produtor;;;;Operações de comercialização, exceto as de desconto de DR e de NPR representativas da comercialização de leite e FGPP.
3.1.30.38-6;i;demais;demais;investimento;2013-07-01;2014-06-30;;;;produtor;invest_acima_300mil;1.25;4.1.32.21-1;Operações de investimento com valor superior a R$300.000,00, nos termos do MCR 6-2-11-“b” - Exigibilidade Própria - Contratadas de 1/7/2013 a 30/6/2014 (MCR 6-2-2 e 6-2-11-“b”).
3.1.30.42-7;i;demais;demais;investimento;;2017-06-30;;;;produtor;pecuario;;;Operações de investimento pecuário - contratadas até 30/6/2017.
3.1.30.43-4;i;demais;demais;comercializacao;;;;;;produtor;desconto_leite;;;Operações de desconto representativas da comercialização de leite.
3.1.30.45-8;i;demais;demais;custeio;;;;;;produtor;;;;Operações de custeio.
3.1.30.46-5;i;demais;demais;investimento;;2017-06-30;;;;produtor;;;;Operações de investimento - contratadas até 30/6/2017.
3.1.30.67-8;i;demais;demais;industrializacao;;;;;;produtor;;;;Operações de industrialização.
3.1.30.49-6;i;demais;demais;investimento;2014-07-01;2015-06-30;;;;produtor;mcr_6_2_17_b;1.25;4.1.33.84-9;Operações de investimento contratadas nos termos do MCR 6-2-17-“b” - Exigibilidade Própria - Contratadas de 1/7/2014 a 30/6/2015 (MCR 6-2-3 e 6-2-17-“b”).
3.1.20.15-2;i;demais;demais;investimento;;2011-06-30;;;;produtor;invest_ate_200mil;1.1;4.1.20.10-6;Aplicações em investimento com valor de até R$200.000,00 - Demais operações (MCR 3-3-14 e 6-2-7- A) - contratadas até 30/6/2011.
3.1.21.00-3;i;demais;demais;investimento;2011-07-01;2012-06-30;;;;produtor;invest_ate_300mil;1.1;4.1.40.47-8;Operações de Investimento com valor de até R$300.000,00 - MCR 3-3 - Contratadas de 1/7/2011 a 30/6/2012 (MCR 3-3-14 e 6-2-7-A), contratadas com prazo de reembolso superior a 2 (dois) anos.
3.1.30.62-3;i;demais;demais;custeio;;;;;;produtor;cultivo_protegido;1.2;4.1.33.92-8;Operações de custeio - Cultivo Protegido (MCR 6-2-17-“a”).
3.1.30.71-9;i;demais;demais;investimento;2018-07-01;2019-06-30;;;;produtor;bovinos;;;Operações de investimento destinadas à aquisição de bovinos ou bubalinos para reprodução ou cria - contratadas de 1/7/2018 a 30/6/2019 - Res. 4.669/2018.
3.1.30.86-7;i;3.1.30.87-4;demais;investimento;2021-07-01;2022-06-30;;;;produtor;abc;;;Operações de investimento - Programa ABC - contratadas de 1/7/2021 a 30/6/2022 - Res. 4.917/2021.
3.1.30.88-1;i;3.1.30.89-8;demais;investimento;2021-07-01;2022-06-30;;;;produtor;pca;;;Operações de investimento - PCA - contratadas de 1/7/2021 a 30/6/2022 - Res. 4.917/2021.
3.1.30.92-2;i;3.1.30.87-4;demais;investimento;2022-07-01;2023-06-30;;;;produtor;abc;;;Operações de investimento - Programa ABC+ - contratadas de 1/7/2022 a 30/6/2023 - Res. 5.029/2022.
3.1.30.94-6;i;3.1.30.87-4;demais;investimento;2023-07-01;2024-06-30;;;;produtor;abc;;;Operações de investimento - Programa RenovAgro - contratadas de 1/7/2023 a 30/6/2024 - Res. 5.087/2023.
3.1.30.97-7;c;demais;demais;custeio;2025-07-01;2026-06-30;;;;produtor;;;;Operações de Custeio - contratadas de 1/7/2025 a 30/6/2026.
3.1.30.98-4;c;demais;demais;industrializacao;2025-07-01;2026-06-30;;;;produtor;;;;Operações de Industrialização - contratadas de 1/7/2025 a 30/6/2026.
3.1.31.07-9;c;demais;demais;custeio;2026-07-01;2027-06-30;;;;produtor;;;;Operações de Custeio - contratadas de 1/7/2026 a 30/6/2027.
3.1.31.08-6;c;demais;demais;industrializacao;2026-07-01;2027-06-30;;;;produtor;;;;Operações de Industrialização - contratadas de 1/7/2026 a 30/6/2027.
3.1.21.31-9;i;demais;demais;custeio;;2026-06-30;;;;cooperativa;;;;Créditos de custeio para cooperativa de produção agropecuária - contratadas até 30/6/2026.
3.1.21.75-9;i;demais;demais;industrializacao;;2026-06-30;;;;cooperativa;;;;Créditos de industrialização para cooperativa de produção agropecuária - contratadas até 30/6/2026.
3.1.21.33-3;i;demais;demais;comercializacao;;2026-06-30;;;;cooperativa;;;;Créditos de comercialização para cooperativa de produção agropecuária, exceto FGPP - contratadas até 30/06/2026.
3.1.21.34-0;i;demais;demais;custeio;;2026-06-30;;;;cooperativa_cooperados;;;;Créditos de custeio para atendimento a cooperados, destinado à aquisição de insumos para fornecimento aos cooperados - contratadas até 30/6/2026.
3.1.21.35-7;i;demais;demais;investimento;;2017-06-30;;;;cooperativa_cooperados;;;;Créditos de investimento destinados à aquisição de bens para fornecimento a cooperados, para utilização na atividade de produção agropecuária, contratados até 30/6/2017.
3.1.21.76-6;i;demais;demais;industrializacao;;2026-06-30;;;;cooperativa_cooperados;;;;Créditos de industrialização a cooperativas para atendimento a cooperados - contratadas até 30/6/2026.
3.1.21.77-3;i;demais;demais;comercializacao;;2026-06-30;;;;cooperativa_cooperados;;;;Créditos de comercialização a cooperativas para atendimento a cooperados, exceto FGPP - contratadas até 30/6/2026.
3.1.20.24-8;i;demais;demais;;;;;;;cooperativa;cotas_partes;;;Crédito a cooperativa de produção agropecuária como financiamento da integralização de cotas-partes do capital social.
3.1.21.79-7;c;demais;demais;industrializacao;2026-07-01;2027-06-30;;;;cooperativa;;;;Créditos de industrialização para cooperativa de produção agropecuária na condição de produtor rural - contratadas de 1/7/2026 a 30/6/2027. (*)
3.1.21.80-7;c;demais;demais;comercializacao;2026-07-01;2027-06-30;;;;cooperativa;;;;Créditos de comercialização para cooperativa de produção agropecuária, exceto FGPP - contratadas de 1/7/2026 a 30/6/2027.
3.1.21.81-4;c;demais;demais;comercializacao;2026-07-01;2027-06-30;;;;cooperativa_servicos;;;;Créditos de comercialização a cooperativas de produção agropecuária na condição de sociedade prestadora de serviços de natureza agropecuária aos seus cooperados, exceto FGPP - contratadas de 1/7/2026 a 30/6/2027. (*)
3.1.21.82-1;c;demais;demais;custeio;2026-07-01;2027-06-30;;;;cooperativa_cooperados;;;;Créditos de custeio para atendimento a cooperados, destinado à aquisição de insumos para fornecimento aos cooperados - contratadas de 1/7/2026 até 30/6/2027.
3.1.21.83-8;c;demais;demais;custeio;2026-07-01;2027-06-30;;;;cooperativa;;;;Créditos de custeio para cooperativa de produção agropecuária - contratadas de 1/7/2026 a 30/6/2027.
3.1.21.84-5;c;demais;demais;industrializacao;2026-07-01;2027-06-30;;;;cooperativa_servicos;;;;Créditos de industrialização para cooperativas de produção agropecuária na condição de sociedade prestadora de serviços de natureza agropecuária aos seus cooperados - contratadas de 1/7/2026 a 30/6/2027. (*)
3.1.20.21-7;i;demais;demais;investimento;;2011-06-30;;;;cooperativa_cooperados;;1.1;4.1.40.01-4;Créditos a cooperativas para repasse a cooperados (MCR 5-5-19 e 6-2-7) - operações de investimento - contratadas até 30/6/2011.
3.1.30.65-4;i;3.1.30.53-7;demais;;;;;;;produtor;renegociacao_2238;;;Renegociação de dívidas rurais - Resolução nº 2.238/1996.
3.1.30.54-4;i;demais;demais;;;;;;;produtor;renegociacao_2238_tn;;;Renegociação de dívidas rurais - Valores cedidos ao Tesouro Nacional - Resolução nº 2.238/1996.
3.1.30.57-5;i;demais;demais;;;;;;;produtor;renegociacao;;;Renegociação de dívidas de operações de crédito rural - Outras.
3.1.30.58-2;i;demais;demais;;;;;;;produtor;proagro;;;Proagro - Ressarcimentos pendentes.
3.1.30.59-9;i;demais;demais;;;;;;;produtor;excluida_subvencao;;;Financiamentos rurais excluídos da base da subvenção do Tesouro Nacional.
3.1.30.61-6;i;demais;demais;;;;;;;produtor;outras_fontes;;;Operações contratadas originalmente com recursos de outras fontes.
3.1.30.73-3;i;demais;demais;investimento;;;;;;produtor;res_4801;;;Operações de investimento prorrogadas - Res. nº 4.801/2020.
3.1.30.75-7;i;demais;demais;custeio;;;;;;produtor;renegociacao_4802;;;Renegociação de operações de custeio - Res. nº 4.802/2020.
3.1.30.76-4;i;demais;demais;investimento;;;;;;produtor;renegociacao_4802;;;Renegociação de operações de investimento - Res. nº 4.802/2020.
3.1.30.96-0;i;demais;demais;custeio;;;;;;produtor;exig_adicional;;;Operações de custeio contratadas no âmbito das exigibilidades adicionais.
3.1.30.99-1;i;demais;demais;;;;;;;produtor;desclassificada;;;Operações desclassificadas parcialmente.
3.1.32.00-9;c;3.1.32.01-6;demais;;;;;;;produtor;res_5330;;;Operações de crédito rural - Resolução CMN nº 5.330/2026.
3.1.40.10-1;i;pronamp;pronamp;custeio;;2010-06-30;;;;produtor;;;;Operações de custeio no Pronamp - Contratadas até 30/6/2010 (MCR 8-1 e MCR 6-2-5).
3.1.40.11-8;i;3.1.41.39-9;pronamp;investimento;;2010-06-30;;;;produtor;;;;Operações de investimento no Pronamp - Contratadas até 30/6/2010 (MCR 8-1 e MCR 6-2-5).
3.1.40.12-5;i;pronamp;pronamp;custeio;2010-07-01;2011-06-30;;;;produtor;;1.13;4.1.11.00-5;Operações de custeio no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2010 a 30/6/2011 (MCR 8-1 e MCR 6-2-5).
3.1.40.13-2;i;3.1.41.39-9;pronamp;investimento;2010-07-01;2011-06-30;;;;produtor;;1.13;4.1.11.00-5;Operações de investimento no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2010 a 30/6/2011 (MCR 8-1 e MCR 6-2-5).
3.1.40.16-3;i;pronamp;pronamp;custeio;2011-07-01;2012-06-30;;;;produtor;;1.1;4.1.11.01-2;Operações de custeio no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2011 a 30/6/2012 (MCR 8-1 e MCR 6-2-5).
3.1.40.17-0;i;3.1.41.39-9;pronamp;investimento;2011-07-01;2012-06-30;;;;produtor;;1.1;4.1.11.01-2;Operações de investimento no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2011 a 30/6/2012 (MCR 8-1 e MCR 6-2-5).
3.1.41.01-4;i;pronamp;pronamp;custeio;2012-07-01;2013-06-30;;;;produtor;;1.11;4.1.11.02-9;Operações de custeio no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2012 a 30/6/2013 (MCR 8-1 e MCR 6-2-5).
3.1.41.02-1;i;3.1.41.39-9;pronamp;investimento;2012-07-01;2013-06-30;;;;produtor;;;;Operações de investimento no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2012 a 30/6/2013 (MCR 8-1 e MCR 6-2-5).
3.1.41.06-9;i;3.1.41.39-9;pronamp;investimento;2013-07-01;2014-06-30;;;;produtor;;;;Operações de investimento no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2013 a 30/6/2014 (MCR 8-1 e MCR 6-2-5).
3.1.41.08-3;i;3.1.41.39-9;pronamp;investimento;2013-07-01;2014-06-30;;;;produtor;semiarido;;;Operações de investimento no Pronamp para o Plano Safra Semiárido 2013/2014 - Resolução nº 4.254/2013 - Subexigibilidade Própria - Contratadas de 1/7/2013 a 30/6/2014 (MCR 8-2-2 e MCR 6-2- 5).
3.1.41.17-9;i;pronamp;pronamp;custeio;2014-07-01;2015-06-30;;;;produtor;;1.11;4.1.11.06-7;Operações de custeio no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2014 a 30/6/2015 (MCR 8-1 e MCR 6-2-9).
3.1.41.18-6;i;3.1.41.39-9;pronamp;investimento;2014-07-01;2015-06-30;;;;produtor;;;;Operações de investimento no Pronamp - Subexigibilidade Própria - Contratadas de 1/7/2014 a 30/6/2015 (MCR 8-1 e MCR 6-2-9).
3.1.41.27-2;i;pronamp;pronamp;custeio;2015-07-01;2016-06-30;;;;produtor;;;;Operações de custeio no Pronamp - Contratadas de 1/7/2015 a 30/6/2016 (MCR 6-2-9 e 8-1).
3.1.41.29-6;i;pronamp;pronamp;custeio;2016-07-01;2017-06-30;;;;produtor;;;;Operações de custeio no Pronamp - Contratadas de 1/7/2016 a 30/6/2017 (MCR 6-2-9 e 8-1).
3.1.41.31-3;i;pronamp;pronamp;custeio;2017-07-01;2018-06-30;;;;produtor;;;;Operações de Custeio - Contratadas de 1/7/2017 a 30/6/2018.
3.1.41.33-7;i;pronamp;pronamp;custeio;2018-07-01;2019-06-30;;;;produtor;;;;Operações de Custeio - Contratadas de 1/7/2018 a 30/6/2019.
3.1.41.35-1;i;3.1.41.39-9;pronamp;investimento;2018-07-01;2019-06-30;;;;produtor;bovinos;;;Operações de investimento vinculadas ao Pronamp destinadas à aquisição de bovinos ou bubalinos para reprodução ou cria - contratadas de 1/7/2018 a 30/6/2019 - Res. 4.669/2018.
3.1.41.37-5;i;pronamp;pronamp;custeio;2019-07-01;2020-06-30;;;;produtor;;;;Operações de Custeio - Contratadas de 1/7/2019 a 30/6/2020.
3.1.41.38-2;i;3.1.41.39-9;pronamp;investimento;2019-07-01;2020-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2019 a 30/6/2020 - Res. 4.730/2019.
3.1.41.40-9;i;pronamp;pronamp;custeio;2020-07-01;2021-06-30;;;;produtor;;;;Operações de Custeio - Contratadas de 1/7/2020 a 30/6/2021.
3.1.41.41-6;i;3.1.41.39-9;pronamp;investimento;2020-07-01;2021-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2020 a 30/6/2021.
3.1.41.42-3;i;pronamp;pronamp;custeio;2021-07-01;2022-06-30;;;;produtor;;;;Operações de Custeio - Contratadas de 1/7/2021 a 30/6/2022.
3.1.41.43-0;i;3.1.41.39-9;pronamp;investimento;2021-07-01;2022-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2021 a 30/6/2022.
3.1.41.44-7;i;pronamp;pronamp;custeio;2022-07-01;2023-06-30;;;;produtor;;;;Operações de custeio - Contratadas de 1/7/2022 a 30/6/2023.
3.1.41.45-4;i;3.1.41.39-9;pronamp;investimento;2022-07-01;2023-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2022 a 30/6/2023.
3.1.41.46-1;i;pronamp;pronamp;custeio;2023-07-01;2024-06-30;;;;produtor;;;;Operações de custeio - contratadas de 1/7/2023 a 30/6/2024.
3.1.41.47-8;i;3.1.41.39-9;pronamp;investimento;2023-07-01;2024-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2023 a 30/6/2024.
3.1.41.48-5;i;pronamp;pronamp;custeio;2024-07-01;2025-06-30;;;;produtor;;;;Operações de custeio - contratadas de 1/7/2024 a 30/6/2025.
3.1.41.49-2;i;3.1.41.39-9;pronamp;investimento;2024-07-01;2025-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2024 a 30/6/2025.
3.1.41.50-2;c;pronamp;pronamp;custeio;2025-07-01;2026-06-30;;;;produtor;;;;Operações de custeio vinculadas ao Pronamp - contratadas de 1/7/2025 a 30/6/2026.
3.1.41.51-9;c;3.1.41.39-9;pronamp;investimento;2025-07-01;2026-06-30;;;;produtor;;;;Operações de investimento vinculadas ao Pronamp - contratadas de 1/7/2025 a 30/6/2026.
3.1.41.52-6;c;pronamp;pronamp;custeio;2026-07-01;2027-06-30;;;;produtor;;;;Operações de custeio - contratadas de 1/7/2026 a 30/6/2027.
3.1.41.53-3;c;3.1.41.39-9;pronamp;investimento;2026-07-01;2027-06-30;;;;produtor;;;;Operações de investimento - contratadas de 1/7/2026 a 30/6/2027.
3.1.40.21-1;i;pronamp;pronamp;;;;;;;produtor;proagro;;;Proagro - Ressarcimentos pendentes.
3.1.40.22-8;i;pronamp;pronamp;;;;;;;produtor;excluida_subvencao;;;Financiamentos rurais excluídos da base da subvenção do Tesouro Nacional.
3.1.40.23-5;i;pronamp;pronamp;;;;;;;produtor;outras_fontes;;;Operações contratadas originalmente com recursos de outras fontes.
3.1.40.24-2;i;pronamp;pronamp;;;;;;;produtor;renegociacao;;;Renegociação de dívidas de operações de crédito rural no âmbito do Pronamp - Outras.
3.1.40.27-3;i;pronamp;pronamp;custeio;;;;;;produtor;res_4801;;;Operações de custeio vinculadas ao Pronamp - Res. nº 4.801/2020.
3.1.40.28-0;i;pronamp;pronamp;custeio;;;;;;produtor;renegociacao_4802;;;Renegociação de operações de custeio vinculadas ao Pronamp - Res. nº 4.802/2020.
3.1.40.29-7;i;3.1.41.39-9;pronamp;investimento;;;;;;produtor;renegociacao_4802;;;Renegociação de operações de investimento vinculadas ao Pronamp - Res. nº 4.802/2020.
3.1.40.33-8;i;pronamp;pronamp;custeio;;;;;;produtor;res_4802;;;Operações de custeio vinculadas ao Pronamp - Res. nº 4.802/2020.
3.1.40.35-2;i;3.1.40.34-5;pronamp;custeio;;;;;;produtor;peq_medio;;;Operações de custeio com pequenos e médios produtores rurais - MCR 6-2.
3.1.40.36-9;i;pronamp;pronamp;;;;;;;produtor;desclassificada;;;Operações desclassificadas parcialmente, contratadas no âmbito do Pronamp.
3.1.40.37-6;c;3.1.40.38-3;pronamp;;;;;;;produtor;res_5330;;;Operações de crédito rural, contratadas no âmbito do Pronamp - Resolução CMN nº 5.330/2026.
3.2.10.09-3;i;credito_rural;;industrializacao;;;;;nao_controlados;;;;;Operações de Industrialização - recursos não controlados.
3.2.10.10-3;i;credito_rural;;custeio;;;;;nao_controlados;;;;;Operações de custeio - recursos não controlados.
3.2.10.11-0;i;credito_rural;;investimento;;;;;nao_controlados;;;;;Operações de investimento - recursos não controlados.
3.2.10.12-7;i;credito_rural;;comercializacao;;;;;nao_controlados;;;;;Operações de comercialização - recursos não controlados.
3.2.10.13-4;i;credito_rural;pronaf;;;2009-06-30;;;;;pronaf_condicoes_6_2;;;Operações formalizadas nas condições do MCR 6-2 no âmbito do Pronaf - contratadas até 30/6/2009 (MCR 6-4).
3.2.10.20-6;i;credito_rural;pronaf;custeio;;2011-06-30;;;controlados;;;;;Operações de custeio formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas até 30/6/2011.
3.2.10.21-3;i;credito_rural;pronaf;investimento;;2011-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4- 5) - contratadas até 30/6/2011.
3.2.10.24-4;i;credito_rural;pronamp;investimento;;2011-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas até 30/6/2011.
3.2.10.29-9;i;credito_rural;;investimento;;;;;;;bndes;;;Operações de investimento nas condições dos programas do BNDES.
3.2.10.30-9;i;credito_rural;;investimento;;;;;;;abc;;;Operações de investimento - Programa para Redução da Emissão de Gases de Efeito Estufa na Agricultura - Programa ABC (Resolução nº 3.896/2010) - Aplica-se exclusivamente ao Banco do Brasil S.A.
3.2.10.34-7;i;credito_rural;pronaf;custeio;2011-07-01;2012-06-30;;;controlados;;;;;Operações de custeio formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2011 até 30/6/2012.
3.2.10.35-4;i;credito_rural;pronaf;investimento;2011-07-01;2012-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4- 5) - contratadas de 1/7/2011 até 30/6/2012.
3.2.10.38-5;i;credito_rural;pronamp;investimento;2011-07-01;2012-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2011 até 30/6/2012.
3.2.10.39-2;i;credito_rural;demais;custeio;2011-07-01;2012-06-30;;;controlados;;;;;Operações de custeio formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2011 até 30/6/2012.
3.2.10.42-6;i;credito_rural;pronaf;custeio;2012-07-01;2013-06-30;;;controlados;;;;;Operações de custeio formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2012 até 30/6/2013.
3.2.10.45-7;i;credito_rural;pronamp;custeio;2012-07-01;2013-06-30;;;controlados;;;;;Operações de custeio formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2012 até 30/6/2013.
3.2.10.46-4;i;credito_rural;pronamp;investimento;2012-07-01;2013-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2012 até 30/6/2013.
3.2.10.47-1;i;credito_rural;demais;custeio;2012-07-01;2013-06-30;;;controlados;;;;;Operações de custeio formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2012 até 30/6/2013.
3.2.10.49-5;i;credito_rural;demais;comercializacao;2012-07-01;2013-06-30;;;controlados;;;;;Operações de comercialização formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2012 até 30/6/2013.
3.2.10.51-2;i;credito_rural;;;2013-07-01;2014-06-30;;;;;armazenagem;;;Financiamentos para Armazenagem (MCR 6-4-5-“a”) - contratadas de 1/7/2013 até 30/6/2014.
3.2.11.01-6;i;3.2.10.54-3;pronaf;custeio;2013-07-01;2014-06-30;;;controlados;;res_4259;;;Operações de custeio formalizadas ao amparo do Pronaf - Resolução nº 4.259/2013 - recursos controlados (MCR 6-1-2, 6-4-5 e 6-4-8) - contratadas de 1/7/2013 até 30/6/2014.
3.2.10.55-0;i;credito_rural;pronaf;investimento;2013-07-01;2014-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4- 5) - contratadas de 1/7/2013 até 30/6/2014.
3.2.10.56-7;i;credito_rural;pronaf;investimento;2013-07-01;2014-06-30;;;controlados;;semiarido;;;Operações de investimento formalizadas ao amparo do Pronaf para o Plano Safra Semiárido 2013/2014 - recursos controlados - Resolução nº 4.248/2013 (MCR 6-1-2, 6-4-5 e 10-18-10) - contratadas de 1/7/2013 até 30/6/2014.
3.2.10.62-2;i;credito_rural;pronamp;investimento;2013-07-01;2014-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2013 até 30/6/2014.
3.2.10.68-4;i;credito_rural;demais;investimento;2013-07-01;2014-06-30;;;controlados;;semiarido;;;Operações de investimento formalizadas com demais produtores para o Plano Safra Semiárido 2013/2014 - recursos controlados (MCR 6-1-2, 6-4-5 e 3-6-13) - Resolução nº 4.254/2013 - contratadas de 1/7/2013 até 30/6/2014.
3.2.11.02-3;i;3.2.10.72-5;;investimento;2013-07-01;2013-12-31;;;controlados;;res_4259;;;Operações de investimento - Resolução nº 4.259/2013 - recursos controlados (MCR 6-1-2, 6-4-5 e 6-4-8) - contratadas de 1/7/2013 até 31/12/2013.
3.2.10.73-2;i;credito_rural;;;2014-07-01;2015-06-30;;;;;armazenagem;;;Financiamentos para Armazenagem (MCR 6-4-5-“a”) - contratadas de 1/7/2014 até 30/6/2015.
3.2.10.75-6;i;credito_rural;pronaf;investimento;2014-07-01;2015-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4- 5) - contratadas de 1/7/2014 até 30/6/2015.
3.2.10.78-7;i;credito_rural;pronamp;investimento;2014-07-01;2015-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2014 até 30/6/2015.
3.2.10.79-4;i;credito_rural;demais;custeio;2014-07-01;2015-06-30;;;controlados;;;;;Operações de custeio formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2014 até 30/6/2015.
3.2.10.80-4;i;credito_rural;demais;investimento;2014-07-01;2015-06-30;;;controlados;;;;;Operações de investimento formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2014 até 30/6/2015.
3.2.10.83-5;i;credito_rural;;;2015-07-01;2016-06-30;;;;;armazenagem;;;Financiamentos para Armazenagem (MCR 6-4-5-“a”) - contratadas de 1/7/2015 até 30/6/2016.
3.2.10.84-2;i;credito_rural;pronaf;custeio;2015-07-01;2016-06-30;;;controlados;;;;;Operações de custeio formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2015 até 30/6/2016.
3.2.10.85-9;i;credito_rural;pronaf;investimento;2015-07-01;2016-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4- 5) - contratadas de 1/7/2015 até 30/6/2016.
3.2.10.88-0;i;credito_rural;pronamp;investimento;2015-07-01;2016-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2015 até 30/6/2016.
3.2.10.89-7;i;credito_rural;demais;custeio;2015-07-01;2016-06-30;;;controlados;;;;;Operações de custeio formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2015 até 30/6/2016.
3.2.10.90-7;i;credito_rural;demais;investimento;2015-07-01;2016-06-30;;;controlados;;;;;Operações de investimento formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2015 até 30/6/2016.
3.2.11.03-0;i;credito_rural;;;2016-07-01;2017-06-30;;;;;armazenagem;;;Financiamento para Armazenagem – MCR 6-4-5-“a” – de 1/7/2016 a 30/6/2017.
3.2.11.04-7;i;credito_rural;pronaf;custeio;2016-07-01;2017-06-30;;;controlados;;;;;Operações de custeio formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2016 até 30/6/2017.
3.2.11.06-1;i;credito_rural;pronaf;investimento;2016-07-01;2017-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronaf - recursos controlados (MCR 6-1-2 e 6-4- 5) - contratadas de 1/7/2016 até 30/6/2017.
3.2.11.10-2;i;credito_rural;pronamp;investimento;2016-07-01;2017-06-30;;;controlados;;;;;Operações de investimento formalizadas ao amparo do Pronamp - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2016 até 30/6/2017.
3.2.11.11-9;i;credito_rural;demais;custeio;2016-07-01;2017-06-30;;;controlados;;;;;Operações de custeio formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6-4-5) - contratadas de 1/7/2016 até 30/6/2017.
3.2.11.13-3;i;credito_rural;demais;investimento;2016-07-01;2017-06-30;;;controlados;;;;;Operações de investimento formalizadas com demais produtores - recursos controlados (MCR 6-1-2 e 6- 4-5) - contratadas de 1/7/2016 até 30/6/2017.
3.2.11.15-7;i;credito_rural;;;2017-07-01;2018-06-30;;;;;armazenagem;;;Financiamentos para Armazenagem - MCR 6-4-5-"a" - contratadas de 1/7/2017 até 30/6/2018.
3.2.11.16-4;i;credito_rural;pronaf;custeio;2017-07-01;2018-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - contratadas de 1/7/2017 até 30/6/2018.
3.2.11.18-8;i;credito_rural;pronaf;investimento;2017-07-01;2018-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - contratadas de 1/7/2017 até 30/6/2018.
3.2.11.20-5;i;credito_rural;pronamp;custeio;2017-07-01;2018-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - contratadas de 1/7/2017 até 30/6/2018.
3.2.11.21-2;i;credito_rural;pronamp;investimento;2017-07-01;2018-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - contratadas de 1/7/2017 até 30/6/2018.
3.2.11.22-9;i;credito_rural;demais;custeio;2017-07-01;2018-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - contratadas de 1/7/2017 até 30/6/2018.
3.2.11.26-7;i;credito_rural;;;2018-07-01;2019-06-30;;;;;armazenagem;;;Financiamentos para Armazenagem - MCR 6-4-5-"a" - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.27-4;i;credito_rural;pronaf;custeio;2018-07-01;2019-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.29-8;i;credito_rural;pronaf;investimento;2018-07-01;2019-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.30-8;i;credito_rural;pronamp;custeio;2018-07-01;2019-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.31-5;i;credito_rural;pronamp;investimento;2018-07-01;2019-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.32-2;i;credito_rural;demais;custeio;2018-07-01;2019-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.34-6;i;credito_rural;demais;investimento;2018-07-01;2019-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - contratadas de 1/7/2018 até 30/6/2019.
3.2.11.37-7;i;credito_rural;pronaf;custeio;2019-07-01;2020-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - contratadas de 1/7/2019 até 30/6/2020.
3.2.11.39-1;i;credito_rural;pronaf;investimento;2019-07-01;2020-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - contratadas de 1/7/2019 até 30/6/2020.
3.2.11.40-1;i;credito_rural;pronamp;custeio;2019-07-01;2020-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - contratadas de 1/7/2019 até 30/6/2020.
3.2.11.41-8;i;credito_rural;pronamp;investimento;2019-07-01;2020-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - contratadas de 1/7/2019 até 30/6/2020.
3.2.11.42-5;i;credito_rural;demais;custeio;2019-07-01;2020-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - contratadas de 1/7/2019 até 30/6/2020.
3.2.11.45-6;i;credito_rural;demais;comercializacao;2019-07-01;2020-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos controlados - contratadas de 1/7/2019 até 30/6/2020.
3.2.11.46-3;i;credito_rural;pronaf;custeio;2019-02-01;2019-06-30;;;;;;2.39;4.2.10.81-0;Operações de custeio rural - Pronaf Adicional - contratadas de 1º/2/2019 até 30/6/2019.
3.2.11.47-0;i;credito_rural;pronamp;custeio;2019-02-01;2019-06-30;;;;;;2.06;4.2.10.82-7;Operações de custeio rural - Pronamp Adicional - contratadas de 1º/2/2019 até 30/6/2019.
3.2.11.48-7;i;credito_rural;;;2020-07-01;2021-06-30;;;;;armazenagem;;;Financiamentos para Armazenagem - MCR 6-4-5-"a" - de 1/7/2020 a 30/6/2021.
3.2.11.49-4;i;credito_rural;pronaf;custeio;2020-07-01;2021-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - de 1/7/2020 a 30/6/2021.
3.2.11.51-1;i;credito_rural;pronaf;investimento;2020-07-01;2021-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - de 1/7/2020 a 30/6/2021.
3.2.11.52-8;i;credito_rural;pronamp;custeio;2020-07-01;2021-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - de 1/7/2020 a 30/6/2021.
3.2.11.53-5;i;credito_rural;pronamp;investimento;2020-07-01;2021-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - de 1/7/2020 a 30/6/2021.
3.2.11.54-2;i;credito_rural;demais;custeio;2020-07-01;2021-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - de 1/7/2020 a 30/6/2021.
3.2.11.56-6;i;credito_rural;demais;investimento;2020-07-01;2021-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - de 1/7/2020 a 30/6/2021.
3.2.11.58-0;i;credito_rural;pronaf;custeio;2021-07-01;2022-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.60-7;i;credito_rural;pronaf;investimento;2021-07-01;2022-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.61-4;i;credito_rural;pronamp;custeio;2021-07-01;2022-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.62-1;i;credito_rural;pronamp;investimento;2021-07-01;2022-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.63-8;i;credito_rural;demais;custeio;2021-07-01;2022-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.65-2;i;credito_rural;demais;investimento;2021-07-01;2022-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.66-9;i;credito_rural;demais;comercializacao;2021-07-01;2022-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos controlados - de 1/7/2021 a 30/6/2022.
3.2.11.67-6;i;credito_rural;pronaf;custeio;2022-07-01;2023-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.69-0;i;credito_rural;pronaf;investimento;2022-07-01;2023-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.70-0;i;credito_rural;pronamp;custeio;2022-07-01;2023-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.71-7;i;credito_rural;pronamp;investimento;2022-07-01;2023-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.72-4;i;credito_rural;demais;custeio;2022-07-01;2023-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.74-8;i;credito_rural;demais;investimento;2022-07-01;2023-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.75-5;i;credito_rural;demais;comercializacao;2022-07-01;2023-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos controlados - de 1/7/2022 a 30/6/2023.
3.2.11.76-2;i;credito_rural;pronaf;custeio;2023-07-01;2024-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.78-6;i;credito_rural;pronaf;investimento;2023-07-01;2024-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.79-3;i;credito_rural;pronamp;custeio;2023-07-01;2024-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.80-3;i;credito_rural;pronamp;investimento;2023-07-01;2024-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.81-0;i;credito_rural;demais;custeio;2023-07-01;2024-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.83-4;i;credito_rural;demais;investimento;2023-07-01;2024-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.84-1;i;credito_rural;demais;comercializacao;2023-07-01;2024-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos controlados - de 1/7/2023 a 30/6/2024.
3.2.11.85-8;i;credito_rural;pronaf;custeio;2024-07-01;2025-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.87-2;i;credito_rural;pronaf;investimento;2024-07-01;2025-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.88-9;i;credito_rural;pronamp;custeio;2024-07-01;2025-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.89-6;i;credito_rural;pronamp;investimento;2024-07-01;2025-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.90-6;i;credito_rural;demais;custeio;2024-07-01;2025-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.92-0;i;credito_rural;demais;investimento;2024-07-01;2025-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.93-7;i;credito_rural;demais;comercializacao;2024-07-01;2025-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos controlados - de 1/7/2024 a 30/6/2025.
3.2.11.94-4;c;credito_rural;pronaf;custeio;2025-07-01;2026-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.11.95-1;c;credito_rural;pronaf;industrializacao;2025-07-01;2026-06-30;;;controlados;;;;;Operações de industrialização - Pronaf - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.11.96-8;c;credito_rural;pronaf;investimento;2025-07-01;2026-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.11.97-5;c;credito_rural;pronamp;custeio;2025-07-01;2026-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.11.98-2;c;credito_rural;pronamp;investimento;2025-07-01;2026-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.11.99-9;c;credito_rural;demais;custeio;2025-07-01;2026-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.12.00-8;c;credito_rural;demais;industrializacao;2025-07-01;2026-06-30;;;controlados;;;;;Operações de industrialização - Demais produtores - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.12.01-5;c;credito_rural;demais;investimento;2025-07-01;2026-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.12.02-2;c;credito_rural;demais;comercializacao;2025-07-01;2026-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos Controlados - de 1/7/2025 a 30/6/2026.
3.2.12.03-9;c;credito_rural;;industrializacao;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de industrialização – Recursos não controlados - de 1/7/2025 a 30/6/2026.
3.2.12.04-6;c;credito_rural;;custeio;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de custeio – Recursos não controlados - de 1/7/2025 a 30/6/2026.
3.2.12.05-3;c;credito_rural;;investimento;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de investimento – Recursos não controlados - de 1/7/2025 a 30/6/2026.
3.2.12.06-0;c;credito_rural;;comercializacao;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de comercialização – Recursos não controlados - de 1/7/2025 a 30/6/2026.
3.2.12.07-7;c;credito_rural;pronaf;custeio;2026-07-01;2027-06-30;;;controlados;;;;;Operações de custeio - Pronaf - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.08-4;c;credito_rural;pronaf;industrializacao;2026-07-01;2027-06-30;;;controlados;;;;;Operações de industrialização - Pronaf - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.09-1;c;credito_rural;pronaf;investimento;2026-07-01;2027-06-30;;;controlados;;;;;Operações de investimento - Pronaf - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.10-1;c;credito_rural;pronamp;custeio;2026-07-01;2027-06-30;;;controlados;;;;;Operações de custeio - Pronamp - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.11-8;c;credito_rural;pronamp;investimento;2026-07-01;2027-06-30;;;controlados;;;;;Operações de investimento - Pronamp - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.12-5;c;credito_rural;pronamp;comercializacao;2026-07-01;2027-06-30;;;controlados;;;;;Operações de comercialização - Pronamp - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.13-2;c;credito_rural;demais;custeio;2026-07-01;2027-06-30;;;controlados;;;;;Operações de custeio - Demais produtores - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.14-9;c;credito_rural;demais;industrializacao;2026-07-01;2027-06-30;;;controlados;;;;;Operações de industrialização - Demais produtores - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.15-6;c;credito_rural;demais;investimento;2026-07-01;2027-06-30;;;controlados;;;;;Operações de investimento - Demais produtores - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.16-3;c;credito_rural;demais;comercializacao;2026-07-01;2027-06-30;;;controlados;;;;;Operações de comercialização - Demais produtores - Recursos controlados - de 1/7/2026 a 30/6/2027.
3.2.12.17-0;c;credito_rural;;industrializacao;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de Industrialização - Recursos não controlados - de 1/7/2026 a 30/6/2027.
3.2.12.18-7;c;credito_rural;;custeio;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de custeio - Recursos não controlados - de 1/7/2026 a 30/6/2027.
3.2.12.19-4;c;credito_rural;;investimento;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de investimento - Recursos não controlados - de 1/7/2026 a 30/6/2027.
3.2.12.20-4;c;credito_rural;;comercializacao;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de comercialização - Recursos não controlados - de 1/7/2026 a 30/6/2027.
3.2.20.25-8;i;credito_rural;pronaf;;;;;;;;renegociacao_4028;2.0;4.2.10.60-7;Composição e renegociação de dívidas de operações de crédito rural no âmbito do Pronaf - Resolução nº 4.028/2011.
3.2.20.26-5;i;credito_rural;;;;;;;;;renegociacao;;;Renegociação de dívidas de operações de crédito rural - Outras.
3.2.20.27-2;i;credito_rural;;;;;;;;;desclassificada;;;Operações desclassificadas parcialmente.
3.2.20.28-9;i;3.2.20.29-6;;;;;;;;;res_5247;;;Operações de crédito rural - Res. CMN nº 5.247/2025 e MCR 6-8-33.
3.2.20.31-3;c;credito_rural;;;;;;;controlados;;res_5330;;;Operações de crédito rural - Resolução CMN nº 5.330/2026 - Recursos controlados.
3.2.20.32-0;c;credito_rural;;;;;;;nao_controlados;;res_5330;;;Operações de crédito rural - Resolução CMN nº 5.330/2026 - Recursos não controlados.
3.2.30.21-7;i;3.2.30.10-7;;;;;;;;;cpr;;;Aplicação mediante aquisição de Cédulas de Produto Rural (CPR).
3.2.30.23-1;i;3.2.30.10-7;;;;;;;;;cpr_5330;;;Aquisição de Cédulas de Produto Rural (CPR) - Resolução CMN nº 5.330/2026.
3.3.30.03-8;i;credito_rural;;custeio;;2017-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas até 30/6/2017.
3.3.30.06-9;i;credito_rural;;investimento;;2017-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas até 30/6/2017.
3.3.30.20-3;i;credito_rural;;custeio;2017-07-01;2018-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas de 1/7/2017 a 30/6/2018.
3.3.30.23-4;i;credito_rural;;investimento;2017-07-01;2018-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas de 1/7/2017 a 30/6/2018.
3.3.30.24-1;i;credito_rural;;custeio;2018-07-01;2019-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas de 1/7/2018 a 30/6/2019.
3.3.30.26-5;i;credito_rural;;comercializacao;2018-07-01;2019-06-30;;;;;;;;Operações de crédito rural de comercialização - contratadas de 1/7/2018 a 30/6/2019.
3.3.30.27-2;i;credito_rural;;investimento;2018-07-01;2019-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas de 1/7/2018 a 30/6/2019.
3.3.30.28-9;i;credito_rural;;custeio;2019-07-01;2020-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas de 1/7/2019 a 30/6/2020.
3.3.30.31-3;i;credito_rural;;investimento;2019-07-01;2020-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas de 1/7/2019 a 30/6/2020.
3.3.30.32-0;i;credito_rural;;custeio;2020-07-01;2021-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas de 1/7/2020 a 30/6/2021.
3.3.30.34-4;i;credito_rural;;comercializacao;2020-07-01;2021-06-30;;;;;;;;Operações de crédito rural de comercialização - contratadas de 1/7/2020 a 30/6/2021.
3.3.30.35-1;i;credito_rural;;investimento;2020-07-01;2021-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas de 1/7/2020 a 30/6/2021.
3.3.30.36-8;i;credito_rural;;custeio;2021-07-01;2022-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas de 1/7/2021 a 30/6/2022.
3.3.30.37-5;i;credito_rural;;industrializacao;2021-07-01;2022-06-30;;;;;;;;Operações de crédito rural de industrialização - contratadas de 1/7/2021 a 30/6/2022.
3.3.30.38-2;i;credito_rural;;comercializacao;2021-07-01;2022-06-30;;;;;;;;Operações de crédito rural de comercialização - contratadas de 1/7/2021 a 30/6/2022.
3.3.30.39-9;i;credito_rural;;investimento;2021-07-01;2022-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas de 1/7/2021 a 30/6/2022.
3.3.30.40-9;i;credito_rural;;custeio;2022-07-01;2023-06-30;;;;;;;;Operações de crédito rural de custeio - contratadas de 1/7/2022 a 30/6/2023.
3.3.30.41-6;i;credito_rural;;industrializacao;2022-07-01;2023-06-30;;;;;;;;Operações de crédito rural de industrialização - contratadas de 1/7/2022 a 30/6/2023.
3.3.30.42-3;i;credito_rural;;comercializacao;2022-07-01;2023-06-30;;;;;;;;Operações de crédito rural de comercialização - contratadas de 1/7/2022 a 30/6/2023.
3.3.30.43-0;i;credito_rural;;investimento;2022-07-01;2023-06-30;;;;;;;;Operações de crédito rural de investimento - contratadas de 1/7/2022 a 30/6/2023.
3.3.30.44-7;i;credito_rural;;custeio;2023-07-01;2024-06-30;;;controlados;;;;;Operações de crédito rural de custeio - Recursos Controlados - contratadas de 1/7/2023 a 30/6/2024.
3.3.30.46-1;i;credito_rural;;comercializacao;2023-07-01;2024-06-30;;;controlados;;;;;Operações de crédito rural de comercialização - Recursos Controlados - contratadas de 1/7/2023 a 30/6/2024.
3.3.30.47-8;i;credito_rural;;investimento;2023-07-01;2024-06-30;;;controlados;;;;;Operações de crédito rural de investimento - Recursos Controlados - contratadas de 1/7/2023 a 30/6/2024.
3.3.30.48-5;i;credito_rural;;custeio;2023-07-01;2024-06-30;;;nao_controlados;;;;;Operações de crédito rural de custeio contratadas de 1/7/2023 a 30/6/2024 nas condições dos recursos livres.
3.3.30.49-2;i;credito_rural;;industrializacao;2023-07-01;2024-06-30;;;nao_controlados;;;;;Operações de crédito rural de industrialização contratadas de 1/7/2023 a 30/6/2024 nas condições dos recursos livres.
3.3.30.50-2;i;credito_rural;;comercializacao;2023-07-01;2024-06-30;;;nao_controlados;;;;;Operações de crédito rural de comercialização contratadas de 1/7/2023 a 30/6/2024 nas condições dos recursos livres.
3.3.30.51-9;i;credito_rural;;investimento;2023-07-01;2024-06-30;;;nao_controlados;;;;;Operações de crédito rural de investimento contratadas de 1/7/2023 a 30/6/2024 nas condições dos recursos livres.
3.3.30.52-6;i;credito_rural;;custeio;2024-07-01;2025-06-30;;;controlados;;;;;Operações de crédito rural de custeio - Recursos Controlados - contratadas de 1/7/2024 a 30/6/2025.
3.3.30.53-3;i;credito_rural;;industrializacao;2024-07-01;2025-06-30;;;controlados;;;;;Operações de crédito rural de industrialização - Recursos Controlados - contratadas de 1/7/2024 a 30/6/2025.
3.3.30.54-0;i;credito_rural;;comercializacao;2024-07-01;2025-06-30;;;controlados;;;;;Operações de crédito rural de comercialização - Recursos Controlados - contratadas de 1/7/2024 a 30/6/2025.
3.3.30.55-7;i;credito_rural;;investimento;2024-07-01;2025-06-30;;;controlados;;;;;Operações de crédito rural de investimento - Recursos Controlados - contratadas de 1/7/2024 a 30/6/2025.
3.3.30.56-4;i;credito_rural;;custeio;2024-07-01;2025-06-30;;;nao_controlados;;;;;Operações de crédito rural de custeio contratadas de 1/7/2024 a 30/6/2025 nas condições dos recursos livres.
3.3.30.57-1;i;credito_rural;;industrializacao;2024-07-01;2025-06-30;;;nao_controlados;;;;;Operações de crédito rural de industrialização contratadas de 1/7/2024 a 30/6/2025 nas condições dos recursos livres.
3.3.30.58-8;i;credito_rural;;comercializacao;2024-07-01;2025-06-30;;;nao_controlados;;;;;Operações de crédito rural de comercialização contratadas de 1/7/2024 a 30/6/2025 nas condições dos recursos livres.
3.3.30.59-5;i;credito_rural;;investimento;2024-07-01;2025-06-30;;;nao_controlados;;;;;Operações de crédito rural de investimento contratadas de 1/7/2024 a 30/6/2025 nas condições dos recursos livres.
3.3.30.60-5;c;credito_rural;;custeio;2025-07-01;2026-06-30;;;controlados;;;;;Operações de crédito rural de custeio – Recursos Controlados - contratadas de 1/7/2025 a 30/6/2026.
3.3.30.61-2;c;credito_rural;;industrializacao;2025-07-01;2026-06-30;;;controlados;;;;;Operações de crédito rural de industrialização – Recursos Controlados - contratadas de 1/7/2025 a 30/6/2026.
3.3.30.62-9;c;credito_rural;;comercializacao;2025-07-01;2026-06-30;;;controlados;;;;;Operações de crédito rural de comercialização – Recursos Controlados - contratadas de 1/7/2025 a 30/6/2026.
3.3.30.63-6;c;credito_rural;;investimento;2025-07-01;2026-06-30;;;controlados;;;;;Operações de crédito rural de investimento – Recursos Controlados - contratadas de 1/7/2025 a 30/6/2026.
3.3.30.64-3;c;credito_rural;;custeio;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de crédito rural de custeio contratadas de 1/7/2025 a 30/6/2026 nas condições dos recursos livres.
3.3.30.65-0;c;credito_rural;;industrializacao;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de crédito rural de industrialização contratadas de 1/7/2025 a 30/6/2026 nas condições dos recursos livres.
3.3.30.66-7;c;credito_rural;;comercializacao;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de crédito rural de comercialização contratadas de 1/7/2025 a 30/6/2026 nas condições dos recursos livres.
3.3.30.67-4;c;credito_rural;;investimento;2025-07-01;2026-06-30;;;nao_controlados;;;;;Operações de crédito rural de investimento contratadas de 1/7/2025 a 30/6/2026 nas condições dos recursos livres.
3.3.30.68-1;c;credito_rural;;custeio;2026-07-01;2027-06-30;;;controlados;;;;;Operações de crédito rural de custeio - Recursos Controlados - contratadas de 1/7/2026 a 30/6/2027.
3.3.30.69-8;c;credito_rural;;industrializacao;2026-07-01;2027-06-30;;;controlados;;;;;Operações de crédito rural de industrialização - Recursos Controlados - contratadas de 1/7/2026 a 30/6/2027.
3.3.30.70-8;c;credito_rural;;comercializacao;2026-07-01;2027-06-30;;;controlados;;;;;Operações de crédito rural de comercialização - Recursos Controlados - contratadas de 1/7/2026 a 30/6/2027.
3.3.30.71-5;c;credito_rural;;investimento;2026-07-01;2027-06-30;;;controlados;;;;;Operações de crédito rural de investimento - Recursos Controlados - contratadas de 1/7/2026 a 30/6/2027.
3.3.30.72-2;c;credito_rural;;custeio;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de crédito rural de custeio contratadas de 1/7/2026 a 30/6/2027 - Recursos não controlados.
3.3.30.73-9;c;credito_rural;;industrializacao;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de crédito rural de industrialização contratadas de 1/7/2026 a 30/6/2027 - Recursos não controlados.
3.3.30.74-6;c;credito_rural;;comercializacao;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de crédito rural de comercialização contratadas de 1/7/2026 a 30/6/2027 - Recursos não controlados.
3.3.30.75-3;c;credito_rural;;investimento;2026-07-01;2027-06-30;;;nao_controlados;;;;;Operações de crédito rural de investimento contratadas de 1/7/2026 a 30/6/2027 - Recursos não controlados.
3.3.20.10-3;i;credito_rural;;;;;;;;;renegociacao;;;Renegociação de dívidas de operações de crédito rural.
3.3.20.16-5;i;3.3.20.17-2;;;;;;;;;res_5247_i;;;Operações de crédito rural - art. 2º, inciso I, da Res. CMN nº 5.247/2025 e MCR 6-8-34.
3.3.20.11-0;i;credito_rural;;;;;;;;;desclassificada;;;Operações desclassificadas parcialmente.
3.3.20.18-9;c;credito_rural;;;;;;;controlados;;res_5330;;;Operações de crédito rural - Resolução CMN nº 5.330/2026 - Recursos controlados.
3.3.20.19-6;c;credito_rural;;;;;;;nao_controlados;;res_5330;;;Operações de crédito rural - Resolução CMN nº 5.330/2026 - Recursos não controlados.
3.3.10.16-8;i;faculdade;;;;2018-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7-5-"b" e do MCR 6-7-5-A - adquiridas até 30/6/2018.
3.3.10.21-6;i;faculdade;;;2018-07-01;2019-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7-5-"b" e do MCR 6-7-5-A - adquiridas de 1/7/2018 a 30/6/2019.
3.3.10.22-3;i;faculdade;;;2019-07-01;2020-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7-5-"b"-I e do MCR 6-7-5-A - adquiridas de 1/7/2019 a 30/6/2020.
3.3.30.12-4;i;faculdade;;;2020-07-01;2021-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7-5-“b”-I e IV e do MCR 6-7-5-A - adquiridas de 1/7/2020 a 30/6/2021.
3.3.30.15-5;i;faculdade;;;2021-07-01;2022-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7 - adquiridas de 1/7/2021 a 30/6/2022.
3.3.30.16-2;i;faculdade;;;2022-07-01;2023-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7 - adquiridas de 1/7/2022 a 30/6/2023.
3.3.30.10-0;i;faculdade;;;;;;;;;cra;;;Aquisição de CRA nas condições do MCR 6-7.
3.3.30.13-1;i;faculdade;;;;;;;;;cda_wa;;;Aquisição de CDA e WA nas condições do MCR 6-7.
3.3.30.14-8;i;faculdade;;;;;;;;;fundo_garantidor;;;Aquisição de quotas de fundos garantidores de crédito com produtores rurais nas condições do MCR 6-7.
3.3.30.07-6;i;faculdade;;;;2021-04-30;;;;;mcr_6_8_6;;;Operações referidas no MCR 6-8-6, contratadas até 30/4/2021.
3.3.30.08-3;i;faculdade;;;;;;;;;cdca;;;Aquisição de CDCA nas condições do MCR 6-7.
3.3.30.17-9;i;faculdade;;;2023-07-01;2024-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7 - adquiridas de 1/7/2023 a 30/6/2024.
3.3.30.18-6;i;faculdade;;;2024-07-01;2025-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7 - adquiridas de 1/7/2024 a 30/6/2025.
3.3.30.19-3;i;faculdade;;;2025-07-01;2026-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7 - adquiridas de 1/7/2025 a 30/6/2026.
3.3.30.76-0;i;faculdade;;;2026-07-01;2027-06-30;;;;;cpr;;;Aquisição de CPR nas condições do MCR 6-7 - adquiridas de 1/7/2026 a 30/6/2027.
3.3.40.08-0;i;3.3.40.09-7;;;;;;;;;res_5247_ii_iii;;;Operações de crédito - art. 2º, incisos II e III, da Res. CMN nº 5.247/2025 e MCR 6-8-34.
3.3.40.11-4;i;faculdade;;;;;;;;;cpr_5330;;;Aquisição de Cédulas de Produto Rural (CPR) - Res. CMN nº 5.330/2026."""
_COLS = ("codigo", "tipo", "cat", "programa", "finalidade", "de", "ate", "taxa_max", "taxa_min", "recursos",
         "beneficiario", "linha", "fator", "ponderador", "titulo")
REGRAS = [dict(zip(_COLS, l.split(";")), fonte=FONTE_DO_CODIGO[l[:3]]) for l in REGRAS_TXT.splitlines()]
CODIGOS = {r["codigo"]: r for r in REGRAS}
ACEITOS = {k: sorted({r[k] for r in REGRAS} - {""}) for k in ("programa", "finalidade", "recursos", "beneficiario", "linha")}

# grupos que a gerência preenche (liberação, % de reembolso, prazo) -> onde a liberação nova entra
GRUPOS = {"RO_PRONAF": ("RO", "pronaf"), "RO_PRONAMP": ("RO", "pronamp"), "RO_DEMAIS": ("RO", "demais"),
          "POUP": ("POUP", "credito_rural"), "LCA": ("LCA", "credito_rural")}
DIR_MODALIDADES = ["geral", "pronaf", "pronamp", "poup", "lca_cr", "lca_livre"]

# linhas dos anexos: chave, código do Documento 6, descrição
LINHAS = {
    "RO": [("vsr_media", "1.1.10.00-9", "Média do VSR à vista (período de cálculo)"),
           ("propria", "2.1.10.00-8", "Exigibilidade - Própria"),
           ("capt_dir_geral", "2.1.20.00-5", "Captação DIR-Geral"),
           ("capt_dir_pronaf", "2.1.20.20-1", "Captação DIR-Pronaf"),
           ("capt_dir_pronamp", "2.1.20.30-4", "Captação DIR-Pronamp"),
           ("exig_total", "2.1.00.00-1", "Exigibilidade - Total"),
           ("sub_pronaf", "2.1.00.20-7", "Subexigibilidade Pronaf - Total"),
           ("sub_pronamp", "2.1.00.30-0", "Subexigibilidade Pronamp - Total"),
           ("geral", "2.1.00.40-3", "Exigibilidade Geral - Total"),
           ("ap_pronaf", "3.1.10.00-7", "Aplicado - Subexigibilidade Pronaf"),
           ("ap_geral", "3.1.30.00-1", "Aplicado - Exigibilidade Geral"),
           ("ap_pronamp", "3.1.40.00-8", "Aplicado - Subexigibilidade Pronamp"),
           ("ap_total", "3.1.00.00-0", "Aplicado - Total"),
           ("def_pronaf", "5.1.11.00-4", "Deficiência Pronaf"), ("exc_pronaf", "5.1.12.00-3", "Excesso Pronaf"),
           ("def_pronamp", "5.1.31.00-8", "Deficiência Pronamp"), ("exc_pronamp", "5.1.32.00-7", "Excesso Pronamp"),
           ("def_geral", "5.1.51.00-2", "Deficiência Geral"), ("exc_geral", "5.1.52.00-1", "Excesso Geral"),
           ("def_total", "5.1.41.00-5", "Deficiência Recursos Obrigatórios"),
           ("exc_total", "5.1.42.00-4", "Excesso Recursos Obrigatórios")],
    "POUP": [("vsr_media", "1.2.10.00-2", "Média do VSR da poupança rural (período de cálculo)"),
             ("propria", "2.2.10.00-1", "Exigibilidade própria (70%)"),
             ("capt_dir", "2.2.20.00-8", "Captação DIR-Poup"),
             ("exig_total", "2.2.00.00-4", "Exigibilidade - Total"),
             ("sub", "2.2.10.10-4", "Subexigibilidade - crédito rural"),
             ("exig_liquida", "2.2.50.00-9", "Exigibilidade - Líquida"),
             ("cpr_limite", "2.2.40.00-2", "Faculdade CPR (5% da própria)"),
             ("ap_cr", "3.2.10.00-0", "Aplicado - crédito rural"),
             ("ap_cpr", "3.2.30.00-4", "Aplicado - faculdade CPR"),
             ("ap_total", "3.2.00.00-3", "Aplicado - Total"),
             ("def_sub", "5.2.11.00-7", "Deficiência crédito rural"), ("exc_sub", "5.2.12.00-6", "Excesso crédito rural"),
             ("def_total", "5.2.01.00-0", "Deficiência Poupança Rural"),
             ("exc_total", "5.2.02.00-9", "Excesso Poupança Rural")],
    "LCA": [("lca_media", "1.3.10.00-5", "Média acumulada das LCA (junho até o mês anterior)"),
            ("propria", "2.3.10.00-4", "Direcionamento - Próprio (60%)"),
            ("capt_dir_cr", "2.3.20.05-6", "Captação DIR-LCA-CR"),
            ("capt_dir_livre", "2.3.20.15-9", "Captação DIR-LCA-Livre"),
            ("exig_total", "2.3.00.00-7", "Direcionamento - Total"),
            ("sub", "2.3.00.10-0", "Subdirecionamento crédito rural - Total"),
            ("faculdade", "2.3.00.20-3", "Faculdade - Total"),
            ("exc_ro", "3.3.20.15-8", "Excesso de Recursos Obrigatórios (só em junho)"),
            ("ap_cr", "3.3.10.00-3", "Aplicado - crédito rural"),
            ("ap_fac", "3.3.40.01-1", "Aplicado - faculdade (até o limite)"),
            ("ap_total", "3.3.00.00-6", "Aplicado - Total"),
            ("def_sub", "5.3.11.00-0", "Deficiência crédito rural"), ("exc_sub", "5.3.12.00-9", "Excesso crédito rural"),
            ("def_total", "5.3.01.00-3", "Deficiência LCA"), ("exc_total", "5.3.02.00-2", "Excesso LCA")],
}
for _cod, (_an, _b, _base, _pct, _exc, _desc) in LIMITES.items():  # linhas dos códigos de limite
    _l = LINHAS[_an]
    _l.insert(next(j for j, x in enumerate(_l) if x[0].startswith("ap_")), ("lim_" + _cod, _cod, _desc))
    if _exc:
        _l.insert(next(j for j, x in enumerate(_l) if x[0].startswith("ap_")), ("exc_" + _cod, "3.1.30.78-8",
                                                                                 "Excedente do 3.1.40.34-5 (vai para a Geral)"))
LINHAS_INFORMADAS = {"1.3.10.00-5", "3.3.20.15-8"}  # as demais linhas de LINHAS o Sisex calcula

# relatório dos diretores: uma linha por exigibilidade, sem soma (rótulo, anexo, chave do excesso, chave da deficiência)
RESUMO = [("Obrigatórios Pronaf", "RO", "exc_pronaf", "def_pronaf"),
          ("Obrigatórios Pronamp", "RO", "exc_pronamp", "def_pronamp"),
          ("Obrigatórios Geral", "RO", "exc_geral", "def_geral"),
          ("Poupança", "POUP", "exc_total", "def_total"),
          ("LCA", "LCA", "exc_total", "def_total")]

# operacional, visão por exigibilidade: blocos empilhados. Linha = chave do anexo (rótulo de LINHAS) ou [chave, rótulo].
# "mes:<baldes>" = saldo médio do mês das aplicações que caem nesses baldes (com ponderador, sem limite);
# "dir:<modalidades>" = aplicação DIR acumulada; "#" = título de uma das duas contas.
BLOCOS = [
    ["Obrigatórios Pronaf", "RO", [["sub_pronaf", "Subexigibilidade"], ["capt_dir_pronaf", "Captação DIR"],
                                   ["dir:pronaf", "Aplicação DIR"], ["mes:ap_pronaf", "Saldo médio do mês"],
                                   ["ap_pronaf", "Aplicação média acumulada"], "lim_3.1.13.48-2", "lim_3.1.12.09-8",
                                   ["exc_pronaf", "Excesso"], ["def_pronaf", "Deficiência"]]],
    ["Obrigatórios Pronamp", "RO", [["sub_pronamp", "Subexigibilidade"], ["capt_dir_pronamp", "Captação DIR"],
                                    ["dir:pronamp", "Aplicação DIR"], ["mes:ap_pronamp", "Saldo médio do mês"],
                                    ["ap_pronamp", "Aplicação média acumulada"], "lim_3.1.41.39-9", "lim_3.1.40.34-5",
                                    "exc_3.1.40.34-5", "lim_3.1.40.38-3", ["exc_pronamp", "Excesso"],
                                    ["def_pronamp", "Deficiência"]]],
    ["Obrigatórios Geral", "RO", [["geral", "Exigibilidade Geral"], ["capt_dir_geral", "Captação DIR"],
                                  ["dir:geral", "Aplicação DIR"], ["mes:ap_geral", "Saldo médio do mês"],
                                  ["ap_geral", "Aplicação média acumulada"], "lim_3.1.30.87-4", "lim_3.1.30.89-8",
                                  "lim_3.1.32.01-6", "lim_3.1.30.53-7", ["exc_geral", "Excesso"],
                                  ["def_geral", "Deficiência"]]],
    ["Poupança", "POUP", [["#", "Conta 1: total"], ["exig_total", "Exigibilidade"], ["capt_dir", "Captação DIR"],
                          ["dir:poup", "Aplicação DIR"], ["mes:ap_cr+ap_cpr", "Saldo médio do mês"],
                          ["ap_total", "Aplicação média acumulada"], ["cpr_limite", "Limite da CPR (5% da própria)"],
                          "lim_3.2.30.10-7", "lim_3.2.10.54-3", "lim_3.2.10.72-5", "lim_3.2.20.29-6",
                          ["exc_total", "Excesso"], ["def_total", "Deficiência"],
                          ["#", "Conta 2: mínimo em crédito rural (95%)"], ["sub", "Subexigibilidade crédito rural"],
                          ["ap_cr", "Aplicação em crédito rural"], ["exc_sub", "Excesso"], ["def_sub", "Deficiência"]]],
    ["LCA", "LCA", [["#", "Conta 1: total"], ["exig_total", "Direcionamento"], ["capt_dir_cr", "Captação DIR-LCA-CR"],
                    ["capt_dir_livre", "Captação DIR-LCA-Livre"], ["dir:lca_cr+lca_livre", "Aplicação DIR"],
                    ["mes:ap_cr+ap_fac", "Saldo médio do mês"], ["ap_total", "Aplicação média acumulada"],
                    ["faculdade", "Faculdade"], ["ap_fac", "Aplicado na faculdade (até o limite)"], "lim_3.3.40.09-7",
                    ["exc_ro", "Excesso de Obrigatórios (só em junho)"], ["exc_total", "Excesso"],
                    ["def_total", "Deficiência"], ["#", "Conta 2: mínimo em crédito rural (45%)"],
                    ["sub", "Subdirecionamento crédito rural"], ["ap_cr", "Aplicação em crédito rural"],
                    "lim_3.3.20.17-2", ["exc_sub", "Excesso"], ["def_sub", "Deficiência"]]],
]
# o que o relatório dos diretores leva da rodada (além da série oficial e dos parâmetros)
CHAVES_DIRETORES = ["versao", "data_base", "mes_base", "meses", "safras", "du", "bases", "pagos", "dir", "grupos",
                    "dir_modalidades", "limites", "balde", "resumo", "rodada", "sem_oficial"]


# ---------------------------------------------------------------- leitura e utilidades

def num(s):
    s = str(s if s is not None else "").strip()
    if not s:
        return 0.0
    if "," in s:  # formato brasileiro 1.234,56
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def dt(s):
    s = str(s).strip()
    if "/" in s:
        dd, mm, aa = s.split("/")
        return date(int(aa), int(mm), int(dd))
    return date.fromisoformat(s[:10])


def meses(ini, n):
    """n meses 'AAAA-MM' a partir de ini."""
    a, m = map(int, ini.split("-"))
    out = []
    for _ in range(n):
        out.append(f"{a:04d}-{m:02d}")
        a, m = (a + 1, 1) if m == 12 else (a, m + 1)
    return out


def mes_anterior(m):
    a, mm = map(int, m.split("-"))
    return f"{a - 1}-12" if mm == 1 else f"{a}-{mm - 1:02d}"


def safra_de(m):
    """Ano em que começa o ano-safra (julho) do mês 'AAAA-MM'."""
    return int(m[:4]) - (m[5:] < "07")


def mes_da_linha(r):
    v = (r.get("mes") or r.get("data") or "").strip()
    return dt(v).strftime("%Y-%m") if "/" in v or len(v) > 7 else v


def dias_uteis(m, feriados):
    a, mm = map(int, m.split("-"))
    return sum(1 for d in range(1, calendar.monthrange(a, mm)[1] + 1)
               if date(a, mm, d).weekday() < 5 and date(a, mm, d) not in feriados)


def ler(caminho):
    """Linhas de um CSV/TXT (separador ; , tab ou | detectado) ou da 1ª aba de um .xlsx, com cabeçalho em minúsculas."""
    caminho = Path(caminho)
    if caminho.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            sys.exit(f"{caminho}: para ler Excel instale o openpyxl (pip install openpyxl) ou salve como CSV")
        linhas = openpyxl.load_workbook(caminho, read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True)
        cab = [str(c or "").strip().lower() for c in next(linhas)]
        yield from ({k: "" if v is None else (v.isoformat() if hasattr(v, "isoformat") else str(v)) for k, v in zip(cab, l)}
                    for l in linhas if any(v is not None for v in l))
        return
    with open(caminho, encoding="utf-8-sig", newline="") as f:  # linha a linha: saldos_diarios pode ter milhões
        amostra = f.read(4096)
        f.seek(0)
        sep = max(";,\t|", key=amostra.count)
        for r in csv.DictReader(f, delimiter=sep):
            yield {(k or "").strip().lower(): v for k, v in r.items()}


def achar(pasta, nome, obrigatorio=False):
    """pasta/nome.csv, .txt ou .xlsx (o primeiro que existir)."""
    for ext in (".csv", ".txt", ".xlsx", ".xlsm"):
        if (Path(pasta) / f"{nome}{ext}").exists():
            return Path(pasta) / f"{nome}{ext}"
    if obrigatorio:
        sys.exit(f"falta {nome}.csv (ou .txt/.xlsx) em {pasta}")
    return None


def br(v):
    return f"{v:.2f}".replace(".", ",") if isinstance(v, float) else v


def br_mil(v):
    return f"{v:>18,.2f}".translate(str.maketrans(",.", ".,"))


def gravar(caminho, cabecalho, linhas):
    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(cabecalho)
        w.writerows([br(c) for c in l] for l in linhas)
    return caminho


def balde(codigo):
    cat = CODIGOS[codigo]["cat"]
    return LIMITES[cat][1] if cat in LIMITES else BALDE[CODIGOS[codigo]["fonte"]][cat]


def grupo(codigo):
    """Grupo da gerência a que o código pertence (para o % de reembolso)."""
    if not codigo:
        return None
    fonte = CODIGOS[codigo]["fonte"]
    if fonte != "RO":
        return fonte
    return {"ap_pronaf": "RO_PRONAF", "ap_pronamp": "RO_PRONAMP"}.get(balde(codigo), "RO_DEMAIS")


def classificar(op):
    """Código do Documento 6 da operação pelas regras de REGRAS_TXT. Devolve (código, empatados).

    Entra a regra cujas condições a operação cumpre (fonte, programa, finalidade, recursos, beneficiário, linha
    especial, período de contratação e taxa); havendo mais de uma, vale a mais específica: período mais curto,
    depois menor teto de taxa, depois mais condições. Sem regra -> "" (não cumpre exigibilidade)."""
    cands = []
    for r in REGRAS:
        if r["fonte"] != op["fonte"] or r["linha"] != op.get("linha", ""):
            continue
        if any(r[k] and r[k] != op.get(k, "") for k in ("programa", "finalidade", "recursos", "beneficiario")):
            continue
        d, tx = op["data_contratacao"], op["taxa_aa"]
        if (r["de"] and d < r["de"]) or (r["ate"] and d > r["ate"]):
            continue
        if (r["taxa_max"] and tx > float(r["taxa_max"])) or (r["taxa_min"] and tx <= float(r["taxa_min"])):
            continue
        dias = (date.fromisoformat(r["ate"] or "2099-12-31") - date.fromisoformat(r["de"] or "1900-01-01")).days
        teto = float(r["taxa_max"]) if r["taxa_max"] else 999.0
        n = sum(bool(r[k]) for k in ("programa", "finalidade", "recursos", "taxa_min"))
        cands.append(((dias, teto, -n), r["codigo"]))
    if not cands:
        return "", []
    cands.sort()
    return cands[0][1], [c for k, c in cands if k == cands[0][0]]


def ler_classificacao(r, onde, data_padrao=None):
    """Lê as colunas de classificação de uma linha de operacoes/liberacoes. Devolve (código, motivo)."""
    fonte = (r.get("fonte") or "").strip().upper()
    if fonte not in FONTES:
        sys.exit(f"{onde}: fonte '{fonte}' inválida. Aceitas: {', '.join(FONTES)}")
    if fonte == "LIVRE":
        return "", "livre"
    op = {"fonte": fonte, "taxa_aa": num(r.get("taxa_aa"))}
    for k in ("programa", "finalidade", "recursos", "beneficiario", "linha"):
        op[k] = (r.get(k) or "").strip().lower()
        if op[k] and op[k] not in ACEITOS[k]:
            sys.exit(f"{onde}: {k} '{op[k]}' inválido. Aceitos: {', '.join(ACEITOS[k])}")
    if fonte == "RO":
        op["beneficiario"] = op["beneficiario"] or "produtor"
    dc = (r.get("data_contratacao") or "").strip()
    op["data_contratacao"] = (dt(dc) if dc else data_padrao or date(1900, 1, 1)).isoformat()
    regra, emp = classificar(op)
    informado = (r.get("codigo") or "").strip()
    if informado:
        if informado not in CODIGOS or CODIGOS[informado]["fonte"] != fonte:
            sys.exit(f"{onde}: código {informado} não é código de aplicação do anexo da fonte {fonte}")
        return informado, "informado" if informado == regra else f"informado (regra daria {regra or 'sem código'})"
    if not regra:
        return "", "sem código no Documento 6"
    return regra, "regra" if len(emp) == 1 else f"regra empatada com {', '.join(emp[1:])}"


# ---------------------------------------------------------------- projeção operação a operação

def projetar(ops, parcelas, libs, cenario, data_base, ate, feriados=frozenset(), hist=None):
    """Saldo diário de cada operação até a liquidação.

    Devolve (res, pagos): res = {(id, mes): [codigo, saldo_medio]}, pagos = {(grupo, mes): valor}.
    Saldo médio = soma dos saldos dos dias úteis / dias úteis do mês, com zero antes da 1ª liberação e depois da
    liquidação (Doc 6, Anexo I, 4.6-b). O código do mês é o do dia 15 (Doc 6, Anexo I, 12.2-b).
    hist: saída de realizado_diario; dá os dias úteis do mês da data-base antes dela e o código até ali.
    """
    hist = hist or {}
    ops = {k: dict(v) for k, v in ops.items()}
    pag = defaultdict(dict)
    for i, d, v in parcelas:
        if d >= data_base:
            pag[i][d] = pag[i].get(d, 0.0) + v
    ultima = {i: max(p) for i, p in pag.items()}
    lib = defaultdict(dict)
    for i, d, v, extra in libs:
        if d >= data_base:
            lib[i][d] = lib[i].get(d, 0.0) + v
            if i not in ops:  # operação nova, ainda sem saldo
                ops[i] = dict(extra, saldo=0.0)
    liq, recl = {}, defaultdict(dict)
    for c in cenario:
        if c["tipo"] == "liquidacao":
            liq[c["id"]] = c["data"]
        elif c["tipo"] == "reclassificacao":
            recl[c["id"]][c["data"]] = c["codigo"]
        else:
            sys.exit(f"cenário: tipo desconhecido '{c['tipo']}' (use liquidacao ou reclassificacao)")
    du = {}

    def n_uteis(m):
        if m not in du:
            du[m] = dias_uteis(m, feriados)
        return du[m]

    util = lambda d: d.weekday() < 5 and d not in feriados
    res, pagos = {}, defaultdict(float)
    # shortcut: loop diário em Python (60 mil ops x 600 dias ~ alguns minutos); vetorizar se ficar lento
    for i, op in ops.items():
        saldo, cod = op["saldo"], op["codigo"]
        fat_dia = (1 + op["taxa_aa"] / 100) ** (1 / 365)  # shortcut: juros compostos por dia corrido; ajustar à regra real
        soma, cls = defaultdict(float), {}
        m0 = data_base.strftime("%Y-%m")
        h = hist.get((i, m0))
        cls[m0] = h[0] if h else cod
        soma[m0] += h[1] * n_uteis(m0) if h else 0.0  # dias antes da data-base sem linha no saldo diário = zero
        ult_lib = max(lib[i], default=None)
        d = data_base
        while d <= ate:
            m = d.strftime("%Y-%m")
            if d > data_base:
                saldo *= fat_dia
            if d in recl[i]:
                cod = recl[i][d]
            saldo += lib[i].get(d, 0.0)
            if d in pag[i]:
                pago = saldo if d == ultima[i] else min(pag[i][d], saldo)  # a última parcela liquida
                saldo -= pago
                pagos[(grupo(cod), m)] += pago
            if liq.get(i) == d:
                saldo = 0.0
            if m not in cls or d.day <= 15:
                cls[m] = cod
            if util(d):
                soma[m] += saldo
            if saldo <= 0.0 and (ult_lib is None or d >= ult_lib):
                break  # liquidada: daqui em diante só zeros
            d += timedelta(days=1)
        for m, s in soma.items():
            if s > 0:
                res[(i, m)] = [cls[m], s / n_uteis(m)]
    return res, {k: v for k, v in pagos.items() if k[0]}


def realizado_diario(linhas, data_base, feriados=frozenset()):
    """Realizado de cada operação a partir do saldo diário (arquivo saldos_diarios).

    Devolve (hist, avisos): hist = {(id, mes): [codigo, saldo_medio]} de julho do ano-safra até o mês da data-base,
    neste só com os dias antes dela. Saldo médio = soma dos saldos dos dias úteis / dias úteis do mês, com zero no
    dia útil sem linha (Doc 6, Anexo I, 4.6-b). Código do mês = regra aplicada à linha do último dia útil até o dia
    15, pois a situação não muda em dia não útil (Doc 6, Anexo I, 12.2-b); operação que só aparece depois do dia 15
    usa a 1ª linha do mês.
    """
    ini = date(safra_de(data_base.strftime("%Y-%m")), 7, 1)
    cache, acc = {}, {}  # acc[(id, mes)] = [soma, nota da linha do código, código]
    fora, sem = 0, set()
    chave = ("fonte", "programa", "finalidade", "recursos", "beneficiario", "linha", "data_contratacao", "taxa_aa", "codigo")
    for n, r in enumerate(linhas, start=2):
        d = dt(r["data"])
        if not ini <= d < data_base:
            continue
        if d.weekday() > 4 or d in feriados:
            fora += 1
            continue
        k = (r["id"].strip(), d.strftime("%Y-%m"))
        a = acc.setdefault(k, [0.0, -99, ""])
        a[0] += num(r["saldo"])
        nota = d.day if d.day <= 15 else -d.day  # vale o dia mais perto do 15 por baixo; sem ele, o 1º do mês
        if nota > a[1]:
            c = tuple((r.get(x) or "").strip() for x in chave)
            if c not in cache:  # a classificação quase não muda de um dia para o outro
                cache[c] = ler_classificacao(r, f"saldos_diarios linha {n}")
            a[1:] = nota, cache[c][0]
            if cache[c][1].startswith("sem código"):
                sem.add(k[0])
    du = {}
    hist = {}
    for (i, m), (soma, _, cod) in acc.items():
        du[m] = du.get(m) or dias_uteis(m, feriados)
        hist[(i, m)] = [cod, soma / du[m]]
    avisos = [f"saldos_diarios: {fora} linhas em dia não útil ignoradas"] if fora else []
    if sem:
        avisos.append(f"saldos_diarios: {len(sem)} operações sem código no Documento 6 em algum mês (não contam)")
    return hist, avisos


def carregar_entradas(pasta, data_base, feriados):
    """Lê a pasta de entradas. Detalhe dos arquivos no README."""
    p = Path(pasta)
    ops, classif = {}, []
    for n, r in enumerate(ler(achar(p, "operacoes", True)), start=2):
        i = r["id"].strip()
        cod, motivo = ler_classificacao(r, f"operacoes linha {n}")
        ops[i] = {"codigo": cod, "saldo": num(r["saldo"]), "taxa_aa": num(r.get("taxa_aa"))}
        classif.append([i, r["fonte"].strip().upper(), cod, CODIGOS[cod]["tipo"] if cod else "", motivo])
    parcelas = [(r["id"].strip(), dt(r["data"]), num(r["valor"])) for r in ler(achar(p, "parcelas", True))]
    libs = []
    for n, r in enumerate(ler(achar(p, "liberacoes")) if achar(p, "liberacoes") else [], start=2):
        i, d = r["id"].strip(), dt(r["data"])
        extra = {}
        if i not in ops and i not in {x[0] for x in libs}:  # operação nova: classifica como em operacoes
            cod, motivo = ler_classificacao(r, f"liberacoes linha {n} (operação nova precisa de fonte e classificação)", d)
            extra = {"codigo": cod, "taxa_aa": num(r.get("taxa_aa"))}
            classif.append([i, r["fonte"].strip().upper(), cod, CODIGOS[cod]["tipo"] if cod else "", motivo])
        libs.append((i, d, num(r["valor"]), extra))
    cenario = []
    for n, r in enumerate(ler(achar(p, "cenario")) if achar(p, "cenario") else [], start=2):
        c = {"tipo": r["tipo"].strip().lower(), "id": r["id"].strip(), "data": dt(r["data"]),
             "codigo": (r.get("codigo") or "").strip()}
        if c["tipo"] == "reclassificacao" and c["codigo"] and c["codigo"] not in CODIGOS:
            sys.exit(f"cenario linha {n}: código {c['codigo']} não é código de aplicação (vazio = sai da exigibilidade)")
        cenario.append(c)
    sd = achar(p, "saldos_diarios")
    hist, avisos = realizado_diario(ler(sd), data_base, feriados) if sd else ({}, [])
    vsr = defaultdict(list)
    for r in ler(achar(p, "vsr", True)):
        vsr[mes_da_linha(r)].append((num(r.get("vsr_avista")), num(r.get("vsr_poupanca"))))
    lca = defaultdict(lambda: [0.0, 0])  # mes -> [soma, peso]
    for r in ler(achar(p, "lca", True)):
        m = mes_da_linha(r)
        peso = 1 if r.get("data") else dias_uteis(m, feriados)  # linha diária pesa 1; mensal pesa os dias úteis
        lca[m][0] += num(r.get("saldo")) * peso
        lca[m][1] += peso
    dirs = defaultdict(dict)
    for n, r in enumerate(ler(achar(p, "dir")) if achar(p, "dir") else [], start=2):
        mod = r["modalidade"].strip().lower()
        if mod not in DIR_MODALIDADES:
            sys.exit(f"dir linha {n}: modalidade '{mod}' inválida. Aceitas: {', '.join(DIR_MODALIDADES)}")
        c, a = dirs[mod].get(mes_da_linha(r), [0.0, 0.0])
        dirs[mod][mes_da_linha(r)] = [c + num(r.get("captacao")), a + num(r.get("aplicacao"))]
    params = json.loads((p / "parametros.json").read_text(encoding="utf-8")) if (p / "parametros.json").exists() else {}
    oficial = defaultdict(lambda: defaultdict(float))  # posicao -> codigo -> valor oficial (Sisex) da posição
    for r in ler(achar(p, "oficial")) if achar(p, "oficial") else []:
        oficial[mes_da_linha({"mes": r.get("posicao") or r.get("mes")})][r["codigo"].strip()] += num(r["valor"])
    return ops, parcelas, libs, cenario, hist, vsr, lca, dict(dirs), params, avisos, classif, {m: dict(v) for m, v in oficial.items()}


# ---------------------------------------------------------------- bases de cálculo (MCR 6-2-2, 6-4-2, 6-7-4)

def pct_ro(safra, tipo_if):
    """MCR 6-2-3-B. safra = ano em que começa o período de cumprimento (julho)."""
    if tipo_if == "cooperativa":
        return {2025: 0.06, 2026: 0.13, 2027: 0.22}.get(safra, 0.315)
    return 0.315


def bases(safra, vsr, lca, tipo_if, avisos):
    """Exigibilidades próprias do ano-safra. VSR: média de jul/S-1 a jun/S. LCA: média de jun/S até o mês anterior à posição."""
    calc = meses(f"{safra - 1}-07", 12)
    linhas = [v for m in calc for v in vsr.get(m, [])]
    faltam = [m for m in calc if m not in vsr]
    if faltam:
        avisos.append(f"{safra}/{safra + 1}: VSR sem os meses {', '.join(faltam)} (média só com os que existem)")
    va = sum(v[0] for v in linhas) / len(linhas) if linhas else 0.0
    vp = sum(v[1] for v in linhas) / len(linhas) if linhas else 0.0
    ro = max(0.0, va - DEDUCAO) * pct_ro(safra, tipo_if)
    out = {"vsr_avista": va, "vsr_poupanca": vp, "ro_propria": ro if ro > ISENCAO else 0.0,
           "poup_propria": 0.70 * vp, "lca_media": {}, "lca_propria": {}}
    for m in meses(f"{safra}-07", 12):
        janela = meses(f"{safra}-06", meses(f"{safra}-06", 13).index(m))  # jun/S até o mês anterior
        soma = sum(lca[k][0] for k in janela if k in lca)
        peso = sum(lca[k][1] for k in janela if k in lca)
        media = soma / peso if peso else 0.0
        lp = max(0.0, media - DEDUCAO) * 0.60
        out["lca_media"][m] = media
        out["lca_propria"][m] = lp if lp > ISENCAO else 0.0
    if not any(k in lca for k in meses(f"{safra}-06", 12)):
        avisos.append(f"{safra}/{safra + 1}: sem saldo de LCA de junho em diante")
    return out


# ---------------------------------------------------------------- parâmetros da gerência e fórmulas do Sisex
# Estas funções têm cópia idêntica em JavaScript no HTML (simulação); o HTML confere que dão o mesmo resultado.

def pct_de(params, g, m):
    """% de reembolso do grupo no mês (premissa gerencial): o da célula, senão o do grupo, senão o geral, senão 80."""
    for v in (params.get("pct", {}).get(g, {}).get(m), params.get("pct_grupo", {}).get(g), params.get("pct_geral")):
        if v is not None:
            return v
    return 80


def com_parametros(apl, pagos, params, meses_h, mes_base):
    """Soma às aplicações as liberações da gerência (voltam em parcela única após o prazo) e o que não for
    reembolsado pelo % de reembolso (fica no saldo até o fim do horizonte). Valores da gerência em reais."""
    out = {k: dict(v) for k, v in apl.items()}
    for g, (fo, ca) in GRUPOS.items():
        prazo = max(1, int(params.get("prazo", {}).get(g, 12)))
        alvo = out.setdefault(f"{fo}|{ca}", {})
        for j, mj in enumerate(meses_h):
            if mj < mes_base:
                continue
            lib = params.get("liberacao", {}).get(g, {}).get(mj, 0)
            pct = pct_de(params, g, mj)
            u = (1 - pct / 100) * pagos.get(g, {}).get(mj, 0)
            if not lib and not u:
                continue
            for k in range(j, len(meses_h)):
                dl = lib * (0.5 if k in (j, j + prazo) else 1) if k <= j + prazo else 0  # meio mês na ida e na volta
                alvo[meses_h[k]] = alvo.get(meses_h[k], 0) + dl + u * (0.5 if k == j else 1)
    return out


def dir_final(dirs, params):
    out = {mod: {m: list(v) for m, v in dirs.get(mod, {}).items()} for mod in DIR_MODALIDADES}
    for mod, porm in params.get("dir", {}).items():
        for m, (c, a) in porm.items():
            out.setdefault(mod, {})[m] = [c, a]
    return out


def acum(serie, safra, m, du):
    """Média acumulada de julho até m, ponderada pelos dias úteis (posição do Documento 6)."""
    s = p = 0.0
    for k in meses(f"{safra}-07", 12):
        if k > m:
            break
        w = du.get(k, 0)
        s += serie.get(k, 0) * w
        p += w
    return s / p if p else 0.0


def aplicar(an, d, c):
    """Soma as aplicações de cada balde (Pronaf, Pronamp, Geral; crédito rural, CPR, faculdade), cortando cada
    código de limite no % da sua base (LIMITES). O excesso do 3.1.40.34-5 vai para a Geral (código 3.1.30.78-8)."""
    for k in set(BALDE[an].values()) | {v[1] for v in LIMITES.values() if v[0] == an}:
        d[k] = 0.0
    for cat, b in BALDE[an].items():
        d[b] += c(f"{an}|{cat}")
    for cod, (a, b, base, pct, exc, _) in LIMITES.items():
        if a == an:
            v = c(f"{an}|{cod}")
            d["lim_" + cod] = min(v, pct * d[base])
            d[b] += d["lim_" + cod]
            if exc:
                d["exc_" + cod] = v - d["lim_" + cod]
                d[exc] += d["exc_" + cod]


def sisex(safra, m, apl, dirs, base, du):
    """Posição do mês m nos anexos II (RO), III (Poupança) e IV (LCA)."""
    c = lambda k: acum(apl.get(k, {}), safra, m, du)
    dc = lambda mod, i: acum({k: v[i] for k, v in dirs.get(mod, {}).items()}, safra, m, du)

    p = base["ro_propria"]
    ro = {"vsr_media": base["vsr_avista"], "propria": p, "capt_dir_geral": dc("geral", 0),
          "capt_dir_pronaf": dc("pronaf", 0), "capt_dir_pronamp": dc("pronamp", 0)}
    ro["exig_total"] = p + ro["capt_dir_geral"] + ro["capt_dir_pronaf"] + ro["capt_dir_pronamp"]
    ro["sub_pronaf"] = 0.35 * p + ro["capt_dir_pronaf"]                    # 6-2-10
    ro["sub_pronamp"] = 0.50 * p + ro["capt_dir_pronamp"]                  # 6-2-8
    ro["geral"] = p - 0.35 * p - 0.50 * p + ro["capt_dir_geral"]
    aplicar("RO", ro, c)
    ro["ap_pronaf"] += dc("pronaf", 1)
    ro["ap_pronamp"] += dc("pronamp", 1)
    ro["ap_geral"] += dc("geral", 1)
    ro["ap_total"] = ro["ap_pronaf"] + ro["ap_geral"] + ro["ap_pronamp"]
    ro["def_pronaf"] = max(0.0, ro["sub_pronaf"] - ro["ap_pronaf"])
    ro["exc_pronaf"] = max(0.0, ro["ap_pronaf"] - ro["sub_pronaf"])
    ro["def_pronamp"] = max(0.0, ro["sub_pronamp"] - ro["ap_pronamp"])
    ro["exc_pronamp"] = max(0.0, ro["ap_pronamp"] - ro["sub_pronamp"])
    ro["def_geral"] = max(0.0, ro["exig_total"] - (ro["ap_total"] + ro["def_pronaf"] + ro["def_pronamp"]))
    ro["exc_geral"] = max(0.0, ro["ap_geral"] - ro["geral"])
    ro["def_total"] = ro["def_geral"] + ro["def_pronaf"] + ro["def_pronamp"]
    ro["exc_total"] = max(0.0, ro["exc_pronaf"] + ro["exc_pronamp"] + ro["ap_geral"] - ro["geral"])

    pp = base["poup_propria"]
    po = {"vsr_media": base["vsr_poupanca"], "propria": pp, "capt_dir": dc("poup", 0)}
    po["exig_total"] = pp + po["capt_dir"]
    po["sub"] = 0.95 * pp + po["capt_dir"]                                  # 6-4-10
    po["cpr_limite"] = 0.05 * pp                                            # 6-4-11; 2.2.40.00-2
    po["exig_liquida"] = po["exig_total"] - dc("poup", 1)                   # 2.2.50.00-9
    aplicar("POUP", po, c)
    po["ap_cr"] += dc("poup", 1)
    po["ap_total"] = po["ap_cr"] + po["ap_cpr"]
    po["def_sub"] = max(0.0, po["sub"] - po["ap_cr"])
    po["exc_sub"] = max(0.0, po["ap_cr"] - po["sub"])
    po["def_total"] = max(0.0, po["exig_total"] - po["ap_total"])
    po["exc_total"] = max(0.0, po["ap_total"] - po["exig_total"])

    lp = base["lca_propria"].get(m, 0.0)
    lc = {"lca_media": base["lca_media"].get(m, 0.0), "propria": lp, "capt_dir_cr": dc("lca_cr", 0),
          "capt_dir_livre": dc("lca_livre", 0)}
    lc["exig_total"] = lp + lc["capt_dir_cr"] + lc["capt_dir_livre"]
    lc["sub"] = 0.45 * lp + lc["capt_dir_cr"]                               # 6-7-7-a
    lc["faculdade"] = 0.55 * lp + lc["capt_dir_livre"]                      # 6-7-7-b
    lc["exc_ro"] = ro["exc_total"] if m[5:] == "06" else 0.0               # 6-7-9-h: excesso ao final do período
    aplicar("LCA", lc, c)
    lc["ap_cr"] += dc("lca_cr", 1) + lc["exc_ro"]
    lc["ap_fac"] = min(lc["ap_fac"] + dc("lca_livre", 1), lc["faculdade"])  # 3.3.40.01-1
    lc["ap_total"] = lc["ap_cr"] + lc["ap_fac"]
    lc["def_sub"] = max(0.0, lc["sub"] - lc["ap_cr"])
    lc["exc_sub"] = max(0.0, lc["ap_cr"] - lc["sub"])
    lc["def_total"] = max(0.0, lc["exig_total"] - lc["ap_total"])
    lc["exc_total"] = max(0.0, lc["ap_total"] - lc["exig_total"])
    return {"RO": ro, "POUP": po, "LCA": lc}


def calcular(dados, params, apl=None):
    """{safra: {mes: sisex}} das duas safras do horizonte (apl: padrão = série calculada)."""
    apl = com_parametros(apl or dados["apl"], dados["pagos"], params, dados["meses"], dados["mes_base"])
    dirs = dir_final(dados["dir"], params)
    return {str(s): {m: sisex(s, m, apl, dirs, dados["bases"][str(s)], dados["du"]) for m in meses(f"{s}-07", 12)}
            for s in dados["safras"]}


def por_categoria(apl_cod):
    """Saldo x fator (ponderador da seção 4) por fonte|categoria do cálculo."""
    apl = defaultdict(dict)
    for cod, porm in apl_cod.items():
        r = CODIGOS[cod]
        for m, v in porm.items():
            k = f"{r['fonte']}|{r['cat']}"
            apl[k][m] = apl[k].get(m, 0.0) + v * float(r["fator"] or 1)
    return dict(apl)


def montar(pasta, data_base, tipo_if="banco"):
    """Lê as entradas, projeta e devolve (dados, res_operacoes, avisos)."""
    db = dt(data_base) if isinstance(data_base, str) else data_base
    feriados_arq = achar(pasta, "feriados")
    feriados = {dt(r["data"]) for r in ler(feriados_arq)} if feriados_arq else set()
    ops, parcelas, libs, cenario, hist, vsr, lca, dirs, params, avisos, classif, oficial = carregar_entradas(pasta, db, feriados)
    s0 = safra_de(db.strftime("%Y-%m"))
    meses_h = meses(f"{s0}-07", 24)
    res, pagos = projetar(ops, parcelas, libs, cenario, db, date(s0 + 2, 6, 30), feriados, hist)
    res = {**{k: v for k, v in hist.items() if v[1] > 0}, **res}  # realizado calculado + projeção
    apl_cod = defaultdict(lambda: defaultdict(float))
    for (i, m), (cod, s) in res.items():
        if cod:
            apl_cod[cod][m] += s
    mes_base = db.strftime("%Y-%m")
    du = {m: dias_uteis(m, feriados) for m in meses_h}
    # série oficial: oficial nos meses passados, projeção nos futuros (o que os diretores veem)
    apl_of = {cod: dict(serie) for cod, serie in apl_cod.items()}
    if oficial:
        for cod in {c for v in oficial.values() for c in v if c in CODIGOS} | set(apl_cod):
            pos = {m: v[cod] for m, v in oficial.items() if cod in v}
            apl_of.setdefault(cod, {}).update(serie_do_oficial(pos, s0, mes_base, du, apl_cod.get(cod, {})))
        sem_of = [m for m in meses_h if m < mes_base and m not in oficial]
        if sem_of:
            avisos.append(f"oficial sem a posição {', '.join(sem_of)}: nesses meses o relatório dos diretores usa o calculado")
    for motivo, txt in (("sem código", "sem código no Documento 6 (não contam)"), ("empatada", "com regra empatada"),
                        ("regra daria", "com código informado diferente da regra")):
        n = sum(1 for c in classif if motivo in c[4])
        if n:
            avisos.append(f"{n} operações {txt}: ver classificacao.csv")
    com_saldo = {m for _, m in hist}
    sem_real = [m for m in meses_h if m < mes_base and m not in com_saldo]
    if sem_real:
        avisos.append(f"sem saldos_diarios para {', '.join(sem_real)}: esses meses entram com zero")
    pg = defaultdict(dict)
    for (g, m), v in pagos.items():
        pg[g][m] = v
    dados = {"versao": VERSAO, "data_base": db.isoformat(), "mes_base": mes_base, "tipo_if": tipo_if,
             "meses": meses_h, "safras": [s0, s0 + 1], "du": du, "oficial": oficial, "apl_oficial": por_categoria(apl_of),
             "bases": {str(s): bases(s, vsr, lca, tipo_if, avisos) for s in (s0, s0 + 1)},
             "apl": por_categoria(apl_cod), "pagos": dict(pg), "dir": dirs, "params": params,
             "grupos": list(GRUPOS), "dir_modalidades": DIR_MODALIDADES, "limites": LIMITES, "balde": BALDE, "linhas": LINHAS, "resumo": RESUMO, "blocos": BLOCOS,
             "chaves_diretores": CHAVES_DIRETORES, "sem_oficial": [m for m in meses_h if m < mes_base and m not in oficial]}
    return dados, res, avisos, classif, apl_cod


CAB_CODIGOS = ["ano_safra", "posicao", "R_P", "anexo", "codigo", "tipo", "titulo", "saldo_medio_mes", "calculado",
               "oficial", "diferenca"]


def serie_do_oficial(pos, s, mes_base, du, calc):
    """Saldo médio de cada mês passado do ano-safra s tirado das posições oficiais acumuladas de um código.

    A posição é a média acumulada de julho até o mês pelos dias úteis (Doc 6, Anexo I, 4.6-c/d), então o mês sai de
    posição x dias úteis acumulados menos o que os meses anteriores já somam. Mês sem posição oficial usa o calculado."""
    out, soma, peso = {}, 0.0, 0
    for m in meses(f"{s}-07", 12):
        if m >= mes_base:
            break
        peso += du[m]
        out[m] = (pos[m] * peso - soma) / du[m] if m in pos else calc.get(m, 0.0)
        soma += out[m] * du[m]
    return out


def com_oficial(linhas, oficial, dados):
    """Preenche oficial e diferença nas linhas do codigos.csv das posições que vieram no arquivo oficial.
    Código que só existe num dos lados aparece com o outro em zero; posição sem oficial fica em branco."""
    vistos = set()
    for l in linhas:
        if l[1] in oficial:
            of = oficial[l[1]].get(l[4], 0.0)
            l[9:] = [round(of, 2), round(of - (l[8] or 0.0), 2)]
            vistos.add((l[1], l[4]))
    desc = {cod: d for ls in LINHAS.values() for _, cod, d in ls}
    for m, cod, of in sorted((m, c, v) for m, porc in oficial.items() for c, v in porc.items()):
        if (m, cod) in vistos or m not in dados["meses"]:
            continue
        r = CODIGOS.get(cod, {})
        linhas.append([f"{safra_de(m)}/{safra_de(m) + 1}", m, "R" if m < dados["mes_base"] else "P",
                       {"1": "RO", "2": "POUP", "3": "LCA"}.get(cod[2:3], ""), cod,
                       "informado" if r.get("tipo") == "i" or cod in LINHAS_INFORMADAS else "calculado",
                       r.get("titulo") or desc.get(cod, ""), "", 0.0, round(of, 2), round(of, 2)])
    return linhas


def linhas_codigos(dados, resultado, apl_cod):
    """Todos os códigos com valor: aplicação (média acumulada de julho até a posição, Doc 6 Anexo I 4.6-c/d),
    ponderadores (seção 4) e as linhas de exigibilidade, limite, deficiência e excesso. Oficial e diferença: com_oficial."""
    tipo = {"i": "informado", "c": "calculado"}
    out = []
    for s, porm in resultado.items():
        for m, v in porm.items():
            rp = "R" if m < dados["mes_base"] else "P"
            for cod in sorted(apl_cod):
                r, serie = CODIGOS[cod], apl_cod[cod]
                val = acum(serie, int(s), m, dados["du"])
                if not val:
                    continue
                out.append([f"{s}/{int(s) + 1}", m, rp, r["fonte"], cod, tipo[r["tipo"]], r["titulo"],
                            round(serie.get(m, 0.0), 2), round(val, 2), "", ""])
                if r["ponderador"]:
                    out.append([f"{s}/{int(s) + 1}", m, rp, r["fonte"], r["ponderador"], "calculado",
                                f"Ponderação de {cod}: {round((float(r['fator']) - 1) * 100, 1)}% do saldo".replace(".", ","),
                                "", round(val * (float(r["fator"]) - 1), 2), "", ""])
            for an, linhas in LINHAS.items():
                for k, cod, desc in linhas:
                    out.append([f"{s}/{int(s) + 1}", m, rp, an, cod,
                                "informado" if cod in LINHAS_INFORMADAS else "calculado", desc, "",
                                round(v[an].get(k, 0.0), 2), "", ""])
    return out


def cmd_rodar(pasta, data_base, tipo_if="banco", rodada=None):
    dados, res, avisos, classif, apl_cod = montar(pasta, data_base, tipo_if)
    rodada = rodada or f"{datetime.now():%Y-%m-%d_%H%M%S}"
    destino = SAIDAS / "rodadas" / rodada  # nunca sobrescreve rodada anterior
    if destino.exists():
        sys.exit(f"rodada {rodada} já existe em {destino}; use outro --rodada")
    resultado = calcular(dados, dados["params"])
    dados["rodada"], dados["avisos"], dados["python"] = rodada, avisos, resultado
    dados["python_oficial"] = oficial = calcular(dados, dados["params"], dados["apl_oficial"])
    gravar(destino / "saldo_operacoes.csv", ["id", "mes", "fonte", "codigo", "saldo_medio"],
           [[i, m, CODIGOS[c]["fonte"] if c else "", c, round(s, 2)] for (i, m), (c, s) in sorted(res.items())])
    gravar(destino / "classificacao.csv", ["id", "fonte", "codigo", "tipo", "motivo", "titulo"],
           [[i, fo, c, {"i": "informado", "c": "calculado"}.get(t, ""), mo, CODIGOS[c]["titulo"] if c else ""]
            for i, fo, c, t, mo in classif])
    cods = com_oficial(linhas_codigos(dados, resultado, apl_cod), dados["oficial"], dados)
    gravar(destino / "codigos.csv", CAB_CODIGOS, cods)
    if dados["oficial"]:
        dif = [l for l in cods if l[10] != "" and abs(l[10]) > 1]
        print(f"Oficial: {len(dif)} códigos com diferença acima de R$ 1,00 (ver codigos.csv)")
    gravar(destino / "aplicacoes.csv", ["fonte_categoria", "mes", "saldo_medio_ponderado"],
           [[k, m, round(v, 2)] for k, porm in sorted(dados["apl"].items()) for m, v in sorted(porm.items())])
    gravar(destino / "exigibilidade.csv", ["ano_safra", "posicao", "tipo", "anexo", "codigo", "item", "valor"],
           [[f"{s}/{int(s) + 1}", m, "R" if m < dados["mes_base"] else "P", an, cod, desc, round(v[an][k], 2)]
            for s, porm in resultado.items() for m, v in porm.items() for an in LINHAS for k, cod, desc in LINHAS[an]])
    (destino / "operacional.html").write_text(html_operacional(dados), encoding="utf-8")
    (destino / "diretores.html").write_text(html_diretores(dados), encoding="utf-8")
    for a in avisos:
        print("AVISO:", a)
    for s, porm in oficial.items():
        jun = porm[f"{int(s) + 1}-06"]
        print(f"\n{s}/{int(s) + 1} - fechamento de junho, como os diretores veem (excesso + / deficiência -):")
        for rot, an, exc, dfc in RESUMO:
            print(f"  {rot:32} {br_mil(jun[an][exc] - jun[an][dfc])}")
    print(f"\nRelatórios: {destino / 'operacional.html'} e {destino / 'diretores.html'}")
    return destino




# ---------------------------------------------------------------- comparação de rodadas

def comparar(a, b, tol=0.01):
    """[(id, mes, motivo, fonte_a, cat_a, saldo_a, fonte_b, cat_b, saldo_b, diferenca)], maior impacto primeiro."""
    vazio = ("", "", 0.0)
    linhas = []
    for k in a.keys() | b.keys():
        fa, ca, sa = a.get(k, vazio)
        fb, cb, sb = b.get(k, vazio)
        if k not in a:
            motivo = "so_na_B"
        elif k not in b:
            motivo = "so_na_A"
        elif (fa, ca) != (fb, cb):
            motivo = "mudou_classificacao"
        elif abs(sb - sa) > tol:
            motivo = "mudou_saldo"
        else:
            continue
        linhas.append((*k, motivo, fa, ca, sa, fb, cb, sb, sb - sa))
    return sorted(linhas, key=lambda l: -max(abs(l[9]), l[5], l[8]) if l[2] == "mudou_classificacao" else -abs(l[9]))


def carregar_rodada(caminho, mes=None):
    caminho = Path(caminho)
    if caminho.is_dir():
        caminho = caminho / "saldo_operacoes.csv"
    return {(r["id"].strip(), r["mes"].strip()): (r["fonte"].strip().upper(), r["codigo"].strip(), num(r["saldo_medio"]))
            for r in ler(caminho) if not mes or r["mes"].strip() == mes}


def cmd_comparar(cam_a, cam_b, mes=None, saida=None):
    linhas = comparar(carregar_rodada(cam_a, mes), carregar_rodada(cam_b, mes))
    saida = gravar(saida or SAIDAS / "diferencas.csv",
                   ["id", "mes", "motivo", "fonte_A", "codigo_A", "saldo_A", "fonte_B", "codigo_B", "saldo_B",
                    "diferenca"], linhas)
    imp = defaultdict(float)  # impacto por fonte/código = soma B - soma A
    for l in linhas:
        if l[3]:
            imp[(l[3], l[4])] -= l[5]
        if l[6]:
            imp[(l[6], l[7])] += l[8]
    print(f"{len(linhas)} diferenças operação/mês -> {saida}\nImpacto B - A por fonte/código:")
    for (fo, ca), v in sorted(imp.items(), key=lambda x: -abs(x[1])):
        print(f"  {fo:5} {ca:16} {br_mil(v)}")
    return linhas


# ---------------------------------------------------------------- estrutura da projeção atual (sem valores)

def _mascarar(s, palavras):
    for p in palavras:
        s = re.sub(re.escape(p), "<mascarado>", s, flags=re.I)
    return s


def _estrutura_excel(caminho, palavras, linhas_cab=2, max_linhas=60):
    try:
        import openpyxl
    except ImportError:
        return "(openpyxl não instalado: pip install openpyxl)\n"
    wb = openpyxl.load_workbook(caminho, data_only=False, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append(f"\n### Aba `{_mascarar(ws.title, palavras)}` ({ws.max_row} linhas x {ws.max_column} colunas)\n")
        vistos = set()  # mesma fórmula na mesma coluna aparece uma vez só
        for r, row in enumerate(ws.iter_rows(max_row=max_linhas), start=1):
            cels = []
            for c in row:
                v = getattr(c, "value", None)
                if v is None:
                    continue
                if isinstance(v, str) and v.startswith("="):
                    chave = (c.column_letter, re.sub(r"\d+", "#", v))
                    if chave not in vistos:
                        vistos.add(chave)
                        cels.append(f"{c.coordinate} `{_mascarar(v, palavras)}`")
                elif isinstance(v, str):  # texto abaixo do cabeçalho pode ser nome de cliente
                    cels.append(f"{c.coordinate} {_mascarar(v, palavras) if r <= linhas_cab else '<txt>'}")
                elif isinstance(v, (int, float)):
                    cels.append(f"{c.coordinate} <n>")
                else:
                    cels.append(f"{c.coordinate} <{type(v).__name__}>")
            if cels:
                out.append(f"- L{r}: " + " | ".join(cels))
    return "\n".join(out) + "\n"


def _estrutura_codigo(caminho, palavras):
    txt = Path(caminho).read_text(encoding="utf-8", errors="replace")
    txt = re.sub(r"'[^'\n]*'", "'<str>'", txt)
    txt = re.sub(r'"[^"\n]*"', '"<str>"', txt)
    return "```\n" + _mascarar(txt, palavras) + "\n```\n"


def cmd_estrutura(caminhos, mascarar="", saida=None):
    palavras = [w.strip() for w in mascarar.split(",") if w.strip()]
    tipos = {".xlsx": _estrutura_excel, ".xlsm": _estrutura_excel,
             **{e: _estrutura_codigo for e in (".sql", ".py", ".r", ".sas", ".sps", ".bas", ".vba", ".txt")}}
    arquivos = []
    for c in map(Path, caminhos):
        arquivos += sorted(f for f in (c.rglob("*") if c.is_dir() else [c]) if f.suffix.lower() in tipos)
    if not arquivos:
        sys.exit("nenhum Excel (.xlsx/.xlsm) ou código (.sql, .py, .txt etc.) encontrado")
    partes = ["# Estrutura da projeção atual (sem valores)\n"]
    for f in arquivos:
        partes.append(f"\n## {f.suffix.lower()} `{_mascarar(f.name, palavras)}`\n")
        partes.append(tipos[f.suffix.lower()](f, palavras))
    saida = Path(saida or SAIDAS / "estrutura_projecao.md")
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text("".join(partes), encoding="utf-8")
    print(f"{len(arquivos)} arquivos -> {saida}\nLEIA ANTES DE SUBIR OU COLAR: a máscara é automática e pode deixar passar algo.")
    return saida


# ---------------------------------------------------------------- relatório HTML (visão dos diretores + gestores)

CSS = r"""
:root{--bg:#f7f7f5;--card:#fff;--tx:#1d1d1b;--mut:#6b6b66;--ln:#e2e1dc;--pos:#1f7a4d;--neg:#b3261e;--acc:#2f5d8a;--posbg:#e6f2ea;--negbg:#fbe9e7;--pa:#efeeea;--at:#fbf1cf;--fu:#e8f0f8}
@media (prefers-color-scheme:dark){:root{--bg:#161615;--card:#20201e;--tx:#ecebe6;--mut:#a3a29c;--ln:#3a3936;--pos:#6cc799;--neg:#f2867c;--acc:#8db4dd;--posbg:#1d3328;--negbg:#3a2220;--pa:#2a2a27;--at:#3a3420;--fu:#1f2a36}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
header{padding:20px 24px 8px}h1{font-size:20px;margin:0 0 4px}h2{font-size:16px;margin:24px 0 8px}h3{font-size:14px;margin:16px 0 6px;color:var(--mut)}
.meta{color:var(--mut);font-size:13px}.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px;margin-left:8px}
.ok{background:var(--posbg);color:var(--pos)}.falhou{background:var(--negbg);color:var(--neg)}
nav{display:flex;gap:4px;padding:0 24px;border-bottom:1px solid var(--ln);flex-wrap:wrap}nav button{background:none;border:0;border-bottom:2px solid transparent;padding:10px 12px;color:var(--mut);font:inherit;cursor:pointer}
nav button.ativo{color:var(--tx);border-color:var(--acc)}main{padding:8px 24px 40px}section{display:none}section.ativo{display:block}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px}.card{background:var(--card);border:1px solid var(--ln);border-radius:8px;padding:12px}
.card .r{color:var(--mut);font-size:12px}.card .v{font-size:20px;font-weight:600;margin-top:4px;font-variant-numeric:tabular-nums}
.tab{overflow-x:auto;background:var(--card);border:1px solid var(--ln);border-radius:8px}table{border-collapse:collapse;width:max-content;min-width:100%}
th,td{padding:5px 8px;border-bottom:1px solid var(--ln);text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}th:first-child,td:first-child{text-align:left;position:sticky;left:0;background:var(--card)}
th{font-weight:600;color:var(--mut);font-size:12px}.pos{color:var(--pos)}.neg{color:var(--neg)}.cod{color:var(--mut);font-size:12px;margin-right:6px}
.pa{background:var(--pa)}.at{background:var(--at)}.fu{background:var(--fu)}th.jun,td.jun{font-weight:700}
tr.conta td{font-weight:600;padding-top:12px}tr.sub td{color:var(--mut);font-size:12px}tr.sub td:first-child{padding-left:22px}
.leg span{display:inline-block;padding:1px 8px;border-radius:4px;margin-right:6px}
input{width:110px;padding:3px 4px;border:1px solid var(--ln);border-radius:4px;background:var(--bg);color:var(--tx);font:inherit;text-align:right}
select{padding:4px;border:1px solid var(--ln);border-radius:4px;background:var(--card);color:var(--tx);font:inherit}
.acoes{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0;align-items:center}.acoes button,.acoes label{padding:6px 12px;border:1px solid var(--ln);border-radius:6px;background:var(--card);color:var(--tx);font:inherit;cursor:pointer}
.avisos{background:var(--negbg);color:var(--neg);border-radius:8px;padding:8px 12px;margin:8px 0;font-size:13px}.nota{color:var(--mut);font-size:12px}
@media (max-width:600px){header,main{padding-left:16px;padding-right:16px}nav{padding:0 16px}}
"""

# Cópia exata em JavaScript de pct_de, com_parametros, dir_final, acum, aplicar e sisex (ADR 0003).
MOTOR_JS = r"""
"use strict";
const D = JSON.parse(document.getElementById("dados").textContent);
const GR = {RO_PRONAF:["RO","pronaf"],RO_PRONAMP:["RO","pronamp"],RO_DEMAIS:["RO","demais"],POUP:["POUP","credito_rural"],LCA:["LCA","credito_rural"]};
function mesesSafra(s){const r=[];let a=s,m=7;for(let i=0;i<12;i++){r.push(a+"-"+String(m).padStart(2,"0"));if(m===12){a++;m=1}else m++}return r}
function g(o,...ks){for(const k of ks){if(o==null)return undefined;o=o[k]}return o}
function pctDe(P,gr,m){return g(P,"pct",gr,m)??g(P,"pct_grupo",gr)??P.pct_geral??80}
function comParametros(apl,pagos,P,M,mesBase){
  const out={};for(const k in apl)out[k]=Object.assign({},apl[k]);
  for(const gr in GR){const [fo,ca]=GR[gr];const prazo=Math.max(1,parseInt(g(P,"prazo",gr)??12));const key=fo+"|"+ca;out[key]=out[key]||{};const alvo=out[key];
    M.forEach((mj,j)=>{if(mj<mesBase)return;const lib=g(P,"liberacao",gr,mj)||0;const pct=pctDe(P,gr,mj);const u=(1-pct/100)*(g(pagos,gr,mj)||0);
      if(!lib&&!u)return;
      for(let k=j;k<M.length;k++){const dl=k<=j+prazo?lib*((k===j||k===j+prazo)?0.5:1):0;alvo[M[k]]=(alvo[M[k]]||0)+dl+u*(k===j?0.5:1)}})}
  return out}
function dirFinal(dirs,P){const out={};for(const mod of D.dir_modalidades){out[mod]={};for(const m in (dirs[mod]||{}))out[mod][m]=dirs[mod][m].slice()}
  for(const mod in (P.dir||{}))for(const m in P.dir[mod]){const [c,a]=P.dir[mod][m];(out[mod]=out[mod]||{})[m]=[c,a]}return out}
function acum(serie,s,m){let t=0,p=0;for(const k of mesesSafra(s)){if(k>m)break;const w=D.du[k]||0;t+=(serie[k]||0)*w;p+=w}return p?t/p:0}
function aplicar(an,d,c){for(const k of Object.values(D.balde[an]))d[k]=0;for(const v of Object.values(D.limites))if(v[0]===an)d[v[1]]=0;
  for(const cat in D.balde[an])d[D.balde[an][cat]]+=c(an+"|"+cat);
  for(const [cod,[a,b,base,pct,exc]] of Object.entries(D.limites))if(a===an){const v=c(an+"|"+cod);d["lim_"+cod]=Math.min(v,pct*d[base]);d[b]+=d["lim_"+cod];
    if(exc){d["exc_"+cod]=v-d["lim_"+cod];d[exc]+=d["exc_"+cod]}}}
function sisex(s,m,apl,dirs,base){
  const c=k=>acum(apl[k]||{},s,m);const dc=(mod,i)=>{const sr={};for(const k in (dirs[mod]||{}))sr[k]=dirs[mod][k][i];return acum(sr,s,m)};
  const p=base.ro_propria;const ro={vsr_media:base.vsr_avista,propria:p,capt_dir_geral:dc("geral",0),capt_dir_pronaf:dc("pronaf",0),capt_dir_pronamp:dc("pronamp",0)};
  ro.exig_total=p+ro.capt_dir_geral+ro.capt_dir_pronaf+ro.capt_dir_pronamp;ro.sub_pronaf=0.35*p+ro.capt_dir_pronaf;ro.sub_pronamp=0.50*p+ro.capt_dir_pronamp;
  ro.geral=p-0.35*p-0.50*p+ro.capt_dir_geral;
  aplicar("RO",ro,c);ro.ap_pronaf+=dc("pronaf",1);ro.ap_pronamp+=dc("pronamp",1);ro.ap_geral+=dc("geral",1);
  ro.ap_total=ro.ap_pronaf+ro.ap_geral+ro.ap_pronamp;
  ro.def_pronaf=Math.max(0,ro.sub_pronaf-ro.ap_pronaf);ro.exc_pronaf=Math.max(0,ro.ap_pronaf-ro.sub_pronaf);
  ro.def_pronamp=Math.max(0,ro.sub_pronamp-ro.ap_pronamp);ro.exc_pronamp=Math.max(0,ro.ap_pronamp-ro.sub_pronamp);
  ro.def_geral=Math.max(0,ro.exig_total-(ro.ap_total+ro.def_pronaf+ro.def_pronamp));ro.exc_geral=Math.max(0,ro.ap_geral-ro.geral);
  ro.def_total=ro.def_geral+ro.def_pronaf+ro.def_pronamp;ro.exc_total=Math.max(0,ro.exc_pronaf+ro.exc_pronamp+ro.ap_geral-ro.geral);
  const pp=base.poup_propria;const po={vsr_media:base.vsr_poupanca,propria:pp,capt_dir:dc("poup",0)};
  po.exig_total=pp+po.capt_dir;po.sub=0.95*pp+po.capt_dir;po.cpr_limite=0.05*pp;po.exig_liquida=po.exig_total-dc("poup",1);aplicar("POUP",po,c);po.ap_cr+=dc("poup",1);
  po.ap_total=po.ap_cr+po.ap_cpr;po.def_sub=Math.max(0,po.sub-po.ap_cr);po.exc_sub=Math.max(0,po.ap_cr-po.sub);
  po.def_total=Math.max(0,po.exig_total-po.ap_total);po.exc_total=Math.max(0,po.ap_total-po.exig_total);
  const lp=base.lca_propria[m]||0;const lc={lca_media:base.lca_media[m]||0,propria:lp,capt_dir_cr:dc("lca_cr",0),capt_dir_livre:dc("lca_livre",0)};
  lc.exig_total=lp+lc.capt_dir_cr+lc.capt_dir_livre;lc.sub=0.45*lp+lc.capt_dir_cr;lc.faculdade=0.55*lp+lc.capt_dir_livre;
  lc.exc_ro=m.slice(5)==="06"?ro.exc_total:0;aplicar("LCA",lc,c);lc.ap_cr+=dc("lca_cr",1)+lc.exc_ro;
  lc.ap_fac=Math.min(lc.ap_fac+dc("lca_livre",1),lc.faculdade);lc.ap_total=lc.ap_cr+lc.ap_fac;
  lc.def_sub=Math.max(0,lc.sub-lc.ap_cr);lc.exc_sub=Math.max(0,lc.ap_cr-lc.sub);lc.def_total=Math.max(0,lc.exig_total-lc.ap_total);lc.exc_total=Math.max(0,lc.ap_total-lc.exig_total);
  return {RO:ro,POUP:po,LCA:lc}}
function calcular(P,apl=D.apl){const a=comParametros(apl,D.pagos,P,D.meses,D.mes_base);const dirs=dirFinal(D.dir,P);const out={};
  for(const s of D.safras){out[s]={};for(const m of mesesSafra(s))out[s][m]=sisex(s,m,a,dirs,D.bases[s])}return out}
// ---- tela (comum aos dois relatórios)
const mi=v=>(v/1e6).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1});
const reais=v=>v===undefined||v===""?"":(+v).toLocaleString("pt-BR",{maximumFractionDigits:2});
const sinal=(v,lim)=>v>lim?"pos":v<-lim?"neg":"";
const rot=m=>["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"][+m.slice(5)-1]+"/"+m.slice(2,4);
const periodo=m=>(m<D.mes_base?"pa":m===D.mes_base?"at":"fu")+(m.slice(5)==="06"?" jun":"");
const cabecalho=(s,primeira="")=>`<tr><th>${primeira}</th>`+mesesSafra(s).map(m=>`<th class="${periodo(m)}">${rot(m)}</th>`).join("")+"</tr>";
const dataBR=d=>d.split("-").reverse().join("/");
"""

DIRETORES_MODELO = r"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Exigibilidade - diretores</title><style>__CSS__</style></head><body>
<header><h1>Exigibilidade do Crédito Rural</h1><div class="meta" id="meta"></div></header>
<main id="corpo"></main>
<script type="application/json" id="dados">__DADOS__</script>
<script>__MOTOR__
const R=calcular(D.params||{});let h=`<p class="nota leg"><span class="pa">meses passados: valor oficial (Sisex)</span><span class="at">mês da data-base</span><span class="fu">projeção, partindo do oficial</span></p>`;
if(D.sem_oficial.length)h+=`<div class="avisos">Sem valor oficial em ${D.sem_oficial.map(rot).join(", ")}: nesses meses foi usado o valor calculado pelo programa.</div>`;
for(const s of D.safras){const jun=(s+1)+"-06";
  h+=`<h2>Ano-safra ${s}/${s+1}</h2><h3>Fechamento de junho/${String(s+1).slice(2)}: excesso (+) / deficiência (−), R$ milhões</h3><div class="cards">`;
  for(const [r,an,e,f] of D.resumo){const v=R[s][jun][an][e]-R[s][jun][an][f];h+=`<div class="card"><div class="r">${r}</div><div class="v ${sinal(v,5e4)}">${mi(v)}</div></div>`}
  h+=`</div><h3>Excesso (+) / deficiência (−) até o mês, R$ milhões</h3><div class="tab"><table>${cabecalho(s)}`;
  for(const [r,an,e,f] of D.resumo)h+=`<tr><td>${r}</td>`+mesesSafra(s).map(m=>{const v=R[s][m][an][e]-R[s][m][an][f];return `<td class="${periodo(m)} ${sinal(v,5e4)}">${mi(v)}</td>`}).join("")+"</tr>";
  h+="</table></div>"}
document.getElementById("corpo").innerHTML=h+`<p class="nota">Posição de cada mês = média acumulada desde julho, como no Sisex (MCR, Documento 6). Parâmetros da gerência: ${D.origem}.</p>`;
document.getElementById("meta").textContent=`Data-base ${dataBR(D.data_base)} · rodada ${D.rodada} · v${D.versao}`;
</script></body></html>"""

OPERACIONAL_MODELO = r"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Exigibilidade - operacional</title><style>__CSS__</style></head><body>
<header><h1>Exigibilidade do Crédito Rural · operacional</h1><div class="meta" id="meta"></div>
<div class="acoes"><label>Ano-safra <select id="safra"></select></label><button id="gerar">Gerar relatório dos diretores com estes parâmetros</button></div>
<p class="nota leg"><span class="pa">passado: calculado pelo programa</span><span class="at">mês da data-base</span><span class="fu">projeção</span> Valores em reais. Onde há valor oficial, as linhas "oficial" e "diferença" (oficial − calculado) aparecem embaixo.</p></header>
<nav><button class="ativo" data-s="blocos">Por exigibilidade</button><button data-s="detalhe">Detalhe por código</button><button data-s="param">Parâmetros e simulação</button></nav>
<main>
<section id="blocos" class="ativo"></section>
<section id="detalhe"></section>
<section id="param">
<p class="nota">Campos em reais, valor exato. Liberação e % de reembolso por grupo e mês; a liberação nova é paga em parcela única depois do número de meses informado. DIR: saldo médio do mês (captação e aplicação). Tudo recalcula na hora.</p>
<div class="acoes"><button id="salvar">Salvar parâmetros (parametros.json)</button><label>Carregar parâmetros <input type="file" id="carregar" accept=".json" hidden></label><button id="restaurar">Voltar aos parâmetros da rodada</button></div>
<div id="campos"></div>
</section>
</main>
<script type="application/json" id="dados">__DADOS__</script>
<script type="application/json" id="modelo-diretores">__MODELO_DIRETORES__</script>
<script>__MOTOR__
const NOME_GR = {RO_PRONAF:"Obrigatórios Pronaf",RO_PRONAMP:"Obrigatórios Pronamp",RO_DEMAIS:"Obrigatórios Geral",POUP:"Poupança",LCA:"LCA"};
const NOME_DIR = {geral:"DIR-Geral",pronaf:"DIR-Pronaf",pronamp:"DIR-Pronamp",poup:"DIR-Poup",lca_cr:"DIR-LCA-CR",lca_livre:"DIR-LCA-Livre"};
const LIN = {};for(const an in D.linhas){LIN[an]={};for(const [k,cod,desc] of D.linhas[an])LIN[an][k]=[cod,desc]}
let P = JSON.parse(JSON.stringify(D.params || {}));
function conferir(){const igual=(py,js)=>{for(const s in py)for(const m in py[s])for(const an in py[s][m])for(const k in py[s][m][an])
    if(Math.abs(py[s][m][an][k]-js[s][m][an][k])>0.01)return false;return true};
  return igual(D.python,calcular(D.params||{}))&&igual(D.python_oficial,calcular(D.params||{},D.apl_oficial))}
const safraSel=()=>+document.getElementById("safra").value;
const r0=v=>(Math.round(v)+0).toLocaleString("pt-BR");  // +0 tira o "-0"
function linhas(s,an,k,rotulo,cod,val){  // linha calculada + oficial e diferença onde há oficial
  let h=`<tr><td>${rotulo}</td>`+mesesSafra(s).map(m=>`<td class="${periodo(m)}">${r0(val(m))}</td>`).join("")+"</tr>";
  if(cod&&mesesSafra(s).some(m=>D.oficial[m])){
    h+=`<tr class="sub"><td>oficial</td>`+mesesSafra(s).map(m=>`<td class="${periodo(m)}">${D.oficial[m]?r0(D.oficial[m][cod]||0):""}</td>`).join("")+"</tr>";
    h+=`<tr class="sub"><td>diferença</td>`+mesesSafra(s).map(m=>{if(!D.oficial[m])return `<td class="${periodo(m)}"></td>`;const d=(D.oficial[m][cod]||0)-val(m);return `<td class="${periodo(m)} ${sinal(d,1)}">${r0(d)}</td>`}).join("")+"</tr>"}
  return h}
function render(){const R=calcular(P),s=safraSel(),apl=comParametros(D.apl,D.pagos,P,D.meses,D.mes_base),dirs=dirFinal(D.dir,P);let b="",d="";
  for(const [nome,an,rows] of D.blocos){b+=`<h2>${nome} · ${s}/${s+1}</h2><div class="tab"><table>${cabecalho(s)}`;
    for(const r of rows){const [k,rotulo]=typeof r==="string"?[r,LIN[an][r][1]]:r;
      if(k==="#"){b+=`<tr class="conta"><td colspan="13">${rotulo}</td></tr>`;continue}
      if(k.startsWith("mes:")){const bs=k.slice(4).split("+");const cats=Object.keys(D.balde[an]).filter(c=>bs.includes(D.balde[an][c]))
          .concat(Object.keys(D.limites).filter(c=>D.limites[c][0]===an&&bs.includes(D.limites[c][1])));
        b+=linhas(s,an,k,rotulo,"",m=>cats.reduce((t,c)=>t+(g(apl,an+"|"+c,m)||0),0));continue}
      if(k.startsWith("dir:")){const mods=k.slice(4).split("+");
        b+=linhas(s,an,k,rotulo,"",m=>mods.reduce((t,mod)=>{const sr={};for(const x in (dirs[mod]||{}))sr[x]=dirs[mod][x][1];return t+acum(sr,s,m)},0));continue}
      b+=linhas(s,an,k,rotulo,LIN[an][k][0],m=>R[s][m][an][k]??0)}
    b+="</table></div>"}
  for(const an of ["RO","POUP","LCA"]){const nome={RO:"Anexo II - Obrigatórios",POUP:"Anexo III - Poupança Rural",LCA:"Anexo IV - LCA"}[an];
    d+=`<h2>${nome} · ${s}/${s+1}</h2><div class="tab"><table>${cabecalho(s)}`;
    for(const [k,cod,desc] of D.linhas[an])d+=linhas(s,an,k,`<span class="cod">${cod}</span>${desc}`,cod,m=>R[s][m][an][k]??0);
    d+="</table></div>"}
  document.getElementById("blocos").innerHTML=b;document.getElementById("detalhe").innerHTML=d+'<p class="nota">Códigos de aplicação um a um: codigos.csv da rodada.</p>'}
function campo(caminho,valor){return `<input data-c="${caminho}" value="${reais(valor)}">`}
function atualizarPct(){document.querySelectorAll('#campos input[data-c^="pct"]').forEach(e=>{const [t,x,m]=e.dataset.c.split("|");
  if(t==="pct")e.value=reais(pctDe(P,x,m));else if(t==="pct_grupo")e.value=reais(g(P,"pct_grupo",x)??P.pct_geral??80)})}
function montarCampos(){const M=D.meses.filter(m=>m>=D.mes_base);const cab="<tr><th></th>"+M.map(m=>`<th>${rot(m)}</th>`).join("")+"</tr>";
  let h="<h2>Liberação nova é paga depois de quantos meses</h2><div class='tab'><table><tr>"+D.grupos.map(x=>`<th>${NOME_GR[x]}</th>`).join("")+"</tr><tr>"+D.grupos.map(x=>`<td>${campo("prazo|"+x,g(P,"prazo",x)??12)}</td>`).join("")+"</tr></table></div>";
  h+="<h2>Liberação da gerência (R$ no mês, valor exato)</h2><div class='tab'><table>"+cab+D.grupos.map(x=>`<tr><td>${NOME_GR[x]}</td>`+M.map(m=>`<td>${campo("liberacao|"+x+"|"+m,g(P,"liberacao",x,m))}</td>`).join("")+"</tr>").join("")+"</table></div>";
  h+="<h2>% de reembolso das parcelas (100 = paga tudo)</h2><p class='nota'>O geral preenche todos os grupos e meses; o do grupo preenche os meses do grupo; cada mês pode ser mudado sozinho.</p><div class='tab'><table><tr><th>Geral</th><td>"+campo("pct_geral",P.pct_geral??80)+"</td></tr></table><table>"
    +"<tr><th></th><th>Grupo</th>"+M.map(m=>`<th>${rot(m)}</th>`).join("")+"</tr>"+D.grupos.map(x=>`<tr><td>${NOME_GR[x]}</td><td>${campo("pct_grupo|"+x,g(P,"pct_grupo",x)??P.pct_geral??80)}</td>`+M.map(m=>`<td>${campo("pct|"+x+"|"+m,pctDe(P,x,m))}</td>`).join("")+"</tr>").join("")+"</table></div>";
  h+="<h2>DIR (R$, saldo médio do mês, valor exato): captação / aplicação</h2><div class='tab'><table>"+cab;
  for(const mod of D.dir_modalidades)for(const [i,nm] of [[0,"captação"],[1,"aplicação"]])
    h+=`<tr><td>${NOME_DIR[mod]} ${nm}</td>`+M.map(m=>{const v=g(P,"dir",mod,m)?P.dir[mod][m][i]:(g(D.dir,mod,m)?D.dir[mod][m][i]:0);return `<td>${campo("dir|"+mod+"|"+m+"|"+i,v)}</td>`}).join("")+"</tr>";
  document.getElementById("campos").innerHTML=h+"</table></div>"}
document.getElementById("campos").addEventListener("input",e=>{const c=e.target.dataset.c;if(!c)return;const v=parseFloat(e.target.value.replace(/\./g,"").replace(",","."));const p=c.split("|");
  if(p[0]==="prazo"){(P.prazo=P.prazo||{})[p[1]]=isNaN(v)?12:v}
  else if(p[0]==="dir"){const [,mod,m,i]=p;P.dir=P.dir||{};P.dir[mod]=P.dir[mod]||{};const atual=P.dir[mod][m]||(g(D.dir,mod,m)||[0,0]).slice();atual[+i]=isNaN(v)?0:v;P.dir[mod][m]=atual}
  else if(p[0]==="pct_geral"){if(isNaN(v))delete P.pct_geral;else P.pct_geral=v;delete P.pct_grupo;delete P.pct;atualizarPct()}
  else if(p[0]==="pct_grupo"){P.pct_grupo=P.pct_grupo||{};if(isNaN(v))delete P.pct_grupo[p[1]];else P.pct_grupo[p[1]]=v;if(P.pct)delete P.pct[p[1]];atualizarPct()}
  else{const [t,x,m]=p;P[t]=P[t]||{};P[t][x]=P[t][x]||{};if(isNaN(v))delete P[t][x][m];else P[t][x][m]=v}
  render()});
const baixar=(nome,txt,tipo)=>{const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([txt],{type:tipo}));a.download=nome;a.click()};
document.getElementById("salvar").onclick=()=>baixar("parametros.json",JSON.stringify(P,null,1),"application/json");
document.getElementById("gerar").onclick=()=>{const dd=Object.fromEntries(D.chaves_diretores.map(k=>[k,D[k]]));
  Object.assign(dd,{apl:D.apl_oficial,params:P,origem:`simulação da tela de ${new Date().toLocaleString("pt-BR")}`});
  const tpl=JSON.parse(document.getElementById("modelo-diretores").textContent);
  baixar("diretores.html",tpl.replace("__DA"+"DOS__",()=>JSON.stringify(dd).replace(/<\//g,"<\\/")),"text/html")};
document.getElementById("carregar").onchange=e=>{const f=e.target.files[0];if(!f)return;f.text().then(t=>{P=JSON.parse(t);montarCampos();render()})};
document.getElementById("restaurar").onclick=()=>{P=JSON.parse(JSON.stringify(D.params||{}));montarCampos();render()};
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>{document.querySelectorAll("nav button,section").forEach(x=>x.classList.remove("ativo"));b.classList.add("ativo");document.getElementById(b.dataset.s).classList.add("ativo")});
const sel=document.getElementById("safra");sel.innerHTML=D.safras.map(s=>`<option value="${s}">${s}/${s+1}</option>`).join("");sel.onchange=render;
const ok=conferir();
document.getElementById("meta").innerHTML=`Data-base ${dataBR(D.data_base)} · rodada ${D.rodada} · v${D.versao}<span class="badge ${ok?"ok":"falhou"}">CHECK simulação = Python: ${ok?"OK":"FALHOU"}</span>`
  +(D.avisos.length?`<div class="avisos">${D.avisos.map(a=>"AVISO: "+a).join("<br>")}</div>`:"");
montarCampos();render();
</script></body></html>"""

_json = lambda x: json.dumps(x, ensure_ascii=False).replace("</", "<\\/")
DIRETORES_HTML = DIRETORES_MODELO.replace("__CSS__", CSS).replace("__MOTOR__", MOTOR_JS)
OPERACIONAL_HTML = OPERACIONAL_MODELO.replace("__CSS__", CSS).replace("__MOTOR__", MOTOR_JS)


def dados_diretores(dados, params, origem):
    """O que o relatório dos diretores recebe: a série oficial (oficial no passado, projeção no futuro) e os parâmetros."""
    return dict({k: dados[k] for k in CHAVES_DIRETORES}, apl=dados["apl_oficial"], params=params, origem=origem)


def html_diretores(dados):
    return DIRETORES_HTML.replace("__DADOS__", _json(dados_diretores(dados, dados["params"], "parametros.json da rodada")))


def html_operacional(dados):
    return OPERACIONAL_HTML.replace("__DADOS__", _json(dados)).replace("__MODELO_DIRETORES__", _json(DIRETORES_HTML))


# ---------------------------------------------------------------- exemplo e autoteste (dados inventados)

def cmd_exemplo(data_base="2026-11-02"):
    ex = PASTA_SCRIPT / "exemplo"
    ex.mkdir(exist_ok=True)
    # fonte, programa, finalidade, recursos, beneficiario, linha, data de contratação, taxa, tamanho
    tipos = [("RO", "pronaf", "custeio", "", "", "", "2026-08-10", 3.0, 30e6),
             ("RO", "pronaf", "custeio", "", "", "", "2025-09-15", 5.0, 20e6),
             ("RO", "pronaf", "comercializacao", "", "", "", "2026-09-01", 4.0, 2e6),
             ("RO", "pronamp", "custeio", "", "", "", "2026-08-05", 8.0, 45e6),
             ("RO", "pronamp", "investimento", "", "", "", "2025-10-20", 8.0, 2e6),
             ("RO", "pronamp", "custeio", "", "", "peq_medio", "2026-07-15", 10.0, 3e6),
             ("RO", "demais", "custeio", "", "", "", "2026-07-20", 10.0, 14e6),
             ("RO", "demais", "investimento", "", "", "abc", "2023-09-01", 8.5, 2e6),
             ("RO", "pronaf", "investimento", "", "", "", "2026-08-01", 3.0, 1e6),  # sem código no Doc 6
             ("POUP", "pronaf", "custeio", "controlados", "", "", "2026-08-20", 3.0, 20e6),
             ("POUP", "demais", "custeio", "nao_controlados", "", "", "2026-09-10", 11.0, 22e6),
             ("POUP", "", "", "", "", "cpr", "2026-10-01", 12.0, 2e6),
             ("LCA", "", "custeio", "controlados", "", "", "2026-08-12", 12.0, 42e6),
             ("LCA", "", "investimento", "nao_controlados", "", "", "2025-11-03", 13.0, 5e6),
             ("LCA", "", "", "", "", "cpr", "2026-10-15", 13.0, 2e6),
             ("LIVRE", "", "", "", "", "", "2026-08-01", 15.0, 2e6)]
    cab = ["id", "fonte", "programa", "finalidade", "recursos", "beneficiario", "linha", "data_contratacao", "taxa_aa",
           "codigo", "saldo"]
    ops, parc = [], []
    for k in range(64):
        fo, pr, fi, rc, be, li, dc, tx, tam = tipos[k % len(tipos)]
        s = tam * (1 + (k * 37 % 11) / 10)
        i = f"OP{k + 1:03d}"
        ops.append([i, fo, pr, fi, rc, be, li, dc, tx, "3.1.13.44-4" if k == 1 else "", s])  # OP002: código informado
        venc = meses("2027-02", 16)[k * 7 % 16]
        if fi == "investimento":  # parcelas anuais
            parc += [[i, f"{int(venc[:4]) + j}{venc[4:]}-10", s / 3 * 1.1] for j in range(3)]
        else:
            parc.append([i, f"{venc}-15", s * 1.08])
    gravar(ex / "operacoes.csv", cab, ops)
    gravar(ex / "liberacoes.csv", ["id", "data", "valor"] + cab[1:9],
           [["OP003", "2026-12-01", 2e6] + [""] * 8,
            ["NOVA01", "2027-01-11", 30e6, "RO", "pronamp", "custeio", "", "", "", "", 8.0],
            ["NOVA02", "2027-03-01", 25e6, "POUP", "demais", "custeio", "controlados", "", "", "", 11.0]])
    gravar(ex / "parcelas.csv", ["id", "data", "valor"], parc + [["NOVA01", "2027-12-15", 32e6], ["NOVA02", "2028-02-15", 27e6]])
    gravar(ex / "cenario.csv", ["tipo", "id", "data", "codigo"],
           [["liquidacao", "OP009", "2027-01-20", ""], ["reclassificacao", "OP005", "2027-02-10", "3.1.41.52-6"]])
    gravar(ex / "feriados.csv", ["data"], [["2026-11-20"], ["2026-12-25"], ["2027-01-01"], ["2027-04-21"], ["2027-12-25"]])
    # saldo diário de jul/26 até o dia útil antes da data-base, a partir da contratação, subindo até o saldo de hoje
    dias = [d for d in (date(2026, 7, 1) + timedelta(days=k) for k in range(124)) if d.weekday() < 5]
    gravar(ex / "saldos_diarios.csv", ["id", "data", "saldo"] + cab[1:10],
           [[o[0], d.isoformat(), round(o[-1] * (0.85 + 0.15 * j / len(dias)), 2)] + o[1:10]
            for o in ops for j, d in enumerate(dias) if d.isoformat() >= o[7]])
    gravar(ex / "vsr.csv", ["mes", "vsr_avista", "vsr_poupanca", "tipo"],
           [[m, 2.2e9 + j * 4e6, 2.4e8 + j * 1e6, "R" if m < "2026-11" else "P"] for j, m in enumerate(meses("2025-07", 36))])
    # LCA: saldo diário realizado (coluna data) e, para os meses projetados, uma linha por mês (coluna mes)
    lca = [[d.isoformat(), "", 6.5e8 + (d - date(2026, 6, 1)).days * 2.5e5] for d in
           (date(2026, 6, 1) + timedelta(days=k) for k in range(154)) if d.weekday() < 5]
    lca += [["", m, 6.9e8 + j * 1.2e7] for j, m in enumerate(meses("2026-11", 20))]
    gravar(ex / "lca.txt", ["data", "mes", "saldo"], lca)
    gravar(ex / "dir.csv", ["mes", "modalidade", "captacao", "aplicacao"],
           [[m, "pronaf", 0.0, 5e6] for m in meses("2026-07", 24)] + [[m, "geral", 3e6, 0.0] for m in meses("2026-07", 24)])
    (ex / "parametros.json").write_text(json.dumps({"prazo": {"RO_PRONAF": 10}, "liberacao": {"RO_PRONAF": {"2027-07": 20e6}},
                                                    "pct": {"RO_PRONAMP": {"2027-03": 60}}}, indent=1), encoding="utf-8")
    # oficial inventado: o calculado das posições passadas com uma diferença em um código e um código a menos
    dados, _, _, _, apl_cod = montar(ex, data_base)
    cods = linhas_codigos(dados, calcular(dados, dados["params"]), apl_cod)
    gravar(ex / "oficial.csv", ["posicao", "codigo", "valor"],
           [[l[1], l[4], round(l[8] * (1.03 if l[4] == "3.1.13.49-9" else 1), 2)] for l in cods
            if l[2] == "R" and l[4] != "3.2.12.18-7"])
    print(f"Entradas inventadas em {ex}\n")
    return cmd_rodar(ex, data_base, rodada=f"exemplo_{datetime.now():%Y-%m-%d_%H%M%S}")


def autoteste():
    checks = []

    def check(desc, ok):
        checks.append(f"CHECK {desc}: {'OK' if ok else 'FALHOU'}")

    perto = lambda a, b: abs(a - b) < 0.01

    # bases de cálculo
    vsr = {m: [(4.13e9, 1e9)] for m in meses("2025-07", 12)}
    lca = {m: [2e9 * 21, 21] for m in meses("2026-06", 5)}
    lca.update({"2026-10": [2.6e9 * 21, 21]})
    b = bases(2026, vsr, lca, "banco", [])
    check("RO própria = 31,5% x (média do VSR de jul/25 a jun/26 - 500 mi)", perto(b["ro_propria"], 3.63e9 * 0.315))
    check("cooperativa 2026/27 = 13%", perto(bases(2026, vsr, lca, "cooperativa", [])["ro_propria"], 3.63e9 * 0.13))
    check("Poupança própria = 70% do VSR da poupança", perto(b["poup_propria"], 7e8))
    check("LCA da posição de out/26 usa jun a set/26", perto(b["lca_media"]["2026-10"], 2e9))
    check("LCA da posição de nov/26 usa jun a out/26", perto(b["lca_media"]["2026-11"], 2.12e9))
    check("isenção: exigibilidade própria até 10 mi vira zero",
          bases(2026, {m: [(5.2e8, 0)] for m in meses("2025-07", 12)}, lca, "banco", [])["ro_propria"] == 0.0)

    # fórmulas do Sisex (anexo II): exigibilidade própria de 100, sem DIR
    du = {m: 20 for m in meses("2026-07", 12)}
    base = {"ro_propria": 100.0, "vsr_avista": 0, "poup_propria": 100.0, "vsr_poupanca": 0,
            "lca_propria": {m: 100.0 for m in du}, "lca_media": {}}
    serie = lambda v: {m: v for m in du}
    apl = {"RO|pronaf": serie(45.0), "RO|pronamp": serie(50.0), "RO|demais": serie(5.0)}
    r = sisex(2026, "2026-09", apl, {}, base, du)["RO"]
    check("excesso Pronaf cobre falta na Geral (Sisex 5.1.51)", perto(r["def_total"], 0) and perto(r["exc_total"], 0)
          and perto(r["def_geral"], 0) and perto(r["exc_pronaf"], 10))
    apl["RO|3.1.41.39-9"] = serie(20.0)
    r = sisex(2026, "2026-09", apl, {}, base, du)["RO"]
    check("investimento Pronamp limitado a 10% da sub", perto(r["ap_pronamp"], 55))
    r = sisex(2026, "2026-09", apl, {"geral": serie([30.0, 0.0]), "pronaf": serie([0.0, 8.0])}, base, du)["RO"]
    check("captação DIR soma na exigibilidade e aplicação DIR cumpre", perto(r["exig_total"], 130) and perto(r["ap_pronaf"], 53))
    po = sisex(2026, "2026-09", {"POUP|credito_rural": serie(90.0), "POUP|3.2.30.10-7": serie(20.0)},
               {"poup": serie([10.0, 0.0])}, base, du)["POUP"]
    check("Poupança: CPR até 5% da própria (3.2.30.10-7) e crédito rural mínimo de 95%",
          perto(po["ap_cpr"], 5) and perto(po["ap_total"], 95) and perto(po["def_sub"], 15))
    r = sisex(2026, "2026-09", {"RO|pronamp": serie(50.0), "RO|3.1.40.34-5": serie(20.0)}, {}, base, du)["RO"]
    check("peq./médios até 10% do Pronamp; o excesso vai para a Geral (3.1.40.34-5 e 3.1.30.78-8)",
          perto(r["ap_pronamp"], 55) and perto(r["ap_geral"], 15))
    r = sisex(2026, "2026-09", {"RO|3.1.12.09-8": serie(30.0)}, {}, base, du)["RO"]
    check("Res. 5.330 no Pronaf até 40% da sub (3.1.12.09-8)", perto(r["ap_pronaf"], 14))
    apl2 = {"RO|demais": serie(30.0), "RO|pronaf": serie(35.0), "RO|pronamp": serie(50.0)}
    du12 = {m: 20 for m in meses("2026-07", 12)}
    lc_mai = sisex(2026, "2027-05", apl2, {}, base, du12)["LCA"]
    lc_jun = sisex(2026, "2027-06", apl2, {}, base, du12)["LCA"]
    check("LCA: excesso de RO só conta na posição de junho", perto(lc_mai["exc_ro"], 0) and perto(lc_jun["exc_ro"], 15))
    check("média acumulada de julho até o mês, pelos dias úteis",
          perto(acum({"2026-07": 10.0, "2026-08": 40.0}, 2026, "2026-08", {"2026-07": 10, "2026-08": 30}), 32.5))

    # projeção
    op = lambda: {"A": {"codigo": "3.2.12.07-7", "saldo": 100.0, "taxa_aa": 0.0}}
    res, pagos = projetar(op(), [("A", date(2026, 12, 15), 100.0)], [], [], date(2026, 11, 2), date(2027, 1, 31))
    du_dez = dias_uteis("2026-12", set())
    uteis_ate_14 = sum(1 for k in range(1, 15) if date(2026, 12, k).weekday() < 5)
    check("saldo zero depois da liquidação (Doc 6, Anexo I, 4.6-b)",
          perto(res[("A", "2026-12")][1], 100.0 * uteis_ate_14 / du_dez) and ("A", "2027-01") not in res)
    check("pagamento entra no total pago do grupo", perto(pagos[("POUP", "2026-12")], 100.0))
    cen = [{"tipo": "reclassificacao", "id": "A", "data": date(2026, 12, 10), "codigo": "3.1.31.07-9"},
           {"tipo": "reclassificacao", "id": "A", "data": date(2027, 1, 20), "codigo": "3.3.30.68-1"}]
    res, _ = projetar(op(), [], [], cen, date(2026, 11, 2), date(2027, 2, 28))
    check("fonte do dia 15 vale para o mês inteiro (Doc 6, Anexo I, 12.2-b)",
          res[("A", "2026-12")][0] == "3.1.31.07-9" and res[("A", "2027-01")][0] == "3.1.31.07-9"
          and res[("A", "2027-02")][0] == "3.3.30.68-1")
    libs = [("N", date(2027, 1, 18), 50.0, {"codigo": "3.1.13.49-9", "taxa_aa": 0.0})]
    res, _ = projetar({}, [], libs, [], date(2026, 11, 2), date(2027, 2, 28))
    uteis_jan_apos = sum(1 for k in range(18, 32) if date(2027, 1, k).weekday() < 5)
    check("liberação futura: zero antes da liberação",
          perto(res[("N", "2027-01")][1], 50.0 * uteis_jan_apos / dias_uteis("2027-01", set()))
          and perto(res[("N", "2027-02")][1], 50.0))
    res, _ = projetar({"A": dict(op()["A"], taxa_aa=10.0)}, [], [], [], date(2026, 1, 1), date(2026, 12, 31))
    check("juros sobem o saldo cerca de 10% ao ano", 100 < res[("A", "2026-12")][1] < 110.1)

    # realizado a partir do saldo diário (Doc 6, Anexo I, 4.6-b e 12.2-b)
    cols = {"fonte": "RO", "finalidade": "custeio", "data_contratacao": "2026-08-01", "taxa_aa": "3"}
    dia = lambda d, prog, s=100.0: dict(cols, id="A", data=d.isoformat(), saldo=str(s), programa=prog)
    ago = [date(2026, 8, k) for k in range(1, 32) if date(2026, 8, k).weekday() < 5]
    linhas = [dia(d, "pronamp" if d.day < 14 or d.day > 15 else "pronaf") for d in ago]  # 15/08/2026 é sábado
    linhas += [dia(d, "pronaf") for d in (date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3))]
    hist, _ = realizado_diario(linhas, date(2026, 11, 2))
    check("dia 15 num sábado usa a classificação da sexta (Doc 6, Anexo I, 12.2-b), em qualquer ordem de linhas",
          hist[("A", "2026-08")][0] == "3.1.13.49-9" and realizado_diario(linhas[::-1], date(2026, 11, 2))[0] == hist)
    check("operação que muda de programa depois do dia 15 fica no código antigo naquele mês",
          hist[("A", "2026-08")][0] == "3.1.13.49-9" and perto(hist[("A", "2026-08")][1], 100.0))
    check("dia útil sem linha conta zero (Doc 6, Anexo I, 4.6-b)",
          perto(hist[("A", "2026-09")][1], 300.0 / dias_uteis("2026-09", set())))
    linhas = [dict(dia(date(2026, 11, k), "pronaf", 50.0), id="B") for k in (2, 3, 4)]
    hist, _ = realizado_diario(linhas, date(2026, 11, 5))
    res, _ = projetar({"B": {"codigo": "3.1.30.45-8", "saldo": 50.0, "taxa_aa": 0.0}}, [], [], [], date(2026, 11, 5),
                      date(2026, 11, 30), hist=hist)
    check("mês da data-base: dias antes dela vêm do saldo diário; o código é o do dia 15",
          perto(res[("B", "2026-11")][1], 50.0) and res[("B", "2026-11")][0] == "3.1.30.45-8")

    # oficial (Sisex) ao lado do calculado
    du2 = {"2026-07": 10, "2026-08": 30, "2026-09": 20}
    sr = serie_do_oficial({"2026-07": 10.0, "2026-08": 32.5, "2026-09": 26.0}, 2026, "2026-10", du2, {})
    check("oficial acumulado vira série mensal e volta às mesmas posições",
          perto(sr["2026-08"], 40.0) and all(perto(acum(sr, 2026, m, du2), v)
                                            for m, v in (("2026-07", 10.0), ("2026-08", 32.5), ("2026-09", 26.0))))
    check("mês passado sem posição oficial usa o calculado",
          perto(serie_do_oficial({"2026-08": 32.5}, 2026, "2026-09", du2, {"2026-07": 10.0})["2026-08"], 40.0))
    lin = [["2026/2027", "2026-08", "R", "RO", "3.1.13.49-9", "calculado", "", 5.0, 7.0, "", ""],
           ["2026/2027", "2026-09", "R", "RO", "3.1.13.49-9", "calculado", "", 5.0, 7.0, "", ""]]
    lin = com_oficial(lin, {"2026-08": {"3.2.12.18-7": 4.0}}, {"mes_base": "2026-11", "meses": meses("2026-07", 24)})
    so_of = [l for l in lin if l[4] == "3.2.12.18-7"]
    check("código só no oficial aparece com calculado zero, e vice-versa; posição sem oficial fica em branco",
          lin[0][9:] == [0.0, -7.0] and lin[1][9:] == ["", ""] and len(so_of) == 1
          and so_of[0][3] == "POUP" and so_of[0][8:] == [0.0, 4.0, 4.0])

    # classificação no código do Documento 6
    cl = lambda **k: classificar(dict({"fonte": "RO", "programa": "", "finalidade": "", "recursos": "",
                                       "beneficiario": "produtor", "linha": "", "taxa_aa": 0.0}, **k))[0]
    check("Pronaf custeio 2026/27 até 3% a.a. -> 3.1.13.49-9 (Sisex calcula), ponderador 37% (4.1.34.19-2)",
          cl(programa="pronaf", finalidade="custeio", data_contratacao="2026-08-10", taxa_aa=3.0) == "3.1.13.49-9"
          and CODIGOS["3.1.13.49-9"]["tipo"] == "c" and CODIGOS["3.1.13.49-9"]["fator"] == "1.37"
          and CODIGOS["3.1.13.49-9"]["ponderador"] == "4.1.34.19-2")
    check("Pronaf custeio 2026/27 acima de 3% -> 3.1.13.50-9 (demais operações)",
          cl(programa="pronaf", finalidade="custeio", data_contratacao="2026-08-10", taxa_aa=3.5) == "3.1.13.50-9")
    check("Geral custeio 2023/24 -> 3.1.30.45-8 (a instituição informa)",
          cl(programa="demais", finalidade="custeio", data_contratacao="2024-03-01") == "3.1.30.45-8"
          and CODIGOS["3.1.30.45-8"]["tipo"] == "i")
    check("Pronaf investimento 2026/27 não tem código no Anexo II",
          cl(programa="pronaf", finalidade="investimento", data_contratacao="2026-08-01") == "")
    check("Pronamp investimento entra no limite 3.1.41.39-9",
          CODIGOS[cl(programa="pronamp", finalidade="investimento", data_contratacao="2026-09-01")]["cat"] == "3.1.41.39-9")
    check("Poupança custeio não controlado 2026/27 -> 3.2.12.18-7; LCA CPR 2026/27 -> 3.3.30.76-0 (faculdade)",
          cl(fonte="POUP", finalidade="custeio", recursos="nao_controlados", beneficiario="",
             data_contratacao="2026-09-10") == "3.2.12.18-7"
          and cl(fonte="LCA", linha="cpr", beneficiario="", data_contratacao="2026-10-15") == "3.3.30.76-0"
          and CODIGOS["3.3.30.76-0"]["cat"] == "faculdade")
    dup = []
    for r in REGRAS:  # uma operação típica de cada regra tem que cair só nela
        de = r["de"] or (str(int(r["ate"][:4]) - 1) + r["ate"][4:] if r["ate"] else "2000-01-01")
        ate = r["ate"] or "2027-06-30"
        meio = (date.fromisoformat(de) + (date.fromisoformat(ate) - date.fromisoformat(de)) / 2).isoformat()
        o = {"fonte": r["fonte"], "linha": r["linha"], "data_contratacao": meio,
             "taxa_aa": float(r["taxa_max"] or (float(r["taxa_min"]) + 1 if r["taxa_min"] else 10)),
             "programa": r["programa"] or "demais", "finalidade": r["finalidade"] or "custeio",
             "recursos": r["recursos"] or "controlados", "beneficiario": r["beneficiario"]}
        c, emp = classificar(o)
        if c != r["codigo"] or len(emp) > 1:
            dup.append(r["codigo"])
    check(f"cada um dos {len(REGRAS)} códigos de aplicação tem regra própria, sem empate" + (f" {dup}" if dup else ""),
          not dup)

    # parâmetros da gerência
    M = meses("2026-07", 6)
    out = com_parametros({}, {"RO_PRONAF": {"2026-08": 10.0}},
                         {"prazo": {"RO_PRONAF": 2}, "liberacao": {"RO_PRONAF": {"2026-07": 1e6}},
                          "pct": {"RO_PRONAF": {"2026-08": 60}}}, M, "2026-07")["RO|pronaf"]
    pg = {"RO_PRONAF": {"2026-08": 10.0}, "POUP": {"2026-08": 10.0, "2026-09": 10.0}}
    niv = com_parametros({}, pg, {"pct_geral": 50, "pct_grupo": {"POUP": 90}, "pct": {"POUP": {"2026-09": 0}}}, M, "2026-07")
    check("% de reembolso: geral preenche todos, grupo preenche o grupo, célula muda só ela",
          perto(niv["RO|pronaf"]["2026-09"], 5) and perto(niv["POUP|credito_rural"]["2026-10"], 1 + 10))
    check("% de reembolso padrão = 80%; o parametros.json salvo vira o padrão da próxima rodada",
          perto(com_parametros({}, pg, {}, M, "2026-07")["RO|pronaf"]["2026-09"], 2)
          and perto(com_parametros({}, pg, {"pct_geral": 70}, M, "2026-07")["RO|pronaf"]["2026-09"], 3))
    check("liberação da gerência: meio mês, prazo cheio, meio mês na volta; reembolso de 60% deixa 40% no saldo",
          perto(out["2026-07"], 5e5) and perto(out["2026-08"], 1e6 + 2) and perto(out["2026-09"], 5e5 + 4)
          and perto(out["2026-10"], 4))
    check("parâmetro da gerência não mexe em mês realizado",
          com_parametros({}, {}, {"liberacao": {"LCA": {"2026-07": 1}}}, M, "2026-08").get("LCA|credito_rural") == {})

    # comparação
    a = {("1", "m"): ("RO", "pronaf", 10.0), ("2", "m"): ("RO", "pronamp", 5.0), ("3", "m"): ("POUP", "cpr", 1.0)}
    b2 = {("1", "m"): ("RO", "pronaf", 12.0), ("2", "m"): ("RO", "pronaf", 5.0), ("4", "m"): ("LCA", "faculdade", 7.0)}
    check("comparação acha saldo, classificação, só na A e só na B",
          {l[0]: l[2] for l in comparar(a, b2)} == {"1": "mudou_saldo", "2": "mudou_classificacao", "3": "so_na_A",
                                                   "4": "so_na_B"})
    return checks


# ---------------------------------------------------------------- menu e linha de comando

def menu():
    print(f"\nExigibilidade v{VERSAO}\n  1 autoteste\n  2 exemplo (dados inventados)\n  3 rodar com as entradas reais\n"
          "  4 comparar duas rodadas\n  5 estrutura da projeção atual (sem valores)")
    e = input("Número: ").strip()
    if e == "1":
        return ["autoteste"]
    if e == "2":
        return ["exemplo"]
    if e == "3":
        return ["rodar", input("Pasta das entradas: ").strip().strip('"'), "--data-base",
                input("Data-base AAAA-MM-DD: ").strip()]
    if e == "4":
        return ["comparar", input("Pasta da rodada A: ").strip().strip('"'), input("Pasta da rodada B: ").strip().strip('"')]
    if e == "5":
        return ["estrutura", input("Pasta da projeção atual: ").strip().strip('"'),
                "--mascarar", input("Nomes a esconder, separados por vírgula (Enter = nenhum): ").strip()]
    sys.exit("opção inválida")


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        argv = menu()
    p = argparse.ArgumentParser(description=f"Exigibilidade do crédito rural v{VERSAO}")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("autoteste")
    sub.add_parser("exemplo")
    s = sub.add_parser("rodar")
    s.add_argument("entradas")
    s.add_argument("--data-base", required=True)
    s.add_argument("--tipo-if", choices=["banco", "cooperativa"], default="banco")
    s.add_argument("--rodada", help="nome da pasta da rodada (padrão: data e hora)")
    s = sub.add_parser("comparar")
    s.add_argument("a")
    s.add_argument("b")
    s.add_argument("--mes")
    s = sub.add_parser("estrutura")
    s.add_argument("caminhos", nargs="+")
    s.add_argument("--mascarar", default="")
    a = p.parse_args(argv)

    if a.cmd == "autoteste":
        checks = autoteste()
        ok = all(c.endswith(": OK") for c in checks)
        txt = [f"exigibilidade.py v{VERSAO} — {datetime.now():%Y-%m-%d %H:%M} — Python {sys.version.split()[0]}",
               *checks, f"RESULTADO: {'TUDO OK' if ok else 'HÁ FALHAS'}"]
        (PASTA_SCRIPT / "resumo_verificacao.txt").write_text("\n".join(txt) + "\n", encoding="utf-8")
        print("\n".join(txt))
        return 0 if ok else 1
    if a.cmd == "exemplo":
        cmd_exemplo()
    elif a.cmd == "rodar":
        cmd_rodar(a.entradas, a.data_base, a.tipo_if, a.rodada)
    elif a.cmd == "comparar":
        cmd_comparar(a.a, a.b, a.mes)
    elif a.cmd == "estrutura":
        cmd_estrutura(a.caminhos, a.mascarar)
    return 0


if __name__ == "__main__":
    sys.exit(main())
