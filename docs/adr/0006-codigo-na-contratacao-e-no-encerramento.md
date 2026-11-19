# Código do mês na contratação e no encerramento

O Doc 6, Anexo I, 12.2-b manda alocar o saldo pela fonte que a operação tem no Sicor no dia 15, mas não diz o que fazer com a operação que não existe no dia 15. Por decisão do usuário: a operação contratada depois do dia 15 entra no mês com o código da fonte da contratação (a 1ª linha do mês no `saldos_diarios`); a operação encerrada antes do dia 15 entra no mês com o código do dia 15 do mês anterior.
