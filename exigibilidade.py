"""Exigibilidades do crédito rural (MCR cap. 6): realizado, projeção operação a operação e comparação de rodadas.

Comandos:
  python exigibilidade.py autoteste        testa tudo com dados inventados (resumo_verificacao.txt)
  python exigibilidade.py exemplo          gera dados inventados em exemplo/ e roda o fluxo inteiro
  python exigibilidade.py projetar OPERACOES PARCELAS --data-base AAAA-MM-DD --ate AAAA-MM-DD
                                   [--cenario CENARIO] [--feriados FERIADOS] [--rodada NOME]
  python exigibilidade.py calcular CAPTACOES APLICACOES [MAIS_APLICACOES ...] [--tipo-if banco|cooperativa]
  python exigibilidade.py comparar SALDO_OPERACOES_A SALDO_OPERACOES_B [--mes AAAA-MM]
  python exigibilidade.py estrutura ARQUIVO_OU_PASTA [...] [--mascarar "NOME1,NOME2"]

Sem argumentos (F5 no Spyder) abre um menu. Saídas em saidas/ ao lado do script. Só usa a biblioteca padrão
(o comando estrutura usa openpyxl para ler Excel). Formatos dos arquivos: README.md e Doc Mestre seção 3.
Regras: MCR Atualização nº 760 (22/10/2026).
"""
import argparse
import csv
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

VERSAO = "0.1.0"
PASTA_SCRIPT = Path(__file__).resolve().parent
SAIDAS = PASTA_SCRIPT / "saidas"

DEDUCAO = 500_000_000  # MCR 6-2-2 e 6-7-4
ISENCAO = 10_000_000   # MCR 6-2-5 e 6-7-6

# categorias aceitas por fonte; LIVRE é aceita na projeção mas não cumpre exigibilidade
CATEGORIAS = {
    "RO": {"pronamp", "pronamp_invest", "pronaf", "pronaf_ind_com", "outras"},
    "POUP": {"credito_rural", "cpr"},
    "LCA": {"credito_rural", "faculdade"},
    "LIVRE": {"livre"},
}


# ---------------------------------------------------------------- leitura e utilidades

def num(s):
    s = str(s if s is not None else "").strip()
    if not s:
        return 0.0
    if "," in s:  # formato brasileiro 1.234,56
        s = s.replace(".", "").replace(",", ".")
    return float(s)


def dt(s):
    s = s.strip()
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


def ler(caminho):
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        amostra = f.read(4096)
        f.seek(0)
        sep = max(";,\t|", key=amostra.count)
        return [{(k or "").strip().lower(): v for k, v in r.items()} for r in csv.DictReader(f, delimiter=sep)]


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


def validar_classe(fonte, cat, onde):
    if cat not in CATEGORIAS.get(fonte, ()):
        sys.exit(f"{onde}: fonte/categoria inválida {fonte}/{cat}. Aceitas: "
                 + "; ".join(f"{f}: {', '.join(sorted(c))}" for f, c in CATEGORIAS.items()))


# ---------------------------------------------------------------- exigibilidades (MCR 6-2, 6-4, 6-7)

def pct_ro(safra, tipo_if):
    """MCR 6-2-3-B. safra = ano em que começa o período de cumprimento (julho)."""
    if tipo_if == "cooperativa":
        return {2025: 0.06, 2026: 0.13, 2027: 0.22}.get(safra, 0.315)
    return 0.315


def media(serie, ms, du):
    """Média dos saldos diários (dias úteis) nos meses ms que existem nos dados."""
    ms = [m for m in ms if m in du]
    peso = sum(du[m] for m in ms)
    return (sum(serie.get(m, 0.0) * du[m] for m in ms) / peso) if peso else 0.0


def calcular(safra, du, capt, apl, tipo_if="banco", ate=None):
    """[(fonte, item, exigido, aplicado)] do ano-safra; ate limita os meses de aplicação (média acumulada)."""
    cumpr = [m for m in meses(f"{safra}-07", 12) if not ate or m <= ate]
    calc = meses(f"{safra - 1}-07", 12)      # MCR 6-2-6-a e 6-4-3-a
    calc_lca = meses(f"{safra - 1}-06", 12)  # MCR 6-7-9-a
    a = {k: media(v, cumpr, du) for k, v in apl.items()}
    g = lambda f, c: a.get((f, c), 0.0)
    out = []

    # Recursos Obrigatórios (MCR 6-2)
    exig = max(0.0, media(capt["vsr_avista"], calc, du) - DEDUCAO) * pct_ro(safra, tipo_if)
    if exig <= ISENCAO:
        exig = 0.0
    sub_pronamp, sub_pronaf = 0.50 * exig, 0.35 * exig  # 6-2-8 e 6-2-10
    # shortcut: 2025/26 teve Pronaf 14,35% de abr a jun (6-8-35/36); não tratado, só vale para o ano passado
    ap_pronamp = g("RO", "pronamp") + min(g("RO", "pronamp_invest"), 0.10 * sub_pronamp)  # 6-2-9
    ap_pronaf = g("RO", "pronaf") + min(g("RO", "pronaf_ind_com"), 0.10 * sub_pronaf)      # 6-2-10-A
    # investimento Pronamp acima do teto não conta (6-2-14); pronaf_ind_com acima do teto conta só na geral
    ap_ro = ap_pronamp + g("RO", "pronaf") + g("RO", "pronaf_ind_com") + g("RO", "outras")
    out += [("RO", "exigibilidade", exig, ap_ro),
            ("RO", "sub_pronamp", sub_pronamp, ap_pronamp),
            ("RO", "sub_pronaf", sub_pronaf, ap_pronaf)]

    # Poupança Rural (MCR 6-4): 70% do VSR, sub 95% em crédito rural, até 5% em CPR
    exig_p = 0.70 * media(capt["vsr_poupanca"], calc, du)
    cr_p = g("POUP", "credito_rural")
    out += [("POUP", "exigibilidade", exig_p, cr_p + min(g("POUP", "cpr"), 0.05 * exig_p)),
            ("POUP", "sub_credito_rural", 0.95 * exig_p, cr_p)]

    # LCA (MCR 6-7): 60% de (média - 500 mi), mín. 45% crédito rural, até 55% faculdade, + excesso de RO (9-h)
    exig_l = max(0.0, media(capt["saldo_lca"], calc_lca, du) - DEDUCAO) * 0.60
    if exig_l <= ISENCAO:
        exig_l = 0.0
    cr_l = g("LCA", "credito_rural")
    excesso_ro = max(0.0, ap_ro - exig)
    out += [("LCA", "exigibilidade", exig_l, cr_l + min(g("LCA", "faculdade"), 0.55 * exig_l) + excesso_ro),
            ("LCA", "sub_credito_rural", 0.45 * exig_l, cr_l)]
    return out


def carregar_calculo(cam_capt, cams_apl):
    du, tipo, capt = {}, {}, defaultdict(dict)
    for r in ler(cam_capt):
        m = r["mes"].strip()
        du[m] = int(num(r["dias_uteis"]))
        tipo[m] = (r.get("tipo") or "R").strip().upper()
        for c in ("vsr_avista", "vsr_poupanca", "saldo_lca"):
            capt[c][m] = num(r.get(c))
    apl = defaultdict(lambda: defaultdict(float))
    for cam in cams_apl:
        for i, r in enumerate(ler(cam), start=2):
            fonte, cat, m = r["fonte"].strip().upper(), r["categoria"].strip().lower(), r["mes"].strip()
            validar_classe(fonte, cat, f"{cam} linha {i}")
            if fonte == "LIVRE":
                continue
            if m not in du:
                sys.exit(f"{cam} linha {i}: mês {m} sem dias_uteis no arquivo de captações")
            # fator = ponderador do MCR (ex. 1,37 Pronaf custeio, 6-2-12); vazio = 1
            apl[(fonte, cat)][m] += num(r["saldo_medio"]) * (num(r.get("fator")) or 1.0)
    return du, tipo, capt, apl


def cmd_calcular(cam_capt, cams_apl, tipo_if="banco", saida=None):
    du, tipo, capt, apl = carregar_calculo(cam_capt, cams_apl)
    meses_apl = sorted({m for s in apl.values() for m in s})
    safras = sorted({int(m[:4]) - (m[5:] < "07") for m in meses_apl})
    linhas = []
    for safra in safras:
        for ate in meses(f"{safra}-07", 12):
            if ate in meses_apl:
                for fonte, item, exigido, aplicado in calcular(safra, du, capt, apl, tipo_if, ate):
                    linhas.append([f"{safra}/{safra + 1}", ate, tipo.get(ate, "R"), fonte, item,
                                   round(exigido, 2), round(aplicado, 2), round(aplicado - exigido, 2)])
    saida = gravar(saida or SAIDAS / "exigibilidade.csv",
                   ["ano_safra", "media_ate", "tipo_mes", "fonte", "item", "exigido", "aplicado", "sobra_falta"],
                   linhas)
    ultimo = {l[0]: l[1] for l in linhas}
    for l in linhas:
        if l[1] == ultimo[l[0]]:
            print(f"{l[0]} até {l[1]} {l[3]:4} {l[4]:18} exigido {br_mil(l[5])} aplicado {br_mil(l[6])}"
                  f" sobra/falta {br_mil(l[7])}")
    print(f"Mês a mês (média acumulada) em {saida}")
    return linhas


# ---------------------------------------------------------------- projeção operação a operação

def projetar(ops, parcelas, cenario, data_base, ate, feriados=()):
    """{(id, 'AAAA-MM'): [fonte, categoria, saldo_medio, fator]}: média dos dias úteis do mês.

    No mês em que a operação liquida, ela entra com o último saldo do mês anterior (não conta os dias zerados).
    """
    ops = {k: dict(v) for k, v in ops.items()}
    pag = defaultdict(lambda: defaultdict(float))
    for i, d, v in parcelas:
        pag[i][d] += v
    ultima = {i: max(p) for i, p in pag.items()}
    lib = defaultdict(lambda: defaultdict(float))
    liq, recl, pcts = {}, defaultdict(dict), []
    for c in cenario:
        t, i, d = c["tipo"], c.get("id", ""), c["data"]
        if t == "liberacao":
            lib[i][d] += c["valor"]
            ops.setdefault(i, {"fonte": c["fonte"], "categoria": c["categoria"], "saldo": 0.0,
                               "taxa_aa": c.get("taxa_aa", 0.0), "fator": c.get("fator") or 1.0})
        elif t == "pct_pagamento":
            pcts.append((d, c.get("fonte", ""), c.get("categoria", ""), c["pct"] / 100))
        elif t == "liquidacao":
            liq[i] = d
        elif t == "reclassificacao":
            recl[i][d] = (c["fonte"], c["categoria"])
        else:
            sys.exit(f"cenário: tipo desconhecido '{t}' (use liberacao, pct_pagamento, liquidacao, reclassificacao)")
    pcts.sort()

    def pct(d, fonte, cat):
        p = 1.0
        for ini, f, c, v in pcts:  # vale o último que começou até a data
            if ini <= d and f in ("", fonte) and c in ("", cat):
                p = v
        return p

    out = {}
    # shortcut: loop diário em Python (60 mil ops x 600 dias ~ alguns minutos); vetorizar se ficar lento
    for i, op in ops.items():
        saldo, fonte, cat = op["saldo"], op["fonte"], op["categoria"]
        fat_dia = (1 + op["taxa_aa"] / 100) ** (1 / 365)  # shortcut: juros compostos por dia corrido; ajustar à regra real
        acum = defaultdict(lambda: [0.0, 0])  # mes -> [soma dos saldos, dias úteis]
        classe, fim_mes_ant, liquidou = {}, {}, None
        d = data_base
        while d <= ate:
            if d > data_base:
                saldo *= fat_dia
            if d in recl[i]:
                fonte, cat = recl[i][d]
            saldo += lib[i].get(d, 0.0)
            if d in pag[i]:
                p = pct(d, fonte, cat)
                saldo -= pag[i][d] * p
                if d == ultima[i] and p >= 1:
                    saldo = 0.0
            if liq.get(i) == d:
                saldo = 0.0
            saldo = max(saldo, 0.0)
            m = d.strftime("%Y-%m")
            if d.weekday() < 5 and d not in feriados:
                acum[m][0] += saldo
                acum[m][1] += 1
                classe[m] = (fonte, cat)
            if saldo == 0.0 and (d in pag[i] or liq.get(i) == d) and d >= max(lib[i], default=d):
                liquidou = m
                break
            prox = d + timedelta(days=1)
            if prox.month != d.month:
                fim_mes_ant[prox.strftime("%Y-%m")] = (saldo, fonte, cat)
            d = prox
        for m, (soma, n) in acum.items():
            if not n:
                continue
            fo, ca = classe[m]
            media_m = soma / n
            if m == liquidou and m in fim_mes_ant:
                media_m, fo, ca = fim_mes_ant[m]
            if media_m:
                out[(i, m)] = [fo, ca, media_m, op["fator"]]
    return out


def carregar_projecao(cam_ops, cam_parc, cam_cen=None, cam_fer=None):
    ops = {}
    for n, r in enumerate(ler(cam_ops), start=2):
        fonte, cat = r["fonte"].strip().upper(), r["categoria"].strip().lower()
        validar_classe(fonte, cat, f"{cam_ops} linha {n}")
        ops[r["id"].strip()] = {"fonte": fonte, "categoria": cat, "saldo": num(r["saldo"]),
                                "taxa_aa": num(r.get("taxa_aa")), "fator": num(r.get("fator")) or 1.0}
    parcelas = [(r["id"].strip(), dt(r["data"]), num(r["valor"])) for r in ler(cam_parc)]
    cenario = []
    for n, r in enumerate(ler(cam_cen) if cam_cen else [], start=2):
        c = {"tipo": r["tipo"].strip().lower(), "id": (r.get("id") or "").strip(), "data": dt(r["data"]),
             "valor": num(r.get("valor")), "fonte": (r.get("fonte") or "").strip().upper(),
             "categoria": (r.get("categoria") or "").strip().lower(), "pct": num(r.get("pct")),
             "taxa_aa": num(r.get("taxa_aa"))}
        if c["tipo"] in ("liberacao", "reclassificacao"):
            validar_classe(c["fonte"], c["categoria"], f"{cam_cen} linha {n}")
        cenario.append(c)
    feriados = {dt(r["data"]) for r in ler(cam_fer)} if cam_fer else set()
    return ops, parcelas, cenario, feriados


def cmd_projetar(cam_ops, cam_parc, data_base, ate, cam_cen=None, cam_fer=None, rodada=None):
    ops, parcelas, cenario, feriados = carregar_projecao(cam_ops, cam_parc, cam_cen, cam_fer)
    res = projetar(ops, parcelas, cenario, dt(data_base), dt(ate), feriados)
    rodada = rodada or f"{datetime.now():%Y-%m-%d_%H%M}"
    pasta = SAIDAS / "rodadas" / rodada  # nunca sobrescreve rodada anterior
    if pasta.exists():
        sys.exit(f"rodada {rodada} já existe em {pasta}; use outro --rodada")
    s_op = gravar(pasta / "saldo_operacoes.csv", ["id", "mes", "fonte", "categoria", "saldo_medio"],
                  [[i, m, fo, ca, round(s, 2)] for (i, m), (fo, ca, s, _) in sorted(res.items())])
    agg = defaultdict(float)
    for (i, m), (fo, ca, s, fat) in res.items():
        agg[(m, fo, ca)] += s * fat
    s_ag = gravar(pasta / "aplicacoes_projetadas.csv", ["mes", "fonte", "categoria", "saldo_medio", "fator", "tipo"],
                  [[m, fo, ca, round(s, 2), "", "P"] for (m, fo, ca), s in sorted(agg.items())])
    print(f"{len(ops)} operações -> {s_op}\nagregado (já com ponderador) -> {s_ag}")
    return s_op, s_ag


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
    return {(r["id"].strip(), r["mes"].strip()): (r["fonte"].strip().upper(), r["categoria"].strip().lower(),
                                                  num(r["saldo_medio"]))
            for r in ler(caminho) if not mes or r["mes"].strip() == mes}


def cmd_comparar(cam_a, cam_b, mes=None, saida=None):
    linhas = comparar(carregar_rodada(cam_a, mes), carregar_rodada(cam_b, mes))
    saida = gravar(saida or SAIDAS / "diferencas.csv",
                   ["id", "mes", "motivo", "fonte_A", "categoria_A", "saldo_A", "fonte_B", "categoria_B", "saldo_B",
                    "diferenca"], linhas)
    imp = defaultdict(float)  # impacto por fonte/categoria = soma B - soma A
    for l in linhas:
        if l[3]:
            imp[(l[3], l[4])] -= l[5]
        if l[6]:
            imp[(l[6], l[7])] += l[8]
    print(f"{len(linhas)} diferenças operação/mês -> {saida}\nImpacto B - A por fonte/categoria:")
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


# ---------------------------------------------------------------- exemplo e autoteste (dados inventados)

def cmd_exemplo():
    ex = PASTA_SCRIPT / "exemplo"
    ex.mkdir(exist_ok=True)
    gravar(ex / "captacoes.csv", ["mes", "dias_uteis", "vsr_avista", "vsr_poupanca", "saldo_lca", "tipo"],
           [[m, 21, 4e9 + i * 2e7, 1.5e9 + i * 5e6, 2e9 + i * 3e7, "R" if m <= "2026-10" else "P"]
            for i, m in enumerate(meses("2025-06", 37))])
    base = {("RO", "pronamp"): 560e6, ("RO", "pronamp_invest"): 30e6, ("RO", "pronaf"): 300e6,
            ("RO", "pronaf_ind_com"): 20e6, ("RO", "outras"): 250e6, ("POUP", "credito_rural"): 1000e6,
            ("POUP", "cpr"): 60e6, ("LCA", "credito_rural"): 400e6, ("LCA", "faculdade"): 500e6}
    gravar(ex / "aplicacoes_realizado.csv", ["mes", "fonte", "categoria", "saldo_medio", "fator", "tipo"],
           [[m, fo, ca, v * (1 + 0.01 * i), "1,37" if ca == "pronaf" else "", "R"]
            for i, m in enumerate(meses("2026-07", 4)) for (fo, ca), v in base.items()])
    ops = [("OP%03d" % k, fo, ca, v / 100, 3.0 if ca == "pronaf" else 9.0, "1,37" if ca == "pronaf" else "")
           for k, ((fo, ca), v) in enumerate(base.items(), start=1) for _ in [0]]
    gravar(ex / "operacoes.csv", ["id", "fonte", "categoria", "saldo", "taxa_aa", "fator"],
           [[i, fo, ca, s * 100, t, f] for i, fo, ca, s, t, f in ops])
    gravar(ex / "parcelas.csv", ["id", "data", "valor"],
           [[i, f"{2027 + k % 2}-0{3 + k % 4}-15", s * 100 * 1.12] for k, (i, *_x, s, _t, _f) in enumerate(
               [(o[0], o[1], o[2], o[3], o[4], o[5]) for o in ops])]
           + [["NOVA01", "2027-05-31", 2e8]])
    gravar(ex / "cenario.csv", ["tipo", "id", "data", "valor", "fonte", "categoria", "pct", "taxa_aa"],
           [["liberacao", "NOVA01", "2026-12-01", 2e8, "RO", "pronamp", "", 8.0],
            ["pct_pagamento", "", "2027-01-01", "", "RO", "pronamp", 60.0, ""],
            ["liquidacao", "OP008", "2027-01-15", "", "", "", "", ""],
            ["reclassificacao", "OP005", "2027-02-01", "", "RO", "pronamp", "", ""]])
    gravar(ex / "feriados.csv", ["data"], [["2026-11-20"], ["2026-12-25"], ["2027-01-01"]])
    print(f"Dados inventados em {ex}\n")
    _, ag_sem = cmd_projetar(ex / "operacoes.csv", ex / "parcelas.csv", "2026-11-02", "2028-06-30",
                             None, ex / "feriados.csv", rodada=f"exemplo_sem_cenario_{datetime.now():%H%M%S}")
    _, ag = cmd_projetar(ex / "operacoes.csv", ex / "parcelas.csv", "2026-11-02", "2028-06-30",
                         ex / "cenario.csv", ex / "feriados.csv", rodada=f"exemplo_com_cenario_{datetime.now():%H%M%S}")
    print()
    cmd_calcular(ex / "captacoes.csv", [ex / "aplicacoes_realizado.csv", ag])
    print()
    cmd_comparar(ag_sem.parent / "saldo_operacoes.csv", ag.parent / "saldo_operacoes.csv")


def autoteste():
    checks = []

    def check(desc, ok):
        checks.append(f"CHECK {desc}: {'OK' if ok else 'FALHOU'}")

    perto = lambda a, b: abs(a - b) < 0.01

    # exigibilidades
    ms = meses("2025-06", 25)
    du = {m: 21 for m in ms}
    capt = {"vsr_avista": {m: 4.13e9 for m in ms}, "vsr_poupanca": {m: 1e9 for m in ms},
            "saldo_lca": {m: 2e9 for m in ms}}
    apl = defaultdict(lambda: defaultdict(float))
    for m in meses("2026-07", 12):
        apl[("RO", "pronamp")][m] = 600e6
        apl[("RO", "pronamp_invest")][m] = 100e6
        apl[("RO", "outras")][m] = 600e6
        apl[("POUP", "cpr")][m] = 100e6
        apl[("LCA", "faculdade")][m] = 900e6
    r = {(f, i): (e, a) for f, i, e, a in calcular(2026, du, capt, apl)}
    exig = (4.13e9 - 5e8) * 0.315
    check("RO = 31,5% x (média VSR - 500 mi)", perto(r[("RO", "exigibilidade")][0], exig))
    check("sub Pronamp = 50% e investimento limitado a 10% da sub",
          perto(r[("RO", "sub_pronamp")][0], exig / 2) and perto(r[("RO", "sub_pronamp")][1], 600e6 + 0.05 * exig))
    check("Poupança = 70% do VSR e CPR limitada a 5%", perto(r[("POUP", "exigibilidade")][0], 700e6)
          and perto(r[("POUP", "exigibilidade")][1], 35e6))
    ap_ro = r[("RO", "exigibilidade")][1]
    check("LCA = 60% x (média - 500 mi), faculdade até 55% + excesso de RO",
          perto(r[("LCA", "exigibilidade")][0], 900e6) and perto(r[("LCA", "exigibilidade")][1], 495e6 + ap_ro - exig))
    rc = {(f, i): e for f, i, e, a in calcular(2026, du, capt, apl, "cooperativa")}
    check("cooperativa 2026/27 = 13%", perto(rc[("RO", "exigibilidade")], (4.13e9 - 5e8) * 0.13))
    capt["vsr_avista"] = {m: 5.2e8 for m in ms}
    check("exigibilidade até 10 mi fica isenta", calcular(2026, du, capt, apl)[0][2] == 0.0)

    # projeção
    op = {"A": {"fonte": "POUP", "categoria": "credito_rural", "saldo": 100.0, "taxa_aa": 0.0, "fator": 1.0}}
    p = projetar(op, [("A", date(2026, 12, 15), 100.0)], [], date(2026, 11, 2), date(2027, 1, 31))
    check("mês da liquidação entra com o saldo do fim do mês anterior",
          perto(p[("A", "2026-12")][2], 100.0) and ("A", "2027-01") not in p)
    cen = [{"tipo": "pct_pagamento", "data": date(2026, 1, 1), "pct": 60, "fonte": "", "categoria": ""},
           {"tipo": "reclassificacao", "id": "A", "data": date(2027, 1, 1), "fonte": "RO", "categoria": "outras"},
           {"tipo": "liberacao", "id": "N", "data": date(2027, 1, 4), "valor": 50.0, "fonte": "LCA",
            "categoria": "credito_rural"}]
    fer = {date(2027, 1, 1)}
    p = projetar(op, [("A", date(2026, 12, 15), 100.0)], cen, date(2026, 11, 2), date(2027, 1, 31), fer)
    check("% de pagamento deixa o resto no saldo", perto(p[("A", "2027-01")][2], 40.0))
    check("reclassificação muda fonte/categoria a partir da data", p[("A", "2027-01")][:2] == ["RO", "outras"])
    check("liberação futura entra no saldo", perto(p[("N", "2027-01")][2], 50.0))
    p = projetar(op, [("A", date(2026, 12, 15), 100.0)],
                 cen + [{"tipo": "liquidacao", "id": "A", "data": date(2027, 1, 20)}], date(2026, 11, 2),
                 date(2027, 1, 31), fer)
    check("liquidação antecipada zera e usa o saldo de dezembro", perto(p[("A", "2027-01")][2], 40.0))
    p = projetar({"A": dict(op["A"], taxa_aa=10.0)}, [], [], date(2026, 1, 1), date(2026, 12, 31))
    check("juros sobem o saldo cerca de 10% ao ano", 100 < p[("A", "2026-12")][2] < 110.1)

    # comparação
    a = {("1", "m"): ("RO", "pronaf", 10.0), ("2", "m"): ("RO", "pronamp", 5.0), ("3", "m"): ("POUP", "cpr", 1.0)}
    b = {("1", "m"): ("RO", "pronaf", 12.0), ("2", "m"): ("RO", "pronaf", 5.0), ("4", "m"): ("LCA", "faculdade", 7.0)}
    check("comparação acha saldo, classificação, só na A e só na B",
          {l[0]: l[2] for l in comparar(a, b)} == {"1": "mudou_saldo", "2": "mudou_classificacao", "3": "so_na_A",
                                                  "4": "so_na_B"})
    return checks


# ---------------------------------------------------------------- menu e linha de comando

def menu():
    print(f"\nExigibilidade v{VERSAO}\n  1 autoteste\n  2 exemplo (dados inventados, fluxo inteiro)\n"
          "  3 estrutura da projeção atual (sem valores)\nOutros comandos: veja o README.")
    e = input("Número: ").strip()
    if e == "1":
        return ["autoteste"]
    if e == "2":
        return ["exemplo"]
    if e == "3":
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
    s = sub.add_parser("projetar")
    s.add_argument("operacoes")
    s.add_argument("parcelas")
    s.add_argument("--data-base", required=True)
    s.add_argument("--ate", required=True)
    s.add_argument("--cenario")
    s.add_argument("--feriados")
    s.add_argument("--rodada", help="nome da pasta da rodada (padrão: data e hora)")
    s = sub.add_parser("calcular")
    s.add_argument("captacoes")
    s.add_argument("aplicacoes", nargs="+", help="realizado e projetado; não repita mês/fonte/categoria")
    s.add_argument("--tipo-if", choices=["banco", "cooperativa"], default="banco")
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
    elif a.cmd == "projetar":
        cmd_projetar(a.operacoes, a.parcelas, a.data_base, a.ate, a.cenario, a.feriados, a.rodada)
    elif a.cmd == "calcular":
        cmd_calcular(a.captacoes, a.aplicacoes, a.tipo_if)
    elif a.cmd == "comparar":
        cmd_comparar(a.a, a.b, a.mes)
    elif a.cmd == "estrutura":
        cmd_estrutura(a.caminhos, a.mascarar)
    return 0


if __name__ == "__main__":
    sys.exit(main())
