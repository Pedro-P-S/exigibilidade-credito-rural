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

VERSAO = "0.2.0"
PASTA_SCRIPT = Path(__file__).resolve().parent
SAIDAS = PASTA_SCRIPT / "saidas"

DEDUCAO = 500_000_000  # MCR 6-2-2 e 6-7-4
ISENCAO = 10_000_000   # MCR 6-2-5 e 6-7-6

# categoria de cada saldo, por fonte (o que ela cumpre e qual limite tem)
CATEGORIAS = {
    "RO": {"pronaf", "pronaf_ind_com", "pronamp", "pronamp_invest", "pronamp_peq_medio",
           "demais", "demais_abc", "demais_pca"},
    "POUP": {"credito_rural", "cpr"},
    "LCA": {"credito_rural", "faculdade"},
    "LIVRE": {"livre"},
}
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
             ("ap_cr", "3.2.10.00-0", "Aplicado - crédito rural"),
             ("ap_cpr", "2.2.40.00-2", "Aplicado - CPR (até 5%)"),
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
# visão dos diretores: rótulo, anexo, chave do excesso, chave da deficiência
RESUMO = [("Recursos Obrigatórios", "RO", "exc_total", "def_total"),
          ("RO - Exigibilidade Geral", "RO", "exc_geral", "def_geral"),
          ("RO - Pronaf", "RO", "exc_pronaf", "def_pronaf"),
          ("RO - Pronamp", "RO", "exc_pronamp", "def_pronamp"),
          ("Poupança Rural", "POUP", "exc_total", "def_total"),
          ("Poupança - crédito rural (95%)", "POUP", "exc_sub", "def_sub"),
          ("LCA", "LCA", "exc_total", "def_total"),
          ("LCA - crédito rural (45%)", "LCA", "exc_sub", "def_sub")]


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
        linhas = list(openpyxl.load_workbook(caminho, read_only=True, data_only=True).worksheets[0]
                      .iter_rows(values_only=True))
        cab = [str(c or "").strip().lower() for c in linhas[0]]
        return [{k: "" if v is None else (v.isoformat() if hasattr(v, "isoformat") else str(v)) for k, v in zip(cab, l)}
                for l in linhas[1:] if any(v is not None for v in l)]
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        amostra = f.read(4096)
        f.seek(0)
        sep = max(";,\t|", key=amostra.count)
        return [{(k or "").strip().lower(): v for k, v in r.items()} for r in csv.DictReader(f, delimiter=sep)]


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


def validar_classe(fonte, cat, onde):
    if cat not in CATEGORIAS.get(fonte, ()):
        sys.exit(f"{onde}: fonte/categoria inválida {fonte}/{cat}. Aceitas: "
                 + "; ".join(f"{f}: {', '.join(sorted(c))}" for f, c in CATEGORIAS.items()))


def grupo(fonte, cat):
    if fonte == "RO":
        return "RO_PRONAF" if cat.startswith("pronaf") else "RO_PRONAMP" if cat.startswith("pronamp") else "RO_DEMAIS"
    return fonte if fonte in GRUPOS else None


# ---------------------------------------------------------------- projeção operação a operação

def projetar(ops, parcelas, libs, cenario, data_base, ate, feriados=frozenset()):
    """Saldo diário de cada operação até a liquidação.

    Devolve (res, pagos): res = {(id, mes): [fonte, categoria, saldo_medio, fator]}, pagos = {(grupo, mes): valor}.
    Saldo médio = soma dos saldos dos dias úteis / dias úteis do mês, com zero antes da 1ª liberação e depois da
    liquidação (Doc 6, Anexo I, 4.6-b). A fonte/categoria do mês é a do dia 15 (Doc 6, Anexo I, 12.2-b).
    """
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
            recl[c["id"]][c["data"]] = (c["fonte"], c["categoria"])
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
        saldo, fonte, cat = op["saldo"], op["fonte"], op["categoria"]
        fat_dia = (1 + op["taxa_aa"] / 100) ** (1 / 365)  # shortcut: juros compostos por dia corrido; ajustar à regra real
        soma, cls = defaultdict(float), {}
        m0 = data_base.strftime("%Y-%m")
        cls[m0] = (fonte, cat)
        # shortcut: dias úteis do mês da data-base antes dela entram com o saldo da data-base
        soma[m0] += saldo * sum(1 for k in range(1, data_base.day) if util(data_base.replace(day=k)))
        ult_lib = max(lib[i], default=None)
        d = data_base
        while d <= ate:
            m = d.strftime("%Y-%m")
            if d > data_base:
                saldo *= fat_dia
            if d in recl[i]:
                fonte, cat = recl[i][d]
            saldo += lib[i].get(d, 0.0)
            if d in pag[i]:
                pago = saldo if d == ultima[i] else min(pag[i][d], saldo)  # a última parcela liquida
                saldo -= pago
                pagos[(grupo(fonte, cat), m)] += pago
            if liq.get(i) == d:
                saldo = 0.0
            if m not in cls or d.day <= 15:
                cls[m] = (fonte, cat)
            if util(d):
                soma[m] += saldo
            if saldo <= 0.0 and (ult_lib is None or d >= ult_lib):
                break  # liquidada: daqui em diante só zeros
            d += timedelta(days=1)
        for m, s in soma.items():
            if s > 0:
                res[(i, m)] = [*cls[m], s / n_uteis(m), op["fator"]]
    return res, {k: v for k, v in pagos.items() if k[0]}


def carregar_entradas(pasta, data_base, feriados):
    """Lê a pasta de entradas. Detalhe dos arquivos no README."""
    p = Path(pasta)
    ops = {}
    for n, r in enumerate(ler(achar(p, "operacoes", True)), start=2):
        fonte, cat = r["fonte"].strip().upper(), r["categoria"].strip().lower()
        validar_classe(fonte, cat, f"operacoes linha {n}")
        ops[r["id"].strip()] = {"fonte": fonte, "categoria": cat, "saldo": num(r["saldo"]),
                                "taxa_aa": num(r.get("taxa_aa")), "fator": num(r.get("fator")) or 1.0}
    parcelas = [(r["id"].strip(), dt(r["data"]), num(r["valor"])) for r in ler(achar(p, "parcelas", True))]
    libs = []
    for n, r in enumerate(ler(achar(p, "liberacoes")) if achar(p, "liberacoes") else [], start=2):
        i = r["id"].strip()
        extra = {"fonte": (r.get("fonte") or "").strip().upper(), "categoria": (r.get("categoria") or "").strip().lower(),
                 "taxa_aa": num(r.get("taxa_aa")), "fator": num(r.get("fator")) or 1.0}
        if i not in ops:
            validar_classe(extra["fonte"], extra["categoria"], f"liberacoes linha {n} (operação nova precisa de fonte/categoria)")
        libs.append((i, dt(r["data"]), num(r["valor"]), extra))
    cenario = []
    for n, r in enumerate(ler(achar(p, "cenario")) if achar(p, "cenario") else [], start=2):
        c = {"tipo": r["tipo"].strip().lower(), "id": r["id"].strip(), "data": dt(r["data"]),
             "fonte": (r.get("fonte") or "").strip().upper(), "categoria": (r.get("categoria") or "").strip().lower()}
        if c["tipo"] == "reclassificacao":
            validar_classe(c["fonte"], c["categoria"], f"cenario linha {n}")
        cenario.append(c)
    mes_base = data_base.strftime("%Y-%m")
    realizado, avisos = defaultdict(lambda: defaultdict(float)), []
    for n, r in enumerate(ler(achar(p, "realizado")) if achar(p, "realizado") else [], start=2):
        fonte, cat, m = r["fonte"].strip().upper(), r["categoria"].strip().lower(), mes_da_linha(r)
        validar_classe(fonte, cat, f"realizado linha {n}")
        if m >= mes_base:
            avisos.append(f"realizado linha {n}: mês {m} não é anterior à data-base; ignorado (vale a projeção)")
        elif fonte != "LIVRE":
            realizado[f"{fonte}|{cat}"][m] += num(r["saldo_medio"]) * (num(r.get("fator")) or 1.0)
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
    return ops, parcelas, libs, cenario, realizado, vsr, lca, dict(dirs), params, avisos


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
# Estas três funções têm cópia idêntica em JavaScript no HTML (simulação); o HTML confere que dão o mesmo resultado.

def com_parametros(apl, pagos, params, meses_h, mes_base):
    """Soma às aplicações as liberações da gerência (voltam em parcela única após o prazo) e o que não for
    reembolsado pelo % de reembolso (fica no saldo até o fim do horizonte). Valores da gerência em R$ milhões."""
    out = {k: dict(v) for k, v in apl.items()}
    for g, (fo, ca) in GRUPOS.items():
        prazo = max(1, int(params.get("prazo", {}).get(g, 12)))
        alvo = out.setdefault(f"{fo}|{ca}", {})
        for j, mj in enumerate(meses_h):
            if mj < mes_base:
                continue
            lib = params.get("liberacao", {}).get(g, {}).get(mj, 0) * 1e6
            pct = params.get("pct", {}).get(g, {}).get(mj, 100)
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
            out.setdefault(mod, {})[m] = [c * 1e6, a * 1e6]
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
    ro["ap_pronaf"] = c("RO|pronaf") + min(c("RO|pronaf_ind_com"), 0.10 * ro["sub_pronaf"]) + dc("pronaf", 1)  # 6-2-10-A
    ro["ap_pronamp"] = (c("RO|pronamp") + min(c("RO|pronamp_invest"), 0.10 * ro["sub_pronamp"])  # 6-2-9
                        + min(c("RO|pronamp_peq_medio"), 0.10 * ro["sub_pronamp"]) + dc("pronamp", 1))  # 6-2-8-b
    ro["ap_geral"] = (c("RO|demais") + min(c("RO|demais_abc"), 0.036 * ro["geral"])  # 6-8-8
                      + min(c("RO|demais_pca"), 0.024 * ro["geral"]) + dc("geral", 1))  # 6-8-9
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
    po["ap_cr"] = c("POUP|credito_rural") + dc("poup", 1)
    po["ap_cpr"] = min(c("POUP|cpr"), 0.05 * po["exig_total"])            # 6-4-11
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
    lc["ap_cr"] = c("LCA|credito_rural") + dc("lca_cr", 1) + lc["exc_ro"]
    lc["ap_fac"] = min(c("LCA|faculdade") + dc("lca_livre", 1), lc["faculdade"])
    lc["ap_total"] = lc["ap_cr"] + lc["ap_fac"]
    lc["def_sub"] = max(0.0, lc["sub"] - lc["ap_cr"])
    lc["exc_sub"] = max(0.0, lc["ap_cr"] - lc["sub"])
    lc["def_total"] = max(0.0, lc["exig_total"] - lc["ap_total"])
    lc["exc_total"] = max(0.0, lc["ap_total"] - lc["exig_total"])
    return {"RO": ro, "POUP": po, "LCA": lc}


def calcular(dados, params):
    """{safra: {mes: sisex}} das duas safras do horizonte."""
    apl = com_parametros(dados["apl"], dados["pagos"], params, dados["meses"], dados["mes_base"])
    dirs = dir_final(dados["dir"], params)
    return {str(s): {m: sisex(s, m, apl, dirs, dados["bases"][str(s)], dados["du"]) for m in meses(f"{s}-07", 12)}
            for s in dados["safras"]}


def montar(pasta, data_base, tipo_if="banco"):
    """Lê as entradas, projeta e devolve (dados, res_operacoes, avisos)."""
    db = dt(data_base) if isinstance(data_base, str) else data_base
    feriados_arq = achar(pasta, "feriados")
    feriados = {dt(r["data"]) for r in ler(feriados_arq)} if feriados_arq else set()
    ops, parcelas, libs, cenario, realizado, vsr, lca, dirs, params, avisos = carregar_entradas(pasta, db, feriados)
    s0 = safra_de(db.strftime("%Y-%m"))
    meses_h = meses(f"{s0}-07", 24)
    res, pagos = projetar(ops, parcelas, libs, cenario, db, date(s0 + 2, 6, 30), feriados)
    apl = defaultdict(dict, {k: dict(v) for k, v in realizado.items()})
    for (i, m), (fo, ca, s, fat) in res.items():
        if fo != "LIVRE":
            apl[f"{fo}|{ca}"][m] = apl[f"{fo}|{ca}"].get(m, 0.0) + s * fat
    mes_base = db.strftime("%Y-%m")
    sem_real = [m for m in meses_h if m < mes_base and not any(m in v for v in realizado.values())]
    if sem_real:
        avisos.append(f"sem realizado para {', '.join(sem_real)}: esses meses entram com zero")
    pg = defaultdict(dict)
    for (g, m), v in pagos.items():
        pg[g][m] = v
    dados = {"versao": VERSAO, "data_base": db.isoformat(), "mes_base": mes_base, "tipo_if": tipo_if,
             "meses": meses_h, "safras": [s0, s0 + 1], "du": {m: dias_uteis(m, feriados) for m in meses_h},
             "bases": {str(s): bases(s, vsr, lca, tipo_if, avisos) for s in (s0, s0 + 1)},
             "apl": dict(apl), "pagos": dict(pg), "dir": dirs, "params": params,
             "grupos": list(GRUPOS), "dir_modalidades": DIR_MODALIDADES, "linhas": LINHAS, "resumo": RESUMO}
    return dados, res, avisos


def cmd_rodar(pasta, data_base, tipo_if="banco", rodada=None):
    dados, res, avisos = montar(pasta, data_base, tipo_if)
    rodada = rodada or f"{datetime.now():%Y-%m-%d_%H%M%S}"
    destino = SAIDAS / "rodadas" / rodada  # nunca sobrescreve rodada anterior
    if destino.exists():
        sys.exit(f"rodada {rodada} já existe em {destino}; use outro --rodada")
    resultado = calcular(dados, dados["params"])
    dados["rodada"], dados["avisos"], dados["python"] = rodada, avisos, resultado
    gravar(destino / "saldo_operacoes.csv", ["id", "mes", "fonte", "categoria", "saldo_medio"],
           [[i, m, fo, ca, round(s, 2)] for (i, m), (fo, ca, s, _) in sorted(res.items())])
    gravar(destino / "aplicacoes.csv", ["fonte_categoria", "mes", "saldo_medio_ponderado"],
           [[k, m, round(v, 2)] for k, porm in sorted(dados["apl"].items()) for m, v in sorted(porm.items())])
    gravar(destino / "exigibilidade.csv", ["ano_safra", "posicao", "tipo", "anexo", "codigo", "item", "valor"],
           [[f"{s}/{int(s) + 1}", m, "R" if m < dados["mes_base"] else "P", an, cod, desc, round(v[an][k], 2)]
            for s, porm in resultado.items() for m, v in porm.items() for an in LINHAS for k, cod, desc in LINHAS[an]])
    (destino / "relatorio.html").write_text(html(dados), encoding="utf-8")
    for a in avisos:
        print("AVISO:", a)
    for s, porm in resultado.items():
        jun = porm[f"{int(s) + 1}-06"]
        print(f"\n{s}/{int(s) + 1} - fechamento de junho (sobra + / falta -):")
        for rot, an, exc, dfc in RESUMO:
            print(f"  {rot:32} {br_mil(jun[an][exc] - jun[an][dfc])}")
    print(f"\nRelatório: {destino / 'relatorio.html'}")
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


# ---------------------------------------------------------------- relatório HTML (visão dos diretores + gestores)

HTML_MODELO = r"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Exigibilidade do Crédito Rural</title>
<style>
:root{--bg:#f7f7f5;--card:#fff;--tx:#1d1d1b;--mut:#6b6b66;--ln:#e2e1dc;--pos:#1f7a4d;--neg:#b3261e;--acc:#2f5d8a;--posbg:#e6f2ea;--negbg:#fbe9e7}
@media (prefers-color-scheme:dark){:root{--bg:#161615;--card:#20201e;--tx:#ecebe6;--mut:#a3a29c;--ln:#3a3936;--pos:#6cc799;--neg:#f2867c;--acc:#8db4dd;--posbg:#1d3328;--negbg:#3a2220}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
header{padding:20px 24px 8px}h1{font-size:20px;margin:0 0 4px}h2{font-size:16px;margin:24px 0 8px}h3{font-size:14px;margin:16px 0 6px;color:var(--mut)}
.meta{color:var(--mut);font-size:13px}.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:12px;margin-left:8px}
.ok{background:var(--posbg);color:var(--pos)}.falhou{background:var(--negbg);color:var(--neg)}
nav{display:flex;gap:4px;padding:0 24px;border-bottom:1px solid var(--ln)}nav button{background:none;border:0;border-bottom:2px solid transparent;padding:10px 12px;color:var(--mut);font:inherit;cursor:pointer}
nav button.ativo{color:var(--tx);border-color:var(--acc)}main{padding:8px 24px 40px}section{display:none}section.ativo{display:block}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px}.card{background:var(--card);border:1px solid var(--ln);border-radius:8px;padding:12px}
.card .r{color:var(--mut);font-size:12px}.card .v{font-size:20px;font-weight:600;margin-top:4px;font-variant-numeric:tabular-nums}
.tab{overflow-x:auto;background:var(--card);border:1px solid var(--ln);border-radius:8px}table{border-collapse:collapse;width:max-content;min-width:100%}
th,td{padding:5px 8px;border-bottom:1px solid var(--ln);text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}th:first-child,td:first-child{text-align:left;position:sticky;left:0;background:var(--card)}
th{font-weight:600;color:var(--mut);font-size:12px}.pos{color:var(--pos)}.neg{color:var(--neg)}.cod{color:var(--mut);font-size:12px;margin-right:6px}
input{width:72px;padding:3px 4px;border:1px solid var(--ln);border-radius:4px;background:var(--bg);color:var(--tx);font:inherit;text-align:right}
.acoes{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.acoes button,.acoes label{padding:6px 12px;border:1px solid var(--ln);border-radius:6px;background:var(--card);color:var(--tx);font:inherit;cursor:pointer}
.avisos{background:var(--negbg);color:var(--neg);border-radius:8px;padding:8px 12px;margin:8px 0;font-size:13px}.nota{color:var(--mut);font-size:12px}
@media (max-width:600px){header,main{padding-left:16px;padding-right:16px}nav{padding:0 16px}}
</style></head><body>
<header><h1>Exigibilidade do Crédito Rural</h1><div class="meta" id="meta"></div></header>
<nav><button class="ativo" data-s="geral">Visão geral</button><button data-s="detalhe">Detalhe dos anexos</button><button data-s="param">Parâmetros e simulação</button></nav>
<main>
<section id="geral" class="ativo"></section>
<section id="detalhe"></section>
<section id="param">
<p class="nota">Valores em R$ milhões. Liberação e % de reembolso por grupo e mês; a liberação volta em parcela única depois do prazo. DIR: saldo médio do mês (captação e aplicação). Tudo recalcula na hora.</p>
<div class="acoes"><button id="salvar">Salvar parâmetros (parametros.json)</button><label>Carregar parâmetros <input type="file" id="carregar" accept=".json" hidden></label><button id="restaurar">Voltar aos parâmetros da rodada</button></div>
<div id="campos"></div>
</section>
</main>
<script type="application/json" id="dados">__DADOS__</script>
<script>
"use strict";
const D = JSON.parse(document.getElementById("dados").textContent);
const GR = {RO_PRONAF:["RO","pronaf"],RO_PRONAMP:["RO","pronamp"],RO_DEMAIS:["RO","demais"],POUP:["POUP","credito_rural"],LCA:["LCA","credito_rural"]};
const NOME_GR = {RO_PRONAF:"RO - Pronaf",RO_PRONAMP:"RO - Pronamp",RO_DEMAIS:"RO - Demais",POUP:"Poupança Rural",LCA:"LCA"};
const NOME_DIR = {geral:"DIR-Geral",pronaf:"DIR-Pronaf",pronamp:"DIR-Pronamp",poup:"DIR-Poup",lca_cr:"DIR-LCA-CR",lca_livre:"DIR-LCA-Livre"};
let P = JSON.parse(JSON.stringify(D.params || {}));

// ---- cópia exata de com_parametros, dir_final, acum e sisex do Python
function mesesSafra(s){const r=[];let a=s,m=7;for(let i=0;i<12;i++){r.push(a+"-"+String(m).padStart(2,"0"));if(m===12){a++;m=1}else m++}return r}
function g(o,...ks){for(const k of ks){if(o==null)return undefined;o=o[k]}return o}
function comParametros(apl,pagos,P,M,mesBase){
  const out={};for(const k in apl)out[k]=Object.assign({},apl[k]);
  for(const gr in GR){const [fo,ca]=GR[gr];const prazo=Math.max(1,parseInt(g(P,"prazo",gr)??12));const key=fo+"|"+ca;out[key]=out[key]||{};const alvo=out[key];
    M.forEach((mj,j)=>{if(mj<mesBase)return;const lib=(g(P,"liberacao",gr,mj)||0)*1e6;const pct=g(P,"pct",gr,mj)??100;const u=(1-pct/100)*(g(pagos,gr,mj)||0);
      if(!lib&&!u)return;
      for(let k=j;k<M.length;k++){const dl=k<=j+prazo?lib*((k===j||k===j+prazo)?0.5:1):0;alvo[M[k]]=(alvo[M[k]]||0)+dl+u*(k===j?0.5:1)}})}
  return out}
function dirFinal(dirs,P){const out={};for(const mod of D.dir_modalidades){out[mod]={};for(const m in (dirs[mod]||{}))out[mod][m]=dirs[mod][m].slice()}
  for(const mod in (P.dir||{}))for(const m in P.dir[mod]){const [c,a]=P.dir[mod][m];(out[mod]=out[mod]||{})[m]=[c*1e6,a*1e6]}return out}
function acum(serie,s,m){let t=0,p=0;for(const k of mesesSafra(s)){if(k>m)break;const w=D.du[k]||0;t+=(serie[k]||0)*w;p+=w}return p?t/p:0}
function sisex(s,m,apl,dirs,base){
  const c=k=>acum(apl[k]||{},s,m);const dc=(mod,i)=>{const sr={};for(const k in (dirs[mod]||{}))sr[k]=dirs[mod][k][i];return acum(sr,s,m)};
  const p=base.ro_propria;const ro={vsr_media:base.vsr_avista,propria:p,capt_dir_geral:dc("geral",0),capt_dir_pronaf:dc("pronaf",0),capt_dir_pronamp:dc("pronamp",0)};
  ro.exig_total=p+ro.capt_dir_geral+ro.capt_dir_pronaf+ro.capt_dir_pronamp;ro.sub_pronaf=0.35*p+ro.capt_dir_pronaf;ro.sub_pronamp=0.50*p+ro.capt_dir_pronamp;
  ro.geral=p-0.35*p-0.50*p+ro.capt_dir_geral;
  ro.ap_pronaf=c("RO|pronaf")+Math.min(c("RO|pronaf_ind_com"),0.10*ro.sub_pronaf)+dc("pronaf",1);
  ro.ap_pronamp=(c("RO|pronamp")+Math.min(c("RO|pronamp_invest"),0.10*ro.sub_pronamp)+Math.min(c("RO|pronamp_peq_medio"),0.10*ro.sub_pronamp)+dc("pronamp",1));
  ro.ap_geral=(c("RO|demais")+Math.min(c("RO|demais_abc"),0.036*ro.geral)+Math.min(c("RO|demais_pca"),0.024*ro.geral)+dc("geral",1));
  ro.ap_total=ro.ap_pronaf+ro.ap_geral+ro.ap_pronamp;
  ro.def_pronaf=Math.max(0,ro.sub_pronaf-ro.ap_pronaf);ro.exc_pronaf=Math.max(0,ro.ap_pronaf-ro.sub_pronaf);
  ro.def_pronamp=Math.max(0,ro.sub_pronamp-ro.ap_pronamp);ro.exc_pronamp=Math.max(0,ro.ap_pronamp-ro.sub_pronamp);
  ro.def_geral=Math.max(0,ro.exig_total-(ro.ap_total+ro.def_pronaf+ro.def_pronamp));ro.exc_geral=Math.max(0,ro.ap_geral-ro.geral);
  ro.def_total=ro.def_geral+ro.def_pronaf+ro.def_pronamp;ro.exc_total=Math.max(0,ro.exc_pronaf+ro.exc_pronamp+ro.ap_geral-ro.geral);
  const pp=base.poup_propria;const po={vsr_media:base.vsr_poupanca,propria:pp,capt_dir:dc("poup",0)};
  po.exig_total=pp+po.capt_dir;po.sub=0.95*pp+po.capt_dir;po.ap_cr=c("POUP|credito_rural")+dc("poup",1);po.ap_cpr=Math.min(c("POUP|cpr"),0.05*po.exig_total);
  po.ap_total=po.ap_cr+po.ap_cpr;po.def_sub=Math.max(0,po.sub-po.ap_cr);po.exc_sub=Math.max(0,po.ap_cr-po.sub);
  po.def_total=Math.max(0,po.exig_total-po.ap_total);po.exc_total=Math.max(0,po.ap_total-po.exig_total);
  const lp=base.lca_propria[m]||0;const lc={lca_media:base.lca_media[m]||0,propria:lp,capt_dir_cr:dc("lca_cr",0),capt_dir_livre:dc("lca_livre",0)};
  lc.exig_total=lp+lc.capt_dir_cr+lc.capt_dir_livre;lc.sub=0.45*lp+lc.capt_dir_cr;lc.faculdade=0.55*lp+lc.capt_dir_livre;
  lc.exc_ro=m.slice(5)==="06"?ro.exc_total:0;lc.ap_cr=c("LCA|credito_rural")+dc("lca_cr",1)+lc.exc_ro;
  lc.ap_fac=Math.min(c("LCA|faculdade")+dc("lca_livre",1),lc.faculdade);lc.ap_total=lc.ap_cr+lc.ap_fac;
  lc.def_sub=Math.max(0,lc.sub-lc.ap_cr);lc.exc_sub=Math.max(0,lc.ap_cr-lc.sub);lc.def_total=Math.max(0,lc.exig_total-lc.ap_total);lc.exc_total=Math.max(0,lc.ap_total-lc.exig_total);
  return {RO:ro,POUP:po,LCA:lc}}
function calcular(P){const apl=comParametros(D.apl,D.pagos,P,D.meses,D.mes_base);const dirs=dirFinal(D.dir,P);const out={};
  for(const s of D.safras){out[s]={};for(const m of mesesSafra(s))out[s][m]=sisex(s,m,apl,dirs,D.bases[s])}return out}

// ---- tela
const mi=v=>(v/1e6).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1});
const cls=v=>v>0.05e6?"pos":v<-0.05e6?"neg":"";
const rot=m=>["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"][+m.slice(5)-1]+"/"+m.slice(2,4);
const tipo=m=>m<D.mes_base?"R":"P";
function conferir(){let ok=true;const r=calcular(D.params||{});
  for(const s in D.python)for(const m in D.python[s])for(const an in D.python[s][m])for(const k in D.python[s][m][an])
    if(Math.abs(D.python[s][m][an][k]-r[s][m][an][k])>0.01)ok=false;return ok}
function cabecalho(s){return "<tr><th></th>"+mesesSafra(s).map(m=>`<th>${rot(m)} ${tipo(m)}</th>`).join("")+"</tr>"}
function render(){const R=calcular(P);let h="",d="";
  for(const s of D.safras){const jun=(s+1)+"-06";
    h+=`<h2>Ano-safra ${s}/${s+1} · fechamento de junho (sobra + / falta −, R$ mi)</h2><div class="cards">`;
    for(const [r,an,e,f] of D.resumo){const v=R[s][jun][an][e]-R[s][jun][an][f];h+=`<div class="card"><div class="r">${r}</div><div class="v ${cls(v)}">${mi(v)}</div></div>`}
    h+=`</div><h3>Posição de cada mês (média acumulada desde julho)</h3><div class="tab"><table>${cabecalho(s)}`;
    for(const [r,an,e,f] of D.resumo){h+=`<tr><td>${r}</td>`+mesesSafra(s).map(m=>{const v=R[s][m][an][e]-R[s][m][an][f];return `<td class="${cls(v)}">${mi(v)}</td>`}).join("")+"</tr>"}
    h+="</table></div>";
    for(const an of ["RO","POUP","LCA"]){const nome={RO:"Anexo II - Recursos Obrigatórios",POUP:"Anexo III - Poupança Rural",LCA:"Anexo IV - LCA"}[an];
      d+=`<h2>${s}/${s+1} · ${nome} (R$ mi)</h2><div class="tab"><table>${cabecalho(s)}`;
      for(const [k,cod,desc] of D.linhas[an])d+=`<tr><td><span class="cod">${cod}</span>${desc}</td>`+mesesSafra(s).map(m=>`<td>${mi(R[s][m][an][k])}</td>`).join("")+"</tr>";
      d+="</table></div>"}}
  document.getElementById("geral").innerHTML=h+'<p class="nota">R = realizado, P = projetado. A fonte de cada operação no mês é a do dia 15.</p>';
  document.getElementById("detalhe").innerHTML=d}
function campo(caminho,valor,mult){return `<input data-c="${caminho}" value="${valor===undefined?"":String(valor).replace(".",",")}">`}
function montarCampos(){const M=D.meses.filter(m=>m>=D.mes_base);const cab="<tr><th></th>"+M.map(m=>`<th>${rot(m)}</th>`).join("")+"</tr>";
  let h="<h2>Prazo de retorno das liberações novas (meses)</h2><div class='tab'><table><tr>"+D.grupos.map(x=>`<th>${NOME_GR[x]}</th>`).join("")+"</tr><tr>"+D.grupos.map(x=>`<td>${campo("prazo|"+x,g(P,"prazo",x)??12)}</td>`).join("")+"</tr></table></div>";
  h+="<h2>Liberação da gerência (R$ mi no mês)</h2><div class='tab'><table>"+cab+D.grupos.map(x=>`<tr><td>${NOME_GR[x]}</td>`+M.map(m=>`<td>${campo("liberacao|"+x+"|"+m,g(P,"liberacao",x,m))}</td>`).join("")+"</tr>").join("")+"</table></div>";
  h+="<h2>% de reembolso das parcelas (100 = paga tudo)</h2><div class='tab'><table>"+cab+D.grupos.map(x=>`<tr><td>${NOME_GR[x]}</td>`+M.map(m=>`<td>${campo("pct|"+x+"|"+m,g(P,"pct",x,m)??100)}</td>`).join("")+"</tr>").join("")+"</table></div>";
  h+="<h2>DIR (R$ mi, saldo médio do mês): captação / aplicação</h2><div class='tab'><table>"+cab;
  for(const mod of D.dir_modalidades)for(const [i,nm] of [[0,"captação"],[1,"aplicação"]])
    h+=`<tr><td>${NOME_DIR[mod]} ${nm}</td>`+M.map(m=>{const v=g(P,"dir",mod,m)?P.dir[mod][m][i]:(g(D.dir,mod,m)?D.dir[mod][m][i]/1e6:0);return `<td>${campo("dir|"+mod+"|"+m+"|"+i,Math.round(v*100)/100)}</td>`}).join("")+"</tr>";
  document.getElementById("campos").innerHTML=h+"</table></div>"}
document.getElementById("campos").addEventListener("input",e=>{const c=e.target.dataset.c;if(!c)return;const v=parseFloat(e.target.value.replace(/\./g,"").replace(",","."));const p=c.split("|");
  if(p[0]==="prazo"){(P.prazo=P.prazo||{})[p[1]]=isNaN(v)?12:v}
  else if(p[0]==="dir"){const [,mod,m,i]=p;P.dir=P.dir||{};P.dir[mod]=P.dir[mod]||{};const atual=P.dir[mod][m]||[g(D.dir,mod,m)?D.dir[mod][m][0]/1e6:0,g(D.dir,mod,m)?D.dir[mod][m][1]/1e6:0];atual[+i]=isNaN(v)?0:v;P.dir[mod][m]=atual}
  else{const [t,x,m]=p;P[t]=P[t]||{};P[t][x]=P[t][x]||{};if(isNaN(v))delete P[t][x][m];else P[t][x][m]=v}
  render()});
document.getElementById("salvar").onclick=()=>{const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([JSON.stringify(P,null,1)],{type:"application/json"}));a.download="parametros.json";a.click()};
document.getElementById("carregar").onchange=e=>{const f=e.target.files[0];if(!f)return;f.text().then(t=>{P=JSON.parse(t);montarCampos();render()})};
document.getElementById("restaurar").onclick=()=>{P=JSON.parse(JSON.stringify(D.params||{}));montarCampos();render()};
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>{document.querySelectorAll("nav button,section").forEach(x=>x.classList.remove("ativo"));b.classList.add("ativo");document.getElementById(b.dataset.s).classList.add("ativo")});
const ok=conferir();
document.getElementById("meta").innerHTML=`Data-base ${D.data_base.split("-").reverse().join("/")} · rodada ${D.rodada} · v${D.versao}<span class="badge ${ok?"ok":"falhou"}">CHECK simulação = Python: ${ok?"OK":"FALHOU"}</span>`
  +(D.avisos.length?`<div class="avisos">${D.avisos.map(a=>"AVISO: "+a).join("<br>")}</div>`:"");
montarCampos();render();
</script></body></html>"""


def html(dados):
    return HTML_MODELO.replace("__DADOS__", json.dumps(dados, ensure_ascii=False).replace("</", "<\\/"))


# ---------------------------------------------------------------- exemplo e autoteste (dados inventados)

def cmd_exemplo(data_base="2026-11-02"):
    ex = PASTA_SCRIPT / "exemplo"
    ex.mkdir(exist_ok=True)
    cats = [("RO", "pronaf", 3.0, "1,37"), ("RO", "pronaf_ind_com", 4.0, ""), ("RO", "pronamp", 8.0, ""),
            ("RO", "pronamp_invest", 8.0, ""), ("RO", "demais", 10.0, ""), ("RO", "demais_abc", 8.5, ""),
            ("POUP", "credito_rural", 11.0, ""), ("POUP", "cpr", 12.0, ""), ("LCA", "credito_rural", 12.0, ""),
            ("LCA", "faculdade", 13.0, ""), ("LIVRE", "livre", 15.0, "")]
    tam = {"pronaf": 30e6, "pronamp": 45e6, "demais": 14e6, "credito_rural": 42e6}
    ops, parc = [], []
    for k in range(60):
        fo, ca, tx, fat = cats[k % len(cats)]
        s = tam.get(ca, 2e6) * (1 + (k * 37 % 11) / 10)
        i = f"OP{k + 1:03d}"
        ops.append([i, fo, ca, s, tx, fat])
        venc = meses("2027-02", 16)[k * 7 % 16]
        if ca in ("pronamp_invest", "demais_abc"):  # investimento: parcelas anuais
            parc += [[i, f"{venc}-10", s / 3 * 1.1], [i, f"{int(venc[:4]) + 1}{venc[4:]}-10", s / 3 * 1.1],
                     [i, f"{int(venc[:4]) + 2}{venc[4:]}-10", s / 3 * 1.1]]
        else:
            parc.append([i, f"{venc}-15", s * 1.08])
    gravar(ex / "operacoes.csv", ["id", "fonte", "categoria", "saldo", "taxa_aa", "fator"], ops)
    gravar(ex / "liberacoes.csv", ["id", "data", "valor", "fonte", "categoria", "taxa_aa", "fator"],
           [["OP003", "2026-12-01", 2e6, "", "", "", ""],
            ["NOVA01", "2027-01-11", 30e6, "RO", "pronamp", 8.0, ""], ["NOVA02", "2027-03-01", 25e6, "POUP", "credito_rural", 11.0, ""]])
    gravar(ex / "parcelas.csv", ["id", "data", "valor"], parc + [["NOVA01", "2027-12-15", 32e6], ["NOVA02", "2028-02-15", 27e6]])
    gravar(ex / "cenario.csv", ["tipo", "id", "data", "fonte", "categoria"],
           [["liquidacao", "OP009", "2027-01-20", "", ""], ["reclassificacao", "OP005", "2027-02-10", "RO", "demais"]])
    gravar(ex / "feriados.csv", ["data"], [["2026-11-20"], ["2026-12-25"], ["2027-01-01"], ["2027-04-21"], ["2027-12-25"]])
    tot = defaultdict(float)
    for i, fo, ca, s, tx, fat in ops:
        if fo != "LIVRE":
            tot[(fo, ca, fat)] += s
    gravar(ex / "realizado.csv", ["mes", "fonte", "categoria", "saldo_medio", "fator"],
           [[m, fo, ca, v * (0.85 + 0.04 * j), fat] for j, m in enumerate(meses("2026-07", 4)) for (fo, ca, fat), v in tot.items()])
    gravar(ex / "vsr.csv", ["mes", "vsr_avista", "vsr_poupanca", "tipo"],
           [[m, 2.2e9 + j * 4e6, 2.4e8 + j * 1e6, "R" if m < "2026-11" else "P"] for j, m in enumerate(meses("2025-07", 36))])
    # LCA: saldo diário realizado (coluna data) e, para os meses projetados, uma linha por mês (coluna mes)
    lca = [[d.isoformat(), "", 6.5e8 + (d - date(2026, 6, 1)).days * 2.5e5] for d in
           (date(2026, 6, 1) + timedelta(days=k) for k in range(154)) if d.weekday() < 5]
    lca += [["", m, 6.9e8 + j * 1.2e7] for j, m in enumerate(meses("2026-11", 20))]
    gravar(ex / "lca.txt", ["data", "mes", "saldo"], lca)
    gravar(ex / "dir.csv", ["mes", "modalidade", "captacao", "aplicacao"],
           [[m, "pronaf", 0.0, 5e6] for m in meses("2026-07", 24)] + [[m, "geral", 3e6, 0.0] for m in meses("2026-07", 24)])
    (ex / "parametros.json").write_text(json.dumps({"prazo": {"RO_PRONAF": 10}, "liberacao": {"RO_PRONAF": {"2027-07": 20}},
                                                    "pct": {"RO_PRONAMP": {"2027-03": 60}}}, indent=1), encoding="utf-8")
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
    apl["RO|pronamp_invest"] = serie(20.0)
    r = sisex(2026, "2026-09", apl, {}, base, du)["RO"]
    check("investimento Pronamp limitado a 10% da sub", perto(r["ap_pronamp"], 55))
    r = sisex(2026, "2026-09", apl, {"geral": serie([30.0, 0.0]), "pronaf": serie([0.0, 8.0])}, base, du)["RO"]
    check("captação DIR soma na exigibilidade e aplicação DIR cumpre", perto(r["exig_total"], 130) and perto(r["ap_pronaf"], 53))
    po = sisex(2026, "2026-09", {"POUP|credito_rural": serie(90.0), "POUP|cpr": serie(20.0)}, {}, base, du)["POUP"]
    check("Poupança: CPR até 5% e crédito rural mínimo de 95%", perto(po["ap_total"], 95) and perto(po["def_sub"], 5))
    apl2 = {"RO|demais": serie(30.0), "RO|pronaf": serie(35.0), "RO|pronamp": serie(50.0)}
    du12 = {m: 20 for m in meses("2026-07", 12)}
    lc_mai = sisex(2026, "2027-05", apl2, {}, base, du12)["LCA"]
    lc_jun = sisex(2026, "2027-06", apl2, {}, base, du12)["LCA"]
    check("LCA: excesso de RO só conta na posição de junho", perto(lc_mai["exc_ro"], 0) and perto(lc_jun["exc_ro"], 15))
    check("média acumulada de julho até o mês, pelos dias úteis",
          perto(acum({"2026-07": 10.0, "2026-08": 40.0}, 2026, "2026-08", {"2026-07": 10, "2026-08": 30}), 32.5))

    # projeção
    op = lambda: {"A": {"fonte": "POUP", "categoria": "credito_rural", "saldo": 100.0, "taxa_aa": 0.0, "fator": 1.0}}
    res, pagos = projetar(op(), [("A", date(2026, 12, 15), 100.0)], [], [], date(2026, 11, 2), date(2027, 1, 31))
    du_dez = dias_uteis("2026-12", set())
    uteis_ate_14 = sum(1 for k in range(1, 15) if date(2026, 12, k).weekday() < 5)
    check("saldo zero depois da liquidação (Doc 6, Anexo I, 4.6-b)",
          perto(res[("A", "2026-12")][2], 100.0 * uteis_ate_14 / du_dez) and ("A", "2027-01") not in res)
    check("pagamento entra no total pago do grupo", perto(pagos[("POUP", "2026-12")], 100.0))
    cen = [{"tipo": "reclassificacao", "id": "A", "data": date(2026, 12, 10), "fonte": "RO", "categoria": "demais"},
           {"tipo": "reclassificacao", "id": "A", "data": date(2027, 1, 20), "fonte": "LCA", "categoria": "credito_rural"}]
    res, _ = projetar(op(), [], [], cen, date(2026, 11, 2), date(2027, 2, 28))
    check("fonte do dia 15 vale para o mês inteiro (Doc 6, Anexo I, 12.2-b)",
          res[("A", "2026-12")][:2] == ["RO", "demais"] and res[("A", "2027-01")][:2] == ["RO", "demais"]
          and res[("A", "2027-02")][:2] == ["LCA", "credito_rural"])
    libs = [("N", date(2027, 1, 18), 50.0, {"fonte": "RO", "categoria": "pronaf", "taxa_aa": 0.0, "fator": 1.0})]
    res, _ = projetar({}, [], libs, [], date(2026, 11, 2), date(2027, 2, 28))
    uteis_jan_apos = sum(1 for k in range(18, 32) if date(2027, 1, k).weekday() < 5)
    check("liberação futura: zero antes da liberação",
          perto(res[("N", "2027-01")][2], 50.0 * uteis_jan_apos / dias_uteis("2027-01", set()))
          and perto(res[("N", "2027-02")][2], 50.0))
    res, _ = projetar({"A": dict(op()["A"], taxa_aa=10.0)}, [], [], [], date(2026, 1, 1), date(2026, 12, 31))
    check("juros sobem o saldo cerca de 10% ao ano", 100 < res[("A", "2026-12")][2] < 110.1)

    # parâmetros da gerência
    M = meses("2026-07", 6)
    out = com_parametros({}, {"RO_PRONAF": {"2026-08": 10.0}},
                         {"prazo": {"RO_PRONAF": 2}, "liberacao": {"RO_PRONAF": {"2026-07": 1}},
                          "pct": {"RO_PRONAF": {"2026-08": 60}}}, M, "2026-07")["RO|pronaf"]
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
