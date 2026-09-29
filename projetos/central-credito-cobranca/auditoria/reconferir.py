#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Reconferencia independente da Central
=====================================

Recalcula os numeros principais DIRETO da Base_bruta.xlsx, por um caminho diferente
do central.py, e compara com o que esta ESCRITO no central.html:

  1. base contra tela  -- totais, vencido, saldo e semana batem com a planilha?
  2. tela contra tela  -- o mesmo cliente aparece com o mesmo valor em todo lugar?

E a prova de que a Central "continua fechando" quando a planilha muda. O conferir.py
na raiz faz outra coisa: ele compara com um retrato dos numeros de uma base
especifica, entao ACUSA DIFERENCA de proposito quando a base e trocada.

Diferencas deliberadas em relacao ao central.py:
  - le a planilha com openpyxl, nao com o leitor proprio do central.py
  - NAO importa central.py nem redacao.py
  - le a tela a partir do HTML gerado, nao dos objetos Python

Uso:
  python auditoria/reconferir.py

Precisa do openpyxl (pip install openpyxl). A Central em si NAO precisa: se a
maquina nao tiver, este script avisa e sai sem conferir nada.

Saida: 0 = tudo fechou | 1 = achou divergencia | 2 = nao conseguiu rodar
"""

import datetime as dt
import html as H
import os
import re
import statistics
import sys
from collections import defaultdict

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    from openpyxl import load_workbook
except ImportError:
    print("[PULADO] o openpyxl nao esta instalado nesta maquina.")
    print("         A Central funciona sem ele; so esta reconferencia precisa.")
    print("         Para rodar: pip install openpyxl")
    sys.exit(2)

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

PASTA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(PASTA, "Base_bruta.xlsx")
PAGINA = os.path.join(PASTA, "central.html")

MESES = {"janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5,
         "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10,
         "novembro": 11, "dezembro": 12}

achados = []


def falha(titulo, detalhe=""):
    achados.append(titulo)
    print(f"  !! {titulo}")
    if detalhe:
        for linha in str(detalhe).split("\n"):
            print(f"       {linha}")


def reais(v):
    """reais inteiros -> 'R$ 1.234,00' (a planilha guarda reais sem centavo)"""
    inteiro = int(v)
    centavos = round((v - inteiro) * 100)
    return "R$ " + f"{inteiro:,}".replace(",", ".") + f",{centavos:02d}"


def valor(txt):
    """'R$ 126.780,00' -> 126780.0"""
    return float(txt.replace("R$", "").replace(".", "").replace(",", ".").strip())


def so_data(v):
    return v.date() if isinstance(v, dt.datetime) else v


def um_ano_antes(d):
    try:
        return d.replace(year=d.year - 1)
    except ValueError:  # 29 de fevereiro
        return d.replace(year=d.year - 1, day=28)


# =========================================================== ler a planilha
for caminho, nome in ((XLSX, "Base_bruta.xlsx"), (PAGINA, "central.html")):
    if not os.path.exists(caminho):
        print(f"[ERRO] nao encontrei {nome} na pasta do projeto.")
        sys.exit(2)

wb = load_workbook(XLSX, data_only=True)


def aba(nome):
    linhas = list(wb[nome].values)
    cab = [str(c).strip() if c is not None else "" for c in linhas[0]]
    return [dict(zip(cab, l)) for l in linhas[1:] if any(c is not None for c in l)]


REF = None
for linha in wb["Leia-me"].values:
    celulas = [str(c) for c in linha if c]
    if celulas and "REFER" in celulas[0].upper():
        for c in celulas[1:]:
            m = re.search(r"(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", c)
            if m:
                REF = dt.date(int(m.group(3)), MESES[m.group(2).lower()], int(m.group(1)))
if REF is None:
    print("[ERRO] nao achei a linha 'DATA DE REFERÊNCIA' na aba Leia-me.")
    sys.exit(2)

CL = {}
for c in aba("Clientes"):
    CL[int(c["Código do cliente"])] = {"nome": c["Nome do cliente"], "limite": int(c["Limite de crédito"])}

TT = []
for t in aba("Títulos"):
    TT.append({
        "cod": int(t["Código do cliente"]),
        "emissao": so_data(t["Data de emissão"]),
        "valor": int(t["Valor do título"]),   # reais inteiros na planilha
        "venc": so_data(t["Data de vencimento"]),
        "pgto": so_data(t["Data de pagamento"]) if t["Data de pagamento"] else None,
    })

pagina = open(PAGINA, encoding="utf-8").read()

print("RECONFERENCIA INDEPENDENTE DA CENTRAL")
print("=" * 70)
print(f"Base: {len(TT)} titulos, {len(CL)} clientes, referencia {REF:%d/%m/%Y}")

# a pagina tem que ter sido gerada com ESTA base. se alguem trocou a planilha e
# esqueceu de rodar o central.py, tudo abaixo diverge e o motivo nao e esse.
m = re.search(r"REF_ISO = '(\d{4}-\d{2}-\d{2})'", pagina)
if not m or m.group(1) != REF.isoformat():
    print()
    print(f"[ERRO] o central.html foi gerado com a data {m.group(1) if m else '?'}, "
          f"mas a planilha e de {REF.isoformat()}.")
    print("       Rode 'python central.py' antes de reconferir.")
    sys.exit(2)

# ============================================== 1. base contra tela
print("\n[1] BASE CONTRA TELA")

abertos = [t for t in TT if t["pgto"] is None]
vencidos = [t for t in abertos if t["venc"] < REF]
tot = {
    "vencido": sum(t["valor"] for t in vencidos),
    "a vencer": sum(t["valor"] for t in abertos if t["venc"] >= REF),
    "em aberto": sum(t["valor"] for t in abertos),
}
for rotulo, v in tot.items():
    ok = reais(v) in pagina
    print(f"  total {rotulo:9s} {reais(v):>16s}  {'ok' if ok else 'NAO APARECE NA TELA'}")
    if not ok:
        falha(f"total {rotulo} da planilha nao aparece na tela", reais(v))

# numeros do alto da aba Hoje: cada total vem marcado com data-total="<qual>"
cartoes = dict(re.findall(r'data-total="([^"]+)">([^<]+)<', pagina))
for chave, esp in (("vencido", tot["vencido"]), ("a vencer", tot["a vencer"]),
                   ("em aberto", tot["em aberto"])):
    if chave not in cartoes:
        falha(f"cartao de resumo '{chave}' sumiu da tela")
    elif abs(valor(cartoes[chave]) - esp) > 0.005:
        falha(f"cartao '{chave}' diverge da planilha", f"tela {cartoes[chave]} | planilha {reais(esp)}")

# por cliente
perfil = {}
for cod, c in CL.items():
    meus = [t for t in TT if t["cod"] == cod]
    ab = [t for t in meus if t["pgto"] is None]
    ve = [t for t in ab if t["venc"] < REF]
    pagos = [t for t in meus if t["pgto"]]

    # atraso medio fora dos meses sazonais (mes elevado em 2+ anos, >= max(5, desvio))
    por_mes = defaultdict(lambda: defaultdict(list))
    for t in pagos:
        por_mes[t["venc"].month][t["venc"].year].append((t["pgto"] - t["venc"]).days)
    sazonais = set()
    for mes, anos in por_mes.items():
        if len(anos) < 2:
            continue
        medias = [statistics.mean(v) for v in anos.values()]
        resto = [(t["pgto"] - t["venc"]).days for t in pagos if t["venc"].month != mes]
        if len(resto) < 3:
            continue
        mr = statistics.mean(resto)
        sr = statistics.pstdev(resto) if len(resto) > 1 else 0
        if all(x > mr for x in medias) and statistics.mean(medias) - mr >= max(5, sr):
            sazonais.add(mes)
    atr = [(t["pgto"] - t["venc"]).days for t in pagos if t["venc"].month not in sazonais]

    perfil[cod] = {
        "saldo": sum(t["valor"] for t in ab),
        "vencido": sum(t["valor"] for t in ve),
        "qtd_vencidos": len(ve),
        "qtd_abertos": len(ab),
        "nivel": statistics.mean(atr) if len(atr) >= 5 else None,
    }

com_vencido = [c for c, p in perfil.items() if p["qtd_vencidos"]]
print(f"  clientes com vencido: {len(com_vencido)} | com saldo zero: "
      f"{sum(1 for p in perfil.values() if p['saldo'] == 0)}")
for cod, p in perfil.items():
    if p["saldo"] == 0:
        # numero redondo: so aceito se a planilha confirma que nao ha titulo em aberto
        print(f"  (saldo zero do cliente {cod} confere: {p['qtd_abertos']} titulos em aberto na planilha)")

# semana que acabou
ini = REF - dt.timedelta(days=6)
receb = sum(t["valor"] for t in TT if t["pgto"] and ini <= t["pgto"] <= REF)
nao_pago = sum(t["valor"] for t in TT if ini <= t["venc"] <= REF and t["pgto"] is None)
for rotulo, v in (("recebido na semana", receb), ("vencido e nao pago na semana", nao_pago)):
    ok = reais(v) in pagina
    print(f"  {rotulo:30s} {reais(v):>16s}  {'ok' if ok else 'NAO APARECE NA TELA'}")
    if not ok:
        falha(f"{rotulo} da planilha nao aparece na tela", reais(v))

# ============================================== 2. tela contra tela
print("\n[2] TELA CONTRA TELA")


def tela(nome):
    # cada tela e um <section class="tela"> ate a proxima tela ou o rodape
    m = re.search(rf'id="tela-{nome}"[^>]*>(.*?)(?=<section class="tela[^"]*" id="tela-|<footer class="rodape")',
                  pagina, re.S)
    if not m:
        falha(f"nao achei a tela {nome} na pagina -- a marcacao mudou? ajuste o regex de tela()")
    return m.group(1) if m else ""


hoje, previsao, regua = tela("hoje"), tela("previsao"), tela("regua")
fichas = {int(m.group(1)): m.group(0) for m in re.finditer(
    r'<div class="ficha ficha-modal" id="ficha-(\d+)".*?(?=<div class="ficha ficha-modal"|<script>)', pagina, re.S)}

# Hoje: uma linha por cliente com vencido, com o valor da planilha
linhas_hoje = {int(m.group(1)): (valor(m.group(2)), int(m.group(3))) for m in re.finditer(
    r'data-cod="(\d+)".*?<div class="valor">([^<]+?)\s*<small>vencido · (\d+)', hoje, re.S)}
if len(linhas_hoje) != len(com_vencido):
    falha("aba Hoje: nao consegui ler uma linha por cliente com vencido",
          f"li {len(linhas_hoje)}, a planilha tem {len(com_vencido)} -- se a marcacao mudou, ajuste o regex")
for cod, (v, q) in linhas_hoje.items():
    p = perfil[cod]
    if abs(v - p["vencido"]) > 0.005 or q != p["qtd_vencidos"]:
        falha(f"aba Hoje, cliente {cod}: vencido diverge da planilha",
              f"tela {reais(v)} em {q} | planilha {reais(p['vencido'])} em {p['qtd_vencidos']}")
print(f"  aba Hoje x planilha: {len(linhas_hoje)} clientes conferidos")

# fichas de Risco: saldo em aberto
if len(fichas) != len(CL):
    falha("fichas de Risco: nao encontrei uma por cliente", f"achei {len(fichas)}, a planilha tem {len(CL)}")
for cod, f in fichas.items():
    m = re.search(r'Saldo em aberto</div>\s*<div class="valor-topo">([^<]+)</div>', f)
    if not m:
        falha(f"ficha {cod}: nao achei o saldo em aberto")
    elif abs(valor(m.group(1)) - perfil[cod]["saldo"]) > 0.005:
        falha(f"ficha {cod}: saldo diverge da planilha",
              f"tela {m.group(1)} | planilha {reais(perfil[cod]['saldo'])}")
print(f"  fichas de Risco x planilha: {len(fichas)} saldos conferidos")

# o MESMO atraso medio na aba Hoje e na ficha de Risco (ja foram duas contas diferentes)
atraso_hoje = {int(m.group(1)): float(m.group(2).replace(",", ".")) for m in re.finditer(
    r'data-cod="(\d+)".*?Atraso médio histórico \(títulos pagos\)</td><td>([\d,.]+) dia', hoje, re.S)}
comparados = 0
for cod, h in atraso_hoje.items():
    m = re.search(r'Atraso médio depois do vencimento</div><div class="valor-fato">(\d+) dia', fichas.get(cod, ""))
    if not m:
        continue
    comparados += 1
    if abs(round(h) - int(m.group(1))) > 1:
        falha(f"cliente {cod}: atraso medio diferente entre Hoje e a ficha de Risco",
              f"Hoje {h} dias | ficha {m.group(1)} dias")
    nivel = perfil[cod]["nivel"]
    if nivel is not None and abs(h - nivel) > 0.06:
        falha(f"cliente {cod}: atraso medio da tela diverge do recalculado",
              f"tela {h} | recalculado {nivel:.2f} (fora dos meses sazonais)")
if atraso_hoje and not comparados:
    falha("nao consegui comparar o atraso medio entre Hoje e as fichas -- a marcacao mudou?")
print(f"  atraso medio Hoje x ficha x planilha: {comparados} clientes conferidos")

# Regua: o valor cobrado de cada cartao
cartoes_regua = re.findall(r'id="regua-(\d+)" data-cod="\d+"\s+data-tom-sugerido="[^"]+" data-valor="(\d+)"\s+data-qtd="(\d+)"',
                           regua)
if len(cartoes_regua) < len(com_vencido):
    falha("aba Regua: menos cartoes que clientes com vencido",
          f"{len(cartoes_regua)} cartoes | {len(com_vencido)} clientes com vencido")
for cod, v, q in cartoes_regua:
    p = perfil[int(cod)]
    esp = p["vencido"] if p["qtd_vencidos"] else p["saldo"]
    esp_q = p["qtd_vencidos"] if p["qtd_vencidos"] else p["qtd_abertos"]
    if abs(int(v) / 100 - esp) > 0.005 or int(q) != esp_q:
        falha(f"aba Regua, cliente {cod}: valor diverge da planilha",
              f"tela {reais(int(v) / 100)} em {q} | planilha {reais(esp)} em {esp_q}")
print(f"  aba Regua x planilha: {len(cartoes_regua)} cartoes conferidos")

# Previsao: todo titulo em aberto aparece -- menos os de cliente sem historico
# (menos de 5 titulos pagos fora de mes sazonal), que ficam fora e viram aviso
sem_hist = [t for t in abertos if perfil[t["cod"]]["nivel"] is None]
projetaveis = [t for t in abertos if perfil[t["cod"]]["nivel"] is not None]
tit_prev = re.findall(r'<td class="num">AUR-[^<]+</td><td>[^<]+</td><td class="dir">([^<]+)</td>', previsao)
if len(tit_prev) != len(projetaveis):
    falha("aba Previsao: nao lista todos os titulos que dava para projetar",
          f"lista {len(tit_prev)} | planilha tem {len(projetaveis)} (fora {len(sem_hist)} de cliente sem historico)")
elif abs(sum(valor(v) for v in tit_prev) - sum(t["valor"] for t in projetaveis)) > 0.005:
    falha("aba Previsao: a soma dos titulos listados nao da o saldo projetavel da planilha")
if sem_hist:
    total_fora = sum(t["valor"] for t in sem_hist)
    if reais(total_fora) not in previsao:
        falha("aba Previsao: titulos de cliente sem historico ficaram fora sem o aviso com o valor certo",
              f"esperado aviso com {reais(total_fora)}")
    print(f"  aba Previsao: {len(sem_hist)} titulo(s) de cliente sem historico fora da conta, com aviso")
semanas = re.findall(r'Semana \d</span>.*?<td class="dir">R\$ ([\d.]+)</td>\s*<td class="dir">R\$ ([\d.]+)</td>', previsao, re.S)
total = re.search(r'Total nas \d+ semanas</td>\s*<td class="dir">R\$ ([\d.]+)</td>\s*<td class="dir">R\$ ([\d.]+)</td>', previsao)
if not semanas or not total:
    falha("aba Previsao: nao consegui ler a tabela semanal -- a marcacao mudou?")
else:
    for i, rot in ((0, "planilha"), (1, "esperado")):
        soma = sum(int(s[i].replace(".", "")) for s in semanas)
        if soma != int(total.group(i + 1).replace(".", "")):
            falha(f"aba Previsao: a soma das semanas ({rot}) nao da o total",
                  f"soma {soma} | total {total.group(i + 1)}")
print(f"  aba Previsao: {len(tit_prev)} titulos e {len(semanas)} semanas conferidos")

# Relatorio: le os mesmos numeros das telas
rel = re.search(r'<script id="texto-relatorio" type="text/plain">(.*?)</script>', pagina, re.S)
if not rel:
    falha("o texto do relatorio da semana sumiu da pagina")
else:
    texto = H.unescape(rel.group(1))
    for rotulo, v in (("carteira vencida", tot["vencido"]), ("saldo em aberto", tot["em aberto"]),
                      ("recebido na semana", receb)):
        if reais(v) not in texto:
            falha(f"relatorio nao traz o {rotulo} da planilha", reais(v))
    print("  relatorio x planilha: conferido")

# ============================================== fim
print("\n" + "=" * 70)
if achados:
    print(f"RESULTADO: {len(achados)} divergencia(s). A tela NAO fecha com a planilha.")
    print("Antes de consertar, descubra qual dos dois esta errado: a conta da Central")
    print("ou esta reconferencia. Os dois caminhos foram escritos separados de proposito.")
    sys.exit(1)
print("RESULTADO: a tela fecha com a planilha, e as telas fecham entre si.")
sys.exit(0)
