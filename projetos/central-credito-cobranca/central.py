#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Central de Credito e Cobranca - Distribuidora Aurora
======================================================

Le Base_bruta.xlsx (somente leitura), calcula a fila de cobranca de hoje
e abre o resultado no navegador padrao como uma pagina HTML.

Nao usa nenhuma biblioteca externa: so a biblioteca padrao do Python
(zipfile + xml.etree para ler o .xlsx, o resto e Python puro). Isso e
proposital -- o objetivo e rodar numa maquina limpa, sem "pip install"
de nada, sem internet e sem conta.

Como rodar:
    python central.py
(ou de dois cliques em "Abrir Central.bat", no Windows)
"""

import base64
import csv
import datetime as dt
import glob
import html
import json
import math
import os
import pathlib
import re
import statistics
import sys
import unicodedata
import webbrowser
import zipfile
from xml.etree import ElementTree as ET

import redacao

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

PASTA = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_BASE = os.path.join(PASTA, "Base_bruta.xlsx")
ARQUIVO_SAIDA = os.path.join(PASTA, "central.html")
# as cobrancas que VOCE registrou moram aqui, fora da planilha da empresa.
# Base_bruta.xlsx e so leitura, sempre.
# CSV de proposito: abre no Excel, entao da para conferir o que a Central vai ler.
ARQUIVO_REGISTROS = os.path.join(PASTA, "cobrancas_registradas.csv")

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

MESES_PT = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9,
    "outubro": 10, "novembro": 11, "dezembro": 12,
}
MESES_PT_INV = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio",
    6: "junho", 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro",
    11: "novembro", 12: "dezembro",
}
DIAS_SEMANA_PT = [
    "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira",
    "sexta-feira", "sábado", "domingo",
]

EXCEL_EPOCH = dt.datetime(1899, 12, 30)  # compensa o bug do ano bissexto de 1900 do Excel

BUILTIN_DATE_FMT_IDS = {14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57}


# ---------------------------------------------------------------------------
# Leitor de .xlsx usando so a biblioteca padrao (sem pandas / openpyxl)
# ---------------------------------------------------------------------------

def col_letters_to_index(letters):
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def excel_serial_to_datetime(serial):
    return EXCEL_EPOCH + dt.timedelta(days=serial)


class LeitorXlsx:
    """Le as abas de um .xlsx simples (sem formulas, sem merges complexos)."""

    def __init__(self, caminho):
        self.zf = zipfile.ZipFile(caminho)
        self._carregar_workbook()
        self._carregar_shared_strings()
        self._carregar_estilos()

    def _carregar_workbook(self):
        wb_xml = ET.fromstring(self.zf.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(self.zf.read("xl/_rels/workbook.xml.rels"))
        rid_para_target = {}
        for rel in rels_xml:
            rid_para_target[rel.get("Id")] = rel.get("Target")
        self.planilhas = {}  # nome da aba -> caminho do sheetN.xml dentro do zip
        for sheet in wb_xml.find(f"{NS}sheets"):
            nome = sheet.get("name")
            rid = sheet.get(f"{NS_R}id")
            target = rid_para_target[rid]
            # o formato .xlsx aceita os dois jeitos de apontar para a aba: relativo
            # a pasta xl/ ("worksheets/sheet1.xml", como o Excel grava) ou absoluto
            # a partir da raiz ("/xl/worksheets/sheet1.xml", como outras ferramentas
            # gravam). Antes so o primeiro funcionava, e a Central nao abria uma
            # planilha que tivesse passado pelo LibreOffice, por exemplo.
            if target.startswith("/"):
                target = target.lstrip("/")
            elif not target.startswith("xl/"):
                target = "xl/" + target
            self.planilhas[nome] = target

    def _carregar_shared_strings(self):
        self.shared_strings = []
        if "xl/sharedStrings.xml" not in self.zf.namelist():
            return
        root = ET.fromstring(self.zf.read("xl/sharedStrings.xml"))
        for si in root:
            textos = si.findall(f".//{NS}t")
            self.shared_strings.append("".join(t.text or "" for t in textos))

    def _carregar_estilos(self):
        self.estilo_e_data = {}  # indice de estilo (s="") -> True/False (representa data?)
        if "xl/styles.xml" not in self.zf.namelist():
            return
        root = ET.fromstring(self.zf.read("xl/styles.xml"))
        formatos_customizados = {}
        num_fmts = root.find(f"{NS}numFmts")
        if num_fmts is not None:
            for nf in num_fmts:
                formatos_customizados[int(nf.get("numFmtId"))] = nf.get("formatCode", "")

        def eh_formato_de_data(num_fmt_id):
            if num_fmt_id in BUILTIN_DATE_FMT_IDS:
                return True
            codigo = formatos_customizados.get(num_fmt_id, "")
            if not codigo or codigo in ("General", "@"):
                return False
            return bool(re.search(r"[dhysm]", codigo, re.IGNORECASE))

        cell_xfs = root.find(f"{NS}cellXfs")
        if cell_xfs is not None:
            for i, xf in enumerate(cell_xfs):
                num_fmt_id = int(xf.get("numFmtId", "0"))
                self.estilo_e_data[i] = eh_formato_de_data(num_fmt_id)

    def ler_aba_bruta(self, nome_aba):
        """Retorna a aba como lista de linhas, cada linha uma lista de valores por coluna
        (posicional, sem assumir cabecalho). Uso: abas de texto livre, como Leia-me."""
        caminho = self.planilhas[nome_aba]
        root = ET.fromstring(self.zf.read(caminho))
        sheet_data = root.find(f"{NS}sheetData")
        linhas = []
        if sheet_data is None:
            return linhas
        for row_el in sheet_data.findall(f"{NS}row"):
            valores = {}
            for c in row_el.findall(f"{NS}c"):
                ref = c.get("r")
                letras = re.match(r"[A-Z]+", ref).group()
                col = col_letters_to_index(letras)
                valores[col] = self._ler_celula(c)
            if not valores:
                linhas.append([])
                continue
            max_col = max(valores.keys())
            linhas.append([valores.get(i) for i in range(max_col + 1)])
        return linhas

    def ler_aba(self, nome_aba):
        """Retorna uma lista de dicionarios: {nome_da_coluna: valor}, uma por linha (sem o cabecalho)."""
        caminho = self.planilhas[nome_aba]
        root = ET.fromstring(self.zf.read(caminho))
        sheet_data = root.find(f"{NS}sheetData")
        linhas_brutas = []
        if sheet_data is None:
            return []
        for row_el in sheet_data.findall(f"{NS}row"):
            valores = {}
            for c in row_el.findall(f"{NS}c"):
                ref = c.get("r")
                letras = re.match(r"[A-Z]+", ref).group()
                col = col_letters_to_index(letras)
                valores[col] = self._ler_celula(c)
            linhas_brutas.append(valores)
        if not linhas_brutas:
            return []

        cabecalho = linhas_brutas[0]
        max_col = max(
            (max(l.keys()) for l in linhas_brutas if l),
            default=-1,
        )
        nomes_colunas = [cabecalho.get(i) for i in range(max_col + 1)]

        registros = []
        for linha in linhas_brutas[1:]:
            if not linha:
                continue
            registro = {}
            for i, nome_col in enumerate(nomes_colunas):
                if nome_col is None:
                    continue
                registro[nome_col] = linha.get(i)
            registros.append(registro)
        return registros

    def _ler_celula(self, c):
        t = c.get("t")
        s = c.get("s")
        estilo_idx = int(s) if s is not None else 0

        if t == "inlineStr":
            is_el = c.find(f"{NS}is")
            if is_el is None:
                return None
            t_el = is_el.find(f"{NS}t")
            return t_el.text if t_el is not None else ""

        v_el = c.find(f"{NS}v")
        if t == "s":
            if v_el is None or v_el.text is None:
                return None
            return self.shared_strings[int(v_el.text)]

        if t == "str":
            return v_el.text if v_el is not None else ""

        if t == "b":
            return bool(int(v_el.text)) if v_el is not None else None

        # numero "puro" (pode ser data, se o estilo indicar formato de data)
        if v_el is None or v_el.text is None:
            return None
        numero = float(v_el.text)
        if self.estilo_e_data.get(estilo_idx, False):
            momento = excel_serial_to_datetime(numero)
            if abs(momento.hour) + momento.minute + momento.second == 0:
                return momento.date()
            return momento
        if numero == int(numero):
            return int(numero)
        return numero


# ---------------------------------------------------------------------------
# Utilitarios de data/formatacao
# ---------------------------------------------------------------------------

def subtrair_meses(data, meses):
    total = data.month - 1 - meses
    ano = data.year + total // 12
    mes = total % 12 + 1
    ultimo_dia = (dt.date(ano + (mes == 12), mes % 12 + 1, 1) - dt.timedelta(days=1)).day
    dia = min(data.day, ultimo_dia)
    return dt.date(ano, mes, dia)


def formatar_reais_de_centavos(centavos):
    negativo = centavos < 0
    centavos = abs(centavos)
    inteiros, centavos_resto = divmod(centavos, 100)
    texto = f"{inteiros:,}".replace(",", ".")
    resultado = f"R$ {texto},{centavos_resto:02d}"
    return ("-" + resultado) if negativo else resultado


def meses_de_historico(titulos, ref):
    """Quantos meses de calendario a base cobre, contando o primeiro e o ultimo,
    do mes da emissao mais antiga ao mes da data de referencia. E a mesma regra
    do 'Periodo coberto' do Leia-me (setembro/2024 a agosto/2026 = 24 meses).
    Antes estava escrito '24' a mao na tela -- com a base do mes seguinte, mentiria."""
    inicio = min(t["dt_emissao"] for t in titulos)
    return (ref.year - inicio.year) * 12 + (ref.month - inicio.month) + 1


def texto_dias(valor, casas=0, sinal=False):
    """'1 dia' / '18 dias' / '2,7 dias'. A concordancia e o separador decimal
    ficam num lugar so -- espalhados pelo codigo, uma hora sai '1 dias'."""
    if valor is None:
        return "—"
    if casas == 0:
        n = int(round(valor))
        num = f"{n:+d}" if sinal else f"{n}"
        return f"{num} {'dia' if abs(n) == 1 else 'dias'}"
    bruto = f"{valor:+.{casas}f}" if sinal else f"{valor:.{casas}f}"
    return bruto.replace(".", ",") + " dias"


def passos_do_eixo(minimo, maximo, alvo=5):
    """Valores redondos para as linhas de grade, em vez de fatias da faixa de dados.
    Eixo com '-8d, 5d, 17d, 29d' nao se le; com '-10, 0, 10, 20, 30' se le.
    Garante o zero quando a faixa cruza zero."""
    faixa = maximo - minimo
    if faixa <= 0:
        return [minimo]
    bruto = faixa / max(1, alvo - 1)
    magnitude = 10 ** math.floor(math.log10(bruto)) if bruto > 0 else 1
    for m in (1, 2, 2.5, 5, 10):
        passo = m * magnitude
        if passo >= bruto:
            break
    inicio = math.floor(minimo / passo) * passo
    valores, v = [], inicio
    while v <= maximo + passo * 0.001:
        valores.append(round(v, 6))
        v += passo
    return valores


def texto_pct(fracao, casas=1):
    """0.187 -> '18,7%'. Em portugues a casa decimal e virgula."""
    if fracao is None:
        return "—"
    return f"{fracao * 100:.{casas}f}".replace(".", ",") + "%"


def texto_decimal(valor, casas=1):
    """5.8 -> '5,8'. Numero quebrado em texto sai sempre com virgula: ja saiu
    '5.8 vezes o habitual' e '3.4 de 15 pts' na tela."""
    return f"{valor:.{casas}f}".replace(".", ",")


def texto_ordinal(posicao):
    """2.0 -> '2ª'; 2.5 -> '2,5ª'. O rank pode sair quebrado quando ha empate."""
    if posicao == int(posicao):
        return f"{int(posicao)}ª"
    return texto_decimal(posicao) + "ª"


def formatar_reais_estimado(centavos):
    """Estimativa nao tem centavo. Arredonda para o real mais proximo e some com
    a casa decimal. Vale so para numero projetado -- titulo fechado mantem centavo."""
    reais = int(round(centavos / 100))
    negativo = reais < 0
    texto = f"{abs(reais):,}".replace(",", ".")
    return ("-" if negativo else "") + f"R$ {texto}"


def formatar_data_pt(data):
    return f"{data.day:02d} de {MESES_PT_INV[data.month]} de {data.year}"


def formatar_data_curta(data):
    return f"{data.day:02d}/{data.month:02d}/{data.year}"


def dia_semana_pt(data):
    return DIAS_SEMANA_PT[data.weekday()]


def extrair_data_referencia(linhas_leiame):
    """Procura, na aba Leia-me, a linha 'DATA DE REFERENCIA' e le a data escrita ao lado."""
    padrao = re.compile(r"(\d{1,2})\s+de\s+([a-zçã]+)\s+de\s+(\d{4})", re.IGNORECASE)
    for linha in linhas_leiame:
        celulas = [v for v in linha if isinstance(v, str)]
        if not celulas:
            continue
        rotulo = celulas[0].strip().upper()
        if "REFER" in rotulo and "DATA" in rotulo:
            for texto in celulas[1:]:
                m = padrao.search(texto)
                if m:
                    dia = int(m.group(1))
                    mes = MESES_PT.get(m.group(2).lower())
                    ano = int(m.group(3))
                    if mes:
                        return dt.date(ano, mes, dia)
    raise RuntimeError("Nao encontrei a data de referencia na aba Leia-me.")


def rank_desc(valores):
    """Rank 1 = maior valor. Empates recebem a media das posicoes empatadas."""
    n = len(valores)
    ordem = sorted(range(n), key=lambda i: -valores[i])
    ranks = [0.0] * n
    i = 0
    while i < n:
        j = i
        while j + 1 < n and valores[ordem[j + 1]] == valores[ordem[i]]:
            j += 1
        media = (i + 1 + j + 1) / 2
        for k in range(i, j + 1):
            ranks[ordem[k]] = media
        i = j + 1
    return ranks


def media(lista):
    return statistics.mean(lista) if lista else None


def desvio_padrao(lista):
    if len(lista) > 1:
        return statistics.pstdev(lista)
    return 0.0


# ---------------------------------------------------------------------------
# Carga dos dados da planilha
# ---------------------------------------------------------------------------

def carregar_dados():
    leitor = LeitorXlsx(ARQUIVO_BASE)

    leiame = leitor.ler_aba_bruta("Leia-me")
    ref = extrair_data_referencia(leiame)

    clientes_raw = leitor.ler_aba("Clientes")
    clientes = {}
    for c in clientes_raw:
        cod = int(c["Código do cliente"])
        clientes[cod] = {
            "cod": cod,
            "nome": c["Nome do cliente"],
            "categoria": c["Categoria"],
            "regiao": c["Região"],
            "uf": c["UF"],
            "prazo_contratado": c["Prazo de pagamento contratado"],
            "prazo_dias": int(c["Prazo em dias"]),
            "limite_credito_centavos": int(c["Limite de crédito"]) * 100,
            "cliente_desde": c["Cliente desde"],
            # quem atende e em que numero -- sem isso a tela diz "ligue primeiro
            # para fulano" e deixa a pessoa procurando o telefone em outro lugar.
            "contato": c.get("Contato") or "",
            "telefone": c.get("Telefone") or "",
        }

    titulos_raw = leitor.ler_aba("Títulos")
    titulos = []
    for t in titulos_raw:
        titulos.append({
            "numero": t["Número do título"],
            "cod_cliente": int(t["Código do cliente"]),
            "dt_emissao": t["Data de emissão"],
            "valor_centavos": int(t["Valor do título"]) * 100,
            "dt_vencimento": t["Data de vencimento"],
            "dt_pagamento": t.get("Data de pagamento"),
        })

    cobrancas_raw = leitor.ler_aba("Cobranças")
    cobrancas = []
    for cb in cobrancas_raw:
        cobrancas.append({
            "cod_cliente": int(cb["Código do cliente"]),
            "dt_hora_contato": cb["Data e hora do contato"],
            "canal": cb["Canal"],
            "tom": cb["Tom usado"],
            "valor_cobrado_centavos": int(cb["Valor cobrado"]) * 100,
            "titulos_na_cobranca": int(cb["Títulos na cobrança"]),
        })

    cobrancas_registradas = carregar_registros()
    return ref, clientes, titulos, cobrancas + cobrancas_registradas


COLUNAS_REGISTRO = ["codigo_cliente", "cliente", "data", "hora", "canal", "tom", "valor", "titulos"]


def _ler_data_hora(data_txt, hora_txt):
    """Aceita mais de um formato de data porque o Excel reescreve a celula quando
    voce abre o CSV e salva por cima."""
    data_txt = (data_txt or "").strip()
    hora_txt = (hora_txt or "").strip() or "09:00"
    dia = None
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            dia = dt.datetime.strptime(data_txt, formato).date()
            break
        except ValueError:
            continue
    if dia is None:
        return None
    hora = dt.time(9, 0)
    for formato in ("%H:%M", "%H:%M:%S"):
        try:
            hora = dt.datetime.strptime(hora_txt, formato).time()
            break
        except ValueError:
            continue
    return dt.datetime.combine(dia, hora)


def _ler_centavos(texto):
    """'126.780,00' -> 12678000. Sem float: dinheiro continua sendo inteiro."""
    limpo = re.sub(r"[^\d,.-]", "", str(texto or "0"))
    if not limpo:
        return 0
    negativo = limpo.startswith("-")
    limpo = limpo.lstrip("-")
    if "," in limpo:
        inteiro, _, decimal = limpo.replace(".", "").partition(",")
    elif limpo.count(".") == 1 and len(limpo.split(".")[1]) == 2:
        inteiro, _, decimal = limpo.partition(".")   # 126780.00
    else:
        inteiro, decimal = limpo.replace(".", ""), "0"
    centavos = int(inteiro or 0) * 100 + int((decimal + "00")[:2])
    return -centavos if negativo else centavos


def registros_fora_do_lugar():
    """Arquivos que PARECEM o registro mas nao vao ser lidos, por causa do nome.
    O caso classico e o 'cobrancas_registradas (1).csv' que o navegador cria quando
    voce baixa de novo sem mover o primeiro -- e a Central ignorava isso calada."""
    esperado = os.path.abspath(ARQUIVO_REGISTROS)
    achados = []
    for caminho in sorted(glob.glob(os.path.join(PASTA, "cobranca*"))):
        if os.path.isdir(caminho) or os.path.abspath(caminho) == esperado:
            continue
        achados.append(os.path.basename(caminho))
    return achados


def carregar_registros():
    """Le as cobrancas que voce registrou pela tela, do CSV a parte.
    Se o arquivo nao existir, a Central roda igual -- so com o que veio da planilha."""
    if not os.path.exists(ARQUIVO_REGISTROS):
        return []
    try:
        # utf-8-sig engole o BOM que o Excel espera encontrar
        with open(ARQUIVO_REGISTROS, encoding="utf-8-sig", newline="") as f:
            linhas = list(csv.DictReader(f, delimiter=";"))
    except (OSError, csv.Error, UnicodeDecodeError):
        return []

    registros = []
    for r in linhas:
        momento = _ler_data_hora(r.get("data"), r.get("hora"))
        if momento is None:
            continue
        try:
            cod = int(str(r.get("codigo_cliente", "")).strip())
        except ValueError:
            continue
        registros.append({
            "cod_cliente": cod,
            "dt_hora_contato": momento,
            "canal": (r.get("canal") or "—").strip(),
            "tom": (r.get("tom") or "—").strip(),
            "valor_cobrado_centavos": _ler_centavos(r.get("valor")),
            "titulos_na_cobranca": int(re.sub(r"\D", "", str(r.get("titulos") or "0")) or 0),
            "registrado_na_tela": True,
        })
    return registros


# ---------------------------------------------------------------------------
# Calculo da fila
# ---------------------------------------------------------------------------

# Janela usada para julgar se um contato de cobranca e "recente" (dias corridos).
# A base so tem granularidade de DIA na data de referencia (sem hora), entao
# tratamos "48 horas" como "2 dias corridos" -- ver nota na conversa com o usuario.
JANELA_CONTATO_RECENTE_DIAS = 2

# Janela usada para comparar o atraso "recente" com o atraso "historico" de cada
# cliente e enxergar tendencia de piora ou melhora.
JANELA_TENDENCIA_MESES = 6


def calcular_fila(ref, clientes, titulos, cobrancas, perfis):
    """A fila de Hoje. O comportamento de cada cliente (quanto ele atrasa, o quanto
    varia, se piorou) vem PRONTO de calcular_perfis_risco -- esta tela nao recalcula
    isso por conta propria, senao duas telas passam a dizer numeros diferentes
    sobre o mesmo cliente."""
    titulos_por_cliente = {}
    for t in titulos:
        titulos_por_cliente.setdefault(t["cod_cliente"], []).append(t)

    ultimo_contato_por_cliente = {}
    for cb in cobrancas:
        atual = ultimo_contato_por_cliente.get(cb["cod_cliente"])
        if atual is None or cb["dt_hora_contato"] > atual["dt_hora_contato"]:
            ultimo_contato_por_cliente[cb["cod_cliente"]] = cb

    corte_tendencia = subtrair_meses(ref, JANELA_TENDENCIA_MESES)

    candidatos = []
    for cod, cliente in clientes.items():
        titulos_cliente = titulos_por_cliente.get(cod, [])

        # tudo isto vem do perfil, que e a unica fonte do comportamento do cliente.
        # antes esta tela calculava a propria media incluindo os meses sazonais,
        # e a tela de Risco calculava outra excluindo -- duas contas, mesmo nome.
        perfil = perfis[cod]
        pagos = perfil["pagos"]
        atraso_medio_historico = perfil["nivel"]
        desvio_atraso_historico = perfil["variabilidade"] or 0.0
        atraso_medio_recente = perfil["atraso_recente"]
        atraso_medio_antigo = perfil["atraso_antigo"]
        tendencia = perfil["diff_mudanca"] if perfil["diff_mudanca"] is not None else 0.0

        vencidos = [
            t for t in titulos_cliente
            if t["dt_pagamento"] is None and t["dt_vencimento"] < ref
        ]
        if not vencidos:
            continue  # so entra na fila quem tem dinheiro vencido hoje

        for t in vencidos:
            t["dias_atraso_atual"] = (ref - t["dt_vencimento"]).days
        vencidos.sort(key=lambda t: -t["dias_atraso_atual"])

        valor_vencido_centavos = sum(t["valor_centavos"] for t in vencidos)
        dias_atraso_max = vencidos[0]["dias_atraso_atual"]
        titulo_mais_antigo = vencidos[0]
        dias_atraso_medio_atual = media([t["dias_atraso_atual"] for t in vencidos])

        anomalia = (
            dias_atraso_medio_atual - atraso_medio_historico
            if atraso_medio_historico is not None else 0.0
        )
        risco_comportamental = max(0.0, anomalia) + max(0.0, tendencia)

        ultimo_contato = ultimo_contato_por_cliente.get(cod)
        dias_desde_contato = None
        contatado_recente = False
        if ultimo_contato is not None:
            dias_desde_contato = (ref - ultimo_contato["dt_hora_contato"].date()).days
            contatado_recente = dias_desde_contato <= JANELA_CONTATO_RECENTE_DIAS

        ultima_compra = max((t["dt_emissao"] for t in titulos_cliente), default=None)
        dias_sem_comprar = (ref - ultima_compra).days if ultima_compra else None

        candidatos.append({
            "cliente": cliente,
            "valor_vencido_centavos": valor_vencido_centavos,
            "qtd_vencidos": len(vencidos),
            "dias_atraso_max": dias_atraso_max,
            "titulo_mais_antigo": titulo_mais_antigo,
            "titulos_vencidos": vencidos,  # o que falar na ligacao: nota a nota
            "dias_atraso_medio_atual": dias_atraso_medio_atual,
            "atraso_medio_historico": atraso_medio_historico,
            "desvio_atraso_historico": desvio_atraso_historico,
            # conta os mesmos titulos que entraram na media, senao o rotulo
            # diz "com base em N" apontando para outro N
            "qtd_titulos_pagos_historico": perfil["qtd_pagos_ns"],
            "atraso_medio_recente": atraso_medio_recente,
            "atraso_medio_antigo": atraso_medio_antigo,
            # os dois julgamentos que a frase da fila faz vem do perfil, iguais aos
            # das abas Risco e Regua. Antes a fila julgava a piora pela propria conta
            # e dizia "padrao normal" de quem a aba Risco mostrava piorando.
            "piora_confirmada": perfil["piora_confirmada"],
            "fora_do_padrao_hoje": perfil["fora_do_padrao_hoje"],
            "tendencia": tendencia,
            "anomalia": anomalia,
            "risco_comportamental": risco_comportamental,
            "ultimo_contato": ultimo_contato,
            "dias_desde_contato": dias_desde_contato,
            "contatado_recente": contatado_recente,
            "dias_sem_comprar": dias_sem_comprar,
        })

    rank_valor = rank_desc([c["valor_vencido_centavos"] for c in candidatos])
    rank_urgencia = rank_desc([c["dias_atraso_max"] for c in candidatos])
    rank_risco = rank_desc([c["risco_comportamental"] for c in candidatos])
    for c, rv, ru, rr in zip(candidatos, rank_valor, rank_urgencia, rank_risco):
        c["rank_valor"] = rv
        c["rank_urgencia"] = ru
        c["rank_risco"] = rr
        c["rank_medio"] = (rv + ru + rr) / 3

    nao_recentes = sorted(
        (c for c in candidatos if not c["contatado_recente"]),
        key=lambda c: c["rank_medio"],
    )
    recentes = sorted(
        (c for c in candidatos if c["contatado_recente"]),
        key=lambda c: c["rank_medio"],
    )
    fila = nao_recentes + recentes
    for i, c in enumerate(fila, start=1):
        c["posicao"] = i

    return fila


# ---------------------------------------------------------------------------
# Frase explicativa de cada posicao na fila
# ---------------------------------------------------------------------------

def plural(qtd, singular, plural_):
    return singular if qtd == 1 else plural_


# O julgamento de cada cliente da fila, numa palavra. A frase e a etiqueta colorida
# do cartao leem daqui, entao nao tem como a cor dizer uma coisa e o texto outra.
# A cor tem um trabalho so em toda a Central: grave = fora do padrao dele,
# atencao = piorou, ok = e o padrao dele, neutro = nada a julgar agora.
SITUACOES_FILA = {
    "aguardando": ("Aguardando retorno do contato", "neutro"),
    "sem_historico": ("Sem histórico", "neutro"),
    "fora": ("Fora do padrão", "grave"),
    "piorou": ("Piorou", "atencao"),
    "padrao": ("Padrão dele", "ok"),
}


def situacao_na_fila(c):
    if c["contatado_recente"]:
        return "aguardando"
    if c["atraso_medio_historico"] is None:
        return "sem_historico"
    if c["fora_do_padrao_hoje"]:
        return "fora"
    if c["piora_confirmada"]:
        return "piorou"
    return "padrao"


def frase_padrao_historico(atraso_medio_historico):
    if atraso_medio_historico < -0.5:
        return f"paga em média {texto_dias(abs(atraso_medio_historico))} antes do vencimento"
    if atraso_medio_historico <= 0.5:
        return "paga em dia"
    return f"atrasa em média {texto_dias(atraso_medio_historico)}"


def gerar_explicacao(c, maior_valor_da_fila_cod):
    """Por que o cliente esta na fila, em duas ou tres linhas. O valor e a quantidade
    de titulos ja estao no cabecalho do cartao; os numeros por tras ficam na dobra."""
    H = c["atraso_medio_historico"]
    qual = "o título" if c["qtd_vencidos"] == 1 else "o mais antigo"
    idade = f"{qual} está vencido há {texto_dias(c['dias_atraso_max'])}."
    situacao = situacao_na_fila(c)

    partes = []
    if c["cliente"]["cod"] == maior_valor_da_fila_cod:
        partes.append("Maior valor vencido da fila.")

    if situacao == "aguardando":
        uc = c["ultimo_contato"]
        dd = c["dias_desde_contato"]
        quando = "hoje" if dd == 0 else f"há {texto_dias(dd)}"
        partes.append(
            f"Cobrado {quando} ({uc['dt_hora_contato']:%d/%m}, {uc['canal']}): desceu para o fim da fila, "
            f"e ligar de novo agora só irritaria."
        )
        partes.append(idade[0].upper() + idade[1:])
        if H is not None and not c["fora_do_padrao_hoje"]:
            partes.append(f"É o padrão dele: {frase_padrao_historico(H)} e sempre paga.")
    elif situacao == "sem_historico":
        partes.append(idade[0].upper() + idade[1:])
        partes.append("Ainda não há pagamentos suficientes para saber se isso foge do padrão dele.")
    elif situacao == "fora":
        partes.append(f"{frase_padrao_historico(H).capitalize()}; {idade}")
        if c["piora_confirmada"]:
            partes.append(
                f"E piorou: passou de {c['atraso_medio_antigo']:.0f} para {texto_dias(c['atraso_medio_recente'])} "
                f"nos últimos {JANELA_TENDENCIA_MESES} meses. Fora do padrão dele."
            )
        else:
            partes.append("Fora do padrão dele: vale checar se é pontual ou o começo de algo novo.")
    elif situacao == "piorou":
        partes.append(
            f"Piorou: o atraso passou de {c['atraso_medio_antigo']:.0f} para {texto_dias(c['atraso_medio_recente'])} "
            f"nos últimos {JANELA_TENDENCIA_MESES} meses, mudança confirmada."
        )
        partes.append(idade[0].upper() + idade[1:])
    else:
        partes.append(f"{frase_padrao_historico(H).capitalize()} e sempre paga; {idade}")
        partes.append("É o padrão dele: não é sinal novo de risco, só o valor pesa na fila.")

    return " ".join(partes)


# ---------------------------------------------------------------------------
# Tela "Risco por cliente": perfil de comportamento calculado sobre os 24
# meses de historico. Nada aqui esta escrito na planilha -- tudo e conta.
# ---------------------------------------------------------------------------

JANELA_FATURAMENTO_MESES = 12

# Pesos do indice de risco (somam 100). Cada um vira uma linha na tela, com
# a conta e a frase que justifica os pontos daquele cliente naquele peso.
PESO_NIVEL = 15
PESO_VARIABILIDADE = 20
PESO_MUDANCA = 30
PESO_COMPRA = 15
PESO_FATURAMENTO = 10
PESO_LIMITE = 10

MIN_TITULOS_HISTORICO = 5          # minimo de titulos pagos p/ falar de nivel/variabilidade
MIN_TITULOS_JANELA_MUDANCA = 3     # minimo de titulos pagos em cada janela p/ comparar tendencia
MIN_ANOS_SAZONALIDADE = 2          # um mes so conta como padrao de calendario se aparecer em 2+ anos
LIMIAR_CV_ESTAVEL = 0.35           # desvio-padrao/media do atraso; abaixo disso, o cliente e "previsivel"


def clip(valor, minimo, maximo):
    return max(minimo, min(maximo, valor))


def detectar_sazonalidade(pagos_cliente):
    """Meses do ano em que o atraso desse cliente e sistematicamente mais alto,
    em pelo menos MIN_ANOS_SAZONALIDADE anos diferentes -- e so nesse caso.
    Um ano ruim isolado nao conta; tem que se repetir."""
    por_mes = {}
    for t in pagos_cliente:
        por_mes.setdefault(t["dt_vencimento"].month, {}).setdefault(t["dt_vencimento"].year, []).append(t["atraso"])

    sazonais = {}
    for mes, anos in por_mes.items():
        medias_por_ano = {ano: media(vals) for ano, vals in anos.items()}
        if len(medias_por_ano) < MIN_ANOS_SAZONALIDADE:
            continue
        resto = [t["atraso"] for t in pagos_cliente if t["dt_vencimento"].month != mes]
        if len(resto) < MIN_TITULOS_JANELA_MUDANCA:
            continue
        media_resto = media(resto)
        std_resto = desvio_padrao(resto)
        media_mes = media(list(medias_por_ano.values()))
        elevado_em_todos_os_anos = all(v > media_resto for v in medias_por_ano.values())
        limiar = max(5.0, std_resto)
        if elevado_em_todos_os_anos and (media_mes - media_resto) >= limiar:
            sazonais[mes] = {
                "anos": medias_por_ano,
                "media_mes": media_mes,
                "media_resto": media_resto,
                "diff": media_mes - media_resto,
            }
    return sazonais


def calcular_perfis_risco(ref, clientes, titulos, cobrancas):
    titulos_por_cliente = {}
    for t in titulos:
        titulos_por_cliente.setdefault(t["cod_cliente"], []).append(t)

    contatos_por_cliente = {}
    for cb in cobrancas:
        contatos_por_cliente.setdefault(cb["cod_cliente"], []).append(cb)

    corte_tendencia = subtrair_meses(ref, JANELA_TENDENCIA_MESES)
    corte_faturamento = subtrair_meses(ref, JANELA_FATURAMENTO_MESES)

    faturamento_12m_por_cliente = {
        cod: sum(t["valor_centavos"] for t in lst if t["dt_emissao"] >= corte_faturamento)
        for cod, lst in titulos_por_cliente.items()
    }
    faturamento_12m_total = sum(faturamento_12m_por_cliente.values())

    brutos = {}
    for cod, cliente in clientes.items():
        titulos_cliente = titulos_por_cliente.get(cod, [])
        pagos = [t for t in titulos_cliente if t["dt_pagamento"] is not None]
        for t in pagos:
            t["atraso"] = (t["dt_pagamento"] - t["dt_vencimento"]).days

        sazonais = detectar_sazonalidade(pagos)
        meses_sazonais = set(sazonais.keys())
        pagos_ns = [t for t in pagos if t["dt_vencimento"].month not in meses_sazonais]

        if len(pagos_ns) >= MIN_TITULOS_HISTORICO:
            atrasos_ns = [t["atraso"] for t in pagos_ns]
            nivel = media(atrasos_ns)
            variabilidade = desvio_padrao(atrasos_ns)
            atraso_min, atraso_max = min(atrasos_ns), max(atrasos_ns)
            # quantos dias corridos ele leva da emissao ate o dinheiro entrar.
            # e o numero que serve para renegociar prazo -- "atraso medio" sozinho
            # e facil de ler errado como se fosse o prazo inteiro.
            prazo_efetivo = media([(t["dt_pagamento"] - t["dt_emissao"]).days for t in pagos_ns])
        else:
            nivel = variabilidade = atraso_min = atraso_max = prazo_efetivo = None

        recentes_ns = [t for t in pagos_ns if t["dt_vencimento"] >= corte_tendencia]
        antigos_ns = [t for t in pagos_ns if t["dt_vencimento"] < corte_tendencia]
        if len(recentes_ns) >= MIN_TITULOS_JANELA_MUDANCA and len(antigos_ns) >= MIN_TITULOS_JANELA_MUDANCA:
            atraso_recente = media([t["atraso"] for t in recentes_ns])
            atraso_antigo = media([t["atraso"] for t in antigos_ns])
            diff_mudanca = atraso_recente - atraso_antigo
            limiar_mudanca = max(5.0, variabilidade or 0.0)
            piora_confirmada = diff_mudanca >= limiar_mudanca
            melhora_confirmada = diff_mudanca <= -limiar_mudanca
        else:
            atraso_recente = atraso_antigo = diff_mudanca = limiar_mudanca = None
            piora_confirmada = melhora_confirmada = False

        # Qual atraso representa esse cliente HOJE. E a base da previsao de caixa:
        # quem mudou de comportamento de forma confirmada e projetado pelo que faz
        # agora, nao pela media de dois anos -- senao projetamos um cliente que nao
        # existe mais. Vale para os dois lados: quem piorou e quem melhorou.
        if (piora_confirmada or melhora_confirmada) and atraso_recente is not None:
            atraso_esperado = atraso_recente
            rumo = "piorou" if piora_confirmada else "melhorou"
            base_atraso_esperado = (
                f"comportamento dos últimos {JANELA_TENDENCIA_MESES} meses, porque ele {rumo} de forma confirmada"
            )
        elif nivel is not None:
            atraso_esperado = nivel
            base_atraso_esperado = "média histórica dele, fora dos meses sazonais"
        else:
            atraso_esperado = None
            base_atraso_esperado = "sem histórico suficiente para projetar"

        # o dia bom e o dia ruim dele: um desvio-padrao para cada lado
        if atraso_esperado is not None and variabilidade is not None:
            atraso_bom = atraso_esperado - variabilidade
            atraso_ruim = atraso_esperado + variabilidade
        else:
            atraso_bom = atraso_ruim = atraso_esperado

        serie_mensal = {}
        for t in pagos:
            chave = (t["dt_vencimento"].year, t["dt_vencimento"].month)
            serie_mensal.setdefault(chave, []).append(t["atraso"])
        serie_mensal = {k: media(v) for k, v in sorted(serie_mensal.items())}

        datas_compra = sorted(t["dt_emissao"] for t in titulos_cliente)
        if len(datas_compra) >= 2:
            gaps = [(datas_compra[i + 1] - datas_compra[i]).days for i in range(len(datas_compra) - 1)]
            intervalo_medio = media(gaps)
        else:
            intervalo_medio = None
        ultima_compra = datas_compra[-1] if datas_compra else None
        dias_sem_comprar = (ref - ultima_compra).days if ultima_compra else None

        recentes_compra = [t for t in titulos_cliente if t["dt_emissao"] >= corte_tendencia]
        antigos_compra = [t for t in titulos_cliente if t["dt_emissao"] < corte_tendencia]
        meses_antigo = ((corte_tendencia - datas_compra[0]).days / 30.44) if datas_compra and corte_tendencia > datas_compra[0] else 0
        taxa_recente = len(recentes_compra) / JANELA_TENDENCIA_MESES
        taxa_antiga = (len(antigos_compra) / meses_antigo) if meses_antigo > 1 else None
        ticket_recente = media([t["valor_centavos"] for t in recentes_compra]) if recentes_compra else None
        ticket_antigo = media([t["valor_centavos"] for t in antigos_compra]) if antigos_compra else None

        abertos_cliente = sorted(
            (t for t in titulos_cliente if t["dt_pagamento"] is None),
            key=lambda t: t["dt_vencimento"],
        )
        saldo_aberto_centavos = sum(t["valor_centavos"] for t in abertos_cliente)
        vencidos_cliente = [t for t in abertos_cliente if t["dt_vencimento"] < ref]
        tem_vencido_hoje = len(vencidos_cliente) > 0
        ocupacao_limite = (
            saldo_aberto_centavos / cliente["limite_credito_centavos"]
            if cliente["limite_credito_centavos"] else None
        )

        faturamento_12m = faturamento_12m_por_cliente.get(cod, 0)
        participacao_faturamento = (faturamento_12m / faturamento_12m_total) if faturamento_12m_total else 0.0

        # quanto o atraso dos titulos vencidos HOJE foge do padrao do proprio cliente.
        # fica aqui, junto do resto do comportamento, para a regua de cobranca nao
        # precisar recalcular isso por conta propria.
        # o atraso de hoje existe mesmo sem historico; o que precisa de historico
        # e a comparacao com o padrao dele
        atraso_medio_hoje = (
            media([(ref - t["dt_vencimento"]).days for t in vencidos_cliente]) if vencidos_cliente else None
        )
        if atraso_medio_hoje is not None and nivel is not None:
            anomalia_atual = atraso_medio_hoje - nivel
            fora_do_padrao_hoje = anomalia_atual > max(variabilidade or 0.0, 1.0)
        else:
            anomalia_atual = None
            fora_do_padrao_hoje = False

        # quantas vezes o intervalo normal de compra ja passou sem pedido
        razao_sem_comprar = (
            dias_sem_comprar / intervalo_medio
            if intervalo_medio and dias_sem_comprar is not None else None
        )

        brutos[cod] = dict(
            cliente=cliente, pagos=pagos, pagos_ns=pagos_ns, sazonais=sazonais,
            nivel=nivel, variabilidade=variabilidade, atraso_min=atraso_min, atraso_max=atraso_max,
            prazo_efetivo=prazo_efetivo,
            atraso_recente=atraso_recente, atraso_antigo=atraso_antigo, diff_mudanca=diff_mudanca,
            limiar_mudanca=limiar_mudanca, piora_confirmada=piora_confirmada, melhora_confirmada=melhora_confirmada,
            atraso_esperado=atraso_esperado, atraso_bom=atraso_bom, atraso_ruim=atraso_ruim,
            base_atraso_esperado=base_atraso_esperado,
            serie_mensal=serie_mensal, intervalo_medio=intervalo_medio, dias_sem_comprar=dias_sem_comprar,
            taxa_recente=taxa_recente, taxa_antiga=taxa_antiga, ticket_recente=ticket_recente, ticket_antigo=ticket_antigo,
            saldo_aberto_centavos=saldo_aberto_centavos, tem_vencido_hoje=tem_vencido_hoje,
            titulos_abertos=abertos_cliente, qtd_vencidos_hoje=len(vencidos_cliente),
            ultimo_contato=(
                max(contatos_por_cliente.get(cod, []), key=lambda cb: cb["dt_hora_contato"])
                if contatos_por_cliente.get(cod) else None
            ),
            ocupacao_limite=ocupacao_limite, faturamento_12m=faturamento_12m,
            participacao_faturamento=participacao_faturamento,
            atraso_medio_hoje=atraso_medio_hoje, anomalia_atual=anomalia_atual,
            fora_do_padrao_hoje=fora_do_padrao_hoje, razao_sem_comprar=razao_sem_comprar,
            foi_cobrado_algum_dia=len(contatos_por_cliente.get(cod, [])) > 0,
            qtd_pagos=len(pagos), qtd_pagos_ns=len(pagos_ns), corte_tendencia=corte_tendencia,
        )

    max_nivel = max((d["nivel"] for d in brutos.values() if d["nivel"]), default=None) or 1.0
    max_variab = max((d["variabilidade"] for d in brutos.values() if d["variabilidade"] is not None), default=None) or 1.0
    max_participacao = max((d["participacao_faturamento"] for d in brutos.values()), default=None) or 1.0
    niveis_validos = sorted(d["nivel"] for d in brutos.values() if d["nivel"] is not None)
    mediana_nivel = statistics.median(niveis_validos) if niveis_validos else 0.0

    perfis = {}
    for cod, d in brutos.items():
        cliente = d["cliente"]
        componentes = []

        if d["nivel"] is not None:
            pts = PESO_NIVEL * clip(max(0.0, d["nivel"]) / max_nivel, 0, 1)
            if d["nivel"] > 0.5:
                frase = (
                    f"Historicamente paga cerca de {texto_dias(d['nivel'])} após o vencimento, fora dos meses "
                    f"sazonais identificados — {pts:.0f} de {PESO_NIVEL} pontos porque é o "
                    f"{'maior' if d['nivel'] >= max_nivel else 'um dos maiores'} atraso estrutural da carteira."
                    if d["nivel"] >= max_nivel * 0.6 else
                    f"Historicamente paga cerca de {texto_dias(d['nivel'])} após o vencimento — abaixo da média da carteira."
                )
            else:
                frase = "Historicamente paga em dia ou adiantado — não pesa no índice."
        else:
            pts = 0.0
            frase = f"Histórico insuficiente (menos de {MIN_TITULOS_HISTORICO} títulos pagos fora dos meses sazonais) para medir o nível estrutural de atraso."
        componentes.append(dict(chave="nivel", nome="Nível de atraso histórico", pontos=round(pts, 1), peso=PESO_NIVEL, frase=frase))

        if d["variabilidade"] is not None:
            pts = PESO_VARIABILIDADE * clip(d["variabilidade"] / max_variab, 0, 1)
            if d["variabilidade"] >= max_variab * 0.5:
                frase = (
                    f"O atraso varia bastante: de {d['atraso_min']} a {texto_dias(d['atraso_max'])} historicamente "
                    f"(desvio-padrão de {texto_dias(d['variabilidade'])}) — imprevisível, mesmo quando a média parece controlada."
                )
            else:
                frase = (
                    f"É consistente: paga quase sempre na mesma faixa, entre {d['atraso_min']} e {texto_dias(d['atraso_max'])} "
                    f"(desvio-padrão de {texto_dias(d['variabilidade'])})."
                )
        else:
            pts = 0.0
            frase = "Histórico insuficiente para medir a variabilidade do atraso."
        componentes.append(dict(chave="variabilidade", nome="Variabilidade (imprevisibilidade)", pontos=round(pts, 1), peso=PESO_VARIABILIDADE, frase=frase))

        if d["diff_mudanca"] is not None:
            limiar = d["limiar_mudanca"]
            excesso = max(0.0, d["diff_mudanca"])
            pts = PESO_MUDANCA * clip(excesso / (3 * limiar), 0, 1)
            if d["piora_confirmada"]:
                frase = (
                    f"Mudança de comportamento confirmada: nos últimos {JANELA_TENDENCIA_MESES} meses o atraso médio foi "
                    f"de {d['atraso_recente']:.0f} dias, contra {d['atraso_antigo']:.0f} antes — {texto_dias(d['diff_mudanca'])} "
                    f"a mais, acima da variação normal dele (limiar de {texto_dias(limiar)})."
                )
            elif excesso > 0:
                frase = (
                    f"Sinal leve: o atraso médio foi de {d['atraso_antigo']:.0f} para {texto_dias(d['atraso_recente'])} nos "
                    f"últimos {JANELA_TENDENCIA_MESES} meses, mas ainda dentro da variação normal dele (limiar de {texto_dias(limiar)}) "
                    f"— não é uma mudança estatisticamente confirmada."
                )
            elif d["melhora_confirmada"]:
                frase = (
                    f"Está melhorando: o atraso médio caiu de {d['atraso_antigo']:.0f} para {texto_dias(d['atraso_recente'])} "
                    f"nos últimos {JANELA_TENDENCIA_MESES} meses."
                )
            else:
                frase = "Sem mudança relevante: o comportamento recente está em linha com o histórico dele."
        else:
            pts = 0.0
            frase = (
                f"Histórico insuficiente (menos de {MIN_TITULOS_JANELA_MUDANCA} títulos pagos em uma das janelas de "
                f"{JANELA_TENDENCIA_MESES} meses) para comparar e afirmar mudança de padrão."
            )
        componentes.append(dict(chave="mudanca", nome="Mudança recente de comportamento", pontos=round(pts, 1), peso=PESO_MUDANCA, frase=frase))

        pontos_compra = []
        frases_compra = []
        if d["intervalo_medio"] and d["dias_sem_comprar"] is not None:
            razao = d["dias_sem_comprar"] / d["intervalo_medio"]
            pontos_compra.append(PESO_COMPRA * clip(razao - 1.0, 0.0, 1.0))
            if razao >= 2.0:
                frases_compra.append(
                    f"parece ter parado de comprar: {texto_dias(d['dias_sem_comprar'])} desde o último pedido, quando o "
                    f"intervalo normal dele é de {texto_dias(d['intervalo_medio'])}"
                )
            elif razao >= 1.3:
                frases_compra.append(
                    f"está demorando mais que o normal para pedir de novo ({texto_dias(d['dias_sem_comprar'])}, ante um "
                    f"intervalo médio de {texto_dias(d['intervalo_medio'])})"
                )
        if d["ticket_recente"] is not None and d["ticket_antigo"]:
            var_ticket = (d["ticket_recente"] - d["ticket_antigo"]) / d["ticket_antigo"]
            pontos_compra.append(PESO_COMPRA * clip(-var_ticket / 0.30, 0.0, 1.0))
            if var_ticket <= -0.15:
                frases_compra.append(f"o pedido médio encolheu {texto_pct(abs(var_ticket), 0)} nos últimos {JANELA_TENDENCIA_MESES} meses")
        if d["taxa_antiga"]:
            var_freq = (d["taxa_recente"] - d["taxa_antiga"]) / d["taxa_antiga"]
            pontos_compra.append(PESO_COMPRA * clip(-var_freq / 0.30, 0.0, 1.0))
            if var_freq <= -0.15:
                frases_compra.append(
                    f"a frequência de pedidos caiu {texto_pct(abs(var_freq), 0)} (de {texto_decimal(d['taxa_antiga'])} para "
                    f"{texto_decimal(d['taxa_recente'])} títulos/mês)"
                )
        pts = max(pontos_compra) if pontos_compra else 0.0
        if frases_compra:
            frase = "Sinal de compra: " + "; ".join(frases_compra) + "."
        elif d["dias_sem_comprar"] is not None:
            frase = (
                f"Comprando normalmente: último pedido há {texto_dias(d['dias_sem_comprar'])}, dentro do intervalo e do "
                f"ritmo de compra habitual dele."
            )
        else:
            frase = "Sem histórico de compras suficiente para avaliar frequência ou tamanho médio do pedido."
        componentes.append(dict(chave="compra", nome="Sinal de compra (frequência e tamanho do pedido)", pontos=round(pts, 1), peso=PESO_COMPRA, frase=frase))

        pts = PESO_FATURAMENTO * clip(d["participacao_faturamento"] / max_participacao, 0, 1)
        frase = (
            f"Responde por {texto_pct(d['participacao_faturamento'])} do faturamento dos últimos "
            f"{JANELA_FATURAMENTO_MESES} meses ({formatar_reais_de_centavos(d['faturamento_12m'])}) — quanto maior a "
            f"fatia, maior o impacto se algo der errado com esse cliente."
        )
        componentes.append(dict(chave="faturamento", nome="Concentração no faturamento", pontos=round(pts, 1), peso=PESO_FATURAMENTO, frase=frase))

        if d["ocupacao_limite"] is not None:
            pts = PESO_LIMITE * clip(d["ocupacao_limite"], 0, 1)
            frase = (
                f"Está usando {texto_pct(d['ocupacao_limite'], 0)} do limite de crédito dele "
                f"({formatar_reais_de_centavos(d['saldo_aberto_centavos'])} de {formatar_reais_de_centavos(cliente['limite_credito_centavos'])})."
            )
            if d["ocupacao_limite"] > 1.0:
                frase += " Já está acima do limite contratado."
        else:
            pts = 0.0
            frase = "Sem limite de crédito cadastrado para calcular ocupação."
        componentes.append(dict(chave="limite", nome="Ocupação do limite de crédito", pontos=round(pts, 1), peso=PESO_LIMITE, frase=frase))

        indice = clip(sum(c["pontos"] for c in componentes), 0, 100)
        # valor em risco e estimativa (saldo x indice): reais inteiros, sem centavo
        valor_em_risco_centavos = round(d["saldo_aberto_centavos"] * (indice / 100) / 100) * 100

        cv_cliente = (d["variabilidade"] / d["nivel"]) if d["nivel"] and d["nivel"] > 0 and d["variabilidade"] is not None else None
        selo_cronico_estavel = (
            d["tem_vencido_hoje"] and d["diff_mudanca"] is not None and not d["piora_confirmada"]
            and d["nivel"] is not None and d["nivel"] >= mediana_nivel
            and cv_cliente is not None and cv_cliente <= LIMIAR_CV_ESTAVEL
        )
        selo_invisivel = (
            not d["tem_vencido_hoje"] and not d["foi_cobrado_algum_dia"]
            and d["diff_mudanca"] is not None and d["diff_mudanca"] > 0
        )

        perfis[cod] = dict(
            d, componentes=componentes, indice=round(indice, 1),
            valor_em_risco_centavos=valor_em_risco_centavos,
            selo_cronico_estavel=selo_cronico_estavel, selo_invisivel=selo_invisivel,
        )

    # entre os "invisiveis" (sem vencido hoje, nunca cobrados, tendencia de piora),
    # so o de maior valor em risco leva o selo -- o pedido foi por UM cliente com nome e sobrenome.
    invisiveis = [cod for cod, p in perfis.items() if p["selo_invisivel"]]
    if len(invisiveis) > 1:
        melhor = max(invisiveis, key=lambda cod: perfis[cod]["valor_em_risco_centavos"])
        for cod in invisiveis:
            if cod != melhor:
                perfis[cod]["selo_invisivel"] = False

    return perfis


# ---------------------------------------------------------------------------
# Tela "Previsao de caixa": quando o dinheiro da carteira atual deve entrar.
# Le o comportamento ja calculado em calcular_perfis_risco -- nao refaz conta.
# ---------------------------------------------------------------------------

SEMANAS_PREVISTAS = 4

# Quanto o indice de risco do cliente pode descontar da expectativa de entrada.
# Risco 100 tira 30% do valor; risco 0 nao tira nada.
PESO_RISCO_NA_ENTRADA = 0.30

# Piso da chance de entrada: por mais velho que esteja, nao zeramos o titulo.
PISO_CHANCE_ENTRADA = 0.25


def chance_de_entrar(p, dias_de_atraso_hoje):
    """Chance de o titulo entrar, entre 0 e 1.

    Cai por dois motivos: o risco do cliente (indice ja calculado) e a idade do
    titulo. A idade e medida contra o proprio historico dele -- um titulo so
    comeca a perder chance depois de passar do maior atraso que aquele cliente
    ja praticou. Cada cliente tem a propria regua."""
    fator_risco = 1 - (p["indice"] / 100) * PESO_RISCO_NA_ENTRADA

    referencia = max(p["atraso_max"] or 0, 7)
    excesso = max(0, dias_de_atraso_hoje - (p["atraso_max"] or 0))
    fator_idade = clip(1 - excesso / (2 * referencia), PISO_CHANCE_ENTRADA, 1.0)

    return fator_risco * fator_idade, fator_risco, fator_idade


def atraso_do_titulo(p, titulo, qual="esperado"):
    """Atraso a aplicar num titulo. Se ele vence num mes sazonal daquele cliente,
    vale o atraso daquele mes -- nao a media geral, que exclui justamente esses meses."""
    mes = titulo["dt_vencimento"].month
    if mes in p["sazonais"]:
        base = p["sazonais"][mes]["media_mes"]
        sazonal = True
    else:
        base = p["atraso_esperado"]
        sazonal = False
    if base is None:
        return None, sazonal
    if qual == "bom":
        return base - (p["variabilidade"] or 0), sazonal
    if qual == "ruim":
        return base + (p["variabilidade"] or 0), sazonal
    return base, sazonal


def calcular_previsao(ref, perfis):
    semanas = []
    for i in range(SEMANAS_PREVISTAS):
        ini = ref + dt.timedelta(days=7 * i)
        semanas.append({"ini": ini, "fim": ini + dt.timedelta(days=6)})
    fim_janela = semanas[-1]["fim"]

    def indice_da_semana(data):
        # o que ja deveria ter entrado cai na semana 1: atrasado nao e perdido,
        # e dinheiro que pode bater na conta a qualquer momento
        if data < ref:
            return 0
        for i, s in enumerate(semanas):
            if s["ini"] <= data <= s["fim"]:
                return i
        return None

    planilha = [0] * SEMANAS_PREVISTAS
    esperado = [0] * SEMANAS_PREVISTAS
    conservador = [0] * SEMANAS_PREVISTAS
    otimista = [0] * SEMANAS_PREVISTAS
    fora_da_janela = 0
    descontado = 0
    linhas = []
    sem_historico = []

    for cod in sorted(perfis):
        p = perfis[cod]
        for t in p["titulos_abertos"]:
            valor = t["valor_centavos"]
            dias_hoje = (ref - t["dt_vencimento"]).days
            chance, f_risco, f_idade = chance_de_entrar(p, dias_hoje)

            atraso_esp, sazonal = atraso_do_titulo(p, t, "esperado")
            if atraso_esp is None:
                # cliente sem historico: nao da para dizer quando ele paga, entao nao
                # se estima. O titulo sai dos DOIS lados (planilha e esperado) e vira
                # aviso na tela. Antes ele entrava so na planilha e inflava a diferenca.
                sem_historico.append({"cliente": p["cliente"], "titulo": t})
                continue

            i_planilha = indice_da_semana(t["dt_vencimento"])
            if i_planilha is not None:
                planilha[i_planilha] += valor

            data_esperada = t["dt_vencimento"] + dt.timedelta(days=round(atraso_esp))
            # estimativa nao tem centavo: arredonda aqui, titulo a titulo, para que
            # semanas e total somem exatamente o que a tela mostra. Arredondar so na
            # hora de exibir fazia a soma das semanas diferir do total em R$ 1.
            valor_ponderado = round(valor * chance / 100) * 100
            descontado += valor - valor_ponderado

            i_esp = indice_da_semana(data_esperada)
            if i_esp is None:
                fora_da_janela += valor_ponderado
            else:
                esperado[i_esp] += valor_ponderado

            for qual, destino in (("ruim", conservador), ("bom", otimista)):
                atraso, _ = atraso_do_titulo(p, t, qual)
                data = t["dt_vencimento"] + dt.timedelta(days=round(atraso))
                i = indice_da_semana(data)
                if i is not None:
                    destino[i] += valor_ponderado

            linhas.append({
                "cliente": p["cliente"],
                "titulo": t,
                "dias_atraso_hoje": dias_hoje,
                "atraso_projetado": atraso_esp,
                "data_esperada": data_esperada,
                "semana": i_esp,
                "chance": chance,
                "fator_risco": f_risco,
                "fator_idade": f_idade,
                "valor_ponderado": valor_ponderado,
                "sazonal": sazonal,
                "base": p["base_atraso_esperado"],
            })

    def acumular(serie):
        total, saida = 0, []
        for v in serie:
            total += v
            saida.append(total)
        return saida

    return {
        "semanas": semanas,
        "fim_janela": fim_janela,
        "planilha": planilha,
        "esperado": esperado,
        "conservador": conservador,
        "otimista": otimista,
        "acum_planilha": acumular(planilha),
        "acum_esperado": acumular(esperado),
        "acum_conservador": acumular(conservador),
        "acum_otimista": acumular(otimista),
        "total_planilha": sum(planilha),
        "total_esperado": sum(esperado),
        "total_conservador": sum(conservador),
        "total_otimista": sum(otimista),
        "diferenca": sum(planilha) - sum(esperado),
        "fora_da_janela": fora_da_janela,
        "descontado": descontado,
        "linhas": sorted(linhas, key=lambda l: (l["semana"] is None, l["semana"], -l["valor_ponderado"])),
        "semanas_invertidas": [
            i + 1 for i in range(SEMANAS_PREVISTAS) if conservador[i] > otimista[i]
        ],
        "sem_historico": sem_historico,
        "sem_historico_centavos": sum(x["titulo"]["valor_centavos"] for x in sem_historico),
    }


# ---------------------------------------------------------------------------
# Tela "Regua de cobranca": qual tom usar com cada cliente, e por que.
# O texto em si mora em redacao.py -- aqui so se decide O QUE dizer.
# ---------------------------------------------------------------------------

# a partir de quantas vezes o intervalo normal de compra consideramos que parou
RAZAO_PAROU_DE_COMPRAR = 2.0


def escolher_tom(p):
    """Devolve (tom, motivo). A ordem importa: o sinal comercial vem antes do
    financeiro, porque quem parou de comprar precisa de conversa antes de cobranca."""
    nome = p["cliente"]["nome"]

    if not p["tem_vencido_hoje"]:
        if p["selo_invisivel"]:
            return "contato_preventivo", (
                f"Não tem nada vencido e nunca foi cobrado, mas o atraso médio subiu de "
                f"{p['atraso_antigo']:.0f} para {texto_dias(p['atraso_recente'])} e o saldo em aberto é de "
                f"{formatar_reais_de_centavos(p['saldo_aberto_centavos'])}. Contato de relacionamento, não de cobrança."
            )
        return "contato_preventivo", (
            f"{nome} não tem título vencido hoje — qualquer contato aqui é de relacionamento, não cobrança."
        )

    if p["nivel"] is None:
        # sem historico nao ha padrao para comparar: nada de "atipico para ele" nem
        # de "sempre pagou assim". So os tons que nao falam do passado servem.
        return "cobranca_objetiva", (
            f"Cliente sem histórico suficiente: só {p['qtd_pagos_ns']} "
            f"{plural(p['qtd_pagos_ns'], 'título pago', 'títulos pagos')} até hoje (o mínimo para ler o padrão "
            f"é {MIN_TITULOS_HISTORICO}). Ainda não dá para saber se esse atraso é normal para ele — cobrança "
            f"direta e respeitosa, sem comparar com um passado que não existe."
        )

    if p["razao_sem_comprar"] and p["razao_sem_comprar"] >= RAZAO_PAROU_DE_COMPRAR:
        return "reaproximacao_comercial", (
            f"Está há {texto_dias(p['dias_sem_comprar'])} sem comprar, quando o intervalo normal dele é de "
            f"{texto_dias(p['intervalo_medio'])} — {texto_decimal(p['razao_sem_comprar'])} vezes o habitual. "
            f"Parar de comprar é sinal comercial: a conversa vem antes da cobrança."
        )

    if p["piora_confirmada"]:
        return "conversa_mudanca", (
            f"O atraso médio subiu de {p['atraso_antigo']:.0f} para {texto_dias(p['atraso_recente'])} nos últimos "
            f"{JANELA_TENDENCIA_MESES} meses — mudança confirmada, acima da variação normal dele. "
            f"Quem está piorando precisa de conversa, não de ameaça."
        )

    if p["selo_cronico_estavel"]:
        return "lembrete_leve", (
            f"Atrasa de forma muito consistente ({texto_dias(p['nivel'])} em média, variação de "
            f"± {p['variabilidade']:.0f}) e nunca falhou em {p['qtd_pagos_ns']} títulos pagos. "
            f"O atraso de hoje está dentro do padrão dele — lembrete leve, com proposta de acertar o prazo no contrato."
        )

    if p["fora_do_padrao_hoje"]:
        return "primeiro_aviso_cordial", (
            f"Historicamente atrasa {texto_dias(p['nivel'])}, e os títulos de hoje estão em média "
            f"{texto_dias(p['atraso_medio_hoje'])} vencidos — bem fora do padrão dele. "
            f"Cliente que costuma pagar bem merece uma pergunta antes de uma cobrança."
        )

    if p["anomalia_atual"] is not None and p["anomalia_atual"] <= 0:
        return "lembrete_leve", (
            f"Os títulos vencidos hoje estão em média {texto_dias(p['atraso_medio_hoje'])} em atraso, dentro do que "
            f"ele já pratica ({texto_dias(p['nivel'])} em média). Nada aqui indica problema novo."
        )

    media_txt = f"{p['atraso_medio_hoje']:.0f}" if p["atraso_medio_hoje"] is not None else "—"
    return "cobranca_objetiva", (
        f"{p['qtd_vencidos_hoje']} {plural(p['qtd_vencidos_hoje'], 'título vencido', 'títulos vencidos')}, "
        f"em média {media_txt} em atraso. Não se explica pelo padrão dele nem por mudança recente de "
        f"comportamento — cobrança direta, sem rodeio e sem ameaça."
    )


def montar_contexto_redacao(ref, p):
    """Traduz os numeros do perfil para o dicionario que redacao.py espera,
    ja formatado para leitura humana."""
    titulos = sorted(p["titulos_abertos"], key=lambda t: t["dt_vencimento"])
    vencidos = [t for t in titulos if t["dt_vencimento"] < ref]
    lista = vencidos if vencidos else titulos

    itens = [{
        "numero": t["numero"],
        "vencimento": formatar_data_curta(t["dt_vencimento"]),
        "valor": formatar_reais_de_centavos(t["valor_centavos"]),
        "dias": (ref - t["dt_vencimento"]).days,
    } for t in lista]

    total_centavos = sum(t["valor_centavos"] for t in lista)
    mais_antigo = lista[0] if lista else None

    return {
        "contato": p["cliente"]["contato"] or p["cliente"]["nome"],
        "empresa": p["cliente"]["nome"],
        "titulos": itens,
        "qtd_titulos": len(itens),
        "total": formatar_reais_de_centavos(total_centavos),
        "total_centavos": total_centavos,
        "dias_mais_antigo": (ref - mais_antigo["dt_vencimento"]).days if mais_antigo else 0,
        "data_mais_antiga": formatar_data_curta(mais_antigo["dt_vencimento"]) if mais_antigo else "—",
        # no meio da frase, gente escreve "22/06", nao "22/06/2026"
        "data_prosa": f"{mais_antigo['dt_vencimento']:%d/%m}" if mais_antigo else "—",
        "atraso_medio": f"{p['nivel']:.0f}" if p["nivel"] is not None else "—",
        "atraso_recente": f"{p['atraso_recente']:.0f}" if p["atraso_recente"] is not None else "—",
        "prazo_contratado": p["cliente"]["prazo_contratado"],
        "prazo_real": f"{p['prazo_efetivo']:.0f}" if p.get("prazo_efetivo") is not None else "—",
        "dias_sem_comprar": p["dias_sem_comprar"] if p["dias_sem_comprar"] is not None else "—",
        "intervalo_medio": f"{p['intervalo_medio']:.0f}" if p["intervalo_medio"] else "—",
        "participacao": f"{p['participacao_faturamento'] * 100:.1f}".replace(".", ",") + "%",
        "sugerir_prazo": bool(p["selo_cronico_estavel"]),
        "tem_historico": p["nivel"] is not None,
    }


def calcular_semana(ref, titulos):
    """O que a semana que acabou de passar produziu. Numero novo desta rodada:
    mora aqui junto dos outros, e tanto o relatorio quanto a tela de previsao leem daqui."""
    inicio = ref - dt.timedelta(days=6)

    recebidos = [
        t for t in titulos
        if t["dt_pagamento"] is not None and inicio <= t["dt_pagamento"] <= ref
    ]
    venceram = [t for t in titulos if inicio <= t["dt_vencimento"] <= ref]
    viraram_vencidos = [t for t in venceram if t["dt_pagamento"] is None]
    emitidos = [t for t in titulos if inicio <= t["dt_emissao"] <= ref]

    atrasos = [(t["dt_pagamento"] - t["dt_vencimento"]).days for t in recebidos]

    return {
        "inicio": inicio,
        "fim": ref,
        "recebido_centavos": sum(t["valor_centavos"] for t in recebidos),
        "recebido_qtd": len(recebidos),
        "atraso_medio_recebido": media(atrasos),
        "venceu_centavos": sum(t["valor_centavos"] for t in venceram),
        "venceu_qtd": len(venceram),
        "nao_entrou_centavos": sum(t["valor_centavos"] for t in viraram_vencidos),
        "nao_entrou_qtd": len(viraram_vencidos),
        "emitido_centavos": sum(t["valor_centavos"] for t in emitidos),
        "emitido_qtd": len(emitidos),
    }


def gerar_alerta_ficha(p):
    return " ".join(partes_do_alerta(p))


def partes_do_alerta(p):
    """O alerta da ficha frase a frase. A chamada da lista de Risco usa a primeira.
    Antes a chamada cortava o texto no primeiro ponto, que as vezes era o separador
    de milhar: 'saldo em aberto de R$ 130.' no lugar de R$ 130.360,00."""
    cliente = p["cliente"]
    nome = cliente["nome"]
    partes = []

    if p["selo_cronico_estavel"]:
        partes.append(
            f"{nome} aparece como atrasado quase todo mês, mas isso não é risco: há {p['qtd_pagos_ns']} títulos pagos, "
            f"ele atrasa de forma muito consistente (desvio-padrão de {texto_dias(p['variabilidade'])} em torno de "
            f"{p['nivel']:.0f} dias) e o atraso atual não fugiu desse padrão."
        )
        prazo_txt = (
            f"na prática o dinheiro entra em torno de {p['prazo_efetivo']:.0f} dias corridos depois da emissão"
            if p["prazo_efetivo"] is not None else
            f"na prática ele paga cerca de {p['nivel']:.0f} dias depois do vencimento"
        )
        partes.append(
            f"O prazo contratado com ele é \"{cliente['prazo_contratado']}\", mas {prazo_txt}. "
            f"Ação: parar de tratar isso como cobrança urgente — se fizer sentido, renegocie o prazo contratado para "
            f"perto do que ele já pratica, e pare de gastar tempo da equipe ligando para quem sempre paga."
        )
    elif p["selo_invisivel"]:
        partes.append(
            f"{nome} nunca apareceu em nenhuma lista de cobrança — não tem título vencido hoje e nunca foi contatado — "
            f"mas o histórico mostra o atraso médio subindo (de {p['atraso_antigo']:.0f} para {p['atraso_recente']:.0f} "
            f"dias nos últimos {JANELA_TENDENCIA_MESES} meses) e o saldo em aberto dele é de "
            f"{formatar_reais_de_centavos(p['saldo_aberto_centavos'])} — {texto_pct(p['participacao_faturamento'])} do "
            f"faturamento dos últimos {JANELA_FATURAMENTO_MESES} meses."
        )
        partes.append(
            "Ação: ligar preventivamente antes que isso vire um problema visível — nenhum sistema de cobrança vai "
            "apontar esse cliente sozinho, porque nada nele está formalmente vencido."
        )
    elif p["indice"] < 15:
        partes.append(f"{nome} tem risco baixo hoje: nenhum sinal do histórico dele preocupa. Ação: nenhuma.")
    else:
        principais = sorted(p["componentes"], key=lambda c: -c["pontos"])[:2]
        principais = [c for c in principais if c["pontos"] > 0]
        if principais:
            nomes_fatores = " e ".join(c["nome"].lower() for c in principais)
            partes.append(f"{nome} pesa no índice principalmente por causa de: {nomes_fatores}.")
        if p["piora_confirmada"]:
            partes.append(
                "Ação: ligar e perguntar diretamente se há um problema pontual — a mudança de comportamento é "
                "estatisticamente real, não é ruído do dia a dia."
            )
        elif p["ocupacao_limite"] is not None and p["ocupacao_limite"] >= 0.8:
            partes.append(
                f"Ação: reavaliar o limite de crédito antes de liberar novos pedidos — já está em "
                f"{texto_pct(p['ocupacao_limite'], 0)} do limite."
            )
        elif p["tem_vencido_hoje"]:
            partes.append("Ação: segue no fluxo normal de cobrança da tela Hoje; nada aqui pede tratamento especial além disso.")
        else:
            partes.append("Ação: nenhuma cobrança necessária agora; vale só acompanhar o saldo em aberto.")

    return partes


# Quantas descobertas vao para o alto da aba Hoje. O pedido foi "as duas ou tres
# coisas mais importantes que a Central descobriu".
MAX_DESCOBERTAS = 3


def calcular_descobertas(perfis):
    """O que a Central achou olhando o comportamento, e que uma lista de vencidos nao
    mostra. Um achado por tipo (o do cliente com mais dinheiro em risco, ou o grupo
    inteiro quando o tipo junta varios), ordenados pelo valor em risco envolvido.
    Nao calcula nada novo: le os mesmos sinais que a aba Risco mostra."""
    def risco(cods):
        return sum(perfis[c]["valor_em_risco_centavos"] for c in cods)

    def nomes(cods):
        n = [perfis[c]["cliente"]["nome"] for c in cods]
        return n[0] if len(n) == 1 else ", ".join(n[:-1]) + " e " + n[-1]

    achados = []

    def maior(filtro):
        cods = [c for c, p in perfis.items() if filtro(p)]
        return max(cods, key=lambda c: perfis[c]["valor_em_risco_centavos"]) if cods else None

    cod = maior(lambda p: p["piora_confirmada"])
    if cod is not None:
        p = perfis[cod]
        achados.append(dict(
            tipo="piorou", rotulo="Piorou", cor="atencao", cods=[cod],
            frase=(f"{p['cliente']['nome']} passou a pagar mais tarde: de {p['atraso_antigo']:.0f} para "
                   f"{texto_dias(p['atraso_recente'])} de atraso nos últimos {JANELA_TENDENCIA_MESES} meses. "
                   f"Mudança confirmada, que pede conversa, não cobrança dura."),
        ))

    cod = maior(lambda p: p["melhora_confirmada"])
    if cod is not None:
        p = perfis[cod]
        achados.append(dict(
            tipo="melhorou", rotulo="Melhorou", cor="ok", cods=[cod],
            frase=(f"{p['cliente']['nome']} passou a pagar mais cedo: de {p['atraso_antigo']:.0f} para "
                   f"{texto_dias(p['atraso_recente'])} de atraso nos últimos {JANELA_TENDENCIA_MESES} meses. "
                   f"A previsão de caixa já conta com o ritmo novo."),
        ))

    cod = maior(lambda p: p["selo_invisivel"])
    if cod is not None:
        p = perfis[cod]
        achados.append(dict(
            tipo="invisivel", rotulo="Fora de qualquer lista", cor="atencao", cods=[cod],
            frase=(f"{p['cliente']['nome']} nunca foi cobrado e não tem nada vencido, mas o atraso subiu de "
                   f"{p['atraso_antigo']:.0f} para {texto_dias(p['atraso_recente'])}. São "
                   f"{formatar_reais_de_centavos(p['saldo_aberto_centavos'])} em aberto que nenhuma lista de "
                   f"vencidos mostra."),
        ))

    cods = sorted((c for c, p in perfis.items() if p["selo_cronico_estavel"]),
                  key=lambda c: -perfis[c]["valor_em_risco_centavos"])
    if cods:
        varios = len(cods) > 1
        achados.append(dict(
            tipo="estavel", rotulo="Padrão, não é risco", cor="ok", cods=cods,
            frase=(f"{nomes(cods)} {'atrasam' if varios else 'atrasa'} sempre igual e "
                   f"{'nunca falharam' if varios else 'nunca falhou'}. Acertar o prazo no contrato economiza "
                   f"cobrança e tempo da equipe."),
        ))

    cod = maior(lambda p: p["razao_sem_comprar"] is not None and p["razao_sem_comprar"] >= RAZAO_PAROU_DE_COMPRAR
                and p["saldo_aberto_centavos"] > 0)
    if cod is not None:
        p = perfis[cod]
        achados.append(dict(
            tipo="parou", rotulo="Parou de comprar", cor="atencao", cods=[cod],
            frase=(f"{p['cliente']['nome']} está há {texto_dias(p['dias_sem_comprar'])} sem comprar, "
                   f"{texto_decimal(p['razao_sem_comprar'])} vezes o intervalo normal dele. A conversa comercial "
                   f"vem antes da cobrança."),
        ))

    cod = maior(lambda p: bool(p["sazonais"]))
    if cod is not None:
        p = perfis[cod]
        meses = [MESES_PT_INV[m] for m in sorted(p["sazonais"])]
        meses_txt = meses[0] if len(meses) == 1 else ", ".join(meses[:-1]) + " e " + meses[-1]
        achados.append(dict(
            tipo="sazonal", rotulo="Padrão de calendário", cor="neutro", cods=[cod],
            frase=(f"{p['cliente']['nome']} atrasa mais em {meses_txt}, todo ano. A Central desconta esses "
                   f"meses para não confundir com piora."),
        ))

    for a in achados:
        a["valor_em_risco_centavos"] = risco(a["cods"])
    achados.sort(key=lambda a: -a["valor_em_risco_centavos"])
    return achados[:MAX_DESCOBERTAS]


def historico_para_desenhar(p):
    """O atraso mes a mes do cartao da fila precisa de pelo menos dois meses com titulo pago e
    de uma media para a faixa normal. Cliente sem historico fica so com a regua. (Na V9 o dono
    pediu o grafico em todos os cartoes, tambem quando a linha e quase reta.)"""
    return p["nivel"] is not None and len(p["serie_mensal"]) >= 2


def gerar_relatorio_texto(ref, clientes, titulos, fila, perfis, prev, semana, registros):
    """Texto corrido para a diretoria. NAO calcula nada: le os mesmos numeros
    que as telas mostram."""
    r = formatar_reais_de_centavos
    e = formatar_reais_estimado

    abertos = [t for t in titulos if t["dt_pagamento"] is None]
    vencidos = [t for t in abertos if t["dt_vencimento"] < ref]
    total_aberto = sum(t["valor_centavos"] for t in abertos)
    total_vencido = sum(t["valor_centavos"] for t in vencidos)

    por_risco = sorted(perfis.values(), key=lambda p: -p["valor_em_risco_centavos"])
    total_em_risco = sum(p["valor_em_risco_centavos"] for p in perfis.values())
    top3 = por_risco[:3]

    cronicos = [p for p in perfis.values() if p["selo_cronico_estavel"]]
    invisiveis = [p for p in perfis.values() if p["selo_invisivel"]]

    linhas = []
    linhas.append("RELATÓRIO DA SEMANA — CRÉDITO E COBRANÇA")
    linhas.append("Distribuidora Aurora")
    linhas.append(f"Semana de {formatar_data_curta(semana['inicio'])} a {formatar_data_curta(semana['fim'])}")
    linhas.append("")

    # 1. o que entrou
    atraso_txt = (
        f", com atraso médio de {semana['atraso_medio_recebido']:.0f} dias"
        if semana["atraso_medio_recebido"] is not None else ""
    )
    linhas.append("O QUE ENTROU")
    linhas.append(
        f"Na semana entraram {r(semana['recebido_centavos'])} em "
        f"{semana['recebido_qtd']} {plural(semana['recebido_qtd'], 'título', 'títulos')}{atraso_txt}. "
        f"No mesmo período foram emitidos {r(semana['emitido_centavos'])} em "
        f"{semana['emitido_qtd']} {plural(semana['emitido_qtd'], 'título', 'títulos')} novos."
    )
    linhas.append("")

    # 2. o que nao entrou
    linhas.append("O QUE NÃO ENTROU")
    linhas.append(
        f"Venceram {r(semana['venceu_centavos'])} na semana e "
        f"{r(semana['nao_entrou_centavos'])} não foram pagos, em "
        f"{semana['nao_entrou_qtd']} {plural(semana['nao_entrou_qtd'], 'título', 'títulos')}. "
        f"A carteira vencida acumulada é de {r(total_vencido)} em "
        f"{len(vencidos)} títulos de {len(fila)} clientes, dentro de um saldo total em aberto de {r(total_aberto)}."
    )
    linhas.append("")

    # 3. onde esta o risco
    linhas.append("ONDE ESTÁ O RISCO")
    pedacos = []
    for p in top3:
        pedacos.append(
            f"{p['cliente']['nome']} ({e(p['valor_em_risco_centavos'])} em risco, "
            f"índice {p['indice']:.0f} de 100, sobre {r(p['saldo_aberto_centavos'])} em aberto)"
        )
    linhas.append(
        f"Ponderando o saldo de cada cliente pelo risco que o histórico dele mostra, "
        f"{e(total_em_risco)} dos {r(total_aberto)} em aberto estão em risco, o equivalente a "
        f"{texto_pct(total_em_risco / total_aberto, 0)} da carteira. "
        f"A concentração está em três nomes: {'; '.join(pedacos)}."
    )
    if invisiveis:
        inv = invisiveis[0]
        linhas.append(
            f"Vale destacar {inv['cliente']['nome']}, que não tem um único título vencido e nunca foi cobrado, "
            f"mas vem pagando cada vez mais tarde — a média subiu de {inv['atraso_antigo']:.0f} para "
            f"{inv['atraso_recente']:.0f} dias — carregando {r(inv['saldo_aberto_centavos'])} em aberto. "
            f"É o tipo de conta que nenhuma lista de vencidos mostra."
        )
    if cronicos:
        nomes = " e ".join(p["cliente"]["nome"] for p in cronicos)
        soma = sum(p["valor_em_risco_centavos"] for p in cronicos)
        linhas.append(
            f"No sentido oposto, {nomes} aparecem como atrasados todo mês, mas pagam com atraso constante há anos "
            f"e nunca falharam. Somam {e(soma)} em risco e consomem tempo da equipe sem necessidade: "
            f"o caminho aqui é acertar o prazo no contrato, não cobrar."
        )
    linhas.append("")

    # 4. quanto entra
    linhas.append("QUANTO ENTRA NAS PRÓXIMAS SEMANAS")
    linhas.append(
        f"Somando os vencimentos, uma planilha apontaria {e(prev['total_planilha'])} para as próximas quatro "
        f"semanas. Lendo o comportamento real de pagamento de cada cliente, a expectativa é de "
        f"{e(prev['total_esperado'])} — uma diferença de {e(abs(prev['diferenca']))}, ou "
        f"{texto_pct(abs(prev['diferenca']) / prev['total_planilha'], 0)} do que a planilha promete. "
        f"Desse total, {e(prev['fora_da_janela'])} só devem entrar depois das quatro semanas. "
        f"A projeção considera apenas a carteira existente e não inclui venda nova."
    )
    linhas.append("")

    # 5. o que fazer
    linhas.append("O QUE FAZER NESTA SEMANA")
    linhas.append(
        f"Em ordem de quanto vale, a fila de contato da semana é esta — "
        f"{len(fila)} {plural(len(fila), 'cliente', 'clientes')} com título vencido:"
    )
    linhas.append("")
    cod_maior = max(fila, key=lambda c: c["valor_vencido_centavos"])["cliente"]["cod"] if fila else None
    ordenada = sorted(fila, key=lambda c: -c["valor_vencido_centavos"])
    for i, c in enumerate(ordenada, start=1):
        p = perfis[c["cliente"]["cod"]]
        tom, _ = escolher_tom(p)
        linhas.append(
            f"{i}. {c['cliente']['nome']} — {r(c['valor_vencido_centavos'])} vencidos em "
            f"{c['qtd_vencidos']} {plural(c['qtd_vencidos'], 'título', 'títulos')}, o mais antigo há "
            f"{c['dias_atraso_max']} dias. Abordagem: {redacao.TONS[tom]['nome'].lower()}. "
            f"Falar com {p['cliente']['contato']}, {p['cliente']['telefone']}."
        )
    linhas.append("")

    # so o que foi registrado DENTRO da semana. O CSV acumula meses, e antes o
    # relatorio de setembro contava cobrancas de agosto como 'do periodo'.
    da_semana = [x for x in registros
                 if semana["inicio"] <= x["dt_hora_contato"].date() <= semana["fim"]]
    if da_semana:
        linhas.append(
            f"Foram registradas {len(da_semana)} {plural(len(da_semana), 'cobrança', 'cobranças')} na semana, "
            f"somando {r(sum(x['valor_cobrado_centavos'] for x in da_semana))}."
        )
        linhas.append("")

    linhas.append(
        f"Todos os números deste relatório vêm da Base_bruta.xlsx, com data de referência congelada em "
        f"{formatar_data_pt(ref)}, e são os mesmos que a Central mostra nas telas. Nada aqui foi recalculado "
        f"para este documento."
    )
    return "\n".join(linhas)


# ---------------------------------------------------------------------------
# Geracao do HTML
# ---------------------------------------------------------------------------
# Daqui para baixo mora a APARENCIA inteira: graficos em SVG montado a mao, o CSS,
# uma funcao render_conteudo_* por tela e render_pagina, que junta tudo com o
# JavaScript. Nada aqui calcula numero de negocio: so le o que as funcoes de cima
# devolveram e decide como mostrar.

# As fontes vao embutidas no proprio HTML (base64), para a pagina abrir igual numa
# maquina sem internet. Sem a pasta fontes/, a Central usa a fonte do sistema.
PASTA_FONTES = os.path.join(PASTA, "fontes")
FONTES = (
    # familia, arquivo, faixa de peso da fonte variavel
    ("Unbounded", "Unbounded-latin.woff2", "300 900"),
    ("Inter", "Inter-latin.woff2", "300 800"),
)

# O arquivo para mandar para alguem da equipe: um arquivo so, que abre com dois
# cliques. Fica numa pasta propria para nao ser confundido com o central.html.
PASTA_ENVIO = os.path.join(PASTA, "para_enviar")
ARQUIVO_ENVIO = os.path.join(PASTA_ENVIO, "Central Aurora.html")

FAIXAS_RISCO = [
    # quartos da escala 0-100 do proprio indice; a legenda na tela diz isso.
    (75, "critico", "risco crítico"),
    (50, "alto", "risco alto"),
    (25, "moderado", "risco moderado"),
    (0, "baixo", "risco baixo"),
]

# Poucas cores, um trabalho so para cada uma, e sempre com a palavra junto:
#   grave   = fora do padrao / risco alto ou critico
#   atencao = piorou / risco moderado / fora de qualquer lista
#   ok      = e o padrao dele / risco baixo
#   neutro  = nada a julgar agora (aguardando retorno, sem historico)
# Fora delas so existem a cor de ACAO (botao, aba atual) e a de DINHEIRO no texto.
COR_DA_FAIXA = {"baixo": "ok", "moderado": "atencao", "alto": "grave", "critico": "grave"}

# redacao.TONS da uma cor para cada um dos seis tons. Na tela elas se reduzem as
# cores de situacao: leve = ok, conversa sobre mudanca = atencao, objetiva = grave,
# o resto neutro. Seis cores de tom seriam seis significados novos para decorar.
COR_DO_TOM = {"verde": "ok", "ambar": "atencao", "vermelho": "grave",
              "azul": "neutro", "roxo": "neutro", "cinza": "neutro"}

ICONE_DA_SITUACAO = {"aguardando": "relogio", "sem_historico": "duvida", "fora": "alerta",
                     "piorou": "subiu", "padrao": "certo"}
ICONE_DA_DESCOBERTA = {"piorou": "subiu", "melhorou": "desceu", "invisivel": "olho",
                       "estavel": "certo", "parou": "carrinho", "sazonal": "calendario"}

# Tracos de 24x24, desenhados uma vez e reusados com <use>. Sem xmlns de proposito:
# dentro do HTML o navegador nao precisa, e a conferencia reprova "http://" na pagina.
ICONES = {
    "hoje": '<circle cx="12" cy="12" r="4"/><path d="M12 2.5v2.2M12 19.3v2.2M2.5 12h2.2M19.3 12h2.2M5.3 5.3l1.5 1.5M17.2 17.2l1.5 1.5M5.3 18.7l1.5-1.5M17.2 6.8l1.5-1.5"/>',
    "risco": '<path d="M12 3 5 6v5.5c0 4.3 3 8 7 9.5 4-1.5 7-5.2 7-9.5V6l-7-3z"/><path d="M12 8.5v4M12 16h.01"/>',
    "previsao": '<path d="M3 17.5 8.5 12l4 3.2L21 6.5"/><path d="M15 6.5h6v6"/>',
    "regua": '<path d="M4 5.5h16v11H9l-5 4v-15z"/><path d="M8 9.5h8M8 12.8h5"/>',
    "relatorio": '<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5M9 12h7M9 16h7"/>',
    "busca": '<circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.2-4.2"/>',
    "certo": '<circle cx="12" cy="12" r="9"/><path d="m8 12.2 2.8 2.8L16.2 9"/>',
    "check": '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
    "alerta": '<path d="M12 3.8 2.8 19.5h18.4L12 3.8z"/><path d="M12 10v4.2M12 17h.01"/>',
    "subiu": '<path d="M6 18 18 6M9.5 6H18v8.5"/>',
    "desceu": '<path d="M6 6l12 12M18 9.5V18H9.5"/>',
    "relogio": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5.2l3.2 2"/>',
    "duvida": '<circle cx="12" cy="12" r="9"/><path d="M9.6 9.4a2.5 2.5 0 1 1 3.4 2.3c-.6.3-1 .8-1 1.5v.4M12 17h.01"/>',
    "olho": '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="2.8"/><path d="M4 4l16 16"/>',
    "carrinho": '<circle cx="9.5" cy="19.5" r="1.4"/><circle cx="17.5" cy="19.5" r="1.4"/><path d="M3 4h2.2l2.3 11h11L21 7.5H6.3"/>',
    "calendario": '<rect x="3.5" y="5" width="17" height="15.5" rx="2.5"/><path d="M3.5 10h17M8 3v4M16 3v4"/>',
    "telefone": '<path d="M5.2 3.5h3.4l1.7 4.4-2.3 1.4a11.5 11.5 0 0 0 6.7 6.7l1.4-2.3 4.4 1.7v3.4a2 2 0 0 1-2.1 2A16.5 16.5 0 0 1 3.2 5.6a2 2 0 0 1 2-2.1z"/>',
    "fechar": '<path d="M6 6l12 12M18 6 6 18"/>',
    "copiar": '<rect x="8.5" y="8.5" width="11.5" height="11.5" rx="2.2"/><path d="M15.5 8.5V5.8a1.8 1.8 0 0 0-1.8-1.8H5.8A1.8 1.8 0 0 0 4 5.8v7.9a1.8 1.8 0 0 0 1.8 1.8h2.7"/>',
    "imprimir": '<path d="M7 8.5V3.5h10v5M7 17H4.5v-7.5h15V17H17"/><rect x="7" y="14" width="10" height="6.5" rx="1"/>',
    "ficha": '<rect x="4.5" y="3" width="15" height="18" rx="2.2"/><path d="M8.5 8h7M8.5 12h7M8.5 16h4"/>',
    "seta": '<path d="M5 12h14M13 6l6 6-6 6"/>',
    "baixar": '<path d="M12 4v11M7.5 10.5 12 15l4.5-4.5M5 20h14"/>',
    "lixo": '<path d="M4.5 7h15M9.5 7V4.5h5V7M6.5 7l1 13h9l1-13"/>',
    "lista": '<path d="M8.5 6.5h11M8.5 12h11M8.5 17.5h11"/><circle cx="4.5" cy="6.5" r=".9"/><circle cx="4.5" cy="12" r=".9"/><circle cx="4.5" cy="17.5" r=".9"/>',
    "info": '<circle cx="12" cy="12" r="9"/><path d="M12 11v5.5M12 7.8h.01"/>',
}


def icone(nome, classe=""):
    return f'<svg class="ic{" " + classe if classe else ""}" aria-hidden="true"><use href="#i-{nome}"/></svg>'


def sprite_icones():
    simbolos = "".join(
        f'<symbol id="i-{nome}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" '
        f'stroke-linecap="round" stroke-linejoin="round">{corpo}</symbol>'
        for nome, corpo in ICONES.items()
    )
    return f'<svg width="0" height="0" style="position:absolute" aria-hidden="true">{simbolos}</svg>'


def faixa_de_risco(indice):
    for minimo, chave, rotulo in FAIXAS_RISCO:
        if indice >= minimo:
            return chave, rotulo
    return "baixo", "risco baixo"


# Todo valor em dinheiro no meio de uma frase ganha a cor de dinheiro, para o olho
# achar os numeros sem ler. "R$ 1.234,56", "R$ 1.234" (estimativa) e com sinal.
RE_DINHEIRO = re.compile(r"(?:[−-]\s?)?R\$\s?\d{1,3}(?:\.\d{3})*(?:,\d{2})?")


def prosa(texto):
    """Texto corrido para a tela: escapa o HTML e pinta cada valor em dinheiro."""
    return RE_DINHEIRO.sub(lambda m: f'<span class="dinheiro">{m.group(0)}</span>', html.escape(texto))


def sem_acento(texto):
    """Para a busca: 'Frescor Hortifrúti' acha com 'hortifruti'."""
    decomposto = unicodedata.normalize("NFD", texto or "")
    return "".join(ch for ch in decomposto if not unicodedata.combining(ch)).lower()


def formatar_eixo_reais(centavos):
    """So para rotulo de eixo, onde precisao atrapalha a leitura: 'R$ 250 mil'."""
    reais = centavos / 100
    if abs(reais) >= 1_000_000:
        return "R$ " + texto_decimal(reais / 1_000_000) + " mi"
    if abs(reais) >= 1000:
        return f"R$ {reais / 1000:.0f} mil"
    return f"R$ {reais:.0f}"


def css_fontes():
    blocos = []
    for familia, arquivo, pesos in FONTES:
        caminho = os.path.join(PASTA_FONTES, arquivo)
        if not os.path.exists(caminho):
            print(f"Aviso: nao achei fontes/{arquivo}. A pagina vai usar a fonte do sistema.")
            continue
        with open(caminho, "rb") as f:
            dados = base64.b64encode(f.read()).decode("ascii")
        blocos.append(
            f"@font-face{{font-family:'{familia}';font-style:normal;font-weight:{pesos};font-display:swap;"
            f"src:url(data:font/woff2;base64,{dados}) format('woff2');}}"
        )
    return "\n".join(blocos)


# ---------------------------------------------------------------------------
# Graficos (SVG a mao). A cor vem do CSS pelas classes, para o mesmo desenho
# servir na tela escura e na impressao em papel branco.
# ---------------------------------------------------------------------------

def gerar_regua_atraso_svg(c, situacao):
    """Na fila de Hoje: o titulo mais antigo contra a faixa onde o PROPRIO cliente
    costuma pagar (media +- 1 desvio-padrao, a mesma faixa do grafico da ficha).
    E o 'por que ele esta aqui' num olhar: ponto longe da faixa = fora do padrao."""
    H = c["atraso_medio_historico"]
    dp = c["desvio_atraso_historico"] or 0.0
    dias = sorted(t["dias_atraso_atual"] for t in c["titulos_vencidos"])
    mais_antigo = dias[-1]
    tem_faixa = H is not None
    lo, hi = (H - dp, H + dp) if tem_faixa else (None, None)

    vmin = min(0.0, lo) if tem_faixa else 0.0
    vmax = max([mais_antigo, 10.0] + ([hi] if tem_faixa else [])) * 1.1
    L, A = 340, 70
    esq, dir_ = 10, 10
    y = 32

    def x(v):
        return esq + (v - vmin) / (vmax - vmin) * (L - esq - dir_)

    cor = SITUACOES_FILA[situacao][1]
    partes = [f'<line class="rg-trilho" x1="{x(vmin):.1f}" y1="{y}" x2="{x(vmax):.1f}" y2="{y}"/>']
    if vmin < 0:
        partes.append(f'<line class="rg-zero" x1="{x(0):.1f}" y1="{y - 9}" x2="{x(0):.1f}" y2="{y + 9}"/>')

    if tem_faixa:
        x1, x2 = x(lo), x(hi)
        partes.append(
            f'<rect class="rg-faixa" x="{x1:.1f}" y="{y - 8}" width="{max(x2 - x1, 6):.1f}" height="16" rx="8"/>'
        )
        partes.append(f'<line class="rg-media" x1="{x(H):.1f}" y1="{y - 8}" x2="{x(H):.1f}" y2="{y + 8}"/>')
        normal = (f"normal dele: até {hi:.0f} dias" if lo < 0.5 else f"normal dele: {lo:.0f} a {hi:.0f} dias")
        ancora, xr = ("start", x1) if x1 < L * 0.55 else ("end", x2)
        partes.append(f'<text class="rg-rotulo" x="{xr:.1f}" y="{y + 30}" text-anchor="{ancora}">{normal}</text>')
    else:
        normal = "sem padrão para comparar"
        partes.append(f'<text class="rg-rotulo" x="{esq}" y="{y + 30}">{normal}</text>')

    for d in dias[:-1]:
        partes.append(f'<circle class="rg-titulo" cx="{x(d):.1f}" cy="{y}" r="3.2"/>')
    xh = x(mais_antigo)
    partes.append(f'<circle class="rg-hoje cor-{cor}" cx="{xh:.1f}" cy="{y}" r="6.5"/>')
    ancora = "end" if xh > L - 50 else ("start" if xh < 50 else "middle")
    xt = xh + 8 if ancora == "end" else (xh - 8 if ancora == "start" else xh)
    partes.append(
        f'<text class="rg-valor" x="{xt:.1f}" y="{y - 14}" text-anchor="{ancora}">{texto_dias(mais_antigo)}</text>'
    )

    rotulo = f"Título mais antigo vencido há {texto_dias(mais_antigo)}; {normal}."
    return (f'<svg class="regua-atraso" viewBox="0 0 {L} {A}" role="img" aria-label="{html.escape(rotulo)}">'
            f'{"".join(partes)}</svg>')


def gerar_linha_atraso_svg(p, cor):
    """Nos cartoes da fila, embaixo da regua: em quantos dias o cliente pagou, mes a mes, com
    a faixa normal dele por tras. So e chamado quando historico_para_desenhar(p). Mesma
    largura da regua (340), para as duas lerem na mesma escala de letra. O ultimo valor vem com
    o mes ("18 d · abr/26"): e o ultimo mes com titulo ja pago, nao o titulo aberto hoje -- esse
    esta na regua, logo acima."""
    serie = list(p["serie_mensal"].items())
    valores = [v for _, v in serie]
    lo, hi = p["nivel"] - (p["variabilidade"] or 0), p["nivel"] + (p["variabilidade"] or 0)
    vmin, vmax = min(valores + [lo, 0]), max(valores + [hi])
    # altura minima de 20 dias: sem isso, quem paga sempre entre 1 e 5 dias ganha uma linha
    # "nervosa", porque o grafico estica a variacao pequena ate ocupar a altura toda
    vmax = max(vmax, vmin + 20)
    folga = max(1.0, (vmax - vmin) * 0.1)
    vmin, vmax = vmin - folga, vmax + folga
    L, A = 340, 66
    esq, dir_, topo, baixo = 10, 104, 22, 4

    def x(i):
        return esq + i / (len(serie) - 1) * (L - esq - dir_)

    def y(v):
        return topo + (1 - (v - vmin) / (vmax - vmin)) * (A - topo - baixo)

    pts = " ".join(f"{x(i):.1f},{y(v):.1f}" for i, v in enumerate(valores))
    (ano_ini, mes_ini), _ = serie[0]
    (ano_fim, mes_fim), _ = serie[-1]
    quando = f"{MESES_PT_INV[mes_fim][:3]}/{ano_fim % 100:02d}"
    xf, yf = x(len(valores) - 1), y(valores[-1])
    dica = (f"Em quantos dias pagou, mês a mês, de {mes_ini:02d}/{ano_ini} a {mes_fim:02d}/{ano_fim} (só títulos já "
            f"pagos): de {min(valores):.0f} a {texto_dias(max(valores))}. A faixa clara é o normal dele "
            f"({max(lo, 0):.0f} a {hi:.0f} dias).")
    return (
        f'<svg class="linha-atraso cor-{cor}" viewBox="0 0 {L} {A}" role="img" aria-label="{html.escape(dica)}">'
        f'<title>{html.escape(dica)}</title>'
        f'<text class="la-rotulo" x="{esq}" y="12">atraso mês a mês</text>'
        f'<rect class="la-faixa" x="{esq}" y="{y(hi):.1f}" width="{L - esq - dir_}" height="{max(y(lo) - y(hi), 2):.1f}" rx="3"/>'
        f'<polyline class="la-linha" points="{pts}"/>'
        f'<circle class="la-fim" cx="{xf:.1f}" cy="{yf:.1f}" r="4"/>'
        f'<text class="la-valor" x="{xf + 8:.1f}" y="{yf + 4.5:.1f}">{valores[-1]:.0f} d · {quando}</text>'
        f'</svg>'
    )


def gerar_dispersao_risco_svg(perfis):
    """Aba Risco: quem deve muito (para a direita) e quem esta piorando (para cima), de uma
    vez so. Eixo de cima: atraso medio dos ultimos 6 meses menos o de antes, a mesma conta
    da 'mudanca recente de comportamento'. Em ambar, quem piorou de forma confirmada.
    Ocupa a largura toda da tela, por isso a area de desenho e larga e baixa."""
    pontos = [p for p in perfis.values() if p["diff_mudanca"] is not None]
    if not pontos:
        return "<p class='sem-grafico'>Nenhum cliente com histórico suficiente para comparar.</p>"
    L, A = 1300, 320
    esq, dir_, topo, baixo = 64, 24, 30, 50
    largura_letra = 8.1  # largura media de uma letra do rotulo (14,5px), para nao encavalar nomes
    xmax = max(p["saldo_aberto_centavos"] for p in pontos) * 1.08 or 1
    difs = [p["diff_mudanca"] for p in pontos]
    ymin, ymax = min(difs + [-5.0]) - 1.5, max(difs + [10.0]) + 1.5

    def x(v):
        return esq + v / xmax * (L - esq - dir_)

    def y(v):
        return topo + (1 - (v - ymin) / (ymax - ymin)) * (A - topo - baixo)

    partes = [f'<rect class="dp-quadrante" x="{esq}" y="{topo}" width="{L - esq - dir_}" height="{y(0) - topo:.1f}"/>']
    for v in passos_do_eixo(0, xmax, 6):
        if v > xmax:
            continue
        partes.append(f'<line class="g-grade" x1="{x(v):.1f}" y1="{topo}" x2="{x(v):.1f}" y2="{A - baixo}"/>')
        partes.append(f'<text class="g-eixo" x="{x(v):.1f}" y="{A - baixo + 20}" text-anchor="middle">{formatar_eixo_reais(v)}</text>')
    for v in passos_do_eixo(ymin, ymax, 5):
        if not (ymin <= v <= ymax):
            continue
        partes.append(f'<line class="{"g-zero" if v == 0 else "g-grade"}" x1="{esq}" y1="{y(v):.1f}" x2="{L - dir_}" y2="{y(v):.1f}"/>')
        rot = "0" if v == 0 else f"{v:+.0f} d"
        partes.append(f'<text class="g-eixo" x="{esq - 10}" y="{y(v) + 4:.1f}" text-anchor="end">{rot}</text>')
    partes.append(f'<text class="g-eixo forte" x="{esq + 8}" y="{topo + 17}">↑ pagando mais tarde que antes</text>')
    partes.append(f'<text class="g-eixo forte" x="{esq + 8}" y="{A - baixo - 9}">↓ pagando mais cedo</text>')
    partes.append(f'<text class="g-eixo forte" x="{L - dir_}" y="{A - 6}" text-anchor="end">saldo em aberto →</text>')

    # quem ganha nome escrito: quem piorou, quem esta fora de qualquer lista e os tres
    # maiores saldos. O resto se le passando o mouse.
    maiores = sorted(pontos, key=lambda p: -p["saldo_aberto_centavos"])[:3]
    rotular = {p["cliente"]["cod"] for p in pontos if p["piora_confirmada"] or p["selo_invisivel"]}
    rotular |= {p["cliente"]["cod"] for p in maiores}
    ocupados = [(x(p["saldo_aberto_centavos"]) - 8, y(p["diff_mudanca"]) - 8,
                 x(p["saldo_aberto_centavos"]) + 8, y(p["diff_mudanca"]) + 8) for p in pontos]

    def livre(caixa):
        x1, y1, x2, y2 = caixa
        if x1 < esq or x2 > L - dir_ + 4 or y1 < topo - 4 or y2 > A - baixo + 2:
            return False
        return all(x2 < a or x1 > c or y2 < b or y1 > d for a, b, c, d in ocupados)

    # quem piorou e quem esta fora de qualquer lista escolhem lugar primeiro: sao o motivo do desenho
    essenciais = {p["cliente"]["cod"] for p in pontos if p["piora_confirmada"] or p["selo_invisivel"]}
    for p in sorted(pontos, key=lambda p: (p["cliente"]["cod"] not in rotular, p["cliente"]["cod"] not in essenciais,
                                           -p["saldo_aberto_centavos"])):
        cx, cy = x(p["saldo_aberto_centavos"]), y(p["diff_mudanca"])
        piorou = p["piora_confirmada"]
        dica = (f"{p['cliente']['nome']}: {formatar_reais_de_centavos(p['saldo_aberto_centavos'])} em aberto; "
                f"atraso {texto_dias(p['diff_mudanca'], 1, sinal=True)} nos últimos {JANELA_TENDENCIA_MESES} meses"
                f"{' — piora confirmada' if piorou else ''}")
        partes.append(f'<circle class="dp-ponto{" piorou" if piorou else ""}" cx="{cx:.1f}" cy="{cy:.1f}" r="7">'
                      f'<title>{html.escape(dica)}</title></circle>')
        if p["cliente"]["cod"] not in rotular:
            continue
        nome = p["cliente"]["nome"] + (" · piorou" if piorou else "")
        w = len(nome) * largura_letra
        for ax, ay, anc, caixa in (
            (cx + 12, cy + 5, "start", (cx + 11, cy - 8, cx + 13 + w, cy + 7)),
            (cx - 12, cy + 5, "end", (cx - 13 - w, cy - 8, cx - 11, cy + 7)),
            (cx, cy - 14, "middle", (cx - w / 2, cy - 27, cx + w / 2, cy - 11)),
            (cx, cy + 24, "middle", (cx - w / 2, cy + 11, cx + w / 2, cy + 27)),
            (cx - 9, cy - 14, "end", (cx - 10 - w, cy - 27, cx - 8, cy - 11)),
            (cx - 9, cy + 24, "end", (cx - 10 - w, cy + 11, cx - 8, cy + 27)),
            (cx + 9, cy - 14, "start", (cx + 8, cy - 27, cx + 10 + w, cy - 11)),
            (cx + 9, cy + 24, "start", (cx + 8, cy + 11, cx + 10 + w, cy + 27)),
        ):
            if livre(caixa):
                partes.append(f'<text class="dp-rotulo{" piorou" if piorou else ""}" x="{ax:.1f}" y="{ay:.1f}" '
                              f'text-anchor="{anc}">{html.escape(nome)}</text>')
                ocupados.append(caixa)
                break

    nota = ""
    if len(pontos) < len(perfis):
        falta = len(perfis) - len(pontos)
        nota = (f'<p class="sub" style="margin:6px 0 0">{falta} {plural(falta, "cliente ficou", "clientes ficaram")} '
                f'fora do desenho: sem histórico suficiente para comparar antes e depois.</p>')
    return (f'<div class="rolagem-grafico"><svg class="dispersao-risco" viewBox="0 0 {L} {A}" role="img" '
            f'aria-label="Saldo em aberto contra a mudança no atraso de cada cliente">{"".join(partes)}</svg></div>{nota}')


def gerar_grafico_svg(p):
    """Atraso medio mensal, com a faixa do padrao historico e os pontos fora do
    padrao (ou sazonais) marcados. O fim da serie e a faixa vem rotulados no proprio
    desenho, para a conclusao estar no grafico e nao so na legenda."""
    serie = list(p["serie_mensal"].items())
    if len(serie) < 2:
        return "<p class='sem-grafico'>Histórico mensal insuficiente para desenhar o gráfico.</p>"

    largura, altura = 720, 250
    pad_esq, pad_dir, pad_topo, pad_baixo = 44, 70, 20, 30
    plot_w = largura - pad_esq - pad_dir
    plot_h = altura - pad_topo - pad_baixo

    tem_faixa = p["nivel"] is not None and p["variabilidade"] is not None
    valores = [v for _, v in serie]
    candidatos_y = list(valores)
    if tem_faixa:
        candidatos_y += [p["nivel"] - p["variabilidade"], p["nivel"] + p["variabilidade"]]
    y_min, y_max = min(candidatos_y), max(candidatos_y)
    folga = max(2.0, (y_max - y_min) * 0.12)
    y_min, y_max = y_min - folga, y_max + folga
    if y_max == y_min:
        y_max += 1

    def esc_x(i):
        return pad_esq + (i / (len(serie) - 1)) * plot_w

    def esc_y(v):
        return pad_topo + (1 - (v - y_min) / (y_max - y_min)) * plot_h

    partes = []
    for valor_grade in passos_do_eixo(y_min, y_max):
        if not (y_min <= valor_grade <= y_max):
            continue  # a grade redonda pode cair fora da area e ir parar em cima da legenda
        yg = esc_y(valor_grade)
        partes.append(f'<line class="g-grade" x1="{pad_esq}" y1="{yg:.1f}" x2="{largura - pad_dir}" y2="{yg:.1f}"/>')
        partes.append(f'<text class="g-eixo" x="{pad_esq - 8}" y="{yg + 3.5:.1f}" text-anchor="end">{valor_grade:.0f}d</text>')

    if tem_faixa:
        lo, hi = p["nivel"] - p["variabilidade"], p["nivel"] + p["variabilidade"]
        yt, yb = esc_y(hi), esc_y(lo)
        partes.append(
            f'<rect class="g-faixa" x="{pad_esq}" y="{yt:.1f}" width="{plot_w}" height="{(yb - yt):.1f}" rx="4">'
            f'<title>Faixa do padrão histórico dele: {lo:.0f} a {hi:.0f} dias</title></rect>'
        )
        yn = esc_y(p["nivel"])
        partes.append(f'<line class="g-media" x1="{pad_esq}" y1="{yn:.1f}" x2="{largura - pad_dir}" y2="{yn:.1f}"/>')
        # o rotulo vai no comeco da faixa, por dentro: na ponta direita ele disputava
        # lugar com o valor do ultimo mes e com o "vencimento"
        partes.append(
            f'<text class="g-anot" x="{pad_esq + 8}" y="{yt - 6:.1f}">padrão dele: {lo:.0f} a {hi:.0f} dias</text>'
        )

    if y_min < 0 < y_max:
        yz = esc_y(0)
        partes.append(f'<line class="g-zero" x1="{pad_esq}" y1="{yz:.1f}" x2="{largura - pad_dir}" y2="{yz:.1f}"/>')
        partes.append(f'<text class="g-eixo" x="{largura - pad_dir + 6}" y="{yz + 3.5:.1f}">vencimento</text>')

    pontos_xy = [(esc_x(i), esc_y(v)) for i, (_, v) in enumerate(serie)]
    partes.append(
        '<polyline class="g-linha" points="' + " ".join(f"{x:.1f},{y:.1f}" for x, y in pontos_xy) + '"/>'
    )

    corte = p["corte_tendencia"]
    sazonais = p["sazonais"]
    for i, ((ano, mes), valor) in enumerate(serie):
        x, y = pontos_xy[i]
        fora_da_faixa = tem_faixa and (
            valor > p["nivel"] + p["variabilidade"] or valor < p["nivel"] - p["variabilidade"]
        )
        eh_recente = dt.date(ano, mes, 1) >= dt.date(corte.year, corte.month, 1)
        if mes in sazonais:
            classe, raio, nota = "g-ponto sazonal", 5, " · mês sazonal, descontado da comparação"
        elif fora_da_faixa and eh_recente and p["piora_confirmada"]:
            classe, raio, nota = "g-ponto fora", 5, " · fora do padrão dele"
        else:
            classe, raio, nota = "g-ponto", 3.2, ""
        dica = f"{MESES_PT_INV[mes]}/{ano}: pagou {texto_dias(valor)} após o vencimento{nota}"
        partes.append(
            f'<circle class="{classe}" cx="{x:.1f}" cy="{y:.1f}" r="{raio}"><title>{html.escape(dica)}</title></circle>'
        )
        if i == 0 or i == len(serie) - 1 or i % 3 == 0:
            partes.append(
                f'<text class="g-eixo" x="{x:.1f}" y="{altura - 9}" text-anchor="middle">{mes:02d}/{str(ano)[2:]}</text>'
            )

    # o ultimo mes rotulado na ponta: e o numero que a pessoa procura
    (ano_f, mes_f), v_f = serie[-1]
    xf, yf = pontos_xy[-1]
    partes.append(
        f'<text class="g-fim" x="{xf + 9:.1f}" y="{yf - 8:.1f}">{texto_dias(v_f)}</text>'
    )

    return (
        f'<svg class="grafico-ficha" viewBox="0 0 {largura} {altura}" width="100%" '
        f'role="img" aria-label="Atraso médio de pagamento mês a mês">{"".join(partes)}</svg>'
    )


def gerar_grafico_previsao_svg(prev):
    """Acumulado das 4 semanas. A linha da planilha corre por cima para mostrar o
    tamanho da promessa; a faixa entre otimista e conservador so fecha aqui. A
    diferenca vem escrita na ponta, entre as duas linhas."""
    n = len(prev["semanas"])
    largura, altura = 760, 330
    pad_esq, pad_dir, pad_topo, pad_baixo = 86, 168, 22, 46
    plot_w = largura - pad_esq - pad_dir
    plot_h = altura - pad_topo - pad_baixo

    topo = (max(prev["acum_planilha"] + prev["acum_otimista"]) or 1) * 1.08

    def x(i):
        return pad_esq + (i / n) * plot_w

    def y(v):
        return pad_topo + (1 - v / topo) * plot_h

    partes = []
    for valor in passos_do_eixo(0, topo):
        yy = y(valor)
        partes.append(f'<line class="g-grade" x1="{pad_esq}" y1="{yy:.1f}" x2="{largura - pad_dir}" y2="{yy:.1f}"/>')
        partes.append(
            f'<text class="g-eixo" x="{pad_esq - 10}" y="{yy + 3.5:.1f}" text-anchor="end">{formatar_eixo_reais(valor)}</text>'
        )

    cima = [(x(0), y(0))] + [(x(i + 1), y(v)) for i, v in enumerate(prev["acum_otimista"])]
    baixo = [(x(i + 1), y(v)) for i, v in enumerate(prev["acum_conservador"])] + [(x(0), y(0))]
    poly = " ".join(f"{px:.1f},{py:.1f}" for px, py in cima + list(reversed(baixo)))
    partes.append(
        f'<polygon class="pv-faixa" points="{poly}"><title>Faixa entre o cenário otimista e o conservador</title></polygon>'
    )

    def linha(serie, classe, rotulo):
        pts = [(x(0), y(0))] + [(x(i + 1), y(v)) for i, v in enumerate(serie)]
        d = " ".join(f"{px:.1f},{py:.1f}" for px, py in pts)
        partes.append(f'<polyline class="{classe}" points="{d}"/>')
        for i, v in enumerate(serie):
            partes.append(
                f'<circle class="{classe}-ponto" cx="{x(i + 1):.1f}" cy="{y(v):.1f}" r="4.2">'
                f"<title>{html.escape(rotulo)} até a semana {i + 1}: {formatar_reais_estimado(v)}</title></circle>"
            )

    linha(prev["acum_planilha"], "pv-planilha", "Planilha")
    linha(prev["acum_esperado"], "pv-esperado", "Esperado")

    # rotulo direto na ponta de cada linha, e a diferenca escrita entre elas
    xf = x(n)
    yp, ye = y(prev["acum_planilha"][-1]), y(prev["acum_esperado"][-1])
    if prev["diferenca"] and abs(ye - yp) > 24:
        partes.append(f'<line class="pv-vao" x1="{xf:.1f}" y1="{min(yp, ye) + 7:.1f}" x2="{xf:.1f}" y2="{max(yp, ye) - 7:.1f}"/>')
        sinal = "−" if prev["diferenca"] > 0 else "+"
        partes.append(f'<text class="pv-vao-rotulo" x="{xf - 12:.1f}" y="{(yp + ye) / 2 + 5:.1f}" text-anchor="end">'
                      f'{sinal} {formatar_reais_estimado(abs(prev["diferenca"]))}</text>')
    if abs(yp - ye) < 34:
        ym = (yp + ye) / 2
        yp, ye = (ym - 17, ym + 17) if yp < ye else (ym + 17, ym - 17)
    partes.append(f'<text class="pv-rotulo" x="{xf + 12:.1f}" y="{yp - 2:.1f}">a planilha promete</text>')
    partes.append(f'<text class="pv-valor planilha" x="{xf + 12:.1f}" y="{yp + 14:.1f}">{formatar_reais_estimado(prev["total_planilha"])}</text>')
    partes.append(f'<text class="pv-rotulo" x="{xf + 12:.1f}" y="{ye - 2:.1f}">esperamos receber</text>')
    partes.append(f'<text class="pv-valor esperado" x="{xf + 12:.1f}" y="{ye + 14:.1f}">{formatar_reais_estimado(prev["total_esperado"])}</text>')

    for i, s in enumerate(prev["semanas"]):
        partes.append(f'<text class="g-eixo forte" x="{x(i + 1):.1f}" y="{altura - 24}" text-anchor="middle">sem {i + 1}</text>')
        partes.append(f'<text class="g-eixo" x="{x(i + 1):.1f}" y="{altura - 9}" text-anchor="middle">{s["fim"]:%d/%m}</text>')
    partes.append(f'<text class="g-eixo" x="{x(0):.1f}" y="{altura - 24}" text-anchor="middle">hoje</text>')

    return (
        f'<svg class="grafico-previsao" viewBox="0 0 {largura} {altura}" width="100%" role="img" '
        f'aria-label="Entrada de caixa acumulada nas próximas 4 semanas">{"".join(partes)}</svg>'
    )


# ---------------------------------------------------------------------------
# CSS. Direcao "Aurora": ceu noturno com faixas de luz no alto, que escurece ate
# o pe da pagina; vidro em tudo; brilho so no que e principal (numero do alto,
# botao principal, aba atual, primeiro da fila).
# ---------------------------------------------------------------------------

CSS = """
__FONTES__
:root {
  --fundo-topo: #0D1336;
  --fundo-meio: #080B1D;
  --fundo-pe: #03040A;
  --tinta: #EEF1FB;
  --tinta-2: #B6BDD6;
  --tinta-3: #8A93B0;
  --linha: rgba(255,255,255,.08);
  --linha-2: rgba(255,255,255,.15);
  --vidro: linear-gradient(180deg, rgba(255,255,255,.075) 0%, rgba(255,255,255,.03) 100%);
  --vidro-borda: rgba(255,255,255,.10);
  --vidro-luz: inset 0 1px 0 rgba(255,255,255,.07);
  --sombra: 0 22px 44px -26px rgba(0,0,0,.85);

  /* cor = um trabalho so, e sempre com a palavra junto */
  --grave: #FF8E84;   --grave-f: rgba(255,142,132,.14);
  --atencao: #F7C766; --atencao-f: rgba(247,199,102,.14);
  --ok: #6BE3AC;      --ok-f: rgba(107,227,172,.13);
  --neutro: #AEB6D0;  --neutro-f: rgba(174,182,208,.13);
  --dinheiro: #7CC6FF;
  --acao: #9585FF; --acao-2: #5C8DFF; --acao-texto: #9585FF;

  --display: 'Unbounded', 'Segoe UI', system-ui, sans-serif;
  --texto: 'Inter', 'Segoe UI', system-ui, -apple-system, Roboto, Arial, sans-serif;
  --raio: 20px;
  --trilho: 292px;
}
* { box-sizing: border-box; }
[hidden] { display: none !important; }
html { -webkit-text-size-adjust: 100%; font-size: 16px; color-scheme: dark; }
@media (min-width: 1700px) { html { font-size: 17px; } }
@media (min-width: 2400px) { html { font-size: 19px; } }
body {
  margin: 0;
  min-height: 100vh;
  font-family: var(--texto);
  color: var(--tinta);
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
  background-color: var(--fundo-pe);
  /* o degrade vai do alto ao pe da pagina inteira, nao so da primeira tela */
  background-image: linear-gradient(180deg, var(--fundo-topo) 0, #0A1030 420px, var(--fundo-meio) 1000px,
                    #05070F 72%, var(--fundo-pe) 100%);
}
h1, h2, h3, h4, p { margin: 0; }
button { font-family: inherit; }
.ic { width: 18px; height: 18px; flex: 0 0 auto; display: inline-block; vertical-align: middle; }
.dinheiro { color: var(--dinheiro); font-weight: 650; white-space: nowrap; font-variant-numeric: tabular-nums; }
:focus-visible { outline: 2px solid #BDB2FF; outline-offset: 2px; border-radius: 6px; }

/* ---------------- a aurora ---------------- */
.ceu {
  position: absolute; left: 0; right: 0; top: 0; height: 860px;
  overflow: hidden; pointer-events: none; z-index: 0;
  -webkit-mask-image: linear-gradient(180deg, #000 0%, #000 38%, transparent 100%);
          mask-image: linear-gradient(180deg, #000 0%, #000 38%, transparent 100%);
}
.ceu .luz { position: absolute; border-radius: 50%; filter: blur(64px); will-change: transform; }
.ceu .l1 { width: 64vw; height: 330px; left: 12vw; top: -80px; opacity: .72;
  background: radial-gradient(closest-side, rgba(52,230,200,.95), rgba(52,230,200,0));
  animation: deriva-a 44s ease-in-out infinite alternate; }
.ceu .l2 { width: 58vw; height: 380px; left: 42vw; top: -30px; opacity: .72;
  background: radial-gradient(closest-side, rgba(128,90,255,.95), rgba(128,90,255,0));
  animation: deriva-b 52s ease-in-out infinite alternate; }
.ceu .l3 { width: 50vw; height: 280px; left: -8vw; top: 40px; opacity: .6;
  background: radial-gradient(closest-side, rgba(60,120,255,.9), rgba(60,120,255,0));
  animation: deriva-c 60s ease-in-out infinite alternate; }
.ceu .l4 { width: 30vw; height: 180px; left: 66vw; top: -60px; opacity: .38;
  background: radial-gradient(closest-side, rgba(236,96,214,.8), rgba(236,96,214,0));
  animation: deriva-a 70s ease-in-out infinite alternate-reverse; }
/* as cortinas: listras verticais finas que dao o desenho de aurora */
.ceu .cortina {
  position: absolute; left: -10%; right: -10%; top: -40px; height: 470px; opacity: .8;
  background: repeating-linear-gradient(90deg, rgba(150,255,230,0) 0 22px, rgba(150,255,230,.07) 22px 24px,
              rgba(150,255,230,0) 24px 51px, rgba(190,170,255,.06) 51px 53px, rgba(190,170,255,0) 53px 80px);
  -webkit-mask-image: radial-gradient(55% 75% at 45% 18%, #000 0%, transparent 100%);
          mask-image: radial-gradient(55% 75% at 45% 18%, #000 0%, transparent 100%);
  transform: skewX(-14deg); filter: blur(1.5px);
  animation: cortina 30s ease-in-out infinite alternate;
}
.ceu .estrelas {
  position: absolute; inset: 0; opacity: .5;
  background-image:
    radial-gradient(1px 1px at 8% 22%, rgba(255,255,255,.8), transparent 60%),
    radial-gradient(1px 1px at 23% 64%, rgba(255,255,255,.6), transparent 60%),
    radial-gradient(1.2px 1.2px at 37% 12%, rgba(255,255,255,.7), transparent 60%),
    radial-gradient(1px 1px at 52% 44%, rgba(255,255,255,.5), transparent 60%),
    radial-gradient(1px 1px at 68% 18%, rgba(255,255,255,.7), transparent 60%),
    radial-gradient(1.3px 1.3px at 81% 58%, rgba(255,255,255,.6), transparent 60%),
    radial-gradient(1px 1px at 93% 30%, rgba(255,255,255,.7), transparent 60%),
    radial-gradient(1px 1px at 15% 88%, rgba(255,255,255,.4), transparent 60%),
    radial-gradient(1px 1px at 74% 82%, rgba(255,255,255,.4), transparent 60%);
}
@keyframes deriva-a { from { transform: translate3d(-5vw,0,0) rotate(-9deg) scale(1); }
                      to   { transform: translate3d(6vw,34px,0) rotate(-2deg) scale(1.14); } }
@keyframes deriva-b { from { transform: translate3d(4vw,-10px,0) rotate(6deg) scale(1.05); }
                      to   { transform: translate3d(-7vw,26px,0) rotate(-3deg) scale(.95); } }
@keyframes deriva-c { from { transform: translate3d(0,0,0) scale(.95); }
                      to   { transform: translate3d(9vw,-18px,0) scale(1.12); } }
@keyframes cortina  { from { transform: skewX(-14deg) translateX(-2%); opacity: .6; }
                      to   { transform: skewX(-8deg) translateX(3%); opacity: .9; } }

/* ---------------- estrutura: menu lateral + palco ---------------- */
.app { position: relative; z-index: 1; display: grid; grid-template-columns: var(--trilho) minmax(0, 1fr); min-height: 100vh; }
.trilho {
  position: sticky; top: 0; height: 100vh; overflow-y: auto;
  display: flex; flex-direction: column; gap: 20px; padding: 26px 18px 22px;
  background: linear-gradient(180deg, rgba(16,20,50,.62), rgba(8,10,26,.78));
  border-right: 1px solid var(--vidro-borda);
  -webkit-backdrop-filter: blur(20px) saturate(140%); backdrop-filter: blur(20px) saturate(140%);
}
.marca { display: flex; align-items: center; gap: 12px; padding: 0 4px; }
.marca-sinal {
  width: 38px; height: 38px; border-radius: 12px; flex: 0 0 auto;
  background: conic-gradient(from 210deg, #34E6C8, #6D7BFF, #B36BFF, #34E6C8);
  box-shadow: 0 0 24px rgba(120,110,255,.55), inset 0 0 0 1px rgba(255,255,255,.25);
  display: grid; place-items: center; font-family: var(--display); font-weight: 700; font-size: 1.05rem; color: #0B0F2A;
}
.marca-nome { font-family: var(--display); font-weight: 600; font-size: .9rem; letter-spacing: -.01em; line-height: 1.2; }
.marca-sub { font-size: .76rem; color: var(--tinta-3); margin-top: 2px; }
.data-ref {
  padding: 12px 14px; border-radius: 14px; background: rgba(255,255,255,.045); border: 1px solid var(--vidro-borda);
  font-size: .82rem; line-height: 1.35;
}
.data-ref span { display: block; color: var(--tinta-3); font-size: .72rem; letter-spacing: .12em; text-transform: uppercase; font-weight: 700; }
.data-ref strong { font-weight: 650; color: var(--tinta); }
.busca { position: relative; }
.busca-campo {
  display: flex; align-items: center; gap: 8px; padding: 0 10px 0 12px; height: 42px; border-radius: 12px;
  background: rgba(0,0,0,.28); border: 1px solid var(--linha-2); color: var(--tinta-3);
}
.busca-campo:focus-within { border-color: rgba(170,160,255,.7); box-shadow: 0 0 0 3px rgba(149,133,255,.18); }
.busca-campo input {
  flex: 1; min-width: 0; background: none; border: none; outline: none; color: var(--tinta);
  font: inherit; font-size: .9rem;
}
.busca-campo input::placeholder { color: var(--tinta-3); }
.busca-campo kbd {
  font-family: var(--texto); font-size: .7rem; color: var(--tinta-3); border: 1px solid var(--linha-2);
  border-radius: 6px; padding: 1px 6px;
}
.busca-resultados {
  position: absolute; left: 0; right: 0; top: calc(100% + 6px); z-index: 20;
  padding: 6px; border-radius: 14px; background: rgba(18,22,50,.97); border: 1px solid var(--linha-2);
  box-shadow: 0 24px 50px -12px rgba(0,0,0,.8);
}
.br-item { padding: 8px 10px; border-radius: 10px; }
.br-item + .br-item { border-top: 1px solid var(--linha); }
.br-nome { font-weight: 650; font-size: .88rem; }
.br-acoes { display: flex; flex-wrap: wrap; gap: 4px 14px; margin-top: 3px; font-size: .82rem; }
.br-vazio { padding: 10px; font-size: .84rem; color: var(--tinta-3); }

.menu { display: flex; flex-direction: column; gap: 4px; }
.menu-titulo { font-size: .7rem; letter-spacing: .14em; text-transform: uppercase; color: var(--tinta-3); font-weight: 700; padding: 0 12px 6px; }
.aba {
  position: relative; display: flex; align-items: center; gap: 12px; padding: 11px 12px; border-radius: 13px;
  color: var(--tinta-2); text-decoration: none; font-weight: 600; font-size: .94rem; border: 1px solid transparent;
  white-space: nowrap; transition: background .15s, color .15s;
}
.aba-curta { display: none; }
.aba .ic { width: 20px; height: 20px; flex: 0 0 auto; }
.aba:hover { background: rgba(255,255,255,.055); color: var(--tinta); }
.aba.ativa {
  color: #fff;
  background: linear-gradient(90deg, rgba(149,133,255,.30), rgba(92,141,255,.10));
  border-color: rgba(170,160,255,.40);
  box-shadow: 0 0 28px -8px rgba(149,133,255,.75), var(--vidro-luz);
}
.aba.ativa::before {
  content: ""; position: absolute; left: -18px; top: 9px; bottom: 9px; width: 3px; border-radius: 3px;
  background: linear-gradient(180deg, #A897FF, #5CD2FF); box-shadow: 0 0 14px #9585FF;
}
.aba-conta {
  margin-left: auto; font-size: .74rem; font-weight: 700; padding: 2px 9px; border-radius: 999px;
  background: rgba(255,255,255,.08); color: var(--tinta-2); font-variant-numeric: tabular-nums;
}
.aba.ativa .aba-conta { background: rgba(255,255,255,.16); color: #fff; }
.trilho-pe { margin-top: auto; font-size: .76rem; color: var(--tinta-3); line-height: 1.5; padding: 0 6px; }
.trilho-pe strong { color: var(--tinta-2); font-weight: 600; }
.copia-selo {
  display: block; margin-bottom: 8px; padding: 8px 10px; border-radius: 10px;
  background: var(--atencao-f); color: var(--atencao); font-weight: 650;
}

.palco { min-width: 0; padding: 30px clamp(20px, 3vw, 56px) 40px; }
.palco-interno { max-width: 96rem; margin: 0 auto; }
.tela { display: block; }
.js .tela { display: none; }
.js .tela.ativa { display: block; animation: surge .35s ease-out both; }
@keyframes surge { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }

/* ---------------- o alto de cada tela ---------------- */
.topo-tela { margin-bottom: 30px; }
.migalha { display: flex; align-items: center; gap: 12px; margin-bottom: 18px; }
.migalha .passo {
  font-family: var(--display); font-size: .7rem; font-weight: 500; padding: 4px 10px; border-radius: 999px;
  border: 1px solid var(--linha-2); color: var(--tinta-2); white-space: nowrap;
}
.migalha h1 { font-family: var(--display); font-weight: 500; font-size: clamp(1.02rem, 1.25vw, 1.32rem); letter-spacing: -.01em; }
.heroi-grade { display: grid; grid-template-columns: minmax(0, 1.05fr) minmax(0, 1fr); gap: 34px; align-items: start; }
.heroi-linha { display: grid; grid-template-columns: minmax(0, 1.55fr) minmax(0, 1fr); gap: 22px 52px; align-items: center; }
.heroi-lado { min-width: 0; }
.heroi-lado .heroi-texto { margin-top: 0; }
.heroi-lado .medidor-carteira { margin: 0 0 16px; }
.heroi-lado .legenda-faixas { margin-top: 14px; }
.stats-coluna { display: grid; border-left: 1px solid var(--linha-2); padding-left: 30px; }
.stats-coluna .stat { padding: 9px 0; }
.stats-coluna .stat + .stat { border-top: 1px solid var(--linha); }
.heroi { container-type: inline-size; min-width: 0; }
.heroi-rotulo { font-size: .76rem; font-weight: 700; letter-spacing: .16em; text-transform: uppercase; color: var(--tinta-2); }
.heroi-numero {
  font-family: var(--display); font-weight: 700; line-height: 1.02; letter-spacing: -.04em; white-space: nowrap;
  font-size: clamp(2.2rem, 4.6vw, 6rem);
  font-size: clamp(2.1rem, 9.4cqi, 7rem);
  margin: .14em 0 .12em; color: #fff;
  text-shadow: 0 0 30px rgba(150,160,255,.55), 0 0 90px rgba(120,90,255,.35);
}
@supports ((-webkit-background-clip: text) or (background-clip: text)) {
  .heroi-numero {
    background: linear-gradient(180deg, #FFFFFF 25%, #C9D4FF 100%);
    -webkit-background-clip: text; background-clip: text; color: transparent; text-shadow: none;
    filter: drop-shadow(0 0 26px rgba(150,160,255,.5)) drop-shadow(0 0 70px rgba(120,90,255,.3));
  }
}
.heroi-numero .unidade { font-size: .42em; font-weight: 500; letter-spacing: -.01em; margin-left: .18em; }
.heroi-sub { font-size: 1.05rem; color: var(--tinta-2); }
.heroi-sub strong { color: var(--tinta); font-weight: 650; }
.heroi-texto { margin-top: 14px; max-width: 62ch; font-size: .98rem; line-height: 1.6; color: var(--tinta-2); }
.heroi-texto strong { color: var(--tinta); }
.heroi-stats { display: flex; flex-wrap: wrap; gap: 14px 34px; margin-top: 22px; padding-top: 18px; border-top: 1px solid var(--linha); }
.stat .numero { font-family: var(--display); font-weight: 600; font-size: 1.2rem; letter-spacing: -.02em; font-variant-numeric: tabular-nums; }
.stat .rotulo { font-size: .8rem; color: var(--tinta-3); margin-top: 2px; }
.stat.menos .numero { color: var(--grave); }
.acoes-heroi { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 22px; }

/* ---------------- vidro ---------------- */
.vidro, .descoberta, .linha, .risco-linha, .painel, .regua-cliente, .relatorio-doc, .regua-item {
  background: var(--vidro);
  border: 1px solid var(--vidro-borda);
  box-shadow: var(--vidro-luz), var(--sombra);
}
.descoberta, .painel-heroi {
  -webkit-backdrop-filter: blur(18px) saturate(150%); backdrop-filter: blur(18px) saturate(150%);
}
/* borda de luz: so no que e principal */
.luz-borda { position: relative; }
.luz-borda::after {
  content: ""; position: absolute; inset: -1px; border-radius: inherit; padding: 1px; pointer-events: none;
  background: linear-gradient(125deg, rgba(170,150,255,.95), rgba(92,210,255,.7) 45%, rgba(255,255,255,.08) 75%);
  -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
  -webkit-mask-composite: xor; mask-composite: exclude;
}

/* ---------------- cores de situacao ---------------- */
.cor-grave { --c: var(--grave); --cf: var(--grave-f); }
.cor-atencao { --c: var(--atencao); --cf: var(--atencao-f); }
.cor-ok { --c: var(--ok); --cf: var(--ok-f); }
.cor-neutro { --c: var(--neutro); --cf: var(--neutro-f); }
.situacao, .selo, .selo-tom, .indice-chip {
  display: inline-flex; align-items: center; gap: 6px; padding: 4px 10px 4px 8px; border-radius: 999px;
  font-size: .72rem; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; white-space: nowrap;
  color: var(--c, var(--tinta-2)); background: var(--cf, rgba(255,255,255,.07));
}
.situacao .ic, .selo .ic { width: 14px; height: 14px; }
.linha .so-registrado { display: none; }
.linha.recente .so-registrado { display: inline-flex; }
.linha.recente .so-registrado ~ .situacao, .linha.recente .situacao:has(~ .so-registrado) { display: none; }
.ponto { width: 8px; height: 8px; border-radius: 50%; background: var(--c); flex: 0 0 auto; }

/* ---------------- descobertas ---------------- */
.descobertas { display: grid; grid-template-columns: minmax(0, 1.3fr) minmax(0, 1fr) minmax(0, 1fr); gap: 14px; margin-top: 28px; }
.descobertas .desc-titulo { grid-column: 1 / -1; }
.desc-titulo { display: flex; align-items: center; gap: 8px; font-size: .74rem; font-weight: 700; letter-spacing: .16em; text-transform: uppercase; color: var(--tinta-2); }
.descoberta { border-radius: 18px; padding: 16px 20px 15px; display: flex; flex-direction: column; }
.descoberta .desc-links { margin-top: auto; }
.descoberta.principal { box-shadow: var(--vidro-luz), 0 0 0 1px rgba(160,150,255,.12), 0 20px 60px -24px rgba(120,100,255,.8); }
.desc-rotulo { display: inline-flex; align-items: center; gap: 7px; font-size: .72rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: var(--c); }
.desc-rotulo .ic { width: 16px; height: 16px; }
.descoberta p { margin: 7px 0 9px; font-size: .95rem; line-height: 1.5; }
.desc-links { display: flex; flex-wrap: wrap; gap: 4px 16px; font-size: .84rem; }

/* ---------------- botoes ---------------- */
.botao {
  display: inline-flex; align-items: center; gap: 8px; font-size: .86rem; font-weight: 650; line-height: 1.2;
  padding: 10px 16px; border-radius: 12px; border: 1px solid var(--linha-2); background: rgba(255,255,255,.045);
  color: var(--tinta); cursor: pointer; text-decoration: none; white-space: nowrap;
  transition: background .15s, border-color .15s, box-shadow .2s, filter .15s, transform .08s;
}
.botao .ic { width: 17px; height: 17px; }
.botao:hover { background: rgba(255,255,255,.09); border-color: rgba(255,255,255,.26); }
.botao:active { transform: translateY(1px); }
.botao.primario {
  border-color: transparent; color: #fff;
  background: linear-gradient(135deg, #A08FFF 0%, #7280FF 48%, #4FA0FF 100%);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.35), 0 8px 20px -10px rgba(123,110,255,.8);
}
.linha.primeiro .botao.primario, .acoes-heroi .botao.primario, .regua-cliente .botao.primario {
  box-shadow: inset 0 1px 0 rgba(255,255,255,.35), 0 10px 26px -8px rgba(123,110,255,.85), 0 0 32px -10px rgba(90,170,255,.9);
}
.botao.primario:hover { filter: brightness(1.1); box-shadow: inset 0 1px 0 rgba(255,255,255,.4), 0 12px 32px -8px rgba(123,110,255,1), 0 0 40px -8px rgba(90,170,255,1); }
.botao.primario.feito {
  background: var(--ok-f); color: var(--ok); box-shadow: none; border-color: rgba(107,227,172,.35); cursor: default; filter: none;
}
.botao.fantasma { background: transparent; border-color: transparent; color: var(--tinta-2); padding: 10px 12px; }
.botao.fantasma:hover { background: rgba(255,255,255,.07); color: var(--tinta); }
.link {
  background: none; border: none; padding: 0; cursor: pointer; font: inherit; font-weight: 650;
  color: var(--acao-texto); text-decoration: none; display: inline-flex; align-items: center; gap: 5px;
}
.link:hover { color: #fff; text-decoration: underline; text-underline-offset: 3px; }
.link .ic { width: 14px; height: 14px; }

/* ---------------- barra de ferramentas e chips ---------------- */
.secao-cabeca { display: flex; flex-wrap: wrap; align-items: baseline; gap: 6px 16px; margin: 6px 0 14px; }
.secao-cabeca h2 { font-family: var(--display); font-weight: 600; font-size: 1.2rem; letter-spacing: -.01em; }
.secao-cabeca p { color: var(--tinta-2); font-size: .9rem; }
.ferramentas { display: flex; flex-wrap: wrap; align-items: center; gap: 10px 14px; margin-bottom: 16px; }
.chips { display: flex; flex-wrap: wrap; gap: 8px; }
.chip {
  display: inline-flex; align-items: center; gap: 8px; padding: 7px 13px; border-radius: 999px;
  border: 1px solid var(--linha-2); background: rgba(255,255,255,.035); color: var(--tinta-2);
  font-size: .84rem; font-weight: 600; cursor: pointer; transition: background .15s, border-color .15s;
}
.chip b { color: var(--tinta); font-weight: 700; font-variant-numeric: tabular-nums; }
.chip:hover { background: rgba(255,255,255,.07); }
.chip.ativo { background: rgba(149,133,255,.2); border-color: rgba(170,160,255,.6); color: #fff; }
.ordem { margin-left: auto; display: flex; align-items: center; gap: 9px; font-size: .84rem; color: var(--tinta-3); }
select {
  font: inherit; font-size: .86rem; color: var(--tinta); cursor: pointer; appearance: none; -webkit-appearance: none;
  padding: 8px 34px 8px 12px; border-radius: 11px; border: 1px solid var(--linha-2); background-color: rgba(255,255,255,.05);
  background-image: linear-gradient(45deg, transparent 50%, var(--tinta-2) 50%), linear-gradient(135deg, var(--tinta-2) 50%, transparent 50%);
  background-position: calc(100% - 17px) 52%, calc(100% - 12px) 52%; background-size: 5px 5px; background-repeat: no-repeat;
}
select option { background: #121633; color: var(--tinta); }
.legenda-regua { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 16px; font-size: .8rem; color: var(--tinta-3); margin: -4px 0 14px; }
.legenda-regua i { display: inline-block; vertical-align: middle; margin-right: 6px; }
.legenda-regua .lg-faixa { width: 26px; height: 10px; border-radius: 5px; background: rgba(255,255,255,.16); }
.legenda-regua .lg-faixa.dinheiro-faixa { background: linear-gradient(90deg, #4FA8F0, var(--dinheiro)); }
.legenda-regua .lg-ponto { width: 11px; height: 11px; border-radius: 50%; background: var(--tinta); box-shadow: 0 0 8px rgba(255,255,255,.5); }
.lista-vazia { padding: 26px; text-align: center; color: var(--tinta-3); border: 1px dashed var(--linha-2); border-radius: var(--raio); }

/* ---------------- fila de Hoje ---------------- */
.fila { display: flex; flex-direction: column; gap: 14px; }
.linha { border-radius: var(--raio); padding: 20px 22px 16px; transition: border-color .2s, box-shadow .2s, opacity .2s; }
.linha:hover { border-color: rgba(255,255,255,.19); }
.linha.primeiro { border-color: rgba(170,160,255,.42); box-shadow: var(--vidro-luz), 0 0 0 1px rgba(149,133,255,.10), 0 22px 60px -30px rgba(120,100,255,.9); }
.linha.recente { opacity: .8; }
.linha.destaque { border-color: rgba(140,224,255,.8); box-shadow: 0 0 0 3px rgba(140,224,255,.2), var(--sombra); }
.linha-grade { display: grid; grid-template-columns: 44px minmax(0, 1fr) 340px 200px; gap: 4px 26px; align-items: start; }
.posicao {
  font-family: var(--display); font-weight: 600; font-size: 1.1rem; width: 42px; height: 42px; border-radius: 13px;
  display: grid; place-items: center; background: rgba(255,255,255,.07); border: 1px solid var(--linha-2); color: #fff;
}
.linha.primeiro .posicao {
  background: linear-gradient(135deg, #A08FFF, #4FA0FF); border-color: transparent;
  box-shadow: 0 0 24px rgba(140,125,255,.7), inset 0 1px 0 rgba(255,255,255,.35);
}
.linha-topo { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; }
.nome { font-size: 1.14rem; font-weight: 700; letter-spacing: -.01em; }
.meta { font-size: .8rem; color: var(--tinta-3); margin-top: 3px; }
.explicacao { margin-top: 9px; font-size: .97rem; line-height: 1.55; color: var(--tinta); max-width: 74ch; }
.contato-bloco { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 12px; margin-top: 10px; font-size: .84rem; color: var(--tinta-2); }
.contato-bloco .ic { width: 15px; height: 15px; color: var(--tinta-3); }
.contato-nome { font-weight: 650; color: var(--tinta); }
.contato-fone { font-weight: 700; color: var(--tinta); font-variant-numeric: tabular-nums; letter-spacing: .01em; }
.ultimo-contato { color: var(--tinta-3); }
.aviso-registro { color: var(--ok); }
.aviso-registro .contato-nome { color: var(--ok); }
.linha-grafico { padding-top: 2px; min-width: 0; }
.regua-atraso { width: 100%; height: auto; display: block; overflow: visible; }
.rg-trilho { stroke: rgba(255,255,255,.14); stroke-width: 2; stroke-linecap: round; }
.rg-zero { stroke: rgba(255,255,255,.28); stroke-width: 1.4; }
.rg-faixa { fill: rgba(255,255,255,.16); }
.rg-media { stroke: rgba(255,255,255,.55); stroke-width: 1.6; }
.rg-titulo { fill: #0B0F24; stroke: var(--tinta-2); stroke-width: 1.4; }
.rg-hoje { fill: var(--c); stroke: #0B0F24; stroke-width: 2.5; filter: drop-shadow(0 0 6px var(--c)); }
.rg-valor { fill: var(--tinta); font-size: 14px; font-weight: 700; font-family: var(--texto); }
.rg-rotulo { fill: var(--tinta-2); font-size: 13px; font-family: var(--texto); }
.valor { text-align: right; font-family: var(--display); font-weight: 600; font-size: 1.38rem; letter-spacing: -.025em; white-space: nowrap; font-variant-numeric: tabular-nums; }
.valor small { display: block; font-family: var(--texto); font-size: .78rem; font-weight: 500; letter-spacing: 0; color: var(--tinta-3); margin-top: 5px; }
.acoes { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 8px; margin-top: 14px; padding-left: 68px; }
.acoes .separa { flex: 1; }
.dobra {
  margin: 14px 0 2px 68px; padding-top: 16px; border-top: 1px solid var(--linha);
  display: grid; grid-template-columns: minmax(0, 1.25fr) minmax(0, 1fr); gap: 20px 34px;
}
.js .dobra:not(.aberta) { display: none; }
.dobra h4, .ficha h4 { font-size: .72rem; font-weight: 700; letter-spacing: .14em; text-transform: uppercase; color: var(--tinta-3); margin-bottom: 8px; }
.tabela-numeros { width: 100%; border-collapse: collapse; font-size: .84rem; }
.tabela-numeros td { padding: 5px 12px 5px 0; vertical-align: top; border-bottom: 1px solid rgba(255,255,255,.05); }
.tabela-numeros td:first-child { color: var(--tinta-3); }
.tabela-numeros td:last-child { font-variant-numeric: tabular-nums; }

/* ---------------- tabelas ---------------- */
.rolagem-tabela { overflow-x: auto; -webkit-overflow-scrolling: touch; max-width: 100%; }
.tabela-titulos { width: 100%; min-width: 360px; border-collapse: collapse; font-size: .85rem; }
.tabela-titulos.ampla { min-width: 760px; }
.tabela-titulos th, .tabela-semanas th, .tabela-registros th {
  text-align: left; font-size: .68rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
  color: var(--tinta-3); padding: 6px 14px 8px 0; border-bottom: 1px solid var(--linha-2);
}
.tabela-titulos td { padding: 8px 14px 8px 0; border-bottom: 1px solid rgba(255,255,255,.055); white-space: nowrap; font-variant-numeric: tabular-nums; }
.tabela-titulos td.num { font-family: ui-monospace, 'Cascadia Mono', Consolas, monospace; font-size: .8rem; color: var(--tinta-2); }
.tabela-titulos th.dir, .tabela-titulos td.dir, .tabela-titulos th.sit, .tabela-titulos td.atraso { text-align: right; }
.tabela-titulos th:last-child, .tabela-titulos td:last-child { padding-right: 0; }
.tabela-titulos tr.venceu td.atraso { color: var(--grave); font-weight: 650; }
.tabela-titulos tr.a-vencer td.atraso { color: var(--tinta-3); }
.tabela-titulos tfoot td { font-weight: 700; border-bottom: none; padding-top: 10px; }
.tabela-titulos td.chance { color: var(--tinta-2); }
.tabela-titulos tr.improvavel td.chance { color: var(--grave); font-weight: 650; }
.tabela-titulos .baixa { font-size: .72rem; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; margin-left: 4px; }

/* ---------------- Risco ---------------- */
.medidor-carteira { margin-top: 18px; max-width: 640px; }
.medidor-trilho { height: 12px; border-radius: 999px; background: rgba(255,255,255,.08); overflow: hidden; border: 1px solid var(--linha); }
.medidor-cheio { height: 100%; border-radius: 999px; background: linear-gradient(90deg, #4FA8F0, var(--dinheiro)); box-shadow: 0 0 16px rgba(124,198,255,.6); }
.medidor-rotulos { display: flex; justify-content: space-between; gap: 12px; font-size: .8rem; color: var(--tinta-3); margin-top: 7px; }
.legenda-faixas { display: flex; flex-wrap: wrap; gap: 6px 14px; font-size: .78rem; color: var(--tinta-3); }
.legenda-faixas span { display: inline-flex; align-items: center; gap: 6px; }
.lista-risco { display: flex; flex-direction: column; gap: 12px; }
.risco-linha {
  display: grid; grid-template-columns: 34px minmax(0, 1fr) minmax(200px, 330px) 190px;
  grid-template-areas: "pos corpo barra numeros" ". teaser teaser abrir";
  gap: 6px 24px; align-items: center; width: 100%; text-align: left; font: inherit; color: inherit; cursor: pointer;
  padding: 18px 22px; border-radius: var(--raio); transition: border-color .2s, transform .15s;
}
.risco-linha:hover { border-color: rgba(170,160,255,.45); transform: translateY(-1px); }
.risco-linha > span { display: block; min-width: 0; }
.rl-pos { grid-area: pos; font-family: var(--display); font-weight: 500; color: var(--tinta-3); font-size: .95rem; }
.rl-corpo { grid-area: corpo; }
.rl-nome { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 10px; font-size: 1.06rem; font-weight: 700; }
.rl-barra { grid-area: barra; }
.rl-trilho { display: block; height: 10px; border-radius: 999px; background: rgba(255,255,255,.12); overflow: hidden; }
.rl-cheio { display: block; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #4FA8F0, var(--dinheiro)); }
.rl-barra-rot { display: block; font-size: .78rem; color: var(--tinta-3); margin-top: 6px; }
.rl-numeros { grid-area: numeros; text-align: right; }
.rl-valor { display: block; font-family: var(--display); font-weight: 600; font-size: 1.3rem; letter-spacing: -.025em; font-variant-numeric: tabular-nums; }
.rl-valor small { font-family: var(--texto); font-size: .76rem; font-weight: 500; color: var(--tinta-3); letter-spacing: 0; }
.rl-numeros .indice-chip { margin-top: 6px; }
.rl-teaser { grid-area: teaser; font-size: .9rem; color: var(--tinta-2); line-height: 1.5; }
.rl-abrir { grid-area: abrir; justify-self: end; font-size: .84rem; font-weight: 650; color: var(--acao-texto); display: inline-flex !important; align-items: center; gap: 6px; }
.rl-abrir .ic { width: 15px; height: 15px; }
.risco-linha:hover .rl-abrir { color: #fff; }

/* ---------------- ficha (painel lateral) ---------------- */
.modal-fundo {
  display: none; position: fixed; inset: 0; z-index: 50; background: rgba(3,4,14,.6);
  -webkit-backdrop-filter: blur(6px); backdrop-filter: blur(6px);
}
.modal-fundo.aberto { display: block; }
.ficha {
  position: absolute; top: 0; right: 0; height: 100%; width: min(880px, 100%); overflow-y: auto;
  padding: 26px 34px 44px; background: linear-gradient(180deg, rgba(24,28,64,.97), rgba(10,12,30,.98));
  border-left: 1px solid var(--vidro-borda); box-shadow: -40px 0 90px rgba(0,0,0,.6);
  animation: entra-lado .28s ease-out both;
}
@keyframes entra-lado { from { transform: translateX(28px); opacity: .5; } to { transform: none; opacity: 1; } }
.ficha-cabeca { display: flex; align-items: flex-start; gap: 16px; margin-bottom: 18px; }
.ficha-cabeca .migalha { margin-bottom: 8px; }
.ficha h3 { font-family: var(--display); font-weight: 600; font-size: 1.55rem; letter-spacing: -.02em; line-height: 1.15; }
.meta-ficha { color: var(--tinta-3); font-size: .86rem; margin-top: 6px; }
.fechar {
  margin-left: auto; flex: 0 0 auto; width: 40px; height: 40px; border-radius: 12px; cursor: pointer;
  background: rgba(255,255,255,.06); border: 1px solid var(--linha-2); color: var(--tinta); display: grid; place-items: center;
}
.fechar:hover { background: rgba(255,255,255,.12); }
.fechar .ic { width: 18px; height: 18px; }
.ficha-topo {
  display: flex; flex-wrap: wrap; gap: 18px 34px; align-items: flex-end; padding: 18px 20px; border-radius: 16px;
  background: rgba(255,255,255,.045); border: 1px solid var(--vidro-borda);
}
.indice-grande { font-family: var(--display); font-weight: 700; font-size: 2.8rem; line-height: 1; letter-spacing: -.04em; }
.indice-grande span { font-size: .9rem; color: var(--tinta-3); font-weight: 500; letter-spacing: 0; margin-left: 2px; }
.indice-legenda { margin-top: 8px; }
.blocos-topo { flex: 1; min-width: 240px; display: flex; flex-wrap: wrap; gap: 14px 30px; }
.rotulo-topo { font-size: .68rem; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: var(--tinta-3); }
.valor-topo { font-family: var(--display); font-weight: 600; font-size: 1.12rem; letter-spacing: -.02em; margin-top: 4px; font-variant-numeric: tabular-nums; }
.ficha-acoes { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }
.ficha .contato-bloco { margin-top: 14px; }
.alerta {
  margin: 18px 0 8px; padding: 16px 18px 16px 20px; border-radius: 16px; font-size: .95rem; line-height: 1.65;
  background: var(--cf, rgba(149,133,255,.12)); border: 1px solid var(--linha); border-left: 3px solid var(--c, var(--acao));
}
.ficha h4 { margin: 28px 0 10px; }
.grafico-ficha, .grafico-previsao { display: block; width: 100%; height: auto; overflow: visible; }
.g-grade { stroke: rgba(255,255,255,.07); stroke-width: 1; }
.g-eixo { fill: var(--tinta-3); font-size: 12px; font-family: var(--texto); }
.grafico-previsao .g-eixo { font-size: 13px; }
.g-eixo.forte { fill: var(--tinta-2); font-weight: 600; }
.g-faixa { fill: rgba(255,255,255,.09); }
.g-media { stroke: rgba(255,255,255,.4); stroke-width: 1.3; }
.g-zero { stroke: rgba(255,255,255,.22); stroke-width: 1; }
.g-anot { fill: var(--tinta-2); font-size: 12px; font-family: var(--texto); font-weight: 600;
  paint-order: stroke; stroke: #151A40; stroke-width: 5px; stroke-linejoin: round; }
.g-fim, .pv-vao-rotulo, .rg-valor { paint-order: stroke; stroke: #11163A; stroke-width: 4px; stroke-linejoin: round; }
.g-linha { fill: none; stroke: rgba(230,234,255,.7); stroke-width: 1.8; stroke-linejoin: round; stroke-linecap: round; }
.g-ponto { fill: #E6EAFF; stroke: #0E1230; stroke-width: 1.5; }
.g-ponto.fora { fill: var(--grave); filter: drop-shadow(0 0 5px rgba(255,142,132,.8)); }
.g-ponto.sazonal { fill: #0E1230; stroke: var(--tinta); stroke-width: 2; }
.g-fim { fill: var(--tinta); font-size: 12px; font-weight: 700; font-family: var(--texto); }
.grafico-legenda { display: flex; flex-wrap: wrap; gap: 6px 18px; font-size: .78rem; color: var(--tinta-3); margin-top: 8px; }
.legenda-item { display: inline-flex; align-items: center; gap: 7px; }
.lg { display: inline-block; flex: 0 0 auto; }
.lg.ponto-normal { width: 9px; height: 9px; border-radius: 50%; background: #E6EAFF; }
.lg.ponto-fora { width: 10px; height: 10px; border-radius: 50%; background: var(--grave); }
.lg.ponto-sazonal { width: 10px; height: 10px; border-radius: 50%; border: 2px solid var(--tinta); }
.lg.faixa { width: 18px; height: 10px; border-radius: 3px; background: rgba(255,255,255,.16); }
.grade-fatos { display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 8px; }
.fato { padding: 11px 13px; border-radius: 13px; background: rgba(255,255,255,.04); border: 1px solid var(--linha); }
.rotulo-fato { font-size: .66rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: var(--tinta-3); }
.valor-fato { font-size: .98rem; font-weight: 650; margin-top: 3px; font-variant-numeric: tabular-nums; }
.fato.largo { grid-column: 1 / -1; }
.fato.largo .valor-fato { font-size: .88rem; font-weight: 450; color: var(--tinta-2); line-height: 1.55; }
.sem-dado { color: var(--tinta-3); font-weight: 500; font-style: italic; }
.sem-grafico { color: var(--tinta-3); font-style: italic; font-size: .9rem; }
.componentes { display: flex; flex-direction: column; gap: 8px; }
.componente { border-radius: 13px; background: rgba(255,255,255,.035); border: 1px solid var(--linha); }
.componente summary {
  list-style: none; cursor: pointer; display: grid; grid-template-columns: minmax(0, 1fr) 140px 96px; gap: 14px;
  align-items: center; padding: 11px 14px; font-size: .9rem; font-weight: 650;
}
.componente summary::-webkit-details-marker { display: none; }
.componente summary .comp-nome::before { content: "▸"; display: inline-block; margin-right: 8px; color: var(--tinta-3); transition: transform .15s; }
.componente[open] summary .comp-nome::before { transform: rotate(90deg); }
.barra-peso { height: 6px; border-radius: 999px; background: rgba(255,255,255,.09); overflow: hidden; }
.barra-peso > span { display: block; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #A08FFF, #5CB8FF); }
.pontos { text-align: right; font-variant-numeric: tabular-nums; color: var(--tinta); white-space: nowrap; }
.pontos.zerado { color: var(--tinta-3); }
.componente .frase { padding: 0 14px 13px 30px; font-size: .88rem; color: var(--tinta-2); line-height: 1.6; }

/* ---------------- Previsao ---------------- */
.painel { border-radius: var(--raio); padding: 20px 22px; }
.painel-heroi { border-radius: 22px; padding: 18px 18px 10px; }
.painel h3 { font-size: 1rem; font-weight: 700; }
.painel .sub { font-size: .86rem; color: var(--tinta-2); margin: 4px 0 12px; line-height: 1.55; }
.pv-faixa { fill: rgba(124,198,255,.14); }
.pv-planilha { fill: none; stroke: rgba(236,238,255,.62); stroke-width: 2; stroke-dasharray: 7 6; stroke-linejoin: round; }
.pv-planilha-ponto { fill: #0E1230; stroke: rgba(236,238,255,.75); stroke-width: 2; }
.pv-esperado { fill: none; stroke: var(--dinheiro); stroke-width: 3; stroke-linejoin: round; filter: drop-shadow(0 0 7px rgba(124,198,255,.7)); }
.pv-esperado-ponto { fill: var(--dinheiro); stroke: #0E1230; stroke-width: 2; }
.pv-vao { stroke: rgba(255,142,132,.75); stroke-width: 2; stroke-dasharray: 3 4; }
.pv-vao-rotulo { fill: var(--grave); font-size: 14px; font-weight: 700; font-family: var(--display); letter-spacing: -.02em; }
.pv-rotulo { fill: var(--tinta-2); font-size: 12.5px; font-family: var(--texto); }
.pv-valor { font-size: 15px; font-weight: 700; font-family: var(--display); letter-spacing: -.02em; }
.pv-valor.planilha { fill: var(--tinta); }
.pv-valor.esperado { fill: var(--dinheiro); }
.avisos { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 12px; margin-bottom: 18px; }
.aviso-escopo {
  padding: 13px 16px; border-radius: 14px; font-size: .88rem; line-height: 1.55; color: var(--tinta-2);
  background: var(--atencao-f); border: 1px solid rgba(247,199,102,.22);
}
.aviso-escopo strong { color: var(--atencao); }
.faixa-semana { display: grid; margin-top: 6px; }
.faixa-semana .stat { padding: 12px 0; }
.faixa-semana .stat + .stat { border-top: 1px solid var(--linha); }
.faixa-semana .stat .numero { font-size: 1.45rem; }
.duas-colunas { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 18px; margin-bottom: 18px; }
.tabela-semanas { width: 100%; border-collapse: collapse; font-size: .9rem; }
.tabela-semanas td { padding: 10px 14px 10px 0; border-bottom: 1px solid rgba(255,255,255,.055); font-variant-numeric: tabular-nums; white-space: nowrap; }
.tabela-semanas th.dir, .tabela-semanas td.dir { text-align: right; }
.tabela-semanas th:last-child, .tabela-semanas td:last-child { padding-right: 0; }
.tabela-semanas tfoot td { font-weight: 700; border-bottom: none; border-top: 1px solid var(--linha-2); padding-top: 12px; }
.tabela-semanas .menos { color: var(--grave); font-weight: 650; }
.semana-rotulo { font-weight: 650; }
.semana-datas { font-size: .8rem; color: var(--tinta-3); font-weight: 400; }
.nota { margin-top: 12px; font-size: .86rem; color: var(--tinta-2); line-height: 1.6; }

/* ---------------- blocos de metodo (fechados) ---------------- */
details.metodo { margin-top: 18px; border-radius: 16px; border: 1px solid var(--linha); background: rgba(255,255,255,.025); }
details.metodo > summary {
  list-style: none; cursor: pointer; padding: 13px 18px; font-size: .88rem; font-weight: 650; color: var(--tinta-2);
  display: flex; align-items: center; gap: 9px;
}
details.metodo > summary::-webkit-details-marker { display: none; }
details.metodo > summary .ic { width: 17px; height: 17px; color: var(--tinta-3); }
details.metodo > summary::after { content: "▸"; margin-left: auto; color: var(--tinta-3); transition: transform .15s; }
details.metodo[open] > summary::after { transform: rotate(90deg); }
details.metodo > summary:hover { color: var(--tinta); }
.metodo-corpo { padding: 0 18px 16px; font-size: .9rem; color: var(--tinta-2); line-height: 1.65; }
.metodo-corpo p + p { margin-top: 8px; }
.metodo-corpo strong { color: var(--tinta); }

/* ---------------- Regua ---------------- */
.regua-layout { display: grid; grid-template-columns: minmax(270px, 350px) minmax(0, 1fr); gap: 20px; align-items: start; }
.regua-lista { display: flex; flex-direction: column; gap: 8px; position: sticky; top: 20px; }
.regua-item {
  display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 4px 12px; align-items: center;
  width: 100%; text-align: left; font: inherit; color: inherit; cursor: pointer; padding: 13px 14px; border-radius: 15px;
  transition: border-color .15s, background .15s;
}
.regua-item:hover { border-color: rgba(255,255,255,.22); }
.regua-item.ativo {
  background: linear-gradient(90deg, rgba(149,133,255,.26), rgba(92,141,255,.08));
  border-color: rgba(170,160,255,.55); box-shadow: 0 0 30px -10px rgba(149,133,255,.8), var(--vidro-luz);
}
.ri-nome { font-weight: 700; font-size: .93rem; }
.ri-valor { font-weight: 650; font-size: .86rem; font-variant-numeric: tabular-nums; text-align: right; }
.ri-linha2 { grid-column: 1 / -1; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.ri-linha2 .selo-tom { font-size: .64rem; padding: 3px 8px; }
.ri-feito { font-size: .74rem; color: var(--ok); font-weight: 650; display: inline-flex; align-items: center; gap: 4px; }
.ri-feito .ic { width: 13px; height: 13px; }
.regua-cliente { border-radius: 22px; padding: 22px 26px 24px; }
.js .regua-cliente:not(.ativo) { display: none; }
.regua-cabecalho { display: flex; flex-wrap: wrap; align-items: flex-start; gap: 12px 20px; }
.regua-cabecalho .nome { font-family: var(--display); font-weight: 600; font-size: 1.3rem; letter-spacing: -.02em; flex: 1; min-width: 220px; }
.regua-cabecalho .nome .selo-tom { font-family: var(--texto); margin-left: 10px; vertical-align: 3px; }
.regua-cabecalho .meta { font-family: var(--texto); font-size: .84rem; margin-top: 8px; letter-spacing: 0; font-weight: 450; }
.regua-cabecalho .valor small { text-transform: none; }
.motivo-tom { margin-top: 16px; padding: 13px 16px; border-radius: 14px; background: rgba(255,255,255,.04); border: 1px solid var(--linha); font-size: .92rem; line-height: 1.6; color: var(--tinta-2); }
.motivo-tom strong { color: var(--tinta); }
.troca-tom { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 12px; margin-top: 14px; font-size: .86rem; }
.troca-tom label { color: var(--tinta-3); font-weight: 650; }
.aviso-manual { display: none; align-items: center; gap: 6px; color: var(--atencao); font-weight: 650; font-size: .8rem; }
.aviso-tons { flex-basis: 100%; color: var(--tinta-3); font-size: .82rem; }
.canais { display: flex; gap: 4px; margin-top: 16px; padding: 4px; border-radius: 13px; background: rgba(0,0,0,.25); border: 1px solid var(--linha); width: max-content; max-width: 100%; overflow-x: auto; }
.canais button {
  font: inherit; font-size: .84rem; font-weight: 650; padding: 8px 14px; border-radius: 10px; border: none;
  background: transparent; color: var(--tinta-3); cursor: pointer; white-space: nowrap;
}
.canais button:hover { color: var(--tinta); }
.canais button.ativo { background: rgba(255,255,255,.1); color: #fff; box-shadow: var(--vidro-luz); }
.caixa-texto { margin-top: 12px; }
.caixa-texto .assunto { font-size: .86rem; color: var(--tinta-2); padding: 10px 16px; border: 1px solid var(--linha); border-bottom: none; border-radius: 14px 14px 0 0; background: rgba(255,255,255,.04); }
.caixa-texto pre {
  margin: 0; white-space: pre-wrap; word-wrap: break-word; font-family: var(--texto); font-size: .95rem; line-height: 1.65;
  color: var(--tinta); background: rgba(0,0,0,.28); border: 1px solid var(--linha); border-radius: 14px; padding: 16px 18px;
}
.caixa-texto .assunto + pre { border-radius: 0 0 14px 14px; }
.acoes-regua { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 10px; margin-top: 14px; }
.acoes-regua .separa { flex: 1; }
.ok-copiado { display: none; align-items: center; gap: 5px; color: var(--ok); font-weight: 650; font-size: .84rem; }
.painel-registros { margin-top: 22px; }
.pr-cabeca { display: flex; flex-wrap: wrap; align-items: baseline; gap: 6px 14px; }
.pr-cabeca h3 { font-family: var(--display); font-weight: 600; font-size: 1.05rem; }
.pr-resumo { font-size: .84rem; color: var(--tinta-3); }
.tabela-registros { width: 100%; border-collapse: collapse; font-size: .86rem; margin-top: 8px; }
.tabela-registros td { padding: 8px 14px 8px 0; border-bottom: 1px solid rgba(255,255,255,.055); }
.tabela-registros th.dir, .tabela-registros td.dir { text-align: right; }
.tabela-registros th:last-child, .tabela-registros td:last-child { padding-right: 0; }
.origem { display: inline-block; font-size: .64rem; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; padding: 2px 8px; border-radius: 999px; margin-left: 8px; vertical-align: 1px; }
.origem.planilha { background: rgba(255,255,255,.08); color: var(--tinta-2); }
.origem.registrado { background: var(--atencao-f); color: var(--atencao); }
.vazio-registros { font-size: .9rem; color: var(--tinta-3); font-style: italic; margin-top: 10px; }
.achado-ok, .achado-erro, .achado-neutro {
  font-size: .86rem; line-height: 1.6; padding: 11px 14px; border-radius: 12px; margin-top: 10px; border: 1px solid var(--linha);
}
.achado-ok { background: var(--ok-f); color: var(--ok); }
.achado-erro { background: var(--grave-f); color: var(--grave); }
.achado-neutro { background: var(--atencao-f); color: var(--atencao); }
.achado-ok code, .achado-erro code, .achado-neutro code {
  font-family: ui-monospace, 'Cascadia Mono', Consolas, monospace; font-size: .8rem; padding: 2px 6px; border-radius: 6px;
  background: rgba(0,0,0,.3); color: var(--tinta); word-break: break-all;
}

/* ---------------- Relatorio ---------------- */
.relatorio-layout { display: grid; grid-template-columns: minmax(0, 860px) minmax(200px, 280px); gap: 34px; align-items: start; }
.relatorio-indice { position: sticky; top: 24px; display: flex; flex-direction: column; gap: 10px; padding: 18px 20px;
  border-radius: 18px; background: rgba(255,255,255,.03); border: 1px solid var(--linha); font-size: .9rem; }
.relatorio-doc h3 { scroll-margin-top: 20px; }
.relatorio-doc { max-width: 860px; border-radius: 24px; padding: 40px 48px 44px; font-size: 1.04rem; line-height: 1.8; }
.relatorio-doc .rel-titulo { font-family: var(--display); font-weight: 600; font-size: 1.35rem; letter-spacing: -.02em; line-height: 1.25; }
.relatorio-doc .rel-sub { color: var(--tinta-3); font-size: .9rem; margin-top: 6px; }
.relatorio-doc h3 { font-size: .74rem; font-weight: 700; letter-spacing: .16em; text-transform: uppercase; color: var(--acao-texto); margin: 30px 0 8px; }
.relatorio-doc p { color: var(--tinta); }
.relatorio-doc p + p { margin-top: 12px; }
.relatorio-doc ol { margin: 12px 0 0; padding-left: 1.4em; }
.relatorio-doc li { margin-bottom: 8px; }
.relatorio-doc .rel-nota { margin-top: 26px; padding-top: 16px; border-top: 1px solid var(--linha); font-size: .86rem; color: var(--tinta-3); }
.relatorio-doc .dinheiro { font-weight: 700; }

/* ---------------- rodape e aviso flutuante ---------------- */
.rodape {
  margin-top: 46px; padding: 18px 22px; border-radius: 18px; background: rgba(255,255,255,.025); border: 1px solid var(--linha);
  font-size: .8rem; color: var(--tinta-3); line-height: 1.65;
}
.rodape strong { color: var(--tinta-2); }
.rodape summary { list-style: none; cursor: pointer; display: flex; align-items: center; gap: 9px; }
.rodape summary::-webkit-details-marker { display: none; }
.rodape summary::after { content: "▸"; margin-left: auto; transition: transform .15s; }
.rodape details[open] summary::after { transform: rotate(90deg); }
.rodape summary .ic { width: 16px; height: 16px; }
.rodape p { margin-top: 10px; }
.aviso-flutuante {
  position: fixed; left: 50%; bottom: 26px; z-index: 80; transform: translate(-50%, 20px); opacity: 0; pointer-events: none;
  display: flex; align-items: center; gap: 16px; padding: 12px 18px; border-radius: 14px; font-size: .9rem;
  background: rgba(22,26,60,.97); border: 1px solid var(--linha-2); box-shadow: 0 20px 50px -10px rgba(0,0,0,.8);
  transition: opacity .2s, transform .2s;
}
.aviso-flutuante.visivel { opacity: 1; transform: translate(-50%, 0); pointer-events: auto; }

/* ---------------- pecas trazidas da V8 ---------------- */
/* Hoje: o atraso mes a mes, embaixo da regua */
.linha-atraso { width: 100%; height: auto; display: block; overflow: visible; margin-top: 10px; }
.la-faixa { fill: rgba(255,255,255,.09); }
.la-linha { fill: none; stroke: rgba(230,234,255,.72); stroke-width: 1.8; stroke-linejoin: round; stroke-linecap: round; }
.la-fim { fill: var(--c, #E6EAFF); stroke: #0E1230; stroke-width: 1.8; }
.la-rotulo { fill: var(--tinta-3); font-size: 12px; font-family: var(--texto); }
.la-valor { fill: var(--tinta); font-size: 13px; font-weight: 700; font-family: var(--texto); paint-order: stroke; stroke: #11163A; stroke-width: 3px; }
/* Risco: quem deve muito e quem esta piorando */
.painel-dispersao { margin: 0 0 26px; }
.painel-dispersao h3 { font-size: 1.05rem; }
.dispersao-risco { display: block; width: 100%; height: auto; overflow: visible; margin-top: 4px; }
.dp-quadrante { fill: rgba(247,199,102,.06); }
.dp-ponto { fill: rgba(200,206,230,.75); stroke: #0E1230; stroke-width: 1.6; }
.dp-ponto.piorou { fill: var(--atencao); filter: drop-shadow(0 0 6px rgba(247,199,102,.7)); }
.dp-rotulo { fill: var(--tinta-2); font-size: 14.5px; font-family: var(--texto); font-weight: 600; paint-order: stroke; stroke: #131840; stroke-width: 4px; stroke-linejoin: round; }
.dp-rotulo.piorou { fill: var(--atencao); }
.dispersao-risco .g-eixo { font-size: 14px; }
.lg.ponto-piorou { width: 10px; height: 10px; border-radius: 50%; background: var(--atencao); }
.lg.ponto-neutro { width: 10px; height: 10px; border-radius: 50%; background: rgba(200,206,230,.75); }
/* Risco: a tabela embaixo dos cartoes */
.secao-tabela-risco { margin-top: 30px; }
.painel-tabela { padding: 8px 22px 14px; }
.tabela-risco { width: 100%; border-collapse: collapse; font-size: .88rem; }
.tabela-risco th {
  text-align: left; font-size: .68rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase;
  color: var(--tinta-3); padding: 8px 14px 8px 0; border-bottom: 1px solid var(--linha-2); white-space: nowrap;
}
.tabela-risco td { padding: 10px 14px 10px 0; border-bottom: 1px solid rgba(255,255,255,.055); vertical-align: middle; font-variant-numeric: tabular-nums; }
.tabela-risco th.dir, .tabela-risco td.dir { text-align: right; }
.tabela-risco th:nth-child(5), .tabela-risco td:nth-child(5) { padding-left: 18px; }
.tabela-risco tbody tr.tr-risco { cursor: pointer; transition: background .12s; }
.tabela-risco tbody tr.tr-risco:hover { background: rgba(255,255,255,.04); }
.tabela-risco td.pos { color: var(--tinta-3); width: 32px; font-family: var(--display); }
.tabela-risco .cli { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 8px; }
.tabela-risco .cli .link { color: var(--tinta); font-size: .92rem; font-weight: 650; white-space: normal; text-align: left; }
.tabela-risco .cli .link:hover { color: #fff; }
.tabela-risco .sit { color: var(--tinta-2); white-space: nowrap; }
.tabela-risco .sit.nada { color: var(--tinta-3); }
.tabela-risco .indice { display: inline-flex; align-items: center; gap: 8px; white-space: nowrap; }
.tabela-risco .indice b { font-family: var(--display); font-weight: 600; min-width: 22px; text-align: right; }
.indice-barra { width: 64px; height: 5px; border-radius: 999px; background: rgba(255,255,255,.1); overflow: hidden; }
.indice-barra > span { display: block; height: 100%; border-radius: 999px; background: var(--c); }
.em-risco { display: inline-flex; align-items: center; gap: 10px; justify-content: flex-end; }
.em-risco b { font-family: var(--display); font-weight: 600; font-size: .95rem; min-width: 84px; text-align: right; }
.barra-dinheiro { width: 90px; height: 6px; border-radius: 999px; background: rgba(255,255,255,.08); overflow: hidden; }
.barra-dinheiro > span { display: block; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #4FA8F0, var(--dinheiro)); }
.tabela-risco tfoot td { padding: 12px 14px 4px 0; font-weight: 700; border-bottom: none; }
.tabela-risco td.seta { color: var(--tinta-3); width: 18px; padding-right: 0; }
.nota-tabela { font-size: .78rem; color: var(--tinta-3); margin-top: 10px; }
/* Relatorio: os tres numeros da semana, abaixo da data */
.kpis { display: grid; padding: 0; }
.kpi { padding: 16px 22px; min-width: 0; }
.kpi + .kpi { border-left: 1px solid var(--linha); }
.kpi-valor { font-family: var(--display); font-weight: 600; font-size: 1.45rem; line-height: 1.15; letter-spacing: -.025em; white-space: nowrap; font-variant-numeric: tabular-nums; }
.kpi-rotulo { color: var(--tinta-2); font-size: .86rem; margin-top: 5px; }
.kpi.principal .kpi-valor { color: #fff; text-shadow: 0 0 22px rgba(150,160,255,.45); }
.resumo-semana { grid-template-columns: repeat(3, minmax(0, 1fr)); margin-top: 22px; max-width: 1174px; }

/* ---------------- telas menores ---------------- */
@media (max-width: 1500px) {
  .heroi-grade { grid-template-columns: minmax(0, 1fr); }
  .linha-grade { grid-template-columns: 44px minmax(0, 1fr) 270px 180px; }
  .relatorio-layout { grid-template-columns: minmax(0, 1fr); }
  .relatorio-indice { display: none; }
}
@media (max-width: 1250px) {
  .kpi-valor { font-size: 1.2rem; }
  .linha-grade { grid-template-columns: 44px minmax(0, 1fr) 180px; grid-template-areas: "pos corpo valor" "pos grafico grafico"; }
  .linha-grade > .posicao { grid-area: pos; } .linha-grade > .linha-corpo { grid-area: corpo; }
  .linha-grade > .linha-grafico { grid-area: grafico; max-width: 420px; margin-top: 8px; }
  .linha-grade > .valor { grid-area: valor; }
  .risco-linha { grid-template-columns: 30px minmax(0, 1fr) 180px; grid-template-areas: "pos corpo numeros" ". barra barra" ". teaser teaser" ". abrir abrir"; }
  .rl-abrir { justify-self: start; }
  .duas-colunas { grid-template-columns: minmax(0, 1fr); }
}
@media (max-width: 1100px) {
  :root { --trilho: 100%; }
  .app { grid-template-columns: minmax(0, 1fr); }
  .trilho {
    position: relative; height: auto; overflow: visible; border-right: none; border-bottom: 1px solid var(--vidro-borda);
    display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 14px; padding: 16px;
  }
  .trilho .marca { grid-column: 1; }
  .trilho .data-ref { grid-column: 2; }
  .trilho .busca, .trilho .menu { grid-column: 1 / -1; }
  .heroi-linha { grid-template-columns: minmax(0, 1fr); }
  .stats-coluna { border-left: none; padding-left: 0; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 0 18px; }
  .stats-coluna .stat + .stat { border-top: none; }
  .menu { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 6px; }
  .menu-titulo { display: none; }
  .aba { flex-direction: column; gap: 4px; padding: 9px 4px; font-size: .78rem; text-align: center; justify-content: center; }
  .aba-nome, .aba-conta { display: none; }
  .aba-curta { display: block; }
  .aba.ativa::before { left: 10px; right: 10px; top: auto; bottom: -1px; width: auto; height: 3px; }
  .trilho-pe { display: none; }
  .descobertas { grid-template-columns: minmax(0, 1fr); }
  .descobertas .desc-titulo { grid-column: auto; }
  .regua-layout { grid-template-columns: minmax(0, 1fr); }
  .regua-lista { position: static; flex-direction: row; overflow-x: auto; padding-bottom: 4px; }
  .regua-item { flex: 0 0 250px; }
}
@media (max-width: 700px) {
  html { font-size: 15px; }
  .palco { padding: 20px 16px 32px; }
  .trilho { grid-template-columns: minmax(0, 1fr); }
  .trilho .data-ref { grid-column: 1; }
  .stats-coluna { grid-template-columns: minmax(0, 1fr); }
  .ceu { height: 560px; }
  .linha { padding: 16px 16px 14px; }
  .linha-grade { grid-template-columns: 38px minmax(0, 1fr); grid-template-areas: "pos corpo" "pos valor" "grafico grafico"; }
  .linha-grade > .valor { text-align: left; margin-top: 8px; }
  .linha-grade > .linha-grafico { max-width: none; }
  .posicao { width: 36px; height: 36px; font-size: .95rem; }
  .acoes, .dobra { padding-left: 0; margin-left: 0; }
  .dobra { grid-template-columns: minmax(0, 1fr); }
  .ordem { margin-left: 0; }
  .risco-linha { grid-template-columns: minmax(0, 1fr); grid-template-areas: "corpo" "numeros" "barra" "teaser" "abrir"; padding: 16px; }
  .rl-pos { display: none !important; }
  .rl-numeros { text-align: left; }
  .ficha { padding: 20px 16px 36px; }
  .componente summary { grid-template-columns: minmax(0, 1fr) 70px; }
  .componente summary .barra-peso { display: none; }
  .relatorio-doc { padding: 24px 18px; font-size: 1rem; }
  .rolagem-grafico { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  .rolagem-grafico > svg { min-width: 620px; }
  .rolagem-grafico > svg.dispersao-risco { min-width: 900px; }
  .resumo-semana { grid-template-columns: minmax(0, 1fr); }
  .resumo-semana .kpi + .kpi { border-left: none; border-top: 1px solid var(--linha); }
  /* no celular a tabela mostra quem, o indice e o valor em risco; situacao e saldo estao na ficha e no cartao */
  .tabela-risco .indice-barra, .tabela-risco .barra-dinheiro { display: none; }
  .tabela-risco th:nth-child(3), .tabela-risco td:nth-child(3),
  .tabela-risco th:nth-child(4), .tabela-risco td:nth-child(4),
  .tabela-risco th:nth-child(7), .tabela-risco td:nth-child(7) { display: none; }
  .tabela-risco td, .tabela-risco th { padding-right: 8px; }
  .tabela-risco th:nth-child(5), .tabela-risco td:nth-child(5) { padding-left: 0; }
  .tabela-risco th { white-space: normal; }
  .tabela-risco .selo { white-space: normal; border-radius: 6px; }
  .tabela-risco .indice { flex-direction: column; align-items: flex-start; gap: 3px; }
  .tabela-risco .em-risco b { min-width: 0; white-space: nowrap; font-size: .8rem; }
  .tabela-risco .indice b { font-size: .85rem; }
  .painel-tabela { padding: 6px 14px 12px; }
  .semana-datas { display: block; }
  .regua-cliente { padding: 18px 16px; }
  .regua-cabecalho .valor { text-align: left; }
}

@media (prefers-reduced-motion: reduce) {
  .ceu .luz, .ceu .cortina { animation: none; }
  .js .tela.ativa, .ficha { animation: none; }
  * { transition: none !important; }
}
@media (prefers-reduced-transparency: reduce) {
  .trilho, .descoberta, .painel-heroi { -webkit-backdrop-filter: none; backdrop-filter: none; background: #12163A; }
}

/* ---------------- impressao: sai so a tela aberta, em papel branco ---------------- */
@media print {
  @page { margin: 14mm; }
  html, body { background: #fff !important; color: #111 !important; font-size: 11pt; }
  :root { --tinta: #111; --tinta-2: #333; --tinta-3: #555; --linha: #ccc; --linha-2: #999; --dinheiro: #0B5F80;
          --grave: #A3261B; --atencao: #8A5A00; --ok: #1B6B43; --neutro: #444;
          --grave-f: #fff; --atencao-f: #fff; --ok-f: #fff; --neutro-f: #fff; }
  .ceu, .trilho, .ferramentas, .acoes, .acoes-heroi, .acoes-regua, .troca-tom, .canais, .legenda-regua,
  .regua-lista, .painel-registros, .modal-fundo, .aviso-flutuante, details.metodo, .desc-links, .rl-abrir { display: none !important; }
  .rodape { border: none; padding: 8px 0 0; margin-top: 16px; }
  .rodape summary::after { content: ""; }
  .app { display: block; }
  .palco { padding: 0; }
  .js .tela { display: none !important; }
  .js .tela.ativa { display: block !important; animation: none; }
  .heroi-numero { color: #111 !important; background: none !important; filter: none !important; text-shadow: none !important; font-size: 30pt !important; }
  .linha, .risco-linha, .descoberta, .painel, .regua-cliente, .relatorio-doc, .regua-item {
    background: #fff !important; box-shadow: none !important; border: 1px solid #bbb !important; break-inside: avoid;
  }
  .js .dobra { display: block !important; margin-left: 0; }
  .dobra > div + div { margin-top: 12px; }
  .tabela-titulos { min-width: 0; }
  .rolagem-tabela { overflow: visible; }
  #tela-relatorio .topo-tela, .relatorio-indice { display: none !important; }
  .relatorio-doc h3 { color: #3B2F99 !important; }
  .relatorio-layout { display: block; }
  .js .regua-cliente { display: block !important; margin-bottom: 10px; }
  .relatorio-doc { max-width: none; padding: 0; border: none !important; }
  .rg-trilho { stroke: #bbb; } .rg-faixa { fill: #ddd; } .rg-titulo { fill: #fff; stroke: #555; } .rg-hoje { stroke: #fff; filter: none; }
  .g-ponto { fill: #444; } .g-linha { stroke: #444; } .g-faixa { fill: #e6e6e6; } .pv-esperado { filter: none; }
  .pv-planilha { stroke: #666; } .pv-faixa { fill: #d7eef7; }
  .la-linha { stroke: #444; } .la-faixa { fill: #e6e6e6; } .la-valor, .dp-rotulo { stroke: #fff; }
  .dp-ponto { fill: #888; stroke: #fff; } .dp-ponto.piorou { fill: #8A5A00; filter: none; } .dp-quadrante { fill: #f6f1e4; }
  .g-anot, .g-fim, .pv-vao-rotulo, .rg-valor { stroke: #fff; }
  .posicao { color: #111 !important; background: #fff !important; box-shadow: none !important; border: 1px solid #999 !important; }
  a[href]::after { content: ""; }
}
"""


# ---------------------------------------------------------------------------
# Pecas comuns das telas
# ---------------------------------------------------------------------------

TELAS = [
    # chave, nome no menu, titulo da tela, icone
    ("hoje", "Hoje", "Hoje · quem eu ligo primeiro", "hoje", "Hoje"),
    ("risco", "Risco por cliente", "Risco por cliente · onde está o meu dinheiro", "risco", "Risco"),
    ("previsao", "Previsão de caixa", "Previsão de caixa · as próximas 4 semanas", "previsao", "Caixa"),
    ("regua", "Régua de cobrança", "Régua de cobrança · o texto pronto, no tom certo", "regua", "Régua"),
    ("relatorio", "Relatório da semana", "Relatório da semana · para a diretoria", "relatorio", "Relatório"),
]


def render_migalha(chave):
    """'2 de 5' e o nome da tela: quem abre qualquer tela ve que existem outras."""
    pos = [t[0] for t in TELAS].index(chave) + 1
    titulo = TELAS[pos - 1][2]
    return (f'<div class="migalha"><span class="passo">{pos} de {len(TELAS)}</span>'
            f'<h1>{html.escape(titulo)}</h1></div>')


def render_metodo(titulo, corpo_html, aberto=False):
    """O bloco que explica o metodo: existe em toda tela, mas fica fechado."""
    return (f'<details class="metodo"{" open" if aberto else ""}><summary>{icone("info")}{html.escape(titulo)}</summary>'
            f'<div class="metodo-corpo">{corpo_html}</div></details>')


def render_situacao(chave):
    rotulo, cor = SITUACOES_FILA[chave]
    return f'<span class="situacao cor-{cor}">{icone(ICONE_DA_SITUACAO[chave])}{html.escape(rotulo)}</span>'


def render_bloco_contato(cliente, ultimo_contato=None, dias_desde_contato=None):
    """Para quem ligar e em que numero -- a acao concreta da tela."""
    nome_contato = cliente.get("contato") or ""
    telefone = cliente.get("telefone") or ""
    if not nome_contato and not telefone:
        return ""
    pedacos = [f'<div class="contato-bloco">{icone("telefone")}']
    if nome_contato:
        pedacos.append(f'<span class="contato-nome">{html.escape(nome_contato)}</span>')
    if telefone:
        pedacos.append(f'<span class="contato-fone">{html.escape(telefone)}</span>')
    if ultimo_contato is not None and dias_desde_contato is not None:
        quando = "hoje" if dias_desde_contato == 0 else f"há {texto_dias(dias_desde_contato)}"
        pedacos.append(
            f'<span class="ultimo-contato">último contato {quando} · '
            f'{html.escape(ultimo_contato["canal"])} · {html.escape(ultimo_contato["tom"])}</span>'
        )
    else:
        pedacos.append('<span class="ultimo-contato">nunca foi cobrado</span>')
    pedacos.append("</div>")
    return "".join(pedacos)


def render_tabela_titulos(titulos_lista, ref):
    """Nota a nota: numero, vencimento, valor e dias. E o que a pessoa le em voz alta."""
    if not titulos_lista:
        return '<p class="sem-dado">Nenhum título em aberto.</p>'
    linhas = []
    total = 0
    for t in titulos_lista:
        dias = (ref - t["dt_vencimento"]).days
        total += t["valor_centavos"]
        if dias > 0:
            situacao, classe = f"{texto_dias(dias)} em atraso", "venceu"
        elif dias == 0:
            situacao, classe = "vence hoje", "a-vencer"
        else:
            situacao, classe = f"vence em {texto_dias(-dias)}", "a-vencer"
        linhas.append(
            f'<tr class="{classe}"><td class="num">{html.escape(t["numero"])}</td>'
            f'<td>{formatar_data_curta(t["dt_vencimento"])}</td>'
            f'<td class="dir">{formatar_reais_de_centavos(t["valor_centavos"])}</td>'
            f'<td class="atraso">{situacao}</td></tr>'
        )
    return (
        '<div class="rolagem-tabela"><table class="tabela-titulos">'
        '<thead><tr><th>Título</th><th>Vencimento</th><th class="dir">Valor</th><th class="sit">Situação</th></tr></thead>'
        f'<tbody>{"".join(linhas)}</tbody>'
        f'<tfoot><tr><td colspan="2">Total</td><td class="dir">{formatar_reais_de_centavos(total)}</td><td></td></tr></tfoot>'
        '</table></div>'
    )


# ---------------------------------------------------------------------------
# Tela Hoje
# ---------------------------------------------------------------------------

def render_descobertas(descobertas, perfis, cods_regua):
    if not descobertas:
        return ""
    itens = []
    for i, d in enumerate(descobertas):
        if len(d["cods"]) > 1:
            links = [f'<button type="button" class="link" onclick="abrirFicha({c})">{icone("ficha")}'
                     f'{html.escape(perfis[c]["cliente"]["nome"])}</button>' for c in d["cods"]]
        else:
            cod = d["cods"][0]
            links = [f'<button type="button" class="link" onclick="abrirFicha({cod})">{icone("ficha")}Abrir ficha</button>']
            if cod in cods_regua:
                links.append(f'<button type="button" class="link" onclick="irParaRegua({cod})">{icone("regua")}Mensagem pronta</button>')
        principal = " principal luz-borda" if i == 0 else ""
        itens.append(
            f'<article class="descoberta cor-{d["cor"]}{principal}" data-tipo="{d["tipo"]}">'
            f'<div class="desc-rotulo">{icone(ICONE_DA_DESCOBERTA[d["tipo"]])}{html.escape(d["rotulo"])}</div>'
            f'<p>{prosa(d["frase"])}</p><div class="desc-links">{"".join(links)}</div></article>'
        )
    return (f'<section class="descobertas" aria-label="O que a Central descobriu">'
            f'<div class="desc-titulo">O que a Central descobriu</div>{"".join(itens)}</section>')


def render_cartao_fila(c, ref, qtd_fila, maior_cod, tom, perfil):
    cliente = c["cliente"]
    cod = cliente["cod"]
    situacao = situacao_na_fila(c)
    qtd = c["qtd_vencidos"]

    tendencia_txt = (
        f"{texto_dias(c['tendencia'], 1, sinal=True)} (últimos {JANELA_TENDENCIA_MESES} meses vs. antes)"
        if c["atraso_medio_antigo"] is not None and c["atraso_medio_recente"] is not None
        else "sem dados suficientes para comparar"
    )
    atraso_hist_txt = texto_dias(c["atraso_medio_historico"], 1) if c["atraso_medio_historico"] is not None else "sem histórico"
    dias_sem_comprar_txt = f"há {texto_dias(c['dias_sem_comprar'])}" if c["dias_sem_comprar"] is not None else "—"
    base_hist = c["qtd_titulos_pagos_historico"]

    numeros = f"""
      <table class="tabela-numeros">
        <tr><td>Posição por valor vencido</td><td>{texto_ordinal(c['rank_valor'])} de {qtd_fila}</td></tr>
        <tr><td>Posição por dias em atraso</td><td>{texto_ordinal(c['rank_urgencia'])} de {qtd_fila}</td></tr>
        <tr><td>Posição por risco comportamental</td><td>{texto_ordinal(c['rank_risco'])} de {qtd_fila}</td></tr>
        <tr><td>Atraso médio histórico (títulos pagos)</td><td>{atraso_hist_txt} — com base em {base_hist} {plural(base_hist, 'título pago', 'títulos pagos')}</td></tr>
        <tr><td>Tendência recente do atraso</td><td>{tendencia_txt}</td></tr>
        <tr><td>Categoria / região</td><td>{html.escape(cliente['categoria'])} · {html.escape(cliente['regiao'])}</td></tr>
        <tr><td>Prazo de pagamento contratado</td><td>{html.escape(cliente['prazo_contratado'])}</td></tr>
        <tr><td>Última compra</td><td>{dias_sem_comprar_txt}</td></tr>
      </table>"""

    rotulo_titulos = "O título vencido" if qtd == 1 else f"Os {qtd} títulos vencidos"
    # se a cobranca for registrada agora, neste navegador, a etiqueta passa a dizer isso
    # (o filtro ja conta o cliente como aguardando retorno; a cor nao pode dizer outra coisa)
    etiqueta_registro = (
        "" if situacao == "aguardando" else
        render_situacao("aguardando").replace('class="situacao', 'class="situacao so-registrado', 1)
    )
    busca = sem_acento(f"{cliente['nome']} {cliente['contato']} {cliente['categoria']} {cliente['regiao']}")
    # embaixo da regua, o atraso mes a mes dos titulos ja pagos (sem historico, fica so a regua)
    linha_atraso = (gerar_linha_atraso_svg(perfil, SITUACOES_FILA[situacao][1])
                    if historico_para_desenhar(perfil) else "")

    return f"""
    <article class="linha{' recente' if c['contatado_recente'] else ''}" data-cod="{cod}" data-rank="{c['rank_medio']:.4f}"
         data-recente="{'1' if c['contatado_recente'] else '0'}" data-valor="{c['valor_vencido_centavos']}"
         data-dias="{c['dias_atraso_max']}" data-risco="{c['risco_comportamental']:.4f}" data-situacao="{situacao}"
         data-tom="{tom}" data-nome="{html.escape(cliente['nome'])}" data-busca="{html.escape(busca)}">
      <div class="linha-grade">
        <div class="posicao">{c['posicao']}</div>
        <div class="linha-corpo">
          <div class="linha-topo"><h3 class="nome">{html.escape(cliente['nome'])}</h3>{render_situacao(situacao)}{etiqueta_registro}</div>
          <div class="meta">{html.escape(cliente['categoria'])} · {html.escape(cliente['regiao'])}</div>
          <p class="explicacao">{prosa(gerar_explicacao(c, maior_cod))}</p>
          {render_bloco_contato(cliente, c.get("ultimo_contato"), c.get("dias_desde_contato"))}
        </div>
        <div class="linha-grafico">{gerar_regua_atraso_svg(c, situacao)}{linha_atraso}</div>
        <div class="valor">{formatar_reais_de_centavos(c['valor_vencido_centavos'])}<small>vencido · {qtd} {plural(qtd, 'título', 'títulos')}</small></div>
      </div>
      <div class="acoes">
        <button type="button" class="botao primario" id="ligou-{cod}" onclick="registrarLigacao({cod})">{icone("check")}<span>Registrei a ligação</span></button>
        <button type="button" class="botao fantasma" onclick="irParaRegua({cod})">{icone("regua")}Mensagem pronta</button>
        <button type="button" class="botao fantasma" onclick="abrirFicha({cod})">{icone("ficha")}Ficha</button>
        <button type="button" class="botao fantasma" aria-expanded="false" aria-controls="dobra-{cod}" onclick="alternarDobra({cod}, this)">{icone("lista")}{rotulo_titulos} e os números</button>
      </div>
      <div class="dobra" id="dobra-{cod}">
        <div><h4>{rotulo_titulos} · o que falar na ligação</h4>{render_tabela_titulos(c['titulos_vencidos'], ref)}</div>
        <div><h4>Os números por trás desta posição</h4>{numeros}</div>
      </div>
    </article>"""


def render_conteudo_hoje(ref, clientes, titulos, fila, perfis, descobertas, cods_regua):
    abertos = [t for t in titulos if t["dt_pagamento"] is None]
    vencidos = [t for t in abertos if t["dt_vencimento"] < ref]
    valor_aberto_centavos = sum(t["valor_centavos"] for t in abertos)
    valor_vencido_centavos = sum(t["valor_centavos"] for t in vencidos)
    maior_cod = max(fila, key=lambda c: c["valor_vencido_centavos"])["cliente"]["cod"] if fila else None
    fora_da_fila = len(clientes) - len(fila)

    cartoes = [render_cartao_fila(c, ref, len(fila), maior_cod, escolher_tom(perfis[c["cliente"]["cod"]])[0],
                                  perfis[c["cliente"]["cod"]])
               for c in fila]

    contagem = {}
    for c in fila:
        s = situacao_na_fila(c)
        contagem[s] = contagem.get(s, 0) + 1
    chips = [f'<button type="button" class="chip ativo" data-filtro="todos" onclick="filtrarHoje(\'todos\')">Todos <b>{len(fila)}</b></button>']
    for chave in ("fora", "piorou", "padrao", "sem_historico", "aguardando"):
        rotulo, cor = SITUACOES_FILA[chave]
        rotulo = rotulo.replace(" do contato", "")
        chips.append(
            f'<button type="button" class="chip cor-{cor}" data-filtro="{chave}" onclick="filtrarHoje(\'{chave}\')"'
            f'{"" if contagem.get(chave) else " hidden"}><i class="ponto"></i>{rotulo} <b>{contagem.get(chave, 0)}</b></button>'
        )

    metodo = render_metodo("Como a fila é montada", f"""
      <p>Entram só os clientes com título vencido em {formatar_data_curta(ref)}. A posição é a
      <strong>média de três posições</strong>, com pesos iguais: quanto ele tem vencido, há quantos dias venceu o
      título mais antigo, e o risco comportamental (quanto o atraso de hoje foge da média do próprio cliente,
      mais a piora recente).</p>
      <p>Quem foi cobrado nos últimos {JANELA_CONTATO_RECENTE_DIAS} dias <strong>desce para o fim da fila</strong>:
      ligar de novo agora só irritaria quem já foi procurado. A linha é uma por cliente, não por título: liga-se para
      uma pessoa, não para uma fatura.</p>
      <p>A etiqueta de cada cartão sai da mesma conta das outras abas: <strong>fora do padrão</strong> quando o atraso
      de hoje passa da variação normal dele, <strong>piorou</strong> quando o atraso dos últimos
      {JANELA_TENDENCIA_MESES} meses subiu de forma confirmada, <strong>padrão dele</strong> quando nada disso
      acontece.</p>""")

    return f"""
<section class="tela ativa" id="tela-hoje" role="tabpanel" aria-labelledby="aba-hoje">
  <header class="topo-tela">
    {render_migalha("hoje")}
    <div class="heroi-linha">
      <div class="heroi">
        <div class="heroi-rotulo">Vencido hoje</div>
        <div class="heroi-numero" data-total="vencido">{formatar_reais_de_centavos(valor_vencido_centavos)}</div>
        <div class="heroi-sub"><strong>{len(vencidos)} títulos</strong> de <strong>{len(fila)} clientes</strong> · vencido hoje</div>
      </div>
      <div class="heroi-lado">
        <div class="stats-coluna">
          <div class="stat"><div class="numero" data-total="a vencer">{formatar_reais_de_centavos(valor_aberto_centavos - valor_vencido_centavos)}</div>
            <div class="rotulo">a vencer · {len(abertos) - len(vencidos)} títulos ainda no prazo</div></div>
          <div class="stat"><div class="numero" data-total="em aberto">{formatar_reais_de_centavos(valor_aberto_centavos)}</div>
            <div class="rotulo">saldo total em aberto · {len(abertos)} títulos</div></div>
          <div class="stat"><div class="numero" data-total="clientes">{len(clientes)}</div>
            <div class="rotulo">clientes ativos · {len(titulos)} títulos nos últimos {meses_de_historico(titulos, ref)} meses</div></div>
        </div>
      </div>
    </div>
    {render_descobertas(descobertas, perfis, cods_regua)}
  </header>

  <div class="secao-cabeca">
    <h2>A fila de hoje</h2>
    <p>{len(fila)} {plural(len(fila), 'cliente', 'clientes')} com título vencido. Os outros {fora_da_fila} não têm nada
    vencido hoje e estão em <a class="link" href="#tela-risco" onclick="return irPara(event, 'risco')">Risco por cliente</a>,
    que olha a carteira inteira.</p>
  </div>
  <div class="ferramentas" role="toolbar" aria-label="Filtrar e ordenar a fila">
    <div class="chips" id="chips-hoje">{"".join(chips)}</div>
    <label class="ordem">Ordenar por
      <select id="ordem-hoje" onchange="ordenarHoje(this.value)">
        <option value="rank">prioridade da Central</option>
        <option value="valor">maior valor vencido</option>
        <option value="dias">título mais antigo</option>
        <option value="risco">risco comportamental</option>
      </select>
    </label>
  </div>
  <div class="legenda-regua"><span><i class="lg-faixa"></i>onde o cliente costuma pagar</span>
    <span><i class="lg-ponto"></i>o título vencido mais antigo dele hoje</span></div>
  <div class="fila">
    {''.join(cartoes)}
  </div>
  <p class="lista-vazia" id="vazio-hoje" hidden>Nenhum cliente da fila com esse filtro ou essa busca.</p>
  {metodo}
</section>
"""


# ---------------------------------------------------------------------------
# Tela Risco por cliente
# ---------------------------------------------------------------------------

def render_conteudo_risco(ref, perfis, meses_historico):
    ordenados = sorted(perfis.values(), key=lambda p: -p["valor_em_risco_centavos"])
    total_em_risco = sum(p["valor_em_risco_centavos"] for p in perfis.values())
    total_saldo = sum(p["saldo_aberto_centavos"] for p in perfis.values())
    proporcao = (total_em_risco / total_saldo) if total_saldo else 0
    maior_saldo = max((p["saldo_aberto_centavos"] for p in perfis.values()), default=0) or 1
    maior_risco = max((p["valor_em_risco_centavos"] for p in perfis.values()), default=0) or 1

    linhas_html = []
    tabela_html = []
    fichas_html = []
    faixas_conta = {}
    com_selo = 0
    for pos, p in enumerate(ordenados, start=1):
        cliente = p["cliente"]
        cod = cliente["cod"]
        tier, rotulo_tier = faixa_de_risco(p["indice"])
        faixas_conta[tier] = faixas_conta.get(tier, 0) + 1
        cor = COR_DA_FAIXA[tier]

        selos = ""
        if p["selo_cronico_estavel"]:
            selos += f'<span class="selo cor-ok">{icone("certo")}padrão dele, não é risco</span>'
        if p["selo_invisivel"]:
            selos += f'<span class="selo cor-atencao">{icone("olho")}nunca apareceu em lista nenhuma</span>'
        if selos:
            com_selo += 1

        teaser = partes_do_alerta(p)[0]
        situacao = (
            f"{p['qtd_vencidos_hoje']} {plural(p['qtd_vencidos_hoje'], 'título vencido', 'títulos vencidos')}"
            if p["qtd_vencidos_hoje"] else "nada vencido hoje"
        )
        largura_saldo = p["saldo_aberto_centavos"] / maior_saldo * 100
        busca = sem_acento(f"{cliente['nome']} {cliente['contato']} {cliente['categoria']} {cliente['regiao']}")

        linhas_html.append(f"""
        <button type="button" class="risco-linha tier-{tier}" onclick="abrirFicha({cod})" data-cod="{cod}"
                data-faixa="{tier}" data-selo="{'1' if selos else '0'}" data-em-risco="{p['valor_em_risco_centavos']}"
                data-indice="{p['indice']}" data-saldo="{p['saldo_aberto_centavos']}" data-busca="{html.escape(busca)}"
                aria-label="Abrir ficha de {html.escape(cliente['nome'])}">
          <span class="rl-pos">{pos}</span>
          <span class="rl-corpo">
            <span class="rl-nome">{html.escape(cliente['nome'])}{selos}</span>
            <span class="meta">{html.escape(cliente['categoria'])} · {html.escape(cliente['regiao'])} · {situacao}</span>
          </span>
          <span class="rl-barra">
            <span class="rl-trilho" style="width:{largura_saldo:.1f}%"><span class="rl-cheio" style="width:{p['indice']:.0f}%"></span></span>
            <span class="rl-barra-rot"><span class="dinheiro">{formatar_reais_estimado(p['valor_em_risco_centavos'])}</span> em risco de {formatar_reais_de_centavos(p['saldo_aberto_centavos'])} em aberto</span>
          </span>
          <span class="rl-numeros">
            <span class="rl-valor">{formatar_reais_estimado(p['valor_em_risco_centavos'])} <small>em risco</small></span>
            <span class="indice-chip cor-{cor}">{p['indice']:.0f}/100 · {rotulo_tier}</span>
          </span>
          <span class="rl-teaser">{prosa(teaser)}</span>
          <span class="rl-abrir">Abrir ficha {icone("seta")}</span>
        </button>""")
        # a mesma carteira em tabela, uma linha por cliente (fica embaixo dos cartoes)
        tabela_html.append(f"""
          <tr class="tr-risco tier-{tier}" onclick="abrirFicha({cod})" data-cod="{cod}">
            <td class="pos">{pos}</td>
            <td><div class="cli"><button type="button" class="link" onclick="event.stopPropagation(); abrirFicha({cod})"
                aria-label="Abrir ficha de {html.escape(cliente['nome'])}">{html.escape(cliente['nome'])}</button>{selos}</div></td>
            <td class="sit{'' if p['qtd_vencidos_hoje'] else ' nada'}">{situacao}</td>
            <td class="dir">{formatar_reais_de_centavos(p['saldo_aberto_centavos'])}</td>
            <td><span class="indice cor-{cor}"><b>{p['indice']:.0f}</b><span class="indice-barra"><span style="width:{p['indice']:.0f}%"></span></span><span class="indice-chip cor-{cor}">{rotulo_tier[6:]}</span></span></td>
            <td class="dir"><span class="em-risco"><span class="barra-dinheiro"><span style="width:{p['valor_em_risco_centavos'] / maior_risco * 100:.0f}%"></span></span><b>{formatar_reais_estimado(p['valor_em_risco_centavos'])}</b></span></td>
            <td class="seta">›</td>
          </tr>""")
        fichas_html.append(render_ficha(ref, p))

    chips = [f'<button type="button" class="chip ativo" data-filtro="todos" onclick="filtrarRisco(\'todos\')">Todos <b>{len(perfis)}</b></button>']
    for minimo, chave, rotulo in FAIXAS_RISCO:
        chips.append(
            f'<button type="button" class="chip cor-{COR_DA_FAIXA[chave]}" data-filtro="{chave}" onclick="filtrarRisco(\'{chave}\')"'
            f'{"" if faixas_conta.get(chave) else " hidden"}><i class="ponto"></i>{rotulo[6:].capitalize()} <b>{faixas_conta.get(chave, 0)}</b></button>'
        )
    chips.append(
        f'<button type="button" class="chip" data-filtro="selo" onclick="filtrarRisco(\'selo\')"'
        f'{"" if com_selo else " hidden"}>Com selo da Central <b>{com_selo}</b></button>'
    )

    metodo = render_metodo("Como o índice e o valor em risco são calculados", f"""
      <p>O <strong>valor em risco</strong> de cada cliente é o saldo em aberto multiplicado pelo índice de risco dele
      (0 a 100), calculado sobre {meses_historico} meses de histórico de pagamento, de compra e de uso do limite.
      A lista segue o valor em risco porque a pergunta é onde está o dinheiro, não quem é o pior cliente.</p>
      <p>O índice soma seis pesos que fecham 100: nível de atraso ({PESO_NIVEL}), variabilidade
      ({PESO_VARIABILIDADE}), mudança recente de comportamento ({PESO_MUDANCA}), sinal de compra ({PESO_COMPRA}),
      concentração no faturamento ({PESO_FATURAMENTO}) e ocupação do limite ({PESO_LIMITE}). A ficha de cada
      cliente mostra os seis, com a frase que justifica os pontos.</p>
      <p>As faixas são os quartos da escala: 0–24 baixo, 25–49 moderado, 50–74 alto, 75–100 crítico.</p>""")

    lista = f"""
<section class="tela" id="tela-risco" role="tabpanel" aria-labelledby="aba-risco">
  <header class="topo-tela">
    {render_migalha("risco")}
    <div class="heroi-linha">
      <div class="heroi">
        <div class="heroi-rotulo">Em risco na carteira</div>
        <div class="heroi-numero">{formatar_reais_estimado(total_em_risco)}</div>
        <div class="heroi-sub">de <strong>{formatar_reais_de_centavos(total_saldo)}</strong> em aberto ·
          <strong>{texto_pct(proporcao, 0)}</strong> da carteira</div>
      </div>
      <div class="heroi-lado">
        <div class="medidor-carteira" role="img" aria-label="{texto_pct(proporcao, 0)} da carteira em aberto está em risco">
          <div class="medidor-trilho"><div class="medidor-cheio" style="width:{proporcao * 100:.1f}%"></div></div>
          <div class="medidor-rotulos"><span><span class="dinheiro">{formatar_reais_estimado(total_em_risco)}</span> em risco</span>
            <span>{formatar_reais_de_centavos(total_saldo)} em aberto</span></div>
        </div>
        <p class="heroi-texto">Saldo em aberto de cada cliente × o índice de risco dele, de 0 a 100. A ordem é
          <strong>onde está o dinheiro</strong>, não quem é o pior cliente. Clique em qualquer cliente para abrir a ficha.</p>
        <div class="legenda-faixas">
          <span><i class="ponto" style="background:var(--ok)"></i>0–24 risco baixo</span>
          <span><i class="ponto" style="background:var(--atencao)"></i>25–49 moderado</span>
          <span><i class="ponto" style="background:var(--grave)"></i>50–74 alto</span>
          <span><i class="ponto" style="background:var(--grave)"></i>75–100 crítico</span>
        </div>
      </div>
    </div>
  </header>
  <div class="painel painel-dispersao">
    <h3>Quem deve muito e quem está piorando</h3>
    <p class="sub">Cada ponto é um cliente: mais à direita, mais saldo em aberto; mais acima, mais atraso nos últimos
      {JANELA_TENDENCIA_MESES} meses do que antes.</p>
    {gerar_dispersao_risco_svg(perfis)}
    <div class="grafico-legenda"><span class="legenda-item"><i class="lg ponto-piorou"></i>piorou de forma confirmada</span>
      <span class="legenda-item"><i class="lg ponto-neutro"></i>sem piora confirmada</span>
      <span class="legenda-item">passe o mouse num ponto para ver o cliente</span></div>
  </div>
  <div class="ferramentas" role="toolbar" aria-label="Filtrar e ordenar os clientes">
    <div class="chips" id="chips-risco">{"".join(chips)}</div>
    <label class="ordem">Ordenar por
      <select id="ordem-risco" onchange="ordenarRisco(this.value)">
        <option value="em-risco">valor em risco</option>
        <option value="indice">índice de risco</option>
        <option value="saldo">saldo em aberto</option>
      </select>
    </label>
  </div>
  <div class="legenda-regua"><span><i class="lg-faixa"></i>saldo em aberto do cliente</span>
    <span><i class="lg-faixa dinheiro-faixa"></i>a parte dele que está em risco</span></div>
  <div class="lista-risco" id="lista-risco">
    {''.join(linhas_html)}
  </div>
  <p class="lista-vazia" id="vazio-risco" hidden>Nenhum cliente com esse filtro ou essa busca.</p>
  <div class="secao-cabeca secao-tabela-risco">
    <h2>Os {len(perfis)} clientes em tabela</h2>
    <p>Sempre a carteira inteira, uma linha por cliente, na ordem do valor em risco. Clique na linha para abrir a ficha.</p>
  </div>
  <div class="painel painel-tabela">
    <div class="rolagem-tabela"><table class="tabela-risco">
      <thead><tr><th>#</th><th>Cliente</th><th>Situação</th><th class="dir">Saldo em aberto</th><th>Índice de risco</th>
        <th class="dir">Valor em risco</th><th></th></tr></thead>
      <tbody>{''.join(tabela_html)}</tbody>
      <tfoot><tr><td></td><td>Carteira inteira</td><td></td>
        <td class="dir">{formatar_reais_de_centavos(total_saldo)}</td><td></td>
        <td class="dir"><span class="dinheiro">{formatar_reais_estimado(total_em_risco)}</span></td><td></td></tr></tfoot>
    </table></div>
    <p class="nota-tabela">valor em risco = saldo em aberto × índice ÷ 100</p>
  </div>
  {metodo}
</section>
"""
    return lista, "".join(fichas_html)


def render_ficha(ref, p):
    cliente = p["cliente"]
    cod = cliente["cod"]
    alerta = gerar_alerta_ficha(p)

    legenda_grafico = """
    <div class="grafico-legenda">
      <span class="legenda-item"><i class="lg ponto-normal"></i>dentro do padrão dele</span>
      <span class="legenda-item"><i class="lg ponto-fora"></i>fora do padrão (mudança confirmada)</span>
      <span class="legenda-item"><i class="lg ponto-sazonal"></i>mês sazonal (descontado da comparação)</span>
      <span class="legenda-item"><i class="lg faixa"></i>faixa do padrão histórico (± 1 desvio-padrão)</span>
      <span class="legenda-item">passe o mouse em cada ponto para ver o mês</span>
    </div>"""

    sazonal_txt = ""
    if p["sazonais"]:
        nomes_meses = ", ".join(MESES_PT_INV[m].capitalize() for m in sorted(p["sazonais"]))
        detalhes_anos = []
        for m, info in sorted(p["sazonais"].items()):
            anos_txt = ", ".join(f"{a}: {v:.0f}d" for a, v in sorted(info["anos"].items()))
            detalhes_anos.append(f"{MESES_PT_INV[m].capitalize()} ({anos_txt}, vs. média de {info['media_resto']:.0f}d no resto do ano)")
        sazonal_txt = (
            f'<div class="fato largo"><div class="rotulo-fato">Padrão de calendário identificado</div>'
            f'<div class="valor-fato">{html.escape(nomes_meses)} — repete em anos diferentes, '
            f'descontado da comparação de tendência. {html.escape("; ".join(detalhes_anos))}.</div></div>'
        )

    def ou_insuficiente(valor, texto="histórico insuficiente"):
        return valor if valor is not None else f'<span class="sem-dado">{texto}</span>'

    fatos = [
        ("Prazo contratado", html.escape(cliente["prazo_contratado"])),
        ("Prazo real (emissão → pagamento)", ou_insuficiente(texto_dias(p["prazo_efetivo"]) if p.get("prazo_efetivo") is not None else None)),
        ("Atraso médio depois do vencimento", ou_insuficiente(texto_dias(p["nivel"]) if p["nivel"] is not None else None)),
        ("Variação do atraso (mín a máx)", ou_insuficiente(f"{p['atraso_min']} a {texto_dias(p['atraso_max'])}" if p["atraso_min"] is not None else None)),
        ("Desvio-padrão do atraso", ou_insuficiente(f"± {texto_dias(p['variabilidade'])}" if p["variabilidade"] is not None else None)),
        ("Saldo em aberto", formatar_reais_de_centavos(p["saldo_aberto_centavos"])),
        ("Limite de crédito", formatar_reais_de_centavos(cliente["limite_credito_centavos"])),
        ("Ocupação do limite", ou_insuficiente(texto_pct(p['ocupacao_limite'], 0) if p["ocupacao_limite"] is not None else None, "sem limite cadastrado")),
        ("Faturamento 12 meses", formatar_reais_de_centavos(p["faturamento_12m"])),
        ("Fatia do faturamento", texto_pct(p['participacao_faturamento'])),
        ("Último pedido", ou_insuficiente(f"há {texto_dias(p['dias_sem_comprar'])}" if p["dias_sem_comprar"] is not None else None, "sem compras")),
        ("Intervalo entre pedidos", ou_insuficiente(f"{texto_dias(p['intervalo_medio'])} em média" if p["intervalo_medio"] else None)),
        ("Cliente desde", formatar_data_curta(cliente["cliente_desde"])),
    ]
    fatos_html = "".join(
        f'<div class="fato"><div class="rotulo-fato">{html.escape(rot)}</div><div class="valor-fato">{val}</div></div>'
        for rot, val in fatos
    )

    componentes_html = "".join(f"""
        <details class="componente">
          <summary><span class="comp-nome">{html.escape(c['nome'])}</span>
            <span class="barra-peso"><span style="width:{(c['pontos'] / c['peso'] * 100) if c['peso'] else 0:.0f}%"></span></span>
            <span class="pontos{' zerado' if c['pontos'] == 0 else ''}">{texto_decimal(c['pontos'])} de {c['peso']} pts</span></summary>
          <div class="frase">{prosa(c['frase'])}</div>
        </details>""" for c in p["componentes"])

    tier, rotulo_tier = faixa_de_risco(p["indice"])
    if p["selo_cronico_estavel"]:
        cor_alerta = "ok"
    elif tier in ("alto", "critico"):
        cor_alerta = "grave"
    elif p["selo_invisivel"]:
        cor_alerta = "atencao"
    else:
        cor_alerta = "neutro"

    qtd_abertos = len(p["titulos_abertos"])
    titulo_secao_titulos = (
        f"{qtd_abertos} {plural(qtd_abertos, 'título em aberto', 'títulos em aberto')}"
        if qtd_abertos else "Títulos em aberto"
    )
    dias_contato = (ref - p["ultimo_contato"]["dt_hora_contato"].date()).days if p.get("ultimo_contato") else None
    acoes = []
    if p["tem_vencido_hoje"]:
        acoes.append(f'<button type="button" class="botao" onclick="irParaFila({cod})">{icone("hoje")}Ver na fila de Hoje</button>')
    if p["tem_vencido_hoje"] or p["selo_invisivel"]:
        acoes.append(f'<button type="button" class="botao" onclick="irParaRegua({cod})">{icone("regua")}Mensagem pronta</button>')

    return f"""
    <div class="ficha ficha-modal" id="ficha-{cod}" style="display:none;" role="dialog" aria-modal="true" aria-label="Ficha de {html.escape(cliente['nome'])}">
      <div class="ficha-cabeca">
        <div>
          <div class="migalha"><span class="passo">Ficha do cliente</span></div>
          <h3>{html.escape(cliente['nome'])}</h3>
          <div class="meta-ficha">{html.escape(cliente['categoria'])} · {html.escape(cliente['regiao'])}/{html.escape(cliente['uf'])} · cliente desde {formatar_data_curta(cliente['cliente_desde'])}</div>
        </div>
        <button class="fechar" onclick="fecharFicha()" aria-label="Fechar ficha">{icone("fechar")}</button>
      </div>
      <div class="ficha-topo">
        <div>
          <div class="indice-grande">{p['indice']:.0f}<span>/100</span></div>
          <div class="indice-legenda"><span class="indice-chip cor-{COR_DA_FAIXA[tier]}">{rotulo_tier}</span></div>
        </div>
        <div class="blocos-topo">
          <div class="bloco-topo">
            <div class="rotulo-topo">Valor em risco</div>
            <div class="valor-topo">{formatar_reais_estimado(p['valor_em_risco_centavos'])}</div>
          </div>
          <div class="bloco-topo">
            <div class="rotulo-topo">Saldo em aberto</div>
            <div class="valor-topo">{formatar_reais_de_centavos(p['saldo_aberto_centavos'])}</div>
          </div>
          <div class="bloco-topo">
            <div class="rotulo-topo">Vencido hoje</div>
            <div class="valor-topo">{p['qtd_vencidos_hoje']} {plural(p['qtd_vencidos_hoje'], 'título', 'títulos')}</div>
          </div>
        </div>
      </div>
      {render_bloco_contato(cliente, p.get("ultimo_contato"), dias_contato)}
      {f'<div class="ficha-acoes">{"".join(acoes)}</div>' if acoes else ''}
      <div class="alerta cor-{cor_alerta}">{prosa(alerta)}</div>

      <h4>Em quantos dias ele pagou, mês a mês</h4>
      <div class="grafico-caixa"><div class="rolagem-grafico">{gerar_grafico_svg(p)}</div>{legenda_grafico}</div>

      <h4>{html.escape(titulo_secao_titulos)}</h4>
      {render_tabela_titulos(p["titulos_abertos"], ref)}

      <h4>Os números do comportamento dele</h4>
      <div class="grade-fatos">{fatos_html}{sazonal_txt}</div>

      <h4>Como o índice de risco foi montado (0 a 100)</h4>
      <div class="componentes">{componentes_html}</div>
    </div>"""


# ---------------------------------------------------------------------------
# Tela Previsao de caixa
# ---------------------------------------------------------------------------

def render_aviso_sem_historico(prev):
    if not prev["sem_historico"]:
        return ""
    qtd = len(prev["sem_historico"])
    nomes = sorted({x["cliente"]["nome"] for x in prev["sem_historico"]})
    return (
        f'<div class="aviso-escopo"><strong>{qtd} {plural(qtd, "título ficou", "títulos ficaram")} fora da '
        f'previsão: {formatar_reais_de_centavos(prev["sem_historico_centavos"])}.</strong> '
        f'{plural(len(nomes), "É de cliente", "São de clientes")} sem histórico suficiente para saber quando '
        f'paga ({html.escape(", ".join(nomes))}: menos de {MIN_TITULOS_HISTORICO} títulos pagos). Em vez de '
        f'estimar sem base, a Central tira esses títulos das duas contas — da planilha e da esperada — para a '
        f'comparação continuar sobre a mesma carteira.</div>'
    )


def render_conteudo_previsao(ref, prev, semana):
    total_planilha = prev["total_planilha"]
    total_esperado = prev["total_esperado"]
    dif = prev["diferenca"]
    pct = (dif / total_planilha) if total_planilha else 0
    fim = prev["fim_janela"]
    faltou = dif > 0

    hero_frase = (
        f"Some os vencimentos das próximas 4 semanas e dá {formatar_reais_estimado(total_planilha)}. Lendo como "
        f"cada cliente realmente paga, esperamos {formatar_reais_estimado(total_esperado)}: "
        f"{texto_pct(abs(pct), 0)} a menos."
    ) if faltou else (
        "É quanto deve entrar além do que a planilha aponta, porque parte do que já venceu entra nestas 4 semanas."
    )

    semanas_html = []
    for i, s in enumerate(prev["semanas"]):
        d = prev["planilha"][i] - prev["esperado"][i]
        semanas_html.append(f"""
        <tr>
          <td><span class="semana-rotulo">Semana {i + 1}</span>
              <span class="semana-datas"> · {s['ini']:%d/%m} a {s['fim']:%d/%m}</span></td>
          <td class="dir">{formatar_reais_estimado(prev['planilha'][i])}</td>
          <td class="dir">{formatar_reais_estimado(prev['esperado'][i])}</td>
          <td class="dir {'menos' if d > 0 else ''}">{'−' if d > 0 else '+'} {formatar_reais_estimado(abs(d))}</td>
        </tr>""")

    linhas_titulos = []
    for l in prev["linhas"]:
        semana_txt = f"semana {l['semana'] + 1}" if l["semana"] is not None else "depois das 4 semanas"
        improvavel = l["chance"] < 0.5
        marca_baixa = '<span class="baixa">baixa</span>' if improvavel else ""
        linhas_titulos.append(
            f'<tr class="{"improvavel" if improvavel else ""}">'
            f'<td class="num">{html.escape(l["titulo"]["numero"])}</td>'
            f'<td>{html.escape(l["cliente"]["nome"])}</td>'
            f'<td class="dir">{formatar_reais_de_centavos(l["titulo"]["valor_centavos"])}</td>'
            f'<td>{formatar_data_curta(l["titulo"]["dt_vencimento"])}</td>'
            f'<td>{formatar_data_curta(l["data_esperada"])} <span class="semana-datas">({semana_txt})</span></td>'
            f'<td class="dir chance">{texto_pct(l["chance"], 0)}{marca_baixa}</td>'
            f'<td class="dir">{formatar_reais_estimado(l["valor_ponderado"])}</td>'
            f"</tr>"
        )

    nota_inversao = ""
    if prev["semanas_invertidas"]:
        quais = " e ".join(str(s) for s in prev["semanas_invertidas"])
        nota_inversao = render_metodo("Por que a faixa de cenários só aparece no acumulado", f"""
          <p>Atrasar não faz o dinheiro sumir, só empurra ele para a semana seguinte. Então, olhando semana isolada,
          o cenário conservador chega a ficar <em>acima</em> do otimista — nesta base isso acontece
          {'nas semanas' if len(prev['semanas_invertidas']) > 1 else 'na semana'} {quais}, porque o dinheiro que o
          otimista já contou antes reaparece ali no conservador. A faixa só é honesta no acumulado, onde o otimista
          nunca fica abaixo do conservador: quem paga mais cedo sempre chegou antes.</p>""")

    atraso_semana = (f" (atraso médio de {texto_dias(semana['atraso_medio_recebido'])})"
                     if semana["atraso_medio_recebido"] is not None else "")

    metodo = render_metodo("Como a previsão é feita", f"""
      <p>Cada título em aberto recebe a data em que <strong>aquele cliente costuma pagar de verdade</strong>: o
      vencimento mais o atraso que ele pratica. Quem mudou de comportamento de forma confirmada, para pior ou para
      melhor, é projetado pelo ritmo dos últimos {JANELA_TENDENCIA_MESES} meses. Título que vence em mês sazonal
      usa o atraso daquele mês.</p>
      <p>"O que a planilha apontaria" é a soma dos vencimentos das 4 semanas, com o que já venceu lançado na
      semana 1, para os dois números cobrirem a mesma carteira. O conservador e o otimista usam o dia ruim e o dia
      bom de cada cliente (atraso esperado ± 1 desvio-padrão).</p>""")

    return f"""
<section class="tela" id="tela-previsao" role="tabpanel" aria-labelledby="aba-previsao">
  <header class="topo-tela">
    {render_migalha("previsao")}
    <div class="heroi-grade">
      <div class="heroi">
        <div class="heroi-rotulo">O que a planilha promete e não entra</div>
        <div class="heroi-numero">{formatar_reais_estimado(abs(dif))}</div>
        <div class="heroi-sub">de {formatar_data_curta(ref)} a {formatar_data_curta(fim)}</div>
        <p class="heroi-texto">{prosa(hero_frase)}</p>
        <div class="heroi-stats">
          <div class="stat"><div class="numero">{formatar_reais_estimado(total_planilha)}</div>
            <div class="rotulo">a planilha apontaria · soma dos vencimentos, com o que já venceu lançado na semana 1</div></div>
          <div class="stat"><div class="numero">{formatar_reais_estimado(total_esperado)}</div>
            <div class="rotulo">esperamos receber · lendo o comportamento de cada cliente</div></div>
          <div class="stat menos"><div class="numero">{'−' if faltou else '+'} {formatar_reais_estimado(abs(dif))}</div>
            <div class="rotulo">diferença · {texto_pct(abs(pct), 0)} do que a planilha promete</div></div>
        </div>
      </div>
      <div class="painel-heroi vidro">
        <div class="rolagem-grafico">{gerar_grafico_previsao_svg(prev)}</div>
        <div class="grafico-legenda">
          <span class="legenda-item"><i class="lg" style="width:18px;border-top:2px dashed rgba(236,238,255,.7)"></i>o que a planilha promete</span>
          <span class="legenda-item"><i class="lg" style="width:18px;height:3px;background:var(--dinheiro)"></i>o que esperamos receber</span>
          <span class="legenda-item"><i class="lg" style="width:18px;height:10px;background:rgba(124,198,255,.25)"></i>faixa entre conservador e otimista</span>
        </div>
        {nota_inversao}
      </div>
    </div>
  </header>

  <div class="avisos">
    <div class="aviso-escopo"><strong>Isto não é projeção comercial.</strong> Só a carteira que já existe — os
      {len(prev['linhas']) + len(prev['sem_historico'])} títulos hoje em aberto. Venda nova <strong>não entra nesta conta</strong>.
      Estimativa aparece sem centavos; valor de título continua com centavo.</div>
    <div class="aviso-escopo"><strong>A chance de entrada é premissa, não medida.</strong> A base não tem nenhum
      título baixado como perda, então não há inadimplência observada para calibrar. Como ela é calculada está em
      "Título a título", abaixo.</div>
    {render_aviso_sem_historico(prev)}
  </div>

  <div class="duas-colunas">
    <div class="painel">
      <h3>Semana a semana</h3>
      <p class="sub">Sem a faixa de cenários aqui, de propósito: o motivo está no gráfico acima.</p>
      <div class="rolagem-tabela"><table class="tabela-semanas">
        <thead><tr><th>Período</th><th class="dir">Planilha</th><th class="dir">Esperado</th><th class="dir">Diferença</th></tr></thead>
        <tbody>{''.join(semanas_html)}</tbody>
        <tfoot><tr>
          <td>Total nas 4 semanas</td>
          <td class="dir">{formatar_reais_estimado(total_planilha)}</td>
          <td class="dir">{formatar_reais_estimado(total_esperado)}</td>
          <td class="dir menos">{'−' if faltou else '+'} {formatar_reais_estimado(abs(dif))}</td>
        </tr></tfoot>
      </table></div>
      <p class="nota">{prosa(f"Da carteira em aberto, {formatar_reais_estimado(prev['fora_da_janela'])} devem entrar só depois destas 4 semanas, e {formatar_reais_estimado(prev['descontado'])} foram descontados pela chance de entrada — títulos velhos demais para o padrão do próprio cliente, ou de cliente com índice de risco alto.")}</p>
    </div>
    <div class="painel">
      <h3>A semana que acabou de passar</h3>
      <p class="sub">De {formatar_data_curta(semana['inicio'])} a {formatar_data_curta(semana['fim'])}.</p>
      <div class="faixa-semana">
        <div class="stat"><div class="numero">{formatar_reais_de_centavos(semana['recebido_centavos'])}</div>
          <div class="rotulo">entraram · {semana['recebido_qtd']} {plural(semana['recebido_qtd'], 'título', 'títulos')}{atraso_semana}</div></div>
        <div class="stat"><div class="numero">{formatar_reais_de_centavos(semana['venceu_centavos'])}</div>
          <div class="rotulo">venceram na semana</div></div>
        <div class="stat menos"><div class="numero">{formatar_reais_de_centavos(semana['nao_entrou_centavos'])}</div>
          <div class="rotulo">não foram pagos · {semana['nao_entrou_qtd']} {plural(semana['nao_entrou_qtd'], 'título', 'títulos')}</div></div>
      </div>
    </div>
  </div>

  <details class="metodo">
    <summary>{icone("lista")}Título a título: os {len(prev['linhas'])} títulos projetados e a data esperada de cada um</summary>
    <div class="metodo-corpo">
      <p>A chance de entrada de cada título cai por dois motivos: o índice de risco do cliente (até
        {texto_pct(PESO_RISCO_NA_ENTRADA, 0)} de desconto) e a idade do título medida contra o maior atraso que
        <em>aquele</em> cliente já praticou. O piso é {texto_pct(PISO_CHANCE_ENTRADA, 0)}: nenhum título é dado como
        perdido. <strong>Esses percentuais são premissa, não medida:</strong> a base não tem nenhum título baixado como
        perda, então não há inadimplência observada para calibrar.</p>
      <div class="rolagem-tabela"><table class="tabela-titulos ampla">
        <thead><tr>
          <th>Título</th><th>Cliente</th><th class="dir">Valor</th><th>Vencimento</th>
          <th>Entrada esperada</th><th class="dir">Chance</th><th class="dir">Ponderado</th>
        </tr></thead>
        <tbody>{''.join(linhas_titulos)}</tbody>
      </table></div>
    </div>
  </details>
  {metodo}
</section>
"""


# ---------------------------------------------------------------------------
# Tela Regua de cobranca
# ---------------------------------------------------------------------------

def clientes_da_regua(perfis):
    """Quem entra na regua: quem tem titulo vencido, mais quem a tela de Risco
    mandou procurar mesmo sem nada vencido."""
    escolhidos = [p for p in perfis.values() if p["tem_vencido_hoje"] or p["selo_invisivel"]]
    return sorted(escolhidos, key=lambda p: -p["valor_em_risco_centavos"])


def situacao_arquivo_registros():
    """O que o terminal diz sobre o arquivo de registros quando o dono roda a Central: onde
    ele tem que ficar, se foi encontrado e quantas cobrancas foram lidas, e se ha um parecido
    com o nome errado na pasta. Isso saiu da tela a pedido do dono: a pagina pode ser aberta
    em outro computador, onde o caminho desta pasta nao faz sentido. Sem acento, como o resto
    do que o central.py escreve no terminal."""
    nome = os.path.basename(ARQUIVO_REGISTROS)
    linhas = []
    if os.path.exists(ARQUIVO_REGISTROS):
        lidos = carregar_registros()
        if lidos:
            total = sum(r["valor_cobrado_centavos"] for r in lidos)
            linhas.append(f"Registros de cobranca: lendo {len(lidos)} {plural(len(lidos), 'cobranca', 'cobrancas')} "
                          f"de {nome} ({formatar_reais_de_centavos(total)}). Ja entraram na fila de Hoje e no relatorio.")
        else:
            linhas.append(f"ATENCAO: existe {nome} nesta pasta, mas nenhuma linha foi lida. "
                          f"Confira se o cabecalho continua o mesmo.")
    else:
        linhas.append(f"Registros de cobranca: ainda nao existe {nome} nesta pasta. O que for registrado na tela "
                      f"fica so no navegador ate o arquivo baixado (botao 'Baixar registros') ser salvo aqui:")
        linhas.append(f"  {ARQUIVO_REGISTROS}")
    perdidos = registros_fora_do_lugar()
    if perdidos:
        linhas.append(f"ATENCAO: arquivo com nome errado nesta pasta: {', '.join(perdidos)}. A Central so le {nome}; "
                      f'se for o registro que voce baixou, renomeie (o navegador costuma acrescentar "(1)").')
    return linhas


def render_estado_copia(ref):
    """Na copia enviada o arquivo de registros nao existe e o download fica
    desligado. A tela diz isso em vez de fingir que funciona."""
    return (
        f'<div class="achado-neutro"><strong>Esta é uma cópia enviada, com os dados de {formatar_data_curta(ref)}.</strong> '
        f'Aqui o registro fica só neste navegador: ele organiza a sua fila da aba Hoje, mas não entra na conta da '
        f'Central. O botão de baixar o arquivo de registros está desligado nesta cópia, porque o registro que conta '
        f'é o da Central de quem te mandou este arquivo, e um arquivo vindo de outra máquina, salvo por cima do '
        f'dela, apagaria cobranças. <strong>Cobrou alguém? Avise quem te mandou a Central.</strong></div>'
    )


def render_conteudo_regua(ref, perfis, para_enviar=False):
    alvos = clientes_da_regua(perfis)

    textos_por_cliente = {}
    itens = []
    blocos = []
    for i, p in enumerate(alvos):
        cod = p["cliente"]["cod"]
        tom_sugerido, motivo = escolher_tom(p)
        ctx = montar_contexto_redacao(ref, p)

        # todos os tons ja vem escritos, para a troca na mao ser instantanea sem
        # que o navegador precise redigir nada -- a redacao mora so em redacao.py
        possiveis = redacao.tons_possiveis(ctx)
        textos_por_cliente[cod] = {tom: redacao.redigir(tom, ctx) for tom in possiveis}

        vencidos = [t for t in p["titulos_abertos"] if t["dt_vencimento"] < ref]
        valor_cobravel = sum(t["valor_centavos"] for t in (vencidos or p["titulos_abertos"]))
        cor = COR_DO_TOM[redacao.TONS[tom_sugerido]["cor"]]
        nome_tom = redacao.TONS[tom_sugerido]["nome"]

        opcoes = "".join(
            f'<option value="{t}"{" selected" if t == tom_sugerido else ""}>{html.escape(redacao.TONS[t]["nome"])}</option>'
            for t in possiveis
        )
        fora = [redacao.TONS[t]["nome"].lower() for t in redacao.TONS if t not in possiveis]
        aviso_tons = (
            f'<span class="aviso-tons">Sem {(", ".join(fora[:-1]) + " e " + fora[-1]) if len(fora) > 1 else fora[0]}: '
            f'{"esses tons citam" if len(fora) > 1 else "esse tom cita"} números que este '
            f'cliente ainda não tem.</span>'
        ) if fora else ""
        situacao = (
            f"{p['qtd_vencidos_hoje']} {plural(p['qtd_vencidos_hoje'], 'título vencido', 'títulos vencidos')}"
            if p["qtd_vencidos_hoje"] else "nada vencido — contato preventivo"
        )
        busca = sem_acento(f"{p['cliente']['nome']} {p['cliente']['contato']}")
        texto_inicial = textos_por_cliente[cod][tom_sugerido]["whatsapp"]

        itens.append(f"""
        <button type="button" class="regua-item{' ativo' if i == 0 else ''}" data-cod="{cod}" data-busca="{html.escape(busca)}"
                onclick="selecionarRegua({cod})" aria-controls="regua-{cod}">
          <span class="ri-nome">{html.escape(p['cliente']['nome'])}</span>
          <span class="ri-valor">{formatar_reais_de_centavos(valor_cobravel)}</span>
          <span class="ri-linha2"><span class="selo-tom cor-{cor}" id="ri-tom-{cod}">{html.escape(nome_tom)}</span>
            <span class="ri-feito" id="ri-feito-{cod}" hidden>{icone("check")}registrado</span></span>
        </button>""")

        blocos.append(f"""
        <div class="regua-cliente{' ativo' if i == 0 else ''}" id="regua-{cod}" data-cod="{cod}"
             data-tom-sugerido="{tom_sugerido}" data-valor="{valor_cobravel}"
             data-qtd="{len(vencidos) or len(p['titulos_abertos'])}"
             data-nome="{html.escape(p['cliente']['nome'])}">
          <div class="regua-cabecalho">
            <div class="nome">{html.escape(p['cliente']['nome'])}<span class="selo-tom cor-{cor}" id="selo-tom-{cod}">{html.escape(nome_tom)}</span>
              <div class="meta">{html.escape(p['cliente']['contato'])} · {html.escape(p['cliente']['telefone'])} · {situacao}</div>
            </div>
            <div class="valor">{formatar_reais_de_centavos(valor_cobravel)}
              <small>{'vencido' if p['qtd_vencidos_hoje'] else 'em aberto'}</small>
            </div>
          </div>

          <p class="motivo-tom"><strong>Por que a Central sugeriu
             {html.escape(nome_tom.lower())}:</strong> {prosa(motivo)}</p>

          <div class="troca-tom">
            <label for="tom-{cod}">Tom</label>
            <select id="tom-{cod}" onchange="trocarTom({cod})">{opcoes}</select>
            <span class="aviso-manual" id="manual-{cod}">{icone("alerta")}tom trocado na mão</span>
            {aviso_tons}
          </div>

          <div class="canais" role="tablist" aria-label="Canal">
            <button type="button" class="ativo" data-canal="whatsapp" onclick="trocarCanal({cod},'whatsapp')">WhatsApp</button>
            <button type="button" data-canal="email" onclick="trocarCanal({cod},'email')">E-mail</button>
            <button type="button" data-canal="ligacao" onclick="trocarCanal({cod},'ligacao')">Roteiro de ligação</button>
          </div>

          <div class="caixa-texto">
            <div class="assunto" id="assunto-{cod}" style="display:none;"></div>
            <pre id="texto-{cod}">{html.escape(texto_inicial)}</pre>
          </div>

          <div class="acoes-regua">
            <button type="button" class="botao primario" onclick="copiarTexto({cod})">{icone("copiar")}Copiar texto</button>
            <span class="ok-copiado" id="ok-{cod}">{icone("check")}copiado</span>
            <span class="separa"></span>
            <select id="canal-registro-{cod}" aria-label="Canal usado">
              <option value="WhatsApp">Registrar como WhatsApp</option>
              <option value="E-mail">Registrar como E-mail</option>
              <option value="Telefone">Registrar como Telefone</option>
            </select>
            <button type="button" class="botao" onclick="registrarCobranca({cod})">{icone("check")}Registrar que cobrei</button>
          </div>
        </div>""")

    dados_js = json.dumps(textos_por_cliente, ensure_ascii=False).replace("</", r"<\/")

    # o que ja esta no CSV vai para a pagina, para o download JUNTAR arquivo e navegador.
    # Sem isso, alguem em outra maquina (navegador vazio) baixava so o que acabou de
    # registrar, salvava por cima e apagava os meses anteriores.
    nomes = {p["cliente"]["cod"]: p["cliente"]["nome"] for p in perfis.values()}
    registros_arquivo = [{
        "cod_cliente": r["cod_cliente"],
        "nome": nomes.get(r["cod_cliente"], f"cliente {r['cod_cliente']}"),
        "dt_hora_contato": r["dt_hora_contato"].strftime("%Y-%m-%dT%H:%M:%S"),
        "canal": r["canal"],
        "tom": r["tom"],
        "valor_cobrado_centavos": r["valor_cobrado_centavos"],
        "titulos_na_cobranca": r["titulos_na_cobranca"],
    } for r in carregar_registros()]
    registros_js = json.dumps(registros_arquivo, ensure_ascii=False).replace("</", r"<\/")

    qtd_vencidos = sum(1 for p in alvos if p["qtd_vencidos_hoje"])
    qtd_outros = len(alvos) - qtd_vencidos

    if para_enviar:
        botoes_registro = (f'<button type="button" class="botao" onclick="limparRegistros()">{icone("lixo")}'
                           f'Limpar os que eu registrei</button>')
        estado = render_estado_copia(ref)
    else:
        botoes_registro = (f'<button type="button" class="botao" onclick="baixarRegistros()">{icone("baixar")}'
                           f'Baixar registros (.csv)</button>'
                           f'<button type="button" class="botao fantasma" onclick="limparRegistros()">{icone("lixo")}'
                           f'Limpar os que eu registrei</button>')
        # o caminho da pasta e o aviso do arquivo sairam da tela (ficam no terminal, ver
        # situacao_arquivo_registros): a pagina pode ser aberta em outro computador
        estado = ""

    tons_txt = "".join(
        f"<p><strong>{html.escape(info['nome'])}:</strong> {html.escape(info['quando'])}</p>"
        for info in redacao.TONS.values()
    )
    metodo = render_metodo("Como a régua escolhe o tom", f"""
      <p>O tom sai dos mesmos números das outras telas, nesta ordem, que é decisão de negócio: nada vencido →
      contato preventivo; parou de comprar → reaproximação comercial, antes de qualquer cobrança; piora confirmada →
      conversa, não ameaça; padrão estável → lembrete leve com sugestão de ajustar o prazo; atraso fora do padrão
      dele → primeiro aviso cordial; dentro do padrão → lembrete leve; senão → cobrança objetiva.</p>
      {tons_txt}
      <p>A mensagem ao cliente nunca diz quanto ele pesa no nosso faturamento: esse número fica só no roteiro
      interno da ligação.</p>""")

    return f"""
<section class="tela" id="tela-regua" role="tabpanel" aria-labelledby="aba-regua">
  <header class="topo-tela">
    {render_migalha("regua")}
    <div class="heroi-linha">
      <div class="heroi">
        <div class="heroi-rotulo">Mensagens prontas</div>
        <div class="heroi-numero">{len(alvos)}<span class="unidade">{plural(len(alvos), 'conversa', 'conversas')}</span></div>
        <div class="heroi-sub"><strong>{qtd_vencidos}</strong> com título vencido hoje{f", mais <strong>{qtd_outros}</strong> que a aba Risco mandou procurar mesmo sem nada vencido" if qtd_outros else ""}.</div>
      </div>
      <div class="heroi-lado">
        <p class="heroi-texto">O tom de cada cliente é escolhido pelos mesmos números das outras telas, e dá para trocar
          na mão que o texto é reescrito na hora. <strong>A Central redige e copia; quem envia é você.</strong> Nenhuma
          mensagem sai daqui sozinha, e não há integração com WhatsApp, e-mail ou qualquer serviço.</p>
      </div>
    </div>
  </header>

  <div class="regua-layout">
    <nav class="regua-lista" aria-label="Clientes da régua">{''.join(itens)}</nav>
    <div class="regua-detalhe">{''.join(blocos)}</div>
  </div>
  <p class="lista-vazia" id="vazio-regua" hidden>Nenhum cliente da régua com essa busca.</p>

  <section class="painel painel-registros" id="painel-registros">
    <div class="pr-cabeca"><h3>Cobranças registradas</h3><span class="pr-resumo" id="pr-resumo"></span></div>
    <p class="sub">Quem foi cobrado, quando, por qual canal, com qual tom e de quanto. Quem aparece aqui com menos de
      {JANELA_CONTATO_RECENTE_DIAS} dias desce para o fim da fila da aba Hoje. O registro é carimbado com
      <strong>{formatar_data_curta(ref)}</strong>, a data de referência da base — nunca com o relógio do computador.
      A <strong>Base_bruta.xlsx não é tocada</strong>: o que você registra fica guardado à parte.</p>
    <div id="lista-registros"></div>
    <div class="acoes-regua">{botoes_registro}</div>
    {estado}
  </section>
  {metodo}
</section>
<script id="dados-regua" type="application/json">{dados_js}</script>
<script id="registros-arquivo" type="application/json">{registros_js}</script>
"""


# ---------------------------------------------------------------------------
# Relatorio da semana (a quinta tela)
# ---------------------------------------------------------------------------

def render_relatorio_html(texto):
    """O mesmo texto do relatorio, com cabecalhos, lista e todo valor em dinheiro
    na cor de dinheiro. O texto em si nao muda: e o que o botao Copiar leva."""
    linhas = texto.split("\n")
    partes = [f'<div class="rel-titulo">{html.escape(linhas[0])}</div>',
              f'<div class="rel-sub">{html.escape(linhas[1])} · {html.escape(linhas[2])}</div>']
    lista_aberta = False
    secao = 0
    for linha in linhas[3:]:
        limpa = linha.strip()
        eh_item = re.match(r"^\d+\.\s", limpa)
        if lista_aberta and not eh_item:
            partes.append("</ol>")
            lista_aberta = False
        if not limpa:
            continue
        if limpa.isupper():
            partes.append(f'<h3 id="rel-{secao}">{html.escape(limpa)}</h3>')
            secao += 1
        elif eh_item:
            if not lista_aberta:
                partes.append("<ol>")
                lista_aberta = True
            item = re.sub(r"^\d+\.\s", "", limpa)
            partes.append(f"<li>{prosa(item)}</li>")
        elif limpa.startswith("Todos os números deste relatório"):
            partes.append(f'<p class="rel-nota">{prosa(limpa)}</p>')
        else:
            partes.append(f"<p>{prosa(limpa)}</p>")
    if lista_aberta:
        partes.append("</ol>")
    return "".join(partes)


def secoes_do_relatorio(texto):
    """Os titulos das secoes (as linhas em maiusculas), para o indice ao lado do texto."""
    return [l.strip() for l in texto.split("\n")[3:] if l.strip() and l.strip().isupper()]


def render_conteudo_relatorio(relatorio, semana, total_vencido_centavos):
    return f"""
<section class="tela" id="tela-relatorio" role="tabpanel" aria-labelledby="aba-relatorio">
  <header class="topo-tela">
    {render_migalha("relatorio")}
    <div class="heroi-linha">
      <div class="heroi">
        <div class="heroi-rotulo">Relatório da semana</div>
        <div class="heroi-numero">{semana['inicio']:%d/%m} a {semana['fim']:%d/%m}</div>
      </div>
      <div class="heroi-lado">
        <p class="heroi-texto">Texto corrido para quem tem quatro minutos e não vai abrir a Central. Os números são os
          mesmos das telas, e todo valor em dinheiro vem em <span class="dinheiro">outra cor</span> para achar sem ler.</p>
        <div class="acoes-heroi">
          <button type="button" class="botao primario" onclick="copiarRelatorio()">{icone("copiar")}Copiar texto</button>
          <button type="button" class="botao" onclick="window.print()">{icone("imprimir")}Imprimir / salvar em PDF</button>
          <span class="ok-copiado" id="ok-relatorio">{icone("check")}copiado</span>
        </div>
      </div>
    </div>
    <div class="painel painel-heroi kpis resumo-semana">
      <div class="kpi principal"><div class="kpi-valor">{formatar_reais_de_centavos(semana['recebido_centavos'])}</div>
        <div class="kpi-rotulo">entraram na semana · {semana['recebido_qtd']} {plural(semana['recebido_qtd'], 'título', 'títulos')}</div></div>
      <div class="kpi"><div class="kpi-valor">{formatar_reais_de_centavos(semana['nao_entrou_centavos'])}</div>
        <div class="kpi-rotulo">venceram e não foram pagos · {semana['nao_entrou_qtd']} {plural(semana['nao_entrou_qtd'], 'título', 'títulos')}</div></div>
      <div class="kpi"><div class="kpi-valor">{formatar_reais_de_centavos(total_vencido_centavos)}</div>
        <div class="kpi-rotulo">vencido acumulado na carteira</div></div>
    </div>
  </header>
  <div class="relatorio-layout">
    <article class="relatorio-doc" id="relatorio-doc">{render_relatorio_html(relatorio)}</article>
    <nav class="relatorio-indice" aria-label="Seções do relatório">
      <div class="desc-titulo">Neste relatório</div>
      {"".join(f'<a class="link" href="#rel-{i}" onclick="return irSecao(event, {i})">{html.escape(s.capitalize())}</a>' for i, s in enumerate(secoes_do_relatorio(relatorio)))}
    </nav>
  </div>
</section>
<script id="texto-relatorio" type="text/plain">{html.escape(relatorio)}</script>
"""


# ---------------------------------------------------------------------------
# Menu lateral, rodape e a pagina inteira
# ---------------------------------------------------------------------------

def render_trilho(ref, periodo_ini, periodo_fim, contagens, para_enviar):
    abas = []
    for i, (chave, nome, _, icn, curto) in enumerate(TELAS):
        conta = contagens.get(chave)
        conta_html = f'<span class="aba-conta">{conta}</span>' if conta is not None else ""
        abas.append(
            f'<a href="#tela-{chave}" class="aba{" ativa" if i == 0 else ""}" id="aba-{chave}" role="tab" '
            f'aria-selected="{"true" if i == 0 else "false"}" aria-controls="tela-{chave}" '
            f'onclick="return irPara(event, \'{chave}\')">{icone(icn)}<span class="aba-nome">{html.escape(nome)}</span>'
            f'<span class="aba-curta">{html.escape(curto)}</span>{conta_html}</a>'
        )
    copia = (f'<span class="copia-selo">Cópia enviada · dados de {formatar_data_curta(ref)}. '
             f'Os números não se atualizam nesta cópia.</span>') if para_enviar else ""
    return f"""
<aside class="trilho" aria-label="Navegação da Central">
  <div class="marca">
    <div class="marca-sinal" aria-hidden="true">A</div>
    <div><div class="marca-nome">Central de Crédito e Cobrança</div><div class="marca-sub">Distribuidora Aurora</div></div>
  </div>
  <div class="data-ref"><span>Hoje é</span><strong>{dia_semana_pt(ref)}, {formatar_data_pt(ref)}</strong></div>
  <div class="busca">
    <label class="busca-campo">{icone("busca")}
      <input id="busca" type="search" placeholder="Buscar cliente" autocomplete="off" aria-label="Buscar cliente"
             aria-controls="busca-resultados">
      <kbd>/</kbd>
    </label>
    <div class="busca-resultados" id="busca-resultados" hidden></div>
  </div>
  <nav class="menu" role="tablist" aria-label="Telas da Central">
    <div class="menu-titulo">As cinco telas</div>
    {''.join(abas)}
  </nav>
  <div class="trilho-pe">
    {copia}
    <strong>Histórico:</strong> {formatar_data_curta(periodo_ini)} a {formatar_data_curta(periodo_fim)}<br>
    Data lida da planilha, não do relógio do computador.
  </div>
</aside>"""


def render_rodape(ref, titulos, para_enviar):
    atualizar = (
        "Esta é uma cópia enviada: os números ficam como estavam no dia em que ela foi gerada. Para números novos, "
        "peça uma cópia atualizada a quem te mandou."
        if para_enviar else
        'Para atualizar, rode <strong>python central.py</strong> de novo (ou dê dois cliques em "Abrir Central.bat").'
    )
    return f"""
<footer class="rodape">
  <details>
    <summary>{icone("info")}<span><strong>Dados de {formatar_data_pt(ref)}</strong> · Base_bruta.xlsx em modo somente
      leitura · nada digitado à mão</span></summary>
    <p>A data de referência é lida da aba Leia-me da planilha, não do relógio do computador. Todo "hoje", "dias em
    atraso" e "há quantos dias" nesta página é contado contra essa data, então os números não mudam sozinhos amanhã.
    A planilha não é alterada. Nenhum valor desta página foi digitado à mão: tudo é calculado a partir dos
    {len(titulos)} títulos da base. {atualizar}</p>
  </details>
</footer>"""


JS_PAGINA = r"""
var TELAS = ['hoje', 'risco', 'previsao', 'regua', 'relatorio'];
var REF_ISO = '__REF_ISO__';
var JANELA_DIAS = __JANELA__;
var CHAVE = '__CHAVE__';
var PARA_ENVIAR = __PARA_ENVIAR__;
var TEXTOS = JSON.parse(document.getElementById('dados-regua').textContent);
var REG_ARQUIVO = JSON.parse(document.getElementById('registros-arquivo').textContent);
var NOMES_TOM = __NOMES_TOM__;
var CORES_TOM = __CORES_TOM__;
var COLUNAS_CSV = __COLUNAS__;
var canalAtual = {};
var ordemHoje = 'rank';
var filtroHoje = 'todos';
var ordemRisco = 'em-risco';
var filtroRisco = 'todos';
var CLIENTES = [];

/* ---------- registro de cobranca: navegador + arquivo ---------- */
function lerRegistros() {
  try { return JSON.parse(localStorage.getItem(CHAVE) || '[]'); } catch (e) { return []; }
}
function gravarRegistros(lista) {
  try { localStorage.setItem(CHAVE, JSON.stringify(lista)); return true; } catch (e) {
    alert('Não consegui guardar o registro neste navegador.' +
          (PARA_ENVIAR ? '' : ' Use o botão "Baixar registros".'));
    return false;
  }
}
function chaveRegistro(r) {
  return [r.cod_cliente, r.dt_hora_contato.substring(0, 16), r.canal, r.tom, r.valor_cobrado_centavos].join('|');
}
function registrosSoNoNavegador() {
  var noArquivo = {};
  REG_ARQUIVO.forEach(function(r) { noArquivo[chaveRegistro(r)] = true; });
  return lerRegistros().filter(function(r) { return !noArquivo[chaveRegistro(r)]; });
}
function todosRegistros() {
  // arquivo + o que so este navegador tem, sem repetir
  return REG_ARQUIVO.concat(registrosSoNoNavegador());
}
function esc(t) {
  return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function formatarReais(centavos) {
  return 'R$ ' + (centavos / 100).toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}
function diasDesdeRef(iso) {
  var refData = new Date(REF_ISO + 'T00:00:00');
  var quando = new Date(iso.substring(0, 10) + 'T00:00:00');
  return Math.round((refData - quando) / 86400000);
}

function registrar(cod, canal, tom) {
  var bloco = document.getElementById('regua-' + cod);
  if (!bloco) return;
  var lista = lerRegistros();
  lista.push({
    cod_cliente: parseInt(cod, 10),
    nome: bloco.dataset.nome,
    // carimbo pelo relogio da BASE, nunca pelo do computador
    dt_hora_contato: REF_ISO + 'T09:00:00',
    canal: canal,
    tom: NOMES_TOM[tom],
    valor_cobrado_centavos: parseInt(bloco.dataset.valor, 10),
    titulos_na_cobranca: parseInt(bloco.dataset.qtd, 10)
  });
  if (!gravarRegistros(lista)) return;
  renderRegistros();
  aplicarRegistrosNaFila();
  marcarRegistrados();
  var naFila = document.querySelector('#tela-hoje .linha[data-cod="' + cod + '"]');
  aviso('Cobrança de ' + bloco.dataset.nome + ' registrada' +
        (naFila ? ': desceu para o fim da fila de Hoje.' : '.'), true);
}
function registrarCobranca(cod) {
  registrar(cod, document.getElementById('canal-registro-' + cod).value, document.getElementById('tom-' + cod).value);
}
function registrarLigacao(cod) {
  var sel = document.getElementById('tom-' + cod);
  var bloco = document.getElementById('regua-' + cod);
  registrar(cod, 'Telefone', sel ? sel.value : bloco.dataset.tomSugerido);
}
function desfazerUltimo() {
  var lista = lerRegistros();
  if (!lista.length) return;
  lista.pop();
  gravarRegistros(lista);
  renderRegistros();
  aplicarRegistrosNaFila();
  marcarRegistrados();
  aviso('Registro desfeito.');
}
function limparRegistros() {
  if (!confirm(PARA_ENVIAR
      ? 'Isso apaga as cobranças que você marcou neste navegador. Confirma?'
      : 'Isso apaga as cobranças guardadas NESTE navegador. As que já estão no arquivo ' +
        'cobrancas_registradas.csv continuam. Confirma?')) return;
  gravarRegistros([]);
  renderRegistros();
  aplicarRegistrosNaFila();
  marcarRegistrados();
}
function campoCsv(v) {
  var t = String(v == null ? '' : v);
  return (t.indexOf(';') >= 0 || t.indexOf('"') >= 0 || t.indexOf('\n') >= 0)
    ? '"' + t.replace(/"/g, '""') + '"' : t;
}
function baixarRegistros() {
  if (PARA_ENVIAR) return;
  var lista = todosRegistros();
  if (!lista.length) { alert('Nenhuma cobrança registrada ainda.'); return; }
  var linhas = [COLUNAS_CSV.join(';')];
  lista.forEach(function(r) {
    var d = r.dt_hora_contato.substring(0, 10).split('-');
    var reais = (r.valor_cobrado_centavos / 100).toFixed(2).replace('.', ',');
    linhas.push([
      r.cod_cliente, r.nome, d[2] + '/' + d[1] + '/' + d[0],
      r.dt_hora_contato.substring(11, 16), r.canal, r.tom, reais, r.titulos_na_cobranca
    ].map(campoCsv).join(';'));
  });
  // BOM + ponto-e-virgula: e o que o Excel em portugues espera encontrar,
  // senao os acentos quebram e tudo cai numa coluna so
  var blob = new Blob(['﻿' + linhas.join('\r\n')], {type: 'text/csv;charset=utf-8'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'cobrancas_registradas.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  alert('Salve o arquivo na mesma pasta da Central, por cima do anterior.\n\n' +
        'Ele abre no Excel, então dá para conferir o que foi registrado.\n\n' +
        'Na próxima vez que rodar "python central.py", esses registros entram na conta de verdade.');
}
function renderRegistros() {
  var alvo = document.getElementById('lista-registros');
  if (!alvo) return;
  var soAqui = {};
  var so = registrosSoNoNavegador();
  so.forEach(function(r) { soAqui[chaveRegistro(r)] = true; });
  var lista = todosRegistros().slice().sort(function(a, b) {
    return a.dt_hora_contato < b.dt_hora_contato ? 1 : -1;
  });
  var resumo = document.getElementById('pr-resumo');
  if (resumo) {
    resumo.textContent = lista.length
      ? lista.length + (lista.length === 1 ? ' registrada' : ' registradas') +
        (so.length ? ' · ' + so.length + ' só neste navegador' : '')
      : 'nenhuma ainda';
  }
  if (!lista.length) {
    alvo.innerHTML = '<p class="vazio-registros">Nenhuma cobrança registrada pela Central ainda. ' +
      'O que aparece na aba Hoje como "último contato" veio da planilha.</p>';
    return;
  }
  var linhas = '';
  lista.forEach(function(r) {
    var d = r.dt_hora_contato.substring(0, 10).split('-');
    var origem = soAqui[chaveRegistro(r)]
      ? (PARA_ENVIAR ? '<span class="origem registrado">só neste navegador</span>'
                     : '<span class="origem registrado">só neste navegador — baixe o arquivo</span>')
      : '<span class="origem planilha">já no arquivo</span>';
    linhas += '<tr><td>' + esc(r.nome) + origem + '</td>' +
      '<td>' + d[2] + '/' + d[1] + '/' + d[0] + '</td>' +
      '<td>' + esc(r.canal) + '</td><td>' + esc(r.tom) + '</td>' +
      '<td class="dir">' + formatarReais(r.valor_cobrado_centavos) + '</td></tr>';
  });
  alvo.innerHTML = '<div class="rolagem-tabela"><table class="tabela-registros"><thead><tr><th>Cliente</th><th>Quando</th>' +
    '<th>Canal</th><th>Tom</th><th class="dir">Valor</th></tr></thead><tbody>' + linhas + '</tbody></table></div>';
}

/* ---------- fila de Hoje: quem foi cobrado desce, e a ordem escolhida ---------- */
function aplicarRegistrosNaFila() {
  // so o que ainda nao esta no arquivo: o que ja esta, o Python ja pos na fila
  var lista = registrosSoNoNavegador();
  var recentes = {};
  for (var i = 0; i < lista.length; i++) {
    if (diasDesdeRef(lista[i].dt_hora_contato) <= JANELA_DIAS) recentes[lista[i].cod_cliente] = lista[i];
  }
  var fila = document.querySelector('#tela-hoje .fila');
  if (!fila) return;
  var linhas = Array.prototype.slice.call(fila.querySelectorAll('.linha'));
  linhas.forEach(function(el) {
    var cod = parseInt(el.dataset.cod, 10);
    var recemRegistrado = !!recentes[cod];
    var jaEraRecente = el.dataset.recente === '1';
    var avisoAntigo = el.querySelector('.aviso-registro');
    if (avisoAntigo) avisoAntigo.remove();
    el.classList.toggle('recente', recemRegistrado || jaEraRecente);
    el.dataset.cobrado = (recemRegistrado || jaEraRecente) ? '1' : '0';
    if (recemRegistrado) {
      var r = recentes[cod];
      var aviso = document.createElement('p');
      aviso.className = 'contato-bloco aviso-registro';
      aviso.innerHTML = '<span class="contato-nome">Você registrou esta cobrança</span>' +
        '<span>' + esc(r.canal) + ' · ' + esc(r.tom) + ' · ' + formatarReais(r.valor_cobrado_centavos) +
        ' — desceu para o fim da fila de hoje</span>';
      var ref = el.querySelector('.explicacao');
      (ref && ref.parentNode ? ref.parentNode : el).insertBefore(aviso, ref);
    }
  });
  ordenarFila();
  atualizarChipsHoje();
}
function ordenarFila() {
  var fila = document.querySelector('#tela-hoje .fila');
  if (!fila) return;
  var linhas = Array.prototype.slice.call(fila.querySelectorAll('.linha'));
  function grupo(el) { return el.dataset.cobrado === '1' ? 1 : 0; }
  function num(el, k) { return parseFloat(el.dataset[k]); }
  // a regra e a mesma do Python: quem foi cobrado desce para o fim,
  // e dentro de cada grupo vale a ordem de prioridade ja calculada
  var prioridade = linhas.slice().sort(function(a, b) {
    return (grupo(a) - grupo(b)) || (num(a, 'rank') - num(b, 'rank'));
  });
  prioridade.forEach(function(el, i) {
    el.querySelector('.posicao').textContent = i + 1;
    el.classList.toggle('primeiro', i === 0);
  });
  var mostrada = prioridade;
  if (ordemHoje !== 'rank') {
    mostrada = prioridade.slice().sort(function(a, b) {
      return (grupo(a) - grupo(b)) || (num(b, ordemHoje) - num(a, ordemHoje)) || (num(a, 'rank') - num(b, 'rank'));
    });
  }
  mostrada.forEach(function(el) { fila.appendChild(el); });
}
function situacaoEfetiva(el) {
  return el.dataset.cobrado === '1' ? 'aguardando' : el.dataset.situacao;
}
function atualizarChipsHoje() {
  if (!document.querySelectorAll) return;
  var cartoes = Array.prototype.slice.call(document.querySelectorAll('#tela-hoje .linha'));
  var conta = {};
  cartoes.forEach(function(el) { var s = situacaoEfetiva(el); conta[s] = (conta[s] || 0) + 1; });
  Array.prototype.slice.call(document.querySelectorAll('#chips-hoje .chip')).forEach(function(ch) {
    var f = ch.dataset.filtro;
    if (f === 'todos') return;
    var n = conta[f] || 0;
    ch.querySelector('b').textContent = n;
    ch.hidden = !n && filtroHoje !== f;
  });
  aplicarFiltroHoje();
}
function filtrarHoje(f) {
  filtroHoje = f;
  marcarChip('chips-hoje', f);
  atualizarChipsHoje();
}
function ordenarHoje(k) { ordemHoje = k; ordenarFila(); }
function marcarChip(id, f) {
  Array.prototype.slice.call(document.querySelectorAll('#' + id + ' .chip')).forEach(function(ch) {
    ch.classList.toggle('ativo', ch.dataset.filtro === f);
  });
}
function textoBusca() {
  var i = document.getElementById('busca');
  return i ? normalizar(i.value) : '';
}
function normalizar(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}
function aplicarFiltroHoje() {
  if (!document.querySelectorAll) return;
  var q = textoBusca();
  var visiveis = 0;
  Array.prototype.slice.call(document.querySelectorAll('#tela-hoje .linha')).forEach(function(el) {
    var ok = (filtroHoje === 'todos' || situacaoEfetiva(el) === filtroHoje) && (!q || el.dataset.busca.indexOf(q) >= 0);
    el.hidden = !ok;
    if (ok) visiveis++;
  });
  var v = document.getElementById('vazio-hoje');
  if (v) v.hidden = visiveis > 0;
}
function alternarDobra(cod, botao) {
  var d = document.getElementById('dobra-' + cod);
  var aberta = d.classList.toggle('aberta');
  botao.setAttribute('aria-expanded', aberta ? 'true' : 'false');
}
function marcarRegistrados() {
  if (!document.querySelectorAll) return;
  var feitos = {};
  todosRegistros().forEach(function(r) {
    if (diasDesdeRef(r.dt_hora_contato) <= JANELA_DIAS) feitos[r.cod_cliente] = true;
  });
  var soAqui = {};
  registrosSoNoNavegador().forEach(function(r) { soAqui[r.cod_cliente] = true; });
  Array.prototype.slice.call(document.querySelectorAll('#tela-hoje .linha')).forEach(function(el) {
    var cod = el.dataset.cod;
    var b = document.getElementById('ligou-' + cod);
    if (!b) return;
    var feito = !!soAqui[cod];
    b.classList.toggle('feito', feito);
    b.disabled = feito;
    b.querySelector('span').textContent = feito ? 'Ligação registrada' : 'Registrei a ligação';
  });
  Array.prototype.slice.call(document.querySelectorAll('.regua-item')).forEach(function(el) {
    var f = document.getElementById('ri-feito-' + el.dataset.cod);
    if (f) f.hidden = !feitos[el.dataset.cod];
  });
}

/* ---------- Risco: filtro por faixa e outra ordem ---------- */
function filtrarRisco(f) { filtroRisco = f; marcarChip('chips-risco', f); aplicarFiltroRisco(); }
function aplicarFiltroRisco() {
  if (!document.querySelectorAll) return;
  var q = textoBusca();
  var visiveis = 0;
  Array.prototype.slice.call(document.querySelectorAll('#lista-risco .risco-linha')).forEach(function(el) {
    var okF = filtroRisco === 'todos' || (filtroRisco === 'selo' ? el.dataset.selo === '1' : el.dataset.faixa === filtroRisco);
    var ok = okF && (!q || el.dataset.busca.indexOf(q) >= 0);
    el.hidden = !ok;
    if (ok) visiveis++;
  });
  var v = document.getElementById('vazio-risco');
  if (v) v.hidden = visiveis > 0;
}
function ordenarRisco(k) {
  ordemRisco = k;
  var lista = document.getElementById('lista-risco');
  var linhas = Array.prototype.slice.call(lista.querySelectorAll('.risco-linha'));
  var campo = {'em-risco': 'emRisco', 'indice': 'indice', 'saldo': 'saldo'}[k];
  linhas.sort(function(a, b) {
    return (parseFloat(b.dataset[campo]) - parseFloat(a.dataset[campo])) ||
           (parseFloat(b.dataset.emRisco) - parseFloat(a.dataset.emRisco));
  });
  linhas.forEach(function(el) { lista.appendChild(el); });
}

/* ---------- Regua: lista de um lado, um cliente do outro ---------- */
function selecionarRegua(cod) {
  cod = String(cod);
  Array.prototype.slice.call(document.querySelectorAll('.regua-item')).forEach(function(el) {
    var sim = el.dataset.cod === cod;
    el.classList.toggle('ativo', sim);
    el.setAttribute('aria-current', sim ? 'true' : 'false');
  });
  Array.prototype.slice.call(document.querySelectorAll('.regua-cliente')).forEach(function(el) {
    el.classList.toggle('ativo', el.dataset.cod === cod);
  });
}
function aplicarFiltroRegua() {
  if (!document.querySelectorAll) return;
  var q = textoBusca();
  var visiveis = 0;
  Array.prototype.slice.call(document.querySelectorAll('.regua-item')).forEach(function(el) {
    var ok = !q || el.dataset.busca.indexOf(q) >= 0;
    el.hidden = !ok;
    if (ok) visiveis++;
  });
  var v = document.getElementById('vazio-regua');
  if (v) v.hidden = visiveis > 0;
}
function trocarTom(cod) {
  var tom = document.getElementById('tom-' + cod).value;
  var bloco = document.getElementById('regua-' + cod);
  // a cor tem que seguir o tom escolhido, senao o cartao diz uma coisa no texto e outra na cor
  ['selo-tom-', 'ri-tom-'].forEach(function(p) {
    var el = document.getElementById(p + cod);
    if (!el) return;
    el.textContent = NOMES_TOM[tom];
    el.className = 'selo-tom cor-' + CORES_TOM[tom];
  });
  document.getElementById('manual-' + cod).style.display =
    (tom === bloco.dataset.tomSugerido) ? 'none' : 'inline-flex';
  mostrarTexto(cod);
}
function trocarCanal(cod, canal) {
  canalAtual[cod] = canal;
  var botoes = document.querySelectorAll('#regua-' + cod + ' .canais button');
  for (var i = 0; i < botoes.length; i++) {
    botoes[i].classList.toggle('ativo', botoes[i].dataset.canal === canal);
  }
  mostrarTexto(cod);
}
function mostrarTexto(cod) {
  var tom = document.getElementById('tom-' + cod).value;
  var canal = canalAtual[cod] || 'whatsapp';
  var t = TEXTOS[cod][tom];
  var caixaAssunto = document.getElementById('assunto-' + cod);
  if (canal === 'email') {
    caixaAssunto.style.display = 'block';
    caixaAssunto.textContent = 'Assunto: ' + t.email_assunto;
  } else {
    caixaAssunto.style.display = 'none';
  }
  document.getElementById('texto-' + cod).textContent = t[canal];
}
function copiar(texto, depois) {
  function porComando() {
    var area = document.createElement('textarea');
    area.value = texto;
    area.setAttribute('readonly', '');
    area.style.position = 'fixed';
    area.style.opacity = '0';
    document.body.appendChild(area);
    area.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) {}
    document.body.removeChild(area);
    depois(ok);
  }
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(texto).then(function() { depois(true); }, porComando);
  } else {
    porComando();
  }
}
function mostrarOk(id) {
  var ok = document.getElementById(id);
  if (!ok) return;
  ok.style.display = 'inline-flex';
  setTimeout(function() { ok.style.display = 'none'; }, 2000);
}
function copiarTexto(cod) {
  copiar(document.getElementById('texto-' + cod).textContent, function(ok) {
    if (ok) { mostrarOk('ok-' + cod); aviso('Texto copiado. Cole no ' + ({whatsapp: 'WhatsApp', email: 'e-mail', ligacao: 'bloco de notas'})[canalAtual[cod] || 'whatsapp'] + '.'); }
    else { aviso('O navegador não deixou copiar. Selecione o texto e use Ctrl+C.'); }
  });
}
function copiarRelatorio() {
  copiar(document.getElementById('texto-relatorio').textContent, function(ok) {
    if (ok) { mostrarOk('ok-relatorio'); aviso('Relatório copiado.'); }
    else { aviso('O navegador não deixou copiar. Selecione o texto e use Ctrl+C.'); }
  });
}

/* ---------- navegacao entre as telas ---------- */
function mostrarTela(nome, semRolar) {
  if (TELAS.indexOf(nome) < 0) nome = 'hoje';
  TELAS.forEach(function(n) {
    var ativa = (n === nome);
    document.getElementById('tela-' + n).classList.toggle('ativa', ativa);
    var aba = document.getElementById('aba-' + n);
    aba.classList.toggle('ativa', ativa);
    aba.setAttribute('aria-selected', ativa ? 'true' : 'false');
  });
  try { history.replaceState(null, '', '#' + nome); } catch (e) {}
  if (!semRolar) window.scrollTo(0, 0);
}
function irPara(evento, nome) {
  if (evento && evento.preventDefault) evento.preventDefault();
  mostrarTela(nome);
  return false;
}
function irParaRegua(cod) {
  fecharFicha();
  mostrarTela('regua');
  selecionarRegua(cod);
}
function irParaFila(cod) {
  fecharFicha();
  var el = document.querySelector('#tela-hoje .linha[data-cod="' + cod + '"]');
  if (!el) { abrirFicha(cod); return; }
  mostrarTela('hoje', true);
  filtroHoje = 'todos';
  marcarChip('chips-hoje', 'todos');
  aplicarFiltroHoje();
  el.hidden = false;
  el.scrollIntoView({block: 'center', behavior: 'smooth'});
  el.classList.add('destaque');
  setTimeout(function() { el.classList.remove('destaque'); }, 1800);
}
function abrirFicha(cod) {
  document.querySelectorAll('.ficha-modal').forEach(function(el) { el.style.display = 'none'; });
  var alvo = document.getElementById('ficha-' + cod);
  if (!alvo) return;
  alvo.style.display = 'block';
  alvo.scrollTop = 0;
  document.getElementById('modal-fundo').classList.add('aberto');
  document.body.style.overflow = 'hidden';
  var fechar = alvo.querySelector('.fechar');
  if (fechar) fechar.focus();
}
function fecharFicha() {
  var fundo = document.getElementById('modal-fundo');
  if (!fundo) return;
  fundo.classList.remove('aberto');
  document.body.style.overflow = '';
}

/* ---------- busca ---------- */
function aoBuscar() {
  aplicarFiltroHoje();
  aplicarFiltroRisco();
  aplicarFiltroRegua();
  var caixa = document.getElementById('busca-resultados');
  var q = textoBusca();
  if (!q) { caixa.hidden = true; caixa.innerHTML = ''; return; }
  var achados = CLIENTES.filter(function(c) { return c.busca.indexOf(q) >= 0; }).slice(0, 8);
  if (!achados.length) {
    caixa.innerHTML = '<div class="br-vazio">Nenhum cliente com esse nome.</div>';
  } else {
    caixa.innerHTML = achados.map(function(c) {
      var acoes = '';
      if (c.fila) acoes += '<button type="button" class="link" onclick="escolherBusca(' + c.cod + ',\'fila\')">Na fila</button>';
      acoes += '<button type="button" class="link" onclick="escolherBusca(' + c.cod + ',\'ficha\')">Ficha</button>';
      if (c.regua) acoes += '<button type="button" class="link" onclick="escolherBusca(' + c.cod + ',\'regua\')">Mensagem</button>';
      return '<div class="br-item"><div class="br-nome">' + esc(c.nome) + '</div><div class="br-acoes">' + acoes + '</div></div>';
    }).join('');
  }
  caixa.hidden = false;
}
function escolherBusca(cod, onde) {
  document.getElementById('busca-resultados').hidden = true;
  if (onde === 'fila') irParaFila(cod);
  else if (onde === 'regua') irParaRegua(cod);
  else abrirFicha(cod);
}
function buscaEnter() {
  var q = textoBusca();
  var c = CLIENTES.filter(function(x) { return x.busca.indexOf(q) >= 0; })[0];
  if (!c) return;
  var tela = (location.hash || '#hoje').replace('#', '').replace('tela-', '');
  if (tela === 'hoje' && c.fila) escolherBusca(c.cod, 'fila');
  else if (tela === 'regua' && c.regua) escolherBusca(c.cod, 'regua');
  else escolherBusca(c.cod, 'ficha');
}

function irSecao(evento, i) {
  if (evento && evento.preventDefault) evento.preventDefault();
  var alvo = document.getElementById('rel-' + i);
  if (alvo) alvo.scrollIntoView({block: 'start', behavior: 'smooth'});
  return false;
}

/* ---------- aviso flutuante ---------- */
function aviso(msg, comDesfazer) {
  var el = document.getElementById('aviso-flutuante');
  if (!el) return;
  el.innerHTML = '<span>' + esc(msg) + '</span>' +
    (comDesfazer ? '<button type="button" class="link" onclick="desfazerUltimo()">Desfazer</button>' : '');
  el.classList.add('visivel');
  clearTimeout(aviso.t);
  aviso.t = setTimeout(function() { el.classList.remove('visivel'); }, comDesfazer ? 7000 : 2800);
}

document.addEventListener('DOMContentLoaded', function() {
  try { CLIENTES = JSON.parse(document.getElementById('dados-clientes').textContent); } catch (e) { CLIENTES = []; }
  Object.keys(TEXTOS).forEach(function(cod) { mostrarTexto(cod); });
  renderRegistros();
  aplicarRegistrosNaFila();
  marcarRegistrados();
  var primeiro = document.querySelector('.regua-item');
  if (primeiro) selecionarRegua(primeiro.dataset.cod);
  var alvo = (location.hash || '').replace('#', '').replace('tela-', '');
  mostrarTela(TELAS.indexOf(alvo) >= 0 ? alvo : 'hoje', true);
  var busca = document.getElementById('busca');
  if (busca) {
    busca.addEventListener('input', aoBuscar);
    busca.addEventListener('keydown', function(e) {
      if (e.key === 'Enter') { e.preventDefault(); buscaEnter(); }
      if (e.key === 'Escape') { busca.value = ''; aoBuscar(); busca.blur(); }
    });
  }
  document.addEventListener('click', function(e) {
    var caixa = document.getElementById('busca-resultados');
    if (caixa && !caixa.hidden && !e.target.closest('.busca')) caixa.hidden = true;
  });
});
// o endereco manda na tela: abrir "...central.html#risco", ou voltar no navegador,
// mostra a tela certa (antes so valia na primeira carga da pagina)
if (window.addEventListener) {
  window.addEventListener('hashchange', function() {
    var alvo = (location.hash || '').replace('#', '').replace('tela-', '');
    if (TELAS.indexOf(alvo) >= 0) mostrarTela(alvo);
  });
}
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') { fecharFicha(); }
  var digitando = e.target && /INPUT|TEXTAREA|SELECT/.test(e.target.tagName);
  if (!digitando && (e.key === '/' || ((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')))) {
    var b = document.getElementById('busca');
    if (b) { e.preventDefault(); b.focus(); b.select(); }
  }
  if (!digitando && (e.key === 'ArrowDown' || e.key === 'ArrowUp') && e.target.classList &&
      e.target.classList.contains('regua-item')) {
    e.preventDefault();
    var itens = Array.prototype.slice.call(document.querySelectorAll('.regua-item')).filter(function(x) { return !x.hidden; });
    var i = itens.indexOf(e.target) + (e.key === 'ArrowDown' ? 1 : -1);
    if (itens[i]) { itens[i].focus(); selecionarRegua(itens[i].dataset.cod); }
  }
});
"""


def render_pagina(ref, clientes, titulos, fila, perfis, prev, semana, relatorio, descobertas, para_enviar=False):
    emissoes = [t["dt_emissao"] for t in titulos]
    periodo_ini, periodo_fim = min(emissoes), max(emissoes)
    alvos_regua = clientes_da_regua(perfis)
    cods_regua = {p["cliente"]["cod"] for p in alvos_regua}
    cods_fila = {c["cliente"]["cod"] for c in fila}

    conteudo_hoje = render_conteudo_hoje(ref, clientes, titulos, fila, perfis, descobertas, cods_regua)
    conteudo_risco, fichas_html = render_conteudo_risco(ref, perfis, meses_de_historico(titulos, ref))
    conteudo_previsao = render_conteudo_previsao(ref, prev, semana)
    conteudo_regua = render_conteudo_regua(ref, perfis, para_enviar)
    total_vencido = sum(t["valor_centavos"] for t in titulos if t["dt_pagamento"] is None and t["dt_vencimento"] < ref)
    conteudo_relatorio = render_conteudo_relatorio(relatorio, semana, total_vencido)

    contagens = {"hoje": len(fila), "risco": len(perfis), "previsao": None, "regua": len(alvos_regua), "relatorio": None}

    # a busca do menu lateral acha qualquer cliente da base e leva para onde ele estiver
    busca_js = json.dumps([{
        "cod": cod,
        "nome": p["cliente"]["nome"],
        "busca": sem_acento(f"{p['cliente']['nome']} {p['cliente']['contato']} {p['cliente']['categoria']} {p['cliente']['regiao']}"),
        "fila": cod in cods_fila,
        "regua": cod in cods_regua,
    } for cod, p in sorted(perfis.items(), key=lambda x: x[1]["cliente"]["nome"])], ensure_ascii=False).replace("</", r"<\/")

    cores_tom = {t: COR_DO_TOM[v["cor"]] for t, v in redacao.TONS.items()}
    js = (JS_PAGINA
          .replace("__REF_ISO__", ref.isoformat())
          .replace("__JANELA__", str(JANELA_CONTATO_RECENTE_DIAS))
          # a copia enviada guarda os registros em outra chave: se alguem abrir as duas na
          # mesma maquina, o que foi marcado na copia nao vai parar no arquivo oficial
          .replace("__CHAVE__", "aurora_cobrancas_copia_enviada" if para_enviar else "aurora_cobrancas_registradas")
          .replace("__PARA_ENVIAR__", "true" if para_enviar else "false")
          .replace("__NOMES_TOM__", json.dumps({t: v["nome"] for t, v in redacao.TONS.items()}, ensure_ascii=False))
          .replace("__CORES_TOM__", json.dumps(cores_tom))
          .replace("__COLUNAS__", json.dumps(COLUNAS_REGISTRO)))

    titulo_aba = "Central de Crédito e Cobrança — Distribuidora Aurora" + (" (cópia enviada)" if para_enviar else "")
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark">
<title>{html.escape(titulo_aba)}</title>
<script id="marca-js">document.documentElement.className += ' js';</script>
<style>{CSS.replace("__FONTES__", css_fontes())}</style>
</head>
<body>
{sprite_icones()}
<div class="ceu" aria-hidden="true"><div class="estrelas"></div><div class="cortina"></div>
  <div class="luz l3"></div><div class="luz l1"></div><div class="luz l2"></div><div class="luz l4"></div></div>
<div class="app">
  {render_trilho(ref, periodo_ini, periodo_fim, contagens, para_enviar)}
  <main class="palco" id="conteudo">
    <div class="palco-interno">
      {conteudo_hoje}
      {conteudo_risco}
      {conteudo_previsao}
      {conteudo_regua}
      {conteudo_relatorio}
      {render_rodape(ref, titulos, para_enviar)}
    </div>
  </main>
</div>
<div class="modal-fundo" id="modal-fundo" onclick="if(event.target===this) fecharFicha()">
  {fichas_html}
</div>
<div class="aviso-flutuante" id="aviso-flutuante" role="status" aria-live="polite"></div>
<script id="dados-clientes" type="application/json">{busca_js}</script>
<script>{js}</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def main():
    if not os.path.exists(ARQUIVO_BASE):
        print(f"Nao encontrei {ARQUIVO_BASE}")
        sys.exit(1)

    ref, clientes, titulos, cobrancas = carregar_dados()
    perfis = calcular_perfis_risco(ref, clientes, titulos, cobrancas)
    fila = calcular_fila(ref, clientes, titulos, cobrancas, perfis)
    previsao = calcular_previsao(ref, perfis)
    semana = calcular_semana(ref, titulos)
    registros = [c for c in cobrancas if c.get("registrado_na_tela")]
    relatorio = gerar_relatorio_texto(ref, clientes, titulos, fila, perfis, previsao, semana, registros)
    descobertas = calcular_descobertas(perfis)
    dados = (ref, clientes, titulos, fila, perfis, previsao, semana, relatorio, descobertas)

    with open(ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
        f.write(render_pagina(*dados))

    # a copia para mandar para alguem da equipe: o mesmo arquivo unico, sem o que
    # so funciona nesta pasta (baixar o CSV, o caminho da pasta, "rode python")
    os.makedirs(PASTA_ENVIO, exist_ok=True)
    with open(ARQUIVO_ENVIO, "w", encoding="utf-8") as f:
        f.write(render_pagina(*dados, para_enviar=True))

    print(f"Central gerada: {ARQUIVO_SAIDA}")
    print(f"Copia para enviar: {ARQUIVO_ENVIO}")
    print(f"Data de referencia usada: {ref.isoformat()}")
    print(f"{len(fila)} {plural(len(fila), 'cliente', 'clientes')} na fila de cobranca de hoje.")
    print(f"{len(perfis)} {plural(len(perfis), 'cliente', 'clientes')} com perfil de risco calculado.")
    for linha in situacao_arquivo_registros():
        print(linha)

    webbrowser.open(pathlib.Path(ARQUIVO_SAIDA).as_uri())


if __name__ == "__main__":
    main()
