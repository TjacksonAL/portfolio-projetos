#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conferencia da Central de Credito e Cobranca
=============================================

Garante que mexer numa tela nao muda silenciosamente o numero de outra, e que
nenhum item do o_que_nao_pode_sumir.md sumiu da pagina (no central.html e na
copia para enviar, para_enviar/Central Aurora.html).

  python conferir.py --gravar-base   grava o retrato dos numeros de hoje
  python conferir.py                 recalcula e compara com o retrato gravado

Saida: 0 = aprovada | 1 = reprovada | 3 = planilha nova, estrutura ok, ainda falta
rodar a auditoria/reconferir.py e regravar a base (passo a passo no CLAUDE.md)

A linha de base fica em conferencia_base.json. Se uma verificacao quebrar, o
conserto e no codigo -- nao na verificacao nem na base.

So usa biblioteca padrao, igual ao central.py.
"""

import html
import json
import os
import re
import sys

import central
import redacao

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

PASTA = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_BASE = os.path.join(PASTA, "conferencia_base.json")


def retrato():
    """Todos os numeros que as telas mostram, num dicionario comparavel."""
    ref, clientes, titulos, cobrancas = central.carregar_dados()
    perfis = central.calcular_perfis_risco(ref, clientes, titulos, cobrancas)
    fila = central.calcular_fila(ref, clientes, titulos, cobrancas, perfis)
    prev = central.calcular_previsao(ref, perfis)

    abertos = [t for t in titulos if t["dt_pagamento"] is None]
    vencidos = [t for t in abertos if t["dt_vencimento"] < ref]

    cod_maior = max(fila, key=lambda c: c["valor_vencido_centavos"])["cliente"]["cod"] if fila else None

    r = {
        "data_referencia": ref.isoformat(),
        "totais": {
            "titulos": len(titulos),
            "clientes": len(clientes),
            "abertos_qtd": len(abertos),
            "abertos_centavos": sum(t["valor_centavos"] for t in abertos),
            "vencidos_qtd": len(vencidos),
            "vencidos_centavos": sum(t["valor_centavos"] for t in vencidos),
        },
        "hoje": [],
        "risco": [],
    }

    for c in fila:
        r["hoje"].append({
            "posicao": c["posicao"],
            "cod": c["cliente"]["cod"],
            "valor_vencido_centavos": c["valor_vencido_centavos"],
            "qtd_vencidos": c["qtd_vencidos"],
            "dias_atraso_max": c["dias_atraso_max"],
            "rank_medio": round(c["rank_medio"], 6),
            "contatado_recente": c["contatado_recente"],
            "explicacao": central.gerar_explicacao(c, cod_maior),
        })

    for cod in sorted(perfis):
        p = perfis[cod]
        r["risco"].append({
            "cod": cod,
            "indice": p["indice"],
            "valor_em_risco_centavos": p["valor_em_risco_centavos"],
            "saldo_aberto_centavos": p["saldo_aberto_centavos"],
            "qtd_vencidos_hoje": p["qtd_vencidos_hoje"],
            "nivel": arredonda(p["nivel"]),
            "variabilidade": arredonda(p["variabilidade"]),
            "prazo_efetivo": arredonda(p.get("prazo_efetivo")),
            "atraso_recente": arredonda(p["atraso_recente"]),
            "atraso_antigo": arredonda(p["atraso_antigo"]),
            "piora_confirmada": p["piora_confirmada"],
            "selo_cronico_estavel": p["selo_cronico_estavel"],
            "selo_invisivel": p["selo_invisivel"],
            "sazonais": sorted(p["sazonais"].keys()),
            "componentes": [[c["chave"], round(c["pontos"], 4)] for c in p["componentes"]],
            "alerta": central.gerar_alerta_ficha(p),
            "atraso_esperado": arredonda(p["atraso_esperado"]),
            "base_atraso_esperado": p["base_atraso_esperado"],
        })

    # regua: o tom escolhido para cada cliente nao pode mudar sozinho
    r["regua"] = []
    for p in central.clientes_da_regua(perfis):
        tom, motivo = central.escolher_tom(p)
        ctx = central.montar_contexto_redacao(ref, p)
        textos = redacao.redigir(tom, ctx)
        r["regua"].append({
            "cod": p["cliente"]["cod"],
            "tom": tom,
            "motivo": motivo,
            "qtd_titulos_no_texto": ctx["qtd_titulos"],
            "total_no_texto": ctx["total"],
            # o tamanho pega reescrita sem querer; o conteudo em si e revisado a olho
            "tamanhos": {k: len(v) for k, v in sorted(textos.items())},
        })

    r["semana"] = {
        k: v for k, v in central.calcular_semana(ref, titulos).items()
        if k not in ("inicio", "fim")
    }

    # o que vai para o alto da aba Hoje: qual achado, de quem, e a frase
    r["descobertas"] = [
        {"tipo": d["tipo"], "cods": d["cods"], "valor_em_risco_centavos": d["valor_em_risco_centavos"],
         "frase": d["frase"]}
        for d in central.calcular_descobertas(perfis)
    ]

    r["previsao"] = {
        "planilha": prev["planilha"],
        "esperado": prev["esperado"],
        "conservador": prev["conservador"],
        "otimista": prev["otimista"],
        "total_planilha": prev["total_planilha"],
        "total_esperado": prev["total_esperado"],
        "diferenca": prev["diferenca"],
        "fora_da_janela": prev["fora_da_janela"],
        "descontado": prev["descontado"],
        "semanas_invertidas": prev["semanas_invertidas"],
        "titulos_projetados": len(prev["linhas"]),
    }
    return r


def arredonda(v):
    return round(v, 6) if isinstance(v, (int, float)) else v


def conferir_html():
    """Verificacoes estruturais na pagina gerada."""
    caminho = central.ARQUIVO_SAIDA
    if not os.path.exists(caminho):
        return ["pagina central.html nao foi gerada ainda"]

    d = open(caminho, encoding="utf-8").read()
    falhas = []

    def exige(condicao, descricao):
        if not condicao:
            falhas.append(descricao)

    exige('charset="UTF-8"' in d, "faltou declarar charset UTF-8")
    exige(not re.search(r">None<|>NaN<|>nan<|>null<", d), "valor None/NaN vazou para a tela")
    exige("undefined" not in d, "valor undefined vazou para a tela")

    estilo = re.search(r"<style>(.*?)</style>", d, re.S)
    script = re.search(r"<script>(.*?)</script>", d, re.S)
    exige(estilo is not None, "faltou o bloco <style>")
    exige(script is not None, "faltou o bloco <script>")
    if estilo:
        exige(estilo.group(1).count("{") == estilo.group(1).count("}"), "chaves desbalanceadas no CSS")
    if script:
        exige(script.group(1).count("{") == script.group(1).count("}"), "chaves desbalanceadas no JavaScript")

    ref, clientes, titulos, cobrancas = central.carregar_dados()
    perfis = central.calcular_perfis_risco(ref, clientes, titulos, cobrancas)
    fila = central.calcular_fila(ref, clientes, titulos, cobrancas, perfis)

    linhas_hoje = len(re.findall(r'class="linha(?: [^"]*)?"', d))
    exige(linhas_hoje == len(fila), f"tela Hoje: {linhas_hoje} linhas na pagina, {len(fila)} na conta")

    linhas_risco = len(re.findall(r'class="risco-linha tier-', d))
    exige(linhas_risco == len(perfis), f"tela Risco: {linhas_risco} linhas na pagina, {len(perfis)} na conta")

    fichas = d.count("ficha ficha-modal")
    exige(fichas == len(perfis), f"fichas: {fichas} na pagina, {len(perfis)} clientes")

    # toda linha clicavel precisa da ficha correspondente
    for cod in perfis:
        exige(f'id="ficha-{cod}"' in d, f"faltou a ficha do cliente {cod}")
        exige(f"abrirFicha({cod})" in d, f"faltou o clique que abre a ficha do cliente {cod}")

    # cada aba precisa da tela e do botao
    for nome in ("hoje", "risco", "previsao", "regua", "relatorio"):
        exige(f'id="tela-{nome}"' in d, f"faltou a tela {nome}")
        exige(f'id="aba-{nome}"' in d, f"faltou a aba {nome}")

    # --- invariantes da previsao de caixa ---
    prev = central.calcular_previsao(ref, perfis)

    # no acumulado, quem paga mais cedo nunca pode ficar atras de quem paga mais tarde.
    # se isso inverter, a faixa da tela esta mentindo.
    for i, (o, cons) in enumerate(zip(prev["acum_otimista"], prev["acum_conservador"]), start=1):
        exige(o >= cons, f"faixa invertida no acumulado da semana {i}: otimista {o} < conservador {cons}")

    # a previsao so pode projetar titulos que existem em aberto
    abertos_qtd = len([t for t in titulos if t["dt_pagamento"] is None])
    exige(
        len(prev["linhas"]) <= abertos_qtd,
        f"previsao projetou {len(prev['linhas'])} titulos, mas so existem {abertos_qtd} em aberto",
    )

    # nada projetado pode passar do valor de face da carteira
    total_aberto = sum(t["valor_centavos"] for t in titulos if t["dt_pagamento"] is None)
    exige(
        prev["total_esperado"] <= total_aberto,
        f"esperado ({prev['total_esperado']}) maior que a carteira em aberto ({total_aberto})",
    )

    # o numero que domina a tela precisa estar escrito nela
    exige(
        central.formatar_reais_estimado(abs(prev["diferenca"])) in d,
        "a diferenca entre planilha e esperado nao aparece na pagina",
    )
    exige("não entra nesta conta" in d, "faltou o aviso de que venda nova nao entra na previsao")

    # --- regua de cobranca ---
    alvos = central.clientes_da_regua(perfis)
    exige(
        d.count('class="regua-cliente') == len(alvos),
        f"regua: {d.count('class=\"regua-cliente')} cartoes na pagina, {len(alvos)} clientes na conta",
    )
    for p in alvos:
        cod = p["cliente"]["cod"]
        exige(f'id="regua-{cod}"' in d, f"faltou o cartao de regua do cliente {cod}")
        exige(html.escape(p["cliente"]["telefone"]) in d, f"telefone do cliente {cod} nao aparece na regua")

    # todo cliente da regua precisa dos 6 tons x 3 canais ja escritos na pagina,
    # senao a troca de tom na mao nao teria o que mostrar
    dados = re.search(r'<script id="dados-regua" type="application/json">(.*?)</script>', d, re.S)
    exige(dados is not None, "faltaram os textos da regua na pagina")
    if dados:
        try:
            textos = json.loads(dados.group(1))
        except ValueError:
            textos = {}
            falhas.append("os textos da regua nao sao JSON valido")
        for p in alvos:
            cod = str(p["cliente"]["cod"])
            exige(cod in textos, f"faltaram os textos do cliente {cod}")
            if cod in textos:
                # cada cliente tem escritos exatamente os tons que da para escrever
                # para ele: nem faltando (a troca na mao quebraria) nem sobrando
                # (texto citando numero que ele nao tem)
                possiveis = set(redacao.tons_possiveis(central.montar_contexto_redacao(ref, p)))
                exige(
                    set(textos[cod]) == possiveis,
                    f"cliente {cod}: tons escritos {sorted(textos[cod])} != tons possiveis {sorted(possiveis)}",
                )
                for tom, t in textos[cod].items():
                    tudo = t.get("whatsapp", "") + t.get("email", "") + t.get("ligacao", "")
                    exige(
                        not re.search(r"— dias|de — para|média de —|em média —", tudo),
                        f"cliente {cod}, tom {tom}: frase com numero faltando ('—' no lugar do numero)",
                    )
                    for canal in ("whatsapp", "email", "ligacao"):
                        exige(t.get(canal), f"cliente {cod}, tom {tom}: faltou o texto de {canal}")
                    exige(
                        redacao.ASSINATURA.split("\n")[0] in t.get("whatsapp", ""),
                        f"cliente {cod}, tom {tom}: WhatsApp sem assinatura",
                    )
                    # o "1" tem que estar sozinho: "51 dias" e "11 dias" estao certos
                    junto = t.get("whatsapp", "") + t.get("email", "") + t.get("ligacao", "")
                    exige(
                        re.search(r"(?<!\d)1 dias", junto) is None,
                        f"cliente {cod}, tom {tom}: concordancia errada em '1 dias'",
                    )
                    exige(
                        re.search(r"(?<!\d)1 (títulos|meses)", junto) is None,
                        f"cliente {cod}, tom {tom}: concordancia errada no plural",
                    )

    # --- registro de cobranca ---
    exige("cobrancas_registradas.csv" in d, "faltou o download dos registros em CSV")
    exige("text/csv" in d, "o download nao esta saindo como CSV")
    exige("localStorage" in d, "faltou a gravacao dos registros no navegador")
    exige('id="registros-arquivo"' in d and "todosRegistros()" in d,
          "o download nao junta o que ja esta no arquivo -- baixar de outro navegador apagaria registros")
    exige(
        f"REF_ISO = '{ref.isoformat()}'" in d,
        "o carimbo do registro nao esta preso a data de referencia da base",
    )
    exige("new Date()" not in d, "o JavaScript pega o relogio do computador em algum lugar")
    exige("aplicarRegistrosNaFila" in d, "a fila de Hoje nao reage aos registros")

    # --- relatorio da semana ---
    rel = re.search(r'<script id="texto-relatorio" type="text/plain">(.*?)</script>', d, re.S)
    exige(rel is not None, "faltou o texto do relatorio na pagina")
    if rel:
        texto = rel.group(1)
        for secao in ("O QUE ENTROU", "O QUE N", "ONDE EST", "QUANTO ENTRA NAS PR", "O QUE FAZER NESTA SEMANA"):
            exige(secao in texto, f"relatorio sem a secao que comeca em '{secao}'")
        exige(len(texto) > 1200, "relatorio curto demais para os quatro minutos de leitura")

    # --- nenhuma IA, nenhuma rede ---
    for proibido in ("openai", "anthropic", "fetch(", "XMLHttpRequest", "https://", "http://"):
        exige(proibido not in d, f"a pagina chama rede ou servico externo: {proibido}")

    # os totais da carteira precisam aparecer escritos na pagina
    abertos = [t for t in titulos if t["dt_pagamento"] is None]
    vencidos = [t for t in abertos if t["dt_vencimento"] < ref]
    for centavos, rotulo in [
        (sum(t["valor_centavos"] for t in vencidos), "total vencido"),
        (sum(t["valor_centavos"] for t in abertos), "total em aberto"),
    ]:
        exige(central.formatar_reais_de_centavos(centavos) in d, f"{rotulo} nao aparece na pagina")
    exige("PARA_ENVIAR = false" in d, "o central.html saiu marcado como copia enviada")

    # --- o que nao pode sumir de cada tela (o_que_nao_pode_sumir.md)
    falhas += conferir_inventario(d)

    # --- a copia para mandar para alguem da equipe
    if not os.path.exists(central.ARQUIVO_ENVIO):
        falhas.append("a copia para enviar (para_enviar/Central Aurora.html) nao foi gerada")
    else:
        c = open(central.ARQUIVO_ENVIO, encoding="utf-8").read()
        for proibido in ("openai", "anthropic", "fetch(", "XMLHttpRequest", "https://", "http://", "new Date()", "undefined"):
            exige(proibido not in c, f"copia enviada: aparece '{proibido}'")
        exige("PARA_ENVIAR = true" in c, "copia enviada: nao esta marcada como copia")
        exige("aurora_cobrancas_copia_enviada" in c,
              "copia enviada: guarda os registros na mesma chave do original e se misturaria com ele")
        exige(f"REF_ISO = '{ref.isoformat()}'" in c, "copia enviada: o carimbo nao esta preso a data de referencia")
        falhas += conferir_inventario(c, copia=True)

    return falhas


def secoes_da_pagina(d):
    """Corta a pagina nas cinco telas e no painel das fichas. Se a marcacao mudar e
    a tela nao for achada, a verificacao reprova avisando -- nao passa em branco."""
    marcas = [(m.start(), m.group(1)) for m in re.finditer(r'<section class="tela[^"]*" id="tela-(\w+)"', d)]
    secoes = {}
    for i, (ini, nome) in enumerate(marcas):
        fim = marcas[i + 1][0] if i + 1 < len(marcas) else d.find('<footer class="rodape"', ini)
        secoes[nome] = d[ini:fim]
    a, b = d.find('<div class="modal-fundo"'), d.find('<div class="aviso-flutuante"')
    fichas = d[a:b] if a >= 0 and b > a else ""
    return secoes, fichas


def conferir_inventario(d, copia=False):
    """A lista do o_que_nao_pode_sumir.md, verificada contra a pagina gerada.
    Encurtar, juntar ou guardar atras de um clique pode; sumir nao. Cada item daqui
    sai da mesma conta que a tela usa, entao vale para a planilha de qualquer mes."""
    E = html.escape
    P = central.prosa
    falhas = []
    onde = "copia enviada: " if copia else ""
    # o navegador mostra igual uma frase com ou sem quebra de linha no meio do HTML,
    # entao a comparacao tambem ignora isso. E o que a TELA diz se le sem os <script>.
    d = re.sub(r"\s+", " ", d)
    tela_visivel = re.sub(r"<script.*?</script>", "", d, flags=re.S)

    def exige(condicao, descricao):
        if not condicao:
            falhas.append(onde + descricao)

    ref, clientes, titulos, cobrancas = central.carregar_dados()
    perfis = central.calcular_perfis_risco(ref, clientes, titulos, cobrancas)
    fila = central.calcular_fila(ref, clientes, titulos, cobrancas, perfis)
    prev = central.calcular_previsao(ref, perfis)
    semana = central.calcular_semana(ref, titulos)
    descobertas = central.calcular_descobertas(perfis)
    alvos = central.clientes_da_regua(perfis)
    abertos = [t for t in titulos if t["dt_pagamento"] is None]
    vencidos = [t for t in abertos if t["dt_vencimento"] < ref]
    r = central.formatar_reais_de_centavos
    e = central.formatar_reais_estimado

    secoes, fichas = secoes_da_pagina(d)
    for nome in ("hoje", "risco", "previsao", "regua", "relatorio"):
        exige(nome in secoes, f"nao achei a tela {nome} na pagina -- a marcacao mudou?")
    exige(fichas, "nao achei o painel das fichas -- a marcacao mudou?")
    if falhas:
        return falhas
    hoje, risco, previsao, regua, relatorio = (secoes[n] for n in ("hoje", "risco", "previsao", "regua", "relatorio"))

    # --- em todas as telas
    exige("Central de Crédito e Cobrança" in d and "Distribuidora Aurora" in d, "sumiu o nome da Central")
    exige(f"{central.dia_semana_pt(ref)}, {central.formatar_data_pt(ref)}".lower() in d.lower(),
          "sumiu a data de referencia por extenso")
    exige("Distribuidora Aurora" in d, "sumiu o nome da empresa do menu")
    exige("não do relógio do computador" in d, "sumiu o aviso de que a data vem da planilha, nao do relogio")
    emissoes = [t["dt_emissao"] for t in titulos]
    exige(central.formatar_data_curta(min(emissoes)) in d and central.formatar_data_curta(max(emissoes)) in d,
          "sumiu o periodo do historico")
    for chave in ("hoje", "risco", "previsao", "regua", "relatorio"):
        exige(f'href="#tela-{chave}"' in d, f"o menu perdeu o caminho para a tela {chave}")
    exige("somente leitura" in d and "Nenhum valor desta página foi digitado à mão" in d
          and f"{len(titulos)} títulos da base" in d, "sumiu a nota de origem dos numeros")
    if copia:
        exige("cópia atualizada" in d, "a copia nao diz que os numeros nao se atualizam nela")
        exige("python central.py" not in tela_visivel, "a copia manda rodar python, que quem recebe nao tem")
    else:
        exige("python central.py" in d, "sumiu como atualizar a Central")

    # --- numeros da carteira, no alto da Hoje
    for rotulo, valor in (("vencido", r(sum(t["valor_centavos"] for t in vencidos))),
                          ("a vencer", r(sum(t["valor_centavos"] for t in abertos if t["dt_vencimento"] >= ref))),
                          ("em aberto", r(sum(t["valor_centavos"] for t in abertos))),
                          ("clientes", str(len(clientes)))):
        exige(f'data-total="{rotulo}">{valor}<' in hoje, f"aba Hoje: sumiu o total '{rotulo}' ({valor})")
    exige(f"{len(vencidos)} títulos</strong> de <strong>{len(fila)} clientes" in hoje,
          "aba Hoje: sumiu quantos titulos e clientes estao vencidos")
    exige(f"{len(abertos) - len(vencidos)} títulos ainda no prazo" in hoje, "aba Hoje: sumiu quantos titulos estao a vencer")
    exige(f"saldo total em aberto · {len(abertos)} títulos" in hoje, "aba Hoje: sumiu quantos titulos estao em aberto")
    exige(f"{len(titulos)} títulos nos últimos {central.meses_de_historico(titulos, ref)} meses" in hoje,
          "aba Hoje: sumiu o tamanho do historico")

    # --- Hoje
    exige("desce para o fim da fila" in hoje, "aba Hoje: sumiu a regra de quem foi cobrado descer na fila")
    exige(f"Os outros {len(clientes) - len(fila)} não têm nada" in hoje and "irPara(event, 'risco')" in hoje,
          "aba Hoje: sumiu quantos ficaram fora da fila e o caminho para a aba Risco")
    for desc in descobertas:
        exige(P(desc["frase"]) in hoje, f"aba Hoje: sumiu a descoberta '{desc['tipo']}'")
    maior = max(fila, key=lambda c: c["valor_vencido_centavos"])["cliente"]["cod"] if fila else None
    exige(hoje.count('class="regua-atraso"') == len(fila), "aba Hoje: falta a regua de atraso em algum cartao")
    com_historico = sum(1 for c in fila if central.historico_para_desenhar(perfis[c["cliente"]["cod"]]))
    linhas_na_pagina = hoje.count('class="linha-atraso')
    exige(linhas_na_pagina == com_historico,
          f"aba Hoje: {com_historico} cartoes tem historico e deviam mostrar o atraso mes a mes; "
          f"a pagina mostra {linhas_na_pagina}")
    for c in fila:
        cl = c["cliente"]
        cod = cl["cod"]
        m = re.search(rf'<article class="linha[^"]*" data-cod="{cod}".*?</article>', hoje, re.S)
        exige(m, f"aba Hoje: sumiu o cartao do cliente {cod}")
        if not m:
            continue
        k = m.group(0)
        qtd = c["qtd_vencidos"]
        itens = [
            (E(cl["nome"]), "o nome"), (E(cl["categoria"]), "a categoria"), (E(cl["regiao"]), "a regiao"),
            (E(cl["contato"]), "o contato"), (E(cl["telefone"]), "o telefone"),
            (f'{r(c["valor_vencido_centavos"])}<small>vencido · {qtd} ', "o valor vencido e a quantidade"),
            (P(central.gerar_explicacao(c, maior)), "a frase que explica o cliente"),
            ("Posição por valor vencido", "a posicao por valor"), ("Posição por dias em atraso", "a posicao por dias"),
            ("Posição por risco comportamental", "a posicao por risco"),
            (f"{central.texto_ordinal(c['rank_valor'])} de {len(fila)}", "o numero da posicao por valor"),
            ("Atraso médio histórico (títulos pagos)", "o atraso medio historico"),
            (f"com base em {c['qtd_titulos_pagos_historico']} ", "em quantos titulos a media se baseia"),
            ("Tendência recente do atraso", "a tendencia"), ("Prazo de pagamento contratado", "o prazo contratado"),
            ("Última compra", "a ultima compra"),
            ("último contato" if c["ultimo_contato"] else "nunca foi cobrado", "o ultimo contato"),
            ("Registrei a ligação", "o botao de registrar a ligacao"), (f"abrirFicha({cod})", "o link da ficha"),
            (f"irParaRegua({cod})", "o link da mensagem pronta"),
        ]
        if c["contatado_recente"]:
            itens.append(("Aguardando retorno do contato", "o selo de aguardando retorno"))
        pf = perfis[cod]
        if central.historico_para_desenhar(pf):
            (ano_fim, mes_fim), ultimo = list(pf["serie_mensal"].items())[-1]
            itens.append((f"{ultimo:.0f} d · {central.MESES_PT_INV[mes_fim][:3]}/{ano_fim % 100:02d}</text>",
                          "o ultimo valor do atraso mes a mes, com o mes"))
        for t in c["titulos_vencidos"]:
            itens.append((E(t["numero"]), f"o titulo {t['numero']}"))
            itens.append((r(t["valor_centavos"]), f"o valor do titulo {t['numero']}"))
        for trecho, oque in itens:
            exige(trecho in k, f"aba Hoje, cliente {cod}: sumiu {oque}")

    # --- Risco
    total_em_risco = sum(p["valor_em_risco_centavos"] for p in perfis.values())
    total_saldo = sum(p["saldo_aberto_centavos"] for p in perfis.values())
    exige(e(total_em_risco) in risco and r(total_saldo) in risco
          and central.texto_pct(total_em_risco / total_saldo, 0) in risco,
          "aba Risco: sumiu o total em risco, o saldo ou a porcentagem da carteira")
    for faixa in ("0–24 risco baixo", "25–49 moderado", "50–74 alto", "75–100 crítico"):
        exige(faixa in risco, f"aba Risco: sumiu a legenda '{faixa}'")
    # o desenho de quem deve muito e quem esta piorando, logo abaixo do numero grande
    exige("dispersao-risco" in risco, "aba Risco: sumiu o grafico de quem deve muito e quem esta piorando")
    # no grafico, quem piorou e quem esta fora de qualquer lista tem o nome escrito (sao o motivo do desenho)
    dispersao = re.search(r'<svg class="dispersao-risco".*?</svg>', risco, re.S)
    rotulos = re.findall(r'<text class="dp-rotulo[^"]*"[^>]*>([^<]*)</text>', dispersao.group(0)) if dispersao else []
    for p in perfis.values():
        if (p["piora_confirmada"] or p["selo_invisivel"]) and p["diff_mudanca"] is not None:
            exige(any(rot.startswith(P(p["cliente"]["nome"])) for rot in rotulos),
                  f"aba Risco: o grafico nao escreve o nome de {p['cliente']['nome']}, que piorou ou esta fora de qualquer lista")
    for cod, p in perfis.items():
        situacao = "nada vencido hoje" if not p["qtd_vencidos_hoje"] else "vencido"
        # o cartao do cliente
        m = re.search(rf'<button type="button" class="risco-linha[^"]*" onclick="abrirFicha\({cod}\)".*?</button>', risco, re.S)
        exige(m, f"aba Risco: sumiu o cartao do cliente {cod} (a marcacao mudou?)")
        if m:
            k = m.group(0)
            for trecho, oque in ((E(p["cliente"]["nome"]), "o nome"), (E(p["cliente"]["categoria"]), "a categoria"),
                                 (E(p["cliente"]["regiao"]), "a regiao"), (e(p["valor_em_risco_centavos"]), "o valor em risco"),
                                 (r(p["saldo_aberto_centavos"]), "o saldo em aberto"), (f"{p['indice']:.0f}/100", "o indice"),
                                 (P(central.partes_do_alerta(p)[0]), "a primeira frase do alerta"),
                                 (situacao, "a situacao"), ("Abrir ficha", "o acesso a ficha")):
                exige(trecho in k, f"aba Risco, cartao do cliente {cod}: sumiu {oque}")
            if p["selo_cronico_estavel"]:
                exige("padrão dele, não é risco" in k, f"aba Risco, cartao do cliente {cod}: sumiu o selo 'padrao dele'")
            if p["selo_invisivel"]:
                exige("nunca apareceu em lista nenhuma" in k, f"aba Risco, cartao do cliente {cod}: sumiu o selo 'nunca apareceu'")
        # a linha dele na tabela, embaixo dos cartoes
        m = re.search(rf'<tr class="tr-risco[^"]*" onclick="abrirFicha\({cod}\)".*?</tr>', risco, re.S)
        exige(m, f"aba Risco: sumiu a linha do cliente {cod} na tabela (a marcacao mudou?)")
        if m:
            k = m.group(0)
            for trecho, oque in ((E(p["cliente"]["nome"]), "o nome"), (e(p["valor_em_risco_centavos"]), "o valor em risco"),
                                 (r(p["saldo_aberto_centavos"]), "o saldo em aberto"), (f"<b>{p['indice']:.0f}</b>", "o indice"),
                                 (central.faixa_de_risco(p["indice"])[1][6:], "a faixa do indice, escrita"),
                                 (situacao, "a situacao")):
                exige(trecho in k, f"aba Risco, tabela, cliente {cod}: sumiu {oque}")
            if p["selo_cronico_estavel"]:
                exige("padrão dele, não é risco" in k, f"aba Risco, tabela, cliente {cod}: sumiu o selo 'padrao dele'")
            if p["selo_invisivel"]:
                exige("nunca apareceu em lista nenhuma" in k, f"aba Risco, tabela, cliente {cod}: sumiu o selo 'nunca apareceu'")
    # o pe da tabela: a carteira inteira
    pe = re.search(r'<table class="tabela-risco">.*?<tfoot>(.*?)</tfoot>', risco, re.S)
    exige(pe and r(total_saldo) in pe.group(1) and e(total_em_risco) in pe.group(1),
          "aba Risco: o pe da tabela nao tem o saldo e o valor em risco da carteira inteira")

    # --- fichas
    rotulos_fatos = ("Prazo contratado", "Prazo real (emissão → pagamento)", "Atraso médio depois do vencimento",
                     "Variação do atraso (mín a máx)", "Desvio-padrão do atraso", "Saldo em aberto", "Limite de crédito",
                     "Ocupação do limite", "Faturamento 12 meses", "Fatia do faturamento", "Último pedido",
                     "Intervalo entre pedidos", "Cliente desde")
    for cod, p in perfis.items():
        m = re.search(rf'<div class="ficha ficha-modal" id="ficha-{cod}".*?(?=<div class="ficha ficha-modal"|$)', fichas, re.S)
        exige(m, f"ficha {cod}: sumiu")
        if not m:
            continue
        k = m.group(0)
        for rot in rotulos_fatos:
            exige(E(rot) in k, f"ficha {cod}: sumiu '{rot}'")
        for comp in p["componentes"]:
            exige(E(comp["nome"]) in k and P(comp["frase"]) in k, f"ficha {cod}: sumiu o peso '{comp['nome']}'")
        exige(P(central.gerar_alerta_ficha(p)) in k, f"ficha {cod}: sumiu o alerta")
        exige('grafico-ficha"' in k or "Histórico mensal insuficiente" in k, f"ficha {cod}: sumiu o grafico")
        exige("mês sazonal (descontado da comparação)" in k and "faixa do padrão histórico" in k,
              f"ficha {cod}: sumiu a legenda do grafico")
        for rot in ("Valor em risco", "Saldo em aberto", "Vencido hoje"):
            exige(rot in k, f"ficha {cod}: sumiu '{rot}' do alto")
        exige(E(p["cliente"]["telefone"]) in k, f"ficha {cod}: sumiu o telefone")
        for t in p["titulos_abertos"]:
            exige(E(t["numero"]) in k, f"ficha {cod}: sumiu o titulo {t['numero']}")
        if p["sazonais"]:
            exige("Padrão de calendário identificado" in k, f"ficha {cod}: sumiu o padrao de calendario")

    # --- Previsao
    dif = prev["diferenca"]
    for trecho, oque in (
        (central.formatar_data_curta(ref), "o inicio do periodo"), (central.formatar_data_curta(prev["fim_janela"]), "o fim do periodo"),
        (e(abs(dif)), "o numero principal"), (e(prev["total_planilha"]), "o total da planilha"),
        (e(prev["total_esperado"]), "o total esperado"), ("não entra nesta conta", "o aviso de venda nova"),
        (f"os {len(prev['linhas']) + len(prev['sem_historico'])} títulos hoje em aberto", "quantos titulos a previsao cobre"),
        ("com o que já venceu lançado na semana 1", "como a planilha e somada"),
        (r(semana["recebido_centavos"]), "quanto entrou na semana"), (r(semana["venceu_centavos"]), "quanto venceu na semana"),
        (r(semana["nao_entrou_centavos"]), "quanto nao foi pago na semana"),
        ('grafico-previsao"', "o grafico do acumulado"), ("o que esperamos receber", "a legenda do grafico"),
        ("faixa entre conservador e otimista", "a legenda da faixa"), ("Total nas 4 semanas", "o total da tabela semanal"),
        (e(prev["fora_da_janela"]), "o que entra depois das 4 semanas"), (e(prev["descontado"]), "o descontado pela chance"),
        ("premissa, não medida", "o aviso de que a chance e premissa"),
    ):
        exige(trecho in previsao, f"aba Previsao: sumiu {oque}")
    if prev["semanas_invertidas"]:
        exige("Por que a faixa de cenários só aparece" in previsao, "aba Previsao: sumiu a nota da faixa")
    for i in range(len(prev["semanas"])):
        exige(f"Semana {i + 1}</span>" in previsao, f"aba Previsao: sumiu a semana {i + 1} da tabela")
    for l in prev["linhas"]:
        exige(E(l["titulo"]["numero"]) in previsao, f"aba Previsao: sumiu o titulo {l['titulo']['numero']}")
    if prev["sem_historico"]:
        exige(r(prev["sem_historico_centavos"]) in previsao, "aba Previsao: sumiu o aviso dos titulos fora da previsao")

    # --- Regua
    for trecho, oque in ((f"{len(alvos)}<span", "quantas conversas"), ("quem envia é você", "quem envia"),
                         ("não há integração", "o aviso de nenhuma integracao"), ("Cobranças registradas", "o painel de registros"),
                         ("Limpar os que eu registrei", "o botao de limpar"), ("Base_bruta.xlsx não é tocada", "o aviso da planilha"),
                         (f"<strong>{central.formatar_data_curta(ref)}</strong>, a data de referência da base",
                          "o carimbo pela data de referencia"),
                         ("Como a régua escolhe o tom", "o metodo da regua")):
        exige(trecho in regua, f"aba Regua: sumiu {oque}")
    if copia:
        exige("Baixar registros (.csv)" not in regua, "aba Regua: o botao de baixar o arquivo aparece na copia, onde nao funciona")
        exige("Esta é uma cópia enviada" in regua and "Avise quem te mandou" in regua,
              "aba Regua: a copia nao explica por que o arquivo de registros esta desligado")
    else:
        exige("Baixar registros (.csv)" in regua, "aba Regua: sumiu o botao de baixar os registros")
    # nenhuma das duas paginas mostra o caminho da pasta do dono: a pagina pode ser aberta em outro
    # computador. Onde salvar o arquivo de registros, e se ele foi lido, o central.py diz no terminal.
    exige(E(central.PASTA) not in d, "a pagina mostra o caminho da pasta do dono (isso agora so sai no terminal)")
    for p in alvos:
        cod = p["cliente"]["cod"]
        m = re.search(rf'<div class="regua-cliente[^"]*" id="regua-{cod}".*?(?=<div class="regua-cliente|</div>\s*</div>\s*<p class="lista-vazia")', regua, re.S)
        exige(m, f"aba Regua: sumiu o texto do cliente {cod}")
        if not m:
            continue
        k = m.group(0)
        tom, motivo = central.escolher_tom(p)
        ctx = central.montar_contexto_redacao(ref, p)
        itens = [(E(p["cliente"]["nome"]), "o nome"), (E(p["cliente"]["telefone"]), "o telefone"),
                 (E(p["cliente"]["contato"]), "o contato"), (P(motivo), "o motivo do tom"),
                 ("WhatsApp</button>", "o canal WhatsApp"), ("E-mail</button>", "o canal e-mail"),
                 ("Roteiro de ligação</button>", "o roteiro de ligacao"), ("Copiar texto", "o botao de copiar"),
                 ("Registrar que cobrei", "o botao de registrar"), ("tom trocado na mão", "o aviso de tom trocado")]
        for t in redacao.tons_possiveis(ctx):
            itens.append((f'<option value="{t}"', f"o tom {t} no seletor"))
        for trecho, oque in itens:
            exige(trecho in k, f"aba Regua, cliente {cod}: sumiu {oque}")
        exige(f'data-cod="{cod}"' in regua and "regua-item" in regua, f"aba Regua: o cliente {cod} sumiu da lista")

    # --- Relatorio
    texto = central.gerar_relatorio_texto(
        ref, clientes, titulos, fila, perfis, prev, semana, [c for c in cobrancas if c.get("registrado_na_tela")])
    for secao in central.secoes_do_relatorio(texto):
        exige(E(secao) in relatorio, f"Relatorio: sumiu a secao '{secao}'")
    for c in fila:
        exige(E(c["cliente"]["nome"]) in relatorio and E(c["cliente"]["telefone"]) in relatorio,
              f"Relatorio: sumiu o cliente {c['cliente']['cod']} da fila de contato")
    doc = re.search(r'<article class="relatorio-doc".*?</article>', relatorio, re.S)
    exige(doc, "Relatorio: nao achei o texto formatado")
    if doc:
        valores = central.RE_DINHEIRO.findall(texto)
        pintados = doc.group(0).count('<span class="dinheiro">')
        exige(pintados == len(valores), f"Relatorio: {len(valores)} valores em dinheiro no texto, {pintados} na cor de dinheiro")
    for trecho in ("Copiar texto", "Imprimir / salvar em PDF", "Todos os números deste relatório"):
        exige(trecho in relatorio, f"Relatorio: sumiu '{trecho}'")
    for valor, oque in ((r(semana["recebido_centavos"]), "quanto entrou"), (r(semana["nao_entrou_centavos"]), "quanto nao foi pago"),
                        (r(sum(t["valor_centavos"] for t in vencidos)), "o vencido acumulado")):
        exige(f'<div class="kpi-valor">{valor}</div>' in relatorio, f"Relatorio: sumiu {oque} do alto")

    return falhas


def conferir_javascript():
    """Roda o teste da fila (teste_fila.js) se o Node estiver na maquina.
    Sem Node a Central funciona igual -- so esta conferencia extra nao roda."""
    import shutil
    import subprocess

    teste = os.path.join(PASTA, "teste_fila.js")
    if not os.path.exists(teste):
        return None, "teste_fila.js nao encontrado"
    if shutil.which("node") is None:
        return None, "Node nao instalado nesta maquina (a Central nao precisa dele)"
    try:
        r = subprocess.run(
            ["node", teste], cwd=PASTA, capture_output=True, text=True, timeout=60,
        )
    except (OSError, subprocess.SubprocessError) as erro:
        return None, f"nao consegui rodar o teste: {erro}"
    return r.returncode == 0, r.stdout.strip()


def comparar(base, agora, caminho=""):
    """Lista as diferencas entre dois retratos, com o caminho de cada uma."""
    difs = []
    if type(base) is not type(agora) and not (
        isinstance(base, (int, float)) and isinstance(agora, (int, float))
    ):
        return [f"{caminho}: tipo mudou de {type(base).__name__} para {type(agora).__name__}"]

    if isinstance(base, dict):
        for chave in sorted(set(base) | set(agora)):
            if chave not in base:
                difs.append(f"{caminho}.{chave}: surgiu (valor {agora[chave]!r})")
            elif chave not in agora:
                difs.append(f"{caminho}.{chave}: sumiu (era {base[chave]!r})")
            else:
                difs += comparar(base[chave], agora[chave], f"{caminho}.{chave}")
    elif isinstance(base, list):
        if len(base) != len(agora):
            difs.append(f"{caminho}: tinha {len(base)} itens, agora tem {len(agora)}")
        for i, (b, a) in enumerate(zip(base, agora)):
            difs += comparar(b, a, f"{caminho}[{i}]")
    elif base != agora:
        difs.append(f"{caminho}: era {base!r}, virou {agora!r}")
    return difs


def main():
    gravar = "--gravar-base" in sys.argv

    if gravar:
        with open(ARQUIVO_BASE, "w", encoding="utf-8") as f:
            json.dump(retrato(), f, ensure_ascii=False, indent=1)
        print(f"Linha de base gravada em {os.path.basename(ARQUIVO_BASE)}")
        return 0

    if not os.path.exists(ARQUIVO_BASE):
        print("Nao existe linha de base. Rode: python conferir.py --gravar-base")
        return 1

    base = json.load(open(ARQUIVO_BASE, encoding="utf-8"))
    agora = retrato()

    print("CONFERENCIA DA CENTRAL")
    print("=" * 62)

    difs = comparar(base, agora, "")
    # base trocada: os numeros TEM que mudar. Isso nao e defeito, e o mes novo.
    # A prova de que continua fechando passa a ser a auditoria/reconferir.py.
    base_trocada = base.get("data_referencia") != agora.get("data_referencia")
    if difs and base_trocada:
        print(f"\n[BASE NOVA] a linha de base e da planilha de {base.get('data_referencia')}, "
              f"e a planilha agora e de {agora.get('data_referencia')}.")
        print(f"     {len(difs)} numero(s) mudaram -- esperado, porque os dados mudaram.")
        print("     Esta parte so volta a valer depois de regravar a base (passo a passo no CLAUDE.md).")
    elif difs:
        print(f"\n[FALHOU] {len(difs)} numero(s) das telas mudaram com a MESMA planilha:\n")
        for d in difs[:40]:
            print("   -", d)
        if len(difs) > 40:
            print(f"   ... e mais {len(difs) - 40}")
    else:
        print("\n[OK] telas Hoje e Risco: todos os numeros identicos a linha de base")
        print(f"     ({len(agora['hoje'])} linhas na fila, {len(agora['risco'])} perfis de risco)")

    falhas = conferir_html()
    if falhas:
        print(f"\n[FALHOU] {len(falhas)} verificacao(oes) estrutural(is) da pagina:\n")
        for f in falhas:
            print("   -", f)
    else:
        print("\n[OK] central.html e a copia para enviar: estrutura integra e nada da lista")
        print("     o_que_nao_pode_sumir.md sumiu de nenhuma tela")

    passou_js, detalhe_js = conferir_javascript()
    if passou_js is None:
        print(f"\n[PULADO] teste da fila no navegador: {detalhe_js}")
    elif passou_js:
        print("\n[OK] fila de Hoje reage ao registro de cobranca (teste_fila.js)")
    else:
        print("\n[FALHOU] o teste da fila de Hoje nao passou:\n")
        for linha in detalhe_js.split("\n"):
            print("   ", linha)

    print("\n" + "=" * 62)
    if falhas or passou_js is False or (difs and not base_trocada):
        print("RESULTADO: conferencia reprovada -- conserte o codigo.")
        return 1
    if difs and base_trocada:
        print("RESULTADO: estrutura aprovada com a planilha nova.")
        print("Proximo passo: python auditoria/reconferir.py -- se ela fechar,")
        print("rode python conferir.py --gravar-base.")
        return 3
    print("RESULTADO: conferencia aprovada.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
