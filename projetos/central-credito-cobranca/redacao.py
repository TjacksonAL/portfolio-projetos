#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Redacao das mensagens de cobranca da Central
==============================================

ESTE E O UNICO LUGAR DO SISTEMA QUE ESCREVE TEXTO PARA O CLIENTE.

Nada aqui sabe ler planilha, calcular atraso ou decidir risco. Este modulo recebe
um dicionario ja pronto (o "contexto") e devolve texto. O resto da Central decide
O QUE dizer; aqui se decide COMO dizer.

E assim de proposito: no dia em que a redacao passar a ser feita por um modelo de
linguagem, basta trocar a funcao `redigir` por uma chamada ao modelo, mantendo a
mesma assinatura. Nenhum outro arquivo precisa mudar.

Hoje NAO existe nenhuma chamada de rede aqui: todo texto e montado a partir dos
numeros que a Central ja calculou. Roda numa maquina sem internet.

    redigir(tom, contexto) -> {"whatsapp": ..., "email_assunto": ...,
                               "email": ..., "ligacao": ...}
"""

ASSINATURA = "Equipe financeira\nDistribuidora Aurora"

# Os tons disponiveis. A escolha automatica de qual usar NAO e feita aqui --
# ela vem pronta no parametro `tom`, decidida pelos numeros do cliente.
TONS = {
    "lembrete_leve": {
        "nome": "Lembrete leve",
        "quando": "O atraso de hoje está dentro do padrão normal dele. Não é sinal de risco.",
        "cor": "verde",
        # numeros do cliente que o texto cita; sem eles o tom nao e oferecido
        "usa": [],
    },
    "primeiro_aviso_cordial": {
        "nome": "Primeiro aviso cordial",
        "quando": "Cliente que costuma pagar bem e desta vez atrasou muito mais do que o normal dele.",
        "cor": "azul",
        # numeros do cliente que o texto cita; sem eles o tom nao e oferecido
        "usa": ["atraso_medio"],
    },
    "conversa_mudanca": {
        "nome": "Conversa sobre a mudança",
        "quando": "O comportamento de pagamento piorou de forma confirmada. Pede conversa, não cobrança.",
        "cor": "ambar",
        # numeros do cliente que o texto cita; sem eles o tom nao e oferecido
        "usa": ["atraso_medio", "atraso_recente"],
    },
    "reaproximacao_comercial": {
        "nome": "Reaproximação comercial",
        "quando": "Parou de comprar. A conversa comercial vem antes da cobrança do financeiro.",
        "cor": "roxo",
        # numeros do cliente que o texto cita; sem eles o tom nao e oferecido
        "usa": ["dias_sem_comprar", "intervalo_medio"],
    },
    "cobranca_objetiva": {
        "nome": "Cobrança objetiva",
        "quando": "Atraso relevante que não se explica pelo padrão dele nem por mudança recente.",
        "cor": "vermelho",
        # numeros do cliente que o texto cita; sem eles o tom nao e oferecido
        "usa": [],
    },
    "contato_preventivo": {
        "nome": "Contato preventivo",
        "quando": "Não tem nada vencido, mas os sinais de pagamento e de compra pioraram.",
        "cor": "cinza",
        # numeros do cliente que o texto cita; sem eles o tom nao e oferecido
        "usa": ["dias_sem_comprar", "intervalo_medio", "atraso_medio", "atraso_recente"],
    },
}


NUMERO_POR_EXTENSO = {
    1: "um", 2: "dois", 3: "três", 4: "quatro", 5: "cinco",
    6: "seis", 7: "sete", 8: "oito", 9: "nove", 10: "dez",
}


def tons_possiveis(ctx):
    """Tons que da para escrever com honestidade para este cliente. Um tom que cita
    'a media de voces e de X dias' nao pode ser oferecido para quem nao tem media:
    o texto sairia 'a media de voces e de — dias'."""
    def tem(chave):
        return str(ctx.get(chave, "—")) not in ("—", "None", "")
    return [t for t, info in TONS.items() if all(tem(c) for c in info["usa"])]


def primeiro_nome(nome_completo):
    return (nome_completo or "").strip().split(" ")[0] if nome_completo else ""


def quantos(n, singular, plural):
    """'um título' / 'dois títulos'. Numero por extenso ate dez, que e como
    gente escreve -- '1 título' num WhatsApp denuncia texto de sistema."""
    palavra = NUMERO_POR_EXTENSO.get(n, f"{n}")
    return f"{palavra} {singular if n == 1 else plural}"


def valor_total(ctx):
    """'somando R$ X' faz sentido com varios titulos; com um so, soa como sistema."""
    return f"somando {ctx['total']}" if ctx["qtd_titulos"] > 1 else f"no valor de {ctx['total']}"


def dias_txt(n):
    return "1 dia" if abs(n) == 1 else f"{n} dias"


def lista_titulos(ctx, marcador="• "):
    """Os titulos um a um, do mais antigo para o mais novo."""
    linhas = []
    for t in ctx["titulos"]:
        if t["dias"] > 0:
            situacao = f"{dias_txt(t['dias'])} em atraso"
        elif t["dias"] == 0:
            situacao = "vence hoje"
        else:
            situacao = f"vence em {dias_txt(-t['dias'])}"
        linhas.append(f"{marcador}{t['numero']} — venceu em {t['vencimento']} — {t['valor']} ({situacao})")
    return "\n".join(linhas)


def _bloco_titulos(ctx):
    if not ctx["titulos"]:
        return ""
    cabecalho = "O título:" if ctx["qtd_titulos"] == 1 else "Os títulos:"
    return f"{cabecalho}\n{lista_titulos(ctx)}\n"


def redigir(tom, ctx):
    """Monta as tres versoes do texto para um cliente, num tom.

    ctx precisa trazer, ja formatado para leitura:
      contato, empresa, titulos[{numero, vencimento, valor, dias}], qtd_titulos,
      total, dias_mais_antigo, data_mais_antiga, atraso_medio, atraso_recente,
      prazo_contratado, prazo_real, dias_sem_comprar, intervalo_medio,
      participacao, sugerir_prazo (bool)
    """
    montador = _MONTADORES.get(tom, _MONTADORES["cobranca_objetiva"])
    texto = montador(ctx)
    texto["whatsapp"] = texto["whatsapp"].strip() + f"\n\n{ASSINATURA}"
    texto["email"] = texto["email"].strip() + f"\n\n{ASSINATURA}"
    return texto


# ---------------------------------------------------------------------------
# Um montador por tom. Tudo abaixo e redacao -- nenhuma regra de negocio.
# ---------------------------------------------------------------------------

def _lembrete_leve(ctx):
    nome = primeiro_nome(ctx["contato"])
    quantidade = quantos(ctx["qtd_titulos"], "título", "títulos")
    plural_t = "títulos" if ctx["qtd_titulos"] > 1 else "título"
    proposta_wpp = ""
    proposta_email = ""
    proposta_ligacao = ""
    if ctx.get("sugerir_prazo"):
        proposta_wpp = (
            f"\nE já que estou te escrevendo: o contrato de vocês está em {ctx['prazo_contratado']}, "
            f"mas na prática o pagamento sai em torno de {ctx['prazo_real']} dias. Se esse é o prazo que "
            f"funciona aí, a gente acerta isso no contrato e para de gerar cobrança à toa dos dois lados. "
            f"O que você acha?\n"
        )
        proposta_email = (
            f"\nAproveito para propor uma coisa: o contrato de vocês está em {ctx['prazo_contratado']}, "
            f"mas o pagamento costuma sair em torno de {ctx['prazo_real']} dias — e sempre sai. "
            f"Se esse for o prazo que funciona para a operação de vocês, podemos ajustar o contrato "
            f"para refletir isso. Vocês param de receber cobrança e a gente para de mandar.\n"
        )
        proposta_ligacao = (
            f"\n• Proposta: ajustar o prazo do contrato de {ctx['prazo_contratado']} para perto de "
            f"{ctx['prazo_real']} dias, que é o que ele já pratica"
        )

    qual_venceu = (
        f"venceu em {ctx['data_prosa']}" if ctx["qtd_titulos"] == 1
        else f"o mais antigo venceu em {ctx['data_prosa']}"
    )

    contexto_ligacao = (
        f"• Contexto: ele atrasa em média {ctx['atraso_medio']} dias e sempre paga — isso aqui é o normal dele"
        if ctx.get("tem_historico", True) else
        "• Contexto: cliente sem histórico suficiente — ainda não dá para dizer qual é o normal dele"
    )

    whatsapp = f"""Oi, {nome}, tudo bem?

Passando pra lembrar de {quantidade} que {'está' if ctx['qtd_titulos'] == 1 else 'estão'} em aberto aqui: {ctx['total']}{' no total' if ctx['qtd_titulos'] > 1 else ''}, {qual_venceu}.

{_bloco_titulos(ctx)}
Nada urgente, só queria confirmar com você se o pagamento já está programado.
{proposta_wpp}
Qualquer coisa é só me chamar por aqui."""

    email = f"""Olá, {nome}, tudo bem?

Estou passando para lembrar de {quantidade} de vocês que {'continua' if ctx['qtd_titulos'] == 1 else 'continuam'} em aberto, {valor_total(ctx)}. {qual_venceu.capitalize()}.

{_bloco_titulos(ctx)}
Não é nada urgente — só queria confirmar se o pagamento já está programado aí.
{proposta_email}
Se precisar que eu reenvie boleto ou nota, é só responder este e-mail."""

    ligacao = f"""Como abrir: lembrete, não cobrança. Ele está dentro do padrão dele.

• Cliente: {ctx['empresa']} · falar com {ctx['contato']}
• Em aberto: {quantidade}, {ctx['total']}
• Mais antigo: {ctx['data_mais_antiga']}, há {dias_txt(ctx['dias_mais_antigo'])}
{contexto_ligacao}{proposta_ligacao}

Se ele disser que já pagou: pedir o comprovante e encerrar, sem insistir.
Se pedir prazo: pode concordar, é o comportamento habitual dele."""

    return {"whatsapp": whatsapp, "email_assunto": f"Títulos em aberto — {ctx['empresa']}",
            "email": email, "ligacao": ligacao}


def _primeiro_aviso_cordial(ctx):
    nome = primeiro_nome(ctx["contato"])
    quantidade = quantos(ctx["qtd_titulos"], "título", "títulos")
    um_so = ctx["qtd_titulos"] == 1
    verbo = "apareceu" if um_so else "apareceram"
    qual_venceu = "vencido em" if um_so else "o mais antigo vencido em"

    whatsapp = f"""Oi, {nome}, tudo bem?

{verbo.capitalize()} {quantidade} de vocês em aberto aqui e me chamou atenção, porque não é o normal de vocês: {ctx['total']}, {qual_venceu} {ctx['data_prosa']} — faz {dias_txt(ctx['dias_mais_antigo'])}.

{_bloco_titulos(ctx)}
Vocês quase sempre pagam no prazo, então preferi perguntar antes de ficar insistindo: ficou preso em alguma coisa aí? Nota, aprovação, algum problema com o boleto?

Se precisar que eu reenvie qualquer coisa, me fala que eu mando agora."""

    email = f"""Olá, {nome}, tudo bem?

Estou escrevendo por causa de {quantidade} de vocês que {verbo} em aberto: {ctx['total']}, {qual_venceu} {ctx['data_prosa']}, há {dias_txt(ctx['dias_mais_antigo'])}.

{_bloco_titulos(ctx)}
O motivo de eu escrever é justamente por ser fora do comum. O histórico de vocês é de pagamento no prazo — a média de vocês é de {ctx['atraso_medio']} dias depois do vencimento, e desta vez passou bastante disso. Por isso prefiro perguntar se houve algum problema no caminho: nota fiscal, aprovação interna, boleto que não chegou.

Se for o caso, reenvio boleto e nota na hora. É só responder."""

    ligacao = f"""Como abrir: perguntar, não cobrar. Este cliente não costuma atrasar.

• Cliente: {ctx['empresa']} · falar com {ctx['contato']}
• Em aberto: {quantidade}, {ctx['total']}
• Mais antigo: {ctx['data_mais_antiga']}, há {dias_txt(ctx['dias_mais_antigo'])}
• Contexto: média histórica de {ctx['atraso_medio']} dias de atraso — este caso está muito fora do padrão dele

Perguntar nesta ordem:
1. Se recebeu a nota e o boleto
2. Se travou em alguma aprovação interna
3. Se é pontual ou se mudou alguma coisa no fluxo de caixa deles

Se ele já pagou: pedir comprovante.
Se travou em documento: reenviar ainda hoje e combinar a nova data."""

    return {"whatsapp": whatsapp, "email_assunto": f"Título em aberto — {ctx['empresa']}",
            "email": email, "ligacao": ligacao}


def _conversa_mudanca(ctx):
    nome = primeiro_nome(ctx["contato"])
    quantidade = quantos(ctx["qtd_titulos"], "título", "títulos")

    whatsapp = f"""Oi, {nome}, tudo bem?

Queria conversar rápido com você sobre os pagamentos — e não é cobrança, é conversa mesmo.

Vocês sempre pagaram numa média de {ctx['atraso_medio']} dias depois do vencimento. Nos últimos meses isso foi para {ctx['atraso_recente']} dias, e hoje tem {quantidade} em aberto, {ctx['total']}.

{_bloco_titulos(ctx)}
Não estou pedindo justificativa. Só quero entender se mudou alguma coisa aí — prazo de recebimento, fluxo de caixa, alguma coisa — pra gente se ajustar junto antes de virar um problema pros dois lados.

Você consegue falar comigo essa semana?"""

    email = f"""Olá, {nome}, tudo bem?

Escrevo para conversar sobre os pagamentos, e faço questão de dizer logo no começo que não é uma cobrança formal.

O histórico de vocês é bom: a média foi de {ctx['atraso_medio']} dias depois do vencimento ao longo do último período. Nos últimos meses essa média subiu para {ctx['atraso_recente']} dias, e hoje há {quantidade} em aberto, {valor_total(ctx)}.

{_bloco_titulos(ctx)}
Mudança de ritmo costuma ter explicação — prazo de recebimento que esticou, sazonalidade, uma negociação em andamento. Preferia entender junto com vocês do que ficar mandando aviso automático.

Podemos conversar esta semana? Me diga o melhor horário que eu ligo."""

    ligacao = f"""Como abrir: conversa, nunca ameaça. O histórico dele é bom — o que mudou foi o ritmo.

• Cliente: {ctx['empresa']} · falar com {ctx['contato']}
• Em aberto: {quantidade}, {ctx['total']}
• Mais antigo: {ctx['data_mais_antiga']}, há {dias_txt(ctx['dias_mais_antigo'])}
• O que mudou: média histórica de {ctx['atraso_medio']} dias, subiu para {ctx['atraso_recente']} nos últimos meses

Perguntar nesta ordem:
1. Se mudou alguma coisa no fluxo de recebimento deles
2. Se o prazo atual ainda funciona ou se ficou apertado
3. Se dá para combinar uma data firme para o que está em aberto

Não falar em: suspensão de crédito, protesto, negativação. Não é o momento.
Sair da ligação com: uma data combinada e o motivo da mudança entendido."""

    return {"whatsapp": whatsapp, "email_assunto": f"Uma conversa sobre os pagamentos — {ctx['empresa']}",
            "email": email, "ligacao": ligacao}


def _reaproximacao_comercial(ctx):
    nome = primeiro_nome(ctx["contato"])
    quantidade = quantos(ctx["qtd_titulos"], "título", "títulos")

    whatsapp = f"""Oi, {nome}, tudo bem?

Faz {dias_txt(ctx['dias_sem_comprar'])} que a gente não recebe pedido de vocês — antes era mais ou menos a cada {ctx['intervalo_medio']} dias. Fiquei com isso na cabeça e resolvi perguntar direto: está tudo certo aí?

Se teve algum problema com entrega, preço ou atendimento da nossa parte, eu prefiro saber. Dá pra resolver.

Tem {quantidade} em aberto também, {ctx['total']}:

{lista_titulos(ctx)}

Mas sinceramente, o que eu queria mesmo era entender o que aconteceu. Me chama quando puder."""

    email = f"""Olá, {nome}, tudo bem?

Escrevo menos pelo financeiro e mais pelo comercial: faz {dias_txt(ctx['dias_sem_comprar'])} que não recebemos um pedido de vocês, quando o intervalo habitual era de cerca de {ctx['intervalo_medio']} dias.

Antes de tratar isso como um número, prefiro perguntar: aconteceu alguma coisa? Problema de entrega, preço que ficou fora, atendimento, mudança de fornecedor. Se foi algo do nosso lado, quero saber para corrigir.

Sobre o financeiro, há {quantidade} em aberto, {valor_total(ctx)}:

{lista_titulos(ctx)}

Podemos tratar disso junto, mas o que me interessa primeiro é entender o afastamento. Fico no aguardo."""

    ligacao = f"""Como abrir: ligação comercial, não de cobrança. O financeiro entra depois, e só se ele abrir espaço.

• Cliente: {ctx['empresa']} · falar com {ctx['contato']}
• Sem comprar há: {dias_txt(ctx['dias_sem_comprar'])} (o normal dele é a cada {ctx['intervalo_medio']} dias)
• Em aberto: {quantidade}, {ctx['total']}
• Mais antigo: {ctx['data_mais_antiga']}, há {dias_txt(ctx['dias_mais_antigo'])}

Perguntar nesta ordem:
1. O que aconteceu para parar de comprar
2. Se teve problema de entrega, preço ou atendimento
3. Se trocou de fornecedor, e por quê

Só depois, e se a conversa permitir:
4. Combinar o que fazer com os títulos em aberto

Se ele estiver com problema de caixa: ouvir antes de propor. Pode ser a causa das duas coisas."""

    return {"whatsapp": whatsapp, "email_assunto": f"Faz um tempo que não recebo pedido de vocês — {ctx['empresa']}",
            "email": email, "ligacao": ligacao}


def _cobranca_objetiva(ctx):
    nome = primeiro_nome(ctx["contato"])
    quantidade = quantos(ctx["qtd_titulos"], "título", "títulos")
    um_so = ctx["qtd_titulos"] == 1
    plural_v = "vencido" if um_so else "vencidos"
    qual_venceu = "Ele venceu em" if um_so else "O mais antigo venceu em"

    contexto_ligacao = (
        f"• Contexto: ele atrasa em média {ctx['atraso_medio']} dias"
        if ctx.get("tem_historico", True) else
        "• Contexto: cliente sem histórico suficiente para saber o padrão dele"
    )

    whatsapp = f"""Oi, {nome}, tudo bem?

Tem {quantidade} de vocês {plural_v} aqui, {valor_total(ctx)}. {qual_venceu} {ctx['data_prosa']}, faz {dias_txt(ctx['dias_mais_antigo'])}.

{_bloco_titulos(ctx)}
Consegue me passar uma previsão de quando entra? Se já tiver pago, me manda o comprovante que eu baixo aqui hoje mesmo."""

    email = f"""Olá, {nome}, tudo bem?

Há {quantidade} de vocês {plural_v}, {valor_total(ctx)}. {qual_venceu} {ctx['data_prosa']}, há {dias_txt(ctx['dias_mais_antigo'])}.

{_bloco_titulos(ctx)}
Preciso de uma previsão de pagamento para atualizar aqui. Se já tiver sido pago, me envie o comprovante que eu dou baixa hoje mesmo.

Se houver alguma divergência na nota ou no boleto, me avise que eu verifico."""

    ligacao = f"""Como abrir: direto, respeitoso, com data como objetivo.

• Cliente: {ctx['empresa']} · falar com {ctx['contato']}
• {plural_v.capitalize()}: {quantidade}, {ctx['total']}
• Mais antigo: {ctx['data_mais_antiga']}, há {dias_txt(ctx['dias_mais_antigo'])}
{contexto_ligacao}

Objetivo da ligação: sair com uma data.

Se disser que já pagou: pedir comprovante na hora.
Se pedir prazo: aceitar data específica, não "semana que vem".
Se houver divergência de nota: anotar e resolver antes de cobrar de novo."""

    return {"whatsapp": whatsapp, "email_assunto": f"Títulos vencidos — {ctx['empresa']}",
            "email": email, "ligacao": ligacao}


def _contato_preventivo(ctx):
    nome = primeiro_nome(ctx["contato"])

    whatsapp = f"""Oi, {nome}, tudo bem?

Já começo dizendo: não é cobrança. Vocês não têm nada vencido aqui.

É que eu estava olhando a conta de vocês e queria confirmar uma coisa. O intervalo entre os pedidos aumentou um pouco (o último foi há {dias_txt(ctx['dias_sem_comprar'])}, normalmente é a cada {ctx['intervalo_medio']}) e o pagamento também esticou um pouquinho — de {ctx['atraso_medio']} para {ctx['atraso_recente']} dias em média.

Pode não ser nada. Mas vocês são um dos nossos maiores clientes, e eu preferi perguntar em vez de deixar passar.

Está tudo certo aí? Precisa de alguma coisa da nossa parte?"""

    email = f"""Olá, {nome}, tudo bem?

Começo pelo mais importante: este e-mail não é uma cobrança. Vocês não têm nenhum título vencido conosco.

Escrevo porque, olhando a conta de vocês, notei dois movimentos pequenos na mesma direção. O intervalo entre pedidos passou de cerca de {ctx['intervalo_medio']} dias para {dias_txt(ctx['dias_sem_comprar'])} desde o último, e o prazo médio de pagamento subiu de {ctx['atraso_medio']} para {ctx['atraso_recente']} dias.

Isoladamente, nenhum dos dois chama atenção. Juntos, achei que valia perguntar — ainda mais com uma conta do tamanho da de vocês.

Está tudo certo por aí? Se houver algo que a gente possa ajustar, de prazo a condição comercial, prefiro conversar agora do que depois."""

    ligacao = f"""Como abrir: contato de relacionamento. NÃO é cobrança — ele não tem nada vencido.

• Cliente: {ctx['empresa']} · falar com {ctx['contato']}
• Nada vencido. Saldo em aberto: {ctx['total']}
• Último pedido: há {dias_txt(ctx['dias_sem_comprar'])} (o normal é a cada {ctx['intervalo_medio']})
• Pagamento: subiu de {ctx['atraso_medio']} para {ctx['atraso_recente']} dias
• Peso: {ctx['participacao']} do faturamento dos últimos 12 meses

Perguntar nesta ordem:
1. Como está o movimento da loja
2. Se o prazo atual ainda atende
3. Se tem algo que a gente possa melhorar no atendimento ou na entrega

Não mencionar cobrança em nenhum momento.
Sair da ligação com: o motivo do afastamento entendido, se houver."""

    return {"whatsapp": whatsapp, "email_assunto": f"Passando para saber como estão as coisas — {ctx['empresa']}",
            "email": email, "ligacao": ligacao}


_MONTADORES = {
    "lembrete_leve": _lembrete_leve,
    "primeiro_aviso_cordial": _primeiro_aviso_cordial,
    "conversa_mudanca": _conversa_mudanca,
    "reaproximacao_comercial": _reaproximacao_comercial,
    "cobranca_objetiva": _cobranca_objetiva,
    "contato_preventivo": _contato_preventivo,
}
