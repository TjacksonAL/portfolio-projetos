# O que não pode sumir da Central

Lista feita antes do redesenho de setembro de 2026, a partir do texto que cada tela mostrava
(com todas as dobras abertas). Serve para conferir, a cada rodada de mudança na aparência, que
nenhum alerta, link, frase ou número desapareceu.

Atualizada para a V9 (a V7 com três peças da V8): o atraso mês a mês nos cartões da fila, o
gráfico e a tabela da aba Risco e os três números do Relatório entraram na lista.

**Ela é conferida sozinha.** O `python conferir.py` verifica cada item daqui contra a página
gerada — no `central.html` e na cópia `para_enviar/Central Aurora.html` — pela função
`conferir_inventario`. Item novo nesta lista tem que entrar lá também, senão ninguém confere.

**Regra:** pode encurtar frase, juntar duas numa, mudar de lugar ou guardar atrás de um clique.
Não pode sumir. Se um item daqui deixar de existir na tela, é defeito, a não ser que o dono tenha
pedido para tirar.

Os valores entre parênteses são os da base de 12/08/2026, só de exemplo. Com a planilha do mês
seguinte eles mudam; o item continua obrigatório.

---

## Em todas as telas

- Nome: Central de Crédito e Cobrança, Distribuidora Aurora.
- A data de referência por extenso, com o dia da semana (quarta-feira, 12 de agosto de 2026).
- O aviso de que essa data vem da planilha (aba Leia-me), não do relógio do computador, e que por
  isso os números não mudam sozinhos amanhã.
- O período do histórico (03/09/2024 a 11/08/2026).
- O caminho para as outras telas, com a tela atual marcada, e o acesso ao Relatório da semana.
- Nota de origem: gerado da Base_bruta.xlsx em modo somente leitura; nenhum valor digitado à mão;
  quantos títulos a base tem (754); como atualizar.

## Números da carteira (no alto da aba Hoje)

- Vencido hoje, com quantos títulos e de quantos clientes (R$ 302.680,00 · 16 títulos · 7 clientes).
- A vencer, com quantos títulos (R$ 679.220,00 · 24).
- Saldo total em aberto, com quantos títulos (R$ 981.900,00 · 40).
- Clientes ativos e títulos nos últimos N meses (12 · 754 · 24 meses).

## Hoje

- Como a fila é montada: cruza valor vencido, risco comportamental e cobrança recente; quem foi
  cobrado nos últimos 2 dias desce para o fim.
- Quantos clientes estão na fila e quantos ficaram de fora por não ter nada vencido, com o
  caminho para a aba Risco, que mostra a carteira inteira.
- O que a Central descobriu: até três achados (piorou, fora de qualquer lista, padrão que não é
  risco...), cada um com a frase e o caminho para a ficha.
- Para cada cliente da fila (7):
  - posição na fila;
  - nome, categoria e região;
  - valor vencido e quantos títulos;
  - contato: nome e telefone;
  - último contato (há quantos dias · canal · tom) ou "nunca foi cobrado";
  - selo "aguardando retorno do contato" quando foi cobrado há até 2 dias;
  - a etiqueta da situação, com a palavra: fora do padrão, piorou, padrão dele, aguardando retorno
    do contato ou sem histórico;
  - **a frase que explica por que ele está ali**, curta (duas ou três linhas), com: dias do título
    mais antigo (o valor fica no cabeçalho do cartão); "maior valor vencido da fila" quando for o
    caso; se o atraso é o padrão dele, se está fora do padrão ou se piorou; "cobrado há N dias,
    desceu para o fim da fila" quando for o caso ("nunca foi cobrado" fica na linha do contato);
  - a régua pequena do atraso de hoje contra a faixa normal do cliente;
  - embaixo da régua, o atraso mês a mês dos títulos já pagos, em todo cartão com histórico, com o
    último valor e o mês dele ("18 d · abr/26");
  - o botão "Registrei a ligação" e os caminhos para a mensagem pronta e para a ficha;
  - os títulos vencidos, um a um (número, vencimento, valor, dias em atraso) e o total;
  - os números por trás da posição: posição por valor, por dias e por risco comportamental;
    atraso médio histórico e em quantos títulos pagos ele se baseia; tendência recente; prazo
    contratado; última compra.
- Depois de registrar uma cobrança: o aviso "você registrou esta cobrança, desceu para o fim da
  fila", com canal, tom e valor.

## Risco por cliente

- Total em risco, sobre quanto em aberto, e a porcentagem da carteira (R$ 318.494 de
  R$ 981.900,00 · 32%).
- Como se chega nele: saldo em aberto × índice de 0 a 100, sobre N meses de histórico; a ordem é
  por onde está o dinheiro, não por quem é o pior cliente.
- A legenda das faixas: 0–24 baixo, 25–49 moderado, 50–74 alto, 75–100 crítico.
- Logo abaixo do número grande, o gráfico "quem deve muito e quem está piorando": um ponto por
  cliente (saldo em aberto contra a mudança do atraso nos últimos 6 meses), com o nome escrito de
  quem piorou e de quem está fora de qualquer lista.
- Para cada cliente (12): posição; nome; selos "padrão dele, não é risco" e "nunca apareceu em
  lista nenhuma"; categoria, região e situação (títulos vencidos ou "nada vencido hoje"); valor em
  risco; saldo em aberto; índice /100; a primeira frase do alerta; o acesso à ficha.
- Embaixo dos cartões, a tabela da carteira inteira (12 linhas, sempre todas): posição; nome com os
  selos; situação; saldo em aberto; índice com a faixa escrita; valor em risco; clicar na linha
  abre a ficha; no pé, o saldo e o valor em risco da carteira inteira. No celular, a tabela mostra
  posição, nome, índice e valor em risco (situação e saldo estão no cartão e na ficha).
- **Ficha de cada cliente (12):**
  - nome, categoria, região/UF, cliente desde;
  - índice e faixa; valor em risco; saldo em aberto; quantos títulos vencidos;
  - contato, telefone e último contato;
  - o alerta inteiro, com a ação recomendada;
  - o gráfico "em quantos dias ele pagou, mês a mês", com a faixa do padrão dele, os pontos fora
    do padrão, os meses sazonais e a legenda;
  - os títulos em aberto, um a um, com o total;
  - os 13 números do comportamento (prazo contratado, prazo real, atraso médio, variação mín–máx,
    desvio-padrão, saldo, limite, ocupação do limite, faturamento 12 meses, fatia do faturamento,
    último pedido, intervalo entre pedidos, cliente desde) e o padrão de calendário quando houver;
  - os 6 componentes do índice, cada um com pontos, peso e a frase que justifica.

## Previsão de caixa

- O período das 4 semanas (12/08/2026 a 08/09/2026) e o método: vencimento mais o atraso que
  aquele cliente pratica.
- O número principal: o que a planilha promete e não entra (R$ 199.955), com a frase que o
  explica (planilha R$ 898.390, esperado R$ 698.435, 22% a menos).
- As três contas: a planilha apontaria (com "o que já venceu lançado na semana 1"), esperamos
  receber, diferença e a porcentagem.
- O aviso "isto não é projeção comercial": só a carteira existente (40 títulos), venda nova não
  entra nesta conta, estimativa sem centavo e título com centavo.
- O aviso de títulos fora da previsão por cliente sem histórico, com o valor, quando houver.
- A semana que acabou de passar: quanto entrou, em quantos títulos, com que atraso médio; quanto
  venceu; quanto não foi pago e em quantos títulos.
- O gráfico da entrada acumulada, com as linhas planilha e esperado, a faixa conservador–otimista
  e a legenda.
- A nota de por que a faixa só aparece no acumulado, com as semanas em que ela inverte.
- A tabela semana a semana (planilha, esperado, diferença) e o total.
- Quanto só entra depois das 4 semanas e quanto foi descontado pela chance de entrada.
- Como a chance de entrada é calculada (até 30% pelo risco, idade contra o maior atraso do
  cliente, piso de 25%) e o aviso "premissa, não medida".
- Os títulos projetados, um a um: título, cliente, valor, vencimento, entrada esperada e semana,
  chance, ponderado.

## Régua de cobrança

- Como o tom é escolhido, que dá para trocar na mão, e quantos clientes estão na régua (8: os 7
  com vencido, mais quem a aba Risco mandou procurar).
- "A Central redige e copia; quem envia é você." Nenhuma integração com WhatsApp, e-mail ou
  qualquer serviço.
- **Painel de cobranças registradas:**
  - para que serve e que quem foi cobrado há menos de 2 dias desce na fila de Hoje;
  - o carimbo é a data de referência, nunca o relógio; a Base_bruta.xlsx não é tocada;
  - a lista de cobranças registradas (cliente, quando, canal, tom, valor) com a origem de cada
    uma: "já no arquivo" ou "só neste navegador, baixe o arquivo";
  - os botões Baixar registros (.csv) e Limpar os que eu registrei;
  - (saíram da tela a pedido do dono: onde o arquivo tem que ficar e o estado dele — encontrado,
    não existe, nome errado. Isso agora sai no terminal, ao rodar `python central.py`, porque a
    página pode ser aberta em outro computador.)
- **Para cada cliente da régua (8):** nome; o tom sugerido; contato, telefone e situação; valor
  (vencido ou em aberto); "por que a Central sugeriu este tom" com o motivo; a troca de tom só com
  os tons possíveis, o aviso "tom trocado na mão" e o aviso dos tons que ficaram de fora; os três
  canais (WhatsApp, e-mail com assunto, roteiro de ligação); o texto; copiar texto e o aviso
  "copiado"; o canal do registro; registrar que cobrei.

## Relatório da semana

- Abaixo da data do alto, os três números da semana: quanto entrou (e em quantos títulos), quanto
  venceu e não foi pago (e em quantos títulos), e o vencido acumulado na carteira.
- As cinco partes: o que entrou, o que não entrou, onde está o risco, quanto entra nas próximas
  semanas, o que fazer nesta semana.
- A fila de contato por valor: nome, valor vencido, títulos, dias do mais antigo, abordagem,
  contato e telefone.
- As cobranças registradas na semana, quando houver.
- A nota final: números da Base_bruta.xlsx, data de referência congelada, os mesmos das telas.
- Copiar o texto e imprimir ou salvar em PDF.
- Todo valor em dinheiro no meio do texto na cor de dinheiro (a verificação conta um por um).

## Só na cópia para enviar

- O aviso de que é uma cópia, com a data dos dados, e de que os números não se atualizam nela.
- No lugar do botão de baixar o arquivo de registros e do caminho da pasta: a explicação de por
  que isso está desligado na cópia e o pedido para avisar quem mandou a Central.
- Nenhum caminho de pasta do computador de quem gerou a cópia (vale também para a Central).
