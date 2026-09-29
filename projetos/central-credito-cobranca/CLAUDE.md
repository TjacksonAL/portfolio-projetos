# Central de Crédito e Cobrança — Distribuidora Aurora

Leia isto inteiro antes de mexer em qualquer arquivo. Ele existe para quem chega sem ter
acompanhado a construção — pessoa da equipe ou Claude numa conversa nova.

## O que é

Uma página HTML gerada por Python a partir da planilha `Base_bruta.xlsx` (carteira de contas
a receber de uma distribuidora de alimentos e bebidas que vende a prazo para o varejo).
Abre no navegador, sem servidor. Tem cinco telas, num menu fixo à esquerda:

| Aba | Pergunta que responde |
|---|---|
| **Hoje** | Quem eu ligo primeiro? Fila dos clientes com título vencido, com a frase que explica cada posição, o contato e os títulos um a um. |
| **Risco por cliente** | Onde está o meu dinheiro? Índice de risco 0–100 por cliente, ordenado por *valor em risco*. Clicar abre a ficha com gráfico e alerta. |
| **Previsão de caixa** | Quanto entra nas próximas 4 semanas lendo o comportamento de cada cliente, contra o que a soma dos vencimentos promete. |
| **Régua de cobrança** | O texto de cobrança pronto, no tom certo, para WhatsApp, e-mail e ligação; e o registro de quem foi cobrado. |
| **Relatório da semana** | Texto corrido para a diretoria, para copiar ou imprimir em PDF. Todo valor em dinheiro vem em outra cor. |

Além disso, gera `para_enviar/Central Aurora.html`: a mesma Central num arquivo só, para mandar
para alguém da equipe (ver "A cópia para enviar").

## Como rodar

```
python central.py
```

ou dois cliques em `Abrir Central.bat`. Gera o `central.html` (e a cópia
`para_enviar/Central Aurora.html`) e abre no navegador padrão.
Se o terminal não reconhecer `python`, use `py` no lugar (o `.bat` já tenta os dois). Não use
`python3`: neste Windows ele é o atalho da Microsoft Store, não o Python instalado.

## Mapa da pasta

| Arquivo | O que é | Pode editar? |
|---|---|---|
| `Base_bruta.xlsx` | Dados da empresa. **Somente leitura** — nada no sistema escreve nele. | Só trocar pela do mês novo (passo a passo no fim). |
| `central.py` | Tudo: lê a planilha, calcula, monta a página. | Sim |
| `redacao.py` | O **único** lugar que escreve texto para o cliente. O cabeçalho dele explica por quê. | Sim |
| `central.html` | Gerado pelo `central.py`. | **Não** — é sobrescrito a cada execução. |
| `para_enviar/Central Aurora.html` | A cópia para mandar para alguém da equipe. Gerada junto. | **Não** — sobrescrita a cada execução. |
| `fontes/` | As fontes da tela (Unbounded nos números e títulos, Inter no texto) e as licenças OFL. O `central.py` embute as fontes dentro do HTML. A Space Grotesk fica porque o `central.py` da V8 usa. | Não precisa. Sem a pasta, a Central roda com a fonte do sistema. |
| `o_que_nao_pode_sumir.md` | A lista, tela por tela, do que nenhuma mudança de aparência pode apagar. O `conferir.py` confere cada item. | Sim — e o item novo entra também em `conferir_inventario`. |
| `cobrancas_registradas.csv` | Cobranças registradas pela tela (só existe depois do primeiro download). **É o único registro dessas cobranças** — a empresa não tem cópia. | Pode abrir e editar no Excel; ver "Registro de cobrança". |
| `conferir.py` + `conferencia_base.json` | Teste de regressão: compara os números com um retrato gravado, confere a estrutura da página e a lista do que não pode sumir, nas duas páginas. Uso no docstring do `conferir.py`. | Sim |
| `teste_fila.js` | Testa, com um DOM simulado, a função do navegador que reordena a fila. Roda pelo `conferir.py` se a máquina tiver Node; sem Node é pulado. | Sim |
| `auditoria/reconferir.py` | Reconferência independente: recalcula direto da planilha e compara com a tela. Uso no cabeçalho dele. | Sim |
| `versoes/V1…V9` | Histórico congelado. Ver "Versões". | **Não** |
| `Abrir Central.bat` | Atalho de duplo clique. | Raramente |

O Leia-me da planilha cita um `Base_analisada.xlsx` para conferência. **Esse arquivo não existe
nesta pasta** e nada depende dele.

## O que não pode mudar de jeito nenhum

Cada regra veio de uma decisão explícita do dono do projeto. O porquê vem junto.

1. **Um comando, numa máquina limpa.** A Central usa só a biblioteca padrão do Python: sem
   `pip install`, sem internet, sem conta, sem chave de API, sem login. Por isso o `.xlsx` é
   lido com `zipfile` + `xml.etree` (classe `LeitorXlsx`), e os gráficos são SVG montado à mão.
   Não traga pandas, openpyxl, Chart.js, CDN, nada. (A `auditoria/` pode usar openpyxl porque
   é ferramenta de conferência, não a Central.) As fontes da tela ficam em `fontes/` e vão
   **embutidas** no HTML; nunca aponte para o Google Fonts nem para outro endereço — a
   conferência reprova qualquer `http://` ou `https://` na página.

2. **"Hoje" é a data de referência da planilha, nunca o relógio do computador.** Ela é lida da
   aba Leia-me (`extrair_data_referencia`). Vale para tudo: dias de atraso, janela de 48 horas,
   semana do relatório, e o carimbo do registro de cobrança no navegador
   (`REF_ISO + 'T09:00:00'`). Com o relógio da máquina, os números mudariam sozinhos amanhã e
   um registro feito agora nasceria semanas depois da referência. O `conferir.py` reprova se
   aparecer `new Date()` no JavaScript.

3. **Dinheiro em centavos, como inteiro.** A planilha guarda **reais** (formato moeda do Excel);
   a leitura multiplica por 100. Nunca some valor em `float`.

4. **Nenhum número escrito à mão.** Tudo que aparece na tela sai de conta sobre a planilha.
   Exemplo real: o texto "24 meses" estava fixo e mentiria com a base do mês seguinte — hoje é
   `meses_de_historico()`.

5. **Estimativa não tem centavo; valor fechado tem.** Estimativa = tudo da Previsão de caixa
   (esperado, conservador, otimista, diferença, descontado, fora da janela) e o valor em risco.
   Fechado = valor de título, vencido, saldo em aberto, cobrado. O arredondamento acontece **na
   conta**, título a título e cliente a cliente, não só na exibição — senão a soma das semanas
   deixa de bater com o total (aconteceu). Use `formatar_reais_estimado` para estimativa e
   `formatar_reais_de_centavos` para valor fechado.

6. **A `Base_bruta.xlsx` é intocável.** O que o usuário registra fica à parte, no CSV.

7. **Uma conta, um lugar.** O comportamento de cada cliente (quanto atrasa, o quanto varia, se
   piorou, sazonalidade, atraso esperado) é calculado **só** em `calcular_perfis_risco`. A fila
   de Hoje, a Régua, a Previsão e o Relatório leem dali. Número novo nasce ali (ou em outra
   função de cálculo compartilhada, como `calcular_semana`) e todas as telas que falam do
   assunto passam a usá-lo. Por quê: a aba Hoje já teve a sua própria média de atraso, incluindo
   meses sazonais, enquanto a de Risco excluía. Para o cliente sazonal as duas davam 11,5 e 7,8
   dias — mesmo nome, mesmo cliente, números diferentes.

8. **Nenhuma inteligência artificial e nenhuma rede.** Todo texto é montado a partir dos
   números já calculados. A redação está isolada em `redacao.py`, com a assinatura
   `redigir(tom, contexto)`, para o dia em que o dono quiser trocá-la por um modelo — mas isso
   é decisão dele. **A Central redige e copia; quem envia WhatsApp ou e-mail é a pessoa.** Não
   integre com serviço nenhum.

9. **Mudou um número de tela? Explique qual e por quê.** Rode a conferência antes e depois e
   diga ao dono o que mudou, em qual tela. Ele confere a conta: a Central inteira se sustenta
   nisso.

## A aparência (redesenho de setembro de 2026)

Direção "Aurora", escolhida pelo dono entre três propostas: céu noturno com faixas de luz no alto
de cada tela, que escurece até o pé da página; menu lateral fixo à esquerda com as cinco telas
(no celular vira uma fileira de cinco ícones); vidro em todo painel; brilho só no que é principal.
Tudo isso mora no fim do `central.py`: a constante `CSS`, uma função `render_conteudo_*` por
tela, `render_trilho` (o menu) e `render_pagina` (com o JavaScript em `JS_PAGINA`).

**A versão atual é a V9:** a V7 (a primeira Aurora) com três peças que o dono escolheu na V8, a
versão enxuta que ele testou depois (ver "Versões"). Da V8 vieram: o atraso mês a mês embaixo da
régua nos cartões da fila; na aba Risco, o gráfico "quem deve muito e quem está piorando" e a
tabela dos clientes; no Relatório, os três números da semana. O resto da V8 (letra de sistema,
menu enxuto, fonte Space Grotesk, cartões de quatro fileiras, gráficos novos da Previsão, listas
de escolha claras) ficou de fora por decisão dele. O CSS do que veio de lá está junto, no bloco
"peças trazidas da V8".

- **Tema escuro sempre.** Decisão do dono.
- **Cada tela abre com uma coisa que domina** (o número grande, na fonte Unbounded), com o
  "N de 5" e o nome da tela logo acima — quem abre qualquer tela vê que existem as outras.
- **Poucas cores, um trabalho só para cada uma, sempre com a palavra junto:**
  vermelho-coral = fora do padrão / risco alto ou crítico; âmbar = piorou / moderado / fora de
  qualquer lista; verde = padrão dele / risco baixo; cinza = nada a julgar (aguardando retorno,
  sem histórico); azul = **dinheiro** (valor no meio de uma frase, barra de valor em risco,
  linha "esperamos receber"); violeta = **ação** (botão principal, aba atual, link). As cores
  foram passadas pelo validador de paleta do guia de visualização: todas se distinguem na visão
  normal; só verde × vermelho fica perto para daltônico, e por isso toda situação leva palavra e
  ícone. Não crie cor nova para um significado novo sem rever isso.
- **Brilho** só em: número grande do alto, botão principal do 1º da fila, aba atual, 1ª
  descoberta, 1º cartão da fila, o primeiro dos três números do Relatório e, no gráfico do Risco,
  o ponto de quem piorou. Se tudo brilha, nada é principal.
- **Tabelas de números** com algarismos de mesma largura; número grande sem isso.
- **Animação só como enfeite:** a aurora e a entrada das telas. Sem ela (ou com "reduzir
  movimento" ligado no computador) o conteúdo aparece igual. Sem JavaScript, as cinco telas
  aparecem uma embaixo da outra.
- **Gráficos, só onde explicam mais rápido que o texto:**
  - Hoje: a régua pequena de cada cartão da fila (o título mais antigo contra a faixa onde o
    próprio cliente costuma pagar — é o "por que ele está aqui" num olhar) e, embaixo dela, o
    atraso mês a mês dos títulos **já pagos**, em todo cartão com histórico
    (`historico_para_desenhar`; pedido do dono). O último valor vem com o mês ("18 d · abr/26")
    para não ser confundido com o título aberto de hoje, que está na régua. Cada gráfico mostra
    pelo menos 20 dias de altura: sem isso, quem paga sempre entre 1 e 5 dias ganhava uma linha
    "nervosa", porque a variação pequena era esticada até a altura toda.
  - Risco: o medidor da carteira; logo abaixo do número grande, "quem deve muito e quem está
    piorando" (saldo em aberto × mudança do atraso nos últimos 6 meses, um ponto por cliente, com
    o nome escrito de quem piorou, de quem está fora de qualquer lista e dos três maiores saldos;
    os dois primeiros escolhem lugar primeiro, e a conferência reprova se o nome deles faltar); a
    barra de valor em risco de cada cliente, nos cartões e na tabela; o mês a mês da ficha.
  - Previsão: o acumulado das 4 semanas, com a diferença escrita entre as linhas.
  - Régua e Relatório não têm gráfico: lá o conteúdo é texto.
- **Defeito conhecido, mantido por decisão do dono:** em alguns Windows, a lista que abre ao
  clicar em "Ordenar por", no tom ou no canal da Régua mostra as opções em cinza-claro sobre fundo
  claro. A V8 tem o conserto (regra `select option, select optgroup` no CSS: lista clara com letra
  escura); o dono decidiu não levá-lo para a V9. Não conserte sem perguntar.
- **Como a tela é conferida olhando:** não há teste automático de tela; depois de mexer, tire foto
  em 1920×1080, 1366×768 e 390 px (o Chrome e o Playwright desta máquina fazem isso) e olhe.

## Por que as coisas são como são

### Aba Hoje
- Entram só clientes com título vencido na data de referência.
- A ordem é a **média de três posições**: valor vencido, dias do título mais antigo e risco
  comportamental (quanto o atraso atual foge da média do próprio cliente, mais a piora recente).
  Pesos iguais, sem número inventado.
- **Quem foi cobrado há até 2 dias corridos desce para o fim** (`JANELA_CONTATO_RECENTE_DIAS`).
  O pedido original foi "48 horas"; virou 2 dias corridos porque a data de referência não tem
  hora.
- A linha é **uma por cliente**, não por título: liga-se para uma pessoa, não para uma fatura.
- Cartões numa coluna só, todos com a mesma cara; o 1º tem só uma borda de luz. Pedido do dono.
- **A frase de cada cliente tem duas ou três linhas** (`gerar_explicacao`). O valor e a
  quantidade ficam no cabeçalho do cartão; títulos e "números por trás da posição" ficam na
  dobra. A frase e a etiqueta colorida saem da mesma função, `situacao_na_fila`, que lê
  `piora_confirmada` e `fora_do_padrao_hoje` do perfil — antes a fila julgava a piora por conta
  própria e dizia "padrão normal" da Vila Nova enquanto a aba Risco mostrava piora confirmada.
- **"O que a Central descobriu"** (`calcular_descobertas`): um achado por tipo — piorou,
  melhorou, fora de qualquer lista, padrão que não é risco, parou de comprar, padrão de
  calendário —, ordenados pelo valor em risco envolvido; entram os três primeiros.
- **Operar a fila:** filtros por situação, ordenação (prioridade da Central, valor, título mais
  antigo, risco comportamental) e busca pelo menu lateral. Em qualquer ordem, quem foi cobrado
  continua no fim e o número do cartão continua sendo a posição na prioridade da Central.
- **"Registrei a ligação"** em cada cartão registra a cobrança como Telefone, no tom sugerido pela
  régua, pelo mesmo caminho do registro da Régua; dá para desfazer pelo aviso que aparece.

### Aba Risco por cliente
- Índice 0–100 com seis pesos que somam 100 (`PESO_NIVEL` 15, `PESO_VARIABILIDADE` 20,
  `PESO_MUDANCA` 30, `PESO_COMPRA` 15, `PESO_FATURAMENTO` 10, `PESO_LIMITE` 10). Cada peso
  aparece na ficha com a frase que o justifica: "índice fechado ninguém confia".
- **Ordena por valor em risco** (saldo em aberto × índice ÷ 100), não pelo índice puro. A
  pergunta do gerente é "onde está o meu dinheiro", não "quem é o pior".
- **Piora só é "confirmada"** se o atraso médio dos últimos `JANELA_TENDENCIA_MESES` (6) meses
  subir pelo menos 5 dias **e** pelo menos 1 desvio-padrão do próprio cliente. Quem atrasava 35
  e passou para 42 não deteriorou.
- **Sazonalidade** (`detectar_sazonalidade`): um mês do ano é sazonal para o cliente se o
  atraso dele naquele mês é mais alto em **todos** os anos em que aparece, em pelo menos 2 anos
  distintos, por pelo menos max(5 dias, 1 desvio-padrão). Meses sazonais saem da média e da
  comparação de tendência, senão sazonalidade vira falsa piora.
- **Selo "padrão dele, não é risco"**: tem vencido, não piorou, atraso estrutural acima da
  mediana e muito previsível (desvio ÷ média ≤ `LIMIAR_CV_ESTAVEL` 0,35). O alerta sugere
  acertar o prazo no contrato em vez de continuar cobrando.
- **Selo "nunca apareceu em lista nenhuma"**: sem vencido, nunca cobrado, atraso subindo. Só
  um cliente leva — o de maior valor em risco.
- **Logo abaixo do número grande, o gráfico "quem deve muito e quem está piorando"**
  (`gerar_dispersao_risco_svg`), pedido do dono na V9: mostra de uma vez onde está o saldo e
  quem está pagando mais tarde que antes. Ocupa a largura toda; no celular rola de lado dentro da
  própria caixa, como os outros gráficos.
- A chamada de cada linha é a primeira frase do alerta (`partes_do_alerta`). Antes era cortada
  no primeiro ponto, que às vezes era o de milhar: "saldo em aberto de R$ 130." no lugar de
  R$ 130.360,00.
- **Embaixo dos 12 cartões, a mesma carteira em tabela** (posição, cliente com selo, situação,
  saldo, índice com a faixa escrita, valor em risco; no pé, a carteira inteira), pedido do dono
  na V9. A tabela é **fixa**: sempre os 12, na ordem do valor em risco; os filtros e a ordenação
  do alto valem só para os cartões. Clicar na linha abre a ficha. As linhas usam a classe
  `tr-risco`, não `risco-linha`, para o JavaScript dos cartões não mexer nelas. No celular a
  tabela mostra posição, cliente, índice e valor em risco, para caber sem rolar de lado.
- Filtros por faixa e "com selo da Central"; ordenação por valor em risco, índice ou saldo. A
  ficha abre como painel à direita; os seis pesos do índice ficam fechados, cada um com a frase.
- "Prazo real" na ficha = dias corridos da **emissão** ao pagamento. Não confundir com "atraso
  médio", que conta a partir do **vencimento**. Já houve rótulo errado aqui.

### Aba Previsão de caixa
- Cada título em aberto recebe a data esperada = vencimento + atraso esperado do cliente
  (`atraso_esperado` do perfil). Quem mudou de comportamento de forma confirmada, para pior
  **ou para melhor**, é projetado pelo recente, não pela média de dois anos. Título que vence
  em mês sazonal usa o atraso daquele mês.
- **Chance de entrar** = fator de risco (índice pode descontar até `PESO_RISCO_NA_ENTRADA`
  30%) × fator de idade (o título só perde chance depois de passar do maior atraso que *aquele*
  cliente já praticou), com piso `PISO_CHANCE_ENTRADA` de 25%. **Isto é premissa, não medida**:
  a base não tem nenhum título baixado como perda, então não há inadimplência observada para
  calibrar. A tela diz isso.
- "O que a planilha apontaria" = soma dos vencimentos das 4 semanas, **com o que já venceu
  lançado na semana 1**, para os dois números cobrirem a mesma carteira.
- Conservador e otimista usam o dia ruim e o dia bom de cada cliente (atraso esperado ± 1
  desvio-padrão).
- Considera **só a carteira existente**, nenhuma venda nova. A tela diz isso.
- **Cliente sem histórico** (menos de `MIN_TITULOS_HISTORICO` = 5 títulos pagos fora de mês
  sazonal): não dá para dizer quando ele paga, então a Central **não estima**. O título dele sai
  das duas contas (planilha e esperada) e a tela mostra um aviso com o valor que ficou de fora.
  Antes ele entrava só na conta da planilha e inflava a diferença em destaque.

### Aba Régua de cobrança
- Entram os clientes com título vencido **mais** o que a aba Risco marcou como "nunca apareceu
  em lista nenhuma".
- Seis tons em `redacao.TONS`. A escolha automática (`escolher_tom`) segue esta ordem, e a
  ordem é decisão de negócio: nada vencido → contato preventivo; parou de comprar
  (`RAZAO_PAROU_DE_COMPRAR` 2× o intervalo normal) → reaproximação comercial, *antes* de
  qualquer cobrança; piora confirmada → conversa, não ameaça; padrão estável → lembrete leve com
  sugestão de ajustar o prazo; atraso fora do padrão dele → primeiro aviso cordial; dentro do
  padrão → lembrete leve; senão → cobrança objetiva.
- Os tons × três canais de cada cliente são escritos **no Python** e embutidos na página. Por
  isso trocar o tom na mão reescreve o texto na hora sem o navegador redigir nada.
- **Cada tom declara de quais números do cliente o texto depende** (chave `usa` em
  `redacao.TONS`), e `redacao.tons_possiveis()` só oferece o tom se esses números existirem.
  Motivo: para um cliente sem histórico, o texto sairia "vocês sempre pagaram numa média de —
  dias". Cliente sem histórico recebe a cobrança objetiva, que não fala do passado, e o seletor
  mostra só os tons possíveis, com um aviso dos que ficaram de fora. Tom novo tem que declarar o
  seu `usa`.
- **Lista de um lado, um cliente do outro.** Antes os oito cartões ficavam empilhados com a
  mensagem inteira aberta. Todos os textos continuam na página; só o selecionado aparece. Setas
  para cima e para baixo trocam de cliente quando a lista tem o foco.
- Assinatura: "Equipe financeira / Distribuidora Aurora".
- **A mensagem ao cliente nunca diz quanto ele pesa no nosso faturamento** — isso dá poder de
  barganha a ele. O número fica só no roteiro interno da ligação.
- Os textos foram relidos como se fosse o cliente recebendo. Nada de "1 dias", "título(s)",
  "prezado cliente, informamos que consta", "somando R$ X" para um título só. Em `redacao.py`,
  use `quantos()` (números até dez por extenso: "dois títulos"), `dias_txt()` e `valor_total()`.
  Mudou um texto? Releia a versão de um título só e a de vários.

### Registro de cobrança
- Uma página aberta por duplo clique **não consegue gravar arquivo** — trava do navegador. Por
  isso o registro é híbrido: vai para o `localStorage` do navegador (a fila de Hoje reordena na
  hora e sobrevive a fechar e abrir) e o botão **Baixar registros (.csv)** gera o arquivo que o
  Python lê na próxima execução. Só depois disso a cobrança entra nas frases, no bloco de
  contato e no relatório, e fica visível para outra pessoa.
- O arquivo tem que se chamar exatamente `cobrancas_registradas.csv` e ficar **nesta pasta**. O
  navegador costuma salvar em Downloads e, na segunda vez, como `cobrancas_registradas (1).csv`.
- **Onde o arquivo tem que ficar, e se ele foi lido, sai no terminal, não na tela.** Ao rodar
  `python central.py`, a última linha diz se o arquivo foi encontrado e quantas cobranças foram
  lidas, ou que ainda não existe (com o caminho exato de onde salvar), e avisa quando há um
  arquivo com o nome errado na pasta (`situacao_arquivo_registros`). O `Abrir Central.bat`
  termina em `pause`, para dar tempo de ler. Antes isso aparecia no painel da Régua, com o
  caminho desta pasta; o dono pediu para tirar, porque a página pode ser aberta em outro
  computador, onde esse caminho não vale. A conferência reprova se o caminho da pasta voltar a
  aparecer em qualquer uma das duas páginas.
- CSV e não JSON para o dono poder abrir no Excel. Separador `;` e BOM UTF-8 porque é o que o
  Excel em português espera. O leitor aceita a data e o valor do jeito que o Excel reescrever
  (`_ler_data_hora`, `_ler_centavos`).
- O CSV **acumula os meses** e continua valendo com a base nova: a aba Cobranças da empresa
  **não** traz o que foi registrado pela Central. Contatos antigos saem da janela de 2 dias
  sozinhos, pela data.
- **O download junta o arquivo e o navegador.** A página recebe o que já está no CSV
  (`<script id="registros-arquivo">`), e o botão baixa tudo — o que está no arquivo mais o que
  só este navegador tem — sem repetir. Motivo: antes o botão baixava só o navegador de quem
  clicava; alguém de outra máquina, com o navegador vazio, registrava uma cobrança, salvava por
  cima e **apagava os meses anteriores**. O painel marca cada linha como "já no arquivo" ou "só
  neste navegador — baixe o arquivo". O `teste_fila.js` prova os dois casos.
- **"Limpar os que eu registrei"** apaga só o que está guardado no navegador. O que já está no
  CSV continua — para tirar uma linha do arquivo, edite o CSV no Excel.

### Relatório da semana
- É a quinta tela (antes abria numa janela nova, que o navegador podia bloquear).
- **Abaixo da data do alto, três números da semana** (quanto entrou, quanto venceu e não foi
  pago, o vencido acumulado), pedido do dono na V9. Saem de `calcular_semana` e do mesmo total da
  aba Hoje. Não vão para a impressão, que sai só com o relatório.
- Texto corrido com os números dentro das frases, terminando na fila de contato ordenada por
  valor. Quem lê tem quatro minutos e não abre a Central.
- **Todo valor em dinheiro vem na cor de dinheiro** (`prosa` com `RE_DINHEIRO`); a conferência
  conta um por um. O texto copiado é o mesmo, sem formatação. Imprimir sai só o relatório, em
  papel branco.
- **Não calcula nada**: lê as mesmas funções das telas. A "semana" é a de 7 dias terminando na
  data de referência (`calcular_semana`), e só conta cobranças registradas dentro dela.

## O que parece defeito e não é

- **Hortifrúti Terra Boa com saldo R$ 0, ocupação 0% e valor em risco zero.** Real: compra à
  vista e não tem título em aberto. A reconferência confirma esse tipo de zero.
- **"Cobrança objetiva" nunca é escolhida** nesta base. É o último caso da régua; a regra é
  alcançável, só não tem cliente nessa faixa.
- **Na tabela semanal da Previsão, o conservador passa do otimista em algumas semanas.**
  Atrasar empurra dinheiro para a semana seguinte. Por isso a faixa só aparece no acumulado,
  onde nunca inverte — e a tela explica.
- **Dois títulos da Mercearia Dona Zica com a mesma chance (22%).** Os dois estão no piso de
  idade; o que sobra é o fator de risco do cliente.
- **Atacado São Jorge em 1º na fila de Hoje, sendo bom pagador.** É exatamente o ponto: atrasa
  em média 3 dias e tem um título com 51.
- **Picos de janeiro e fevereiro no gráfico do Congelados Polo Norte.** Sazonalidade detectada
  e descontada (anéis vazados no gráfico).
- **Atraso médio "2,7 dias" na aba Hoje e "3 dias" na ficha.** Mesmo número, casas decimais
  diferentes.
- **A tabela "título a título", a tabela semana a semana e os gráficos rolam de lado no
  celular.** Proposital: rolam dentro da própria caixa para não empurrar a página nem encolher
  até ficar ilegível.
- **Ordenando a fila por valor, os números dos cartões ficam fora de ordem (3, 1, 6…).** É o
  número da prioridade da Central, de propósito: a ordem da tela mudou, a prioridade não.
- **Um cliente registrado agora aparece "Aguardando retorno do contato" mas a frase ainda diz
  "fora do padrão".** A frase é a análise do Python; o registro do navegador só entra na análise
  depois de baixar o CSV e rodar de novo.
- **No celular, o menu mostra "Caixa" em vez de "Previsão de caixa".** Nome curto para caberem
  as cinco telas.
- **Mercearia Dona Zica com 112 dias na régua e "18 d · abr/26" no atraso mês a mês.** São duas
  coisas: a régua mostra o título aberto hoje (venceu em 22/04 e não foi pago); o atraso mês a mês
  mostra só títulos já pagos, e o último mês com título pago é abril (16 e 21 dias, média 18,5).
  O título aberto só entra no gráfico quando for pago. A linha para em abril porque nada que
  venceu depois foi pago.
- **Atacado São Jorge e Adega Serra Azul com o atraso mês a mês quase reto.** É o retrato certo:
  pagam sempre igual. A altura mínima de 20 dias impede que a variação pequena pareça instável.
- **Filtrando a aba Risco por faixa, a tabela de baixo continua com os 12.** De propósito: ela é
  sempre a carteira inteira, com o total no pé. Os filtros valem para os cartões.
- **No celular, a tabela de Risco não mostra situação nem saldo.** Para caber sem rolar de lado; os
  dois estão no cartão logo acima e na ficha.
- **`python conferir.py` termina com saída 3 depois de trocar a planilha.** Esperado — ver o
  passo a passo.
- **Aviso "... fora da previsão" na Previsão e régua com menos de seis tons para um cliente.**
  É cliente sem histórico suficiente (na régua, também quem não tem pagamento recente para
  comparar). A Central diz que não sabe em vez de estimar. Na base de agosto não acontece; com
  cliente novo na base, vai acontecer.
- **Cliente novo sempre com "cobrança objetiva".** Mesmo motivo: sem histórico, os tons que
  comparam com o passado não se aplicam. A Régua explica o porquê do tom ("Cliente sem
  histórico suficiente: só 2 títulos pagos até hoje…").

## Como mexer (tela, aba nova, número novo)

`central.py` é longo; as seções estão na ordem do fluxo: leitor de `.xlsx` → utilitários de
texto e número → `carregar_dados` → `calcular_fila` e `gerar_explicacao` →
`calcular_perfis_risco` e o que deriva dele (previsão, régua, semana) → CSS (constante `CSS`)
→ uma função `render_conteudo_*` por aba → `render_pagina` (com o JavaScript) → `main`.

- **Texto para o cliente:** só em `redacao.py`.
- **Dias e porcentagem em qualquer texto:** sempre `texto_dias()` e `texto_pct()`. Nunca
  `f"{x} dias"` nem "dia(s)" à mão — já houve "1 dias" em vinte lugares e "13.0%" com ponto.
- **Tela nova:** uma linha em `TELAS` (o menu e o "N de 5" saem daí), um
  `<section class="tela" id="tela-NOME">`, o nome no array `TELAS` do `JS_PAGINA`, e o nome no
  laço de abas e em `conferir_inventario` do `conferir.py`.
- **Texto que some ou encolhe:** confira o `o_que_nao_pode_sumir.md`. Encurtar e guardar atrás de
  um clique pode; sumir não. O `conferir.py` reprova se sumir.
- **JavaScript:** nada no nível de cima do `JS_PAGINA` pode mexer na página além das duas linhas
  que leem `TEXTOS` e `REG_ARQUIVO` — o `teste_fila.js` roda esse código num DOM de mentira que
  não tem o resto. Leitura de página vai dentro de `DOMContentLoaded`.
- **Número novo:** nasce na função de cálculo compartilhada, as telas e o relatório leem dali,
  e ele entra em `retrato()` do `conferir.py`.
- **Mudou a marcação HTML?** `conferir.py` e `auditoria/reconferir.py` localizam as coisas por
  regex. Se passarem a não achar, eles reprovam avisando ("a marcação mudou?"). Ajuste o regex;
  não apague a verificação.
- **Não existe teste automático de tela no navegador nesta pasta.** Depois de mexer na cara,
  abra a página e olhe em tela larga e estreita (390 px): nada rolando de lado, nenhum texto
  por cima de outro. Um botão já ficou por cima do título por estar em `position: absolute`; no
  redesenho, um ícone sem tamanho definido cobriu a busca inteira e o número grande saiu cortado.
- **f-string:** nada de barra invertida dentro das chaves `{}` de uma f-string no `central.py`:
  isso só roda no Python 3.12 em diante, e a Central tem que abrir em máquina limpa com Python
  mais antigo.
- **Ao terminar qualquer mudança:** `python central.py`, `python conferir.py`,
  `python auditoria/reconferir.py`. Se números mudaram de propósito, diga ao dono quais e
  regrave a base (`python conferir.py --gravar-base`).

## Formato que a planilha precisa ter

O código lê as colunas **pelo nome do cabeçalho**. Se a empresa renomear alguma, a Central
quebra ao abrir; ajuste os nomes em `carregar_dados` (e, se for uma das colunas abaixo marcadas
com \*, também em `auditoria/reconferir.py`).

- Arquivo `Base_bruta.xlsx`, com as abas **Leia-me**, **Clientes**, **Títulos** e **Cobranças**.
- **Leia-me:** uma linha cuja primeira célula contém "DATA DE REFERÊNCIA" e, ao lado, a data
  por extenso: "12 de agosto de 2026".
- **Clientes:** Código do cliente\*, Nome do cliente\*, Categoria, Região, UF, Prazo de
  pagamento contratado, Prazo em dias, Limite de crédito\*, Cliente desde, Contato, Telefone.
- **Títulos:** Número do título, Código do cliente\*, Data de emissão\*, Valor do título\*,
  Data de vencimento\*, Data de pagamento\* (**vazia = em aberto**).
- **Cobranças:** Código do cliente, Data e hora do contato, Canal, Tom usado, Valor cobrado,
  Títulos na cobrança.
- Valores em **reais**. Datas como data do Excel.
- A planilha pode ter passado por Excel, LibreOffice ou Google Planilhas: o leitor aceita os
  dois jeitos de o `.xlsx` apontar para as abas. (Antes não aceitava, e uma planilha
  re-salva fora do Excel nem abria.)

## Como trabalhar com o dono do projeto

- **Pergunte antes de escolher** quando houver dúvida sobre o que ele quer — principalmente
  arquitetura, formato de entrega e qualquer texto que sai da empresa. Ele prefere parar dois
  minutos a descobrir depois que a Central mandou a mensagem errada para o cliente errado.
  Dúvidas pequenas e reversíveis podem seguir um padrão sensato, desde que você diga qual foi.
- **Achou defeito?** Procure a causa antes de remendar a aparência, confirme que é defeito antes
  de consertar, e não conserte o que não está quebrado. Se uma verificação reprovar, o conserto
  é no código — a menos que você demonstre que a verificação é que está errada.
- Desconfie de número redondo demais (0%, 100%, dias idênticos, categoria vazia): quase sempre
  é conta que não rodou. Confira antes de aceitar.
- Ao entregar, diga o que mudou de número e em qual tela.

## A cópia para enviar

`python central.py` grava também `para_enviar/Central Aurora.html`: um arquivo só, com fontes,
dados e textos dentro, que abre com dois cliques em qualquer computador com navegador — sem
Python, sem internet, sem instalar nada. É isso que se manda para alguém da equipe (e-mail,
WhatsApp, pendrive, pasta compartilhada).

O que muda na cópia, e por quê (a tela diz cada coisa; nada fica fingindo que funciona):

- **É um retrato do dia em que foi gerada.** Os números não se atualizam nela; a cópia avisa no
  menu e no rodapé. Para números novos, gere e mande de novo.
- **Registrar cobrança marca só no navegador de quem recebeu** (decisão do dono): a fila dele se
  reorganiza e fica marcado quem ele já ligou, mas isso não entra na conta da Central. O botão
  "Baixar registros (.csv)" **não existe** na cópia, com a explicação no painel: um CSV vindo de
  outra máquina, salvo por cima do oficial, apagaria cobranças. A cópia guarda os registros numa
  chave própria do navegador (`aurora_cobrancas_copia_enviada`), para não se misturar com a
  Central de quem gerou, se as duas forem abertas no mesmo computador.
- **Não mostra o caminho da pasta** de quem gerou, nem manda "rodar python".
- No celular, depende do aplicativo: no navegador funciona; em pré-visualização de anexo que não
  roda JavaScript, as cinco telas aparecem uma embaixo da outra, sem os botões de operar.

## Versões

- `versoes/V1…Vn`, uma pasta por versão. **Só cria versão quando o dono pedir** — com uma
  exceção: **antes de trocar a planilha, sempre congele a versão atual.**
- O que vai na pasta: `central.html`, `central.py`, `redacao.py` (o `central.py` não roda sem
  ele) e, quando o congelamento é por troca de planilha, também a `Base_bruta.xlsx` daquele mês
  e o `cobrancas_registradas.csv` se existir — assim a versão fica reproduzível.
  `conferir.py`, `teste_fila.js`, `conferencia_base.json` e `auditoria/` não vão, salvo pedido.
- V1 e V2 têm só o HTML: os scripts daquela época foram sobrescritos antes de existir esta regra.
  V3 e V4 são anteriores ao `redacao.py`. Nada disso é defeito.
- V6 é a última versão com a aparência antiga (clara, abas no alto), congelada antes do
  redesenho de setembro de 2026.
- V7 é a primeira versão Aurora (letra grande, menu com legenda e bloco "nesta tela", lista de
  Risco em cartões, fonte Unbounded), congelada a pedido do dono antes da versão enxuta. A pasta
  guarda também a cópia para enviar daquela versão, em `para_enviar/`.
- V8 é a versão enxuta (letra de sistema, menu só com as telas, Risco em tabela, fonte Space
  Grotesk, gráficos novos na Previsão e no Risco), congelada a pedido do dono. Ele gostou, mas
  escolheu levar só alguns pontos dela para a V7; o resultado dessa mistura vira a V9. Com a
  cópia para enviar em `para_enviar/`.
- V9 é a versão atual: a V7 com o atraso mês a mês nos cartões da fila, o gráfico e a tabela da
  aba Risco e os três números do Relatório, trazidos da V8 (ver "A aparência"). Com a cópia para
  enviar em `para_enviar/`.
- A pasta não é repositório git: sobrescrever sem congelar perde o anterior de verdade.

## Passo a passo: trocar a planilha pela do mês novo

Tudo isto foi ensaiado de ponta a ponta com uma planilha imitando o mês seguinte.

**1. Confira que o estado atual está bom.**
```
python conferir.py
```
Tem que terminar em `RESULTADO: conferencia aprovada.` Se não terminar, pare: o problema é
anterior à planilha nova, e trocar agora mistura as duas coisas.

**2. Congele a versão atual.** Descubra o próximo número em `versoes/` (se a última é V5, esta é
V6) e rode, trocando `V6` no fim pelo número certo:
```
python -c "import os,shutil,filecmp,sys; v=sys.argv[1]; d=os.path.join('versoes',v); os.makedirs(d); fs=[f for f in ('central.html','central.py','redacao.py','Base_bruta.xlsx','cobrancas_registradas.csv') if os.path.exists(f)]; [shutil.copy2(f,d) for f in fs]; [print(' ',f,'identico' if filecmp.cmp(f,os.path.join(d,f),shallow=False) else 'DIFERENTE') for f in fs]" V6
```
Ele cria a pasta, copia `central.html`, `central.py`, `redacao.py`, `Base_bruta.xlsx` e
`cobrancas_registradas.csv` (se existir) e compara byte a byte. Todos têm que sair `identico`.
Se a pasta já existir, ele para com erro em vez de sobrescrever: escolha o número seguinte.

**3. Troque a planilha.** Salve a nova **por cima** de `Base_bruta.xlsx`, com esse nome exato.
**Não** mexa no `cobrancas_registradas.csv`: ele continua valendo no mês novo.

**4. Gere a Central.**
```
python central.py
```
O terminal mostra `Data de referencia usada:` — tem que ser a data do Leia-me da planilha nova.
Se der erro ao ler, o motivo mais provável é coluna ou aba renomeada (ver "Formato que a
planilha precisa ter").

**5. Rode a conferência.**
```
python conferir.py
```
O esperado é `[BASE NOVA] ... numero(s) mudaram -- esperado` e
`RESULTADO: estrutura aprovada com a planilha nova.` (saída 3). Os números **têm** que mudar: a
comparação é com a base anterior. O que **não** pode é `[FALHOU]` na estrutura da página ou no
teste da fila — isso é defeito.

**6. Prove que continua fechando.**
```
python auditoria/reconferir.py
```
Recalcula tudo direto da planilha nova por outro caminho e compara com a tela e as telas entre
si. Tem que terminar em `RESULTADO: a tela fecha com a planilha, e as telas fecham entre si.`
A saída diz o que aconteceu:
- **0** — fechou. Siga.
- **1** — divergência (linhas com `!!`). **Não regrave nada.** Descubra qual dos dois está
  errado, a conta da Central ou a reconferência, e avise o dono antes de continuar: a Central
  não pode ir para o mês novo com número que não fecha.
- **2** — não conseguiu conferir. Se diz `[PULADO] o openpyxl nao esta instalado`, rode
  `pip install openpyxl` (precisa de internet; a Central em si não precisa dele). Se não der para
  instalar, confira à mão pelo menos o total em aberto e o vencido da aba Hoje contra a planilha
  filtrada no Excel, e diga ao dono que a reconferência não rodou. Se diz que o `central.html`
  foi gerado com outra data, você pulou o passo 4.

**7. Grave a base nova e confirme.**
```
python conferir.py --gravar-base
python conferir.py
```
Agora tem que voltar a `RESULTADO: conferencia aprovada.`

**8. Olhe a página.** Abra a Central e passe pelas cinco telas. Desconfie do
que vier redondo demais. Confira, na saída do terminal do passo 4, a linha sobre o arquivo de
registros ("Registros de cobranca: lendo N cobrancas…" ou "ainda nao existe…").

**9. Conte ao dono** a data de referência nova, que a reconferência fechou, e qualquer coisa
estranha que você tenha visto no passo 8.
