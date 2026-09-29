---
name: apresentacao-de-impacto
description: Cria apresentações, propostas comerciais e decks visuais de tela cheia em HTML — palco fixo que escala sozinho, cartões de vidro, gráficos SVG desenhados em JS, um controle interativo que recalcula as telas seguintes, e verificação automática de encaixe antes de entregar. Use sempre que o usuário pedir apresentação, slides, deck, pitch, proposta comercial, orçamento para cliente, one-pager, relatório visual ou "material bonito para apresentar" — mesmo que ele não diga HTML e mesmo que só mande o conteúdo bruto. Use também quando pedir para melhorar, redesenhar, "dar impacto" ou consertar uma apresentação ou proposta que já existe.
---

# Apresentação de impacto

Esta skill produz apresentações que se defendem sozinhas na frente de um cliente: uma tela cheia
por ideia, o número que importa em tamanho de manchete, a conta que sustenta esse número logo
abaixo dele, e gráficos que se desenham na hora em que você fala sobre eles.

O que separa uma apresentação incrível de um amontoado de slides não é efeito visual — é **foco**.
Cada tela carrega uma ideia só, e o título já entrega a conclusão. O resto da tela existe para
provar o título. Se você se pegar escrevendo um título que anuncia o assunto ("Análise de
conversão"), troque pela conclusão ("A rede não perdeu público. Perdeu fechamento."). Essa troca
sozinha muda mais a qualidade percebida do que qualquer gradiente.

## O caminho mais curto

Os caminhos abaixo são relativos à **pasta desta skill** (o caminho aparece como "Base directory"
quando a skill é carregada). Use o caminho completo a partir dela.

1. Copie `assets/modelo.html` para o destino e renomeie. Ele já vem com palco, moldura,
   navegação, ícones, cartões, animações e o motor de cálculo — tudo funcionando.
2. Troque os tokens de marca no topo do CSS (bloco `:root`) se a apresentação não for da Praxis.
3. Substitua as telas de exemplo pelo conteúdo real, uma ideia por tela.
4. Rode `node <pasta-da-skill>/scripts/conferir_telas.mjs <arquivo.html>` e conserte o que ele apontar.
5. Só então entregue.

O passo 4 não é opcional. O palco tem altura fixa e nada avisa quando o conteúdo passa do fim —
o texto simplesmente cai por cima do rodapé e você só descobre na reunião. O script pega isso em
segundos e é o que separa um material profissional de um constrangimento.

## A regra que governa todo o resto: o orçamento de altura

O palco é fixo em **1600 × 900** e escala inteiro para caber na janela (`transform: scale`).
Isso é o que faz a apresentação parecer a mesma coisa em qualquer tela, projetor ou PDF — e é
também a restrição que você precisa respeitar o tempo todo.

Cada tela tem `padding: 112px 88px 100px`. Sobram **688px de altura útil e 1424px de largura**.
Antes de escrever uma tela, faça a conta grosseira:

| Peça | Altura aproximada |
|---|---|
| Cabeçalho com sobrancelha + título de 1 linha | 100px |
| Cada linha extra do título (56px) | 60px |
| Linha de apoio (`.lede`, 21px) | 32px por linha |
| Rodapé de procedência (`.src`, mono 14px) | 22px por linha + 18px de respiro |
| Cartão com rótulo, número grande e texto | 300–360px |
| Gráfico confortável | 340–430px |

Se a soma passar de 688, corte conteúdo antes de encolher fonte. Encolher fonte resolve o encaixe
e estraga a leitura à distância — e apresentação é lida a três metros, não a trinta centímetros.
Quando não der para cortar, prefira: tirar uma linha do rodapé, reduzir o gráfico, e só por último
baixar o corpo de texto (nunca abaixo de 16px, nunca o número grande).

## O arco narrativo

Para proposta comercial, esta sequência funciona e você pode adaptá-la a qualquer material — o que
importa é a lógica: **estabelecer o problema em números antes de falar de solução, e falar de preço
só depois que o tamanho do problema já está claro.**

1. **Capa** — nome do cliente grande, uma frase do que se quer resolver, tira de quatro números que
   respondem "quanto, quanto tempo, quanto de entrada, até quando" antes de qualquer pergunta.
2. **O problema em três números** — três cartões, três números grandes. Não quatro, não seis.
3. **Por que o problema existe** — uma decomposição, de preferência em cascata, que separe o que o
   cliente controla do que ele não controla.
4. **Hipóteses** — o que você acha que está acontecendo, com o teste que derrubaria cada hipótese.
   Dizer o que refutaria sua própria tese é o que faz um diagnóstico soar honesto em vez de vendido.
5. **Entregas** — o que continua existindo depois que você sai da sala.
6. **Fases** — linha do tempo + cartões que abrem um por vez conforme você avança.
7. **Investimento** — tabela fechada, condição de pagamento explícita, e o que o preço *não* inclui.
8. **O retorno (interativo)** — o cliente arrasta um controle e vê a conta mudar. Ver isso ao vivo
   vale mais que dez slides de projeção, porque o número passa a ser dele, não seu.
9. **Quando se paga** — curva acumulada com faixa de cenários e o ponto de equilíbrio marcado.
10. **Fora do escopo** — seis itens com o motivo de cada um.
11. **Premissas** — de onde veio cada número, e o que muda se a premissa cair.
12. **Para começar** — onde o cliente chega e os passos concretos, com prazo.

Para material que não é proposta (relatório, prestação de contas, apresentação institucional),
mantenha o esqueleto: **número → causa → o que fazer → o que custa → o que acontece depois**.

## Disciplina de conteúdo

**Uma ideia por tela.** Se a tela tem dois assuntos, ela vira duas telas. Tela com duas ideias
não tem nenhuma.

**O título diz a conclusão.** A sobrancelha (`.eyebrow`, mono, cor de acento) diz o assunto; o
título (`h2`, serifada fina, com um trecho em itálico dourado) diz o que você concluiu. Ponha o
itálico no pedaço que carrega o argumento, não numa palavra qualquer.

**Todo número grande vem com a conta.** Embaixo de cada número, um bloco `.conta` em mono, com a
divisão que gerou aquele número. Isso é o que transforma "confie em mim" em "confira comigo" — e
é o que sobrevive quando o cliente leva o PDF para o sócio.

**Procedência no rodapé.** Toda tela que usa dado do cliente termina com uma linha `.src` dizendo
de onde veio e se foi auditado. Se você interpolou, desenhou uma tendência ou escolheu um cenário,
diga na própria tela. Um gráfico bonito com dado inventado é o jeito mais rápido de perder a conta.

**Cor com significado.** Coral = problema e perda. Âmbar = dinheiro e preço. Menta = ganho e
resultado. Violeta = processo e fases. Azul = entregas e método. Use o acento do cartão para dizer
ao olho o que o texto vai dizer ao cérebro.

## As quatro partes do sistema

Leia o arquivo de referência quando for mexer naquela parte — não precisa ler tudo de uma vez.

- **`references/sistema-visual.md`** — tokens, anatomia do cartão de vidro, escala tipográfica,
  as três luzes de fundo que trocam por tela, grão/grade/vinheta, moldura com régua e contador, e
  as classes de animação. Leia antes de mexer em CSS ou de trocar a identidade visual.
- **`references/graficos.md`** — as cinco receitas de gráfico em SVG desenhado por JS (declínio,
  cascata, linha do tempo, curva de payback, medidores), a interpolação cúbica monótona, a busca do
  ponto de cruzamento e como sincronizar a animação com o traçado. Leia antes de desenhar qualquer
  gráfico.
- **`references/interatividade.md`** — como montar a espinha aritmética única e fazer um controle
  recalcular as telas seguintes, com os três estados de veredito. Leia antes de criar a tela
  interativa.
- **`references/narrativa.md`** — modelos de texto por tipo de tela, com exemplos de título fraco
  contra título forte. Leia quando estiver escrevendo o conteúdo, não a estrutura.

## Interatividade: um número manda em tudo

A tela que mais impressiona é aquela em que o cliente mexe. A regra para ela não virar um monstro:
**uma única fonte de verdade**. Todas as constantes ficam num bloco só no topo do script, todo
número exibido deriva delas por função, e uma função `propagar()` reescreve as telas seguintes
sempre que o controle muda. Nunca escreva um número calculado dentro do HTML — se ele aparece em
duas telas, ele tem que vir da mesma função nas duas.

O limiar que decide se o projeto se paga também é **calculado**, por busca binária, nunca escrito à
mão. No dia em que o preço mudar, a marca na barra anda sozinha. Detalhes em
`references/interatividade.md`.

## Verificação antes de entregar

```bash
node <pasta-da-skill>/scripts/conferir_telas.mjs caminho/para/apresentacao.html
```

O script abre cada tela num navegador sem interface, tira uma foto e reporta três coisas: elemento
que passou da área útil, conteúdo que vazou do próprio quadro (apontando **qual filho** está
forçando a altura) e rótulo de gráfico cortado pela borda do SVG, dizendo de que lado aumentar a
margem. Também coleta erros de JavaScript. Conserte tudo antes de mostrar o arquivo para alguém.

Quando ele acusar "o conteúdo passa Xpx do próprio quadro", olhe o filho que ele nomeia — numa
grade de duas colunas o culpado quase sempre é a coluna **mais alta**, não o gráfico ao lado. Reduzir
o gráfico nesse caso não muda nada.

Depois de rodar o script, **olhe as fotos**. O script pega colisão e estouro; ele não tem opinião
sobre uma tela com um buraco de 200px no meio, um título que quebrou numa palavra feia, ou um
gráfico que ficou espremido. Esses você só vê olhando.

### Quando o script não puder rodar

Ele precisa de Node 18+ e de um Chrome ou Edge instalados. No Claude web, ou em qualquer ambiente
sem isso, o script não roda — e aí a disciplina do orçamento de altura deixa de ser recomendação e
vira a única defesa. Nesse caso:

1. Some as alturas de cada tela **antes** de escrever, com a tabela acima, e deixe 40px de folga.
2. Prefira menos conteúdo por tela a fontes menores.
3. Entregue dizendo ao usuário que a conferência automática não rodou e que vale abrir o arquivo
   em tela cheia antes de apresentar — é honesto e leva dez segundos para ele fazer.

Se houver como rodar um navegador sem interface no ambiente, rode: uma tela estourada descoberta na
reunião custa muito mais que dois minutos de verificação.

## Armadilhas que custam caro

Estas são reais — todas já quebraram uma tela pronta:

- **Regras de CSS com especificidade maior anulando o respiro do ícone.** O ícone do cartão mora no
  canto superior direito, e o conteúdo desvia dele com `padding-right`. Se você criar uma regra mais
  específica para aquele elemento, o texto volta a passar por baixo do ícone. Ao adicionar padding
  para uma tela específica, reponha o `padding-right` do primeiro e do segundo filho.
- **Rodapé longo demais.** `.src` com três linhas é a causa número um de estouro. Duas linhas é o
  teto prático.
- **Rótulo de eixo mais largo que o próprio trecho.** Em linhas do tempo, a primeira fase costuma
  ser curta e o nome dela invade o nome da segunda. Escalone os rótulos em duas alturas.
- **Texto dentro de SVG escalado.** Se o `viewBox` tiver largura diferente da largura de render, a
  tipografia do gráfico encolhe junto e fica ilegível. Case a largura do `viewBox` com a largura
  real em pixels do container.
- **Margem do gráfico menor que o texto das pontas.** O rótulo do fim da barra mais comprida é o
  primeiro a ser cortado, e some sem deixar rastro. Dimensione `L` pelo rótulo mais longo e `R`
  pelo valor mais a nota. O verificador acusa isso desde que você o rode.
- **Chamada bloqueada dentro de iframe derrubando o resto da tela.** Esta e a pior de todas,
  porque nao deixa rastro. No preview do claude.ai, no Notion e em qualquer `<iframe sandbox>`,
  chamadas como `history.replaceState` e `requestFullscreen` **lancam erro**. Se uma delas estiver
  no meio da funcao que troca de tela, tudo que vem depois para de rodar — inclusive o desenho dos
  graficos. A tela troca normalmente, a navegacao funciona, e os graficos aparecem em branco sem
  nenhuma mensagem. O modelo ja traz a correcao: a atualizacao da URL vai **por ultimo** dentro de
  `show()` e dentro de um `try/catch`. Ao acrescentar qualquer chamada de API do navegador nessa
  funcao, pergunte se ela sobrevive num iframe — e isole se houver duvida. O verificador testa isso
  automaticamente, injetando a falha de proposito.
- **Topo dos algarismos cortado no número grande.** É a armadilha mais traiçoeira do sistema:
  `background-clip:text` só pinta dentro da caixa do elemento, e `line-height:.86` deixa o alto dos
  algarismos fora dela. O número aparece decapitado e **nenhum verificador pega** — o texto está lá,
  só não foi pintado. O modelo já traz a correção (`padding-top:.18em` com `margin-top:-.18em` nas
  classes de gradiente); se você criar uma classe nova com gradiente em texto, repita o par.
- **Cartão do simulador com buraco no meio.** Se a tela interativa sobrar espaço, dê
  `flex:1;display:flex;flex-direction:column;justify-content:space-between` ao cartão para os
  blocos se distribuírem. O verificador não enxerga buraco — só a foto mostra.
- **Acento em itálico serifado.** Evite pôr palavras com maiúscula acentuada (Í, Á, É) dentro do
  `<em>` do título; o acento fica frágil no corpo grande. Reescreva a frase.
- **Sequência de teclas brigando com o controle deslizante.** Quando o foco está no `input[range]`,
  as setas devem mover o controle, não trocar de tela. Trate isso no `keydown`.

## Quando o pedido não é uma apresentação inteira

Se o usuário só quer **uma tela** (um gráfico para colar em outro lugar, um comparativo, um
one-pager), use o mesmo modelo com uma tela só e remova a moldura de navegação. O sistema visual
funciona igual e você entrega em minutos.

Se o usuário mandar **uma apresentação pronta para melhorar**, leia o arquivo inteiro antes de
mexer, identifique o que já funciona e preserve o conteúdo. Reestilizar é trocar o sistema visual
mantendo os textos; refazer é reescrever os textos. Pergunte qual dos dois ele quer se não estiver
claro — são trabalhos de tamanhos muito diferentes.
