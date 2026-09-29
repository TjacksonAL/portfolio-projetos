# Narrativa e texto

Índice: [A regra do título](#a-regra-do-título) · [Modelos por tela](#modelos-por-tela)
· [O bloco da conta](#o-bloco-da-conta) · [Procedência](#procedência) · [Hipóteses](#hipóteses-refutáveis)
· [Fora do escopo](#fora-do-escopo) · [Premissas](#premissas) · [Voz](#voz)

Este arquivo é sobre as palavras. O visual dá o impacto inicial; o texto é o que faz o cliente
confiar e é o que sobrevive quando ele reabre o arquivo sozinho, uma semana depois.

## A regra do título

Cada tela tem três camadas de texto e cada uma faz um trabalho diferente:

- **Sobrancelha** (`.eyebrow`) — diz o assunto. É o crachá da tela.
- **Título** (`h2`) — diz a **conclusão**. É a frase que a pessoa vai repetir.
- **Linha de apoio** (`.lede`) — diz o que a tela vai provar, ou como ler o que vem abaixo.

A diferença entre uma apresentação boa e uma medíocre está quase toda nesta linha:

| Fraco (anuncia o assunto) | Forte (entrega a conclusão) |
|---|---|
| Análise de conversão | A rede não perdeu público. Perdeu fechamento. |
| Decomposição do custo | A visita subiu 14%. A matrícula subiu 58%. |
| Retorno sobre investimento | Um número manda em toda a conta. |
| Cronograma do projeto | Diagnóstico, implementação e operação assistida. |
| Escopo não incluído | O que **não** está incluído. |
| Considerações finais | O projeto inteiro custa o equivalente a onze matrículas. |

O itálico dourado (`<em>`) vai no pedaço que carrega o argumento, não numa palavra decorativa.
Em "A visita subiu 14%. *A matrícula subiu 58%.*", o itálico é o contraste inteiro.

## Modelos por tela

**Capa.** Nome do cliente em corpo enorme; uma frase dizendo o que se quer resolver, no infinitivo
("Recuperar a conversão de visita em matrícula nas nove unidades"); e a tira de quatro números que
responde antes da pergunta: valor total, prazo, entrada, validade. Cliente relaxa quando vê o preço
na primeira tela — esconder preço até o fim cria uma tensão que atrapalha a escuta.

**Três números.** Três cartões, três números grandes, nem mais nem menos. Cada um com: rótulo curto
em mono, número, uma frase de 1 a 2 linhas que diz o que aquele número significa em gente ou em
dinheiro, e a conta. Escolha os três números que contam uma história em sequência — o estado de
hoje, a queda, e o que a queda custa.

**Decomposição.** Um gráfico de cascata à esquerda e, à direita, dois cartõezinhos: um com as
variações em números, outro com a leitura ("O preço da visita subiu, e a rede não controla isso. O
que ela controla é quantas dessas visitas fecham."). Separar o controlável do não controlável é o
movimento retórico mais forte de uma proposta de consultoria.

**Entregas.** Liste artefatos, não atividades. "Roteiro único de recepção, escrito e treinado" é
artefato; "alinhamento com a equipe" é fumaça. Feche com o que você **não** entrega — normalmente
"resultado garantido" — porque dizer isso em voz alta aumenta a confiança em tudo o mais.

**Fases.** Cartões que abrem um por vez conforme as setas avançam, com a linha do tempo acima
acendendo o trecho correspondente. Cada fase: nome, duração, 4 itens no máximo.

**Investimento.** Tabela com fase, prazo, condição de pagamento e valor; total destacado em corpo
grande. Embaixo, três cartõezinhos: entrada, validade e **o que o preço não tem** (taxa por
resultado, comissão, cobrança por hora). Terminar com o que não é cobrado desarma a desconfiança
sobre cobrança escondida.

**Fechamento.** De onde para onde o cliente vai (com faixa, não número único), quem fica no comando
no fim, e os passos concretos numerados com prazo real de início.

## O bloco da conta

Embaixo de todo número grande, um bloco `.conta` em mono, separado por um fio, com a aritmética:

```
143 ÷ 622 = 23,0%
32% − 23% = 9 pontos perdidos
9% × 622 = 56 matrículas por mês
```

Três linhas no máximo. Isso responde "de onde você tirou isso?" antes de a pergunta ser feita, e
posiciona você como alguém que fez a conta, não como alguém que escolheu um número que soa bem.

## Procedência

Toda tela com dado do cliente termina com uma linha `.src` em mono, no rodapé, com no máximo duas
linhas. O que ela precisa dizer:

- de onde vieram os números ("informados pela Rede Ímpeto na reunião de diagnóstico");
- se foram auditados ("**ainda não foram auditados pela Praxis**");
- o que é decisão sua e não medição ("pior caso entrega 60% do ganho escolhido — decisão da Praxis");
- quando a conta será conferida ("auditar é a primeira entrega da Fase 1").

Isso parece contraintuitivo numa peça de venda. Na prática é o contrário: a proposta que declara o
que não sabe é a que passa a impressão de ter olhado para os dados de verdade.

## Hipóteses refutáveis

Quando você ainda não diagnosticou, não escreva conclusões — escreva hipóteses, e para cada uma
escreva **o teste que a derrubaria**:

> **A visita chega sem hora marcada.** A pessoa aparece quando quer. Se cai no horário cheio, é
> atendida por quem estiver livre e às vezes espera.
>
> *Como confirma:* a Fase 1 abre a conversão por horário e por dia da semana, unidade a unidade. Se
> a conversão do horário cheio for igual à do horário vazio, a hipótese cai.

Dizer o que refutaria a própria tese é o que faz um diagnóstico soar honesto em vez de vendido. E
protege você: se a Fase 1 derrubar a hipótese, estava escrito desde o começo que podia acontecer.

Feche a tela dizendo o que acontece se todas caírem: "Se o diagnóstico apontar outra causa, o plano
da Fase 2 muda — com aprovação do cliente e **sem alteração de preço**."

## Fora do escopo

Seis itens, cada um com o motivo. Não é uma lista defensiva, é uma lista de clareza: o cliente
precisa saber o que vai ter que resolver por fora para não descobrir no meio do projeto.

Inclua sempre o item difícil — **garantia de resultado** — e escreva a versão honesta: "Quem atende
no balcão é o time da rede. A consultoria responde por método, por prazo e pelo que vai ser medido."

## Premissas

Duas colunas: **origem dos números** e **como a projeção foi construída**. Numeradas, corpo pequeno
(13,5 a 14,5px), e cada uma em uma ou duas linhas.

As que não podem faltar:

1. quem informou os dados e se foram auditados;
2. como a conta fecha consigo mesma (um teste de coerência: se X e Y são verdade, Z tem que dar
   isso — e dá);
3. a margem unitária, com a multiplicação explícita;
4. a taxa de saída ou evasão, e a permanência média que ela implica;
5. o que ficou **constante** na projeção (volume, verba, preço);
6. quando o ganho começa a contar;
7. o teto, e por que aquele teto;
8. a faixa de execução e que ela é decisão sua;
9. o que a conta ignora (reajuste, sazonalidade, receitas adicionais);
10. que nada ali é promessa.

Termine com uma linha que cita o valor escolhido no controle interativo — atualizada por JS, para a
tela de premissas nunca contradizer a tela do retorno.

## Voz

- **Frases curtas.** Em apresentação, frase longa não é lida, é pulada.
- **Português concreto.** "479 pessoas por mês entram na unidade e saem sem fechar" em vez de
  "oportunidades não convertidas".
- **Sem jargão de consultoria.** Nada de sinergia, alavancar, endereçar, acionável.
- **Gente e dinheiro, não percentual sozinho.** "9 pontos perdidos" vira real quando vira "56
  matrículas por mês que a rede já fez e não faz mais".
- **Segunda pessoa com parcimônia.** "vocês" funciona nas telas de compromisso (passos, validade);
  nas telas de diagnóstico prefira o nome do cliente, que é menos acusatório.
- **Números por extenso quando são a manchete**, algarismos quando são dado. "onze matrículas" na
  frase de efeito; "143 matrículas" na conta.
