# Interatividade

Índice: [Por que vale a pena](#por-que-vale-a-pena) · [A espinha aritmética](#a-espinha-aritmética)
· [O controle](#o-controle) · [Estados e veredito](#estados-e-veredito) · [Propagação](#propagação)
· [Varredura de entrada](#varredura-de-entrada) · [Teclado](#teclado-sem-briga) · [Adaptar](#adaptar-a-outro-assunto)

## Por que vale a pena

Numa proposta, a objeção real quase nunca é o preço — é "será que dá para chegar nesse número?".
Uma tela onde o cliente arrasta um controle e vê a conta inteira se refazer resolve isso sem
discussão: ele escolhe o patamar em que acredita e a apresentação mostra o que acontece. O número
passa a ser dele.

Construa **uma** tela dessas. Duas já viram brinquedo e o cliente para de ouvir.

## A espinha aritmética

Um bloco único, no topo do script, com todas as constantes e todas as funções derivadas. Nenhum
número calculado é escrito no HTML — se ele aparece em duas telas, vem da mesma função nas duas.

```js
/* ===== constantes: tudo que veio do cliente ===== */
var VISITAS=622, MENS=149, MARGC=0.62, EVASAO=0.04;
var HOJE=23.0, TETO=32.0, PADRAO=28.0, INICIO_GANHO=4, JANELA=12, PRECO=24800;
var MARGEM_ALUNO = MENS*MARGC;

/* ===== custo mês a mês, do jeito que será faturado ===== */
var CUSTO_MES = new Array(JANELA+1).fill(0);
CUSTO_MES[1] = PRECO*0.30 + 5400*0.70;
for(var m=2;m<=3;m++) CUSTO_MES[m] += (12200/2)*0.70;
for(var m=4;m<=12;m++) CUSTO_MES[m] += 800*0.70;

/* ===== funções derivadas ===== */
function pontosEf(t,f){ return Math.min((t-HOJE)*(f||1), TETO-HOJE); }
function ganhoMes(t,f,m){ /* estoque com evasão corroendo cada safra */ }
function serie(t,f){ /* resultado líquido acumulado, mês a mês */ }
```

Dois cuidados que separam uma projeção séria de uma otimista:

- **Teto.** Nem o melhor cenário projeta acima do patamar que o cliente já praticou algum dia. Isso
  é o `Math.min(..., TETO-HOJE)` e é a primeira coisa que um cliente cético procura.
- **Atraso do ganho.** O ganho só começa a contar quando a implementação termina. Os meses iniciais
  entram com ganho zero e custo cheio — é isso que faz a curva afundar antes de subir, e é essa
  barriga que dá credibilidade ao gráfico.

### Limiar calculado, nunca escrito à mão

O ponto a partir do qual o projeto se paga é uma função do preço e do modelo. Calcule por bisseção:

```js
function limiar(f){
  var lo=HOJE, hi=TETO;
  for(var k=0;k<60;k++){
    var md=(lo+hi)/2;
    if(serie(md,f)[JANELA-1] < 0) lo=md; else hi=md;
  }
  return hi;
}
var LIM_PROV = limiar(1), LIM_PIOR = limiar(0.6);
```

Depois posicione a marca na barra com o valor calculado:

```js
var pos = ((LIM_PIOR-HOJE)/(TETO-HOJE)).toFixed(5);
slider.style.setProperty("--pt", pos);
```

No dia em que o preço mudar, a marca anda sozinha. Marca escrita à mão no CSS é uma mentira à espera
de acontecer.

## O controle

Um `input[type=range]` invisível por cima do desenho. O visual é todo seu: trilho, preenchimento,
agulha e o número grande que anda junto com ela.

```css
.slider input[type=range]{ position:absolute; opacity:0; width:100%; height:62px; cursor:ew-resize }
.slider .fill  { width: calc(16px + var(--p,0) * (100% - 32px)) }
.slider .needle{ left:  calc(16px + var(--p,0) * (100% - 32px)) }
.slider .bigval{ left:  calc(16px + var(--pv,0) * (100% - 32px)); transform:translateX(-50%) }
```

`--p` é a posição real; `--pv` é a mesma posição limitada entre 0,045 e 0,955, para o número grande
não sair da tela nas pontas. Esse detalhe é a diferença entre "funciona" e "funciona bem".

Escreva na escala, embaixo da barra, três marcas com nome: **onde está hoje**, **o limiar** e **o
teto já praticado**. Sem esses três textos o cliente não sabe o que está arrastando.

## Estados e veredito

O resultado tem três estados, e cada um tem cor, degradê, humor de fundo e frase:

| Estado | Condição | Cor | Frase |
|---|---|---|---|
| vermelho | abaixo do limiar do cenário provável | coral | não cobre o investimento nem no provável; faltam N pontos |
| âmbar | entre os dois limiares | âmbar | cobre no provável, mas não no pior caso; faltam N pontos |
| verde | acima do limiar do pior caso | menta | cobre até no pior caso |

Sempre diga **quanto falta** para o próximo estado. "Faltam 1,8 ponto de conversão" é acionável;
"insuficiente" é só um julgamento.

Ao trocar de estado, mude junto: cor do preenchimento e da agulha, degradê do número grande, fundo e
borda do veredito, gradiente dos medidores e **a luz de fundo da tela**. Quando tudo vira coral ao
mesmo tempo, a sala inteira entende sem você dizer nada.

## Propagação

Uma função `pintar(valor, leve)` e uma `propagar(valor)`:

```js
function pintar(c, leve){
  taxa = c;
  /* ... posição, cores, número grande, veredito, medidores ... */
  if(leve) return;      // durante a animação, só o que é barato
  propagar(c);          // ao soltar, o resto da apresentação
}

function propagar(c){
  /* tela do payback: título dinâmico + gráfico redesenhado */
  /* tela de premissas: a linha que cita a taxa escolhida */
  /* tela de fechamento: faixa de chegada, custo e margem */
}
```

O argumento `leve` existe porque redesenhar um gráfico a 60 quadros por segundo trava. Durante o
arrasto e a varredura, atualize só o que é barato; ao soltar, refaça tudo.

O efeito que impressiona: o cliente arrasta na tela 8, avança, e as telas 9, 11 e 12 **já estão
falando do número que ele escolheu**. Vale mais que qualquer animação.

## Varredura de entrada

Quando a tela interativa entra, não mostre o valor pronto: faça a agulha correr do valor de hoje até
o valor padrão em um segundo, com todos os números se resolvendo junto.

```js
function sweep(){
  var alvo = parseFloat(range.value), t0 = null, dur = 1000;
  slider.classList.add("sweeping");   // desliga as transições de CSS durante a varredura
  raf = requestAnimationFrame(function passo(t){
    if(t0===null) t0=t;
    var k = Math.min(1,(t-t0)/dur), e = 1-Math.pow(1-k,3);
    pintar(HOJE + (alvo-HOJE)*e, k<1);
    if(k<1) raf=requestAnimationFrame(passo);
    else { raf=null; slider.classList.remove("sweeping"); pintar(alvo,false); }
  });
}
```

Isso ensina o controle sem instrução: a pessoa vê a barra se mover e entende na hora que pode mexer.
Cancele a varredura ao primeiro toque no controle, para não brigar com o cliente.

## Teclado sem briga

A apresentação navega com as setas, e o controle também usa as setas. Quando o foco está no
controle, deixe as setas para ele:

```js
document.addEventListener("keydown", function(e){
  var ae = document.activeElement;
  if(ae && ae.id === "range" && e.key.indexOf("Arrow") === 0) return;
  /* ... navegação ... */
});
```

E tire o foco do controle assim que o cliente soltar (`blur()` no `pointerup`), senão a próxima seta
move a barra em vez de trocar de tela — e isso confunde no meio da reunião.

## Adaptar a outro assunto

A estrutura vale para qualquer proposta com um número-chave. Troque a variável e o resto segue:

| Negócio | Variável do controle | Limiar |
|---|---|---|
| academia, escola, clínica | conversão de visita em matrícula | onde a margem cobre o projeto |
| e-commerce | taxa de conversão do checkout | idem |
| indústria | disponibilidade da linha, refugo | onde a economia cobre o investimento |
| serviço recorrente | cancelamento mensal | onde a retenção paga o esforço |
| time comercial | taxa de fechamento por proposta | idem |

Se o projeto não tem um número-chave único, não force a tela interativa — ela só funciona quando
existe **uma** alavanca que manda no resultado. Nesse caso, prefira a tela de medidores com a faixa
de cenários.
