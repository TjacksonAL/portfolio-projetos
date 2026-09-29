# Gráficos

Índice: [Princípios](#princípios) · [Escolher a forma](#escolher-a-forma) · [Esqueleto comum](#esqueleto-comum)
· [Linha de declínio](#1-linha-de-declínio) · [Cascata](#2-cascata) · [Linha do tempo](#3-linha-do-tempo)
· [Curva de payback](#4-curva-de-payback) · [Medidores](#5-medidores) · [Interpolação monótona](#interpolação-monótona)
· [Honestidade](#honestidade-no-desenho)

Todos os gráficos são SVG. Os simples ficam escritos à mão no HTML; os que dependem de dados que
mudam são desenhados por JS numa função por gráfico, chamada quando a tela entra.

## Princípios

**O gráfico se desenha enquanto você fala.** Um gráfico que já está pronto quando a tela aparece é
uma figura; um que se desenha da esquerda para a direita é um argumento. Sincronize a entrada das
marcas com a chegada do traço: se a linha leva 1,5s para ir de x=20 a x=1150, o rótulo que fica em
x=600 aparece quando a linha passa por lá, não antes.

```js
var tAt = function(x){ return 0.95 + 1.5*(x-20)/1130; };  // atraso da marca em x
```

**Menos eixo, mais rótulo.** Linhas de grade em `rgba(255,255,255,.06)` a `.075`, quase invisíveis.
Em vez de eixo Y cheio de marcações, escreva o valor final direto na ponta da linha. O cliente
quer saber onde chega, não ler uma régua.

**Cor com o mesmo significado do resto da apresentação.** Coral para a curva do pior caso, menta
para o melhor, degradê coral→âmbar→menta para o cenário provável — que atravessa do problema ao
resultado, e é exatamente o que a curva conta.

**Texto do gráfico em mono.** Números e rótulos dentro do SVG sempre na fonte mono, entre 13 e 16px.
Isso os separa visualmente do texto da tela e mantém dígitos alinhados.

## Escolher a forma

| A pergunta da tela | Forma |
|---|---|
| Qual item pesa mais? Onde está concentrado? | **barras horizontais** — a receita que serve à maioria |
| Para onde isso foi, ao longo do tempo? | linha com área |
| De onde veio essa diferença entre dois números? | cascata |
| O que acontece quando, e por quanto tempo? | linha do tempo em barras |
| Quando o acumulado passa de negativo para positivo? | curva acumulada com faixa de cenários |
| Quais são os quatro números que resumem isso? | medidores (não é gráfico, e quase sempre é melhor) |

Antes de desenhar, pergunte se um número grande não resolveria melhor. Gráfico com três pontos é
uma tabela mal desenhada.

## 0. Barras horizontais — comece por aqui

O modelo traz `barrasH()` pronto, e ela resolve ranking, comparação, composição e orçamento. Numa
apresentação de dezessete telas, essa única função desenhou sete dos oito gráficos. Use-a antes de
escrever qualquer desenho novo.

```js
barrasH("svg-x", { W:900, H:300, L:230, R:170, T:20, bh:34, dados:[
  {rot:"Item mais pesado", val:48, cor:"var(--rose)", txt:"48%", forte:true, nota:"o que puxa o total"},
  {rot:"Segundo item",     val:22, cor:"var(--gold)", txt:"22%"}
]});
```

- `rot` é o rótulo à esquerda, `txt` o valor à direita, `nota` a linha de apoio embaixo do valor.
- `forte: true` clareia rótulo e valor — use na linha que carrega o argumento, no máximo duas.
- `fraca: true` esmaece a barra, para itens de contexto.
- A maior barra define a escala; as outras são proporcionais.
- Horizontal, não vertical: rótulo de categoria cabe por extenso, sem girar texto.

**Reserve margem para o texto das pontas.** `L` precisa caber o rótulo mais longo e `R` precisa caber
o valor **mais a nota** da barra mais comprida — que é justamente a que sobra menos espaço. Regra
prática: `R ≥ 12 + (caracteres da nota mais longa × 6,2px)`. Errar isso corta a nota sem aviso; foi
o defeito mais frequente no uso real desta skill, e hoje o verificador acusa com o lado a corrigir.

## Esqueleto comum

```js
function desenharX(){
  var s = document.getElementById("svg-x"); if(!s) return;
  s.innerHTML = "";
  var W=918, H=370, L=92, R=16, T=24, B=82;          // margens internas
  var px = function(v){ return L + (W-L-R) * v/12; };  // valor → x
  var py = function(v){ return T + (H-T-B) * (1 - v/topo); };  // valor → y
  // grade, depois marcas, depois dados, depois rótulos
}
```

Regras que evitam retrabalho:

- **Case a largura do `viewBox` com a largura real de render.** Se o cartão tem 982px de largura
  interna, use `viewBox="0 0 982 ..."`. Quando as duas diferem, o texto dentro do SVG escala junto e
  some. Essa é a causa número um de gráfico ilegível.
- Use `class="graf"` (`display:block; width:100%; height:auto`) para o SVG respeitar a proporção.
- Desenhe na ordem: grade → faixas de fundo → áreas → linhas → pontos → rótulos. Quem vem depois
  fica por cima.
- Helpers `svgel(tag, attrs)` e `txt(x, y, texto, attrs)` estão no modelo; use-os em vez de
  concatenar string, exceto nos gráficos grandes onde montar `innerHTML` de uma vez é mais rápido.

## 1. Linha de declínio

Para "isto piorou ao longo do tempo". Área com degradê que some para baixo, traço com degradê
âmbar→coral, máscara suave nas pontas, e os dois extremos marcados com ponto e rótulo.

Peças que fazem a diferença:

```xml
<linearGradient id="cvFill" x1="0" y1="0" x2="0" y2="1">
  <stop offset="0%" stop-color="#F5B24A" stop-opacity=".28"/>
  <stop offset="100%" stop-color="#FF6B5A" stop-opacity="0"/>
</linearGradient>
<!-- máscara que apaga as pontas, para a linha não "começar" nem "terminar" de repente -->
<linearGradient id="cvSoft" x1="0" y1="0" x2="1" y2="0">
  <stop offset="0%" stop-color="#000"/><stop offset="9%" stop-color="#fff"/>
  <stop offset="88%" stop-color="#fff"/><stop offset="100%" stop-color="#000"/>
</linearGradient>
```

O traço leva `class="draw glowline"`, e os rótulos das pontas entram com `.fadein` atrasado para
coincidir com a chegada da linha. O ponto final em coral com `.pop` fecha a narrativa.

Se você só tem dois pontos medidos (começo e fim), **não invente o caminho entre eles**. Ou desenhe
uma reta tracejada, deixando claro que é interpolação, ou desenhe a oscilação e escreva na tela que
o caminho é ilustrativo. Ver [Honestidade](#honestidade-no-desenho).

## 2. Cascata

Para "de onde veio esta diferença". A forma mais persuasiva da apresentação inteira, porque separa
o que o cliente controla do que ele não controla.

```js
var CASC = [
  ["Há dezoito meses", 137.50, "base", ""],
  ["+ visita mais cara", 18.75, "sobe", "23% da alta"],
  ["+ conversão mais baixa", 53.80, "sobe", "67% da alta"],
  ["+ as duas juntas", 7.34, "misto", "9% da alta"],
  ["Hoje", 217.39, "base", ""]
];
```

Cada barra "sobe" começa onde a anterior terminou. As duas barras "base" (início e fim) vão do zero,
em âmbar; as intermediárias em coral.

**Quando a decomposição tem duas variáveis, existe um termo cruzado** — o efeito de as duas mudarem
ao mesmo tempo. Ele não pertence a nenhum dos dois lados, então desenhe em cinza e nomeie "as duas
juntas". Jogar esse termo dentro de uma das parcelas infla aquela parcela e alguém vai conferir a
conta.

A conta fecha assim, para custo = preço ÷ taxa:

```
parcela do preço  = preço_novo ÷ taxa_antiga − base
parcela da taxa   = preço_antigo ÷ taxa_nova − base
termo cruzado     = total − base − parcela do preço − parcela da taxa
```

Escreva o peso de cada parcela ("67% da alta") embaixo da barra. É o que o cliente vai repetir para
o sócio dele.

Anime com `transform: scaleY` a partir da base de cada barra, escalonando 440ms entre elas — as
barras sobem em sequência e a soma fica visível.

## 3. Linha do tempo

Barras arredondadas lado a lado, uma por fase, proporcionais à duração. Cor por fase, opacidade
baixa para as inativas e alta para a fase aberta — o que permite ligar a linha do tempo aos cartões
de fase que abrem um por vez.

Duas armadilhas:

- A primeira fase costuma ser curta e o nome dela é mais largo que a barra, invadindo o nome da
  segunda. **Escalone os rótulos em duas alturas** (`y-13` e `y-31`).
- Rótulo de duração e rótulo de mês competem pela mesma linha de base. Ponha duração à esquerda de
  cada trecho e mês à direita do eixo, com `text-anchor="end"`.

Anime com `transform: scaleX(0) → scaleX(1)` a partir da esquerda, 620ms, escalonando 380ms.

## 4. Curva de payback

A mais trabalhosa e a que mais vende. São três curvas (pior, provável, melhor caso), uma faixa
preenchida entre a pior e a melhor, e a zona de payback destacada.

Estrutura:

- **Zona de payback**: retângulo em menta translúcida entre o cruzamento do melhor caso e o do pior,
  com as duas bordas tracejadas e o rótulo `PAYBACK` no topo. É o "entre o 6º e o 8º mês" desenhado.
- **Faixa entre cenários**: `path` fechado indo pela curva do melhor caso e voltando pela do pior.
- **Curvas**: as extremas em branco a 34% de opacidade e 1,6 de espessura; a provável com 3,4 de
  espessura, degradê coral→âmbar→menta e `class="glowline"`.
- **Pontos de cruzamento**: círculo âmbar em cada lugar onde a curva cruza o zero, com brilho.
- **Linha do zero** com `growx` e o rótulo `R$ 0` à direita.
- **Valores de chegada** à direita de cada curva, com o nome do cenário embaixo em 13px. Se as três
  pontas ficarem muito próximas, empurre-as para manter 34px entre elas — o modelo faz isso.
- **Barras das fases** embaixo do eixo, alinhadas aos meses, com um traço vertical ligando cada
  barra ao seu rótulo.

Reserve entre 200 e 210px de margem direita para os rótulos de chegada. Sem isso eles cortam.

## 5. Medidores

Quatro cartõezinhos em linha, cada um com rótulo em mono, número grande em serifada com gradiente e
uma linha de apoio em mono com a faixa. Não é gráfico, mas quase sempre comunica melhor que um.

Mantenha os rótulos curtos: o rótulo divide espaço com o ícone no canto, e um rótulo de três linhas
empurra todos os quatro medidores para baixo. Se não couber em duas linhas, encurte o rótulo e
mande o resto para a linha de apoio.

## Interpolação monótona

Curvas acumuladas ficam duras quando ligadas por segmentos retos e mentirosas quando suavizadas com
spline comum — a spline inventa picos que os dados não têm, e num gráfico de payback um pico falso
move o ponto de equilíbrio.

Use interpolação cúbica monótona (Fritsch–Carlson): ela suaviza sem nunca ultrapassar os valores dos
pontos vizinhos.

```js
function mono(xs, ys){
  var n=xs.length, dx=[], m=[], t=new Array(n), k;
  for(k=0;k<n-1;k++){ dx[k]=xs[k+1]-xs[k]; m[k]=(ys[k+1]-ys[k])/dx[k]; }
  t[0]=m[0]; t[n-1]=m[n-2];
  for(k=1;k<n-1;k++){
    if(m[k-1]*m[k]<=0) t[k]=0;
    else{ var w1=2*dx[k]+dx[k-1], w2=dx[k]+2*dx[k-1]; t[k]=(w1+w2)/(w1/m[k-1]+w2/m[k]); }
  }
  return {xs:xs, ys:ys, t:t};
}
```

E encontre o cruzamento do zero **na curva desenhada**, por bisseção, não no vetor de pontos:

```js
function cruzamento(c){
  var a=c.xs[0], b=c.xs[c.xs.length-1], ant=avaliar(c,a), par=null;
  for(var k=1;k<=600;k++){
    var x=a+(b-a)*k/600, v=avaliar(c,x);
    if(ant<0 && v>=0){ par=[a+(b-a)*(k-1)/600, x]; break; }
    ant=v;
  }
  if(!par) return null;
  var lo=par[0], hi=par[1];
  for(var j=0;j<48;j++){ var md=(lo+hi)/2; if(avaliar(c,md)<0) lo=md; else hi=md; }
  return (lo+hi)/2;
}
```

Assim o ponto marcado no desenho é exatamente onde a linha cruza o zero. Se você calcular pelo vetor
e desenhar pela curva, o ponto fica ao lado da linha e alguém vai reparar.

## Honestidade no desenho

Estas regras protegem a proposta quando alguém for conferir:

- **Não desenhe pontos que você não mediu.** Dois extremos informados não viram uma série mensal.
  Se precisar do caminho, deixe explícito na tela que é ilustração.
- **Não corte o eixo Y** para fazer uma queda parecer maior. Se a queda é pequena, o argumento tem
  que ser outro.
- **Escreva o intervalo, não só o número bonito.** Projeção com faixa (pior/provável/melhor) é mais
  convincente que número único, porque mostra que você pensou no que pode dar errado.
- **Diga o que é decisão sua.** "Pior caso entrega 60% do ganho" é escolha do consultor, não medição
  do cliente. Diga isso na tela de premissas e no rodapé do gráfico.
