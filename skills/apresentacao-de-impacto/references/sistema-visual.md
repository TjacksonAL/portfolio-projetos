# Sistema visual

Índice: [Tokens](#tokens) · [Trocar a identidade](#trocar-a-identidade) · [Palco e moldura](#palco-e-moldura)
· [Luzes de fundo](#luzes-de-fundo) · [Cartão de vidro](#cartão-de-vidro) · [Tipografia](#tipografia)
· [Números](#números) · [Animação](#animação) · [Ícones](#ícones)

O sistema inteiro está montado em `assets/modelo.html`. Este arquivo explica **por que** cada peça é
assim, para você conseguir mudar sem quebrar.

## Tokens

```css
:root{
  --bg-0:#070A13;                      /* fundo do palco */
  --txt:#EFF2F9; --txt-2:#A9B3CA; --txt-3:#6E7A98;   /* três níveis, nunca mais */
  --gold:#F5B24A;  --rose:#FF6B5A;  --mint:#39DEAE;
  --violet:#8B6BFF; --blue:#5B8DEF;
  --grad-brass: /* branco → creme → âmbar → âmbar escuro */;
  --grad-oxide: /* branco → rosa → coral → coral escuro */;
  --grad-moss:  /* branco → menta clara → menta → menta escura */;
  --grad-cool:  /* branco → cinza-azulado → cinza médio */;
  --glass: linear-gradient(158deg, rgba(255,255,255,.085), rgba(255,255,255,.022) 58%, rgba(255,255,255,.045));
}
```

Três níveis de texto bastam: `--txt` para o que precisa ser lido, `--txt-2` para o corpo, `--txt-3`
para rótulo e procedência. Um quarto nível só cria indecisão.

Os gradientes existem porque **número grande com cor chapada fica morto**. O degradê de branco para
a cor faz o número parecer iluminado por cima, como metal — é o que dá a sensação de material caro.

### Semântica da cor

| Acento | Significa | Telas típicas |
|---|---|---|
| `--rose` | problema, perda, risco | diagnóstico, hipóteses, o que não está incluído |
| `--gold` | dinheiro, preço, valor | custo, investimento, totais |
| `--mint` | ganho, resultado, retorno | projeção, payback, fechamento |
| `--violet` | processo, tempo, fases | cronograma, etapas |
| `--blue` | entregas, método, premissas | escopo, o que fica na mão do cliente |

Aplique com as classes `.ac-rose`, `.ac-gold`, `.ac-mint`, `.ac-violet`, `.ac-blue` no cartão. Elas
só trocam a variável `--ac`, e tudo dentro do cartão (fio de luz, ícone) segue junto.

## Trocar a identidade

Para outra marca, mexa em três lugares e mais nada:

1. **As cinco cores de acento** no `:root`. Mantenha a lógica semântica acima; se a marca não tem
   cinco cores, repita — é melhor ter dois acentos coerentes que cinco aleatórios.
2. **Os quatro gradientes.** A receita é sempre: `#FFFFFF 0%` → cor clara `34%` → cor da marca `74%`
   → cor escura `100%`, na diagonal `145deg`. Só troque as cores do meio.
3. **As três fontes.** Uma serifada fina para títulos e números, uma sem serifa para o corpo, uma
   mono para rótulo, conta e procedência. O contraste entre as três é metade do charme.

O fundo deve continuar muito escuro (luminosidade abaixo de 10%). Todo o sistema — vidro, brilho,
degradês, luzes — depende de fundo escuro. Em fundo claro nada disso funciona e você precisa de
outro sistema.

## Palco e moldura

```css
#stage{ position:absolute; top:50%; left:50%; width:1600px; height:900px;
        transform-origin:50% 50%; overflow:hidden }
```

```js
function fit(){
  var s = Math.min(innerWidth/1600, innerHeight/900);
  stage.style.transform = 'translate(-50%,-50%) scale(' + s + ')';
}
```

Tudo é posicionado em pixels de um palco de 1600×900 e o palco inteiro escala. A vantagem: você
projeta uma vez e funciona em notebook, monitor ultrawide, projetor e PDF. A desvantagem: não existe
rolagem, então conteúdo que não cabe simplesmente vaza. Daí o orçamento de altura e o script de
conferência.

A moldura (`.frame`) fica fora das telas, é `pointer-events:none` e contém quatro coisas fixas:

- **marca**, topo esquerdo, mono com espaçamento largo;
- **identificação do documento**, topo direito, duas linhas;
- **rótulo da tela**, rodapé esquerdo, vindo de `data-label`;
- **contador e régua**, rodapé direito, com um traço por tela — o traço da tela atual cresce e
  acende. É o que dá ao cliente a noção de quanto falta sem precisar de número de slide gritante.

Sobreposições do palco, em ordem: grade de 100px mascarada em radial, vinheta escura nas bordas, e
uma camada de **grão** em SVG com 5% de opacidade. O grão é o detalhe mais barato e mais eficaz do
sistema: sem ele os degradês mostram faixas (banding) e a tela parece digital demais.

## Luzes de fundo

Três camadas de gradientes radiais, uma por humor, empilhadas e com transição de opacidade de 1,25s:

- `.mood-neutral` — violeta e âmbar, equilibrada;
- `.mood-problem` — coral e magenta, tensa;
- `.mood-gain` — menta e âmbar, aberta.

Cada tela declara `data-mood` e a troca acontece na transição. O efeito é sutil de propósito: o
cliente não percebe conscientemente, mas a sala fica mais tensa nas telas de problema e abre nas de
resultado. Nas telas interativas, a luz deve seguir **o estado da conta**, não a tela — quando o
cliente arrasta o controle para um patamar ruim, o fundo esquenta. Esse é o momento em que as
pessoas percebem que a apresentação está viva.

Um `#glow` envolvendo as três camadas faz uma deriva lenta de 44s (`@keyframes drift`), para o fundo
nunca ficar completamente parado.

## Cartão de vidro

```css
.card{ --ac:var(--gold); position:relative;
       background:var(--glass); border:1px solid rgba(255,255,255,.11); border-radius:22px;
       backdrop-filter:blur(26px) saturate(165%); box-shadow:var(--shadow); padding:28px 32px }
.card::after{ /* fio de luz no topo */
       content:""; position:absolute; left:20px; right:20px; top:-1px; height:2px;
       background:linear-gradient(90deg,transparent,var(--ac) 20%,var(--ac) 80%,transparent);
       filter:drop-shadow(0 0 6px var(--ac)) drop-shadow(0 0 16px var(--ac)) }
```

Três coisas fazem o cartão parecer vidro de verdade e não um retângulo cinza:

1. **O gradiente diagonal** em `--glass`, mais claro no topo esquerdo — simula luz batendo de cima.
2. **`backdrop-filter` com saturação em 165%**, que puxa a cor da luz de fundo para dentro do vidro.
   É por isso que o mesmo cartão parece diferente em telas de humor diferente.
3. **O fio de luz no topo**, com duas sombras de brilho. É o que dá identidade ao cartão e o que
   carrega a cor semântica.

O ícone fica absoluto no canto superior direito com 62% de opacidade. Para o texto não passar por
baixo dele, o primeiro e o segundo filho do cartão recebem `padding-right:56px`. Se você criar uma
regra mais específica para aquele elemento (uma classe de tela, por exemplo), **reponha o
`padding-right`** — essa é a quebra mais comum do sistema.

Variantes: `.card.tight` (padding menor, para a tira da capa) e `.card.flat` (sem vidro, quando você
quer a estrutura sem a moldura visual).

## Tipografia

| Classe | Uso | Tamanho |
|---|---|---|
| `h1` | nome do cliente na capa | 112px, serifada peso 200 |
| `h2` | título da tela, com `<em>` dourado no argumento | 56px, serifada peso 200 |
| `.eyebrow` | assunto da tela, acima do título | 14px mono, maiúsculas, espaçamento .24em |
| `.lede` | linha de apoio abaixo do título | 21px, peso 300 |
| `.qtxt` | corpo dentro de cartão | 17,5px, peso 300, entrelinha 1,6 |
| `.lbl` | rótulo dentro de cartão | 13px mono, maiúsculas |
| `.conta` | a conta que sustenta o número | 13,5px mono, com borda no topo |
| `.src` | procedência, rodapé da tela | 14px mono |

Os títulos usam `background-clip:text` com `--grad-cool` para ganhar o mesmo brilho metálico dos
números, e `filter:drop-shadow` para descolar do fundo.

A hierarquia inteira se sustenta em **três famílias e três pesos**. Resista a acrescentar um quarto
tamanho para resolver um caso específico — é sempre melhor cortar texto.

## Números

```html
<div class="n g figure">23,0<small>%</small></div>
```

- `.num` / `.n` — serifada peso 200, `font-variant-numeric: tabular-nums`, entrelinha 0,86.
- `.g` / `.gr` / `.gg` — aplicam o gradiente âmbar / coral / frio dentro do texto.
- `<small>` para unidade e símbolo, sempre menor que o número.
- `.figure` liga a animação de entrada (sobe, desfoca e assenta).

`tabular-nums` importa mais do que parece: sem ele, os números dançam quando o valor muda no
controle interativo.

## Animação

A regra geral é **entrada escalonada, nunca simultânea**. Cada elemento com `.anim` recebe
`style="--d:.18s"` e entra um pouco depois do anterior. O olho acompanha a ordem em que você quer
que a tela seja lida.

| Classe | O que faz | Onde usar |
|---|---|---|
| `.anim` + `--d` | sobe 18px e aparece | qualquer bloco da tela |
| `.figure` | sobe, desfoca e assenta | números grandes |
| `.after` | aparece depois do número | texto que explica o número |
| `.draw` + `--len` | desenha o traço do começo ao fim | linhas de gráfico |
| `.fadein` | só aparece | eixos, rótulos, áreas |
| `.growx` / `.growy` | cresce da esquerda / da base | barras |
| `.pop` | surge com um leve exagero | pontos e marcadores |

`.draw` precisa saber o comprimento do próprio traçado. Meça com `getTotalLength()` e escreva em
`--len` — no carregamento, para SVGs escritos à mão, e logo depois de gerar o HTML, para SVGs
desenhados por JS.

Tudo isso é desligado em bloco quando o sistema pede movimento reduzido:

```css
@media (prefers-reduced-motion:reduce){
  .anim{transition:none!important;opacity:1;transform:none}
  *,*::before,*::after{animation:none!important}
}
```

Uma apresentação que ignora essa preferência passa mal em sala com alguém sensível a movimento, e o
custo de respeitar é uma regra de CSS.

## Ícones

Não use biblioteca de ícones. O modelo traz um dicionário de caminhos SVG de traço fino
(`var ICONES = {...}`) e injeta em qualquer elemento com `data-ico="nome"`. O ícone herda a cor do
acento do cartão via `currentColor`, então ele sempre combina.

Ícones disponíveis: `coin`, `calendar`, `clock`, `funnel`, `down`, `up`, `percent`, `chart`,
`users`, `drop`, `search`, `tune`, `loop`, `check`, `shield`, `nope`, `gauge`.

Para acrescentar um, desenhe num `viewBox` de 24×24 com traço de 1,35 e adicione ao dicionário. Não
use ícone preenchido — a coerência do conjunto vem de todos serem de traço.
