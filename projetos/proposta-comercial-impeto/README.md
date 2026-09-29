# Proposta Comercial

Uma proposta de consultoria feita como apresentação de tela cheia, num único arquivo HTML, para ser aberta numa reunião e passada com as setas do teclado. A ideia central é que o cliente **arraste um controle e veja o retorno do projeto ser recalculado**, nessa tela e nas seguintes.

**[Ver no ar](https://proposta-comercial-impeto.vercel.app)** (use as setas do teclado)

![Capa da proposta](img/tela-01.png)

## O que tem nela

Doze telas, do diagnóstico ao preço fechado, ao retorno e ao próximo passo. Duas delas merecem atenção:

- **A capa** já abre com a queda da taxa de conversão de visita em matrícula (de 32% para 23% em 18 meses), desenhada em SVG, ao lado de quatro números-chave: valor, prazo, entrada e validade.
- **A tela do retorno** tem uma barra que o cliente arrasta até a conversão que considera possível. Matrículas a mais por mês, custo por matrícula, margem em 12 meses e retorno por real investido se recalculam juntos, com a faixa de pior e melhor caso, e o efeito segue para as telas seguintes.

![Tela do retorno, com o controle interativo](img/tela-08.png)

## Decisões de projeto

- **A confiança no número vem antes da beleza.** Cada valor da proposta se liga aos outros por uma aritmética explícita, e as premissas completas ficam numa tela própria.
- **Pensada para reunião.** Poucos elementos por tela, tipografia grande e contraste alto, para ser lida de longe numa TV ou projetor.
- **Arquivo único, sem dependências.** Todo o JavaScript e os gráficos em SVG estão dentro do HTML. Só as fontes vêm do Google Fonts.

## No celular e no tablet

A proposta é um palco horizontal 16:9, feito para tela grande, mas abre em celular e tablet:

- **Celular em pé:** aparece o aviso "Gire o celular", porque o palco ficaria pequeno demais para ler. Deitado, ele cabe inteiro na tela, e as telas passam deslizando o dedo para o lado.
- **Modo leve automático.** Em aparelhos de toque a página tira os efeitos gráficos mais caros: o vidro desfocado dos cartões, os filtros animados, as sombras dos títulos, o ruído e o movimento do fundo. O desenho continua o mesmo. No computador, o visual é o completo.

O modo leve existe por um motivo concreto. No iPhone, o Safari e o Chrome (que usam o mesmo motor de renderização, o WebKit) encerravam a página ao chegar à tela 6, "As fases". Ali, os três quadros juntavam vidro desfocado e um filtro de cor animado no mesmo elemento, dentro de um palco grande numa tela de alta densidade. Com o modo leve, a proposta passou a abrir sem travar em iPhone (Safari e Chrome) e em Android.

Para testar, o endereço aceita `?leve=1`, que força o modo leve em qualquer aparelho, e `?leve=0`, que força o visual completo. Em celular, `?leve=0` pode travar a página de novo.

## Dados

Empresa, cliente e números são **fictícios**, criados para fins didáticos.

## Stack

HTML, CSS e JavaScript puros, SVG gerado por código.
