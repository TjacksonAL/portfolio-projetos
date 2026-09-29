# Portfólio de projetos

Ferramentas de dados, painéis, apresentações interativas e skills construídos com o [Claude Code](https://claude.com/claude-code). Cada projeto tem sua própria pasta, com README, capturas de tela e, quando faz sentido, o código completo e uma demo que abre direto no navegador.

## Projetos

| Projeto | O que é | Stack | Ver |
|---|---|---|---|
| [**Central de Crédito e Cobrança**](projetos/central-credito-cobranca/) | Lê uma planilha de contas a receber e gera uma central de cinco telas: quem cobrar hoje, risco por cliente, previsão de caixa, régua de cobrança e relatório da semana. | Python (só biblioteca padrão), HTML, CSS e JS, SVG feito à mão | [Ver no ar](https://central-credito-cobranca.vercel.app) · [Código](projetos/central-credito-cobranca/) |
| [**Apura, painel eleitoral**](projetos/apura-painel-eleitoral/) | Painel de BI eleitoral de Alagoas, com resultados de 2022 e cenário de 2026, feito com dados abertos do TSE e do TRE-AL. | HTML de arquivo único, JavaScript, Chart.js, scripts em Python e PowerShell, Vercel | [Ver no ar](https://websitecampanha.vercel.app) · [Código](https://github.com/TjacksonAL/Website_Campanha) |
| [**Site de campanha**](projetos/site-campanha/) | Site estático de uma candidata a deputada estadual em Alagoas (eleições 2026). | HTML, CSS, Vercel | [Site no ar](https://walkiria22707.vercel.app/) |
| [**Proposta Comercial**](projetos/proposta-comercial-impeto/) | Proposta de consultoria em 12 telas, navegável por setas, em que o cliente arrasta um controle e o retorno do projeto é recalculado em todas as telas. | HTML, CSS e JS num arquivo só, SVG | [Ver no ar](https://proposta-comercial-impeto.vercel.app) · [Pasta](projetos/proposta-comercial-impeto/) |
| [**Painel financeiro Aurora**](projetos/painel-financeiro-aurora/) | Painel de caixa do semestre de uma distribuidora, com filtros por região e categoria, gráficos e uma aba de insights. | HTML, CSS e JS, SVG | [Ver no ar](https://painel-financeiro-aurora.vercel.app) · [Pasta](projetos/painel-financeiro-aurora/) |

## Skills do Claude

| Skill | O que faz |
|---|---|
| [**Apresentação de impacto**](skills/apresentacao-de-impacto/) | Ensina o Claude a criar apresentações e propostas de tela cheia em HTML, com palco fixo que escala sozinho, gráficos em SVG, um controle interativo que recalcula as telas seguintes e um script que confere se cada tela cabe antes da entrega. |

## Sobre os dados

- Empresas, clientes, valores e telefones dos projetos Aurora e da proposta comercial são **fictícios**, criados para fins didáticos.
- O painel Apura usa **dados públicos** do TSE (Portal de Dados Abertos) e do TRE-AL.
- Parte dos projetos nasceu de exercícios de um curso intensivo de Claude. Os materiais do curso, como apostilas e gabaritos, não fazem parte deste repositório.

## Como abrir os projetos

Os arquivos `.html` são páginas completas. No GitHub eles aparecem como código, por isso os links **Ver no ar** acima apontam para a Vercel, que os publica como página e se atualiza a cada mudança neste repositório. Para ver localmente, baixe o repositório e abra o arquivo no navegador.

---

Feito por [@TjacksonAL](https://github.com/TjacksonAL).
