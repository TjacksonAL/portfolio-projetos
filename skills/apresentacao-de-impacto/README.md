# Skill: Apresentação de impacto

Uma skill do Claude que produz apresentações, propostas comerciais e decks visuais de **tela cheia em HTML**: um palco fixo que escala sozinho, cartões de vidro, gráficos SVG desenhados em JavaScript, um controle interativo que recalcula as telas seguintes e uma verificação automática de encaixe antes da entrega.

O [SKILL.md](SKILL.md) é o que o Claude lê. Ele parte de uma ideia simples: o que separa uma apresentação incrível de um amontoado de slides é **foco**. Cada tela carrega uma ideia só, o título já entrega a conclusão ("A rede não perdeu público. Perdeu fechamento." em vez de "Análise de conversão"), o número que importa aparece em tamanho de manchete e a conta que o sustenta fica logo abaixo.

## O que a skill ensina

- **O caminho mais curto:** copiar o modelo, trocar os tokens de marca no CSS, substituir as telas de exemplo por conteúdo real e rodar o script de conferência antes de entregar.
- **O orçamento de altura.** O palco é fixo em 1600×900 e escala inteiro para caber na janela, o que sobra é uma área útil de 688 px de altura. A skill traz a tabela de quanto cada peça gasta e a regra de cortar conteúdo antes de encolher a fonte, porque apresentação é lida a três metros.
- **O arco narrativo** e a **disciplina de conteúdo** de cada tela.
- **Interatividade: um número manda em tudo.** Um controle único cujo valor recalcula as telas seguintes.
- **Armadilhas que custam caro.** Uma lista de erros reais, todos já ocorridos numa tela pronta.

## O que tem na pasta

| Arquivo | Para que serve |
|---|---|
| `SKILL.md` | As instruções da skill. |
| `references/sistema-visual.md` | Tokens, anatomia do cartão de vidro, escala tipográfica, luzes de fundo, moldura e animações. |
| `references/graficos.md` | Receitas de gráfico em SVG desenhado por JS: declínio, cascata, linha do tempo, curva de payback e medidores. |
| `references/interatividade.md` | Como montar a espinha aritmética única e fazer um controle recalcular as telas. |
| `references/narrativa.md` | Modelos de texto por tipo de tela, com título fraco contra título forte. |
| `assets/modelo.html` | Modelo de apresentação: palco, moldura, navegação, ícones, cartões, animações e motor de cálculo já funcionando. |
| `scripts/conferir_telas.mjs` | Abre a apresentação num navegador sem interface, mede cada tela, aponta o que estoura a área útil ou vaza pelas laterais, coleta erros de JavaScript e salva uma foto de cada tela. Sai com código 1 se achar problema, então serve em automação. Precisa de Node 18+ e Chrome ou Edge. |
| `evals/evals.json` | Casos de teste para medir a skill fora do caso que a originou. |
| `pacote/apresentacao-de-impacto.skill` | A skill empacotada, pronta para instalar. |

## Como instalar

- **Claude Code:** copie esta pasta para `~/.claude/skills/apresentacao-de-impacto/`.
- **Claude.ai:** envie o arquivo `pacote/apresentacao-de-impacto.skill` nas configurações de skills.

Depois é só pedir algo como: *"monta uma apresentação para este cliente, com estes números"*.
