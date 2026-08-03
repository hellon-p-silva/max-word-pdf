---
name: max-word-pdf
description: >-
  Transforma áudios (voz gravada no microfone do chat) em documentos PDF ou
  Word bem formatados, com base no que a pessoa fala e pede. Use SEMPRE que o
  usuário enviar um áudio/gravação de voz e quiser gerar um documento, ou
  quando disser coisas como "transforma esse áudio em PDF", "faz um documento
  do que eu falei", "gera um Word disso", "transcreve e formata esse áudio",
  "manda em PDF e Word", "gravei um recado, coloca num documento", "escreve
  isso aí como carta/proposta/relatório/ata". Vale para qualquer tipo de
  documento (carta, proposta, orçamento, ofício, relatório, ata, resumo,
  anotação) — a skill se adapta ao que foi pedido no áudio. Acione mesmo que a
  pessoa não diga a palavra "skill" e mesmo que só mande o áudio com um pedido
  curto de "faz um documento disso".
---

# max-word-pdf

Transforma o conteúdo falado num áudio enviado pelo chat em um documento **PDF**
e/ou **Word** limpo, organizado e pronto para usar.

## O que esta skill faz

Quando a pessoa grava um áudio, ela fala de forma natural: hesita, repete,
mistura instruções ("faz uma carta pro meu cliente...") com o conteúdo em si
("...avisando que o prazo mudou pra sexta"). Seu trabalho é **entender o que ela
quer**, **separar as instruções do conteúdo** e entregar um documento
profissional que reflita a intenção — não uma transcrição crua.

O fluxo é sempre o mesmo: entender → organizar → escolher formato → gerar →
entregar.

## Passo 1 — Entender o áudio

Você recebe o conteúdo do áudio como texto no chat. Leia tudo e identifique duas
camadas:

- **Instruções (meta):** o tipo de documento (carta, proposta, ata, relatório,
  anotação…), o destinatário, o tom (formal/informal), o formato desejado
  (PDF, Word ou ambos), título, quem assina, prazos.
- **Conteúdo:** a informação que efetivamente deve ir no documento.

Se a pessoa **não disse** que tipo de documento quer, infira pelo contexto (um
recado para um cliente → carta/comunicado; uma lista de tarefas faladas →
lista; ideias soltas sobre um assunto → texto/relatório com seções).

## Passo 2 — Organizar o conteúdo

Este é o coração da skill. A fala informal vira texto escrito bem-feito:

- **Limpe os vícios de fala:** remova "é...", "tipo", "né", "então assim",
  hesitações e repetições que não agregam.
- **Corrija a gramática e a pontuação**, mantendo o sentido e a voz da pessoa.
  Não deixe o texto robótico nem excessivamente formal se o tom pedido for
  simples.
- **Dê estrutura:** crie um título, divida em seções com cabeçalhos quando fizer
  sentido, transforme enumerações faladas ("primeiro... depois... e por fim")
  em listas.
- **Preserve tudo que é substantivo:** nomes, datas, valores, números de
  telefone, endereços, condições. Não invente fatos, dados nem números que a
  pessoa não falou.
- **Sinalize lacunas:** se faltar algo importante (nome do destinatário, data,
  valor), use um marcador claro como `[preencher: nome do cliente]` no
  documento **e** avise a pessoa no chat. Se a lacuna for central, é melhor
  perguntar antes de gerar.

O idioma do documento deve ser o mesmo do áudio (português, por padrão).

Escreva o conteúdo organizado em um arquivo Markdown (ex.: no diretório de
scratchpad ou de trabalho). O gerador entende esta sintaxe simples:

```markdown
# Título do documento
## Seção
### Subseção
Parágrafo normal, com **negrito** onde precisar.
- item de lista com marcador
1. item de lista numerada
---   (linha sozinha = quebra de página)
```

## Passo 3 — Escolher o formato

Respeite o que a pessoa pediu no áudio:

- Falou "em PDF" → só PDF (`--pdf`).
- Falou "em Word" / "editável" → só Word (`--docx`).
- Falou "os dois" / "PDF e Word" → ambos.
- **Não especificou** → gere **ambos** (é o mais útil e não custa nada); avise
  no chat que entregou PDF e Word e que é só pedir se quiser só um.

## Passo 4 — Gerar os arquivos

Use o script incluído. Ele produz PDF (com fonte Unicode do sistema, acentos
corretos e texto pesquisável) e/ou DOCX (Word editável):

```bash
python "<skill>/scripts/build_document.py" CONTEUDO.md --out "nome-do-arquivo" --docx --pdf
```

- `--out` define o nome base (sem extensão). Escolha um nome descritivo a partir
  do conteúdo (ex.: `carta-cliente-prazo`, `proposta-consultoria-marco`).
- Passe `--docx`, `--pdf`, ou ambos. Sem nenhuma flag, ele gera os dois.
- `--outdir` define a pasta de saída (padrão: pasta atual).

**Dependências:** o script precisa de `python-docx` (para Word) e `reportlab`
(para PDF). Se der erro de módulo faltando, instale uma vez:

```bash
pip install python-docx reportlab
```

## Passo 5 — Entregar

Envie os arquivos gerados para a pessoa com a ferramenta de envio de arquivos
(`SendUserFile`) para que ela possa baixar/abrir. Em seguida, num resumo curto,
diga o que foi gerado e aponte qualquer lacuna que ficou marcada como
`[preencher: ...]`.

## Exemplo

**Áudio (fala informal):**
> "Ó, faz pra mim uma carta pro seu João, é... avisando que a entrega que era
> pra segunda vai atrasar, tipo, vai só na quarta agora por causa do fornecedor.
> Pede desculpa, fala que a gente cobre o frete. Manda em PDF e Word. Assina como
> Maria, da Loja Central."

**Markdown organizado:**
```markdown
# Comunicado de Alteração de Prazo de Entrega

Prezado Sr. João,

Informamos que a entrega inicialmente prevista para **segunda-feira** precisará
ser reagendada para **quarta-feira**, em razão de um atraso do nosso fornecedor.

Pedimos desculpas pelo transtorno. Como forma de compensação, **a Loja Central
arcará com o custo do frete**.

Ficamos à disposição para qualquer esclarecimento.

Atenciosamente,
Maria
Loja Central
```

**Comando:**
```bash
python "<skill>/scripts/build_document.py" carta-joao.md --out carta-atraso-entrega --docx --pdf
```

Depois: enviar `carta-atraso-entrega.pdf` e `carta-atraso-entrega.docx` com
`SendUserFile` e confirmar no chat.

## Observações

- **Não invente conteúdo.** Formate e organize o que foi dito; não acrescente
  fatos, cláusulas ou valores que a pessoa não mencionou.
- Se a pessoa pedir **transcrição literal** ("transcreve exatamente o que eu
  falei"), aí sim mantenha as palavras dela, apenas com pontuação — não
  reorganize.
- Para documentos que exigem visual mais elaborado (logotipo, tabelas
  complexas, marca d'água, papel timbrado), as skills `docx` e `pdf` da
  Anthropic oferecem recursos mais ricos e podem ser usadas em conjunto — mas
  para o caso comum de "áudio vira documento", este script já entrega um
  resultado limpo e completo.
