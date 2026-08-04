# max-word-pdf

Skill para o **Claude** que transforma **áudios gravados no microfone do chat**
em documentos **PDF** e/ou **Word** bem formatados, com base no que a pessoa
fala e pede.

Você grava um áudio (ex.: *"faz uma carta pro meu cliente avisando que o prazo
mudou pra sexta, manda em PDF e Word"*) e a skill entende o pedido, organiza a
fala em texto escrito profissional e gera os arquivos prontos para baixar.

## O que ela faz

- **Entende o áudio** e separa as *instruções* (tipo de documento, destinatário,
  tom, formato) do *conteúdo* que vai no documento.
- **Organiza a fala informal** em texto limpo: remove vícios de fala ("é...",
  "tipo", "né"), corrige gramática e pontuação, cria título, seções e listas —
  sem inventar dados que não foram ditos.
- **Escolhe o formato** conforme o pedido: PDF, Word, ou ambos (padrão quando
  não é especificado).
- **Gera e entrega** os arquivos: PDF com acentos corretos e texto pesquisável,
  e Word (`.docx`) editável.

## Instalação

Copie a pasta da skill para o diretório de skills do seu Claude:

```bash
# Linux / macOS
cp -r max-word-pdf ~/.claude/skills/max-word-pdf
```

```powershell
# Windows (PowerShell)
Copy-Item -Recurse . "$env:USERPROFILE\.claude\skills\max-word-pdf"
```

> A skill deve ficar em `~/.claude/skills/max-word-pdf/` com o `SKILL.md` na
> raiz dessa pasta.

## Dependências

O gerador de documentos usa duas bibliotecas Python:

```bash
pip install -r requirements.txt
# ou:
pip install python-docx reportlab
```

## Uso

Dentro do Claude, basta gravar um áudio e pedir um documento — a skill dispara
sozinha. Exemplos de frases que acionam a skill:

- "transforma esse áudio em PDF"
- "faz um documento do que eu falei"
- "gera um Word disso"
- "gravei um recado, coloca numa carta"

### Uso direto do gerador (opcional)

O script também funciona sozinho, a partir de um arquivo Markdown:

```bash
python scripts/build_document.py conteudo.md --out meu-documento --docx --pdf
```

Sintaxe Markdown suportada: `#`/`##`/`###` (títulos), `-`/`*` (lista com
marcador), `1.` (lista numerada), `**negrito**`, `---` (quebra de página) e
**tabelas** no formato com `|` (ótimas para orçamentos — cabeçalho destacado,
linhas zebradas e alinhamento de colunas via `:--`, `--:`, `:--:`):

```markdown
| Item        | Qtd | Valor    |
| ----------- | --: | -------: |
| Consultoria |   1 | R$ 500   |
| Suporte     |   2 | R$ 300   |
```

As cores da tabela seguem as constantes `BRAND_HEADER_HEX`, `BRAND_ZEBRA_HEX`
e `BRAND_GRID_HEX` no topo de `scripts/build_document.py` — ajuste-as para a
identidade visual do cliente.

## Estrutura

```
max-word-pdf/
├── SKILL.md                  # instruções da skill (lidas pelo Claude)
├── scripts/
│   └── build_document.py     # gerador de PDF/Word a partir de Markdown
├── requirements.txt
└── README.md
```

## Licença

[MIT](LICENSE)
