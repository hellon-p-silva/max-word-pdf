---
description: Gera um orçamento da MasterLeds no template oficial (logo, rodapé e assinatura) usando a skill max-word-pdf
argument-hint: [cliente, local, equipamentos e valor — ou grave um áudio e digite só /orcamento]
---

Use a skill **max-word-pdf** para gerar um **orçamento da MasterLeds** no
template oficial embutido — **NÃO peça o arquivo modelo ao usuário**, ele já
está incluído na skill (`templates/orcamento-masterleds.docx`).

Dados do orçamento (pode estar vazio se o usuário só enviou um áudio): $ARGUMENTS

Como proceder:
- Extraia dos dados acima (ou do áudio da conversa): cliente, data do evento,
  data de montagem, local, itens de equipamento, valor total e forma de
  pagamento.
- Monte um `dados.json` e rode `scripts/orcamento_masterleds.py` (que usa o
  template embutido por padrão). Não passe `--template`.
- **Não invente dados.** O que o usuário não informar, deixe como está no
  modelo; se faltar algo central (ex.: valor total), pergunte antes de gerar.
- Entregue o `.docx` gerado com `SendUserFile` e lembre que o PDF fiel sai
  exportando pelo Word ("Salvar como PDF").
