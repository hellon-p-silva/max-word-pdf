---
description: Gera um documento (PDF/Word) a partir de áudio ou texto usando a skill max-word-pdf
argument-hint: [descreva o documento — ou grave um áudio e digite só /documento]
---

Use a skill **max-word-pdf** para gerar o documento pedido.

Pedido do usuário (pode estar vazio se ele só enviou um áudio): $ARGUMENTS

Como proceder:
- Se houver um **áudio** nesta conversa, use o conteúdo falado como base.
- Se o pedido acima descrever um documento por texto, use esse texto.
- Siga o fluxo da skill: entender → organizar (limpar vícios de fala, dar
  estrutura, não inventar dados) → escolher o formato (PDF, Word ou ambos,
  conforme pedido; se não especificado, gere ambos) → gerar → entregar os
  arquivos com `SendUserFile`.
- Se for um **orçamento da MasterLeds**, use o template embutido da skill
  (`templates/orcamento-masterleds.docx`) via `scripts/orcamento_masterleds.py`
  — **não peça nenhum arquivo modelo ao usuário**.
- Se faltar um dado importante (destinatário, data, valor), pergunte antes de
  gerar em vez de inventar.
