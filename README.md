# ⌨️ AutoHotkey Excel — Envio Automático de Texto para o Excel

Script AutoHotkey que captura o texto selecionado em qualquer aplicativo e o envia automaticamente para a próxima linha disponível em uma coluna específica do Excel, ao pressionar a tecla `Insert`.

---

## 🚀 Funcionalidades

- Captura o texto selecionado com `Ctrl+C` ao pressionar `Insert`
- Detecta automaticamente a instância ativa do Microsoft Excel via COM
- Identifica a última linha preenchida na coluna C e insere o texto na linha seguinte
- Exibe aviso caso o Excel não esteja aberto

---

## 🛠️ Tecnologias

- [AutoHotkey v1](https://www.autohotkey.com/) — automação de teclado e interface Windows
- COM Object (`ComObjActive`) — integração direta com o Excel em execução

---

## 🔧 Personalização

Para alterar a coluna de destino, edite a linha abaixo no script substituindo o número `3` pelo índice da coluna desejada (A=1, B=2, C=3...):

```autohotkey
lastRow := sheet.Cells(sheet.Rows.Count, 3).End(-4162).Row
...
sheet.Cells(nextRow, 3).Value := Clipboard
```

---

## 👤 Autor

**Felipe Pipelmo**  
Estudante de Engenharia de Controle e Automação — focado em Automação de Processos, Engenharia de Dados e Integrações de APIs.

[![GitHub](https://img.shields.io/badge/GitHub-FelipePipelmo-181717?logo=github)](https://github.com/FelipePipelmo)
