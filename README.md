# Trello JSON → Excel

Aplicação em Streamlit que converte um arquivo JSON exportado do Trello em um Excel com múltiplas abas. A conclusão de cards é baseada exclusivamente no campo `dueComplete` do Trello. Os campos personalizados (Custom Fields) do quadro são suportados de forma dinâmica – novos campos como “Frente” e “Defeito” passam a aparecer automaticamente nas tabelas e na exportação.

## Recursos

- Upload de JSON do Trello e visualização dos dados em tabelas.
- Exportação para Excel com as abas:
  - Cards
  - Checklists
  - ChecklistItems
  - FlatExport (união dos cards com checklist items)
- Suporte dinâmico a campos personalizados (Custom Fields) do Trello.
- Nomes de colunas do FlatExport mais intuitivos em português.

## Publicação

- Repositório publicado no Streamlit Cloud: https://trello-json-convert-to-excel.streamlit.app/

## Requisitos

- Python 3.10+
- Dependências:
  - streamlit
  - pandas
  - XlsxWriter (ou openpyxl como fallback)

Veja [requirements.txt](file:///c:/Repositórios/trello-json-convert-to-excel/requirements.txt).

## Instalação e execução

Forma recomendada (PowerShell no Windows):

```powershell
powershell -ExecutionPolicy Bypass -File .\start-local.ps1
```

- O script cria/usa `.venv`, instala dependências e inicia o Streamlit em modo headless.
- Acesse a aplicação em: http://localhost:8501

Sem usar o venv do projeto (usa Python do sistema):

```powershell
powershell -ExecutionPolicy Bypass -File .\start-local.ps1 -NoVenv
```

Execução manual (alternativa):

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Uso

1. Abra a aplicação em http://localhost:8501.
2. Faça upload do JSON exportado do Trello.
3. Visualize as tabelas no navegador.
4. Clique em “Baixar Excel completo” para gerar o arquivo `.xlsx` com todas as abas.

## Abas e colunas

### Cards
Informações principais do card (lista, título, conclusão por `dueComplete`, prazos, membros, etiquetas, links, datas de atividade e início, descrição) e também os campos personalizados do board (ex.: “Frente”, “Defeito”).  
Fonte: parser em [app.py](file:///c:/Repositórios/trello-json-convert-to-excel/app.py).

### Checklists
Uma linha por checklist com totais e contagem de concluídos/pendentes.

### ChecklistItems
Uma linha por item de checklist, com estado e responsável (quando presente).

### FlatExport
Junção dos Cards com ChecklistItems quando existirem (caso contrário, espelha Cards). Os nomes das colunas foram renomeados para português intuitivo. Exemplos do mapeamento aplicado:

- `card_id` → `Card ID`
- `idShort` → `Número`
- `card_name` → `Título`
- `list_name` → `Lista`
- `dueComplete` → `Concluído`
- `card_due` → `Data de Entrega`
- `labels` → `Etiquetas`
- `members` → `Membros`
- `shortLink` → `Link curto`
- `card_dateLastActivity` → `Última atividade`
- `card_start` → `Início`
- `card_desc` → `Descrição`
- `checklist_name` → `Checklist`
- `checkitem_name` → `Item`
- `state` → `Status`
- `responsavel` → `Responsável`

Os campos personalizados do Trello são mantidos com os nomes definidos no próprio quadro, por exemplo `Frente`, `Defeito`.

## Campos personalizados (dinâmicos)

- O parser identifica automaticamente os Custom Fields retornados no JSON do Trello, e adiciona cada campo como coluna com o nome configurado no quadro.
- Tipos suportados: texto, número, data, checkbox, opção de lista (dropdown).  
- A lógica está no parser em [app.py](file:///c:/Repositórios/trello-json-convert-to-excel/app.py).


## Dicas e observações

- O Excel é gerado usando `XlsxWriter` quando disponível, com fallback para `openpyxl`. Se nenhuma engine estiver instalada, o app solicitará instalação.
- Para evitar problemas de permissão, o script `start-local.ps1` usa o Python do `.venv` e instala dependências dentro do ambiente virtual.
- A visão “Explore” (pendências) foi removida da UI e da exportação conforme solicitado, mantendo foco nas abas principais.

## Desenvolvimento

- Código principal: [app.py](file:///c:/Repositórios/trello-json-convert-to-excel/app.py)
  - Helpers de Excel (`dfs_to_xlsx_bytes`, sanitização de dados).
  - Helpers de Trello (datas, labels e conclusão por `dueComplete`).
  - `parse_trello`: constrói os DataFrames `df_cards`, `df_checklists`, `df_items`, `df_flat`.
  - UI em Streamlit com upload de arquivo e download do Excel.

## Licença

MIT License — livre para usar, copiar, modificar e distribuir.

# Trello JSON → Excel (English)

Streamlit application that converts a Trello exported JSON into an Excel file with multiple sheets. Card completion is based solely on Trello’s `dueComplete` flag. Board Custom Fields are supported dynamically — new fields such as “Frente” and “Defeito” automatically appear in tables and in the export.

## Features
- Upload Trello JSON and view data in tables.
- Export to Excel with sheets:
  - Cards
  - Checklists
  - ChecklistItems
  - FlatExport (cards joined with checklist items)
- Dynamic support for Trello Custom Fields.
- Intuitive Portuguese column names in FlatExport.
- “Explore” view has been removed as requested.

## Deployment
- Published on Streamlit Cloud: https://trello-json-convert-to-excel.streamlit.app/

## Requirements
- Python 3.10+
- Dependencies:
  - streamlit
  - pandas
  - XlsxWriter (or openpyxl as fallback)
See [requirements.txt](file:///c:/Repositórios/trello-json-convert-to-excel/requirements.txt).

## Install & Run
Recommended (PowerShell on Windows):

```powershell
powershell -ExecutionPolicy Bypass -File .\start-local.ps1
```

- The script creates/uses `.venv`, installs dependencies, and starts Streamlit in headless mode.
- Access the app at: http://localhost:8501

Without project venv (use system Python):

```powershell
powershell -ExecutionPolicy Bypass -File .\start-local.ps1 -NoVenv
```

Manual run (alternative):

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Usage
1. Open http://localhost:8501.
2. Upload the Trello exported JSON.
3. Browse the tables in the browser.
4. Click “Baixar Excel completo” to download the `.xlsx` with all sheets.

## Sheets & Columns
### Cards
Card main information (list, title, `dueComplete`-based completion, due dates, members, labels, links, activity and start dates, description) and also board Custom Fields (e.g., “Frente”, “Defeito”).  
Source: parser in [app.py](file:///c:/Repositórios/trello-json-convert-to-excel/app.py).

### Checklists
One row per checklist with totals and counts of completed/pending.

### ChecklistItems
One row per checklist item, with status and assignee (when present).

### FlatExport
Join of Cards with ChecklistItems when they exist (otherwise mirrors Cards). Column names were renamed to intuitive Portuguese. Examples:
- `card_id` → `Card ID`
- `idShort` → `Número`
- `card_name` → `Título`
- `list_name` → `Lista`
- `dueComplete` → `Concluído`
- `card_due` → `Data de Entrega`
- `labels` → `Etiquetas`
- `members` → `Membros`
- `shortLink` → `Link curto`
- `card_dateLastActivity` → `Última atividade`
- `card_start` → `Início`
- `card_desc` → `Descrição`
- `checklist_name` → `Checklist`
- `checkitem_name` → `Item`
- `state` → `Status`
- `responsavel` → `Responsável`

Trello Custom Fields are kept with their board-defined names (e.g., `Frente`, `Defeito`).

## Custom Fields (dynamic)
- The parser automatically detects Custom Fields present in Trello JSON and adds each field as a column with the board-configured name.
- Supported types: text, number, date, checkbox, list option (dropdown).  
- Logic is in the parser at [app.py](file:///c:/Repositórios/trello-json-convert-to-excel/app.py).

## Model File
- An example Trello exported JSON is available (ignored by git):  
  [modelo - sumitomo-mirai.json](file:///c:/Repositórios/trello-json-convert-to-excel/modelo%20-%20sumitomo-mirai.json)

## Tips
- Excel is generated using `XlsxWriter` when available, with fallback to `openpyxl`. If no engine is installed, the app will request installation.
- To avoid permission issues, `start-local.ps1` uses `.venv` Python and installs dependencies inside the virtual environment.
- The “Explore” view (pendências) was removed from the UI and export.

## Development
- Main code: [app.py](file:///c:/Repositórios/trello-json-convert-to-excel/app.py)
  - Excel helpers (`dfs_to_xlsx_bytes`, data sanitization).
  - Trello helpers (dates, labels, and `dueComplete` completion).
  - `parse_trello`: builds `df_cards`, `df_checklists`, `df_items`, `df_flat`.
  - Streamlit UI with file upload and Excel download.

## License
MIT License — free to use, copy, modify, and distribute.
