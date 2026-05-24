# ui/desktop/ — Apps Tkinter do ExtratorUnificado

Esta pasta contém **dois apps Tkinter independentes**. Eles usam a mesma
tecnologia (Tkinter), mas têm propósitos distintos.

## `main.py` + `janela.py` — App principal de extração

App desktop **antigo**, que orquestra a extração de boletos via Selenium.

- `main.py`: entry point (`from ui.desktop.janela import App; App().mainloop()`)
- `janela.py`: janela Tkinter principal com radio buttons, combo boxes etc.

Fluxo: a janela chama `core.extracao.Extrator`, que decide qual módulo de
extração rodar conforme o tipo selecionado (`Biptu`, `condominios`,
`taxabombeiros`, ...).

**Como rodar:**

```powershell
# Da raiz do projeto:
python -m ui.desktop.main
```

## `helper_upload.py` — Ponte WMB → webapp

App desktop **independente** que serve como ponte entre o banco local
(`.WMB` / MS Access) e o webapp Flask (`ui/web/webapp/`).

**Por que existe:** o webapp roda num servidor que não tem acesso ao arquivo
`.WMB` no PC do usuário. Esse helper:

1. Conecta no `.WMB` local via `pyodbc`
2. Executa o SELECT escolhido pelo usuário (Prefeitura/IPTU, Bombeiros, etc.)
3. Envia o resultado como CSV para o webapp via API HTTP

```
┌──────────────────────────┐       ┌────────────────────┐
│  PC do usuário           │       │  Servidor          │
│                          │       │                    │
│  Scai.WMB (MS Access)    │       │  webapp Flask      │
│         │                │       │  (ui/web/webapp/)  │
│         │ lê via pyodbc  │       │         │          │
│         ▼                │       │         │          │
│  helper_upload.py        │──────▶│  POST /api/upload  │
│  (Tkinter desktop)       │ CSV   │  (recebe CSV)      │
└──────────────────────────┘ HTTP  └────────────────────┘
```

**Como rodar:**

```powershell
# Da raiz do projeto:
python -m ui.desktop.helper_upload
```

## Resumo

| Arquivo | Função | Quando usar |
|---|---|---|
| `main.py` + `janela.py` | App principal de extração | Para extrair boletos via UI desktop |
| `helper_upload.py` | Ponte de dados WMB → webapp | Para popular o webapp com dados do banco local |

Os dois apps são totalmente independentes — `helper_upload.py` **não** chama
nenhum dos extratores, e a janela principal **não** sabe da existência do
webapp.
