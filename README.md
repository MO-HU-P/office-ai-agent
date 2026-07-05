# Office AI Agent

[🇯🇵 日本語](README.ja.md) | **🇬🇧 English**

**Create and edit Word, Excel, and PowerPoint files by chatting with an LLM — running locally with Docker.**

Ask in plain language ("make a monthly sales spreadsheet with a totals row") and the AI agent creates or edits real `.docx` / `.xlsx` / `.pptx` files in your workspace, with a live preview in the browser. Files open normally in Microsoft Office or LibreOffice.

> The UI is currently Japanese-only: the primary audience is non-engineer Japanese users. The detailed end-user guide is in [README.ja.md](README.ja.md).

## Features

- **Excel** — build tables, formulas, aggregation, formatting, conditional highlighting, charts
- **Word** — documents with headings/lists/tables, batch edits, copying the look & feel of an existing document to another
- **PowerPoint** — multi-slide decks with shapes, images, tables, and charts
- **Self-review loop** — with a vision-capable model, the agent renders slides/pages to images and fixes overflow or overlapping layout by looking at them
- **Mail merge** — fill `{{placeholders}}` in a template to mass-produce documents
- **Data analysis & statistics** — pandas / NumPy / SciPy / statsmodels / seaborn via a sandboxed `run_python` tool: hypothesis tests, regression with full summary tables, ANOVA, GLM, time series (ARIMA), R-style formulas (`y ~ x1 + x2`), and Japanese-ready charts — then written up as Word/PowerPoint reports
- **Document QA** — checks for missing headings, broken Excel formulas, overcrowded slides
- **File handling** — upload your own files (the agent works on a copy), copy/rename/delete via chat
- **Voice input** — browser speech recognition (Chrome/Edge)

## Architecture

Three containers via Docker Compose, exposing only port 3000:

```
frontend (nginx + React/TypeScript, Vite)
   └─ proxies /api and /ws → backend (FastAPI, Python 3.12)
                                └─ ollama (local mode only, internal network)
```

- **Agent**: a hand-rolled ReAct loop (LangChain core + Ollama, deliberately no LangGraph) with streaming tokens, tool events, and error recovery over WebSocket
- **File editing**: python-docx / openpyxl / python-pptx directly on files — no Office automation; every save is atomic (temp file + `os.replace`)
- **Preview**: Excel via a custom grid, Word via docx-preview, PowerPoint via LibreOffice headless → PDF → PNG
- **LLM**: Ollama Cloud (API key) or fully local Ollama — switchable at runtime from the settings UI, no restart needed

## Quick start

Requires Docker. For cloud mode, get an API key at [ollama.com/settings/keys](https://ollama.com/settings/keys).

```bash
cp .env.example .env       # put OLLAMA_API_KEY=... in .env (leave empty for local mode)
docker compose up -d --build
```

Open **http://localhost:3000**. Switch between cloud/local mode, pick or pull models, and tune reasoning from the gear icon in the header.

Stop with `docker compose down`. Generated files persist in `./workspace` on the host.

### Development without Docker

```bash
# backend (Python 3.12)
cd backend && pip install -r requirements.txt
WORKSPACE_DIR=../workspace OLLAMA_BASE_URL=http://localhost:11434 uvicorn app.main:app --reload

# frontend
cd frontend && npm install && npm run dev   # http://localhost:5173
```

## Notes

- **No authentication** — designed for localhost / trusted networks only. The agent executes Python for data analysis, so do not expose it to the internet.
- API keys live only in `.env`; they are never returned by the API, logged, or shown in the UI.
- Tool docstrings and the system prompt are intentionally in Japanese: they are part of the LLM prompt, and matching the language of user requests improves tool selection — especially for small local models.
- **Big requests work best step by step** (create the table → add the analysis → write the report). If a cloud call fails or stalls mid-run, generated files are kept — just ask the agent to continue. Stalled LLM streams are cut off and retried automatically (tunable via `agent.llm_idle_timeout` in `config.toml`).
- No GPU? Remove the `gpus: all` line in `docker-compose.yml` (local mode will run on CPU).

## License

[Apache License 2.0](LICENSE)
