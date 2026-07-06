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

- **Agent**: a hand-rolled ReAct loop (LangChain core, deliberately no LangGraph) with streaming tokens, tool events, and error recovery over WebSocket
- **File editing**: python-docx / openpyxl / python-pptx directly on files — no Office automation; every save is atomic (temp file + `os.replace`)
- **Preview**: Excel via a custom grid, Word via docx-preview, PowerPoint via LibreOffice headless → PDF → PNG
- **LLM providers**: Ollama (local or cloud) is the zero-setup, no-credit-card default; users who have an API key can also switch to OpenAI. Model construction, reasoning mapping, and vision detection are centralized in `agent/providers.py` (add a provider by writing one `_build_*` function). Provider/model/reasoning are switchable at runtime from the settings UI, no restart needed.

## Quick start

Requires Docker. For cloud mode, get an API key at [ollama.com/settings/keys](https://ollama.com/settings/keys).

```bash
cp .env.example .env       # OLLAMA_API_KEY for Ollama Cloud (empty = local);
                           # OPENAI_API_KEY optional, enables the OpenAI provider
docker compose up -d --build
```

Open **http://localhost:3000**. Pick the provider (Ollama cloud/local, or OpenAI if a key is set), choose or pull models, and tune reasoning from the gear icon in the header. The OpenAI option only appears once `OPENAI_API_KEY` is set.

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
- API keys live only in `.env` (`OLLAMA_API_KEY`, `OPENAI_API_KEY`); they are never returned by the API, logged, or shown in the UI — only a `*_key_configured` boolean is exposed.
- Tool docstrings and the system prompt are intentionally in Japanese: they are part of the LLM prompt, and matching the language of user requests improves tool selection — especially for small local models.
- **Big requests work best step by step** (create the table → add the analysis → write the report). If a cloud call fails or stalls mid-run, generated files are kept — just ask the agent to continue. Transient server errors (e.g. HTTP 500) are retried with exponential backoff, and stalled/silent LLM streams are cut off by both an idle and a total-response timeout, then retried (all tunable in `config.toml`: `agent.max_steps`, `llm_idle_timeout`, `llm_step_timeout`, `llm_max_attempts`, `llm_retry_backoff_cap`). For long multi-step runs a hosted provider (Ollama Cloud or OpenAI) is more reliable than free-tier capacity — switch providers in the settings if one is flaky.
- No GPU? Remove the `gpus: all` line in `docker-compose.yml` (local mode will run on CPU).

## License

[Apache License 2.0](LICENSE)
