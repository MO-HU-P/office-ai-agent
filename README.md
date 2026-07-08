# Office AI Agent

[🇯🇵 日本語](README.ja.md) | **🇬🇧 English**

**Create and edit Word, Excel, and PowerPoint files by chatting with an LLM — running locally with Docker.**

Ask in plain language ("make a monthly sales spreadsheet with a totals row") and the AI agent creates or edits real `.docx` / `.xlsx` / `.pptx` files in your workspace, with a live preview in the browser. Files open normally in Microsoft Office or LibreOffice.

> The UI is currently Japanese-only: the primary audience is non-engineer Japanese users. The detailed end-user guide is in [README.ja.md](README.ja.md).

## Features

- **Excel** — build tables, formulas, aggregation, formatting, conditional highlighting, charts
- **Word** — documents with headings/lists/tables, batch edits, copying the look & feel of an existing document to another
- **PowerPoint** — multi-slide decks with shapes, images, tables, and charts
- **Review, don't overwrite (Word)** — the agent can propose fixes as tracked changes you accept/reject in real Word (rendered as red strike-through / green underline in the preview), or leave margin comments, instead of editing in place
- **Translate (Word / PowerPoint / Excel)** — translate a document into another language; the agent first asks whether to overwrite the original text or keep it and show the translation side-by-side, then preserves formatting. Numbers, dates, and Excel formulas are left untouched, and translated slides are checked for overflow with a vision model
- **Compare versions** — diff two `.docx` / `.xlsx` / `.pptx` files (by content) into a plain-language summary of what changed
- **Automatic backups, change review & undo** — every overwrite or deletion snapshots the previous state to `workspace/.history` first (20 generations per file). A history button in the preview shows what the last request changed (paragraph-level for Word, cell-level for Excel, per-slide text for PowerPoint) and rolls the file back with one click; you can also just ask the agent to "undo that edit". Restores are themselves undoable
- **Anonymize before sending to the cloud (your call)** — to avoid handing raw personal data to a cloud model, you can redact structured PII (email, phone, URL, postal code, My Number, card numbers) into a copy first, then work on that copy. The masking runs deterministically in Python and is *never* sent to the LLM; deciding whether/what to anonymize is up to the user, and names/addresses are out of scope and flagged for manual review
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
- **Safety net**: the single atomic-save choke point also versions the previous state of any overwritten/deleted file into `workspace/.history` (`services/history.py`; `run_python` bypasses that path, so the workspace is cheaply snapshotted before it runs, skipping unchanged files). Diffs and rollbacks are grouped per user request: `GET /api/files/{name}/changes`, `POST /api/files/{name}/restore`, plus `restore_file` / `list_file_versions` agent tools. Hidden dot-paths (incl. `.history`) are rejected by the workspace path resolver. Backend tests live in `backend/tests` (pytest)
- **Preview**: Excel via a custom grid, Word via docx-preview, PowerPoint via LibreOffice headless → PDF → PNG
- **LLM providers**: Ollama (local or cloud) is the zero-setup, no-credit-card default; users who have an API key can also switch to OpenAI or Google Gemini. Model construction, reasoning mapping, and vision detection are centralized in `agent/providers.py` (add a provider by writing one `_build_*` function). Provider/model/reasoning are switchable at runtime from the settings UI, no restart needed. The OpenAI and Gemini model shortlists are config-driven (`config.toml` `llm.openai_models` / `llm.gemini_models`, edit when models get deprecated); any other model can be typed in the UI and is remembered per install (in `data/settings.json`).

## Quick start

Requires Docker. For cloud mode, get an API key at [ollama.com/settings/keys](https://ollama.com/settings/keys).

```bash
cp .env.example .env       # OLLAMA_API_KEY for Ollama Cloud (empty = local);
                           # OPENAI_API_KEY / GEMINI_API_KEY optional, enable the OpenAI / Gemini providers
docker compose up -d --build
```

Open **http://localhost:3000**. Pick the provider (Ollama cloud/local, or OpenAI / Google Gemini if a key is set), choose or pull models, and tune reasoning from the gear icon in the header. The OpenAI and Gemini options only appear once `OPENAI_API_KEY` / `GEMINI_API_KEY` are set.

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
- **Mind what you upload.** With a cloud provider, a file's contents are sent to the provider when you ask the agent to work on it. Avoid uploading files that contain personal information **or confidential/secret material** (internal-only or unpublished data) as-is. Anonymization can redact structured personal data first, but it does **not** remove confidential content — for sensitive material, switch to local mode, where nothing leaves the machine.
- API keys live only in `.env` (`OLLAMA_API_KEY`, `OPENAI_API_KEY`, `GEMINI_API_KEY`); they are never returned by the API, logged, or shown in the UI — only a `*_key_configured` boolean is exposed. Model-name inputs reject anything that looks like a key (`sk-…`, `AIza…`) so a mis-paste can't be persisted to `settings.json`.
- Tool docstrings and the system prompt are intentionally in Japanese: they are part of the LLM prompt, and matching the language of user requests improves tool selection — especially for small local models.
- **Big requests work best step by step** (create the table → add the analysis → write the report). If a cloud call fails or stalls mid-run, generated files are kept — just ask the agent to continue. Transient server errors (e.g. HTTP 500) are retried with exponential backoff, and stalled/silent LLM streams are cut off by both an idle and a total-response timeout, then retried (all tunable in `config.toml`: `agent.max_steps`, `llm_idle_timeout`, `llm_step_timeout`, `llm_max_attempts`, `llm_retry_backoff_cap`). For long multi-step runs a hosted provider (Ollama Cloud, OpenAI, or Google Gemini) is more reliable than free-tier capacity — switch providers in the settings if one is flaky.
- No GPU? Remove the `gpus: all` line in `docker-compose.yml` (local mode will run on CPU).

## Trademarks

This project is not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft, Office, Word, Excel, and PowerPoint are trademarks of Microsoft Corporation. Any use of these names here is purely descriptive, to indicate the file formats the app works with.

## License

[Apache License 2.0](LICENSE)
