# Day 03. Budget vs Actual Tracker.

![AI provider: Anthropic Claude](https://img.shields.io/badge/AI-Anthropic_Claude-D97757?style=flat-square)
![Stack: Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![Local-only](https://img.shields.io/badge/Hosted-Local_only_(127.0.0.1:1003)-555555?style=flat-square)
![Status](https://img.shields.io/badge/Status-MVP_complete-16A34A?style=flat-square)

Local web app + CLI for budget vs actual variance analysis with direction-aware RAG, AI-drafted management commentary, Excel autofilter as drill-through, and an optional Power BI Power Query template for live refresh.

This is **Day 03 of a 30-day finance and AI portfolio sprint** where each
project is shipped end-to-end in a single day, runs locally on its own
loopback port, and integrates the AI in a way that is not cosmetic. The
series alternates between Claude (Days 1 to 3) and DeepSeek V4 (Days 4
onwards) with a deliberate provider switch documented in the README.

---

## AI provider

**This project uses the Anthropic Claude API.**

- **Model**: Claude Haiku 4.5 (via the Anthropic SDK).
- **Cost target**: around $0.003 to $0.005 per analysis.
- **Why Claude here**: Days 1 to 3 of the 30-day series ran on Claude Haiku 4.5 because the Anthropic SDK is mature, vision input is first-class (used in Day 02 for invoice OCR), and the pricing was acceptable at the start of the series. Days 4 onwards switched to DeepSeek V4 to cut costs by roughly 5x to 15x; see those repos for the comparison.

The shared client is `shared/llm_client.py` and exposes `ask_claude_json_with_stats(prompt, system, max_tokens, model)`. It returns the parsed JSON plus a `CallResult` with token counts and cost in USD so the cost log is honest.

---

## What it does

Local web app + CLI for budget vs actual variance analysis with direction-aware RAG, AI-drafted management commentary, Excel autofilter as drill-through, and an optional Power BI Power Query template for live refresh.

The MVP is intentionally compact:

1. The user uploads a file (or picks a bundled sample).
2. The pipeline parses, validates and normalises the input.
3. The deterministic finance maths runs locally (no AI involved).
4. A single AI call writes the narrative around those numbers.
5. The web UI renders the result; Excel + PDF + CSV exports drop into `outputs/`.

Every analytical figure is computed by Python, not the LLM. The AI is
asked only to interpret and explain.

## Quickstart (Windows)

One-time setup, from the repo root:

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
copy .env.example .env
notepad .env       :: paste your API key
```

Then run:

```bat
start.bat
```

A browser window opens at `http://127.0.0.1:1003/`.

## Quickstart (macOS / Linux)

```bash
python -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
cp .env.example .env
$EDITOR .env       # paste your API key
chmod +x start.sh
./start.sh
```

## Environment variables

- `ANTHROPIC_API_KEY` -- Required. Get one at https://console.anthropic.com/.
- `CLAUDE_MODEL_FAST` -- Optional override.

The `.env` file is git-ignored. **Never commit it.** A `.env.example` lives
next to it with placeholder values that you can copy.

## Stack

- Python 3.11+
- openpyxl, python-pptx, anthropic, flask
- System fonts only, no CDN dependencies, loopback only.

## File layout

```
day-03-budget-tracker/
  server.py            Flask web server (port 1003)
  main.py              CLI entry point
  pipeline.py          orchestrator
  shared/              vendored shared modules (config, AI client, etc.)
  static/              frontend CSS + JS + favicon
  templates/           Jinja2 HTML
  sample_data/         deterministic samples that round-trip the parser
  tests/               pytest suite
  outputs/             generated artefacts (gitignored)
  uploads/             user uploads (gitignored)
  logs/                rotating server log (gitignored)
  start.bat / start.sh launchers
  requirements.txt
  README.md (this file)
  .env.example         placeholders for the env vars above
  .gitignore
  LICENSE              MIT
```

## Running tests

```bat
.venv\Scripts\python.exe -m pytest
```

The tests do not call the AI provider; the LLM is stubbed where the
pipeline crosses the network.

## Security and privacy notes

- All processing is local. The server binds to `127.0.0.1` only; no
  inbound traffic is accepted from the network.
- The only outbound call is to the AI provider's API endpoint
  (`api.anthropic.com` for Claude or `api.deepseek.com` for DeepSeek).
- Uploaded files stay in `uploads/` and are git-ignored.
- The exception scrubber strips API keys and absolute paths from any
  error surfaced to the UI before the user sees it.
- CSRF double-submit cookie + single-flight semaphore on the analyse
  route prevent CSRF and accidental double-submission.

## Author

Safwan Ahmed. Day 03 of the 30-day finance + AI sprint, 2026.

## License

MIT. See `LICENSE`.
