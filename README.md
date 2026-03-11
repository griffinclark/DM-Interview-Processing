# HollyPlanner

HollyPlanner is a locked-workbook intake app for financial planners. It turns planner-facing PDF financial plans into the exact Excel model shipped in this repository, keeps the workbook structure on rails, and gives reviewers a resumable import flow instead of a one-shot black box.

The current build is a three-stage pipeline:

1. OCR
2. Workbook Entry
3. Financial Calculations

LangGraph-driven sheet entry, deterministic workbook writes, resumable question handling, and workbook validation all live in one Streamlit surface.

## License

This repository is intentionally **non-commercial only** for interview and evaluation use.

- License notice: [LICENSE](/Users/griffin/Desktop/DM Interview Processing/LICENSE)
- Required notice: [NOTICE](/Users/griffin/Desktop/DM Interview Processing/NOTICE)

If you need commercial rights, you must secure a separate written license from the copyright holder.

## What It Does Now

- Uses the exact workbook template at [assets/template/Copy of FP Case Study Model - Alicia and Tom Smith - March 5, 10_01 AM.xlsx](/Users/griffin/Desktop/DM Interview Processing/assets/template/Copy%20of%20FP%20Case%20Study%20Model%20-%20Alicia%20and%20Tom%20Smith%20-%20March%205,%2010_01%E2%80%AFAM.xlsx)
- Verifies the locked template SHA-256 before every run
- Renders PDF pages locally with PyMuPDF and runs OCR across parallel lanes with retries, heartbeat updates, and shared cooldown handling after rate limits
- Fills the workbook one sheet at a time with a LangGraph agent that prioritizes `Data Input`, carries forward prior writes, and only writes to whitelisted cells and row blocks
- Uses staged evidence escalation during workbook entry: workbook context first, structured OCR next, raw-PDF rereview last, then a user-facing decision gate only if ambiguity still remains
- Gives the `Expenses` sheet a read-only `query_transactions` tool over `Transactions Raw`, exposed as an in-memory SQLite dataset so the agent can inspect or aggregate ledger rows without flooding the prompt window
- Keeps `Transactions Raw` out of preloaded prompt context by default to preserve tokens while still preserving that sheet in the workbook output
- Persists resumable job artifacts so a paused run can continue after a user answer or an explicit "let the agent decide" handoff
- Rebuilds the workbook from canonical assignments after each sheet pass so downstream sheets see current workbook state instead of stale context
- Preserves formulas, styles, sheet order, named ranges, validations, and layout
- Runs formula-count validation plus structural drift checks before declaring success
- Applies review gates based on unresolved critical sheets, minimum supported coverage, and unresolved-target count
- Produces a detailed `review_report.json` with warnings, assignments, assumptions, questions asked, user answers, coverage summary, drift results, and calculation validation results
- Streams live status in a custom Streamlit shell with sticky stage chrome, rate-limit countdowns, agent progress heartbeats, and downloadable final artifacts

## End-to-End Flow

### 1. OCR

- The app verifies the locked template, renders the uploaded PDF, and fans OCR work out across `N` parallel lanes.
- Lanes refill immediately as pages complete, so shorter pages do not leave workers idle.
- Retry events are surfaced with cooldown timing and provider-specific messaging.

### 2. Workbook Entry

- A LangGraph sheet-entry agent walks the template in sheet order, with `Data Input` pulled to the front.
- The agent starts with workbook-native context when possible, escalates to structured OCR when needed, and only falls back to raw OCR text for a final targeted rereview.
- If a sheet is still materially ambiguous, the run pauses behind a "Decision Gate" question instead of silently guessing.
- Resuming supports both explicit human answers and agent delegation.
- For `Expenses`, the agent can query `Transactions Raw` through the read-only `query_transactions` tool instead of relying on prompt-dumped sample rows.

### 3. Financial Calculations

- The output workbook is rebuilt from canonical assignments.
- LibreOffice recalculation is attempted when `soffice` is available; otherwise the workbook is left for Excel recalc-on-open.
- Formula counts are checked against the locked template baseline.
- Structural drift is checked against the template so unsupported sheet/layout edits fail hard.
- Final success requires both a valid workbook and acceptable coverage.

## Output Artifacts

Each run creates a job directory under `tmp/jobs/<job_id>/` containing:

- The source PDF
- `filled_financial_plan.xlsx`
- `ocr_results.json`
- `entry_state.json`
- `review_report.json`

The persisted `entry_state.json` and `ocr_results.json` are what make paused workbook-entry sessions resumable.

## Streamlit UI

The UI is no longer a bare upload form. The current app includes:

- A custom Streamlit shell with a sticky taskbar that shows current stage, status tone, and stage completion
- Live OCR/retry countdowns rendered via a small HTML bridge instead of static timestamps
- A dedicated workbook-entry question surface for human clarification and agent-delegation resumes
- Download actions for both the filled workbook and the review report
- A runtime settings dialog for per-session OCR lane count changes
- Provider branding in the UI, while still keeping provider credentials in `.env`

## Local Setup

Create and use the project virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

Set up the environment file:

```bash
cp example.env .env
```

Then edit `.env` and replace the placeholder with your real API key. The app loads `example.env` first and overrides it with `.env` when present.

Minimum `.env`:

```bash
OPENAI_API_KEY="your-real-openai-key"
```

Optional Anthropic entry:

```bash
# ANTHROPIC_API_KEY="your-anthropic-key"
```

Runtime settings stay locked in code, and the UI currently exposes only the OCR parallel-worker count for the current browser session. Current defaults:

```bash
LLM_PROVIDER="openai"
OPENAI_MODEL="gpt-5.2"
ANTHROPIC_MODEL="claude-sonnet-4-5-20250929"
OCR_PARALLEL_WORKERS="6"
LLM_TIMEOUT_SECONDS="240"
LLM_MAX_RETRIES="2"
LLM_RETRY_BASE_SECONDS="10"
LLM_RETRY_MAX_SECONDS="300"
MAX_PAGES="40"
MIN_SUPPORTED_COVERAGE="0.70"
MAX_UNRESOLVED_SUPPORTED_TARGETS="10"
LOG_LEVEL="INFO"
```

Anthropic support still exists in the runtime plumbing, but this build keeps OpenAI active in the UI. If you want to switch providers, add `ANTHROPIC_API_KEY` to `.env` and change the default provider in code.

The locked template path and checksum stay in code and are not exposed in the settings dialog.

Run locally:

```bash
source .venv/bin/activate
.venv/bin/streamlit run app.py
```

## Docker

Build the image:

```bash
docker build -t hollyplanner .
```

Start the container with the local `.env` file:

```bash
docker run --rm -p 8501:8501 --env-file .env hollyplanner
```

Then open `http://localhost:8501` in your browser.

If you prefer to pass variables explicitly instead of using `--env-file`:

```bash
docker run --rm -p 8501:8501 \
  -e OPENAI_API_KEY="$OPENAI_API_KEY" \
  hollyplanner
```

If you do not already have `.env`, create it from the template before starting the container:

```bash
cp example.env .env
```

Then edit `.env` and set:

```bash
OPENAI_API_KEY="your-real-openai-key"
```

To run with Anthropic instead, also pass `-e ANTHROPIC_API_KEY="$ANTHROPIC_API_KEY"` and change the default provider in code before starting the app.

The container includes LibreOffice so the app can attempt a headless recalculation pass during stage 3.

## Testing

Run the test suite from the project virtual environment:

```bash
source .venv/bin/activate
.venv/bin/pytest
```

The current tests cover:

- LangGraph sheet-entry behavior, including workbook-only vs OCR escalation and raw-PDF rereview
- The `query_transactions` tool and its read-only SQL guardrails
- Resumable job flow after workbook-entry questions
- Coverage-based review gating
- Workbook write behavior and drift detection
- Streamlit UI rendering for taskbar, countdowns, provider rail, and question/resume handling

## Notes For Reviewers

- HollyPlanner is **template-specific by design**. It intentionally rejects alternate workbook layouts.
- Missing or ambiguous PDF data is left unresolved and surfaced for review instead of being guessed into the workbook.
- `Transactions Raw` is preserved as workbook data but is not dumped into prompt context by default; the agent accesses it through a tool when it needs ledger detail.
- Structural workbook drift is treated as a hard failure.
