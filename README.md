# HollyPlanner

HollyPlanner is a Streamlit app that turns a planner-facing PDF financial plan into the locked Excel workbook shipped in this repository. The app is intentionally template-specific: it verifies the bundled workbook checksum before every run and fails final validation if the output workbook drifts outside the allowed write surface.

## Current Behavior

The current build runs a three-step pipeline:

1. OCR
   - Renders PDF pages locally with PyMuPDF.
   - Extracts structured page evidence through the configured LLM client.
   - Processes pages across parallel lanes with retries, heartbeat updates, and shared cooldown handling after rate limits.
2. Workbook entry
   - Uses a LangGraph agent to fill sheets in this order: `Data Input`, `Net Worth`, `Expenses`, `Retirement Accounts`, `Taxable Accounts`, `Education Accounts`.
   - Always prioritizes `Data Input` first.
   - Rebuilds the workbook from canonical assignments after each sheet pass so later sheets read current workbook state, not stale prompt state.
   - Pauses behind a planner-facing question when ambiguity still blocks a supported target.
   - Supports resume with either a human answer or an explicit "let the agent decide" handoff.
3. Final review
   - Rebuilds the workbook from the locked template plus canonical assignments.
   - Attempts LibreOffice recalculation when `soffice` is available.
   - Checks output formula counts against the locked template baseline.
   - Runs structural drift checks against the locked template.

Additional runtime rules that matter:

- `Transactions Raw` is preserved in the workbook and is never prompt-dumped by default.
- On the `Expenses` sheet, the agent gets a read-only `query_transactions` tool backed by an in-memory SQLite view of `Transactions Raw`.
- The agent only writes to whitelisted cells and row blocks defined in code.
- A run is only marked successful when workbook validation passes and coverage rules pass.
- Current coverage gates are `MIN_SUPPORTED_COVERAGE=0.70` and `MAX_UNRESOLVED_SUPPORTED_TARGETS=10`.
- Current critical sheets for review gating are `Data Input` and `Expenses`.

## UI State

The Streamlit UI currently exposes:

- A custom shell with sticky top-level workflow chrome.
- Live OCR lane status, retry countdowns, and workbook-entry heartbeats.
- A planner question surface that appears inline when the run pauses.
- A run settings dialog that only exposes the OCR parallel-lane count for the current browser session.
- Final result states of `Workbook ready` or `Review required` after finalization.

One important detail: the UI shows two top-level workflow stages, `Document review` and `Workbook entry`. The underlying `Financial Calculations` step runs inside the workbook stage view as the final review phase.

## Requirements

- Python 3 with `venv`
- An `OPENAI_API_KEY` for the current default UI/runtime path
- Optional: `soffice` on your local machine if you want LibreOffice-based recalculation during final review

Anthropic runtime plumbing and tests still exist in the codebase, but the settings dialog forces OpenAI in this build. If you want Anthropic end-to-end, you need to change the code, not just `.env`.

## Local Setup

Create and activate the virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

Create the local environment file:

```bash
cp example.env .env
```

`example.env` is loaded first and `.env` overrides it when present.

Minimum `.env`:

```bash
OPENAI_API_KEY=replace-me
```

Run the app:

```bash
source .venv/bin/activate
.venv/bin/streamlit run app.py
```

## Runtime Defaults

These defaults are currently locked in code in [`planlock/config.py`](planlock/config.py):

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

The only runtime setting currently exposed in the UI is `OCR_PARALLEL_WORKERS`, and that change applies only to the current browser session.

## Artifacts

Every job gets a directory under `tmp/jobs/<job_id>/`.

Artifact timing is important:

- The uploaded source PDF is copied into the job directory at job start.
- `filled_financial_plan.xlsx` is created when workbook entry is initialized by copying the locked template.
- `ocr_results.json` is written after OCR completes successfully.
- `entry_state.json` is written when workbook entry starts and is updated as the run advances or pauses.
- `review_report.json` is only written after the run reaches final review.

That means a paused or early-failed run can have a valid job directory without a final `review_report.json`.

UI download behavior is also conditional:

- `review_report.json` is downloadable after finalization.
- `filled_financial_plan.xlsx` is only offered as a download when the workbook passes formula and drift validation, even though a working workbook file may still exist on disk in the job directory.

## Docker

Build the image:

```bash
docker build -t hollyplanner .
```

Run it with your local `.env`:

```bash
docker run --rm -p 8501:8501 --env-file .env hollyplanner
```

Then open [http://localhost:8501](http://localhost:8501).

The Docker image installs LibreOffice, so the container can attempt headless recalculation during final review.

## Testing

Run the test suite from the virtual environment:

```bash
source .venv/bin/activate
.venv/bin/pytest
```

The current tests cover:

- Job-runner stage flow, OCR concurrency, retries, token/progress heartbeats, and resumable question handling
- LangGraph sheet-entry behavior, including workbook-context-first prompting, structured OCR escalation, and raw-PDF rereview
- The `query_transactions` tool and its read-only SQL guardrails
- Workbook writes, preserved formula surfaces, and template drift detection
- Streamlit UI rendering, settings behavior, result/download states, and HTML-rendering helper usage

## Repository Assets

- Locked workbook template: `assets/template/Copy of FP Case Study Model - Alicia and Tom Smith - March 5, 10_01 AM.xlsx`
- Sample PDF: `assets/samples/Domain Money Mock Financial Plan.pdf`

## Review Notes

- HollyPlanner is template-specific by design and is not a general workbook mapper.
- `Transactions Raw` is kept as workbook data but is treated as a read-only evidence source.
- Missing or ambiguous PDF data is surfaced for review instead of being silently guessed into the workbook.
- Structural workbook drift is treated as a hard validation failure.

## License

This repository is noncommercial only.

- License: [`LICENSE`](LICENSE)
- Required notice: [`NOTICE`](NOTICE)

If you need commercial rights, you need a separate written license from the copyright holder.
