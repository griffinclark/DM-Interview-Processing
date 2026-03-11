# PlanLock

PlanLock is a locked workbook intake app for financial planners. It imports planner-facing PDF financial plans into the exact Excel workbook template shipped in this repository, enforces a no-drift structure policy, and exposes a two-step workflow:

1. OCR
2. Workbook Entry
   LangGraph sheet entry and deterministic workbook validation share one UI surface.

## License

This repository is intentionally **non-commercial only** for interview and evaluation use.

- License notice: [LICENSE](/Users/griffin/Desktop/DM Interview Processing/LICENSE)
- Required notice: [NOTICE](/Users/griffin/Desktop/DM Interview Processing/NOTICE)

If you need commercial rights, you must secure a separate written license from the copyright holder.

## What It Does

- Uses the exact workbook template at [assets/template/Copy of FP Case Study Model - Alicia and Tom Smith - March 5, 10_01 AM.xlsx](/Users/griffin/Desktop/DM Interview Processing/assets/template/Copy%20of%20FP%20Case%20Study%20Model%20-%20Alicia%20and%20Tom%20Smith%20-%20March%205,%2010_01%E2%80%AFAM.xlsx)
- Validates the template SHA-256 before startup
- Runs OCR across parallel lanes with retries, immediately refilling each lane as pages finish
- Shares OpenAI cooldown windows across lanes so retries do not keep hammering the API after a 429
- Writes only to whitelisted workbook cells and repeating row blocks
- Preserves formulas, styles, sheet order, named ranges, validations, and layout
- Generates a sidecar `review_report.json` for planner review
- Streams logs and two-step progress in the Streamlit UI

## Local Setup

Create and use the project virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
cp example.env .env
```

Then update `.env` with your real API key. The app automatically loads `example.env` first and overrides it with `.env` when present.

Required environment variables:

```bash
OPENAI_API_KEY="your-key"
```

Runtime settings stay locked in code, and the app’s `Settings` modal currently exposes the parallel OCR lane count for the current browser session. The defaults are:

```bash
LLM_PROVIDER="openai"
OPENAI_MODEL="gpt-5.2"
ANTHROPIC_MODEL="claude-sonnet-4-5-20250929"
OCR_PARALLEL_WORKERS="3"
LLM_TIMEOUT_SECONDS="120"
LLM_MAX_RETRIES="2"
LLM_RETRY_BASE_SECONDS="10"
LLM_RETRY_MAX_SECONDS="300"
MAX_PAGES="40"
LOG_LEVEL="INFO"
```

Anthropic support remains in the runtime plumbing, but this build keeps OpenAI active in the UI. If you want to switch providers, add `ANTHROPIC_API_KEY` to `.env` and change the default provider in code.

The locked template path and checksum stay in code and are not exposed in the settings modal.

Run locally:

```bash
source .venv/bin/activate
streamlit run app.py
```

## Docker

Build and run:

```bash
docker build -t planlock .
docker run --rm -p 8501:8501 \
  -e OPENAI_API_KEY="$OPENAI_API_KEY" \
  planlock
```

To run with Anthropic instead, also pass `-e ANTHROPIC_API_KEY="$ANTHROPIC_API_KEY"` and change the default provider in code before starting the app.

The container includes LibreOffice so the app can attempt a headless recalculation pass during stage 3.

## Testing

```bash
source .venv/bin/activate
pytest
```

## Notes For Reviewers

- PlanLock is **template-specific by design**. It intentionally rejects alternate workbook layouts.
- Missing or ambiguous PDF data is left blank and pushed into the review report instead of being guessed.
- Structural workbook drift is treated as a hard failure.
