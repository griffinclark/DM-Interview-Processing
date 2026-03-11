# PlanLock

PlanLock is a locked workbook intake app for financial planners. It imports planner-facing PDF financial plans into the exact Excel workbook template shipped in this repository, enforces a no-drift structure policy, and exposes a three-stage pipeline:

1. OCR
2. Data Entry
3. Financial Calculations

## License

This repository is intentionally **non-commercial only** for interview and evaluation use.

- License notice: [LICENSE](/Users/griffin/Desktop/DM Interview Processing/LICENSE)
- Required notice: [NOTICE](/Users/griffin/Desktop/DM Interview Processing/NOTICE)

If you need commercial rights, you must secure a separate written license from the copyright holder.

## What It Does

- Uses the exact workbook template at [assets/template/Copy of FP Case Study Model - Alicia and Tom Smith - March 5, 10_01 AM.xlsx](/Users/griffin/Desktop/DM Interview Processing/assets/template/Copy%20of%20FP%20Case%20Study%20Model%20-%20Alicia%20and%20Tom%20Smith%20-%20March%205,%2010_01%E2%80%AFAM.xlsx)
- Validates the template SHA-256 before startup
- Runs OCR across parallel lanes with retries, immediately refilling each lane as pages finish
- Writes only to whitelisted workbook cells and repeating row blocks
- Preserves formulas, styles, sheet order, named ranges, validations, and layout
- Generates a sidecar `review_report.json` for planner review
- Streams logs and three-stage progress in the Streamlit UI

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
ANTHROPIC_API_KEY="your-key"
```

Runtime settings now live in the app’s `Settings` modal. The defaults are:

```bash
MODEL_OCR="claude-sonnet-4-6"
MODEL_MAPPING="claude-opus-4-6"
OCR_PARALLEL_WORKERS="3"
LLM_TIMEOUT_SECONDS="120"
LLM_MAX_RETRIES="2"
LLM_RETRY_BASE_SECONDS="2"
LLM_RETRY_MAX_SECONDS="12"
MAX_PAGES="40"
LOG_LEVEL="INFO"
```

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
  -e ANTHROPIC_API_KEY="$ANTHROPIC_API_KEY" \
  planlock
```

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
