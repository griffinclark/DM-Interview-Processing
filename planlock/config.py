from __future__ import annotations

import hashlib
import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_FILENAME = (
    "Copy of FP Case Study Model - Alicia and Tom Smith - March 5, 10_01\u202fAM.xlsx"
)
DEFAULT_MODEL_OCR = "claude-sonnet-4-6"
DEFAULT_MODEL_MAPPING = "claude-opus-4-6"
DEFAULT_OCR_PARALLEL_WORKERS = 3
DEFAULT_LLM_TIMEOUT_SECONDS = 120.0
DEFAULT_LLM_MAX_RETRIES = 2
DEFAULT_LLM_RETRY_BASE_SECONDS = 2.0
DEFAULT_LLM_RETRY_MAX_SECONDS = 12.0
DEFAULT_MAX_PAGES = 40
DEFAULT_LOG_LEVEL = "INFO"
DEFAULT_TEMPLATE_PATH = PROJECT_ROOT / "assets" / "template" / TEMPLATE_FILENAME
DEFAULT_TEMPLATE_SHA256 = (
    "98e7c1a9d8015c723ad1e020d070f9ae5c6c4b78b6a920f8487a7108ec98f917"
)
DEFAULT_SAMPLE_PDF = PROJECT_ROOT / "assets" / "samples" / "Domain Money Mock Financial Plan.pdf"
DEFAULT_ENV_PATH = PROJECT_ROOT / ".env"
DEFAULT_EXAMPLE_ENV_PATH = PROJECT_ROOT / "example.env"


load_dotenv(DEFAULT_EXAMPLE_ENV_PATH, override=False)
load_dotenv(DEFAULT_ENV_PATH, override=True)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


@dataclass(frozen=True)
class Settings:
    anthropic_api_key: str | None
    model_ocr: str
    model_mapping: str
    ocr_parallel_workers: int
    llm_timeout_seconds: float
    llm_max_retries: int
    llm_retry_base_seconds: float
    llm_retry_max_seconds: float
    template_path: Path
    template_sha256: str
    max_pages: int
    log_level: str
    jobs_dir: Path
    sample_pdf_path: Path
    workbook_output_name: str = "filled_financial_plan.xlsx"
    review_report_name: str = "review_report.json"

    @classmethod
    def from_env(cls) -> "Settings":
        return cls(
            anthropic_api_key=os.getenv("ANTHROPIC_API_KEY"),
            model_ocr=DEFAULT_MODEL_OCR,
            model_mapping=DEFAULT_MODEL_MAPPING,
            ocr_parallel_workers=DEFAULT_OCR_PARALLEL_WORKERS,
            llm_timeout_seconds=DEFAULT_LLM_TIMEOUT_SECONDS,
            llm_max_retries=DEFAULT_LLM_MAX_RETRIES,
            llm_retry_base_seconds=DEFAULT_LLM_RETRY_BASE_SECONDS,
            llm_retry_max_seconds=DEFAULT_LLM_RETRY_MAX_SECONDS,
            template_path=DEFAULT_TEMPLATE_PATH,
            template_sha256=DEFAULT_TEMPLATE_SHA256,
            max_pages=DEFAULT_MAX_PAGES,
            log_level=DEFAULT_LOG_LEVEL,
            jobs_dir=PROJECT_ROOT / "tmp" / "jobs",
            sample_pdf_path=DEFAULT_SAMPLE_PDF,
        )

    def ensure_runtime_dirs(self) -> None:
        self.jobs_dir.mkdir(parents=True, exist_ok=True)

    def validate_template_lock(self) -> str:
        if not self.template_path.exists():
            raise FileNotFoundError(f"Locked template not found: {self.template_path}")

        actual = sha256_file(self.template_path)
        if actual != self.template_sha256:
            raise ValueError(
                "Locked template checksum mismatch. "
                f"Expected {self.template_sha256}, found {actual}."
            )
        return actual
