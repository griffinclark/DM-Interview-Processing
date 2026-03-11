from __future__ import annotations

import hashlib
import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

from planlock.throttle import RequestThrottleCoordinator


PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_FILENAME = (
    "Copy of FP Case Study Model - Alicia and Tom Smith - March 5, 10_01\u202fAM.xlsx"
)
LLM_PROVIDER_OPENAI = "openai"
LLM_PROVIDER_ANTHROPIC = "anthropic"
LLM_PROVIDER_OPTIONS = (LLM_PROVIDER_OPENAI, LLM_PROVIDER_ANTHROPIC)
DEFAULT_LLM_PROVIDER = LLM_PROVIDER_OPENAI
LOCKED_PROVIDER_MODELS = {
    LLM_PROVIDER_OPENAI: "gpt-5.2",
    LLM_PROVIDER_ANTHROPIC: "claude-sonnet-4-5-20250929",
}
DEFAULT_MODEL_OCR = LOCKED_PROVIDER_MODELS[DEFAULT_LLM_PROVIDER]
DEFAULT_MODEL_MAPPING = LOCKED_PROVIDER_MODELS[DEFAULT_LLM_PROVIDER]
DEFAULT_OCR_PARALLEL_WORKERS = 3
DEFAULT_LLM_TIMEOUT_SECONDS = 120.0
DEFAULT_LLM_MAX_RETRIES = 2
DEFAULT_LLM_RETRY_BASE_SECONDS = 10.0
DEFAULT_LLM_RETRY_MAX_SECONDS = 300.0
DEFAULT_MAX_PAGES = 40
DEFAULT_LOG_LEVEL = "INFO"
DEFAULT_MIN_SUPPORTED_COVERAGE = 0.70
DEFAULT_MAX_UNRESOLVED_SUPPORTED_TARGETS = 10
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


def normalize_llm_provider(value: str | None) -> str:
    provider = (value or DEFAULT_LLM_PROVIDER).strip().lower()
    if provider not in LLM_PROVIDER_OPTIONS:
        raise ValueError(
            f"Unsupported LLM provider '{provider}'. Expected one of: {', '.join(LLM_PROVIDER_OPTIONS)}."
        )
    return provider


def provider_display_name(provider: str | None) -> str:
    normalized = normalize_llm_provider(provider)
    if normalized == LLM_PROVIDER_OPENAI:
        return "OpenAI"
    return "Anthropic"


def locked_model_for_provider(provider: str | None) -> str:
    return LOCKED_PROVIDER_MODELS[normalize_llm_provider(provider)]


@dataclass(frozen=True)
class Settings:
    anthropic_api_key: str | None
    openai_api_key: str | None
    llm_provider: str
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
    debug_cache_path: Path
    workbook_output_name: str = "filled_financial_plan.xlsx"
    review_report_name: str = "review_report.json"
    ocr_results_name: str = "ocr_results.json"
    entry_state_name: str = "entry_state.json"
    min_supported_coverage: float = DEFAULT_MIN_SUPPORTED_COVERAGE
    max_unresolved_supported_targets: int = DEFAULT_MAX_UNRESOLVED_SUPPORTED_TARGETS
    request_throttle: RequestThrottleCoordinator = field(
        default_factory=RequestThrottleCoordinator,
        compare=False,
        repr=False,
    )

    @classmethod
    def from_env(cls) -> "Settings":
        default_model = locked_model_for_provider(DEFAULT_LLM_PROVIDER)
        return cls(
            anthropic_api_key=os.getenv("ANTHROPIC_API_KEY"),
            openai_api_key=os.getenv("OPENAI_API_KEY"),
            llm_provider=DEFAULT_LLM_PROVIDER,
            model_ocr=default_model,
            model_mapping=default_model,
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
            debug_cache_path=PROJECT_ROOT / "debug_cache.txt",
            min_supported_coverage=DEFAULT_MIN_SUPPORTED_COVERAGE,
            max_unresolved_supported_targets=DEFAULT_MAX_UNRESOLVED_SUPPORTED_TARGETS,
        )

    def normalized_llm_provider(self) -> str:
        return normalize_llm_provider(self.llm_provider)

    def llm_provider_display_name(self) -> str:
        return provider_display_name(self.llm_provider)

    def locked_model_name(self) -> str:
        return locked_model_for_provider(self.llm_provider)

    def llm_api_key(self) -> str | None:
        if self.normalized_llm_provider() == LLM_PROVIDER_OPENAI:
            return self.openai_api_key
        return self.anthropic_api_key

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
