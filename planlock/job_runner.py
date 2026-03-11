from __future__ import annotations

from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from datetime import UTC, datetime
import json
from pathlib import Path
from queue import Empty, Queue
import threading
from typing import Iterable
from uuid import uuid4

from planlock.canonicalizer import merge_page_mappings
from planlock.calculations import run_calculation_validation
from planlock.config import Settings
from planlock.llm_pipeline import ClaudeExtractionClient, ExtractionClient, RetryNotifier
from planlock.models import (
    ImportArtifacts,
    ImportWarning,
    PageOcrResult,
    ReviewReport,
    RunEvent,
    Severity,
    Stage,
)
from planlock.pdf_renderer import RenderedPage, render_pdf_pages
from planlock.template_guard import check_for_drift
from planlock.template_schema import ALLOWED_WRITE_CELLS_BY_SHEET
from planlock.workbook_writer import (
    apply_assignments_to_workbook,
    build_assignments,
    copy_locked_template,
)


class JobRunner:
    def __init__(self, settings: Settings, extraction_client: ExtractionClient | None = None) -> None:
        self.settings = settings
        self.extraction_client = extraction_client
        self._ocr_thread_local = threading.local()

    def _client(self) -> ExtractionClient:
        if self.extraction_client is None:
            self.extraction_client = ClaudeExtractionClient(self.settings)
        return self.extraction_client

    def _job_dir(self) -> tuple[str, Path]:
        job_id = datetime.now(UTC).strftime("%Y%m%d-%H%M%S") + "-" + uuid4().hex[:8]
        job_dir = self.settings.jobs_dir / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        return job_id, job_dir

    def _ocr_thread_client(self) -> ExtractionClient:
        if self.extraction_client is not None:
            return self.extraction_client

        client = getattr(self._ocr_thread_local, "client", None)
        if client is None:
            client = ClaudeExtractionClient(self.settings)
            self._ocr_thread_local.client = client
        return client

    @staticmethod
    def _drain_retry_queue(
        retry_queue: Queue[dict[str, object]],
        *,
        completed_pages: int,
        total_pages: int,
        pipe_total: int,
    ) -> Iterable[RunEvent]:
        while True:
            try:
                payload = retry_queue.get_nowait()
            except Empty:
                break

            page_number = int(payload["page_number"])
            pipe_number = int(payload["pipe_number"])
            attempt_number = int(payload["attempt_number"])
            max_attempts = int(payload["max_attempts"])
            retry_delay_seconds = float(payload["retry_delay_seconds"])
            error_text = str(payload["error"])
            retry_reason = str(payload.get("retry_reason", "transient"))
            if retry_reason == "rate_limit":
                message = (
                    f"Anthropic rate limit hit on page {page_number}/{total_pages} in lane {pipe_number}. "
                    f"Cooling down for {retry_delay_seconds:.1f}s before pass {attempt_number}/{max_attempts}."
                )
            else:
                message = (
                    f"Running another pass on page {page_number}/{total_pages} in lane {pipe_number} "
                    f"(pass {attempt_number}/{max_attempts}) after {retry_delay_seconds:.1f}s: {error_text}"
                )
            yield RunEvent(
                stage=Stage.OCR,
                message=message,
                severity=Severity.WARNING,
                stage_completed=completed_pages,
                stage_total=total_pages,
                page_number=page_number,
                page_total=total_pages,
                pipe_number=pipe_number,
                pipe_total=pipe_total,
                attempt_number=attempt_number,
                max_attempts=max_attempts,
                retry_delay_seconds=retry_delay_seconds,
                retry_reason="rate_limit" if retry_reason == "rate_limit" else "transient",
                phase="retry",
            )

    def run(self, pdf_bytes: bytes, original_filename: str) -> Iterable[RunEvent]:
        self.settings.ensure_runtime_dirs()
        template_sha = self.settings.validate_template_lock()
        job_id, job_dir = self._job_dir()

        source_pdf = job_dir / Path(original_filename).name
        source_pdf.write_bytes(pdf_bytes)

        yield RunEvent(stage=Stage.OCR, message="Locked template verified.", stage_completed=0, stage_total=1)
        pages = render_pdf_pages(pdf_bytes, self.settings.max_pages)
        if not pages:
            raise ValueError("The uploaded PDF contains no readable pages.")

        total_pages = len(pages)
        parallel_workers = max(1, min(self.settings.ocr_parallel_workers, total_pages))
        ocr_results: list[PageOcrResult | None] = [None] * total_pages
        completed_ocr_pages = 0
        yield RunEvent(
            stage=Stage.OCR,
            message=f"Starting document review with {parallel_workers} parallel lane(s).",
            stage_completed=0,
            stage_total=total_pages,
            pipe_total=parallel_workers,
        )

        with ThreadPoolExecutor(
            max_workers=parallel_workers,
            thread_name_prefix="ocr-pipe",
        ) as executor:
            retry_queue: Queue[dict[str, object]] = Queue()

            def build_retry_notifier(
                *,
                current_page_number: int,
                current_pipe_number: int,
            ) -> RetryNotifier:
                def notify(
                    operation_name: str,
                    attempt_number: int,
                    max_attempts: int,
                    retry_delay_seconds: float,
                    error: Exception,
                ) -> None:
                    retry_queue.put(
                        {
                            "page_number": current_page_number,
                            "pipe_number": current_pipe_number,
                            "attempt_number": attempt_number,
                            "max_attempts": max_attempts,
                            "retry_delay_seconds": retry_delay_seconds,
                            "error": f"{operation_name}: {error}",
                            "retry_reason": (
                                "rate_limit"
                                if "rate_limit_error" in str(error).lower() or "rate limit" in str(error).lower()
                                else "transient"
                            ),
                        }
                    )

                return notify

            def run_ocr_task(
                current_page: RenderedPage,
                retry_notifier: RetryNotifier,
            ) -> PageOcrResult:
                return self._ocr_thread_client().ocr_page(
                    current_page,
                    retry_notifier=retry_notifier,
                )

            def start_page(
                future_map: dict[Future[PageOcrResult], tuple[int, int, RenderedPage]],
                *,
                page: RenderedPage,
                pipe_number: int,
            ) -> tuple[Future[PageOcrResult], RunEvent]:
                future = executor.submit(
                    run_ocr_task,
                    page,
                    build_retry_notifier(
                        current_page_number=page.page_number,
                        current_pipe_number=pipe_number,
                    ),
                )
                future_map[future] = (page.page_number - 1, pipe_number, page)
                return future, RunEvent(
                    stage=Stage.OCR,
                    message=f"Reviewing page {page.page_number}/{total_pages} in lane {pipe_number}.",
                    stage_completed=completed_ocr_pages,
                    stage_total=total_pages,
                    page_number=page.page_number,
                    page_total=total_pages,
                    pipe_number=pipe_number,
                    pipe_total=parallel_workers,
                    phase="start",
                )

            future_map: dict[Future[PageOcrResult], tuple[int, int, RenderedPage]] = {}
            page_iterator = iter(pages)
            for pipe_number in range(1, parallel_workers + 1):
                try:
                    page = next(page_iterator)
                except StopIteration:
                    break
                _, start_event = start_page(future_map, page=page, pipe_number=pipe_number)
                yield start_event

            pending = set(future_map)
            while pending:
                done, pending = wait(pending, timeout=0.1, return_when=FIRST_COMPLETED)
                retry_events = list(
                    self._drain_retry_queue(
                        retry_queue,
                        completed_pages=completed_ocr_pages,
                        total_pages=total_pages,
                        pipe_total=parallel_workers,
                    )
                )
                yield from retry_events
                if not done and not retry_events:
                    yield RunEvent(
                        stage=Stage.OCR,
                        message="Document review in progress.",
                        stage_completed=completed_ocr_pages,
                        stage_total=total_pages,
                        pipe_total=parallel_workers,
                        phase="heartbeat",
                    )
                for future in done:
                    result_index, pipe_number, page = future_map.pop(future)
                    try:
                        ocr_results[result_index] = future.result()
                    except Exception as exc:
                        yield RunEvent(
                            stage=Stage.OCR,
                            message=f"Page review failed for page {page.page_number}/{total_pages} in lane {pipe_number}: {exc}",
                            severity=Severity.ERROR,
                            stage_completed=completed_ocr_pages,
                            stage_total=total_pages,
                            page_number=page.page_number,
                            page_total=total_pages,
                            pipe_number=pipe_number,
                            pipe_total=parallel_workers,
                            phase="failed",
                        )
                        raise
                    completed_ocr_pages += 1
                    yield RunEvent(
                        stage=Stage.OCR,
                        message=f"Finished reviewing page {page.page_number}/{total_pages} in lane {pipe_number}.",
                        stage_completed=completed_ocr_pages,
                        stage_total=total_pages,
                        page_number=page.page_number,
                        page_total=total_pages,
                        pipe_number=pipe_number,
                        pipe_total=parallel_workers,
                        phase="complete",
                    )

                    try:
                        next_page = next(page_iterator)
                    except StopIteration:
                        continue

                    next_future, next_start_event = start_page(future_map, page=next_page, pipe_number=pipe_number)
                    pending.add(next_future)
                    yield next_start_event

            yield from self._drain_retry_queue(
                retry_queue,
                completed_pages=completed_ocr_pages,
                total_pages=total_pages,
                pipe_total=parallel_workers,
            )

        client = self._client()
        final_ocr_results = [result for result in ocr_results if result is not None]
        if len(final_ocr_results) != total_pages:
            raise RuntimeError("Document review finished without results for every page.")

        mapping_results = []
        mapping_total = total_pages + 2
        yield RunEvent(
            stage=Stage.DATA_ENTRY,
            message="Beginning workbook entry.",
            stage_completed=0,
            stage_total=mapping_total,
        )
        for index, (page, ocr_result) in enumerate(zip(pages, final_ocr_results), start=1):
            mapping_results.append(client.map_page(page, ocr_result))
            yield RunEvent(
                stage=Stage.DATA_ENTRY,
                message=f"Entered page {index}/{total_pages} into the workbook.",
                stage_completed=index,
                stage_total=mapping_total,
            )

        canonical_document, merge_warnings = merge_page_mappings(mapping_results)
        yield RunEvent(
            stage=Stage.DATA_ENTRY,
            message="Combined page entries into one household record.",
            stage_completed=total_pages + 1,
            stage_total=mapping_total,
        )

        workbook_path = copy_locked_template(self.settings, job_dir)
        assignments, write_warnings = build_assignments(canonical_document)
        apply_assignments_to_workbook(workbook_path, assignments)
        yield RunEvent(
            stage=Stage.DATA_ENTRY,
            message=f"Wrote {len(assignments)} values into the locked workbook copy.",
            stage_completed=mapping_total,
            stage_total=mapping_total,
        )

        calc_total = 3
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Running workbook checks.",
            stage_completed=1,
            stage_total=calc_total,
        )
        calculation_validation = run_calculation_validation(self.settings.template_path, workbook_path)
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Formula review complete.",
            stage_completed=2,
            stage_total=calc_total,
        )
        drift_check = check_for_drift(
            self.settings.template_path,
            workbook_path,
            ALLOWED_WRITE_CELLS_BY_SHEET,
        )
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Workbook structure review complete.",
            stage_completed=3,
            stage_total=calc_total,
        )

        warnings: list[ImportWarning] = [*merge_warnings, *write_warnings]
        warnings.extend(
            ImportWarning(
                code="calculation_warning",
                message=warning,
                severity=Severity.WARNING,
                stage=Stage.FINANCIAL_CALCULATIONS,
            )
            for warning in calculation_validation.warnings
        )
        warnings.extend(
            ImportWarning(
                code="drift_violation",
                message=violation,
                severity=Severity.ERROR,
                stage=Stage.FINANCIAL_CALCULATIONS,
            )
            for violation in drift_check.violations
        )

        success = calculation_validation.passed and drift_check.passed
        report = ReviewReport(
            job_id=job_id,
            template_sha256=template_sha,
            success=success,
            warnings=warnings,
            mapped_assignments=assignments,
            unmapped_items=canonical_document.unmapped_items,
            assumptions=canonical_document.assumptions,
            drift_check=drift_check,
            calculation_validation=calculation_validation,
        )
        report_path = job_dir / self.settings.review_report_name
        report_path.write_text(json.dumps(report.model_dump(mode="json"), indent=2), encoding="utf-8")

        artifacts = ImportArtifacts(
            success=success,
            job_id=job_id,
            job_dir=job_dir,
            output_workbook_path=workbook_path if success else None,
            review_report_path=report_path,
            review_report=report,
        )
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Workbook is ready." if success else "Run failed because the workbook structure or formulas need attention.",
            severity=Severity.INFO if success else Severity.ERROR,
            stage_completed=calc_total,
            stage_total=calc_total,
            artifacts=artifacts,
        )
