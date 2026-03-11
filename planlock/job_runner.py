from __future__ import annotations

from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from datetime import UTC, datetime
import inspect
import json
from pathlib import Path
from queue import Empty, Queue
import threading
from typing import Iterable
from uuid import uuid4

from planlock.calculations import run_calculation_validation
from planlock.canonicalizer import merge_page_mappings
from planlock.config import Settings
from planlock.llm_pipeline import (
    describe_retry_error,
    ExtractionClient,
    OpenAICompatibleExtractionClient,
    RetryNotifier,
    is_rate_limit_error,
    retry_reason_for_error,
)
from planlock.models import (
    EntrySessionState,
    ImportArtifacts,
    ImportWarning,
    PageOcrResult,
    ReviewReport,
    RunEvent,
    Severity,
    Stage,
)
from planlock.pdf_renderer import RenderedPage, render_pdf_pages
from planlock.template_entry_agent import (
    LangGraphTemplateEntryAgent,
    TemplateEntryAgent,
    answer_from_question,
    coverage_summary_for_state,
    default_sheet_order,
    load_entry_state,
    load_ocr_results,
    persist_ocr_results,
    save_entry_state,
    sheet_result_to_page_mapping_result,
    touched_cells_for_assignments,
)
from planlock.template_guard import check_for_drift
from planlock.template_schema import ALLOWED_WRITE_CELLS_BY_SHEET
from planlock.workbook_writer import apply_assignments_to_workbook, build_assignments, copy_locked_template


class JobRunner:
    def __init__(
        self,
        settings: Settings,
        extraction_client: ExtractionClient | None = None,
        entry_agent: TemplateEntryAgent | None = None,
    ) -> None:
        self.settings = settings
        self.extraction_client = extraction_client
        self.entry_agent = entry_agent
        self._ocr_thread_local = threading.local()

    def _client(self) -> ExtractionClient:
        if self.extraction_client is None:
            self.extraction_client = OpenAICompatibleExtractionClient(self.settings)
        return self.extraction_client

    def _entry_agent(self) -> TemplateEntryAgent:
        if self.entry_agent is None:
            self.entry_agent = LangGraphTemplateEntryAgent(self.settings)
        return self.entry_agent

    def _job_dir(self) -> tuple[str, Path]:
        job_id = datetime.now(UTC).strftime("%Y%m%d-%H%M%S") + "-" + uuid4().hex[:8]
        job_dir = self.settings.jobs_dir / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        return job_id, job_dir

    def _job_dir_for(self, job_id: str) -> Path:
        job_dir = self.settings.jobs_dir / job_id
        if not job_dir.exists():
            raise FileNotFoundError(f"Unknown job id: {job_id}")
        return job_dir

    def _ocr_thread_client(self) -> ExtractionClient:
        if self.extraction_client is not None:
            return self.extraction_client

        client = getattr(self._ocr_thread_local, "client", None)
        if client is None:
            client = OpenAICompatibleExtractionClient(self.settings)
            self._ocr_thread_local.client = client
        return client

    @staticmethod
    def _drain_retry_queue(
        retry_queue: Queue[dict[str, object]],
        *,
        provider_label: str = "OpenAI",
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
            detail_message = str(payload["detail_message"])
            retry_reason = str(payload.get("retry_reason", "transient"))
            current_timeout_seconds = payload.get("current_timeout_seconds")
            next_timeout_seconds = payload.get("next_timeout_seconds")
            if retry_reason == "rate_limit":
                message = (
                    f"{provider_label} rate limit hit on page {page_number}/{total_pages} in lane {pipe_number}. "
                    f"Cooling down for {retry_delay_seconds:.1f}s before pass {attempt_number}/{max_attempts}."
                )
            elif retry_reason == "timeout":
                current_timeout_label = (
                    f"{float(current_timeout_seconds):.0f}s"
                    if isinstance(current_timeout_seconds, (int, float))
                    else "the current"
                )
                next_timeout_label = (
                    f"{float(next_timeout_seconds):.0f}s"
                    if isinstance(next_timeout_seconds, (int, float))
                    else "a higher"
                )
                message = (
                    f"Page {page_number}/{total_pages} in lane {pipe_number} exceeded the {current_timeout_label} processing limit. "
                    f"Retrying immediately with {next_timeout_label} timeout (pass {attempt_number}/{max_attempts})."
                )
            else:
                message = (
                    f"Restarting lane {pipe_number} for page {page_number}/{total_pages} "
                    f"(pass {attempt_number}/{max_attempts}) after {retry_delay_seconds:.1f}s."
                )
            yield RunEvent(
                stage=Stage.OCR,
                message=message,
                detail_message=detail_message,
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
                retry_reason=(
                    "rate_limit"
                    if retry_reason == "rate_limit"
                    else "timeout" if retry_reason == "timeout" else "transient"
                ),
                phase="retry",
            )

    def _ocr_results_path(self, job_dir: Path) -> Path:
        return job_dir / self.settings.ocr_results_name

    def _entry_state_path(self, job_dir: Path) -> Path:
        return job_dir / self.settings.entry_state_name

    def _report_path(self, job_dir: Path) -> Path:
        return job_dir / self.settings.review_report_name

    @staticmethod
    def _drain_entry_retry_queue(
        retry_queue: Queue[dict[str, object]],
        *,
        provider_label: str = "OpenAI",
        current_sheet_name: str,
        stage_completed: int,
        stage_total: int,
        agent_total_tokens: int | None = None,
    ) -> Iterable[RunEvent]:
        while True:
            try:
                payload = retry_queue.get_nowait()
            except Empty:
                break

            attempt_number = int(payload["attempt_number"])
            max_attempts = int(payload["max_attempts"])
            retry_delay_seconds = float(payload["retry_delay_seconds"])
            detail_message = str(payload["detail_message"])
            retry_reason = str(payload.get("retry_reason", "transient"))
            current_timeout_seconds = payload.get("current_timeout_seconds")
            next_timeout_seconds = payload.get("next_timeout_seconds")
            if retry_reason == "rate_limit":
                message = (
                    f"{provider_label} rate limit hit while filling {current_sheet_name}. "
                    f"Resuming pass {attempt_number}/{max_attempts} in {retry_delay_seconds:.1f}s."
                )
            elif retry_reason == "timeout":
                current_timeout_label = (
                    f"{float(current_timeout_seconds):.0f}s"
                    if isinstance(current_timeout_seconds, (int, float))
                    else "the current"
                )
                next_timeout_label = (
                    f"{float(next_timeout_seconds):.0f}s"
                    if isinstance(next_timeout_seconds, (int, float))
                    else "a higher"
                )
                message = (
                    f"{current_sheet_name} exceeded the {current_timeout_label} processing limit. "
                    f"Retrying immediately with {next_timeout_label} timeout (pass {attempt_number}/{max_attempts})."
                )
            else:
                message = (
                    f"Restarting workbook entry for {current_sheet_name} "
                    f"(pass {attempt_number}/{max_attempts}) after {retry_delay_seconds:.1f}s."
                )
            yield RunEvent(
                stage=Stage.DATA_ENTRY,
                message=message,
                sheet_name=current_sheet_name,
                agent_total_tokens=agent_total_tokens,
                detail_message=detail_message,
                severity=Severity.WARNING,
                stage_completed=stage_completed,
                stage_total=stage_total,
                attempt_number=attempt_number,
                max_attempts=max_attempts,
                retry_delay_seconds=retry_delay_seconds,
                retry_reason=(
                    "rate_limit"
                    if retry_reason == "rate_limit"
                    else "timeout" if retry_reason == "timeout" else "transient"
                ),
                phase="retry",
            )

    @staticmethod
    def _drain_entry_progress_queue(progress_queue: Queue[dict[str, object]]) -> str | None:
        latest_progress_message: str | None = None
        while True:
            try:
                payload = progress_queue.get_nowait()
            except Empty:
                break

            progress_message = str(payload.get("progress_message") or "").strip()
            if progress_message:
                latest_progress_message = progress_message
        return latest_progress_message

    def _run_ocr(self, pdf_bytes: bytes) -> Iterable[RunEvent]:
        provider_label = self.settings.llm_provider_display_name()
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
                            "detail_message": f"{operation_name}: {describe_retry_error(error)}",
                            "retry_reason": retry_reason_for_error(error),
                            "current_timeout_seconds": getattr(error, "planlock_timeout_seconds", None),
                            "next_timeout_seconds": getattr(error, "planlock_next_timeout_seconds", None),
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
                        provider_label=provider_label,
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
                        message = f"Page review failed for page {page.page_number}/{total_pages} in lane {pipe_number}."
                        yield RunEvent(
                            stage=Stage.OCR,
                            message=message,
                            detail_message=str(exc),
                            severity=Severity.ERROR,
                            stage_completed=completed_ocr_pages,
                            stage_total=total_pages,
                            page_number=page.page_number,
                            page_total=total_pages,
                            pipe_number=pipe_number,
                            pipe_total=parallel_workers,
                            phase="failed",
                        )
                        raise RuntimeError(message) from exc
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
                provider_label=provider_label,
                completed_pages=completed_ocr_pages,
                total_pages=total_pages,
                pipe_total=parallel_workers,
            )

        final_ocr_results = [result for result in ocr_results if result is not None]
        if len(final_ocr_results) != total_pages:
            raise RuntimeError("Document review finished without results for every page.")
        return final_ocr_results

    def _rebuild_workbook(self, workbook_path: Path, assignments) -> None:
        workbook_path.write_bytes(self.settings.template_path.read_bytes())
        if assignments:
            apply_assignments_to_workbook(workbook_path, assignments)

    def _refresh_assignment_state(self, state: EntrySessionState) -> None:
        page_results = [sheet_result_to_page_mapping_result(result) for result in state.sheet_results]
        canonical_document, _ = merge_page_mappings(page_results)
        assignments, _ = build_assignments(canonical_document)
        self._rebuild_workbook(state.workbook_path, assignments)
        state.mapped_assignments = assignments
        state.unmapped_items = canonical_document.unmapped_items
        state.assumptions = canonical_document.assumptions
        for summary in state.sheet_summaries:
            summary.touched_cells = touched_cells_for_assignments(assignments, summary.sheet_name)

    def _entry_artifacts(
        self,
        *,
        state: EntrySessionState,
        job_dir: Path,
        pending_question=None,
    ) -> ImportArtifacts:
        return ImportArtifacts(
            success=False,
            job_id=state.job_id,
            job_dir=job_dir,
            output_workbook_path=state.workbook_path if state.workbook_path.exists() else None,
            entry_state_path=self._entry_state_path(job_dir),
            ocr_results_path=state.ocr_results_path,
            pending_question=pending_question,
        )

    def _pause_event(
        self,
        *,
        state: EntrySessionState,
        job_dir: Path,
        message: str,
        agent_total_tokens: int | None = None,
    ) -> RunEvent:
        return RunEvent(
            stage=Stage.DATA_ENTRY,
            message=message,
            sheet_name=state.pending_question.sheet_name if state.pending_question is not None else None,
            agent_total_tokens=agent_total_tokens,
            severity=Severity.WARNING,
            stage_completed=state.current_sheet_index,
            stage_total=len(state.sheet_order),
            phase="paused",
            artifacts=self._entry_artifacts(state=state, job_dir=job_dir, pending_question=state.pending_question),
        )

    def _finalize_job(
        self,
        *,
        job_dir: Path,
        state: EntrySessionState,
    ) -> Iterable[RunEvent]:
        page_results = [sheet_result_to_page_mapping_result(result) for result in state.sheet_results]
        canonical_document, merge_warnings = merge_page_mappings(page_results)
        assignments, write_warnings = build_assignments(canonical_document)
        self._rebuild_workbook(state.workbook_path, assignments)
        state.mapped_assignments = assignments
        state.unmapped_items = canonical_document.unmapped_items
        state.assumptions = canonical_document.assumptions
        state.coverage_summary = coverage_summary_for_state(state)

        calc_total = 3
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Running workbook checks.",
            stage_completed=1,
            stage_total=calc_total,
        )
        calculation_validation = run_calculation_validation(self.settings.template_path, state.workbook_path)
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Formula review complete.",
            stage_completed=2,
            stage_total=calc_total,
        )
        drift_check = check_for_drift(
            self.settings.template_path,
            state.workbook_path,
            ALLOWED_WRITE_CELLS_BY_SHEET,
        )
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message="Workbook structure review complete.",
            stage_completed=3,
            stage_total=calc_total,
        )

        review_required_reasons: list[str] = []
        if state.coverage_summary.unresolved_critical_sheet_names:
            review_required_reasons.append(
                "Critical sheets still have unresolved supported targets: "
                + ", ".join(state.coverage_summary.unresolved_critical_sheet_names)
            )
        if state.coverage_summary.coverage_ratio < self.settings.min_supported_coverage:
            review_required_reasons.append(
                "Supported target coverage fell below "
                f"{self.settings.min_supported_coverage:.0%}."
            )
        if (
            state.coverage_summary.unresolved_supported_target_count
            > self.settings.max_unresolved_supported_targets
        ):
            review_required_reasons.append(
                "Supported unresolved targets exceeded "
                f"{self.settings.max_unresolved_supported_targets}."
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
        warnings.extend(
            ImportWarning(
                code="review_required",
                message=reason,
                severity=Severity.WARNING,
                stage=Stage.DATA_ENTRY,
            )
            for reason in review_required_reasons
        )

        state.review_required_reasons = review_required_reasons
        valid_workbook = calculation_validation.passed and drift_check.passed
        success = valid_workbook and not review_required_reasons
        report = ReviewReport(
            job_id=state.job_id,
            template_sha256=state.template_sha256,
            success=success,
            warnings=warnings,
            mapped_assignments=assignments,
            unmapped_items=canonical_document.unmapped_items,
            assumptions=canonical_document.assumptions,
            sheet_summaries=state.sheet_summaries,
            user_answers=state.user_answers,
            questions_asked=state.questions_asked,
            coverage_summary=state.coverage_summary,
            review_required_reasons=review_required_reasons,
            drift_check=drift_check,
            calculation_validation=calculation_validation,
        )
        report_path = self._report_path(job_dir)
        report_path.write_text(json.dumps(report.model_dump(mode="json"), indent=2), encoding="utf-8")
        save_entry_state(self._entry_state_path(job_dir), state)

        artifacts = ImportArtifacts(
            success=success,
            job_id=state.job_id,
            job_dir=job_dir,
            output_workbook_path=state.workbook_path if valid_workbook else None,
            review_report_path=report_path,
            review_report=report,
            entry_state_path=self._entry_state_path(job_dir),
            ocr_results_path=state.ocr_results_path,
        )
        if not valid_workbook:
            message = "Run failed because the workbook structure or formulas need attention."
            severity = Severity.ERROR
        elif review_required_reasons:
            message = "Workbook entry finished, but planner review is required."
            severity = Severity.WARNING
        else:
            message = "Workbook is ready."
            severity = Severity.INFO
        yield RunEvent(
            stage=Stage.FINANCIAL_CALCULATIONS,
            message=message,
            severity=severity,
            stage_completed=calc_total,
            stage_total=calc_total,
            phase="complete",
            artifacts=artifacts,
        )

    def _advance_entry_session(self, job_dir: Path) -> Iterable[RunEvent]:
        state_path = self._entry_state_path(job_dir)
        state = load_entry_state(state_path)
        ocr_results = load_ocr_results(state.ocr_results_path)
        stage_total = len(state.sheet_order)
        provider_label = self.settings.llm_provider_display_name()
        agent_total_tokens = 0
        usage_lock = threading.Lock()

        def total_agent_tokens() -> int:
            with usage_lock:
                return agent_total_tokens

        yield RunEvent(
            stage=Stage.DATA_ENTRY,
            message="Beginning workbook entry.",
            agent_total_tokens=total_agent_tokens(),
            stage_completed=state.current_sheet_index,
            stage_total=stage_total,
            artifacts=self._entry_artifacts(state=state, job_dir=job_dir),
        )

        while not state.completed and state.pending_question is None:
            current_sheet_name = state.sheet_order[state.current_sheet_index]
            retry_queue: Queue[dict[str, object]] = Queue()
            progress_queue: Queue[dict[str, object]] = Queue()
            entry_agent = self._entry_agent()

            yield RunEvent(
                stage=Stage.DATA_ENTRY,
                message=f"LangGraph is working on {current_sheet_name}.",
                sheet_name=current_sheet_name,
                agent_total_tokens=total_agent_tokens(),
                detail_message="Reviewing OCR evidence, workbook context, and prior answers before proposing writes.",
                stage_completed=state.current_sheet_index,
                stage_total=stage_total,
                phase="start",
                artifacts=self._entry_artifacts(state=state, job_dir=job_dir),
            )

            def notify_retry(
                operation_name: str,
                attempt_number: int,
                max_attempts: int,
                retry_delay_seconds: float,
                error: Exception,
            ) -> None:
                retry_queue.put(
                    {
                        "attempt_number": attempt_number,
                        "max_attempts": max_attempts,
                        "retry_delay_seconds": retry_delay_seconds,
                        "detail_message": f"{operation_name}: {describe_retry_error(error)}",
                        "retry_reason": retry_reason_for_error(error),
                        "current_timeout_seconds": getattr(error, "planlock_timeout_seconds", None),
                        "next_timeout_seconds": getattr(error, "planlock_next_timeout_seconds", None),
                    }
                )

            def notify_usage(_operation_name: str, usage: dict[str, int]) -> None:
                nonlocal agent_total_tokens
                total_tokens = int(usage.get("total_tokens") or 0)
                if total_tokens <= 0:
                    return
                with usage_lock:
                    agent_total_tokens += total_tokens

            def notify_progress(_operation_name: str, progress_message: str) -> None:
                progress_text = str(progress_message).strip()
                if not progress_text:
                    return
                progress_queue.put({"progress_message": progress_text})

            advance_kwargs: dict[str, object] = {"retry_notifier": notify_retry}
            try:
                advance_parameters = inspect.signature(entry_agent.advance).parameters
            except (TypeError, ValueError):
                advance_parameters = {}
            accepts_var_kwargs = any(
                parameter.kind == inspect.Parameter.VAR_KEYWORD
                for parameter in advance_parameters.values()
            )
            if accepts_var_kwargs or "usage_notifier" in advance_parameters:
                advance_kwargs["usage_notifier"] = notify_usage
            if accepts_var_kwargs or "progress_notifier" in advance_parameters:
                advance_kwargs["progress_notifier"] = notify_progress

            with ThreadPoolExecutor(max_workers=1, thread_name_prefix="entry-pass") as executor:
                future = executor.submit(
                    entry_agent.advance,
                    state,
                    ocr_results,
                    **advance_kwargs,
                )
                pending = {future}
                while pending:
                    done, pending = wait(pending, timeout=0.1, return_when=FIRST_COMPLETED)
                    retry_events = list(
                        self._drain_entry_retry_queue(
                            retry_queue,
                            provider_label=provider_label,
                            current_sheet_name=current_sheet_name,
                            stage_completed=state.current_sheet_index,
                            stage_total=stage_total,
                            agent_total_tokens=total_agent_tokens(),
                        )
                    )
                    latest_progress_message = self._drain_entry_progress_queue(progress_queue)
                    if retry_events:
                        yield from retry_events
                        continue
                    if latest_progress_message is not None:
                        yield RunEvent(
                            stage=Stage.DATA_ENTRY,
                            message="Workbook entry in progress.",
                            sheet_name=current_sheet_name,
                            agent_total_tokens=total_agent_tokens(),
                            progress_message=latest_progress_message,
                            stage_completed=state.current_sheet_index,
                            stage_total=stage_total,
                            phase="heartbeat",
                        )
                    if not done and latest_progress_message is None:
                        yield RunEvent(
                            stage=Stage.DATA_ENTRY,
                            message="Workbook entry in progress.",
                            sheet_name=current_sheet_name,
                            agent_total_tokens=total_agent_tokens(),
                            stage_completed=state.current_sheet_index,
                            stage_total=stage_total,
                            phase="heartbeat",
                        )
                        continue

                    if not done:
                        continue

                    try:
                        state = future.result()
                    except Exception as exc:
                        root_cause = exc.__cause__ or exc
                        if is_rate_limit_error(root_cause) or is_rate_limit_error(exc):
                            message = (
                                f"Workbook entry stopped on {current_sheet_name} after repeated {provider_label} rate limits."
                            )
                        else:
                            message = f"Workbook entry failed on {current_sheet_name}."
                        yield RunEvent(
                            stage=Stage.DATA_ENTRY,
                            message=message,
                            sheet_name=current_sheet_name,
                            agent_total_tokens=total_agent_tokens(),
                            detail_message=str(root_cause),
                            severity=Severity.ERROR,
                            stage_completed=state.current_sheet_index,
                            stage_total=stage_total,
                            phase="failed",
                        )
                        raise RuntimeError(message) from exc

                yield from self._drain_entry_retry_queue(
                    retry_queue,
                    provider_label=provider_label,
                    current_sheet_name=current_sheet_name,
                    stage_completed=state.current_sheet_index,
                    stage_total=stage_total,
                    agent_total_tokens=total_agent_tokens(),
                )

            self._refresh_assignment_state(state)
            save_entry_state(state_path, state)

            if state.pending_question is not None:
                yield self._pause_event(
                    state=state,
                    job_dir=job_dir,
                    message=f"Workbook entry paused on {current_sheet_name}: {state.pending_question.prompt}",
                    agent_total_tokens=total_agent_tokens(),
                )
                return

            yield RunEvent(
                stage=Stage.DATA_ENTRY,
                message=f"Completed sheet {current_sheet_name}.",
                sheet_name=current_sheet_name,
                agent_total_tokens=total_agent_tokens(),
                stage_completed=state.current_sheet_index,
                stage_total=stage_total,
            )

        if state.completed:
            yield RunEvent(
                stage=Stage.DATA_ENTRY,
                message="Workbook entry complete.",
                agent_total_tokens=total_agent_tokens(),
                stage_completed=stage_total,
                stage_total=stage_total,
                phase="complete",
            )
            yield from self._finalize_job(job_dir=job_dir, state=state)

    def run(self, pdf_bytes: bytes, original_filename: str) -> Iterable[RunEvent]:
        yield from self.start_job(pdf_bytes, original_filename)

    def _start_entry_job(
        self,
        *,
        job_id: str,
        job_dir: Path,
        template_sha: str,
        ocr_results: list[PageOcrResult],
    ) -> Iterable[RunEvent]:
        if not ocr_results:
            raise RuntimeError("Phase one completed without persisted results.")

        workbook_path = copy_locked_template(self.settings, job_dir)
        ocr_results_path = self._ocr_results_path(job_dir)
        persist_ocr_results(ocr_results_path, ocr_results)
        state = EntrySessionState(
            job_id=job_id,
            template_sha256=template_sha,
            workbook_path=workbook_path,
            ocr_results_path=ocr_results_path,
            sheet_order=default_sheet_order(),
        )
        save_entry_state(self._entry_state_path(job_dir), state)
        yield from self._advance_entry_session(job_dir)

    def start_job(self, pdf_bytes: bytes, original_filename: str) -> Iterable[RunEvent]:
        self.settings.ensure_runtime_dirs()
        template_sha = self.settings.validate_template_lock()
        job_id, job_dir = self._job_dir()

        source_pdf = job_dir / Path(original_filename).name
        source_pdf.write_bytes(pdf_bytes)

        ocr_results = yield from self._run_ocr(pdf_bytes)
        yield from self._start_entry_job(
            job_id=job_id,
            job_dir=job_dir,
            template_sha=template_sha,
            ocr_results=ocr_results,
        )

    def resume_job(
        self,
        job_id: str,
        answer: str,
        *,
        source: str = "option",
    ) -> Iterable[RunEvent]:
        self.settings.ensure_runtime_dirs()
        job_dir = self._job_dir_for(job_id)
        state_path = self._entry_state_path(job_dir)
        state = load_entry_state(state_path)
        if state.pending_question is None:
            raise ValueError("This job is not waiting on a user question.")

        state.user_answers.append(
            answer_from_question(
                state.pending_question,
                answer=answer,
                source=source,
            )
        )
        answered_sheet = state.pending_question.sheet_name
        state.pending_question = None
        save_entry_state(state_path, state)
        resume_message = (
            f"Letting the agent resolve {answered_sheet}. Resuming workbook entry."
            if source == "agent"
            else f"Captured answer for {answered_sheet}. Resuming workbook entry."
        )
        yield RunEvent(
            stage=Stage.DATA_ENTRY,
            message=resume_message,
            sheet_name=answered_sheet,
            stage_completed=state.current_sheet_index,
            stage_total=len(state.sheet_order),
        )
        yield from self._advance_entry_session(job_dir)
