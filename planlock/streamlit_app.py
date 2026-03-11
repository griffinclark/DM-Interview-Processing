from __future__ import annotations

from dataclasses import replace
import html
import json
import time
from textwrap import dedent

import streamlit as st
import streamlit.components.v1 as components

from planlock.config import Settings
from planlock.job_runner import JobRunner
from planlock.models import ImportArtifacts, RunEvent, Severity, Stage
from planlock.pdf_renderer import render_pdf_previews


st.set_page_config(
    page_title="PlanLock",
    layout="wide",
)


COOLDOWN_HANDOFF_ANIMATION_MS = 420


LIVE_COUNTDOWN_BRIDGE_HTML = """
<script>
(() => {
  try {
    const parentWindow = window.parent;
    if (!parentWindow || !parentWindow.document) {
      return;
    }

    const intervalKey = "__planlockLiveCountdownInterval";
    if (parentWindow[intervalKey]) {
      parentWindow.clearInterval(parentWindow[intervalKey]);
    }

    const formatSeconds = (deltaMs) => `${(Math.max(0, deltaMs) / 1000).toFixed(1)}s`;

    const updateTimers = () => {
      const countdowns = parentWindow.document.querySelectorAll(".ocr-live-countdown[data-countdown-target-ms]");
      countdowns.forEach((node) => {
        const targetMs = Number(node.dataset.countdownTargetMs);
        if (!Number.isFinite(targetMs)) {
          return;
        }
        const deltaMs = targetMs - Date.now();
        if (deltaMs <= 0) {
          if (node.dataset.countdownState !== "finishing") {
            node.dataset.countdownState = "finishing";
            node.classList.add("is-finishing");
            node.textContent = "0.0s";
          }
          return;
        }
        if (node.dataset.countdownState === "finishing") {
          node.dataset.countdownState = "";
          node.classList.remove("is-finishing");
        }
        node.textContent = formatSeconds(deltaMs);
      });

      const elapsed = parentWindow.document.querySelectorAll(".ocr-live-elapsed[data-elapsed-started-at-ms]");
      elapsed.forEach((node) => {
        const startedAtMs = Number(node.dataset.elapsedStartedAtMs);
        if (!Number.isFinite(startedAtMs)) {
          return;
        }
        node.textContent = formatSeconds(Date.now() - startedAtMs);
      });
    };

    updateTimers();
    parentWindow[intervalKey] = parentWindow.setInterval(updateTimers, 100);

    const iframe = window.frameElement;
    if (iframe) {
      iframe.style.width = "0";
      iframe.style.height = "0";
      iframe.style.border = "0";
      iframe.style.position = "absolute";
    }
  } catch (error) {
    console.debug("PlanLock countdown bridge unavailable", error);
  }
})();
</script>
"""


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Public+Sans:wght@400;500;600;700&display=swap');

        :root {
            --bg: #eef2ef;
            --surface: rgba(255, 255, 255, 0.94);
            --surface-soft: #f6f8f5;
            --ink: #15211d;
            --muted: #61706a;
            --line: #d7e0da;
            --accent: #1b6f69;
            --accent-soft: #dcece7;
            --success: #2e7d57;
            --warn: #b97822;
            --error: #b24e46;
            --radius-xl: 26px;
            --radius-lg: 18px;
            --shadow: 0 18px 40px rgba(21, 33, 29, 0.06);
        }

        html, body, [class*="css"] {
            font-family: "Public Sans", sans-serif;
            color: var(--ink);
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(27, 111, 105, 0.12), transparent 26%),
                linear-gradient(180deg, var(--bg) 0%, #f7faf7 100%);
        }

        .block-container {
            max-width: 1120px;
            padding-top: 1.75rem;
            padding-bottom: 3rem;
        }

        [data-testid="stHeader"] {
            background: transparent;
        }

        [data-testid="stToolbar"] {
            top: 1rem;
            right: 1rem;
        }

        .shell {
            display: grid;
            gap: 1rem;
        }

        .masthead {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 1rem;
            padding-bottom: 0.55rem;
            border-bottom: 1px solid rgba(97, 112, 106, 0.18);
        }

        .eyebrow {
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            color: var(--accent);
            font-weight: 700;
        }

        .masthead-title {
            margin: 0.38rem 0 0;
            font-family: "Instrument Serif", serif;
            font-size: clamp(2.15rem, 4vw, 3.3rem);
            line-height: 0.92;
            letter-spacing: -0.05em;
        }

        .section-card,
        .status-card,
        .log-card,
        .stage-chip {
            border-radius: var(--radius-xl);
            border: 1px solid var(--line);
            background: var(--surface);
            box-shadow: var(--shadow);
        }

        .section-card,
        .log-card {
            padding: 1.2rem;
        }

        .status-card {
            padding: 1rem 1.1rem;
        }

        .status-card.warning {
            background: #fff8ef;
            border-color: #ecd5b4;
        }

        .status-card.error {
            background: #fff3f1;
            border-color: #e7c3bd;
        }

        .status-label {
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 700;
            color: var(--muted);
        }

        .status-message {
            margin-top: 0.3rem;
            font-size: 0.98rem;
            line-height: 1.5;
            color: var(--ink);
        }

        .section-title {
            margin: 0;
            font-size: 1.35rem;
            line-height: 1.05;
            letter-spacing: -0.04em;
            color: var(--ink);
        }

        .section-copy {
            margin-top: 0.5rem;
            color: var(--muted);
            font-size: 0.9rem;
            line-height: 1.5;
        }

        .stage-rail {
            display: flex;
            gap: 0.8rem;
            align-items: stretch;
        }

        .stage-chip {
            flex: 0 0 8.5rem;
            padding: 0.8rem 0.85rem;
            background: var(--surface-soft);
            border-radius: 18px;
        }

        .stage-chip.active {
            flex: 1 1 auto;
            padding: 0.95rem 1rem;
            background: #e6f1ee;
            border-color: #9fc3ba;
            border-radius: var(--radius-xl);
        }

        .stage-chip.complete {
            background: #eef7f1;
            border-color: #cfe1d5;
        }

        .stage-chip.compact {
            box-shadow: none;
        }

        .stage-chip-number {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            color: var(--muted);
            font-weight: 700;
        }

        .stage-chip-name {
            margin-top: 0.42rem;
            font-size: 1.02rem;
            line-height: 1.1;
            font-weight: 700;
            letter-spacing: -0.03em;
            color: var(--ink);
        }

        .stage-chip-state {
            margin-top: 0.55rem;
            font-size: 0.82rem;
            color: var(--muted);
        }

        .stage-chip.compact .stage-chip-name {
            margin-top: 0.28rem;
            font-size: 0.9rem;
            line-height: 1.05;
        }

        .stage-chip.compact .stage-chip-state {
            display: none;
        }

        .metric-row {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.75rem;
            margin-top: 1rem;
        }

        .metric {
            padding: 0.85rem 0.9rem;
            border-radius: var(--radius-lg);
            background: var(--surface-soft);
            border: 1px solid var(--line);
        }

        .metric span {
            display: block;
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            color: var(--muted);
            font-weight: 700;
        }

        .metric strong {
            display: block;
            margin-top: 0.42rem;
            font-size: 1rem;
            line-height: 1.3;
            letter-spacing: -0.03em;
            color: var(--ink);
            word-break: break-word;
        }

        .log-title {
            margin: 0;
            font-size: 1rem;
            letter-spacing: -0.03em;
        }

        .log-shell {
            height: 26rem;
            max-height: 26rem;
            margin-top: 0.8rem;
            border-radius: 20px;
            border: 1px solid var(--line);
            background: #f7faf8;
            padding: 1rem 1.05rem;
            color: #21302a;
            font-family: "SFMono-Regular", "Menlo", monospace;
            font-size: 0.83rem;
            line-height: 1.65;
            white-space: pre-wrap;
            overflow-y: auto;
            overflow-x: hidden;
            overscroll-behavior: contain;
        }

        .ocr-grid {
            display: grid;
            gap: 0.7rem;
            margin-top: 0.95rem;
        }

        .ocr-row {
            display: grid;
            grid-template-columns: minmax(0, 110px) minmax(0, 1fr) minmax(0, 120px);
            gap: 0.8rem;
            align-items: center;
            padding: 0.85rem 0.95rem;
            border-radius: var(--radius-lg);
            background: var(--surface-soft);
            border: 1px solid var(--line);
        }

        .ocr-row.just-completed {
            animation: laneCelebrate 900ms ease;
        }

        .ocr-row.inactive {
            background: #eff3f0;
            border-color: #dde5e0;
        }

        .ocr-row.inactive .ocr-pipe,
        .ocr-row.inactive .ocr-status,
        .ocr-row.inactive .ocr-timing,
        .ocr-row.inactive .ocr-preview-caption {
            color: #85938d;
        }

        .ocr-pipe {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-weight: 700;
            color: var(--ink);
        }

        .ocr-dot {
            width: 0.72rem;
            height: 0.72rem;
            border-radius: 999px;
            background: #bfcbc5;
            flex: 0 0 auto;
        }

        .ocr-dot.running,
        .ocr-dot.retrying {
            background: var(--accent);
        }

        .ocr-dot.complete {
            background: var(--success);
        }

        .ocr-dot.failed {
            background: var(--error);
        }

        .ocr-status {
            color: var(--muted);
            font-size: 0.88rem;
            line-height: 1.45;
        }

        .ocr-timing {
            text-align: right;
            font-size: 0.88rem;
            color: var(--muted);
        }

        .ocr-live-countdown,
        .ocr-live-elapsed,
        .ocr-live-handoff {
            display: inline-block;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }

        .ocr-live-countdown.is-finishing {
            animation: cooldownCountdownExit 220ms ease forwards;
        }

        .ocr-live-handoff {
            color: var(--accent);
            animation: cooldownHandoff 420ms cubic-bezier(0.22, 1, 0.36, 1) both;
        }

        .ocr-preview-shell {
            grid-column: 1 / -1;
            display: grid;
            grid-template-columns: 5.4rem minmax(0, 1fr);
            gap: 0.8rem;
            align-items: end;
            padding-top: 0.15rem;
        }

        .ocr-preview {
            width: 5.4rem;
            aspect-ratio: 0.72;
            border-radius: 14px;
            overflow: hidden;
            border: 1px solid var(--line);
            background:
                linear-gradient(180deg, rgba(27, 111, 105, 0.06) 0%, rgba(27, 111, 105, 0.02) 100%),
                #ffffff;
            box-shadow: 0 10px 24px rgba(21, 33, 29, 0.08);
        }

        .ocr-row.just-completed .ocr-preview {
            animation: previewPulse 900ms ease;
        }

        .ocr-row.inactive .ocr-preview {
            background: #f2f5f3;
            border-color: #dce4df;
            box-shadow: none;
        }

        .ocr-row.inactive .ocr-preview img {
            filter: grayscale(1) saturate(0.15) brightness(1.03);
        }

        .ocr-preview img {
            display: block;
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .ocr-preview.empty {
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 0.5rem;
            border-style: dashed;
            box-shadow: none;
            color: var(--muted);
            font-size: 0.7rem;
            line-height: 1.35;
            text-align: center;
        }

        .ocr-preview-caption {
            font-size: 0.77rem;
            color: var(--muted);
            line-height: 1.35;
        }

        .stFileUploader > div > div {
            border-radius: 18px !important;
            border: 1px dashed #9bbab2 !important;
            background: var(--surface-soft) !important;
            padding: 0.9rem !important;
        }

        .stFileUploader label,
        .stButton button,
        .stDownloadButton button,
        .stCaption,
        .stMarkdown,
        .stText,
        .stExpander {
            font-family: "Public Sans", sans-serif !important;
        }

        .stButton > button,
        .stDownloadButton > button {
            min-height: 3rem !important;
            border-radius: 999px !important;
            border: 1px solid var(--line) !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em !important;
            box-shadow: none !important;
            transition: background 180ms ease, border-color 180ms ease, color 180ms ease !important;
        }

        .stButton > button[kind="primary"] {
            background: var(--ink) !important;
            color: white !important;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover {
            border-color: #aac2bc !important;
        }

        .stDownloadButton > button {
            background: rgba(255, 255, 255, 0.7) !important;
            color: var(--ink) !important;
        }

        div[data-testid="stDataFrame"],
        div[data-testid="stJson"],
        [data-testid="stExpander"] {
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid var(--line);
            background: var(--surface);
        }

        @keyframes rise {
            from {
                opacity: 0;
                transform: translateY(10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }

        @keyframes laneCelebrate {
            0% {
                border-color: #9fc3ba;
                box-shadow: 0 0 0 0 rgba(27, 111, 105, 0.18);
            }
            45% {
                border-color: #70a79e;
                box-shadow: 0 0 0 8px rgba(27, 111, 105, 0.06);
            }
            100% {
                border-color: var(--line);
                box-shadow: none;
            }
        }

        @keyframes previewPulse {
            0% {
                transform: scale(0.96);
                border-color: #8bb6ad;
            }
            55% {
                transform: scale(1.02);
                border-color: #5f9f95;
            }
            100% {
                transform: scale(1);
                border-color: var(--line);
            }
        }

        @keyframes cooldownCountdownExit {
            0% {
                opacity: 1;
                transform: translateY(0) scale(1);
                filter: blur(0);
            }
            100% {
                opacity: 0;
                transform: translateY(-0.35rem) scale(0.92);
                filter: blur(2px);
            }
        }

        @keyframes cooldownHandoff {
            0% {
                opacity: 0;
                transform: translateY(0.45rem) scale(0.94);
                filter: blur(2px);
            }
            55% {
                opacity: 1;
                transform: translateY(-0.05rem) scale(1.02);
                filter: blur(0);
            }
            100% {
                opacity: 1;
                transform: translateY(0) scale(1);
                filter: blur(0);
            }
        }

        .section-card,
        .status-card,
        .log-card,
        .stage-chip {
            animation: rise 240ms ease both;
        }

        .live-card {
            animation: none !important;
        }

        @media (prefers-reduced-motion: reduce) {
            .section-card,
            .status-card,
            .log-card,
            .stage-chip,
            .ocr-live-countdown,
            .ocr-live-handoff,
            .stButton > button,
            .stDownloadButton > button {
                animation: none !important;
                transition: none !important;
            }
        }

        @media (max-width: 980px) {
            .masthead {
                flex-direction: column;
                align-items: flex-start;
            }

            .stage-rail,
            .metric-row {
                display: grid;
                grid-template-columns: 1fr;
            }

            .stage-chip,
            .stage-chip.active {
                flex: 1 1 auto;
            }

            .ocr-row {
                grid-template-columns: 1fr;
            }

            .ocr-timing {
                text-align: left;
            }

            .ocr-preview-shell {
                grid-template-columns: 4.8rem minmax(0, 1fr);
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_state(settings: Settings) -> None:
    st.session_state.setdefault("logs", [])
    st.session_state.setdefault("stage_progress", {stage.value: (0, 1) for stage in Stage})
    st.session_state.setdefault("active_stage", Stage.OCR.value)
    st.session_state.setdefault("result", None)
    st.session_state.setdefault("last_status", None)
    st.session_state.setdefault("ocr_pipeline", build_ocr_pipeline_state())
    st.session_state.setdefault("run_started", False)
    st.session_state.setdefault("is_running", False)
    st.session_state.setdefault("source_filename", None)
    st.session_state.setdefault("page_previews", {})
    runtime_defaults = build_runtime_settings_state(settings)
    runtime_settings = st.session_state.setdefault("runtime_settings", runtime_defaults)
    for key, value in runtime_defaults.items():
        runtime_settings.setdefault(key, value)


def build_ocr_pipeline_state(pipe_total: int = 3) -> dict:
    pipe_count = max(1, pipe_total)
    return {
        "pipe_total": pipe_count,
        "page_total": 0,
        "completed_pages": 0,
        "pipes": [build_ocr_pipe_state(pipe_number) for pipe_number in range(1, pipe_count + 1)],
    }


def build_ocr_pipe_state(pipe_number: int) -> dict:
    return {
        "pipe_number": pipe_number,
        "status": "idle",
        "page_number": None,
        "started_at_ms": None,
        "last_completed_page": None,
        "completed_at_ms": None,
        "attempt_number": None,
        "max_attempts": None,
        "retry_delay_seconds": None,
        "retry_until_ms": None,
        "cooldown_finished_at_ms": None,
        "retry_reason": None,
        "last_error": None,
        "flash_complete": False,
    }


def build_runtime_settings_state(settings: Settings) -> dict:
    return {
        "model_ocr": settings.model_ocr,
        "model_mapping": settings.model_mapping,
        "ocr_parallel_workers": settings.ocr_parallel_workers,
        "llm_timeout_seconds": settings.llm_timeout_seconds,
        "llm_max_retries": settings.llm_max_retries,
        "llm_retry_base_seconds": settings.llm_retry_base_seconds,
        "llm_retry_max_seconds": settings.llm_retry_max_seconds,
        "max_pages": settings.max_pages,
        "log_level": settings.log_level,
    }


def apply_runtime_settings(base_settings: Settings) -> Settings:
    runtime_settings = st.session_state["runtime_settings"]
    return replace(
        base_settings,
        model_ocr=str(runtime_settings["model_ocr"]).strip() or base_settings.model_ocr,
        model_mapping=str(runtime_settings["model_mapping"]).strip() or base_settings.model_mapping,
        ocr_parallel_workers=int(runtime_settings["ocr_parallel_workers"]),
        llm_timeout_seconds=float(runtime_settings["llm_timeout_seconds"]),
        llm_max_retries=int(runtime_settings["llm_max_retries"]),
        llm_retry_base_seconds=float(runtime_settings["llm_retry_base_seconds"]),
        llm_retry_max_seconds=float(runtime_settings["llm_retry_max_seconds"]),
        max_pages=int(runtime_settings["max_pages"]),
        log_level=str(runtime_settings["log_level"]),
    )


def ensure_ocr_pipe_slots(pipeline: dict, pipe_total: int) -> None:
    desired = max(1, pipe_total)
    pipes = pipeline.setdefault("pipes", [])
    while len(pipes) < desired:
        pipes.append(build_ocr_pipe_state(len(pipes) + 1))
    if len(pipes) > desired:
        del pipes[desired:]
    pipeline["pipe_total"] = desired


def reset_run_state(pipe_total: int, source_filename: str, page_previews: dict[int, str]) -> None:
    st.session_state["logs"] = []
    st.session_state["stage_progress"] = {stage.value: (0, 1) for stage in Stage}
    st.session_state["active_stage"] = Stage.OCR.value
    st.session_state["result"] = None
    st.session_state["last_status"] = {
        "index": 1,
        "stage": display_stage_name(Stage.OCR),
        "message": "Job queued. Preparing document review.",
        "severity": Severity.INFO,
    }
    st.session_state["ocr_pipeline"] = build_ocr_pipeline_state(pipe_total)
    st.session_state["run_started"] = True
    st.session_state["is_running"] = True
    st.session_state["source_filename"] = source_filename
    st.session_state["page_previews"] = page_previews


def stage_index(stage: Stage) -> int:
    return list(Stage).index(stage) + 1


def display_stage_name(stage: Stage | str, *, compact: bool = False) -> str:
    stage_enum = stage if isinstance(stage, Stage) else Stage(stage)
    names = {
        Stage.OCR: ("Document review", "Review"),
        Stage.DATA_ENTRY: ("Workbook entry", "Entry"),
        Stage.FINANCIAL_CALCULATIONS: ("Workbook checks", "Checks"),
    }
    full_name, compact_name = names[stage_enum]
    return compact_name if compact else full_name


def status_tone(severity: Severity) -> str:
    if severity == Severity.ERROR:
        return "error"
    if severity == Severity.WARNING:
        return "warning"
    return "info"


def format_seconds_label(seconds: float) -> str:
    return f"{seconds:.1f}s"


def build_timing_markup(
    label: str,
    *,
    countdown_target_ms: int | None = None,
    elapsed_started_at_ms: int | None = None,
    extra_classes: tuple[str, ...] = (),
    animation_offset_ms: int | None = None,
) -> str:
    safe_label = html.escape(label)
    class_names = list(extra_classes)
    attrs: list[str] = []
    style_parts: list[str] = []

    if countdown_target_ms is not None:
        class_names.insert(0, "ocr-live-countdown")
        attrs.append(f'data-countdown-target-ms="{int(countdown_target_ms)}"')
    if elapsed_started_at_ms is not None:
        class_names.insert(0, "ocr-live-elapsed")
        attrs.append(f'data-elapsed-started-at-ms="{int(elapsed_started_at_ms)}"')
    if animation_offset_ms is not None and animation_offset_ms > 0:
        style_parts.append(f"animation-delay: -{int(animation_offset_ms)}ms;")

    if not class_names and not attrs and not style_parts:
        return safe_label

    class_attr = f' class="{" ".join(class_names)}"' if class_names else ""
    attr_text = f" {' '.join(attrs)}" if attrs else ""
    style_attr = f' style="{" ".join(style_parts)}"' if style_parts else ""
    return f"<span{class_attr}{attr_text}{style_attr}>{safe_label}</span>"


def mount_live_countdown_bridge() -> None:
    components.html(LIVE_COUNTDOWN_BRIDGE_HTML, height=0, width=0)


@st.dialog("Run settings")
def render_settings_dialog(base_settings: Settings) -> None:
    defaults = build_runtime_settings_state(base_settings)
    current = dict(st.session_state["runtime_settings"])

    with st.form("runtime_settings_form", border=False):
        st.caption("Only the API key lives in `.env`. These settings apply to the current browser session.")
        model_ocr = st.text_input("Document review model", value=str(current["model_ocr"]))
        model_mapping = st.text_input("Workbook entry model", value=str(current["model_mapping"]))

        col1, col2 = st.columns(2, gap="small")
        ocr_parallel_workers = col1.number_input(
            "Parallel lanes",
            min_value=1,
            max_value=10,
            value=int(current["ocr_parallel_workers"]),
            step=1,
        )
        max_pages = col2.number_input(
            "Page limit",
            min_value=1,
            max_value=200,
            value=int(current["max_pages"]),
            step=1,
        )

        col3, col4 = st.columns(2, gap="small")
        llm_timeout_seconds = col3.number_input(
            "Timeout per call (seconds)",
            min_value=10.0,
            max_value=600.0,
            value=float(current["llm_timeout_seconds"]),
            step=10.0,
        )
        llm_max_retries = col4.number_input(
            "Retry count",
            min_value=0,
            max_value=10,
            value=int(current["llm_max_retries"]),
            step=1,
        )

        col5, col6 = st.columns(2, gap="small")
        llm_retry_base_seconds = col5.number_input(
            "First wait between tries (seconds)",
            min_value=0.5,
            max_value=60.0,
            value=float(current["llm_retry_base_seconds"]),
            step=0.5,
        )
        llm_retry_max_seconds = col6.number_input(
            "Longest wait between tries (seconds)",
            min_value=0.5,
            max_value=120.0,
            value=float(current["llm_retry_max_seconds"]),
            step=0.5,
        )

        log_level = st.selectbox(
            "Log detail",
            options=["DEBUG", "INFO", "WARNING", "ERROR"],
            index=["DEBUG", "INFO", "WARNING", "ERROR"].index(str(current["log_level"])),
        )

        save_col, reset_col = st.columns(2, gap="small")
        save_clicked = save_col.form_submit_button("Save settings", use_container_width=True)
        reset_clicked = reset_col.form_submit_button("Use defaults", use_container_width=True)

    if reset_clicked:
        st.session_state["runtime_settings"] = defaults
        st.rerun()

    if save_clicked:
        st.session_state["runtime_settings"] = {
            "model_ocr": model_ocr.strip() or defaults["model_ocr"],
            "model_mapping": model_mapping.strip() or defaults["model_mapping"],
            "ocr_parallel_workers": int(ocr_parallel_workers),
            "llm_timeout_seconds": float(llm_timeout_seconds),
            "llm_max_retries": int(llm_max_retries),
            "llm_retry_base_seconds": float(llm_retry_base_seconds),
            "llm_retry_max_seconds": float(llm_retry_max_seconds),
            "max_pages": int(max_pages),
            "log_level": log_level,
        }
        st.rerun()


def render_masthead() -> bool:
    title_col, action_col = st.columns([0.82, 0.18], gap="small")
    with title_col:
        st.markdown(
            f"""
            <div class="masthead">
                <div>
                    <div class="eyebrow">PlanLock</div>
                    <h1 class="masthead-title">Turn planner PDFs into locked workbooks</h1>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with action_col:
        st.markdown("<div style='height:0.25rem'></div>", unsafe_allow_html=True)
        return st.button("Settings", key="open_run_settings", use_container_width=True)


def render_status(status_placeholder) -> None:
    status = st.session_state.get("last_status")
    if status is None:
        status_placeholder.markdown(
            """
            <div class="status-card info">
                <div class="status-label">System ready</div>
                <div class="status-message">Upload a planner PDF to start the three-step intake.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    tone = status_tone(status["severity"])
    status_placeholder.markdown(
        f"""
        <div class="status-card {tone}">
            <div class="status-label">Stage {status["index"]} of 3 • {html.escape(status["stage"])}</div>
            <div class="status-message">{html.escape(status["message"])}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_stage_progress() -> None:
    progress = st.session_state["stage_progress"]
    active_stage = st.session_state["active_stage"]
    is_running = bool(st.session_state.get("is_running"))
    run_started = bool(st.session_state.get("run_started"))
    last_status = st.session_state.get("last_status")

    cards: list[str] = []
    for index, stage in enumerate(Stage, start=1):
        completed, total = progress.get(stage.value, (0, 1))
        is_complete = completed >= total and total > 0
        is_active = stage.value == active_stage
        classes = ["stage-chip"]
        state_label = "Waiting"

        if is_active:
            classes.append("active")
        else:
            classes.append("compact")

        if is_complete:
            classes.append("complete")
            state_label = "Complete"
        elif is_active:
            if is_running:
                state_label = "In progress"
            elif run_started and last_status is not None and last_status["severity"] == Severity.ERROR:
                state_label = "Stopped"
            else:
                state_label = "Ready"
        elif not run_started and index == 1:
            state_label = "Ready"

        cards.append(
            dedent(
                f"""
                <div class="{' '.join(classes)}">
                    <div class="stage-chip-number">Step {index}</div>
                    <div class="stage-chip-name">{html.escape(display_stage_name(stage, compact=not is_active))}</div>
                    <div class="stage-chip-state">{state_label}</div>
                </div>
                """
            ).strip()
        )

    st.markdown(f'<div class="stage-rail">{"".join(cards)}</div>', unsafe_allow_html=True)


def append_event(event: RunEvent) -> None:
    st.session_state["active_stage"] = event.stage.value
    st.session_state["stage_progress"][event.stage.value] = (event.stage_completed, event.stage_total)
    if event.phase != "heartbeat":
        st.session_state["last_status"] = {
            "index": stage_index(event.stage),
            "stage": display_stage_name(event.stage),
            "message": event.message,
            "severity": event.severity,
        }
    if event.stage == Stage.OCR:
        pipeline = st.session_state["ocr_pipeline"]
        if event.pipe_total is not None:
            ensure_ocr_pipe_slots(pipeline, event.pipe_total)
        if event.page_total is not None:
            pipeline["page_total"] = event.page_total
        pipeline["completed_pages"] = event.stage_completed
        now_ms = int(time.time() * 1000)

        if event.pipe_number is not None and 1 <= event.pipe_number <= len(pipeline["pipes"]):
            pipe = pipeline["pipes"][event.pipe_number - 1]
            if event.phase == "start":
                pipe["status"] = "running"
                pipe["page_number"] = event.page_number
                pipe["started_at_ms"] = now_ms
                pipe["completed_at_ms"] = None
                pipe["attempt_number"] = 1
                pipe["max_attempts"] = event.max_attempts
                pipe["retry_delay_seconds"] = None
                pipe["retry_until_ms"] = None
                pipe["cooldown_finished_at_ms"] = None
                pipe["retry_reason"] = None
                pipe["last_error"] = None
                pipe["flash_complete"] = False
            elif event.phase == "retry":
                pipe["status"] = "retrying"
                pipe["page_number"] = event.page_number
                if pipe.get("started_at_ms") is None:
                    pipe["started_at_ms"] = now_ms
                pipe["completed_at_ms"] = None
                pipe["attempt_number"] = event.attempt_number
                pipe["max_attempts"] = event.max_attempts
                pipe["retry_delay_seconds"] = event.retry_delay_seconds
                pipe["retry_until_ms"] = (
                    now_ms + int(event.retry_delay_seconds * 1000)
                    if event.retry_delay_seconds is not None
                    else None
                )
                pipe["cooldown_finished_at_ms"] = None
                pipe["retry_reason"] = event.retry_reason
                pipe["last_error"] = event.message
                pipe["flash_complete"] = False
            elif event.phase == "complete":
                pipe["status"] = "complete"
                pipe["page_number"] = None
                pipe["started_at_ms"] = None
                pipe["last_completed_page"] = event.page_number
                pipe["completed_at_ms"] = now_ms
                pipe["attempt_number"] = None
                pipe["max_attempts"] = None
                pipe["retry_delay_seconds"] = None
                pipe["retry_until_ms"] = None
                pipe["cooldown_finished_at_ms"] = None
                pipe["retry_reason"] = None
                pipe["last_error"] = None
                pipe["flash_complete"] = True
            elif event.phase == "failed":
                pipe["status"] = "failed"
                pipe["page_number"] = event.page_number
                pipe["started_at_ms"] = None
                pipe["completed_at_ms"] = None
                pipe["attempt_number"] = event.attempt_number
                pipe["max_attempts"] = event.max_attempts
                pipe["retry_delay_seconds"] = None
                pipe["retry_until_ms"] = None
                pipe["cooldown_finished_at_ms"] = None
                pipe["retry_reason"] = None
                pipe["last_error"] = event.message
                pipe["flash_complete"] = False

        for pipe in pipeline["pipes"]:
            retry_until_ms = pipe.get("retry_until_ms")
            if (
                pipe.get("status") == "retrying"
                and retry_until_ms is not None
                and int(retry_until_ms) <= now_ms
                and pipe.get("cooldown_finished_at_ms") is None
            ):
                pipe["cooldown_finished_at_ms"] = int(retry_until_ms)
    if event.phase != "heartbeat":
        prefix = {
            Severity.INFO: "[INFO]",
            Severity.WARNING: "[WARN]",
            Severity.ERROR: "[ERROR]",
        }[event.severity]
        st.session_state["logs"].append(f"{prefix} {display_stage_name(event.stage)}: {event.message}")
    if event.artifacts is not None:
        st.session_state["result"] = event.artifacts


def render_logs(log_placeholder) -> None:
    log_text = chr(10).join(st.session_state["logs"][-200:]) or "Logs will appear here once the pipeline starts."
    log_placeholder.markdown(
        f"""
        <div class="log-card">
            <h3 class="log-title">Activity log</h3>
            <div class="log-shell">{html.escape(log_text)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_ocr_parallel() -> None:
    pipeline = st.session_state["ocr_pipeline"]
    page_previews = st.session_state.get("page_previews", {})
    active_statuses = {"running", "retrying"}
    pipes = sorted(
        pipeline.get("pipes", []),
        key=lambda pipe: (0 if str(pipe.get("status", "idle")) in active_statuses else 1, int(pipe["pipe_number"])),
    )
    page_total = int(pipeline.get("page_total", 0) or 0)
    completed_pages = int(pipeline.get("completed_pages", 0) or 0)
    pipe_total = int(pipeline.get("pipe_total", len(pipes)) or len(pipes) or 1)
    active_pipes = sum(1 for pipe in pipes if pipe.get("status") in active_statuses)
    remaining_pages = max(0, page_total - completed_pages)
    now_ms = int(time.time() * 1000)

    rows: list[str] = []
    for pipe in pipes:
        status = str(pipe.get("status", "idle"))
        page_number = pipe.get("page_number")
        last_page = pipe.get("last_completed_page")
        started_at_ms = pipe.get("started_at_ms")
        completed_at_ms = pipe.get("completed_at_ms")
        attempt_number = pipe.get("attempt_number")
        max_attempts = pipe.get("max_attempts")
        retry_delay_seconds = pipe.get("retry_delay_seconds")
        retry_until_ms = pipe.get("retry_until_ms")
        cooldown_finished_at_ms = pipe.get("cooldown_finished_at_ms")
        retry_reason = pipe.get("retry_reason")
        preview_page_number = page_number if page_number is not None else last_page
        preview_src = page_previews.get(int(preview_page_number)) if preview_page_number is not None else None
        just_completed = bool(pipe.get("flash_complete"))
        recently_completed = bool(
            status == "complete"
            and completed_at_ms is not None
            and (now_ms - int(completed_at_ms)) < 1400
        )
        is_active = status in active_statuses

        status_label = "Idle"
        detail_label = "Waiting for pages"
        timing_markup = build_timing_markup("Open")

        if status == "running" and page_number is not None:
            status_label = f"Reviewing page {page_number}"
            if started_at_ms is not None:
                elapsed_seconds = round((now_ms - started_at_ms) / 1000, 1)
                timing_markup = build_timing_markup(
                    format_seconds_label(elapsed_seconds),
                    elapsed_started_at_ms=int(started_at_ms),
                )
        elif status == "retrying" and page_number is not None:
            status_label = (
                f"Cooling down page {page_number}"
                if retry_reason == "rate_limit"
                else f"Retrying page {page_number}"
            )
            if attempt_number is not None and max_attempts is not None:
                detail_label = f"Pass {attempt_number}/{max_attempts}"
            if retry_reason == "rate_limit":
                detail_label = "Waiting for the Anthropic token window to clear"
            if retry_until_ms is not None:
                remaining_ms = int(retry_until_ms) - now_ms
                if remaining_ms > 0:
                    timing_markup = build_timing_markup(
                        format_seconds_label(remaining_ms / 1000),
                        countdown_target_ms=int(retry_until_ms),
                    )
                else:
                    finished_at_ms = int(cooldown_finished_at_ms or retry_until_ms)
                    handoff_elapsed_ms = max(0, now_ms - finished_at_ms)
                    if handoff_elapsed_ms < COOLDOWN_HANDOFF_ANIMATION_MS:
                        timing_markup = build_timing_markup(
                            "Retrying",
                            extra_classes=("ocr-live-handoff",),
                            animation_offset_ms=handoff_elapsed_ms,
                        )
                    else:
                        timing_markup = build_timing_markup("Retrying")
            elif retry_delay_seconds is not None:
                timing_markup = build_timing_markup(format_seconds_label(retry_delay_seconds))
        elif status == "complete" and last_page is not None:
            status_label = f"Finished page {last_page}"
            detail_label = "Lane ready for the next page"
            timing_markup = build_timing_markup("0.0s" if recently_completed else "Ready")
        elif status == "failed" and page_number is not None:
            status_label = f"Page {page_number} needs attention"
            detail_label = "See log for the error"
            timing_markup = build_timing_markup("Failed")

        row_classes = ["ocr-row"]
        if not is_active:
            row_classes.append("inactive")
        if just_completed:
            row_classes.append("just-completed")

        preview_markup = '<div class="ocr-preview empty">Waiting</div><div class="ocr-preview-caption">No page assigned</div>'
        if preview_src is not None and preview_page_number is not None:
            preview_state = "current" if page_number is not None else "last"
            preview_caption = (
                f"Page {preview_page_number} in progress"
                if page_number is not None
                else f"Last page {preview_page_number}"
            )
            preview_markup = (
                f'<div class="ocr-preview {preview_state}">'
                f'<img src="{html.escape(preview_src, quote=True)}" alt="Preview of page {int(preview_page_number)}" />'
                "</div>"
                f'<div class="ocr-preview-caption">{html.escape(preview_caption)}</div>'
            )

        rows.append(
            dedent(
                f"""
                <div class="{' '.join(row_classes)}">
                    <div class="ocr-pipe">
                        <span class="ocr-dot {html.escape(status)}"></span>
                        Lane {int(pipe["pipe_number"])}
                    </div>
                    <div class="ocr-status">
                        <strong>{html.escape(status_label)}</strong><br>{html.escape(detail_label)}
                    </div>
                    <div class="ocr-timing">{timing_markup}</div>
                    <div class="ocr-preview-shell">{preview_markup}</div>
                </div>
                """
            ).strip()
        )
        if just_completed:
            pipe["flash_complete"] = False

    rows_markup = "".join(rows)

    st.markdown(
        dedent(
            f"""
            <div class="section-card live-card">
                <h3 class="section-title">Parallel document review</h3>
                <div class="metric-row">
                    <div class="metric">
                        <span>Pages complete</span>
                        <strong>{completed_pages}/{page_total or 0}</strong>
                    </div>
                    <div class="metric">
                        <span>Pages remaining</span>
                        <strong>{remaining_pages}</strong>
                    </div>
                    <div class="metric">
                        <span>Active lanes</span>
                        <strong>{active_pipes}/{pipe_total}</strong>
                    </div>
                </div>
                <div class="ocr-grid">{rows_markup}</div>
            </div>
            """
        ).strip(),
        unsafe_allow_html=True,
    )


def render_upload_panel(
    *,
    context_key: str,
    title: str,
    copy: str | None = None,
) -> tuple[object | None, bool]:
    copy_markup = f'<div class="section-copy">{html.escape(copy)}</div>' if copy else ""
    st.markdown(
        f"""
        <div class="section-card">
            <h3 class="section-title">{html.escape(title)}</h3>
            {copy_markup}
        """,
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader(
        "Planner PDF",
        type=["pdf"],
        key=f"pdf_upload_{context_key}",
    )
    run_clicked = st.button(
        "Build workbook",
        key=f"run_import_{context_key}",
        type="primary",
        use_container_width=True,
        disabled=uploaded_file is None,
    )
    st.markdown("</div>", unsafe_allow_html=True)
    return uploaded_file, run_clicked


def render_stage_focus() -> None:
    stage = Stage(st.session_state["active_stage"])
    source_filename = st.session_state.get("source_filename") or "Current upload"

    if stage == Stage.OCR:
        render_ocr_parallel()
        return

    label = "Writing workbook" if stage == Stage.DATA_ENTRY else "Validating output"
    st.markdown(
        f"""
        <div class="section-card">
            <h3 class="section-title">{html.escape(label)}</h3>
            <div class="section-copy">{html.escape(source_filename)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_result(result: ImportArtifacts | None) -> None:
    if result is None or result.review_report is None:
        return

    report = result.review_report
    headline = "Workbook ready" if result.success else "Review required"
    summary_copy = f"{len(report.mapped_assignments)} mapped, {len(report.unmapped_items)} unresolved."

    st.markdown(
        f"""
        <div class="section-card">
            <h3 class="section-title">{html.escape(headline)}</h3>
            <div class="section-copy">{html.escape(summary_copy)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div style='height:0.85rem'></div>", unsafe_allow_html=True)
    download_col, report_col = st.columns(2, gap="small")
    with download_col:
        if result.output_workbook_path and result.output_workbook_path.exists():
            st.download_button(
                "Download filled workbook",
                data=result.output_workbook_path.read_bytes(),
                file_name=result.output_workbook_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with report_col:
        st.download_button(
            "Download review report",
            data=json.dumps(report.model_dump(mode="json"), indent=2),
            file_name=result.review_report_path.name if result.review_report_path else "review_report.json",
            mime="application/json",
            use_container_width=True,
        )

    if report.warnings:
        with st.expander(f"Warnings ({len(report.warnings)})", expanded=False):
            for warning in report.warnings:
                st.markdown(f"- `{display_stage_name(warning.stage)}`: {warning.message}")

    with st.expander(f"Mapped assignments ({len(report.mapped_assignments)})", expanded=False):
        st.dataframe(
            [
                {
                    "sheet": assignment.sheet_name,
                    "cell": assignment.cell,
                    "semantic_key": assignment.semantic_key,
                    "value": assignment.value,
                    "comment": assignment.comment,
                }
                for assignment in report.mapped_assignments
            ],
            use_container_width=True,
            hide_index=True,
        )

    if report.unmapped_items:
        with st.expander(f"Unmapped items ({len(report.unmapped_items)})", expanded=False):
            st.json(report.unmapped_items)

    if report.assumptions:
        with st.expander(f"Assumptions ({len(report.assumptions)})", expanded=False):
            st.json(report.assumptions)


def render_work_area(work_placeholder, settings: Settings) -> tuple[object | None, bool]:
    uploaded_file = None
    run_clicked = False
    result = st.session_state.get("result")
    is_running = bool(st.session_state.get("is_running"))
    run_started = bool(st.session_state.get("run_started"))
    status = st.session_state.get("last_status")

    with work_placeholder.container():
        if is_running:
            render_stage_focus()
            return None, False

        if result is not None and result.review_report is not None:
            render_result(result)
            st.markdown("<div style='height:0.85rem'></div>", unsafe_allow_html=True)
            with st.expander("Import another PDF", expanded=False):
                uploaded_file, run_clicked = render_upload_panel(
                    context_key="next",
                    title="Process another PDF",
                    copy=f"Current settings: {settings.ocr_parallel_workers} lanes, {settings.max_pages}-page limit.",
                )
            return uploaded_file, run_clicked

        if run_started and status is not None and status["severity"] == Severity.ERROR:
            uploaded_file, run_clicked = render_upload_panel(
                context_key="retry",
                title="Run failed",
                copy=status["message"],
            )
            return uploaded_file, run_clicked

        uploaded_file, run_clicked = render_upload_panel(
            context_key="primary",
            title="Upload PDF",
            copy=(
                "Use Settings to change models, page limit, parallel lanes, or retry timing. "
                "Only the API key stays in .env."
            ),
        )
    return uploaded_file, run_clicked


def main() -> None:
    inject_styles()
    mount_live_countdown_bridge()
    base_settings = Settings.from_env()
    init_state(base_settings)
    settings = apply_runtime_settings(base_settings)

    try:
        base_settings.validate_template_lock()
    except Exception as exc:
        st.error(str(exc))
        return

    st.markdown('<div class="shell">', unsafe_allow_html=True)
    settings_clicked = render_masthead()
    if settings_clicked:
        render_settings_dialog(base_settings)

    status_placeholder = st.empty()
    render_status(status_placeholder)

    st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)
    progress_placeholder = st.empty()
    with progress_placeholder.container():
        render_stage_progress()

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    work_col, log_col = st.columns([1.05, 0.95], gap="large")

    with work_col:
        work_placeholder = st.empty()
        uploaded_file, run_clicked = render_work_area(work_placeholder, settings)

    with log_col:
        log_placeholder = st.empty()
        render_logs(log_placeholder)

    if run_clicked and uploaded_file is not None:
        uploaded_bytes = uploaded_file.getvalue()
        try:
            page_previews = render_pdf_previews(uploaded_bytes, settings.max_pages)
        except Exception:
            page_previews = {}
        reset_run_state(settings.ocr_parallel_workers, uploaded_file.name, page_previews)
        render_status(status_placeholder)
        progress_placeholder.empty()
        with progress_placeholder.container():
            render_stage_progress()
        render_logs(log_placeholder)
        render_work_area(work_placeholder, settings)

        runner = JobRunner(settings)
        try:
            for event in runner.run(uploaded_bytes, uploaded_file.name):
                append_event(event)
                if event.phase == "heartbeat":
                    continue
                render_status(status_placeholder)
                progress_placeholder.empty()
                with progress_placeholder.container():
                    render_stage_progress()
                render_logs(log_placeholder)
                render_work_area(work_placeholder, settings)
        except Exception as exc:
            st.session_state["logs"].append(f"[ERROR] Pipeline: {exc}")
            st.session_state["last_status"] = {
                "index": stage_index(Stage(st.session_state["active_stage"])),
                "stage": display_stage_name(st.session_state["active_stage"]),
                "message": f"Run failed: {exc}",
                "severity": Severity.ERROR,
            }
        finally:
            st.session_state["is_running"] = False
            render_status(status_placeholder)
            progress_placeholder.empty()
            with progress_placeholder.container():
                render_stage_progress()
            render_logs(log_placeholder)
            render_work_area(work_placeholder, settings)

    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
