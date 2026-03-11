from __future__ import annotations

from dataclasses import replace
import html
import json
import time
from textwrap import dedent

import streamlit as st
import streamlit.components.v1 as components

from planlock import APP_NAME
from planlock.config import (
    DEFAULT_LLM_PROVIDER,
    DEFAULT_OCR_PARALLEL_WORKERS,
    LLM_PROVIDER_OPTIONS,
    Settings,
    locked_model_for_provider,
    provider_display_name,
)
from planlock.job_runner import JobRunner
from planlock.models import AgentQuestion, EntrySessionState, ImportArtifacts, RunEvent, Severity, Stage
from planlock.pdf_renderer import render_pdf_previews
from planlock.template_entry_agent import load_entry_state
from planlock.template_schema import ALLOWED_WRITE_CELLS_BY_SHEET, TEMPLATE_SHEET_ORDER


st.set_page_config(
    page_title=APP_NAME,
    layout="wide",
)


COOLDOWN_HANDOFF_ANIMATION_MS = 420
QUESTION_DIALOG_TITLE = "Decision Gate"
QUESTION_SUBMISSION_STATE_KEY = "entry_question_submission"
QUESTION_PENDING_RESUME_STATE_KEY = "entry_question_pending_resume"
QUESTION_ACTIVE_RESUME_STATE_KEY = "entry_question_active_resume"
QUESTION_WIDGET_STATE_KEY = "entry_question_widgets"
PROVIDER_LOGO_URLS = {
    "openai": "https://us1.discourse-cdn.com/openai1/original/4X/3/2/1/321a1ba297482d3d4060d114860de1aa5610f8a9.png",
    "anthropic": "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRlh7pFp_23fKHEyUCA6K6V44mSYZNcboaY9A&s",
}


LIVE_COUNTDOWN_BRIDGE_HTML = """
<script>
(() => {
  try {
    const parentWindow = window.parent;
    if (!parentWindow || !parentWindow.document) {
      return;
    }

    const intervalKey = "__hollyplannerLiveCountdownInterval";
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
    console.debug("HollyPlanner countdown bridge unavailable", error);
  }
})();
</script>
"""


def render_html_block(markup: str) -> None:
    # Streamlit markdown can coerce nested preview HTML into code blocks; use raw HTML when available.
    html_renderer = getattr(st, "html", None)
    if callable(html_renderer):
        html_renderer(markup)
        return
    st.markdown(markup, unsafe_allow_html=True)


def rerun_app() -> None:
    rerun = getattr(st, "rerun")
    try:
        rerun(scope="app")
    except TypeError:
        rerun()


def inject_styles() -> None:
    render_html_block(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

        :root {
            --display-font: "Fraunces", "Iowan Old Style", "Palatino Linotype", serif;
            --body-font: "IBM Plex Sans", sans-serif;
            --bg: #f3eee4;
            --surface: rgba(255, 251, 244, 0.97);
            --surface-soft: #efe5d6;
            --ink: #0f1720;
            --muted: #39495b;
            --line: #d8ccb9;
            --line-strong: #bfa983;
            --accent: #17324d;
            --accent-strong: #0d2135;
            --accent-soft: #d8e1ea;
            --signal: #c99532;
            --signal-strong: #9c6c19;
            --signal-soft: #f5e6c4;
            --success: #1d7348;
            --warn: #a66420;
            --warn-soft: #fff3e3;
            --warn-line: #e3c18b;
            --error: #a1413d;
            --error-soft: #fff0ec;
            --error-line: #deaea7;
            --progress-pending-start: #e8decd;
            --progress-pending-end: #d8ccb8;
            --progress-complete-start: #2e7b52;
            --progress-complete-end: #153524;
            --progress-active-start: #f5dfa4;
            --progress-active-mid: #c99532;
            --progress-active-end: #17324d;
            --progress-active-glow: rgba(201, 149, 50, 0.26);
            --radius-xl: 28px;
            --radius-lg: 20px;
            --shadow: 0 28px 60px rgba(15, 23, 32, 0.14);
        }

        html, body, [class*="css"] {
            font-family: var(--body-font);
            color: var(--ink);
        }

        .stApp,
        .stApp * {
            font-family: var(--body-font) !important;
        }

        h1, h2, h3, h4, h5, h6 {
            font-weight: 600;
            letter-spacing: -0.03em;
            color: var(--ink);
        }

        .stApp {
            background:
                radial-gradient(circle at 16% 10%, rgba(23, 50, 77, 0.1), transparent 24%),
                radial-gradient(circle at 88% 8%, rgba(201, 149, 50, 0.14), transparent 16%),
                linear-gradient(180deg, #fbf7f0 0%, var(--bg) 52%, #efe5d6 100%);
        }

        .block-container {
            max-width: 100%;
            padding-top: 1.1rem;
            padding-bottom: 3rem;
            padding-left: clamp(1rem, 2.8vw, 2.5rem);
            padding-right: clamp(1rem, 2.8vw, 2.5rem);
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
            gap: 0.85rem;
        }

        .taskbar-shell {
            position: sticky;
            top: 0.85rem;
            z-index: 30;
            margin-bottom: 0.2rem;
        }

        .taskbar-bar {
            position: relative;
            display: grid;
            grid-template-columns: minmax(14rem, 1fr) minmax(18rem, 1.35fr) auto;
            gap: 0.75rem;
            align-items: center;
            padding: 0.7rem 0.82rem 0.7rem 0.95rem;
            overflow: hidden;
            isolation: isolate;
            border-radius: 24px;
            border: 1px solid rgba(23, 50, 77, 0.14);
            background:
                linear-gradient(135deg, rgba(255, 251, 244, 0.92) 0%, rgba(248, 241, 231, 0.88) 48%, rgba(243, 233, 219, 0.92) 100%);
            backdrop-filter: blur(18px) saturate(1.05);
            box-shadow:
                0 18px 38px rgba(15, 23, 32, 0.12),
                inset 0 1px 0 rgba(255, 255, 255, 0.66);
        }

        .taskbar-bar::before {
            content: "";
            position: absolute;
            inset: 0 auto 0 0;
            width: 0.36rem;
            background: linear-gradient(180deg, var(--signal) 0%, var(--accent) 100%);
        }

        .taskbar-bar::after {
            content: "";
            position: absolute;
            inset: auto 10% -3rem auto;
            width: 10rem;
            height: 10rem;
            background: radial-gradient(circle, rgba(201, 149, 50, 0.18) 0%, rgba(201, 149, 50, 0) 68%);
            pointer-events: none;
        }

        .taskbar-brand-lockup {
            min-width: 0;
            display: grid;
            gap: 0.16rem;
            padding-left: 0.2rem;
        }

        .taskbar-brand-row {
            display: flex;
            align-items: baseline;
            gap: 0.55rem;
            min-width: 0;
        }

        .taskbar-brand {
            flex: 0 0 auto;
            font-family: var(--display-font);
            font-size: clamp(1.1rem, 2vw, 1.45rem);
            line-height: 0.95;
            letter-spacing: -0.05em;
            font-weight: 700;
            color: var(--ink);
            white-space: nowrap;
        }

        .taskbar-strap {
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-size: 0.68rem;
            line-height: 1.2;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            font-weight: 800;
            color: var(--signal-strong);
        }

        .taskbar-context {
            font-size: 0.72rem;
            line-height: 1.2;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            font-weight: 700;
            color: rgba(23, 50, 77, 0.76);
        }

        .taskbar-status {
            position: relative;
            min-width: 0;
            display: flex;
            align-items: center;
            gap: 0.55rem;
            padding: 0.48rem 0.58rem;
            border-radius: 18px;
            border: 1px solid rgba(23, 50, 77, 0.08);
            background: rgba(255, 255, 255, 0.52);
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.56);
        }

        .taskbar-status::before {
            content: "";
            flex: 0 0 auto;
            width: 0.52rem;
            height: 0.52rem;
            border-radius: 999px;
            background: var(--accent);
            box-shadow: 0 0 0 0.26rem rgba(23, 50, 77, 0.08);
        }

        .taskbar-status-pill {
            flex: 0 0 auto;
            display: inline-flex;
            align-items: center;
            padding: 0.22rem 0.54rem;
            border-radius: 999px;
            background: rgba(23, 50, 77, 0.1);
            color: var(--accent);
            font-size: 0.64rem;
            line-height: 1;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            font-weight: 800;
        }

        .taskbar-status-message {
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            font-size: 0.88rem;
            line-height: 1.25;
            font-weight: 600;
            color: var(--ink);
        }

        .taskbar-status-meta,
        .taskbar-status-timer {
            flex: 0 0 auto;
            font-size: 0.72rem;
            line-height: 1.1;
            color: var(--muted);
            font-weight: 700;
            white-space: nowrap;
        }

        .taskbar-status-timer {
            font-variant-numeric: tabular-nums;
            color: var(--accent);
        }

        .taskbar-status.warning,
        .taskbar-status.throttle {
            border-color: rgba(166, 100, 32, 0.18);
            background: rgba(255, 243, 227, 0.92);
        }

        .taskbar-status.warning::before,
        .taskbar-status.throttle::before {
            background: var(--warn);
            box-shadow: 0 0 0 0.26rem rgba(166, 100, 32, 0.12);
        }

        .taskbar-status.warning .taskbar-status-pill,
        .taskbar-status.throttle .taskbar-status-pill {
            background: rgba(166, 100, 32, 0.14);
            color: var(--warn);
        }

        .taskbar-status.error {
            border-color: rgba(161, 65, 61, 0.16);
            background: rgba(255, 240, 236, 0.94);
        }

        .taskbar-status.error::before {
            background: var(--error);
            box-shadow: 0 0 0 0.26rem rgba(161, 65, 61, 0.1);
        }

        .taskbar-status.error .taskbar-status-pill {
            background: rgba(161, 65, 61, 0.12);
            color: var(--error);
        }

        .taskbar-status.complete {
            border-color: rgba(29, 115, 72, 0.18);
            background: rgba(236, 247, 240, 0.94);
        }

        .taskbar-status.complete::before {
            background: var(--success);
            box-shadow: 0 0 0 0.26rem rgba(29, 115, 72, 0.1);
        }

        .taskbar-status.complete .taskbar-status-pill {
            background: rgba(29, 115, 72, 0.12);
            color: var(--success);
        }

        .taskbar-stage-rail {
            display: flex;
            gap: 0.48rem;
            flex-wrap: wrap;
            align-items: center;
            justify-content: flex-end;
        }

        .taskbar-stage {
            display: inline-flex;
            align-items: center;
            gap: 0.48rem;
            padding: 0.42rem 0.62rem 0.42rem 0.48rem;
            border-radius: 999px;
            border: 1px solid rgba(23, 50, 77, 0.1);
            background: rgba(255, 251, 244, 0.72);
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.46);
            transition: background 180ms ease, border-color 180ms ease, box-shadow 180ms ease, transform 180ms ease;
        }

        .taskbar-stage.is-complete {
            border-color: rgba(29, 115, 72, 0.16);
            background: rgba(236, 247, 240, 0.88);
        }

        .taskbar-stage.is-active {
            border-color: rgba(201, 149, 50, 0.4);
            background: linear-gradient(135deg, var(--accent) 0%, #21486d 100%);
            box-shadow: 0 10px 24px rgba(23, 50, 77, 0.18);
        }

        .taskbar-stage-index {
            display: grid;
            place-items: center;
            width: 1.35rem;
            height: 1.35rem;
            border-radius: 999px;
            background: rgba(23, 50, 77, 0.1);
            color: var(--accent);
            font-size: 0.68rem;
            line-height: 1;
            font-weight: 800;
        }

        .taskbar-stage.is-complete .taskbar-stage-index {
            background: rgba(29, 115, 72, 0.12);
            color: var(--success);
        }

        .taskbar-stage.is-active .taskbar-stage-index {
            background: rgba(245, 230, 196, 0.18);
            color: #fff9ef;
        }

        .taskbar-stage-copy {
            display: grid;
            gap: 0.02rem;
        }

        .taskbar-stage-label {
            font-size: 0.82rem;
            line-height: 1.05;
            font-weight: 700;
            letter-spacing: -0.02em;
            color: var(--ink);
        }

        .taskbar-stage-state {
            font-size: 0.68rem;
            line-height: 1.1;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            font-weight: 700;
            color: rgba(23, 50, 77, 0.68);
        }

        .taskbar-stage.is-active .taskbar-stage-label,
        .taskbar-stage.is-active .taskbar-stage-state {
            color: rgba(255, 255, 255, 0.92);
        }

        .masthead {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 1rem;
            padding-bottom: 0.95rem;
            border-bottom: 1px solid rgba(23, 50, 77, 0.18);
        }

        .masthead-title {
            margin: 0;
            font-size: clamp(1.85rem, 3.4vw, 3.1rem);
            line-height: 0.98;
            letter-spacing: -0.05em;
            font-weight: 700;
            max-width: none;
            white-space: normal;
            color: var(--ink) !important;
            text-shadow: none;
        }

        .masthead-brand {
            margin-bottom: 0.45rem;
            font-size: 0.72rem;
            line-height: 1;
            letter-spacing: 0.22em;
            text-transform: uppercase;
            font-weight: 800;
            color: var(--signal-strong);
        }

        .section-card,
        .status-card,
        .stage-chip {
            position: relative;
            overflow: hidden;
            isolation: isolate;
            border-radius: var(--radius-xl);
            border: 1px solid rgba(23, 50, 77, 0.1);
            background: linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(247, 241, 231, 0.98) 100%);
            box-shadow: 0 18px 36px rgba(15, 23, 32, 0.08);
        }

        .section-card::before,
        .status-card::before,
        .stage-chip::before {
            content: "";
            position: absolute;
            inset: 0 0 auto;
            height: 0.24rem;
            background: linear-gradient(90deg, var(--signal) 0%, rgba(201, 149, 50, 0) 72%);
            opacity: 0.92;
            pointer-events: none;
        }

        .section-card {
            padding: 1.2rem;
        }

        .status-card {
            padding: 1rem 1.1rem;
        }

        .status-card.warning {
            background: var(--warn-soft);
            border-color: var(--warn-line);
        }

        .status-card.warning::before {
            background: linear-gradient(90deg, var(--warn) 0%, rgba(166, 100, 32, 0) 72%);
        }

        .status-card.error {
            background: var(--error-soft);
            border-color: var(--error-line);
        }

        .status-card.error::before {
            background: linear-gradient(90deg, var(--error) 0%, rgba(161, 65, 61, 0) 72%);
        }

        .status-card.throttle {
            border-color: rgba(166, 100, 32, 0.32);
            background:
                linear-gradient(132deg, rgba(255, 245, 222, 0.98) 0%, rgba(255, 237, 204, 0.98) 46%, rgba(242, 224, 189, 0.98) 100%);
        }

        .status-card.throttle::before {
            background:
                repeating-linear-gradient(
                    90deg,
                    var(--warn) 0,
                    var(--warn) 1.1rem,
                    rgba(255, 255, 255, 0) 1.1rem,
                    rgba(255, 255, 255, 0) 1.55rem
                );
            opacity: 1;
        }

        .status-label {
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 700;
            color: var(--accent);
        }

        .status-message {
            margin-top: 0.3rem;
            font-size: 0.98rem;
            line-height: 1.5;
            color: var(--ink);
            font-weight: 500;
        }

        .status-detail {
            margin-top: 0.45rem;
            font-size: 0.84rem;
            line-height: 1.55;
            color: var(--muted);
            font-weight: 500;
        }

        .throttle-status-layout {
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            gap: 1rem;
            align-items: end;
            margin-top: 0.48rem;
        }

        .throttle-status-kicker {
            display: inline-flex;
            align-items: center;
            gap: 0.4rem;
            padding: 0.22rem 0.55rem;
            border-radius: 999px;
            background: rgba(166, 100, 32, 0.12);
            color: var(--warn);
            font-size: 0.68rem;
            font-weight: 800;
            letter-spacing: 0.14em;
            text-transform: uppercase;
        }

        .throttle-status-title {
            margin-top: 0.55rem;
            font-size: clamp(1.2rem, 2vw, 1.6rem);
            line-height: 0.98;
            letter-spacing: -0.04em;
            font-weight: 700;
            color: var(--ink);
        }

        .throttle-status-timer-shell {
            display: grid;
            gap: 0.28rem;
            justify-items: end;
            min-width: 7rem;
        }

        .throttle-status-timer-label {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 700;
            color: var(--warn);
        }

        .throttle-status-timer {
            font-size: clamp(1.35rem, 2.4vw, 2rem);
            line-height: 0.92;
            letter-spacing: -0.05em;
            font-weight: 700;
            color: var(--accent);
            font-variant-numeric: tabular-nums;
        }

        .section-title {
            margin: 0;
            font-size: 1.35rem;
            line-height: 1.05;
            letter-spacing: -0.04em;
            color: var(--ink);
            font-weight: 700;
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
            background: linear-gradient(180deg, rgba(255, 251, 244, 0.94) 0%, rgba(239, 229, 214, 0.94) 100%);
            border-radius: 18px;
        }

        .stage-chip.active {
            flex: 1 1 auto;
            padding: 0.95rem 1rem;
            background: linear-gradient(145deg, var(--accent) 0%, #21486d 100%);
            border-color: rgba(201, 149, 50, 0.48);
            border-radius: var(--radius-xl);
            box-shadow: 0 18px 36px rgba(23, 50, 77, 0.24);
        }

        .stage-chip.complete {
            background: linear-gradient(180deg, rgba(236, 247, 240, 0.96) 0%, rgba(226, 242, 233, 0.96) 100%);
            border-color: rgba(29, 115, 72, 0.18);
        }

        .stage-chip.compact {
            box-shadow: none;
        }

        .stage-chip-number {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            color: var(--accent);
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

        .stage-chip.active .stage-chip-number {
            color: rgba(245, 230, 196, 0.92);
        }

        .stage-chip.active,
        .stage-chip.active .stage-chip-name {
            color: #ffffff;
        }

        .stage-chip.active .stage-chip-state,
        .stage-chip.active .stage-chip-name *,
        .stage-chip.active .stage-chip-state * {
            color: rgba(255, 255, 255, 0.82);
        }

        .stage-chip.compact .stage-chip-name {
            margin-top: 0.28rem;
            font-size: 0.9rem;
            line-height: 1.05;
        }

        .stage-chip.compact .stage-chip-state {
            display: none;
        }

        .chunk-progress {
            margin-top: 1rem;
            padding: 0.9rem 0.95rem;
            border-radius: var(--radius-lg);
            border: 1px solid var(--line);
            background:
                radial-gradient(circle at top left, rgba(201, 149, 50, 0.12), transparent 38%),
                linear-gradient(180deg, rgba(255, 251, 244, 0.94) 0%, rgba(240, 232, 217, 0.98) 100%);
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.72);
        }

        .chunk-progress-track {
            display: flex;
            gap: 0.32rem;
            align-items: center;
            width: 100%;
        }

        .chunk-progress-block {
            display: block;
            position: relative;
            flex: 1 1 0;
            height: 0.96rem;
            border-radius: 999px;
            overflow: hidden;
            transform-origin: center;
            background: linear-gradient(180deg, var(--progress-pending-start) 0%, var(--progress-pending-end) 100%);
            box-shadow: inset 0 0 0 1px rgba(255, 255, 255, 0.58);
        }

        .chunk-progress-block::after {
            content: "";
            position: absolute;
            inset: 0;
            border-radius: inherit;
            background: linear-gradient(90deg, rgba(255, 255, 255, 0) 12%, rgba(255, 255, 255, 0.42) 50%, rgba(255, 255, 255, 0) 88%);
            opacity: 0;
            transform: translateX(-120%);
            pointer-events: none;
        }

        .chunk-progress-block.complete {
            background: linear-gradient(180deg, var(--progress-complete-start) 0%, var(--progress-complete-end) 100%);
            box-shadow:
                inset 0 0 0 1px rgba(255, 255, 255, 0.14),
                0 0 0 1px rgba(37, 58, 45, 0.08);
        }

        .chunk-progress-block.active {
            background: linear-gradient(135deg, var(--progress-active-start) 0%, var(--progress-active-mid) 45%, var(--progress-active-end) 100%);
            box-shadow:
                inset 0 0 0 1px rgba(255, 255, 255, 0.28),
                0 0 0 1px rgba(201, 149, 50, 0.24);
            animation: progressPulse 1.45s cubic-bezier(0.37, 0, 0.22, 1) infinite;
            animation-delay: var(--progress-delay, 0s);
        }

        .chunk-progress-block.active::after {
            opacity: 0.82;
            animation: progressSheen 1.45s ease-in-out infinite;
            animation-delay: var(--progress-delay, 0s);
        }

        .chunk-progress-block.pending {
            background: linear-gradient(180deg, var(--progress-pending-start) 0%, var(--progress-pending-end) 100%);
            opacity: 0.9;
        }

        .workbook-stage-card {
            background:
                radial-gradient(circle at top right, rgba(201, 149, 50, 0.14), transparent 34%),
                linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(244, 236, 223, 0.98) 100%);
        }

        .workbook-retry-banner {
            position: relative;
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            gap: 0.95rem;
            align-items: center;
            margin-top: 0.95rem;
            padding: 1rem 1.05rem 1rem 1.25rem;
            border-radius: 22px;
            border: 1px solid rgba(166, 100, 32, 0.28);
            background:
                linear-gradient(135deg, rgba(255, 245, 222, 0.98) 0%, rgba(255, 236, 198, 0.96) 100%);
            overflow: hidden;
        }

        .workbook-retry-banner.is-error {
            border-color: rgba(161, 65, 61, 0.26);
            background:
                linear-gradient(135deg, rgba(255, 240, 238, 0.98) 0%, rgba(255, 228, 224, 0.96) 100%);
        }

        .workbook-retry-banner::after {
            content: "";
            position: absolute;
            inset: auto -2rem -2rem auto;
            width: 8rem;
            height: 8rem;
            background: radial-gradient(circle, rgba(166, 100, 32, 0.14) 0%, rgba(166, 100, 32, 0) 70%);
            pointer-events: none;
        }

        .workbook-retry-banner.is-error::after {
            background: radial-gradient(circle, rgba(161, 65, 61, 0.16) 0%, rgba(161, 65, 61, 0) 70%);
        }

        .workbook-retry-notch {
            position: absolute;
            inset: 0 auto 0 0;
            width: 0.5rem;
            background:
                repeating-linear-gradient(
                    180deg,
                    rgba(166, 100, 32, 0.95) 0,
                    rgba(166, 100, 32, 0.95) 0.75rem,
                    rgba(240, 214, 166, 0.25) 0.75rem,
                    rgba(240, 214, 166, 0.25) 1.05rem
                );
        }

        .workbook-retry-banner.is-error .workbook-retry-notch {
            background:
                repeating-linear-gradient(
                    180deg,
                    rgba(161, 65, 61, 0.95) 0,
                    rgba(161, 65, 61, 0.95) 0.75rem,
                    rgba(255, 220, 214, 0.28) 0.75rem,
                    rgba(255, 220, 214, 0.28) 1.05rem
                );
        }

        .workbook-retry-copy {
            position: relative;
            z-index: 1;
            min-width: 0;
        }

        .workbook-retry-kicker {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--warn);
        }

        .workbook-retry-banner.is-error .workbook-retry-kicker,
        .workbook-retry-banner.is-error .workbook-retry-timer-label,
        .workbook-retry-banner.is-error .workbook-retry-timer {
            color: var(--error);
        }

        .workbook-retry-title {
            margin-top: 0.36rem;
            font-size: clamp(1.1rem, 1.8vw, 1.4rem);
            line-height: 1;
            letter-spacing: -0.04em;
            font-weight: 700;
            color: var(--ink);
        }

        .workbook-retry-body {
            margin-top: 0.45rem;
            font-size: 0.84rem;
            line-height: 1.5;
            color: var(--muted);
        }

        .workbook-retry-detail {
            margin-top: 0.5rem;
            font-size: 0.8rem;
            line-height: 1.55;
            color: var(--ink);
        }

        .workbook-retry-timer-shell {
            position: relative;
            z-index: 1;
            display: grid;
            gap: 0.22rem;
            justify-items: end;
            min-width: 6.8rem;
        }

        .workbook-retry-timer-label {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.14em;
            font-weight: 800;
            color: var(--warn);
        }

        .workbook-retry-timer {
            font-size: clamp(1.35rem, 2.2vw, 1.9rem);
            line-height: 0.92;
            letter-spacing: -0.05em;
            font-weight: 700;
            color: var(--accent);
            font-variant-numeric: tabular-nums;
        }

        .workbook-stage-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 1rem;
        }

        .workbook-stage-pill {
            flex: 0 0 auto;
            padding: 0.45rem 0.75rem;
            border-radius: 999px;
            border: 1px solid rgba(201, 149, 50, 0.38);
            background: rgba(245, 230, 196, 0.74);
            color: var(--accent);
            font-size: 0.75rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }

        .workbook-phase-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.85rem;
            margin-top: 0.95rem;
        }

        .workbook-stage-body {
            display: grid;
            grid-template-columns: minmax(0, 1fr) minmax(18rem, 24rem);
            gap: 1rem;
            align-items: start;
            margin-top: 0.95rem;
        }

        .workbook-phase-card {
            padding: 1rem;
            border-radius: var(--radius-lg);
            border: 1px solid var(--line);
            background: rgba(255, 251, 244, 0.92);
            transition: transform 180ms ease, border-color 180ms ease, background 180ms ease, box-shadow 180ms ease;
        }

        .workbook-phase-card.active {
            border-color: rgba(201, 149, 50, 0.36);
            background: linear-gradient(180deg, rgba(255, 248, 233, 0.98) 0%, rgba(244, 236, 223, 0.98) 100%);
            box-shadow: 0 14px 32px rgba(201, 149, 50, 0.16);
        }

        .workbook-phase-card.complete {
            border-color: rgba(29, 115, 72, 0.18);
            background: rgba(236, 247, 240, 0.96);
        }

        .workbook-phase-card.attention {
            border-color: var(--warn-line);
            background: rgba(255, 243, 227, 0.96);
        }

        .workbook-phase-kicker {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            color: var(--muted);
            font-weight: 700;
        }

        .workbook-phase-title {
            margin-top: 0.38rem;
            font-size: 1.02rem;
            line-height: 1.1;
            font-weight: 700;
            letter-spacing: -0.03em;
            color: var(--ink);
        }

        .workbook-phase-state {
            margin-top: 0.42rem;
            font-size: 0.82rem;
            color: var(--muted);
        }

        .workbook-phase-meter {
            position: relative;
            height: 0.7rem;
            margin-top: 0.8rem;
            border-radius: 999px;
            overflow: hidden;
            background: rgba(15, 23, 32, 0.1);
        }

        .workbook-phase-meter-fill {
            display: block;
            height: 100%;
            border-radius: inherit;
            background: linear-gradient(135deg, var(--signal) 0%, #e0bb64 48%, var(--accent) 100%);
            transition: width 220ms ease;
        }

        .workbook-phase-card.complete .workbook-phase-meter-fill {
            background: linear-gradient(135deg, #2e7d57 0%, #67b48b 100%);
        }

        .workbook-phase-card.attention .workbook-phase-meter-fill {
            background: linear-gradient(135deg, #b97822 0%, #dfa760 100%);
        }

        .workbook-phase-copy {
            margin-top: 0.62rem;
            font-size: 0.82rem;
            line-height: 1.45;
            color: var(--muted);
        }

        .workbook-stage-message {
            margin-top: 0.9rem;
            padding-top: 0.85rem;
            border-top: 1px solid rgba(23, 50, 77, 0.12);
        }

        .immersive-workbook-shell {
            position: relative;
            min-height: auto;
            margin: 0;
            padding: 0;
            background: transparent;
            overflow: visible;
        }

        .immersive-workbook-shell::before {
            display: none;
        }

        .immersive-workbook-grid {
            position: relative;
            z-index: 1;
            display: grid;
            grid-template-columns: minmax(18.5rem, 21rem) minmax(0, 1fr);
            gap: 1rem;
            min-height: 0;
            align-items: start;
        }

        .immersive-pane {
            min-width: 0;
            min-height: 0;
        }

        .immersive-pane-label {
            margin-bottom: 0.48rem;
            font-size: 0.66rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--muted);
        }

        .immersive-agent-pane .immersive-pane-label {
            color: var(--muted);
        }

        .immersive-sheet-pane {
            position: relative;
            padding: 0;
            border-radius: 0;
            border: none;
            background: transparent;
            box-shadow: none;
        }

        .immersive-sheet-pane::before {
            display: none;
        }

        .immersive-sheet-pane > * {
            position: relative;
            z-index: 1;
        }

        .immersive-alert-row {
            position: relative;
            z-index: 1;
            margin-bottom: 0.95rem;
        }

        .workbook-setup-shell {
            position: relative;
            overflow: hidden;
            min-height: clamp(34rem, 72vh, 48rem);
            padding: clamp(1.1rem, 2vw, 1.65rem);
            border-radius: 32px;
            border: 1px solid rgba(218, 229, 241, 0.14);
            background:
                radial-gradient(circle at 14% 18%, rgba(241, 193, 95, 0.2), transparent 20%),
                radial-gradient(circle at 84% 12%, rgba(120, 172, 221, 0.16), transparent 22%),
                linear-gradient(135deg, rgba(7, 18, 31, 0.98) 0%, rgba(10, 32, 56, 0.98) 44%, rgba(16, 46, 71, 0.96) 100%);
            box-shadow:
                inset 0 1px 0 rgba(255, 255, 255, 0.06),
                0 34px 72px rgba(8, 19, 30, 0.28);
            isolation: isolate;
        }

        .workbook-setup-shell::before {
            content: "";
            position: absolute;
            inset: 0;
            background:
                linear-gradient(135deg, rgba(255, 255, 255, 0.04), transparent 32%),
                repeating-linear-gradient(
                    180deg,
                    rgba(255, 255, 255, 0.05) 0,
                    rgba(255, 255, 255, 0.05) 1px,
                    rgba(255, 255, 255, 0) 1px,
                    rgba(255, 255, 255, 0) 30px
                );
            opacity: 0.16;
            pointer-events: none;
        }

        .workbook-setup-grid {
            position: relative;
            z-index: 1;
            display: grid;
            grid-template-columns: minmax(0, 1.3fr) minmax(20rem, 0.92fr);
            gap: 1rem;
            min-height: inherit;
        }

        .workbook-setup-hero,
        .workbook-setup-console {
            position: relative;
            padding: 1.2rem;
            border-radius: 28px;
            border: 1px solid rgba(218, 229, 241, 0.1);
            background:
                linear-gradient(180deg, rgba(10, 24, 40, 0.84) 0%, rgba(12, 30, 48, 0.72) 100%);
            backdrop-filter: blur(18px);
        }

        .workbook-setup-hero {
            display: grid;
            align-content: space-between;
            gap: 1.1rem;
        }

        .workbook-setup-console {
            display: grid;
            gap: 0.9rem;
            align-content: start;
        }

        .workbook-setup-eyebrow,
        .workbook-setup-console-kicker,
        .workbook-setup-stage-kicker,
        .workbook-setup-step-kicker,
        .workbook-setup-metric span,
        .workbook-setup-radar-core span {
            font-size: 0.64rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            font-weight: 800;
        }

        .workbook-setup-eyebrow,
        .workbook-setup-console-kicker,
        .workbook-setup-stage-kicker,
        .workbook-setup-step-kicker {
            color: rgba(232, 217, 184, 0.8);
        }

        .workbook-setup-title {
            margin-top: 0.45rem;
            max-width: 11ch;
            font-family: var(--display-font) !important;
            font-size: clamp(3rem, 8vw, 5.4rem);
            line-height: 0.92;
            letter-spacing: -0.07em;
            color: #f8fbff;
            text-wrap: balance;
        }

        .workbook-setup-copy,
        .workbook-setup-console-copy {
            color: rgba(226, 235, 244, 0.82);
            line-height: 1.6;
        }

        .workbook-setup-copy {
            max-width: 38rem;
            font-size: 0.98rem;
        }

        .workbook-setup-console-title {
            margin-top: 0.32rem;
            font-size: 1.18rem;
            line-height: 1.1;
            letter-spacing: -0.04em;
            font-weight: 700;
            color: #f7fbff;
        }

        .workbook-setup-console-copy {
            font-size: 0.84rem;
        }

        .workbook-setup-chip-row {
            display: flex;
            flex-wrap: wrap;
            gap: 0.55rem;
            margin-top: 0.2rem;
        }

        .workbook-setup-chip {
            min-width: min(14rem, 100%);
            padding: 0.68rem 0.8rem 0.74rem;
            border-radius: 18px;
            border: 1px solid rgba(218, 229, 241, 0.11);
            background: rgba(255, 255, 255, 0.05);
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.04);
        }

        .workbook-setup-chip span {
            display: block;
            font-size: 0.58rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            font-weight: 800;
            color: rgba(218, 229, 241, 0.62);
        }

        .workbook-setup-chip strong {
            display: block;
            margin-top: 0.3rem;
            font-size: 0.88rem;
            line-height: 1.35;
            color: #fff7e6;
            font-weight: 700;
            overflow-wrap: anywhere;
        }

        .workbook-setup-stage {
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            gap: 0.85rem;
            align-items: end;
            padding: 0.95rem 1rem;
            border-radius: 22px;
            border: 1px solid rgba(241, 193, 95, 0.15);
            background: rgba(255, 255, 255, 0.05);
        }

        .workbook-setup-stage-title {
            margin-top: 0.36rem;
            font-size: 1.08rem;
            line-height: 1.08;
            letter-spacing: -0.04em;
            font-weight: 700;
            color: #fff7e6;
        }

        .workbook-setup-stage-copy {
            margin-top: 0.45rem;
            max-width: 34rem;
            font-size: 0.82rem;
            line-height: 1.55;
            color: rgba(226, 235, 244, 0.78);
        }

        .workbook-setup-stage-indicator {
            justify-self: end;
            padding: 0.38rem 0.64rem;
            border-radius: 999px;
            border: 1px solid rgba(241, 193, 95, 0.18);
            background: rgba(241, 193, 95, 0.08);
            color: rgba(255, 247, 230, 0.92);
            font-size: 0.68rem;
            font-weight: 800;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            white-space: nowrap;
        }

        .workbook-setup-stage-meter {
            position: relative;
            height: 0.62rem;
            margin-top: 0.8rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.08);
            overflow: hidden;
        }

        .workbook-setup-stage-fill {
            position: absolute;
            inset: 0 auto 0 -18%;
            width: 44%;
            border-radius: inherit;
            background: linear-gradient(90deg, rgba(241, 193, 95, 0) 0%, rgba(241, 193, 95, 0.92) 48%, rgba(255, 255, 255, 0.22) 100%);
            box-shadow: 0 0 28px rgba(241, 193, 95, 0.18);
            animation: setupSweep 2.8s ease-in-out infinite;
        }

        .workbook-setup-radar {
            position: relative;
            display: grid;
            place-items: center;
            min-height: clamp(15rem, 34vw, 19rem);
            overflow: hidden;
            border-radius: 28px;
            border: 1px solid rgba(218, 229, 241, 0.11);
            background: radial-gradient(circle at center, rgba(19, 52, 79, 0.9) 0%, rgba(7, 18, 31, 0.84) 66%, rgba(4, 10, 18, 0.98) 100%);
        }

        .workbook-setup-radar-sweep {
            position: absolute;
            inset: -40%;
            background: conic-gradient(
                from 180deg,
                rgba(241, 193, 95, 0) 0deg,
                rgba(241, 193, 95, 0.24) 42deg,
                rgba(241, 193, 95, 0) 90deg
            );
            mix-blend-mode: screen;
            filter: blur(2px);
            animation: setupRadar 4.8s linear infinite;
        }

        .workbook-setup-radar-ring {
            position: absolute;
            left: 50%;
            top: 50%;
            border-radius: 999px;
            border: 1px solid rgba(218, 229, 241, 0.12);
            transform: translate(-50%, -50%);
        }

        .workbook-setup-radar-ring.ring-a {
            width: 11rem;
            height: 11rem;
        }

        .workbook-setup-radar-ring.ring-b {
            width: 17rem;
            height: 17rem;
            opacity: 0.8;
        }

        .workbook-setup-radar-ring.ring-c {
            width: 23rem;
            height: 23rem;
            opacity: 0.54;
        }

        .workbook-setup-radar-core {
            position: relative;
            display: grid;
            gap: 0.42rem;
            justify-items: center;
            padding: 1rem 1.2rem;
            border-radius: 22px;
            border: 1px solid rgba(241, 193, 95, 0.2);
            background: rgba(8, 19, 30, 0.82);
            box-shadow:
                inset 0 1px 0 rgba(255, 255, 255, 0.05),
                0 18px 36px rgba(0, 0, 0, 0.26);
            text-align: center;
            animation: setupFloat 3.6s ease-in-out infinite;
        }

        .workbook-setup-radar-core span {
            color: rgba(232, 217, 184, 0.72);
        }

        .workbook-setup-radar-core strong {
            max-width: 14ch;
            font-size: 1.1rem;
            line-height: 1.2;
            letter-spacing: -0.03em;
            color: #fff7e6;
            font-weight: 700;
            text-wrap: balance;
        }

        .workbook-setup-radar-node {
            position: absolute;
            width: 0.7rem;
            height: 0.7rem;
            border-radius: 999px;
            background: #f1c15f;
            box-shadow: 0 0 0 0 rgba(241, 193, 95, 0.24);
            animation: agentPulse 1.6s ease-in-out infinite;
        }

        .workbook-setup-radar-node.node-a {
            top: 21%;
            left: 29%;
        }

        .workbook-setup-radar-node.node-b {
            top: 33%;
            right: 26%;
            animation-delay: 0.42s;
        }

        .workbook-setup-radar-node.node-c {
            bottom: 22%;
            left: 34%;
            animation-delay: 0.84s;
        }

        .workbook-setup-radar-node.node-d {
            bottom: 28%;
            right: 30%;
            animation-delay: 1.12s;
        }

        .workbook-setup-step-list {
            display: grid;
            gap: 0.72rem;
        }

        .workbook-setup-step {
            position: relative;
            padding: 0.92rem 1rem 0.98rem 3rem;
            border-radius: 22px;
            border: 1px solid rgba(218, 229, 241, 0.1);
            background: rgba(255, 255, 255, 0.05);
            overflow: hidden;
            animation: rise 240ms ease both;
            animation-delay: var(--setup-delay, 0s);
        }

        .workbook-setup-step::before {
            content: "";
            position: absolute;
            left: 1.14rem;
            top: 1.08rem;
            width: 0.78rem;
            height: 0.78rem;
            border-radius: 999px;
            background: #f1c15f;
            box-shadow: 0 0 0 0 rgba(241, 193, 95, 0.22);
            animation: agentPulse 1.4s ease-in-out infinite;
            animation-delay: var(--setup-delay, 0s);
        }

        .workbook-setup-step::after {
            content: "";
            position: absolute;
            left: 1.51rem;
            top: 2rem;
            bottom: -0.78rem;
            width: 1px;
            background: linear-gradient(180deg, rgba(241, 193, 95, 0.44), rgba(255, 255, 255, 0));
        }

        .workbook-setup-step:last-child::after {
            display: none;
        }

        .workbook-setup-step-title {
            margin-top: 0.3rem;
            font-size: 0.98rem;
            line-height: 1.15;
            font-weight: 700;
            letter-spacing: -0.03em;
            color: #f8fbff;
        }

        .workbook-setup-step-copy {
            margin-top: 0.36rem;
            font-size: 0.82rem;
            line-height: 1.52;
            color: rgba(226, 235, 244, 0.76);
        }

        .workbook-setup-step-meter {
            position: relative;
            height: 0.42rem;
            margin-top: 0.72rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.08);
            overflow: hidden;
        }

        .workbook-setup-step-fill {
            position: absolute;
            inset: 0 auto 0 -12%;
            width: 58%;
            border-radius: inherit;
            background: linear-gradient(90deg, rgba(241, 193, 95, 0.08) 0%, rgba(241, 193, 95, 0.88) 54%, rgba(255, 255, 255, 0.18) 100%);
            animation: setupMeter 1.9s ease-in-out infinite;
            animation-delay: var(--setup-delay, 0s);
        }

        .workbook-setup-metrics {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.72rem;
        }

        .workbook-setup-metric {
            padding: 0.86rem 0.94rem 0.9rem;
            border-radius: 18px;
            border: 1px solid rgba(218, 229, 241, 0.1);
            background: rgba(255, 255, 255, 0.04);
        }

        .workbook-setup-metric span {
            display: block;
            color: rgba(218, 229, 241, 0.58);
        }

        .workbook-setup-metric strong {
            display: block;
            margin-top: 0.32rem;
            font-size: 0.94rem;
            line-height: 1.3;
            color: #fff7e6;
            font-weight: 700;
            overflow-wrap: anywhere;
        }

        .sheet-desk {
            display: grid;
            gap: 0.9rem;
            height: auto;
            padding: 1rem;
            border-radius: var(--radius-xl);
            border: 1px solid rgba(23, 50, 77, 0.1);
            background: linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(247, 241, 231, 0.98) 100%);
            box-shadow: 0 18px 34px rgba(15, 23, 32, 0.08);
        }

        .sheet-desk-header {
            display: grid;
            gap: 0.55rem;
        }

        .sheet-desk-kicker {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            font-weight: 800;
            color: var(--signal-strong);
        }

        .sheet-desk-title {
            margin-top: 0.18rem;
            font-family: var(--display-font) !important;
            font-size: clamp(1.75rem, 2.4vw, 2.4rem);
            line-height: 0.95;
            letter-spacing: -0.06em;
            color: var(--ink);
        }

        .sheet-desk-copy {
            margin-top: 0;
            max-width: none;
            color: var(--muted);
            font-size: 0.88rem;
            line-height: 1.5;
        }

        .sheet-desk-summary {
            display: grid;
            grid-template-columns: minmax(0, 1.2fr) minmax(0, 1fr);
            gap: 0.75rem;
        }

        .sheet-desk-summary-card {
            display: grid;
            gap: 0.52rem;
            padding: 0.9rem 0.95rem 0.95rem;
            border-radius: 20px;
            border: 1px solid rgba(23, 50, 77, 0.08);
            background: rgba(255, 255, 255, 0.76);
        }

        .sheet-desk-summary-card.is-primary {
            background: linear-gradient(135deg, rgba(255, 248, 233, 0.98) 0%, rgba(244, 236, 223, 0.98) 100%);
            border-color: rgba(201, 149, 50, 0.3);
        }

        .sheet-desk-summary-label {
            display: block;
            font-size: 0.64rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--muted);
        }

        .sheet-desk-summary-title {
            display: block;
            font-size: 1.1rem;
            line-height: 1.15;
            letter-spacing: -0.03em;
            font-weight: 700;
            color: var(--ink);
        }

        .sheet-desk-summary-copy {
            font-size: 0.8rem;
            line-height: 1.45;
            color: var(--muted);
        }

        .sheet-desk-summary-meter {
            position: relative;
            height: 0.5rem;
            border-radius: 999px;
            overflow: hidden;
            background: rgba(15, 23, 32, 0.08);
        }

        .sheet-desk-summary-meter span {
            display: block;
            height: 100%;
            border-radius: inherit;
            background: linear-gradient(135deg, var(--signal) 0%, #e0bb64 48%, var(--accent) 100%);
        }

        .sheet-desk-surface {
            display: grid;
            grid-template-columns: 1fr;
            gap: 0.8rem;
            min-height: 0;
            flex: 1 1 auto;
        }

        .sheet-preview-card,
        .sheet-queue-card {
            padding: 0.95rem;
            border-radius: 20px;
            border: 1px solid rgba(23, 50, 77, 0.1);
            background: rgba(255, 253, 248, 0.96);
            box-shadow: none;
        }

        .sheet-preview-card {
            min-height: 0;
        }

        .sheet-preview-head,
        .sheet-queue-head {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 0.7rem;
        }

        .sheet-preview-kicker,
        .sheet-queue-kicker {
            font-size: 0.64rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            font-weight: 800;
            color: var(--muted);
        }

        .sheet-preview-title,
        .sheet-queue-title {
            margin-top: 0.35rem;
            font-size: 1.08rem;
            line-height: 1.05;
            letter-spacing: -0.04em;
            font-weight: 700;
            color: var(--ink);
        }

        .sheet-preview-pill {
            display: inline-flex;
            align-items: center;
            min-height: 2rem;
            padding: 0.3rem 0.65rem;
            border-radius: 999px;
            border: 1px solid rgba(201, 149, 50, 0.32);
            background: rgba(245, 230, 196, 0.74);
            color: var(--accent);
            font-size: 0.66rem;
            font-weight: 800;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            white-space: nowrap;
        }

        .sheet-preview-copy,
        .sheet-queue-copy {
            margin-top: 0.7rem;
            color: var(--muted);
            font-size: 0.84rem;
            line-height: 1.5;
        }

        .sheet-cell-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(15rem, 1fr));
            gap: 0.72rem;
            margin-top: 0.9rem;
        }

        .sheet-cell {
            padding: 0.8rem 0.82rem;
            border-radius: 18px;
            border: 1px solid rgba(23, 50, 77, 0.08);
            background: rgba(255, 255, 255, 0.7);
            min-width: 0;
        }

        .sheet-cell.is-filled {
            border-color: rgba(23, 50, 77, 0.14);
            background:
                linear-gradient(135deg, rgba(255, 251, 244, 0.98) 0%, rgba(239, 229, 214, 0.98) 100%);
        }

        .sheet-cell-label {
            display: inline-flex;
            align-items: center;
            min-height: 1.55rem;
            padding: 0 0.42rem;
            border-radius: 999px;
            background: rgba(23, 50, 77, 0.08);
            color: var(--accent);
            font-size: 0.62rem;
            font-weight: 800;
            letter-spacing: 0.12em;
            text-transform: uppercase;
        }

        .sheet-cell-title {
            margin-top: 0.62rem;
            font-size: 0.88rem;
            line-height: 1.25;
            letter-spacing: -0.01em;
            font-weight: 700;
            color: var(--ink);
        }

        .sheet-cell-value {
            margin-top: 0.34rem;
            font-size: 1rem;
            line-height: 1.35;
            color: var(--ink);
            font-weight: 600;
            overflow-wrap: anywhere;
        }

        .sheet-cell-detail {
            margin-top: 0.34rem;
            font-size: 0.74rem;
            line-height: 1.4;
            color: var(--muted);
            overflow-wrap: anywhere;
        }

        .sheet-cell.is-empty .sheet-cell-value {
            color: #7a8794;
            font-weight: 500;
        }

        .sheet-cell-empty-state {
            grid-column: 1 / -1;
            padding: 1rem;
            background: linear-gradient(180deg, rgba(255, 255, 255, 0.82) 0%, rgba(247, 241, 231, 0.98) 100%);
        }

        .sheet-cell-empty-state .sheet-cell-title {
            margin-top: 0;
            font-size: 1rem;
        }

        .sheet-cell-empty-state .sheet-cell-detail {
            margin-top: 0;
            font-size: 0.82rem;
        }

        .sheet-queue-list {
            display: grid;
            gap: 0.55rem;
            margin-top: 0.9rem;
        }

        .sheet-queue-item {
            padding: 0.75rem 0.82rem;
            border-radius: 18px;
            border: 1px solid rgba(23, 50, 77, 0.08);
            background: rgba(255, 255, 255, 0.66);
        }

        .sheet-queue-item.active {
            border-color: rgba(201, 149, 50, 0.34);
            background: linear-gradient(135deg, rgba(255, 248, 233, 0.98) 0%, rgba(244, 236, 223, 0.98) 100%);
            box-shadow: none;
        }

        .sheet-queue-item.complete {
            border-color: rgba(29, 115, 72, 0.18);
            background: rgba(236, 247, 240, 0.88);
        }

        .sheet-queue-item.needs-input {
            border-color: rgba(166, 100, 32, 0.24);
            background: rgba(255, 243, 227, 0.92);
        }

        .sheet-queue-item.skipped {
            border-style: dashed;
        }

        .sheet-queue-name {
            font-size: 0.96rem;
            line-height: 1.2;
            font-weight: 700;
            color: var(--ink);
        }

        .sheet-queue-state {
            margin-top: 0.32rem;
            font-size: 0.64rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--muted);
        }

        .sheet-queue-meta {
            margin-top: 0.36rem;
            font-size: 0.76rem;
            line-height: 1.45;
            color: var(--muted);
        }

        .agent-window {
            position: relative;
            overflow: hidden;
            padding: 1rem;
            border-radius: 24px;
            border: 1px solid rgba(23, 50, 77, 0.12);
            background: linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(241, 234, 222, 0.98) 100%);
            color: var(--ink);
            box-shadow: 0 18px 34px rgba(15, 23, 32, 0.08);
        }

        .agent-window::before {
            display: none;
        }

        .agent-window > * {
            position: relative;
            z-index: 1;
        }

        .agent-window-head {
            display: flex;
            justify-content: space-between;
            gap: 0.85rem;
            align-items: flex-start;
        }

        .agent-window-head-meta {
            display: grid;
            gap: 0.45rem;
            justify-items: end;
        }

        .agent-window-kicker {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            font-weight: 800;
            color: var(--signal-strong);
        }

        .agent-window-title {
            margin-top: 0.35rem;
            font-size: 1.28rem;
            line-height: 1.02;
            letter-spacing: -0.04em;
            font-weight: 700;
            color: var(--ink);
        }

        .agent-window-status {
            display: inline-flex;
            align-items: center;
            gap: 0.38rem;
            padding: 0.38rem 0.62rem;
            border-radius: 999px;
            border: 1px solid rgba(23, 50, 77, 0.12);
            background: rgba(23, 50, 77, 0.06);
            color: var(--ink);
            font-size: 0.68rem;
            font-weight: 800;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            white-space: nowrap;
        }

        .agent-window-status-dot {
            width: 0.5rem;
            height: 0.5rem;
            border-radius: 999px;
            background: #9fb1c4;
            box-shadow: 0 0 0 0 rgba(159, 177, 196, 0.35);
        }

        .agent-window-status.running .agent-window-status-dot,
        .agent-window-status.validating .agent-window-status-dot {
            background: #f1c15f;
            box-shadow: 0 0 0 0 rgba(241, 193, 95, 0.3);
            animation: agentPulse 1.2s ease-in-out infinite;
        }

        .agent-window-status.retrying .agent-window-status-dot,
        .agent-window-status.review .agent-window-status-dot,
        .agent-window-status.needs_input .agent-window-status-dot {
            background: #f0a657;
        }

        .agent-window-status.complete .agent-window-status-dot {
            background: #73c792;
        }

        .agent-window-status.error .agent-window-status-dot {
            background: #e38f87;
        }

        .agent-window-metrics {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.7rem;
            margin-top: 0.95rem;
        }

        .agent-window-metric {
            padding: 0.72rem 0.78rem;
            border-radius: 16px;
            background: rgba(255, 255, 255, 0.72);
            border: 1px solid rgba(23, 50, 77, 0.08);
            box-shadow: none;
        }

        .agent-window-metric span {
            display: block;
            font-size: 0.62rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--muted);
        }

        .agent-window-metric strong {
            display: block;
            margin-top: 0.38rem;
            font-size: 0.9rem;
            line-height: 1.3;
            color: var(--ink);
            overflow-wrap: anywhere;
        }

        .agent-window-copy {
            margin-top: 0.9rem;
            font-size: 0.88rem;
            line-height: 1.52;
            color: var(--muted);
        }

        .agent-window-summary {
            margin-top: 0.85rem;
            padding: 0.82rem 0.88rem;
            border-radius: 18px;
            border: 1px solid rgba(23, 50, 77, 0.08);
            background: rgba(255, 255, 255, 0.74);
        }

        .agent-window-summary.empty {
            background: rgba(242, 236, 226, 0.76);
        }

        .agent-window-summary-kicker {
            font-size: 0.62rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--muted);
        }

        .agent-window-summary-copy {
            margin-top: 0.36rem;
            font-size: 0.84rem;
            line-height: 1.5;
            color: var(--ink);
            white-space: pre-wrap;
        }

        .agent-window-summary.empty .agent-window-summary-copy {
            color: var(--muted);
        }

        .agent-window-timeline {
            display: grid;
            gap: 0.55rem;
            margin-top: 0.95rem;
        }

        .agent-window-line {
            padding: 0.72rem 0.8rem;
            border-radius: 16px;
            border: 1px solid rgba(23, 50, 77, 0.08);
            background: rgba(255, 255, 255, 0.72);
        }

        .agent-window-line.success {
            border-color: rgba(29, 115, 72, 0.16);
            background: rgba(236, 247, 240, 0.9);
        }

        .agent-window-line.warn {
            border-color: rgba(166, 100, 32, 0.16);
            background: rgba(255, 243, 227, 0.92);
        }

        .agent-window-line.error {
            border-color: rgba(161, 65, 61, 0.16);
            background: rgba(255, 240, 236, 0.92);
        }

        .agent-window-line-kicker {
            font-size: 0.62rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--muted);
        }

        .agent-window-line-copy {
            margin-top: 0.28rem;
            font-size: 0.82rem;
            line-height: 1.45;
            color: var(--ink);
        }

        .agent-window-line.empty .agent-window-line-copy {
            color: var(--muted);
        }

        .agent-window-footer {
            margin-top: 0.8rem;
            padding-top: 0.75rem;
            border-top: 1px solid rgba(23, 50, 77, 0.1);
            font-size: 0.74rem;
            line-height: 1.45;
            color: var(--muted);
        }

        .immersive-agent-pane .agent-window {
            min-height: 0;
            padding: 1rem;
        }

        .question-card {
            background:
                linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(246, 238, 225, 0.98) 100%);
        }

        .question-shell {
            display: grid;
            gap: 0.65rem;
        }

        .question-shell.is-inline {
            margin-top: 0.85rem;
        }

        .question-shell.is-modal {
            margin-top: 0.15rem;
        }

        .question-panel {
            padding: 0.95rem 1rem 0.9rem;
            border-radius: var(--radius-xl);
            border: 1px solid rgba(23, 50, 77, 0.1);
            background:
                linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(248, 242, 233, 0.98) 100%);
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.86);
        }

        .question-header {
            display: grid;
            gap: 0.55rem;
        }

        .question-header-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem;
            flex-wrap: wrap;
        }

        .question-kicker {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            color: var(--signal-strong);
            font-weight: 700;
        }

        .question-meta-strip {
            display: flex;
            flex-wrap: wrap;
            gap: 0.4rem;
            align-items: center;
        }

        .question-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.28rem 0.55rem;
            border-radius: 999px;
            border: 1px solid rgba(23, 50, 77, 0.14);
            background: rgba(23, 50, 77, 0.05);
            color: var(--accent);
            font-size: 0.66rem;
            font-weight: 700;
            letter-spacing: 0.1em;
            text-transform: uppercase;
            white-space: nowrap;
        }

        .question-chip.signal {
            border-color: rgba(201, 149, 50, 0.3);
            background: rgba(245, 230, 196, 0.88);
        }

        .question-title {
            margin: 0;
            font-family: var(--display-font) !important;
            font-size: clamp(1.45rem, 1.95vw, 1.95rem);
            line-height: 1.02;
            letter-spacing: -0.04em;
            color: var(--ink);
        }

        [data-testid="stDialog"] .question-panel {
            padding: 0.85rem 0.95rem 0.75rem;
        }

        [data-testid="stDialog"] .question-title {
            font-size: clamp(1.75rem, 2.35vw, 2.35rem);
            line-height: 0.97;
        }

        [data-testid="stDialog"] label[data-testid="stWidgetLabel"] {
            font-size: 0.68rem !important;
            letter-spacing: 0.16em !important;
            text-transform: uppercase;
            font-weight: 800 !important;
            color: var(--accent) !important;
        }

        [data-testid="stDialog"] [data-testid="stMarkdownContainer"] p {
            color: var(--ink) !important;
        }

        [data-testid="stDialog"] form {
            margin-top: 0.15rem;
        }

        .ocr-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(18.25rem, 1fr));
            gap: 0.85rem;
            align-items: stretch;
            margin-top: 0.95rem;
        }

        .ocr-row {
            display: grid;
            grid-template-columns: minmax(0, 1fr) auto;
            grid-template-areas:
                "pipe timing"
                "status status"
                "warning warning"
                "preview preview";
            gap: 0.78rem 1rem;
            align-items: start;
            height: 100%;
            padding: 0.95rem 1rem 1rem;
            border-radius: var(--radius-lg);
            background: var(--surface-soft);
            border: 1px solid var(--line);
            position: relative;
            overflow: hidden;
        }

        .ocr-row.just-completed {
            animation: laneCelebrate 900ms ease;
        }

        .ocr-row.inactive {
            background: #f4ede1;
            border-color: #dfd1bc;
        }

        .ocr-row.has-warning {
            border-color: rgba(191, 129, 30, 0.34);
            background:
                linear-gradient(180deg, rgba(255, 249, 239, 0.98) 0%, rgba(246, 237, 223, 0.98) 100%);
            box-shadow: inset 0 0 0 1px rgba(223, 167, 96, 0.16);
        }

        .ocr-row.has-warning .ocr-timing {
            padding-right: 2.4rem;
        }

        .ocr-row.inactive .ocr-pipe,
        .ocr-row.inactive .ocr-status,
        .ocr-row.inactive .ocr-timing,
        .ocr-row.inactive .ocr-preview-caption {
            color: #6d7885;
        }

        .ocr-pipe {
            grid-area: pipe;
            display: flex;
            align-items: center;
            gap: 0.55rem;
            font-weight: 700;
            color: var(--ink);
            min-width: 0;
        }

        .ocr-dot {
            width: 0.72rem;
            height: 0.72rem;
            border-radius: 999px;
            background: #c4ced8;
            flex: 0 0 auto;
        }

        .ocr-dot.running,
        .ocr-dot.retrying {
            background: var(--signal);
        }

        .ocr-dot.paused {
            background: var(--warn);
        }

        .ocr-dot.complete {
            background: var(--success);
        }

        .ocr-dot.failed {
            background: var(--error);
        }

        .ocr-status {
            grid-area: status;
            min-width: 0;
            color: var(--muted);
            font-size: 0.88rem;
            line-height: 1.45;
        }

        .ocr-status strong {
            display: block;
            margin-bottom: 0.12rem;
            font-size: 0.97rem;
            line-height: 1.2;
            letter-spacing: -0.03em;
            color: var(--ink);
        }

        .ocr-detail {
            display: block;
        }

        .ocr-warning-card {
            grid-area: warning;
            display: grid;
            gap: 0.32rem;
            padding: 0.78rem 0.82rem 0.84rem;
            border-radius: 16px;
            border: 1px solid rgba(166, 100, 32, 0.24);
            background:
                linear-gradient(135deg, rgba(255, 244, 220, 0.98) 0%, rgba(255, 233, 196, 0.96) 100%);
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.6);
        }

        .ocr-warning-kicker {
            font-size: 0.64rem;
            text-transform: uppercase;
            letter-spacing: 0.16em;
            font-weight: 800;
            color: var(--warn);
        }

        .ocr-warning-copy {
            font-size: 0.76rem;
            line-height: 1.42;
            color: #694a18;
            white-space: pre-wrap;
            overflow-wrap: anywhere;
        }

        .ocr-warning-badge {
            position: absolute;
            top: 0.72rem;
            right: 0.72rem;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 3.2rem;
            height: 1.8rem;
            padding: 0 0.55rem;
            border-radius: 999px;
            background: rgba(255, 244, 220, 0.98);
            border: 1px solid rgba(166, 100, 32, 0.26);
            box-shadow: 0 10px 20px rgba(110, 76, 20, 0.08);
            font-size: 0.62rem;
            line-height: 1;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            font-weight: 800;
            color: var(--warn);
        }

        .ocr-timing {
            grid-area: timing;
            justify-self: end;
            text-align: right;
            font-size: 0.88rem;
            color: var(--muted);
            white-space: nowrap;
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
            color: var(--signal-strong);
            animation: cooldownHandoff 420ms cubic-bezier(0.22, 1, 0.36, 1) both;
        }

        .ocr-preview-shell {
            grid-area: preview;
            display: grid;
            grid-template-columns: 6rem minmax(0, 1fr);
            gap: 0.9rem;
            align-items: center;
            padding-top: 0.3rem;
        }

        .ocr-preview {
            width: 6rem;
            aspect-ratio: 0.72;
            border-radius: 14px;
            overflow: hidden;
            border: 1px solid var(--line);
            background:
                linear-gradient(180deg, rgba(23, 50, 77, 0.12) 0%, rgba(201, 149, 50, 0.08) 100%),
                #ffffff;
            box-shadow: 0 10px 24px rgba(22, 32, 43, 0.08);
            background-position: center top;
            background-repeat: no-repeat;
            background-size: cover;
        }

        .ocr-row.just-completed .ocr-preview {
            animation: previewPulse 900ms ease;
        }

        .ocr-row.inactive .ocr-preview {
            background: #f6f0e5;
            border-color: #e0d3bf;
            box-shadow: none;
        }

        .ocr-preview.has-image {
            border-color: rgba(23, 50, 77, 0.14);
        }

        .ocr-row.inactive .ocr-preview.has-image {
            filter: grayscale(1) saturate(0.15) brightness(1.03);
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
            align-self: center;
            font-size: 0.77rem;
            color: var(--muted);
            line-height: 1.35;
        }

        .stFileUploader > div > div {
            border-radius: 18px !important;
            border: 1px dashed rgba(201, 149, 50, 0.58) !important;
            background: linear-gradient(135deg, rgba(255, 251, 244, 0.96) 0%, rgba(239, 229, 214, 0.96) 100%) !important;
            padding: 0.9rem !important;
            box-shadow: inset 0 0 0 1px rgba(23, 50, 77, 0.04) !important;
        }

        [data-testid="stDialog"] div[role="dialog"] {
            border: 1px solid var(--line) !important;
            border-radius: 24px !important;
            background:
                radial-gradient(circle at 0% 0%, rgba(201, 149, 50, 0.18), rgba(201, 149, 50, 0) 24rem),
                linear-gradient(180deg, rgba(255, 251, 244, 0.99) 0%, rgba(244, 235, 220, 0.99) 100%) !important;
            box-shadow: var(--shadow) !important;
        }

        [data-testid="stDialog"] div[role="dialog"] > div:first-child p {
            font-family: var(--display-font) !important;
            letter-spacing: -0.04em;
        }

        label[data-testid="stWidgetLabel"],
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        .stCaption,
        .stText,
        .stExpander {
            color: var(--ink) !important;
        }

        div[data-baseweb="input"],
        div[data-baseweb="base-input"],
        div[data-baseweb="select"] > div,
        div[data-baseweb="textarea"] {
            border-radius: 16px !important;
            border: 1px solid rgba(23, 50, 77, 0.14) !important;
            background: rgba(255, 251, 244, 0.9) !important;
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.8) !important;
            transition: border-color 180ms ease, box-shadow 180ms ease, background 180ms ease !important;
        }

        div[data-baseweb="input"]:focus-within,
        div[data-baseweb="base-input"]:focus-within,
        div[data-baseweb="select"]:focus-within > div,
        div[data-baseweb="textarea"]:focus-within {
            border-color: rgba(201, 149, 50, 0.88) !important;
            box-shadow: 0 0 0 3px rgba(201, 149, 50, 0.18) !important;
            background: #ffffff !important;
        }

        div[data-baseweb="input"] input,
        div[data-baseweb="base-input"] input,
        div[data-baseweb="select"] input,
        div[data-baseweb="textarea"] textarea {
            color: var(--ink) !important;
            background: transparent !important;
        }

        div[data-baseweb="select"] svg,
        [data-testid="stFileUploaderDropzoneInstructions"] svg {
            color: var(--signal-strong) !important;
        }

        .stButton > button,
        .stDownloadButton > button,
        .stFormSubmitButton > button {
            min-height: 3rem !important;
            border-radius: 999px !important;
            border: 1px solid rgba(201, 149, 50, 0.42) !important;
            font-weight: 700 !important;
            letter-spacing: -0.01em !important;
            box-shadow: 0 10px 20px rgba(15, 23, 32, 0.08) !important;
            background: rgba(255, 251, 244, 0.92) !important;
            color: var(--accent) !important;
            transition: background 180ms ease, border-color 180ms ease, color 180ms ease, transform 180ms ease, box-shadow 180ms ease !important;
        }

        .stButton > button[kind="primary"],
        .stFormSubmitButton > button {
            background: linear-gradient(145deg, var(--accent) 0%, #21486d 100%) !important;
            color: #fffaf1 !important;
            -webkit-text-fill-color: #fffaf1 !important;
            border-color: rgba(201, 149, 50, 0.64) !important;
            box-shadow:
                inset 0 1px 0 rgba(255, 255, 255, 0.08),
                0 16px 26px rgba(23, 50, 77, 0.22) !important;
        }

        .stButton > button[kind="primary"] *,
        .stFormSubmitButton > button *,
        .stButton > button[kind="primary"] [data-testid="stMarkdownContainer"] *,
        .stFormSubmitButton > button [data-testid="stMarkdownContainer"] * {
            color: #fffaf1 !important;
            -webkit-text-fill-color: #fffaf1 !important;
        }

        .stButton > button:hover,
        .stDownloadButton > button:hover,
        .stFormSubmitButton > button:hover {
            border-color: rgba(201, 149, 50, 0.72) !important;
            transform: translateY(-1px);
            box-shadow: 0 18px 28px rgba(15, 23, 32, 0.12) !important;
        }

        .stDownloadButton > button {
            background: rgba(255, 251, 244, 0.94) !important;
            color: var(--accent) !important;
        }

        .stButton > button[kind="primary"]:hover,
        .stFormSubmitButton > button:hover {
            background: var(--accent-strong) !important;
            border-color: rgba(201, 149, 50, 0.72) !important;
        }

        .stButton > button:focus-visible,
        .stDownloadButton > button:focus-visible,
        .stFormSubmitButton > button:focus-visible {
            outline: none !important;
            box-shadow: 0 0 0 3px rgba(201, 149, 50, 0.18) !important;
        }

        .stButton > button:disabled,
        .stDownloadButton > button:disabled,
        .stFormSubmitButton > button:disabled {
            background: #e6dccb !important;
            border-color: #e6dccb !important;
            color: #85796a !important;
            transform: none !important;
            cursor: not-allowed !important;
            box-shadow: none !important;
        }

        input[type="radio"] {
            accent-color: var(--signal);
        }

        .provider-rail {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.65rem;
            margin: 0.45rem 0 1rem;
        }

        .provider-card {
            position: relative;
            display: grid;
            grid-template-columns: 44px minmax(0, 1fr);
            gap: 0.7rem;
            align-items: center;
            padding: 0.7rem 0.8rem;
            border-radius: 16px;
            border: 1px solid rgba(23, 50, 77, 0.14);
            background: rgba(255, 251, 244, 0.96);
            box-shadow: 0 10px 18px rgba(15, 23, 32, 0.06);
        }

        .provider-card.is-selected {
            border-color: rgba(23, 50, 77, 0.28);
            background: linear-gradient(145deg, rgba(255, 251, 244, 0.98), rgba(244, 236, 222, 0.98));
            box-shadow: 0 14px 26px rgba(15, 23, 32, 0.08);
        }

        .provider-card.is-disabled {
            opacity: 0.48;
            background: rgba(239, 229, 214, 0.68);
            border-color: rgba(57, 73, 91, 0.14);
            box-shadow: none;
        }

        .provider-mark {
            position: relative;
            width: 44px;
            height: 44px;
            display: grid;
            place-items: center;
            border-radius: 14px;
            color: #13293f;
            background: linear-gradient(145deg, rgba(255, 255, 255, 0.82), rgba(229, 217, 197, 0.96));
            box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.7);
        }

        .provider-logo-image {
            width: 22px;
            height: 22px;
            display: block;
            object-fit: contain;
            border-radius: 6px;
        }

        .provider-logo-fallback {
            font-size: 0.9rem;
            font-weight: 700;
            color: var(--accent);
        }

        .provider-mark-openai {
            color: #0f1720;
        }

        .provider-mark-anthropic {
            color: #3c2916;
            background: linear-gradient(145deg, rgba(255, 247, 233, 0.92), rgba(235, 213, 185, 0.98));
        }

        .provider-card.is-disabled .provider-logo-image {
            filter: grayscale(1);
        }

        .provider-copy {
            min-width: 0;
            display: grid;
            gap: 0.16rem;
        }

        .provider-name-row {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 0.5rem;
        }

        .provider-name {
            font-size: 0.98rem;
            font-weight: 700;
            letter-spacing: -0.03em;
            text-transform: capitalize;
            color: var(--ink);
        }

        .provider-state {
            font-size: 0.66rem;
            font-weight: 700;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            color: var(--muted);
        }

        .provider-card.is-selected .provider-state {
            color: var(--signal-strong);
        }

        .provider-model-name {
            font-size: 0.8rem;
            line-height: 1.3;
            font-weight: 500;
            color: rgba(23, 50, 77, 0.8);
            word-break: break-word;
        }

        @media (prefers-reduced-motion: no-preference) {
            .provider-card {
                transition: box-shadow 180ms ease, border-color 180ms ease, background 180ms ease;
            }
        }

        @media (max-width: 860px) {
            .provider-rail {
                grid-template-columns: 1fr;
            }

            .provider-card {
                grid-template-columns: 40px minmax(0, 1fr);
                padding: 0.68rem 0.78rem;
            }

            .provider-mark {
                width: 40px;
                height: 40px;
                border-radius: 13px;
            }

            .provider-logo-image {
                width: 20px;
                height: 20px;
            }
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
                border-color: #d1ac60;
                box-shadow: 0 0 0 0 rgba(201, 149, 50, 0.14);
            }
            45% {
                border-color: #c99532;
                box-shadow: 0 0 0 8px rgba(201, 149, 50, 0.08);
            }
            100% {
                border-color: var(--line);
                box-shadow: none;
            }
        }

        @keyframes previewPulse {
            0% {
                transform: scale(0.96);
                border-color: #d1ac60;
            }
            55% {
                transform: scale(1.02);
                border-color: #8a6c2a;
            }
            100% {
                transform: scale(1);
                border-color: var(--line);
            }
        }

        @keyframes progressPulse {
            0%, 100% {
                transform: translateY(0) scaleY(1);
                filter: saturate(1) brightness(1);
                box-shadow:
                    inset 0 0 0 1px rgba(255, 255, 255, 0.28),
                    0 0 0 1px rgba(201, 149, 50, 0.22),
                    0 0 0 0 var(--progress-active-glow);
            }
            50% {
                transform: translateY(-1px) scaleY(1.14);
                filter: saturate(1.08) brightness(1.04);
                box-shadow:
                    inset 0 0 0 1px rgba(255, 255, 255, 0.34),
                    0 0 0 1px rgba(201, 149, 50, 0.28),
                    0 0 0 4px rgba(201, 149, 50, 0.16),
                    0 10px 20px rgba(23, 50, 77, 0.18);
            }
        }

        @keyframes progressSheen {
            0% {
                opacity: 0;
                transform: translateX(-120%);
            }
            28% {
                opacity: 0.7;
            }
            62% {
                opacity: 0.22;
                transform: translateX(120%);
            }
            100% {
                opacity: 0;
                transform: translateX(120%);
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

        @keyframes agentPulse {
            0%, 100% {
                transform: scale(1);
                box-shadow: 0 0 0 0 rgba(241, 193, 95, 0.18);
            }
            50% {
                transform: scale(1.08);
                box-shadow: 0 0 0 8px rgba(241, 193, 95, 0.08);
            }
        }

        @keyframes setupSweep {
            0% {
                transform: translateX(-58%) scaleX(0.72);
                opacity: 0.56;
            }
            50% {
                transform: translateX(42%) scaleX(1);
                opacity: 1;
            }
            100% {
                transform: translateX(152%) scaleX(0.7);
                opacity: 0.52;
            }
        }

        @keyframes setupMeter {
            0%, 100% {
                transform: translateX(-26%) scaleX(0.82);
            }
            50% {
                transform: translateX(34%) scaleX(1);
            }
        }

        @keyframes setupRadar {
            0% {
                transform: rotate(0deg);
            }
            100% {
                transform: rotate(360deg);
            }
        }

        @keyframes setupFloat {
            0%, 100% {
                transform: translateY(0);
            }
            50% {
                transform: translateY(-0.42rem);
            }
        }

        .section-card,
        .status-card,
        .stage-chip {
            animation: rise 240ms ease both;
        }

        .section-card:hover,
        .status-card:hover,
        .stage-chip:hover,
        .workbook-phase-card:hover,
        .agent-window:hover,
        .question-panel:hover,
        .ocr-row:hover {
            transform: translateY(-2px);
            box-shadow: 0 22px 42px rgba(15, 23, 32, 0.12);
        }

        .live-card {
            animation: none !important;
        }

        @media (prefers-reduced-motion: reduce) {
            .section-card,
            .status-card,
            .stage-chip,
            .ocr-live-countdown,
            .ocr-live-handoff,
            .agent-window-status-dot,
            .chunk-progress-block.active,
            .chunk-progress-block.active::after,
            .workbook-setup-stage-fill,
            .workbook-setup-radar-sweep,
            .workbook-setup-radar-core,
            .workbook-setup-step,
            .workbook-setup-step::before,
            .workbook-setup-step-fill,
            .stButton > button,
            .stDownloadButton > button {
                animation: none !important;
                transition: none !important;
            }
        }

        @media (max-width: 980px) {
            .taskbar-shell {
                top: 0.55rem;
            }

            .taskbar-bar {
                grid-template-columns: 1fr;
                align-items: stretch;
                padding: 0.78rem 0.82rem;
            }

            .taskbar-brand-row,
            .taskbar-status {
                flex-wrap: wrap;
            }

            .taskbar-strap,
            .taskbar-status-message {
                white-space: normal;
            }

            .taskbar-stage-rail {
                justify-content: flex-start;
            }

            .masthead {
                flex-direction: column;
                align-items: flex-start;
            }

            .immersive-workbook-shell {
                min-height: auto;
                margin: 0;
                padding: 0;
                background: transparent;
            }

            .immersive-workbook-grid,
            .workbook-setup-grid,
            .sheet-desk-summary,
            .sheet-desk-surface,
            .sheet-cell-grid,
            .workbook-setup-metrics {
                grid-template-columns: 1fr;
            }

            .immersive-workbook-grid {
                min-height: auto;
            }

            .workbook-setup-shell {
                min-height: auto;
                padding: 1rem;
            }

            .stage-rail {
                display: grid;
                grid-template-columns: 1fr;
            }

            .stage-chip,
            .stage-chip.active {
                flex: 1 1 auto;
            }

            .workbook-stage-body,
            .workbook-stage-header,
            .workbook-phase-grid,
            .question-grid,
            .agent-window-metrics {
                grid-template-columns: 1fr;
            }

            .sheet-desk-header,
            .sheet-preview-head,
            .sheet-queue-head {
                display: grid;
            }

            .question-shell.is-modal {
                grid-template-columns: 1fr;
            }

            .throttle-status-layout,
            .workbook-retry-banner {
                grid-template-columns: 1fr;
            }

            .throttle-status-timer-shell,
            .workbook-retry-timer-shell {
                justify-items: start;
            }

            .workbook-stage-header {
                display: grid;
            }

            .agent-window-head {
                display: grid;
            }

            .agent-window-head-meta {
                justify-items: start;
            }

            .workbook-setup-title {
                max-width: none;
                font-size: clamp(2.5rem, 13vw, 3.8rem);
            }

            .workbook-setup-stage {
                grid-template-columns: 1fr;
            }

            .workbook-setup-stage-indicator {
                justify-self: start;
                white-space: normal;
            }

            .workbook-setup-radar {
                min-height: 14rem;
            }

            .ocr-grid {
                grid-template-columns: 1fr;
            }

            .ocr-row {
                grid-template-columns: minmax(0, 1fr);
                grid-template-areas:
                    "pipe"
                    "timing"
                    "status"
                    "preview";
            }

            .ocr-timing {
                justify-self: start;
                text-align: left;
            }

            .ocr-preview-shell {
                grid-template-columns: 5rem minmax(0, 1fr);
            }
        }
        </style>
        """
    )


def init_state(settings: Settings) -> None:
    reset_run_scoped_widget_keys()
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
    st.session_state.setdefault("current_job_id", None)
    st.session_state.setdefault("workbook_retry", None)
    st.session_state.setdefault("agent_trace", build_agent_trace_state())
    st.session_state.setdefault(QUESTION_SUBMISSION_STATE_KEY, None)
    st.session_state.setdefault(QUESTION_PENDING_RESUME_STATE_KEY, None)
    st.session_state.setdefault(QUESTION_ACTIVE_RESUME_STATE_KEY, None)
    st.session_state.setdefault(QUESTION_WIDGET_STATE_KEY, {})
    runtime_defaults = build_runtime_settings_state(settings)
    runtime_settings = st.session_state.setdefault("runtime_settings", runtime_defaults)
    for key, value in runtime_defaults.items():
        runtime_settings.setdefault(key, value)


def reset_run_scoped_widget_keys() -> None:
    st.session_state["_render_key_counts"] = {}


def run_scoped_widget_key(base_key: str) -> str:
    key_counts = st.session_state.setdefault("_render_key_counts", {})
    count = int(key_counts.get(base_key, 0)) + 1
    key_counts[base_key] = count
    return base_key if count == 1 else f"{base_key}__{count}"


def build_ocr_pipeline_state(pipe_total: int = DEFAULT_OCR_PARALLEL_WORKERS) -> dict:
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


def build_workbook_retry_state() -> dict:
    return {
        "sheet_name": None,
        "message": None,
        "detail_message": None,
        "phase": None,
        "severity": None,
        "attempt_number": None,
        "max_attempts": None,
        "retry_delay_seconds": None,
        "retry_until_ms": None,
        "retry_reason": None,
    }


def build_agent_trace_state() -> dict:
    return {
        "status": "idle",
        "current_sheet": None,
        "message": "This panel will show status updates once workbook entry starts.",
        "token_count": 0,
        "started_at_ms": None,
        "retry_until_ms": None,
        "live_summary": None,
        "live_summary_updated_at_ms": None,
        "last_rendered_summary": None,
        "recent_events": [],
    }


def agent_trace_status_label(status: str) -> str:
    labels = {
        "idle": "Standby",
        "running": "In progress",
        "retrying": "Paused",
        "needs_input": "Needs your input",
        "validating": "Final check",
        "review": "Needs review",
        "complete": "Complete",
        "error": "Attention needed",
    }
    return labels.get(status, "Standby")


def push_agent_trace_event(*, label: str, message: str, tone: str) -> None:
    trace = st.session_state["agent_trace"]
    trace["recent_events"] = [
        *trace.get("recent_events", []),
        {
            "label": label,
            "message": message,
            "tone": tone,
        },
    ][-4:]


def format_token_count(token_count: int) -> str:
    return f"{max(0, token_count):,}"


def clear_agent_trace_live_summary(trace: dict[str, object]) -> None:
    trace["live_summary"] = None
    trace["live_summary_updated_at_ms"] = None
    trace["last_rendered_summary"] = None


def mark_agent_trace_summary_rendered() -> None:
    trace = st.session_state.get("agent_trace")
    if not isinstance(trace, dict):
        return
    trace["last_rendered_summary"] = trace.get("live_summary")


def summarize_agent_trace_event(event: RunEvent) -> tuple[str, str, str] | None:
    if event.phase == "heartbeat":
        return None

    if event.stage == Stage.DATA_ENTRY:
        if event.phase == "start":
            return (
                "Now working",
                event.detail_message or event.message,
                "info",
            )
        if event.phase == "retry":
            return ("Paused", event.detail_message or event.message, "warn")
        if event.phase == "paused":
            pending_question = event.artifacts.pending_question if event.artifacts is not None else None
            return (
                "Needs your input",
                pending_question.prompt if pending_question is not None else event.message,
                "warn",
            )
        if event.phase == "failed":
            return ("Entry stopped", event.detail_message or event.message, "error")
        if event.phase == "complete" and event.message == "Workbook entry complete.":
            return ("Next step", "Section entry is complete. Final review is now running.", "success")
        if event.sheet_name and event.message.startswith("Completed sheet "):
            return ("Section saved", f"{event.sheet_name} has been written into the workbook.", "success")
        return ("Workbook entry", event.message, "info")

    if event.stage == Stage.FINANCIAL_CALCULATIONS:
        tone = "error" if event.severity == Severity.ERROR else "warn" if event.severity == Severity.WARNING else "success"
        return ("Final review", event.message, tone)

    return None


def sync_agent_trace_from_event(event: RunEvent) -> None:
    if event.stage not in {Stage.DATA_ENTRY, Stage.FINANCIAL_CALCULATIONS}:
        return

    trace = st.session_state["agent_trace"]
    if event.agent_total_tokens is not None:
        trace["token_count"] = int(event.agent_total_tokens)

    progress_message = str(event.progress_message or "").strip()
    if progress_message and progress_message != str(trace.get("live_summary") or ""):
        trace["live_summary"] = progress_message
        trace["live_summary_updated_at_ms"] = int(time.time() * 1000)

    if event.phase == "heartbeat":
        return

    now_ms = int(time.time() * 1000)
    clear_agent_trace_live_summary(trace)
    summary = summarize_agent_trace_event(event)
    if summary is not None:
        label, message, tone = summary
        push_agent_trace_event(label=label, message=message, tone=tone)

    if event.stage == Stage.DATA_ENTRY:
        current_sheet = event.sheet_name or trace.get("current_sheet")
        if event.phase == "start":
            trace.update(
                {
                    "status": "running",
                    "current_sheet": current_sheet,
                    "message": event.detail_message
                    or "Reviewing the source material before adding workbook entries.",
                    "started_at_ms": now_ms,
                    "retry_until_ms": None,
                }
            )
            return

        if event.phase == "retry":
            trace.update(
                {
                    "status": "retrying",
                    "current_sheet": current_sheet,
                    "message": event.detail_message or event.message,
                    "started_at_ms": trace.get("started_at_ms") or now_ms,
                    "retry_until_ms": (
                        now_ms + int(event.retry_delay_seconds * 1000)
                        if event.retry_delay_seconds is not None
                        else None
                    ),
                }
            )
            return

        if event.phase == "paused":
            pending_question = event.artifacts.pending_question if event.artifacts is not None else None
            trace.update(
                {
                    "status": "needs_input",
                    "current_sheet": (
                        pending_question.sheet_name if pending_question is not None else current_sheet
                    ),
                    "message": (
                        pending_question.prompt
                        if pending_question is not None
                        else "Workbook entry is waiting for planner input."
                    ),
                    "started_at_ms": None,
                    "retry_until_ms": None,
                }
            )
            return

        if event.phase == "failed":
            trace.update(
                {
                    "status": "error",
                    "current_sheet": current_sheet,
                    "message": event.detail_message or event.message,
                    "started_at_ms": None,
                    "retry_until_ms": None,
                }
            )
            return

        if event.phase == "complete" and event.message == "Workbook entry complete.":
            trace.update(
                {
                    "status": "validating",
                    "current_sheet": None,
                    "message": "Workbook entry is complete. Final review is running now.",
                    "started_at_ms": now_ms,
                    "retry_until_ms": None,
                }
            )
            return

        if event.message.startswith("Completed sheet "):
            trace.update(
                {
                    "status": "running",
                    "current_sheet": current_sheet,
                    "message": "That section is saved. Preparing the next one now.",
                    "started_at_ms": None,
                    "retry_until_ms": None,
                }
            )
            return

        trace.update(
            {
                "status": "running",
                "current_sheet": current_sheet,
                "message": event.message,
                "started_at_ms": now_ms if trace.get("started_at_ms") is None else trace.get("started_at_ms"),
                "retry_until_ms": None,
            }
        )
        return

    if event.stage == Stage.FINANCIAL_CALCULATIONS:
        if event.phase == "complete":
            final_status = (
                "error"
                if event.severity == Severity.ERROR
                else "review" if event.severity == Severity.WARNING else "complete"
            )
            trace.update(
                {
                    "status": final_status,
                    "current_sheet": None,
                    "message": event.message,
                    "started_at_ms": None,
                    "retry_until_ms": None,
                }
            )
            return

        trace.update(
            {
                "status": "validating",
                "current_sheet": None,
                "message": event.message,
                "started_at_ms": now_ms if trace.get("status") != "validating" else trace.get("started_at_ms"),
                "retry_until_ms": None,
            }
        )


def build_agent_trace_timer_markup(trace: dict[str, object]) -> tuple[str, str]:
    status = str(trace.get("status") or "idle")
    now_ms = int(time.time() * 1000)
    retry_until_ms = trace.get("retry_until_ms")
    started_at_ms = trace.get("started_at_ms")

    if status == "retrying" and retry_until_ms is not None:
        remaining_ms = max(int(retry_until_ms) - now_ms, 0)
        return (
            "Retry in",
            build_timing_markup(
                format_seconds_label(remaining_ms / 1000),
                countdown_target_ms=int(retry_until_ms),
            ),
        )

    if status in {"running", "validating"} and started_at_ms is not None:
        elapsed_seconds = max((now_ms - int(started_at_ms)) / 1000, 0.0)
        return (
            "Elapsed",
            build_timing_markup(
                format_seconds_label(elapsed_seconds),
                elapsed_started_at_ms=int(started_at_ms),
            ),
        )

    return ("State", html.escape(agent_trace_status_label(status)))


def build_agent_trace_headline(trace: dict[str, object]) -> str:
    status = str(trace.get("status") or "idle")
    current_sheet = trace.get("current_sheet")
    if status == "running" and current_sheet:
        return f"Reviewing {current_sheet}"
    if status == "retrying" and current_sheet:
        return f"Paused on {current_sheet}"
    if status == "needs_input" and current_sheet:
        return f"Waiting for your answer on {current_sheet}"
    if status == "validating":
        return "Checking the completed workbook"
    if status == "review":
        return "Workbook needs review"
    if status == "complete":
        return "Workbook ready"
    if status == "error" and current_sheet:
        return f"{current_sheet} needs attention"
    if status == "error":
        return "Run needs attention"
    return "Waiting to start"


def active_entry_question_resume_job_id() -> str | None:
    job_id = st.session_state.get(QUESTION_ACTIVE_RESUME_STATE_KEY)
    if not isinstance(job_id, str) or not job_id.strip():
        return None
    return job_id


def question_resume_in_flight_for(job_id: str | None) -> bool:
    return isinstance(job_id, str) and job_id == active_entry_question_resume_job_id()


def result_waiting_on_question(result: ImportArtifacts | None) -> bool:
    return bool(
        result is not None
        and getattr(result, "pending_question", None) is not None
        and getattr(result, "review_report", None) is None
        and not question_resume_in_flight_for(getattr(result, "job_id", None))
    )


def sanitize_entry_state_for_active_resume(state: EntrySessionState) -> EntrySessionState:
    if not question_resume_in_flight_for(state.job_id):
        return state

    blocked_sheet = state.pending_question.sheet_name if state.pending_question is not None else None
    sanitized_state = state.model_copy(deep=True)
    sanitized_state.pending_question = None
    if blocked_sheet is None:
        return sanitized_state

    sanitized_state.sheet_summaries = [
        summary.model_copy(update={"status": "completed", "message": None})
        if summary.sheet_name == blocked_sheet and summary.status == "needs_input"
        else summary
        for summary in sanitized_state.sheet_summaries
    ]
    return sanitized_state


def load_live_entry_state() -> EntrySessionState | None:
    result = st.session_state.get("result")
    entry_state_path = getattr(result, "entry_state_path", None)
    if entry_state_path is None or not entry_state_path.exists():
        return None
    try:
        return sanitize_entry_state_for_active_resume(load_entry_state(entry_state_path))
    except Exception:
        return None


def determine_focus_sheet(entry_state: EntrySessionState | None, active_stage: Stage) -> str | None:
    trace = st.session_state.get("agent_trace", {})
    current_sheet = trace.get("current_sheet")
    if isinstance(current_sheet, str) and current_sheet.strip():
        return current_sheet

    if entry_state is not None and entry_state.pending_question is not None:
        return entry_state.pending_question.sheet_name

    if entry_state is None or not entry_state.sheet_order:
        return None

    if active_stage == Stage.DATA_ENTRY and entry_state.current_sheet_index < len(entry_state.sheet_order):
        return entry_state.sheet_order[entry_state.current_sheet_index]

    focus_index = min(max(entry_state.current_sheet_index - 1, 0), len(entry_state.sheet_order) - 1)
    return entry_state.sheet_order[focus_index]


def format_sheet_value(value: object) -> str:
    if value in (None, ""):
        return "Waiting for value"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def describe_assignment_title(assignment: object) -> str:
    semantic_key = str(getattr(assignment, "semantic_key", "") or "").strip()
    cell_ref = str(getattr(assignment, "cell", "") or "").strip()
    if not semantic_key:
        return f"Workbook field {cell_ref}" if cell_ref else "Workbook field"

    parts = [part for part in semantic_key.split(".") if part]
    title = parts[-1].replace("_", " ").strip().title() if parts else semantic_key
    if len(parts) >= 2 and parts[-2].startswith("client_"):
        owner = parts[-2].replace("_", " ").strip().title()
        return f"{title} ({owner})"
    return title


def describe_assignment_detail(assignment: object) -> str:
    parts = [f"Workbook cell {getattr(assignment, 'cell', '')}"]
    source_pages = list(getattr(assignment, "source_pages", []) or [])
    if source_pages:
        page_label = "Source page" if len(source_pages) == 1 else "Source pages"
        parts.append(f"{page_label} {', '.join(str(page) for page in source_pages)}")
    return " • ".join(part for part in parts if part.strip())


def build_sheet_queue_markup(
    *,
    entry_state: EntrySessionState | None,
    focus_sheet: str | None,
    active_stage: Stage,
) -> str:
    if entry_state is None or not entry_state.sheet_order:
        sheet_order = list(TEMPLATE_SHEET_ORDER)
        current_sheet_index = 0
        pending_sheet = None
        summary_map: dict[str, object] = {}
    else:
        sheet_order = entry_state.sheet_order
        current_sheet_index = entry_state.current_sheet_index
        pending_sheet = entry_state.pending_question.sheet_name if entry_state.pending_question is not None else None
        summary_map = {summary.sheet_name: summary for summary in entry_state.sheet_summaries}

    items: list[tuple[int, str, str]] = []
    for index, sheet_name in enumerate(sheet_order):
        summary = summary_map.get(sheet_name)
        status_label = "Up next"
        tone = "pending"
        if getattr(summary, "status", None) == "skipped":
            status_label = "No changes"
            tone = "skipped"
        elif pending_sheet == sheet_name or getattr(summary, "status", None) == "needs_input":
            status_label = "Needs your input"
            tone = "needs-input"
        elif active_stage == Stage.FINANCIAL_CALCULATIONS or index < current_sheet_index:
            status_label = "Done"
            tone = "complete"
        elif focus_sheet == sheet_name and active_stage == Stage.DATA_ENTRY:
            status_label = "Now"
            tone = "active"

        mapped_count = int(getattr(summary, "mapped_count", 0) or 0)
        unresolved_count = int(getattr(summary, "unresolved_count", 0) or 0)
        touched_cells = len(getattr(summary, "touched_cells", []) or [])
        if tone == "complete":
            meta = (
                f"{mapped_count} entries saved"
                if mapped_count > 0
                else f"{touched_cells} cells updated"
                if touched_cells > 0
                else "This section is complete."
            )
        elif tone == "needs-input":
            meta = getattr(summary, "message", None) or "We need one answer before this section can continue."
        elif tone == "active":
            meta = getattr(summary, "message", None) or (
                f"{mapped_count} entries saved so far"
                if mapped_count > 0
                else "This section is being filled now."
            )
        elif tone == "skipped":
            meta = getattr(summary, "message", None) or "No updates were needed for this section."
        else:
            meta = (
                f"{unresolved_count} items still need support"
                if unresolved_count > 0
                else "This section starts after the current one."
            )

        items.append(
            (
                index,
                tone,
                dedent(
                    f"""
                    <div class="sheet-queue-item {html.escape(tone)}">
                        <div class="sheet-queue-name">{html.escape(sheet_name)}</div>
                        <div class="sheet-queue-state">{html.escape(status_label)}</div>
                        <div class="sheet-queue-meta">{html.escape(meta)}</div>
                    </div>
                    """
                ).strip(),
            )
        )

    return "".join(markup for _, _, markup in items)


def build_sheet_cells_markup(
    *,
    entry_state: EntrySessionState | None,
    focus_sheet: str | None,
    active_stage: Stage,
) -> tuple[str, str, str]:
    if focus_sheet is None:
        return (
            "Final review",
            "The completed workbook is being checked before release.",
            (
                '<div class="sheet-cell is-empty sheet-cell-empty-state">'
                '<div class="sheet-cell-title">Final review is in progress</div>'
                '<div class="sheet-cell-detail">The completed workbook is being checked before it is released.</div>'
                "</div>"
            ),
        )

    summary = None
    assignments_by_cell: dict[str, object] = {}
    if entry_state is not None:
        summary = next((item for item in entry_state.sheet_summaries if item.sheet_name == focus_sheet), None)
        assignments_by_cell = {
            assignment.cell: assignment
            for assignment in entry_state.mapped_assignments
            if assignment.sheet_name == focus_sheet
        }

    visible_cells = list(getattr(summary, "touched_cells", []) or [])
    if not visible_cells:
        visible_cells = sorted(ALLOWED_WRITE_CELLS_BY_SHEET.get(focus_sheet, set()))[:12]
    else:
        visible_cells = visible_cells[:12]

    if not visible_cells:
        empty_copy = (
            "This area becomes active when the runner reaches a workbook section with editable fields."
            if active_stage == Stage.DATA_ENTRY
            else "Final review is running on the completed workbook."
        )
        return (
            focus_sheet,
            getattr(summary, "message", None) or empty_copy,
            (
                '<div class="sheet-cell is-empty sheet-cell-empty-state">'
                f'<div class="sheet-cell-title">{html.escape("Entries will appear here" if active_stage == Stage.DATA_ENTRY else "Final review is in progress")}</div>'
                f'<div class="sheet-cell-detail">{html.escape(empty_copy)}</div>'
                "</div>"
            ),
        )

    filled_cells = [cell_ref for cell_ref in visible_cells if cell_ref in assignments_by_cell][:8]
    if not filled_cells:
        empty_copy = getattr(summary, "message", None) or (
            "This section is being reviewed. Saved entries will appear here as soon as they are ready."
            if active_stage == Stage.DATA_ENTRY
            else "The completed workbook is being checked before release."
        )
        empty_title = "Entries will appear here" if active_stage == Stage.DATA_ENTRY else "Final review is in progress"
        return (
            focus_sheet,
            empty_copy,
            dedent(
                f"""
                <div class="sheet-cell is-empty sheet-cell-empty-state">
                    <div class="sheet-cell-title">{html.escape(empty_title)}</div>
                    <div class="sheet-cell-detail">{html.escape(empty_copy)}</div>
                </div>
                """
            ).strip(),
        )

    cells: list[str] = []
    for cell_ref in filled_cells:
        assignment = assignments_by_cell[cell_ref]
        cells.append(
            dedent(
                f"""
                <div class="sheet-cell is-filled">
                    <div class="sheet-cell-label">Cell {html.escape(cell_ref)}</div>
                    <div class="sheet-cell-title">{html.escape(describe_assignment_title(assignment))}</div>
                    <div class="sheet-cell-value">{html.escape(format_sheet_value(getattr(assignment, "value", None)))}</div>
                    <div class="sheet-cell-detail">{html.escape(describe_assignment_detail(assignment))}</div>
                </div>
                """
            ).strip()
        )

    copy = getattr(summary, "message", None) or "Latest saved entries for this section appear here."
    return focus_sheet, copy, "".join(cells)


def build_sheet_desk_markup(
    *,
    entry_state: EntrySessionState | None,
    active_stage: Stage,
    source_filename: str,
    current_phase: str,
    workbook_message: str,
    mapping_completed: int,
    mapping_total: int,
    checks_completed: int,
    checks_total: int,
    mapping_state: str,
    mapping_copy: str,
    checks_state: str,
    checks_copy: str,
) -> str:
    focus_sheet = determine_focus_sheet(entry_state, active_stage)
    desk_title = focus_sheet if focus_sheet is not None else "Workbook review"
    queue_markup = build_sheet_queue_markup(
        entry_state=entry_state,
        focus_sheet=focus_sheet,
        active_stage=active_stage,
    )
    preview_title, preview_copy, cell_markup = build_sheet_cells_markup(
        entry_state=entry_state,
        focus_sheet=focus_sheet,
        active_stage=active_stage,
    )
    progress_total = mapping_total if active_stage == Stage.DATA_ENTRY else checks_total
    progress_completed = mapping_completed if active_stage == Stage.DATA_ENTRY else checks_completed
    progress_ratio = min(max(progress_completed / progress_total, 0), 1) if progress_total > 0 else 0.0
    progress_title = (
        f"{mapping_completed} of {mapping_total} sections complete"
        if active_stage == Stage.DATA_ENTRY
        else f"{checks_completed} of {checks_total} final checks complete"
    )
    progress_copy = mapping_copy if active_stage == Stage.DATA_ENTRY else checks_copy
    now_title = current_phase
    now_copy = mapping_copy if active_stage == Stage.DATA_ENTRY else checks_copy
    next_title = "Final review" if active_stage == Stage.DATA_ENTRY else "Release workbook"
    next_copy = (
        checks_copy
        if active_stage == Stage.DATA_ENTRY
        else "The workbook will be ready as soon as the final review finishes."
    )
    next_state = checks_state if active_stage == Stage.DATA_ENTRY else "In progress"

    return dedent(
        f"""
        <section class="sheet-desk">
            <div class="sheet-desk-header">
                <div>
                    <div class="sheet-desk-kicker">Workbook</div>
                    <div class="sheet-desk-title">{html.escape(desk_title)}</div>
                    <div class="sheet-desk-copy">{html.escape(workbook_message)}</div>
                </div>
                <div class="workbook-stage-pill">{html.escape(current_phase)}</div>
            </div>
            <div class="sheet-desk-copy">Source file: {html.escape(source_filename)}</div>
            <div class="sheet-desk-summary">
                <div class="sheet-desk-summary-card is-primary">
                    <div class="sheet-desk-summary-label">Now</div>
                    <div class="sheet-desk-summary-title">{html.escape(now_title)}</div>
                    <div class="sheet-desk-summary-copy">{html.escape(now_copy)} ({html.escape(mapping_state if active_stage == Stage.DATA_ENTRY else checks_state)})</div>
                </div>
                <div class="sheet-desk-summary-card">
                    <div class="sheet-desk-summary-label">Progress</div>
                    <div class="sheet-desk-summary-title">{html.escape(progress_title)}</div>
                    <div class="sheet-desk-summary-meter"><span style="width: {progress_ratio * 100:.1f}%"></span></div>
                    <div class="sheet-desk-summary-copy">{html.escape(next_title)}: {html.escape(next_copy)} ({html.escape(next_state)})</div>
                </div>
            </div>
            <div class="sheet-desk-surface">
                <section class="sheet-preview-card">
                    <div class="sheet-preview-head">
                        <div>
                            <div class="sheet-preview-kicker">Current section</div>
                            <div class="sheet-preview-title">{html.escape(preview_title)}</div>
                        </div>
                        <div class="sheet-preview-pill">{html.escape(current_phase)}</div>
                    </div>
                    <div class="sheet-preview-copy">{html.escape(preview_copy)}</div>
                    <div class="sheet-cell-grid">{cell_markup}</div>
                </section>
                <aside class="sheet-queue-card">
                    <div class="sheet-queue-head">
                        <div>
                            <div class="sheet-queue-kicker">Roadmap</div>
                            <div class="sheet-queue-title">Coming up</div>
                        </div>
                    </div>
                    <div class="sheet-queue-copy">The remaining sections stay visible here so it is clear what is done and what is next.</div>
                    <div class="sheet-queue-list">{queue_markup}</div>
                </aside>
            </div>
        </section>
        """
    ).strip()


def should_render_workbook_setup_shell(
    *,
    active_stage: Stage,
    is_running: bool,
    trace: dict[str, object],
    workbook_retry: object | None,
) -> bool:
    if active_stage != Stage.DATA_ENTRY or not is_running or workbook_retry is not None:
        return False

    status = str(trace.get("status") or "idle")
    current_sheet = str(trace.get("current_sheet") or "").strip()
    recent_events = list(trace.get("recent_events", []))
    started_at_ms = trace.get("started_at_ms")
    token_count = int(trace.get("token_count") or 0)
    live_summary = str(trace.get("live_summary") or "").strip()
    return (
        status == "idle"
        and not current_sheet
        and not recent_events
        and started_at_ms is None
        and token_count == 0
        and not live_summary
    )


def build_workbook_setup_markup(
    *,
    source_filename: str,
    status_message: str,
    mapping_total: int,
    checks_total: int,
) -> str:
    runtime_settings = dict(st.session_state.get("runtime_settings", {}))
    lane_count = max(int(runtime_settings.get("ocr_parallel_workers") or DEFAULT_OCR_PARALLEL_WORKERS), 1)
    retry_budget = max(int(runtime_settings.get("llm_max_retries") or 0), 0)
    sheet_total = max(int(mapping_total or 0), len(TEMPLATE_SHEET_ORDER))
    checks_total_safe = max(int(checks_total or 0), 1)
    setup_copy = status_message.strip() or "Preparing the workbook and waiting for the first section to begin."

    chip_specs = (
        ("File", source_filename),
        ("Review lanes", f"{lane_count} parallel"),
        ("Sections", f"{sheet_total} sections"),
        ("Final checks", f"{checks_total_safe} queued"),
    )
    chip_markup = "".join(
        dedent(
            f"""
            <div class="workbook-setup-chip">
                <span>{html.escape(label)}</span>
                <strong>{html.escape(value)}</strong>
            </div>
            """
        ).strip()
        for label, value in chip_specs
    )

    step_specs = (
        (
            "Step 1",
            "Check the workbook format",
            "Confirming the workbook template before any values are added.",
        ),
        (
            "Step 2",
            "Prepare the first section",
            f"Setting up {sheet_total} sections so entry can begin in the right place.",
        ),
        (
            "Step 3",
            "Queue the final review",
            f"{checks_total_safe} workbook checks will run automatically after section entry is finished.",
        ),
        (
            "Step 4",
            "Wait for the first update",
            "The live workbook view will open as soon as the first section starts.",
        ),
    )
    step_markup = "".join(
        dedent(
            f"""
            <div class="workbook-setup-step" style="--setup-delay: {index * 0.16:.2f}s;">
                <div class="workbook-setup-step-kicker">{html.escape(kicker)}</div>
                <div class="workbook-setup-step-title">{html.escape(title)}</div>
                <div class="workbook-setup-step-copy">{html.escape(copy)}</div>
                <div class="workbook-setup-step-meter"><span class="workbook-setup-step-fill"></span></div>
            </div>
            """
        ).strip()
        for index, (kicker, title, copy) in enumerate(step_specs)
    )

    metric_specs = (
        ("Sections", f"{sheet_total} total"),
        ("Final checks", f"{checks_total_safe} queued"),
        ("Auto-retries", f"{retry_budget} allowed"),
        ("Review lanes", f"{lane_count} active"),
    )
    metric_markup = "".join(
        dedent(
            f"""
            <div class="workbook-setup-metric">
                <span>{html.escape(label)}</span>
                <strong>{html.escape(value)}</strong>
            </div>
            """
        ).strip()
        for label, value in metric_specs
    )

    return dedent(
        f"""
        <section class="workbook-setup-shell">
            <div class="workbook-setup-grid">
                <div class="workbook-setup-hero">
                    <div>
                        <div class="workbook-setup-eyebrow">Preparing workbook</div>
                        <div class="workbook-setup-title">Getting the workbook ready</div>
                        <div class="workbook-setup-copy">{html.escape(setup_copy)}</div>
                    </div>
                    <div class="workbook-setup-chip-row">{chip_markup}</div>
                    <div>
                        <div class="workbook-setup-stage">
                            <div>
                                <div class="workbook-setup-stage-kicker">Current status</div>
                                <div class="workbook-setup-stage-title">Waiting for the first workbook update</div>
                                <div class="workbook-setup-stage-copy">The live workbook view opens as soon as the first section starts.</div>
                            </div>
                            <div class="workbook-setup-stage-indicator">Ready to start</div>
                        </div>
                        <div class="workbook-setup-stage-meter"><span class="workbook-setup-stage-fill"></span></div>
                    </div>
                    <div class="workbook-setup-radar" aria-hidden="true">
                        <div class="workbook-setup-radar-sweep"></div>
                        <div class="workbook-setup-radar-ring ring-a"></div>
                        <div class="workbook-setup-radar-ring ring-b"></div>
                        <div class="workbook-setup-radar-ring ring-c"></div>
                        <div class="workbook-setup-radar-core">
                            <span>First live update</span>
                            <strong>Ready for section entry</strong>
                        </div>
                        <div class="workbook-setup-radar-node node-a"></div>
                        <div class="workbook-setup-radar-node node-b"></div>
                        <div class="workbook-setup-radar-node node-c"></div>
                        <div class="workbook-setup-radar-node node-d"></div>
                    </div>
                </div>
                <aside class="workbook-setup-console">
                    <div>
                        <div class="workbook-setup-console-kicker">What happens first</div>
                        <div class="workbook-setup-console-title">Preparing the live workbook view</div>
                        <div class="workbook-setup-console-copy">This stays in setup mode until the first section begins, then the live progress board takes over.</div>
                    </div>
                    <div class="workbook-setup-step-list">{step_markup}</div>
                    <div class="workbook-setup-metrics">{metric_markup}</div>
                </aside>
            </div>
        </section>
        """
    ).strip()


def build_agent_trace_markup(
    *,
    active_stage: Stage,
    mapping_completed: int,
    mapping_total: int,
    checks_completed: int,
    checks_total: int,
) -> str:
    trace = st.session_state.get("agent_trace", build_agent_trace_state())
    status = str(trace.get("status") or "idle")
    timer_label, timer_markup = build_agent_trace_timer_markup(trace)
    status_label = agent_trace_status_label(status)
    scope_label = str(trace.get("current_sheet") or display_workbook_phase_name(active_stage))
    live_summary = str(trace.get("live_summary") or "").strip()
    progress_label = (
        f"{checks_completed}/{checks_total} final checks"
        if active_stage == Stage.FINANCIAL_CALCULATIONS
        else f"{mapping_completed}/{mapping_total} sections"
    )
    if live_summary:
        summary_markup = dedent(
            f"""
            <div class="agent-window-summary">
                <div class="agent-window-summary-kicker">Reasoning summary</div>
                <div class="agent-window-summary-copy">{html.escape(live_summary)}</div>
            </div>
            """
        ).strip()
    else:
        summary_markup = (
            '<div class="agent-window-summary empty">'
            '<div class="agent-window-summary-kicker">Reasoning summary</div>'
            '<div class="agent-window-summary-copy">OpenAI summaries will appear here while workbook entry is in progress.</div>'
            "</div>"
        )
    recent_events = list(trace.get("recent_events", []))
    if recent_events:
        event_markup = "".join(
            dedent(
                f"""
                <div class="agent-window-line {html.escape(str(item.get('tone') or 'info'))}">
                    <div class="agent-window-line-kicker">{html.escape(str(item.get("label") or "Update"))}</div>
                    <div class="agent-window-line-copy">{html.escape(str(item.get("message") or ""))}</div>
                </div>
                """
            ).strip()
            for item in reversed(recent_events)
        )
    else:
        event_markup = (
            '<div class="agent-window-line empty">'
            '<div class="agent-window-line-kicker">Standby</div>'
            '<div class="agent-window-line-copy">Status updates will appear here once workbook entry begins.</div>'
            "</div>"
        )

    return dedent(
        f"""
        <aside class="agent-window">
            <div class="agent-window-head">
                <div>
                    <div class="agent-window-kicker">Status updates</div>
                    <div class="agent-window-title">{html.escape(build_agent_trace_headline(trace))}</div>
                </div>
                <div class="agent-window-head-meta">
                    <div class="agent-window-status {html.escape(status)}">
                        <span class="agent-window-status-dot"></span>
                        {html.escape(status_label)}
                    </div>
                </div>
            </div>
            <div class="agent-window-metrics">
                <div class="agent-window-metric">
                    <span>Current section</span>
                    <strong>{html.escape(scope_label)}</strong>
                </div>
                <div class="agent-window-metric">
                    <span>Progress</span>
                    <strong>{html.escape(progress_label)}</strong>
                </div>
                <div class="agent-window-metric">
                    <span>{html.escape(timer_label)}</span>
                    <strong>{timer_markup}</strong>
                </div>
            </div>
            <div class="agent-window-copy">{html.escape(str(trace.get("message") or ""))}</div>
            {summary_markup}
            <div class="agent-window-timeline">{event_markup}</div>
            <div class="agent-window-footer">Recent updates and any blockers appear here.</div>
        </aside>
        """
    ).strip()


def build_runtime_settings_state(settings: Settings) -> dict:
    return {
        "llm_provider": settings.normalized_llm_provider(),
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
    llm_provider = DEFAULT_LLM_PROVIDER
    runtime_settings["llm_provider"] = llm_provider
    locked_model = locked_model_for_provider(llm_provider)
    return replace(
        base_settings,
        llm_provider=llm_provider,
        model_ocr=locked_model,
        model_mapping=locked_model,
        ocr_parallel_workers=int(runtime_settings["ocr_parallel_workers"]),
        llm_timeout_seconds=float(runtime_settings["llm_timeout_seconds"]),
        llm_max_retries=int(runtime_settings["llm_max_retries"]),
        llm_retry_base_seconds=float(runtime_settings["llm_retry_base_seconds"]),
        llm_retry_max_seconds=float(runtime_settings["llm_retry_max_seconds"]),
        max_pages=int(runtime_settings["max_pages"]),
        log_level=str(runtime_settings["log_level"]),
    )


def current_provider_display_name() -> str:
    return provider_display_name(DEFAULT_LLM_PROVIDER)


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
    st.session_state["last_status"] = build_status_snapshot(
        Stage.OCR,
        message="Job queued. Preparing document review.",
        severity=Severity.INFO,
    )
    st.session_state["ocr_pipeline"] = build_ocr_pipeline_state(pipe_total)
    st.session_state["run_started"] = True
    st.session_state["is_running"] = True
    st.session_state["source_filename"] = source_filename
    st.session_state["page_previews"] = page_previews
    st.session_state["current_job_id"] = None
    st.session_state["workbook_retry"] = None
    st.session_state["agent_trace"] = build_agent_trace_state()
    st.session_state[QUESTION_SUBMISSION_STATE_KEY] = None
    st.session_state[QUESTION_PENDING_RESUME_STATE_KEY] = None
    st.session_state[QUESTION_ACTIVE_RESUME_STATE_KEY] = None
    st.session_state[QUESTION_WIDGET_STATE_KEY] = {}


WORKFLOW_STAGES: tuple[Stage, ...] = (Stage.OCR, Stage.DATA_ENTRY)


def workflow_stage(stage: Stage | str) -> Stage:
    stage_enum = stage if isinstance(stage, Stage) else Stage(stage)
    return Stage.OCR if stage_enum == Stage.OCR else Stage.DATA_ENTRY


def stage_index(stage: Stage) -> int:
    return WORKFLOW_STAGES.index(workflow_stage(stage)) + 1


def display_stage_name(stage: Stage | str, *, compact: bool = False) -> str:
    stage_enum = stage if isinstance(stage, Stage) else Stage(stage)
    names = {
        Stage.OCR: ("Document review", "Review"),
        Stage.DATA_ENTRY: ("Workbook entry", "Entry"),
        Stage.FINANCIAL_CALCULATIONS: ("Workbook entry", "Entry"),
    }
    full_name, compact_name = names[stage_enum]
    return compact_name if compact else full_name


def display_workbook_phase_name(stage: Stage | str) -> str:
    stage_enum = stage if isinstance(stage, Stage) else Stage(stage)
    if stage_enum == Stage.FINANCIAL_CALCULATIONS:
        return "Final review"
    return "Fill workbook"


def build_status_snapshot(
    stage: Stage | str,
    *,
    message: str,
    severity: Severity,
    detail_message: str | None = None,
) -> dict[str, object]:
    stage_enum = stage if isinstance(stage, Stage) else Stage(stage)
    return {
        "index": stage_index(stage_enum),
        "stage": display_stage_name(stage_enum),
        "message": message,
        "severity": severity,
        "detail_message": detail_message,
    }


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


def normalize_issue_detail(detail_message: object | None, message_text: str | None = None) -> str | None:
    if detail_message is None:
        return None
    detail_text = str(detail_message).strip()
    if not detail_text:
        return None
    if message_text is not None and detail_text == str(message_text).strip():
        return None
    return detail_text


def workbook_issue_scope_label(issue_state: dict[str, object]) -> str:
    scope_label = str(issue_state.get("sheet_name") or "").strip()
    return scope_label or "Workbook entry"


def workbook_issue_severity(issue_state: dict[str, object], fallback: Severity = Severity.WARNING) -> Severity:
    severity_value = issue_state.get("severity")
    if isinstance(severity_value, Severity):
        return severity_value
    if isinstance(severity_value, str):
        try:
            return Severity(severity_value)
        except ValueError:
            return fallback
    return fallback


def active_workbook_issue_snapshot() -> dict[str, object] | None:
    workbook_retry = st.session_state.get("workbook_retry")
    if (
        isinstance(workbook_retry, dict)
        and (
            workbook_retry.get("phase") in {"retry", "failed"}
            or workbook_retry.get("retry_reason") is not None
            or workbook_issue_severity(workbook_retry, Severity.INFO) == Severity.ERROR
        )
    ):
        return workbook_retry
    return None


def active_rate_limit_snapshot() -> dict[str, object] | None:
    workbook_retry = st.session_state.get("workbook_retry")
    if (
        isinstance(workbook_retry, dict)
        and workbook_retry.get("retry_reason") == "rate_limit"
        and workbook_retry.get("retry_until_ms") is not None
    ):
        return {
            "scope": "workbook",
            "label": workbook_retry.get("sheet_name") or "Workbook entry",
            "retry_until_ms": int(workbook_retry["retry_until_ms"]),
            "attempt_number": workbook_retry.get("attempt_number"),
            "max_attempts": workbook_retry.get("max_attempts"),
        }

    pipeline = st.session_state.get("ocr_pipeline", {})
    rate_limited_pipes = [
        pipe
        for pipe in pipeline.get("pipes", [])
        if pipe.get("status") == "retrying"
        and pipe.get("retry_reason") == "rate_limit"
        and pipe.get("retry_until_ms") is not None
    ]
    if not rate_limited_pipes:
        return None

    active_pipe = min(rate_limited_pipes, key=lambda pipe: int(pipe["retry_until_ms"]))
    return {
        "scope": "ocr",
        "label": f"Lane {int(active_pipe['pipe_number'])}",
        "retry_until_ms": int(active_pipe["retry_until_ms"]),
        "attempt_number": active_pipe.get("attempt_number"),
        "max_attempts": active_pipe.get("max_attempts"),
    }


def describe_workbook_issue(
    issue_state: dict[str, object],
    *,
    fallback_message: str | None = None,
    fallback_severity: Severity = Severity.WARNING,
) -> dict[str, str | None]:
    sheet_label = workbook_issue_scope_label(issue_state)
    severity = workbook_issue_severity(issue_state, fallback_severity)
    retry_reason = str(issue_state.get("retry_reason") or "").strip()
    phase = str(issue_state.get("phase") or "").strip()
    message_text = str(issue_state.get("message") or fallback_message or "").strip()
    detail_text = normalize_issue_detail(issue_state.get("detail_message"), message_text)
    attempt_number = issue_state.get("attempt_number")
    max_attempts = issue_state.get("max_attempts")
    retry_delay_seconds = float(issue_state.get("retry_delay_seconds") or 0.0)
    retry_until_ms = issue_state.get("retry_until_ms")

    if retry_reason == "rate_limit":
        progress_copy = (
            f"Pass {attempt_number}/{max_attempts}"
            if attempt_number is not None and max_attempts is not None
            else "Automatic retry"
        )
        attempt_copy = (
            f"Pass {attempt_number}/{max_attempts} is queued and will resume automatically."
            if attempt_number is not None and max_attempts is not None
            else "Workbook entry will resume automatically."
        )
        return {
            "tone_class": "throttle",
            "banner_class": "workbook-retry-banner",
            "kicker": "Rate limit cooldown",
            "title": f"{current_provider_display_name()} rate limit detected",
            "body": f"{sheet_label} is paused. {attempt_copy}",
            "detail": detail_text,
            "timer_label": progress_copy,
            "timer_markup": build_timing_markup(
                "0.0s",
                countdown_target_ms=int(retry_until_ms),
            )
            if retry_until_ms is not None
            else build_timing_markup(format_seconds_label(retry_delay_seconds)),
        }

    if retry_reason == "timeout":
        timer_markup = build_timing_markup("Immediate")
        if retry_until_ms is not None and retry_delay_seconds > 0:
            timer_markup = build_timing_markup(
                format_seconds_label(retry_delay_seconds),
                countdown_target_ms=int(retry_until_ms),
            )
        elif retry_delay_seconds > 0:
            timer_markup = build_timing_markup(format_seconds_label(retry_delay_seconds))
        return {
            "tone_class": "warning",
            "banner_class": "workbook-retry-banner",
            "kicker": "Processing timeout",
            "title": f"{sheet_label} is retrying",
            "body": message_text or f"{sheet_label} exceeded the current processing limit.",
            "detail": detail_text,
            "timer_label": "Next pass",
            "timer_markup": timer_markup,
        }

    if phase == "retry":
        timer_markup = build_timing_markup("Queued")
        if retry_until_ms is not None and retry_delay_seconds > 0:
            timer_markup = build_timing_markup(
                format_seconds_label(retry_delay_seconds),
                countdown_target_ms=int(retry_until_ms),
            )
        elif retry_delay_seconds > 0:
            timer_markup = build_timing_markup(format_seconds_label(retry_delay_seconds))
        return {
            "tone_class": "warning",
            "banner_class": "workbook-retry-banner",
            "kicker": "Retry queued",
            "title": f"{sheet_label} is retrying",
            "body": message_text or f"Another pass is queued for {sheet_label}.",
            "detail": detail_text,
            "timer_label": "Next pass",
            "timer_markup": timer_markup,
        }

    banner_class = "workbook-retry-banner is-error" if severity == Severity.ERROR else "workbook-retry-banner"
    return {
        "tone_class": "error" if severity == Severity.ERROR else "warning",
        "banner_class": banner_class,
        "kicker": "Run stopped" if severity == Severity.ERROR else "Needs attention",
        "title": f"{sheet_label} needs attention",
        "body": message_text or f"{sheet_label} could not continue.",
        "detail": detail_text,
        "timer_label": "State",
        "timer_markup": build_timing_markup("Stopped" if severity == Severity.ERROR else "Attention"),
    }


def build_workbook_issue_markup(
    issue_state: dict[str, object],
    *,
    fallback_message: str | None = None,
    fallback_severity: Severity = Severity.WARNING,
) -> str:
    issue_copy = describe_workbook_issue(
        issue_state,
        fallback_message=fallback_message,
        fallback_severity=fallback_severity,
    )
    detail_markup = ""
    if issue_copy["detail"]:
        detail_markup = f'<div class="workbook-retry-detail">{html.escape(issue_copy["detail"])}</div>'
    return dedent(
        f"""
        <div class="{html.escape(issue_copy['banner_class'])}">
            <div class="workbook-retry-notch"></div>
            <div class="workbook-retry-copy">
                <div class="workbook-retry-kicker">{html.escape(issue_copy["kicker"])}</div>
                <div class="workbook-retry-title">{html.escape(issue_copy["title"])}</div>
                <div class="workbook-retry-body">{html.escape(issue_copy["body"])}</div>
                {detail_markup}
            </div>
            <div class="workbook-retry-timer-shell">
                <span class="workbook-retry-timer-label">{html.escape(issue_copy["timer_label"])}</span>
                <div class="workbook-retry-timer">{issue_copy["timer_markup"]}</div>
            </div>
        </div>
        """
    ).strip()


def provider_logo_markup(provider: str) -> str:
    label = provider_display_name(provider)
    logo_url = html.escape(PROVIDER_LOGO_URLS.get(provider, ""), quote=True)
    if not logo_url:
        return f'<span class="provider-logo-fallback" aria-hidden="true">{html.escape(label[:1])}</span>'
    return (
        f'<img class="provider-logo-image" src="{logo_url}" '
        f'alt="{html.escape(label)} logo" loading="lazy" />'
    )


def build_provider_selector_markup(selected_provider: str) -> str:
    cards: list[str] = []
    for provider in LLM_PROVIDER_OPTIONS:
        enabled = provider == DEFAULT_LLM_PROVIDER
        selected = enabled and provider == selected_provider
        label = provider_display_name(provider)
        model_name = locked_model_for_provider(provider)
        cards.append(
            dedent(
                f"""
            <div class="provider-card{' is-selected' if selected else ''}{' is-disabled' if not enabled else ''}">
                <div class="provider-mark provider-mark-{provider}">
                    {provider_logo_markup(provider)}
                </div>
                <div class="provider-copy">
                    <div class="provider-name-row">
                        <span class="provider-name">{html.escape(label)}</span>
                        <span class="provider-state">{'Active' if selected else 'Unavailable'}</span>
                    </div>
                    <div class="provider-model-name">{html.escape(model_name)}</div>
                </div>
            </div>
            """
            ).strip()
        )
    return f'<div class="provider-rail">{"".join(cards)}</div>'


@st.dialog("Run settings")
def render_settings_dialog(base_settings: Settings) -> None:
    defaults = build_runtime_settings_state(base_settings)
    current = {**defaults, **dict(st.session_state["runtime_settings"])}
    current_provider = DEFAULT_LLM_PROVIDER
    current["llm_provider"] = current_provider

    with st.form("runtime_settings_form", border=False):
        st.caption(
            "Only provider credentials live in `.env`. OpenAI is active for this build, and these settings apply to the current browser session."
        )
        st.markdown("**Provider**")
        st.markdown(build_provider_selector_markup(current_provider), unsafe_allow_html=True)
        ocr_parallel_workers = st.number_input(
            "Parallel lanes",
            min_value=1,
            max_value=10,
            value=int(current["ocr_parallel_workers"]),
            step=1,
        )

        save_col, reset_col = st.columns(2, gap="small")
        save_clicked = save_col.form_submit_button("Save settings", width="stretch")
        reset_clicked = reset_col.form_submit_button("Use defaults", width="stretch")

    if reset_clicked:
        st.session_state["runtime_settings"] = defaults
        st.rerun()

    if save_clicked:
        st.session_state["runtime_settings"] = {
            **defaults,
            **current,
            "llm_provider": DEFAULT_LLM_PROVIDER,
            "ocr_parallel_workers": int(ocr_parallel_workers),
        }
        st.rerun()


def build_workflow_stage_cards() -> list[dict[str, object]]:
    progress = st.session_state.get("stage_progress", {})
    active_stage = workflow_stage(st.session_state.get("active_stage", Stage.OCR.value))
    is_running = bool(st.session_state.get("is_running"))
    run_started = bool(st.session_state.get("run_started"))
    last_status = st.session_state.get("last_status")
    result = st.session_state.get("result")
    needs_input = result_waiting_on_question(result)

    cards: list[dict[str, object]] = []
    for index, stage in enumerate(WORKFLOW_STAGES, start=1):
        member_stages = (stage,) if stage == Stage.OCR else (Stage.DATA_ENTRY, Stage.FINANCIAL_CALCULATIONS)
        is_complete = all(
            progress.get(member_stage.value, (0, 1))[0] >= progress.get(member_stage.value, (0, 1))[1]
            and progress.get(member_stage.value, (0, 1))[1] > 0
            for member_stage in member_stages
        )
        is_active = stage == active_stage
        state_label = "Waiting"

        if is_complete:
            state_label = "Complete"
        elif is_active:
            if stage == Stage.DATA_ENTRY and needs_input:
                state_label = "Needs input"
            elif is_running:
                state_label = "In progress"
            elif run_started and last_status is not None and last_status["severity"] == Severity.ERROR:
                state_label = "Stopped"
            else:
                state_label = "Ready"
        elif not run_started and index == 1:
            state_label = "Ready"

        cards.append(
            {
                "index": index,
                "stage": stage,
                "name": display_stage_name(stage),
                "compact_name": display_stage_name(stage, compact=True),
                "state_label": state_label,
                "is_active": is_active,
                "is_complete": is_complete,
            }
        )

    return cards


def build_taskbar_markup() -> str:
    cards = build_workflow_stage_cards()
    active_card = next((card for card in cards if card["is_active"]), cards[0])
    last_status = st.session_state.get("last_status")
    run_started = bool(st.session_state.get("run_started"))
    is_running = bool(st.session_state.get("is_running"))

    context_label = f"Stage {active_card['index']} of {len(WORKFLOW_STAGES)} • {active_card['name']}"
    status_label = str(active_card["state_label"])
    status_message = "Upload a planner PDF to start the two-step intake."
    tone = "info"
    meta_parts: list[str] = []

    if last_status is not None:
        context_label = f"Stage {last_status['index']} of {len(WORKFLOW_STAGES)} • {last_status['stage']}"
        status_message = str(last_status["message"])
        tone = status_tone(last_status["severity"])

    if status_label == "Complete" and tone == "info" and run_started and not is_running:
        tone = "complete"

    rate_limit_snapshot = (
        active_rate_limit_snapshot()
        if last_status is not None
        and last_status["severity"] == Severity.WARNING
        and "rate limit" in str(last_status["message"]).lower()
        else None
    )
    if rate_limit_snapshot is not None:
        tone = "throttle"
        status_label = "Cooldown"
        status_message = f"{rate_limit_snapshot.get('label') or 'Workbook entry'} paused for automatic retry."
        attempt_number = rate_limit_snapshot.get("attempt_number")
        max_attempts = rate_limit_snapshot.get("max_attempts")
        if attempt_number is not None and max_attempts is not None:
            meta_parts.append(f'<span class="taskbar-status-meta">Pass {attempt_number}/{max_attempts}</span>')
        meta_parts.append(
            '<span class="taskbar-status-timer">'
            + build_timing_markup("0.0s", countdown_target_ms=int(rate_limit_snapshot["retry_until_ms"]))
            + "</span>"
        )

    stage_markup = "".join(
        dedent(
            f"""
            <div class="taskbar-stage{' is-active' if card['is_active'] else ''}{' is-complete' if card['is_complete'] else ''}">
                <span class="taskbar-stage-index">{card['index']}</span>
                <span class="taskbar-stage-copy">
                    <span class="taskbar-stage-label">{html.escape(str(card['name'] if card['is_active'] else card['compact_name']))}</span>
                    <span class="taskbar-stage-state">{html.escape(str(card['state_label']))}</span>
                </span>
            </div>
            """
        ).strip()
        for card in cards
    )

    return dedent(
        f"""
        <div class="taskbar-shell">
            <div class="taskbar-bar">
                <div class="taskbar-brand-lockup">
                    <div class="taskbar-brand-row">
                        <span class="taskbar-brand">{html.escape(APP_NAME)}</span>
                        <span class="taskbar-strap">Turn planner PDFs into locked workbooks</span>
                    </div>
                    <div class="taskbar-context">{html.escape(context_label)}</div>
                </div>
                <div class="taskbar-status {html.escape(tone)}">
                    <span class="taskbar-status-pill">{html.escape(status_label)}</span>
                    <span class="taskbar-status-message" title="{html.escape(status_message, quote=True)}">{html.escape(status_message)}</span>
                    {''.join(meta_parts)}
                </div>
                <div class="taskbar-stage-rail">{stage_markup}</div>
            </div>
        </div>
        """
    ).strip()


def render_masthead() -> bool:
    title_col, action_col = st.columns([0.88, 0.12], gap="small")
    with title_col:
        render_html_block(build_taskbar_markup())
    with action_col:
        st.markdown("<div style='height:0.18rem'></div>", unsafe_allow_html=True)
        return st.button("Settings", key=run_scoped_widget_key("open_run_settings"), width="stretch")


def render_status(status_placeholder) -> None:
    status = st.session_state.get("last_status")
    if status is None:
        status_placeholder.markdown(
            """
            <div class="status-card info">
                <div class="status-label">System ready</div>
                <div class="status-message">Upload a planner PDF to start the two-step intake.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    tone = status_tone(status["severity"])
    message_text = str(status["message"])
    detail_text = normalize_issue_detail(status.get("detail_message"), message_text)
    workbook_issue = active_workbook_issue_snapshot()
    if workbook_issue is not None:
        issue_copy = describe_workbook_issue(
            workbook_issue,
            fallback_message=message_text,
            fallback_severity=status["severity"],
        )
        detail_markup = ""
        if issue_copy["detail"]:
            detail_markup = f'<div class="status-detail">{html.escape(issue_copy["detail"])}</div>'
        status_placeholder.markdown(
            f"""
            <div class="status-card {html.escape(issue_copy['tone_class'])}">
                <div class="status-label">Stage {status["index"]} of {len(WORKFLOW_STAGES)} • {html.escape(str(status["stage"]))}</div>
                <div class="throttle-status-layout">
                    <div>
                        <div class="throttle-status-kicker">{html.escape(issue_copy["kicker"])}</div>
                        <div class="throttle-status-title">{html.escape(issue_copy["title"])}</div>
                        <div class="status-message">{html.escape(issue_copy["body"])}</div>
                        {detail_markup}
                    </div>
                    <div class="throttle-status-timer-shell">
                        <span class="throttle-status-timer-label">{html.escape(issue_copy["timer_label"])}</span>
                        <div class="throttle-status-timer">{issue_copy["timer_markup"]}</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    rate_limit_snapshot = (
        active_rate_limit_snapshot()
        if status["severity"] == Severity.WARNING and "rate limit" in message_text.lower()
        else None
    )
    if rate_limit_snapshot is not None:
        attempt_number = rate_limit_snapshot.get("attempt_number")
        max_attempts = rate_limit_snapshot.get("max_attempts")
        timer_markup = build_timing_markup(
            "0.0s",
            countdown_target_ms=int(rate_limit_snapshot["retry_until_ms"]),
        )
        progress_copy = (
            f"Pass {attempt_number}/{max_attempts}"
            if attempt_number is not None and max_attempts is not None
            else "Automatic retry"
        )
        status_placeholder.markdown(
            f"""
            <div class="status-card throttle">
                <div class="status-label">Stage {status["index"]} of {len(WORKFLOW_STAGES)} • {html.escape(str(status["stage"]))}</div>
                <div class="throttle-status-layout">
                    <div>
                        <div class="throttle-status-kicker">Rate limit cooldown</div>
                        <div class="throttle-status-title">{html.escape(str(rate_limit_snapshot["label"]))} is paused</div>
                        <div class="status-message">{html.escape(message_text)}</div>
                    </div>
                    <div class="throttle-status-timer-shell">
                        <span class="throttle-status-timer-label">{html.escape(progress_copy)}</span>
                        <div class="throttle-status-timer">{timer_markup}</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    status_placeholder.markdown(
        f"""
        <div class="status-card {tone}">
            <div class="status-label">Stage {status["index"]} of {len(WORKFLOW_STAGES)} • {html.escape(status["stage"])}</div>
            <div class="status-message">{html.escape(message_text)}</div>
            {'<div class="status-detail">' + html.escape(detail_text) + '</div>' if detail_text else ''}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_stage_progress() -> None:
    cards: list[str] = []
    for card in build_workflow_stage_cards():
        classes = ["stage-chip"]
        if card["is_active"]:
            classes.append("active")
        else:
            classes.append("compact")

        if card["is_complete"]:
            classes.append("complete")

        cards.append(
            dedent(
                f"""
                <div class="{' '.join(classes)}">
                    <div class="stage-chip-number">Step {card['index']}</div>
                    <div class="stage-chip-name">{html.escape(str(card['name'] if card['is_active'] else card['compact_name']))}</div>
                    <div class="stage-chip-state">{html.escape(str(card['state_label']))}</div>
                </div>
                """
            ).strip()
        )

    st.markdown(f'<div class="stage-rail">{"".join(cards)}</div>', unsafe_allow_html=True)


def render_shell_chrome(chrome_placeholder) -> bool:
    chrome_placeholder.empty()

    with chrome_placeholder.container():
        settings_clicked = render_masthead()
    return settings_clicked


def append_event(event: RunEvent) -> None:
    st.session_state["active_stage"] = event.stage.value
    st.session_state["stage_progress"][event.stage.value] = (event.stage_completed, event.stage_total)
    sync_agent_trace_from_event(event)
    if event.phase != "heartbeat":
        st.session_state["last_status"] = build_status_snapshot(
            event.stage,
            message=event.message,
            severity=event.severity,
            detail_message=event.detail_message,
        )
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
                pipe["last_error"] = event.detail_message or event.message
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
                pipe["last_error"] = event.detail_message or event.message
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
    elif event.stage in {Stage.DATA_ENTRY, Stage.FINANCIAL_CALCULATIONS} and event.phase != "heartbeat":
        if event.phase in {"retry", "failed"}:
            now_ms = int(time.time() * 1000)
            retry_state = build_workbook_retry_state()
            retry_state.update(
                {
                    "sheet_name": event.sheet_name or display_stage_name(event.stage),
                    "message": event.message,
                    "detail_message": event.detail_message,
                    "phase": event.phase,
                    "severity": event.severity,
                    "attempt_number": event.attempt_number,
                    "max_attempts": event.max_attempts,
                    "retry_delay_seconds": event.retry_delay_seconds,
                    "retry_until_ms": (
                        now_ms + int(event.retry_delay_seconds * 1000)
                        if event.retry_delay_seconds is not None
                        else None
                    ),
                    "retry_reason": event.retry_reason,
                }
            )
            st.session_state["workbook_retry"] = retry_state
        else:
            st.session_state["workbook_retry"] = None
    if event.phase != "heartbeat":
        prefix = {
            Severity.INFO: "[INFO]",
            Severity.WARNING: "[WARN]",
            Severity.ERROR: "[ERROR]",
        }[event.severity]
        log_line = f"{prefix} {display_stage_name(event.stage)}: {event.message}"
        detail_text = normalize_issue_detail(event.detail_message, event.message)
        if detail_text is not None:
            log_line = f"{log_line} Details: {detail_text}"
        st.session_state["logs"].append(log_line)
    if event.artifacts is not None:
        st.session_state["result"] = event.artifacts
        st.session_state["current_job_id"] = event.artifacts.job_id
        if active_entry_question_resume_job_id() == event.artifacts.job_id:
            st.session_state[QUESTION_ACTIVE_RESUME_STATE_KEY] = None


def render_logs(_log_placeholder) -> None:
    return


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
    active_page_slots = sum(
        1
        for pipe in pipes
        if pipe.get("status") in active_statuses and pipe.get("page_number") is not None
    )
    progress_total = max(page_total, completed_pages, len(page_previews))
    visible_active_slots = min(max(progress_total - completed_pages, 0), active_page_slots)
    chunk_count = progress_total or max(pipe_total * 4, 12)
    progress_chunks: list[str] = []
    active_chunk_offset = 0
    for index in range(chunk_count):
        if index < completed_pages:
            progress_chunks.append('<span class="chunk-progress-block complete" aria-hidden="true"></span>')
            continue
        if index < completed_pages + visible_active_slots:
            progress_chunks.append(
                f'<span class="chunk-progress-block active" aria-hidden="true" style="--progress-delay: {active_chunk_offset * 0.16:.2f}s"></span>'
            )
            active_chunk_offset += 1
            continue
        progress_chunks.append('<span class="chunk-progress-block pending" aria-hidden="true"></span>')
    progress_chunks_markup = "".join(progress_chunks)
    progress_label = (
        f"{completed_pages} of {progress_total} pages complete"
        if progress_total
        else "Preparing pages for review"
    )
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
        status_class = "paused" if status == "retrying" and retry_reason == "rate_limit" else status
        preview_page_number = page_number if page_number is not None else last_page
        preview_src = page_previews.get(int(preview_page_number)) if preview_page_number is not None else None
        just_completed = bool(pipe.get("flash_complete"))
        recently_completed = bool(
            status == "complete"
            and completed_at_ms is not None
            and (now_ms - int(completed_at_ms)) < 1400
        )
        is_active = status in active_statuses
        warning_message = (
            str(pipe["last_error"]).strip()
            if status in {"retrying", "failed"} and pipe.get("last_error")
            else None
        )

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
                f"Paused on page {page_number}"
                if retry_reason == "rate_limit"
                else f"Retrying page {page_number}"
            )
            if attempt_number is not None and max_attempts is not None:
                detail_label = f"Pass {attempt_number}/{max_attempts}"
            if retry_reason == "rate_limit":
                detail_label = (
                    f"{current_provider_display_name()} rate limit hit. Pass {attempt_number}/{max_attempts} will resume automatically."
                    if attempt_number is not None and max_attempts is not None
                    else f"{current_provider_display_name()} rate limit hit. This lane will resume automatically."
                )
            if retry_until_ms is not None:
                remaining_ms = int(retry_until_ms) - now_ms
                if remaining_ms > 0:
                    timing_markup = (
                        "Retry in "
                        + build_timing_markup(
                            format_seconds_label(remaining_ms / 1000),
                            countdown_target_ms=int(retry_until_ms),
                        )
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
                timing_markup = "Retry in " + build_timing_markup(format_seconds_label(retry_delay_seconds))
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
        if warning_message:
            row_classes.append("has-warning")

        preview_markup = '<div class="ocr-preview empty">Waiting</div><div class="ocr-preview-caption">No page assigned</div>'
        if preview_src is not None and preview_page_number is not None:
            preview_state = "current" if page_number is not None else "last"
            preview_caption = (
                f"Page {preview_page_number} in progress"
                if page_number is not None
                else f"Last page {preview_page_number}"
            )
            preview_markup = (
                f'<div class="ocr-preview has-image {preview_state}" role="img" aria-label="Preview of page {int(preview_page_number)}" '
                f'style="background-image: url(\'{html.escape(preview_src, quote=True)}\');"></div>'
                f'<div class="ocr-preview-caption">{html.escape(preview_caption)}</div>'
            )

        warning_markup = ""
        if warning_message:
            warning_markup = (
                '<div class="ocr-warning-badge" aria-hidden="true">Alert</div>'
                '<div class="ocr-warning-card">'
                '<span class="ocr-warning-kicker">Lane warning</span>'
                f'<div class="ocr-warning-copy">{html.escape(warning_message)}</div>'
                "</div>"
            )

        rows.append(
            dedent(
                f"""
                <div class="{' '.join(row_classes)}">
                    <div class="ocr-pipe">
                        <span class="ocr-dot {html.escape(status_class)}"></span>
                        Lane {int(pipe["pipe_number"])}
                    </div>
                    <div class="ocr-status">
                        <strong>{html.escape(status_label)}</strong>
                        <span class="ocr-detail">{html.escape(detail_label)}</span>
                    </div>
                    <div class="ocr-timing">{timing_markup}</div>
                    {warning_markup}
                    <div class="ocr-preview-shell">{preview_markup}</div>
                </div>
                """
            ).strip()
        )
        if just_completed:
            pipe["flash_complete"] = False

    rows_markup = "".join(rows)

    render_html_block(
        dedent(
            f"""
            <div class="section-card live-card">
                <h3 class="section-title">Parallel document review</h3>
                <div class="chunk-progress">
                    <div
                        class="chunk-progress-track"
                        role="progressbar"
                        aria-label="Document progress"
                        aria-valuemin="0"
                        aria-valuemax="{progress_total or chunk_count}"
                        aria-valuenow="{min(completed_pages, progress_total or chunk_count)}"
                        aria-valuetext="{html.escape(progress_label, quote=True)}"
                    >
                        {progress_chunks_markup}
                    </div>
                </div>
                <div class="ocr-grid">{rows_markup}</div>
            </div>
            """
        ).strip()
    )


def render_workbook_stage() -> None:
    progress = st.session_state["stage_progress"]
    active_stage = Stage(st.session_state["active_stage"])
    is_running = bool(st.session_state.get("is_running"))
    source_filename = st.session_state.get("source_filename") or "Current upload"
    last_status = st.session_state.get("last_status")
    result = st.session_state.get("result")
    needs_input = result_waiting_on_question(result)
    workbook_retry = st.session_state.get("workbook_retry")
    trace = st.session_state.get("agent_trace", build_agent_trace_state())

    mapping_completed, mapping_total = progress.get(Stage.DATA_ENTRY.value, (0, 1))
    checks_completed, checks_total = progress.get(Stage.FINANCIAL_CALCULATIONS.value, (0, 1))
    checks_started = (
        active_stage == Stage.FINANCIAL_CALCULATIONS
        or checks_completed > 0
        or checks_total > 1
    )

    mapping_done = mapping_completed >= mapping_total and mapping_total > 0
    if needs_input and active_stage == Stage.DATA_ENTRY:
        mapping_state = "Needs your input"
        mapping_copy = "Waiting for your answer before workbook entry can continue."
    elif mapping_done:
        mapping_state = "Complete"
        mapping_copy = f"{mapping_completed}/{mapping_total} sections are complete."
    elif active_stage == Stage.DATA_ENTRY and is_running:
        mapping_state = "In progress"
        mapping_copy = f"{mapping_completed}/{mapping_total} sections are complete so far."
    else:
        mapping_state = "Queued"
        mapping_copy = f"{mapping_completed}/{mapping_total} sections are ready."

    checks_done = checks_completed >= checks_total and checks_total > 0
    if checks_done:
        checks_state = "Complete"
        checks_copy = f"{checks_completed}/{checks_total} final checks are complete."
    elif active_stage == Stage.FINANCIAL_CALCULATIONS and is_running:
        checks_state = "In progress"
        checks_copy = f"{checks_completed}/{checks_total} final checks are complete so far."
    elif checks_started:
        checks_state = "Queued"
        checks_copy = f"{checks_completed}/{checks_total} final checks are complete."
    else:
        checks_state = "Queued"
        checks_copy = "Starts after all sections are entered."

    current_phase = display_workbook_phase_name(active_stage)
    workbook_message = (
        str(last_status["message"])
        if last_status is not None and workflow_stage(active_stage) == Stage.DATA_ENTRY
        else "Workbook entry and final review share one view."
    )
    if should_render_workbook_setup_shell(
        active_stage=active_stage,
        is_running=is_running,
        trace=trace,
        workbook_retry=workbook_retry,
    ):
        render_html_block(
            dedent(
                f"""
                <section class="immersive-workbook-shell">
                    {build_workbook_setup_markup(
                        source_filename=source_filename,
                        status_message=workbook_message,
                        mapping_total=mapping_total,
                        checks_total=checks_total,
                    )}
                </section>
                """
            ).strip()
        )
        mark_agent_trace_summary_rendered()
        return

    entry_state = load_live_entry_state()
    workbook_retry_markup = ""
    if isinstance(workbook_retry, dict):
        workbook_retry_markup = (
            '<div class="immersive-alert-row">'
            + build_workbook_issue_markup(
                workbook_retry,
                fallback_message=workbook_message,
                fallback_severity=last_status["severity"] if last_status is not None else Severity.WARNING,
            )
            + "</div>"
        )
    agent_trace_markup = build_agent_trace_markup(
        active_stage=active_stage,
        mapping_completed=mapping_completed,
        mapping_total=mapping_total,
        checks_completed=checks_completed,
        checks_total=checks_total,
    )
    sheet_desk_markup = build_sheet_desk_markup(
        entry_state=entry_state,
        active_stage=active_stage,
        source_filename=source_filename,
        current_phase=current_phase,
        workbook_message=workbook_message,
        mapping_completed=mapping_completed,
        mapping_total=mapping_total,
        checks_completed=checks_completed,
        checks_total=checks_total,
        mapping_state=mapping_state,
        mapping_copy=mapping_copy,
        checks_state=checks_state,
        checks_copy=checks_copy,
    )

    render_html_block(
        dedent(
            f"""
            <section class="immersive-workbook-shell">
                {workbook_retry_markup}
                <div class="immersive-workbook-grid">
                    <div class="immersive-pane immersive-agent-pane">
                        <div class="immersive-pane-label">Status</div>
                        {agent_trace_markup}
                    </div>
                    <div class="immersive-pane immersive-sheet-pane">
                        {sheet_desk_markup}
                    </div>
                </div>
            </section>
            """
        ).strip()
    )
    mark_agent_trace_summary_rendered()


def render_upload_panel(
    *,
    context_key: str,
    title: str,
    copy: str | None = None,
) -> tuple[object | None, str | None]:
    def render_launch_state() -> None:
        st.markdown(
            f"""
            <div class="section-card">
                <h3 class="section-title">Starting workbook build</h3>
                <div class="section-copy">Preparing document review.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    upload_key = run_scoped_widget_key(f"pdf_upload_{context_key}")
    run_key = run_scoped_widget_key(f"run_import_{context_key}")
    copy_markup = f'<div class="section-copy">{html.escape(copy)}</div>' if copy else ""
    panel_placeholder = st.empty()
    with panel_placeholder.container():
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
            key=upload_key,
        )
        run_clicked = st.button(
            "Build workbook",
            key=run_key,
            type="primary",
            width="stretch",
            disabled=uploaded_file is None,
        )
        st.markdown("</div>", unsafe_allow_html=True)
    if run_clicked:
        panel_placeholder.empty()
        with panel_placeholder.container():
            render_launch_state()
        return uploaded_file, "upload"
    return uploaded_file, None


def render_stage_focus() -> None:
    stage = Stage(st.session_state["active_stage"])

    if stage == Stage.OCR:
        render_ocr_parallel()
        return

    render_workbook_stage()


def build_entry_question_context(result: ImportArtifacts) -> dict[str, str]:
    question = result.pending_question
    if question is None:
        return {"progress_text": "Progress unavailable"}

    entry_state = None
    if result.entry_state_path is not None and result.entry_state_path.exists():
        entry_state = load_entry_state(result.entry_state_path)

    completed = entry_state.current_sheet_index if entry_state is not None else 0
    total = len(entry_state.sheet_order) if entry_state is not None else 0
    remaining = max(0, total - completed)
    progress_text = f"{completed}/{total} completed • {remaining} remaining" if total > 0 else "Progress unavailable"

    return {"progress_text": progress_text}


def entry_question_signature(question: AgentQuestion) -> str:
    return json.dumps(
        {
            "id": question.id,
            "sheet_name": question.sheet_name,
            "prompt": question.prompt,
            "affected_targets": question.affected_targets,
            "options": [option.model_dump(mode="json") for option in question.options],
            "allow_free_text": question.allow_free_text,
        },
        sort_keys=True,
    )


def entry_question_widget_keys(job_id: str, question_id: str) -> dict[str, str]:
    return {
        "form": f"entry_question_{job_id}_{question_id}",
        "options": f"question_options_{job_id}",
        "free_text": f"question_free_text_{job_id}",
    }


def sync_entry_question_widget_state(job_id: str, question: AgentQuestion) -> dict[str, str]:
    widget_keys = entry_question_widget_keys(job_id, question.id)
    question_signature = entry_question_signature(question)
    tracked_questions = st.session_state.setdefault(QUESTION_WIDGET_STATE_KEY, {})
    previous_question_signature = tracked_questions.get(job_id)
    if previous_question_signature != question_signature:
        st.session_state.pop(widget_keys["options"], None)
        st.session_state.pop(widget_keys["free_text"], None)
        tracked_questions[job_id] = question_signature
    return widget_keys


def pop_entry_question_submission(job_id: str, question_signature: str) -> tuple[str | None, str | None, bool]:
    submission = st.session_state.pop(QUESTION_SUBMISSION_STATE_KEY, None)
    if not isinstance(submission, dict):
        return None, None, False

    if submission.get("job_id") != job_id or submission.get("question_signature") != question_signature:
        return None, None, False

    answer = submission.get("answer")
    source = submission.get("source")
    if not isinstance(answer, str) or not isinstance(source, str):
        return None, None, False
    return answer, source, True


def pop_entry_question_resume() -> tuple[str | None, str | None, str | None]:
    pending_resume = st.session_state.pop(QUESTION_PENDING_RESUME_STATE_KEY, None)
    if not isinstance(pending_resume, dict):
        return None, None, None

    job_id = pending_resume.get("job_id")
    answer = pending_resume.get("answer")
    source = pending_resume.get("source")
    if not isinstance(job_id, str) or not isinstance(answer, str) or not isinstance(source, str):
        return None, None, None
    return job_id, answer, source


def render_entry_question_form(
    result: ImportArtifacts,
    question_context: dict[str, str],
    *,
    modal: bool,
) -> tuple[str | None, str | None, bool]:
    question = result.pending_question
    if question is None:
        return None, None, False
    question_signature = entry_question_signature(question)
    widget_keys = sync_entry_question_widget_state(result.job_id, question)

    progress_chip = ""
    if question_context["progress_text"] != "Progress unavailable":
        progress_chip = f'<div class="question-chip">{html.escape(question_context["progress_text"])}</div>'

    rereview_chip = ""
    if question.pdf_rereviewed:
        rereview_chip = '<div class="question-chip signal">Raw PDF re-reviewed</div>'

    render_html_block(
        dedent(
            f"""
            <div class="question-shell {'is-modal' if modal else 'is-inline'}">
                <div class="question-panel">
                    <div class="question-header">
                        <div class="question-header-row">
                            <div class="question-kicker">Planner decision required</div>
                            <div class="question-meta-strip">
                                <div class="question-chip">{html.escape(question.sheet_name)}</div>
                                {progress_chip}
                                {rereview_chip}
                            </div>
                        </div>
                        <h3 class="question-title">{html.escape(question.prompt)}</h3>
                    </div>
                </div>
            </div>
            """
        ).strip()
    )

    with st.form(widget_keys["form"], border=False):
        selected_option_value = None
        if question.options:
            option_labels = [option.label for option in question.options]
            option_values = {option.label: option.value for option in question.options}
            option_captions = [
                option.description or "Recommended from the current document context."
                for option in question.options
            ]
            selected_option_label = st.radio(
                "Choose an answer",
                options=option_labels,
                index=0,
                key=widget_keys["options"],
                captions=None if modal else option_captions,
                label_visibility="collapsed" if modal else "visible",
                width="stretch",
            )
            selected_option_value = option_values[selected_option_label]

        free_text = ""
        if question.allow_free_text:
            free_text = st.text_input(
                "Or write your own answer",
                key=widget_keys["free_text"],
                placeholder="Or type a custom answer",
                label_visibility="collapsed" if modal else "visible",
            )

        submit_col, delegate_col = st.columns([0.64, 0.36], gap="small")
        submitted = submit_col.form_submit_button(
            "Submit answer and continue",
            type="primary",
            width="stretch",
        )
        delegated = delegate_col.form_submit_button(
            "Figure it out",
            type="secondary",
            help="Let the agent make the best supported choice from the document.",
            width="stretch",
        )

    answer = None
    source = None
    if delegated:
        answer = ""
        source = "agent"
    elif submitted and free_text.strip():
        answer = free_text.strip()
        source = "free_text"
    elif submitted and selected_option_value is not None:
        answer = selected_option_value
        source = "option"

    if answer is None or source is None:
        return None, None, False

    if not modal:
        return answer, source, True

    st.session_state[QUESTION_SUBMISSION_STATE_KEY] = {
        "job_id": result.job_id,
        "question_signature": question_signature,
        "answer": answer,
        "source": source,
    }
    rerun_app()
    return None, None, False


def render_entry_question(result: ImportArtifacts) -> tuple[str | None, str | None, bool]:
    if result.pending_question is None:
        return None, None, False

    render_workbook_stage()
    queued_submission = pop_entry_question_submission(
        result.job_id,
        entry_question_signature(result.pending_question),
    )
    if queued_submission[2]:
        return queued_submission

    question_context = build_entry_question_context(result)
    dialog_renderer = getattr(st, "dialog", None)
    if callable(dialog_renderer):

        @dialog_renderer(QUESTION_DIALOG_TITLE, width="large", dismissible=False)
        def render_entry_question_dialog() -> None:
            render_entry_question_form(result, question_context, modal=True)

        render_entry_question_dialog()
        return None, None, False

    st.markdown("<div style='height:0.85rem'></div>", unsafe_allow_html=True)
    return render_entry_question_form(result, question_context, modal=False)


def render_result(result: ImportArtifacts | None) -> None:
    if result is None or result.review_report is None:
        return

    report = result.review_report
    headline = "Workbook ready" if result.success else "Review required"
    summary_copy = f"{len(report.mapped_assignments)} mapped, {len(report.unmapped_items)} unresolved."

    render_html_block(
        f"""
        <div class="section-card">
            <h3 class="section-title">{html.escape(headline)}</h3>
            <div class="section-copy">{html.escape(summary_copy)}</div>
        </div>
        """
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
                width="stretch",
            )
    with report_col:
        st.download_button(
            "Download review report",
            data=json.dumps(report.model_dump(mode="json"), indent=2),
            file_name=result.review_report_path.name if result.review_report_path else "review_report.json",
            mime="application/json",
            width="stretch",
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
                    "value": "" if assignment.value is None else str(assignment.value),
                    "comment": assignment.comment or "",
                }
                for assignment in report.mapped_assignments
            ],
            width="stretch",
            hide_index=True,
        )

    if report.unmapped_items:
        with st.expander(f"Unmapped items ({len(report.unmapped_items)})", expanded=False):
            st.json(report.unmapped_items)

    if report.assumptions:
        with st.expander(f"Assumptions ({len(report.assumptions)})", expanded=False):
            st.json(report.assumptions)


def render_work_area(work_placeholder, settings: Settings) -> tuple[object | None, str | None, tuple[str, str] | None]:
    uploaded_file = None
    run_mode = None
    resume_payload: tuple[str, str] | None = None
    result = st.session_state.get("result")
    is_running = bool(st.session_state.get("is_running"))
    run_started = bool(st.session_state.get("run_started"))
    status = st.session_state.get("last_status")

    with work_placeholder.container():
        if is_running:
            render_stage_focus()
            return None, None, None

        if result_waiting_on_question(result):
            answer, source, submitted = render_entry_question(result)
            if submitted and answer is not None and source is not None:
                resume_payload = (answer, source)
            return None, None, resume_payload

        if result is not None and result.review_report is not None:
            render_result(result)
            st.markdown("<div style='height:0.85rem'></div>", unsafe_allow_html=True)
            with st.expander("Import another PDF", expanded=False):
                uploaded_file, run_mode = render_upload_panel(
                    context_key="next",
                    title="Process another PDF",
                    copy=f"Current settings: {settings.ocr_parallel_workers} parallel lanes.",
                )
            return uploaded_file, run_mode, None

        if run_started and status is not None and status["severity"] == Severity.ERROR:
            failure_copy = str(status["message"])
            detail_text = normalize_issue_detail(status.get("detail_message"), failure_copy)
            if detail_text is not None:
                failure_copy = f"{failure_copy} Details: {detail_text}"
            uploaded_file, run_mode = render_upload_panel(
                context_key="retry",
                title="Run failed",
                copy=failure_copy,
            )
            return uploaded_file, run_mode, None

        uploaded_file, run_mode = render_upload_panel(
            context_key="primary",
            title="Upload PDF",
            copy=(
                "Use Settings to change the number of parallel lanes. "
                "Only provider credentials stay in `.env`."
            ),
        )
    return uploaded_file, run_mode, None


def consume_runner_events(
    *,
    events,
    settings: Settings,
    chrome_placeholder,
    log_placeholder,
    work_placeholder,
) -> None:
    try:
        for event in events:
            prior_agent_tokens = int(st.session_state.get("agent_trace", {}).get("token_count") or 0)
            prior_live_summary = str(st.session_state.get("agent_trace", {}).get("live_summary") or "")
            append_event(event)
            if event.phase == "heartbeat":
                current_agent_tokens = int(st.session_state.get("agent_trace", {}).get("token_count") or 0)
                current_live_summary = str(st.session_state.get("agent_trace", {}).get("live_summary") or "")
                if event.stage != Stage.DATA_ENTRY or (
                    current_agent_tokens == prior_agent_tokens and current_live_summary == prior_live_summary
                ):
                    continue
                render_shell_chrome(chrome_placeholder)
                render_logs(log_placeholder)
                render_work_area(work_placeholder, settings)
                continue
            render_shell_chrome(chrome_placeholder)
            render_logs(log_placeholder)
            render_work_area(work_placeholder, settings)
    except Exception as exc:
        st.session_state["logs"].append(f"[ERROR] Pipeline: {exc}")
        prior_status = st.session_state.get("last_status")
        status_message = f"Run failed: {exc}"
        detail_message = None
        if isinstance(prior_status, dict):
            detail_message = normalize_issue_detail(prior_status.get("detail_message"), status_message)
        root_cause = exc.__cause__ or exc
        if detail_message is None and root_cause is not exc:
            detail_message = normalize_issue_detail(str(root_cause), status_message)
        st.session_state["last_status"] = build_status_snapshot(
            Stage(st.session_state["active_stage"]),
            message=status_message,
            severity=Severity.ERROR,
            detail_message=detail_message,
        )
    finally:
        st.session_state["is_running"] = False
        st.session_state[QUESTION_ACTIVE_RESUME_STATE_KEY] = None
        render_shell_chrome(chrome_placeholder)
        render_logs(log_placeholder)
        render_work_area(work_placeholder, settings)


def unpack_work_area_result(result):
    if isinstance(result, tuple) and len(result) == 2:
        uploaded_file, second = result
        if isinstance(second, bool):
            return uploaded_file, ("upload" if second else None), None
        return uploaded_file, second, None
    uploaded_file, run_mode, resume_payload = result
    return uploaded_file, run_mode, resume_payload


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
    chrome_placeholder = st.empty()
    settings_clicked = render_shell_chrome(chrome_placeholder)
    result = st.session_state.get("result")
    question_waiting = result_waiting_on_question(result)
    if settings_clicked and not question_waiting:
        render_settings_dialog(base_settings)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    work_placeholder = st.empty()
    uploaded_file, run_mode, resume_payload = unpack_work_area_result(
        render_work_area(work_placeholder, settings)
    )
    log_placeholder = None
    render_logs(log_placeholder)

    if resume_payload is not None:
        current_job_id = st.session_state.get("current_job_id")
        if current_job_id is None:
            st.session_state["logs"].append("[ERROR] Workbook entry cannot resume because the job id is missing.")
        else:
            current_result = st.session_state.get("result")
            pending_question = getattr(current_result, "pending_question", None)
            if isinstance(current_result, ImportArtifacts) and current_result.job_id == current_job_id:
                st.session_state["result"] = current_result.model_copy(update={"pending_question": None})
            st.session_state[QUESTION_ACTIVE_RESUME_STATE_KEY] = current_job_id
            st.session_state[QUESTION_PENDING_RESUME_STATE_KEY] = {
                "job_id": current_job_id,
                "answer": resume_payload[0],
                "source": resume_payload[1],
            }
            st.session_state["is_running"] = True
            st.session_state["active_stage"] = Stage.DATA_ENTRY.value
            st.session_state["last_status"] = build_status_snapshot(
                Stage.DATA_ENTRY,
                message="Resuming workbook entry.",
                severity=Severity.INFO,
            )
            trace = st.session_state.get("agent_trace")
            if isinstance(trace, dict):
                trace.update(
                    {
                        "status": "running",
                        "current_sheet": (
                            pending_question.sheet_name
                            if isinstance(pending_question, AgentQuestion)
                            else trace.get("current_sheet")
                        ),
                        "message": "Letting the agent resolve the outstanding section and continue workbook entry.",
                        "started_at_ms": int(time.time() * 1000),
                        "retry_until_ms": None,
                        "live_summary": None,
                        "live_summary_updated_at_ms": None,
                        "last_rendered_summary": None,
                    }
                )
            st.rerun()
            return

    pending_resume_job_id, pending_resume_answer, pending_resume_source = pop_entry_question_resume()
    if (
        pending_resume_job_id is not None
        and pending_resume_answer is not None
        and pending_resume_source is not None
    ):
        runner = JobRunner(settings)
        consume_runner_events(
            events=runner.resume_job(
                pending_resume_job_id,
                pending_resume_answer,
                source=pending_resume_source,
            ),
            settings=settings,
            chrome_placeholder=chrome_placeholder,
            log_placeholder=log_placeholder,
            work_placeholder=work_placeholder,
        )
        return

    if run_mode == "upload" and uploaded_file is not None:
        uploaded_bytes = uploaded_file.getvalue()
        try:
            page_previews = render_pdf_previews(uploaded_bytes, settings.max_pages)
        except Exception:
            page_previews = {}
        reset_run_state(settings.ocr_parallel_workers, uploaded_file.name, page_previews)
        render_shell_chrome(chrome_placeholder)
        render_logs(log_placeholder)
        render_work_area(work_placeholder, settings)

        runner = JobRunner(settings)
        consume_runner_events(
            events=(
                runner.start_job(uploaded_bytes, uploaded_file.name)
                if hasattr(runner, "start_job")
                else runner.run(uploaded_bytes, uploaded_file.name)
            ),
            settings=settings,
            chrome_placeholder=chrome_placeholder,
            log_placeholder=log_placeholder,
            work_placeholder=work_placeholder,
        )

    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
