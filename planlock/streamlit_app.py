from __future__ import annotations

from dataclasses import replace
import html
import json
import time
from textwrap import dedent

import streamlit as st
import streamlit.components.v1 as components

from planlock.config import (
    DEFAULT_LLM_PROVIDER,
    LLM_PROVIDER_OPTIONS,
    Settings,
    locked_model_for_provider,
    provider_display_name,
)
from planlock.job_runner import JobRunner
from planlock.models import ImportArtifacts, RunEvent, Severity, Stage
from planlock.pdf_renderer import render_pdf_previews
from planlock.template_entry_agent import load_entry_state


st.set_page_config(
    page_title="PlanLock",
    layout="wide",
)


COOLDOWN_HANDOFF_ANIMATION_MS = 420
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
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

        :root {
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
            font-family: "IBM Plex Sans", sans-serif;
            color: var(--ink);
        }

        .stApp,
        .stApp * {
            font-family: "IBM Plex Sans", sans-serif !important;
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
            padding-top: 1.75rem;
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
            gap: 1.1rem;
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
            font-size: clamp(0.95rem, 4vw, 4rem);
            line-height: 0.96;
            letter-spacing: -0.06em;
            font-weight: 700;
            max-width: none;
            white-space: nowrap;
            color: var(--ink) !important;
            text-shadow: none;
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

        .workbook-retry-banner::after {
            content: "";
            position: absolute;
            inset: auto -2rem -2rem auto;
            width: 8rem;
            height: 8rem;
            background: radial-gradient(circle, rgba(166, 100, 32, 0.14) 0%, rgba(166, 100, 32, 0) 70%);
            pointer-events: none;
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

        .question-card {
            background:
                linear-gradient(180deg, rgba(255, 251, 244, 0.98) 0%, rgba(246, 238, 225, 0.98) 100%);
        }

        .question-kicker {
            font-size: 0.7rem;
            text-transform: uppercase;
            letter-spacing: 0.18em;
            color: var(--signal-strong);
            font-weight: 700;
        }

        .question-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem;
            flex-wrap: wrap;
        }

        .question-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.32rem 0.6rem;
            border-radius: 999px;
            border: 1px solid rgba(201, 149, 50, 0.3);
            background: rgba(245, 230, 196, 0.88);
            color: var(--accent);
            font-size: 0.7rem;
            font-weight: 700;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            white-space: nowrap;
        }

        .question-title {
            margin: 0.55rem 0 0;
            font-size: clamp(1.7rem, 2.4vw, 2.2rem);
            line-height: 0.96;
            letter-spacing: -0.04em;
            color: var(--ink);
        }

        .question-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.8rem;
            margin-top: 1rem;
        }

        .question-meta {
            padding: 0.85rem 0.9rem;
            border-radius: var(--radius-lg);
            border: 1px solid var(--line);
            background: rgba(255, 251, 244, 0.92);
        }

        .question-meta-label {
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.15em;
            color: var(--muted);
            font-weight: 700;
        }

        .question-meta-value {
            display: block;
            margin-top: 0.42rem;
            font-size: 0.9rem;
            line-height: 1.45;
            color: var(--ink);
        }

        .question-actions-note {
            margin-top: 0.55rem;
            font-size: 0.78rem;
            line-height: 1.45;
            color: var(--muted);
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
            background: var(--surface) !important;
            box-shadow: var(--shadow) !important;
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
            border-color: rgba(201, 149, 50, 0.64) !important;
            box-shadow:
                inset 0 1px 0 rgba(255, 255, 255, 0.08),
                0 16px 26px rgba(23, 50, 77, 0.22) !important;
        }

        .stButton > button[kind="primary"] *,
        .stFormSubmitButton > button * {
            color: inherit !important;
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

        .section-card,
        .status-card,
        .stage-chip {
            animation: rise 240ms ease both;
        }

        .section-card:hover,
        .status-card:hover,
        .stage-chip:hover,
        .workbook-phase-card:hover,
        .question-meta:hover,
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
            .chunk-progress-block.active,
            .chunk-progress-block.active::after,
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

            .stage-rail {
                display: grid;
                grid-template-columns: 1fr;
            }

            .stage-chip,
            .stage-chip.active {
                flex: 1 1 auto;
            }

            .workbook-stage-header,
            .workbook-phase-grid,
            .question-grid {
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
    st.session_state.setdefault("current_job_id", None)
    st.session_state.setdefault("workbook_retry", None)
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


def build_workbook_retry_state() -> dict:
    return {
        "sheet_name": None,
        "message": None,
        "detail_message": None,
        "attempt_number": None,
        "max_attempts": None,
        "retry_delay_seconds": None,
        "retry_until_ms": None,
        "retry_reason": None,
    }


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


def reset_cached_run_state(pipe_total: int, source_filename: str, page_total: int) -> None:
    stage_progress = {stage.value: (0, 1) for stage in Stage}
    stage_progress[Stage.OCR.value] = (page_total, page_total or 1)
    st.session_state["logs"] = []
    st.session_state["stage_progress"] = stage_progress
    st.session_state["active_stage"] = Stage.DATA_ENTRY.value
    st.session_state["result"] = None
    st.session_state["last_status"] = build_status_snapshot(
        Stage.DATA_ENTRY,
        message="Loaded cached phase-one output. Preparing workbook entry.",
        severity=Severity.INFO,
    )
    pipeline = build_ocr_pipeline_state(pipe_total)
    pipeline["page_total"] = page_total
    pipeline["completed_pages"] = page_total
    st.session_state["ocr_pipeline"] = pipeline
    st.session_state["run_started"] = True
    st.session_state["is_running"] = True
    st.session_state["source_filename"] = source_filename
    st.session_state["page_previews"] = {}
    st.session_state["current_job_id"] = None
    st.session_state["workbook_retry"] = None


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
        return "Workbook validation"
    return "LangGraph sheet entry"


def build_status_snapshot(stage: Stage | str, *, message: str, severity: Severity) -> dict[str, object]:
    stage_enum = stage if isinstance(stage, Stage) else Stage(stage)
    return {
        "index": stage_index(stage_enum),
        "stage": display_stage_name(stage_enum),
        "message": message,
        "severity": severity,
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


def build_workbook_retry_markup(retry_state: dict[str, object]) -> str:
    retry_until_ms = retry_state.get("retry_until_ms")
    countdown_markup = build_timing_markup(
        format_seconds_label(float(retry_state.get("retry_delay_seconds") or 0.0)),
        countdown_target_ms=int(retry_until_ms) if retry_until_ms is not None else None,
    )
    attempt_number = retry_state.get("attempt_number")
    max_attempts = retry_state.get("max_attempts")
    attempt_copy = (
        f"Pass {attempt_number}/{max_attempts} is queued and will resume automatically."
        if attempt_number is not None and max_attempts is not None
        else "Workbook entry will resume automatically."
    )
    return dedent(
        f"""
        <div class="workbook-retry-banner">
            <div class="workbook-retry-notch"></div>
            <div class="workbook-retry-copy">
                <div class="workbook-retry-kicker">Throttle hold</div>
                <div class="workbook-retry-title">{current_provider_display_name()} rate limit detected</div>
                <div class="workbook-retry-body">
                    {html.escape(str(retry_state.get("sheet_name") or "Workbook entry"))} is paused. {html.escape(attempt_copy)}
                </div>
            </div>
            <div class="workbook-retry-timer-shell">
                <span class="workbook-retry-timer-label">Retry in</span>
                <div class="workbook-retry-timer">{countdown_markup}</div>
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
        save_clicked = save_col.form_submit_button("Save settings", use_container_width=True)
        reset_clicked = reset_col.form_submit_button("Use defaults", use_container_width=True)

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


def render_masthead() -> bool:
    title_col, action_col = st.columns([0.82, 0.18], gap="small")
    with title_col:
        st.markdown(
            f"""
            <div class="masthead">
                <div>
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
                <div class="status-message">Upload a planner PDF to start the two-step intake.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    tone = status_tone(status["severity"])
    message_text = str(status["message"])
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
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_stage_progress() -> None:
    progress = st.session_state["stage_progress"]
    active_stage = workflow_stage(st.session_state["active_stage"])
    is_running = bool(st.session_state.get("is_running"))
    run_started = bool(st.session_state.get("run_started"))
    last_status = st.session_state.get("last_status")
    result = st.session_state.get("result")
    needs_input = bool(result is not None and result.pending_question is not None)

    cards: list[str] = []
    for index, stage in enumerate(WORKFLOW_STAGES, start=1):
        member_stages = (stage,) if stage == Stage.OCR else (Stage.DATA_ENTRY, Stage.FINANCIAL_CALCULATIONS)
        is_complete = all(
            progress.get(member_stage.value, (0, 1))[0] >= progress.get(member_stage.value, (0, 1))[1]
            and progress.get(member_stage.value, (0, 1))[1] > 0
            for member_stage in member_stages
        )
        is_active = stage == active_stage
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
        st.session_state["last_status"] = build_status_snapshot(
            event.stage,
            message=event.message,
            severity=event.severity,
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
        if event.phase == "retry":
            now_ms = int(time.time() * 1000)
            retry_state = build_workbook_retry_state()
            retry_state.update(
                {
                    "sheet_name": (
                        str(event.message).split(" while filling ", 1)[1].split(".", 1)[0]
                        if " while filling " in str(event.message)
                        else display_stage_name(event.stage)
                    ),
                    "message": event.message,
                    "detail_message": event.detail_message,
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
        st.session_state["logs"].append(f"{prefix} {display_stage_name(event.stage)}: {event.message}")
    if event.artifacts is not None:
        st.session_state["result"] = event.artifacts
        st.session_state["current_job_id"] = event.artifacts.job_id


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

    st.markdown(
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
        ).strip(),
        unsafe_allow_html=True,
    )


def render_workbook_stage() -> None:
    progress = st.session_state["stage_progress"]
    active_stage = Stage(st.session_state["active_stage"])
    is_running = bool(st.session_state.get("is_running"))
    source_filename = st.session_state.get("source_filename") or "Current upload"
    last_status = st.session_state.get("last_status")
    result = st.session_state.get("result")
    needs_input = bool(result is not None and result.pending_question is not None)
    workbook_retry = st.session_state.get("workbook_retry")

    mapping_completed, mapping_total = progress.get(Stage.DATA_ENTRY.value, (0, 1))
    checks_completed, checks_total = progress.get(Stage.FINANCIAL_CALCULATIONS.value, (0, 1))
    checks_started = (
        active_stage == Stage.FINANCIAL_CALCULATIONS
        or checks_completed > 0
        or checks_total > 1
    )

    def phase_markup(
        *,
        kicker: str,
        title: str,
        state_label: str,
        copy: str,
        completed: int,
        total: int,
        tone: str,
    ) -> str:
        classes = ["workbook-phase-card", tone]
        if total > 0:
            ratio = min(max(completed / total, 0), 1)
        else:
            ratio = 0.0
        return dedent(
            f"""
            <div class="{' '.join(classes)}">
                <div class="workbook-phase-kicker">{html.escape(kicker)}</div>
                <div class="workbook-phase-title">{html.escape(title)}</div>
                <div class="workbook-phase-state">{html.escape(state_label)}</div>
                <div class="workbook-phase-meter">
                    <span class="workbook-phase-meter-fill" style="width: {ratio * 100:.1f}%"></span>
                </div>
                <div class="workbook-phase-copy">{html.escape(copy)}</div>
            </div>
            """
        ).strip()

    mapping_done = mapping_completed >= mapping_total and mapping_total > 0
    if needs_input and active_stage == Stage.DATA_ENTRY:
        mapping_state = "Needs input"
        mapping_tone = "attention"
        mapping_copy = "Waiting on planner input before the LangGraph agent can continue."
    elif mapping_done:
        mapping_state = "Complete"
        mapping_tone = "complete"
        mapping_copy = f"{mapping_completed}/{mapping_total} sheets completed."
    elif active_stage == Stage.DATA_ENTRY and is_running:
        mapping_state = "Running"
        mapping_tone = "active"
        mapping_copy = f"{mapping_completed}/{mapping_total} sheets completed so far."
    else:
        mapping_state = "Queued"
        mapping_tone = "pending"
        mapping_copy = f"{mapping_completed}/{mapping_total} sheets prepared."

    checks_done = checks_completed >= checks_total and checks_total > 0
    if checks_done:
        checks_state = "Complete"
        checks_tone = "complete"
        checks_copy = f"{checks_completed}/{checks_total} validation checks complete."
    elif active_stage == Stage.FINANCIAL_CALCULATIONS and is_running:
        checks_state = "Running"
        checks_tone = "active"
        checks_copy = f"{checks_completed}/{checks_total} validation checks complete so far."
    elif checks_started:
        checks_state = "Queued"
        checks_tone = "pending"
        checks_copy = f"{checks_completed}/{checks_total} validation checks complete."
    else:
        checks_state = "Queued"
        checks_tone = "pending"
        checks_copy = "Queued after sheet entry completes."

    validation_metric = f"{checks_completed}/{checks_total}" if checks_started or checks_done else "Queued"
    current_phase = display_workbook_phase_name(active_stage)
    workbook_message = (
        str(last_status["message"])
        if last_status is not None and workflow_stage(active_stage) == Stage.DATA_ENTRY
        else "The workbook stage keeps the LangGraph agent and deterministic checks in one surface."
    )
    workbook_retry_markup = ""
    if isinstance(workbook_retry, dict) and workbook_retry.get("retry_reason") == "rate_limit":
        workbook_retry_markup = build_workbook_retry_markup(workbook_retry)

    st.markdown(
        dedent(
            f"""
            <div class="section-card live-card workbook-stage-card">
                <div class="workbook-stage-header">
                    <div>
                        <h3 class="section-title">Workbook entry</h3>
                        <div class="section-copy">{html.escape(source_filename)}</div>
                    </div>
                    <div class="workbook-stage-pill">{html.escape(current_phase)}</div>
                </div>
                <div class="metric-row">
                    <div class="metric">
                        <span>Sheets complete</span>
                        <strong>{mapping_completed}/{mapping_total}</strong>
                    </div>
                    <div class="metric">
                        <span>Validation checks</span>
                        <strong>{html.escape(validation_metric)}</strong>
                    </div>
                    <div class="metric">
                        <span>Current phase</span>
                        <strong>{html.escape(current_phase)}</strong>
                    </div>
                </div>
                {workbook_retry_markup}
                <div class="workbook-phase-grid">
                    {phase_markup(
                        kicker="Phase one",
                        title="LangGraph sheet entry",
                        state_label=mapping_state,
                        copy=mapping_copy,
                        completed=mapping_completed,
                        total=mapping_total,
                        tone=mapping_tone,
                    )}
                    {phase_markup(
                        kicker="Phase two",
                        title="Workbook validation",
                        state_label=checks_state,
                        copy=checks_copy,
                        completed=checks_completed,
                        total=checks_total,
                        tone=checks_tone,
                    )}
                </div>
                <div class="section-copy workbook-stage-message">{html.escape(workbook_message)}</div>
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
    cache_path,
) -> tuple[object | None, str | None]:
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
    cache_available = cache_path.exists()
    run_col, cache_col = st.columns(2, gap="small")
    run_clicked = run_col.button(
        "Build workbook",
        key=f"run_import_{context_key}",
        type="primary",
        use_container_width=True,
        disabled=uploaded_file is None,
    )
    run_from_cache_clicked = cache_col.button(
        "Run from cache",
        key=f"run_from_cache_{context_key}",
        type="secondary",
        use_container_width=True,
        disabled=not cache_available,
        help=(
            f"Use {cache_path.name} from the last completed document review."
            if cache_available
            else f"{cache_path.name} is not available yet."
        ),
    )
    st.markdown("</div>", unsafe_allow_html=True)
    if run_clicked:
        return uploaded_file, "upload"
    if run_from_cache_clicked:
        return uploaded_file, "cache"
    return uploaded_file, None


def render_stage_focus() -> None:
    stage = Stage(st.session_state["active_stage"])

    if stage == Stage.OCR:
        render_ocr_parallel()
        return

    render_workbook_stage()


def render_entry_question(result: ImportArtifacts) -> tuple[str | None, str | None, bool]:
    if result.pending_question is None:
        return None, None, False

    entry_state = None
    if result.entry_state_path is not None and result.entry_state_path.exists():
        entry_state = load_entry_state(result.entry_state_path)

    completed = entry_state.current_sheet_index if entry_state is not None else 0
    total = len(entry_state.sheet_order) if entry_state is not None else 0
    remaining = max(0, total - completed)
    last_summary = None
    if entry_state is not None:
        last_summary = next(
            (summary for summary in entry_state.sheet_summaries if summary.sheet_name == result.pending_question.sheet_name),
            None,
        )

    render_workbook_stage()
    st.markdown("<div style='height:0.85rem'></div>", unsafe_allow_html=True)
    with st.form(f"entry_question_{result.job_id}", border=False):
        summary_message = last_summary.message if last_summary and last_summary.message else "Waiting on user input."
        rereview_chip = ""
        if result.pending_question.pdf_rereviewed:
            rereview_chip = '<div class="question-chip">Raw PDF re-reviewed</div>'
        st.markdown(
            f"""
            <div class="section-card question-card">
                <div class="question-header">
                    <div class="question-kicker">Question for you</div>
                    {rereview_chip}
                </div>
                <h3 class="question-title">{html.escape(result.pending_question.prompt)}</h3>
                <div class="section-copy">{html.escape(result.pending_question.rationale)}</div>
                <div class="question-grid">
                    <div class="question-meta">
                        <div class="question-meta-label">Current sheet</div>
                        <span class="question-meta-value">{html.escape(result.pending_question.sheet_name)}</span>
                    </div>
                    <div class="question-meta">
                        <div class="question-meta-label">Progress</div>
                        <span class="question-meta-value">{completed}/{total} completed • {remaining} remaining</span>
                    </div>
                    <div class="question-meta">
                        <div class="question-meta-label">Affected targets</div>
                        <span class="question-meta-value">{html.escape(', '.join(result.pending_question.affected_targets) or 'Not specified')}</span>
                    </div>
                    <div class="question-meta">
                        <div class="question-meta-label">Last write summary</div>
                        <span class="question-meta-value">{html.escape(summary_message)}</span>
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )
        selected_option_value = None
        if result.pending_question.options:
            option_labels = [option.label for option in result.pending_question.options]
            option_values = {option.label: option.value for option in result.pending_question.options}
            option_captions = [option.description or "Recommended from the current document context." for option in result.pending_question.options]
            selected_option_label = st.radio(
                "Choose an answer",
                options=option_labels,
                index=0,
                key=f"question_options_{result.job_id}",
                captions=option_captions,
                width="stretch",
            )
            selected_option_value = option_values[selected_option_label]
        free_text = ""
        if result.pending_question.allow_free_text:
            free_text = st.text_input(
                "Or write your own answer",
                key=f"question_free_text_{result.job_id}",
            )

        submit_col, delegate_col = st.columns([0.64, 0.36], gap="small")
        submitted = submit_col.form_submit_button(
            "Submit answer and continue",
            type="primary",
            use_container_width=True,
        )
        delegated = delegate_col.form_submit_button(
            "Figure it out",
            type="secondary",
            help="Let the agent make the best supported choice from the document.",
            use_container_width=True,
        )
        st.markdown(
            '<div class="question-actions-note">Use <strong>Figure it out</strong> when you want the agent to make the call and keep moving.</div></div>',
            unsafe_allow_html=True,
        )

    if delegated:
        return "", "agent", True
    if not submitted:
        return None, None, False
    if free_text.strip():
        return free_text.strip(), "free_text", True
    if selected_option_value is not None:
        return selected_option_value, "option", True
    return None, None, False


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
            return None, False, None

        if result is not None and result.pending_question is not None and result.review_report is None:
            answer, source, submitted = render_entry_question(result)
            if submitted and answer is not None and source is not None:
                resume_payload = (answer, source)
            return None, False, resume_payload

        if result is not None and result.review_report is not None:
            render_result(result)
            st.markdown("<div style='height:0.85rem'></div>", unsafe_allow_html=True)
            with st.expander("Import another PDF", expanded=False):
                uploaded_file, run_mode = render_upload_panel(
                    context_key="next",
                    title="Process another PDF",
                    copy=f"Current settings: {settings.ocr_parallel_workers} parallel lanes.",
                    cache_path=settings.debug_cache_path,
                )
            return uploaded_file, run_mode, None

        if run_started and status is not None and status["severity"] == Severity.ERROR:
            uploaded_file, run_mode = render_upload_panel(
                context_key="retry",
                title="Run failed",
                copy=status["message"],
                cache_path=settings.debug_cache_path,
            )
            return uploaded_file, run_mode, None

        uploaded_file, run_mode = render_upload_panel(
            context_key="primary",
            title="Upload PDF",
            copy=(
                "Use Settings to change the number of parallel lanes. "
                f"Only provider credentials stay in .env. Use {settings.debug_cache_path.name} to rerun phase two without redoing document review."
            ),
            cache_path=settings.debug_cache_path,
        )
    return uploaded_file, run_mode, None


def consume_runner_events(
    *,
    events,
    settings: Settings,
    status_placeholder,
    progress_placeholder,
    log_placeholder,
    work_placeholder,
) -> None:
    try:
        for event in events:
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
        st.session_state["last_status"] = build_status_snapshot(
            Stage(st.session_state["active_stage"]),
            message=f"Run failed: {exc}",
            severity=Severity.ERROR,
        )
    finally:
        st.session_state["is_running"] = False
        render_status(status_placeholder)
        progress_placeholder.empty()
        with progress_placeholder.container():
            render_stage_progress()
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
            answer, source = resume_payload
            st.session_state["is_running"] = True
            st.session_state["active_stage"] = Stage.DATA_ENTRY.value
            st.session_state["last_status"] = build_status_snapshot(
                Stage.DATA_ENTRY,
                message="Resuming workbook entry.",
                severity=Severity.INFO,
            )
            render_status(status_placeholder)
            progress_placeholder.empty()
            with progress_placeholder.container():
                render_stage_progress()
            render_logs(log_placeholder)
            render_work_area(work_placeholder, settings)

            runner = JobRunner(settings)
            consume_runner_events(
                events=runner.resume_job(current_job_id, answer, source=source),
                settings=settings,
                status_placeholder=status_placeholder,
                progress_placeholder=progress_placeholder,
                log_placeholder=log_placeholder,
                work_placeholder=work_placeholder,
            )

    if run_mode == "upload" and uploaded_file is not None:
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
        consume_runner_events(
            events=(
                runner.start_job(uploaded_bytes, uploaded_file.name)
                if hasattr(runner, "start_job")
                else runner.run(uploaded_bytes, uploaded_file.name)
            ),
            settings=settings,
            status_placeholder=status_placeholder,
            progress_placeholder=progress_placeholder,
            log_placeholder=log_placeholder,
            work_placeholder=work_placeholder,
        )
    elif run_mode == "cache":
        runner = JobRunner(settings)
        try:
            cache = runner.load_phase_one_cache(settings.debug_cache_path)
        except Exception as exc:
            st.session_state["logs"].append(f"[ERROR] Document review: Cache run failed: {exc}")
            st.session_state["last_status"] = build_status_snapshot(
                Stage.OCR,
                message=f"Cache run failed: {exc}",
                severity=Severity.ERROR,
            )
            st.session_state["is_running"] = False
            render_status(status_placeholder)
            progress_placeholder.empty()
            with progress_placeholder.container():
                render_stage_progress()
            render_logs(log_placeholder)
            render_work_area(work_placeholder, settings)
        else:
            reset_cached_run_state(
                settings.ocr_parallel_workers,
                cache.source_filename,
                len(cache.ocr_results),
            )
            render_status(status_placeholder)
            progress_placeholder.empty()
            with progress_placeholder.container():
                render_stage_progress()
            render_logs(log_placeholder)
            render_work_area(work_placeholder, settings)

            consume_runner_events(
                events=runner.start_job_from_cache(settings.debug_cache_path, cache=cache),
                settings=settings,
                status_placeholder=status_placeholder,
                progress_placeholder=progress_placeholder,
                log_placeholder=log_placeholder,
                work_placeholder=work_placeholder,
            )

    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
