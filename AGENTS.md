# AGENTS.md

## Streamlit HTML Rendering Guardrail

- Follow the Streamlit docs for `st.markdown` and `st.html`: if a block is pure HTML/CSS, prefer `st.html(...)` over `st.markdown(..., unsafe_allow_html=True)`.
- Use `st.markdown` for Markdown-first content. Do not use it for nested component markup such as cards, grids, meters, or repeated `<div>` trees, because Streamlit can escape or coerce that content into literal text/code-like output.
- Keep raw HTML rendering behind a small helper such as `render_html_block()` so the app has one consistent fallback path when `st.html` is unavailable.
- Escape dynamic text before interpolating it into HTML. `st.html` is not iframed, and JavaScript is ignored by default unless `unsafe_allow_javascript=True`, so only enable script execution deliberately and never with untrusted input.
- When adding or refactoring an HTML-heavy surface, add a regression test that asserts the `st.html` path is used when available and that the fallback path still renders the same markup shape.

References:
- [Streamlit `st.html` docs](https://docs.streamlit.io/develop/api-reference/text/st.html)
- [Streamlit `st.markdown` docs](https://docs.streamlit.io/develop/api-reference/text/st.markdown)
