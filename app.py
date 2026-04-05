"""
Research Paper Writing Agent
A persistent, style-aware academic writing assistant built with Streamlit + Anthropic Claude.
Run: streamlit run app.py
Requires: ANTHROPIC_API_KEY environment variable
"""

import io
import json
import re
import uuid
from datetime import datetime
from pathlib import Path

import anthropic
import requests
import streamlit as st
from bs4 import BeautifulSoup

try:
    import pypdf
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# ==============================================================================
# PAPER_CONTEXT
# Fill in these fields to give the agent context about your specific paper.
# When all fields are empty, the agent operates as a general academic writing assistant.
# ==============================================================================
PAPER_CONTEXT = {
    "paper_title": "",           # e.g., "Attention Is All You Need"
    "research_question": "",     # The main research question or hypothesis
    "models_methods": "",        # Models or methods used
    "dataset_description": "",   # Dataset name, size, and key properties
    "key_findings": "",          # Primary results and contributions
    "target_journal": "",        # Target venue (e.g., "NeurIPS 2025")
    "co_authors": "",            # Author names, separated by commas
}
# ==============================================================================

SESSIONS_FILE = "sessions.json"
MODEL = "claude-sonnet-4-0"
SECTIONS = ["Introduction", "Methods", "Results", "Discussion", "Conclusion", "Abstract", "Title"]
# Max characters of document content injected into system prompt per document
DOC_CONTENT_LIMIT = 12_000


# ── Document fetching ──────────────────────────────────────────────────────────

def extract_pdf_text(pdf_bytes: bytes) -> str:
    if not PDF_SUPPORT:
        raise RuntimeError("pypdf not installed")
    reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
    pages = [page.extract_text() or "" for page in reader.pages]
    return "\n\n".join(p for p in pages if p.strip())


def fetch_url_content(url: str) -> tuple[str, str]:
    """
    Fetch content from a URL. Returns (document_name, text_content).
    Handles Google Docs, Google Drive files, direct PDF URLs, and regular web pages.
    """
    url = url.strip()

    # Google Docs shared link → export as plain text
    gdoc_match = re.search(r"docs\.google\.com/document/d/([^/?\s]+)", url)
    if gdoc_match:
        doc_id = gdoc_match.group(1)
        export_url = f"https://docs.google.com/document/d/{doc_id}/export?format=txt"
        resp = requests.get(export_url, timeout=30)
        resp.raise_for_status()
        return f"Google Doc ({doc_id[:12]})", resp.text.strip()

    # Google Drive file link → try to download
    gdrive_match = re.search(r"drive\.google\.com/file/d/([^/?\s]+)", url)
    if gdrive_match:
        file_id = gdrive_match.group(1)
        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        resp = requests.get(download_url, timeout=30, allow_redirects=True)
        resp.raise_for_status()
        ctype = resp.headers.get("content-type", "")
        if "pdf" in ctype:
            return f"Google Drive PDF ({file_id[:12]})", extract_pdf_text(resp.content)
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style"]):
            tag.decompose()
        return f"Google Drive file ({file_id[:12]})", soup.get_text("\n", strip=True)

    # Fetch the URL
    headers = {"User-Agent": "Mozilla/5.0 (compatible; PaperAgent/1.0)"}
    resp = requests.get(url, timeout=30, headers=headers)
    resp.raise_for_status()
    ctype = resp.headers.get("content-type", "")

    # Direct PDF URL
    if "pdf" in ctype or url.lower().endswith(".pdf"):
        name = url.rstrip("/").split("/")[-1] or "document.pdf"
        return name, extract_pdf_text(resp.content)

    # Regular web page
    soup = BeautifulSoup(resp.text, "html.parser")
    for tag in soup(["script", "style", "nav", "footer", "header", "aside"]):
        tag.decompose()
    text = soup.get_text("\n", strip=True)
    text = re.sub(r"\n{3,}", "\n\n", text)
    title_tag = soup.find("title")
    name = title_tag.get_text(strip=True)[:80] if title_tag else url.split("/")[-1] or url
    return name, text


# ── System prompt ──────────────────────────────────────────────────────────────

def build_system_prompt(session: dict | None = None) -> str:
    prompt = """\
You are a research paper writing assistant. Your job is to draft and revise sections of \
academic research papers. You write in a specific style—internalize it and apply it to \
everything you produce.

WRITING STYLE — MATCH THIS EXACTLY:

Characteristics to apply:
- Direct, confident sentences. No fluff, no over-hedging.
- Thoughts flow naturally from one to the next without heavy transitional phrases.
- Comfortable mixing a high-level claim with a specific, grounded example right after it.
- Makes the point and moves on. No over-explaining.
- Occasional short punchy sentences after longer ones for emphasis.
- Acknowledges tradeoffs and limitations honestly rather than glossing over them.
- When listing things, groups them in a natural run rather than bullet-pointing everything.

Patterns to NEVER use — these read as AI-generated:
- "It is important to note that..."
- "Furthermore," / "Moreover," / "Additionally," / "Notably," as sentence starters
- "It is worth mentioning..."
- Restating the conclusion at the end of every paragraph
- Overly balanced constructions: "while X, it is also true that Y"
- Passive voice overuse
- Ending sections with a summary sentence that just repeats what was said

This is academic writing for a research paper. Be scholarly, but stay direct.\
"""

    # Static paper context (from PAPER_CONTEXT at top of file)
    ctx = PAPER_CONTEXT
    if any(v.strip() for v in ctx.values()):
        prompt += "\n\nPAPER CONTEXT:\n"
        if ctx["paper_title"]:       prompt += f"Title: {ctx['paper_title']}\n"
        if ctx["research_question"]: prompt += f"Research question: {ctx['research_question']}\n"
        if ctx["models_methods"]:    prompt += f"Models/methods: {ctx['models_methods']}\n"
        if ctx["dataset_description"]: prompt += f"Dataset: {ctx['dataset_description']}\n"
        if ctx["key_findings"]:      prompt += f"Key findings: {ctx['key_findings']}\n"
        if ctx["target_journal"]:    prompt += f"Target venue: {ctx['target_journal']}\n"
        if ctx["co_authors"]:        prompt += f"Co-authors: {ctx['co_authors']}\n"

    # Attached documents (stored in session)
    if session and session.get("documents"):
        prompt += "\n\nREFERENCE DOCUMENTS (use these as source material when writing):\n"
        for i, doc in enumerate(session["documents"], 1):
            content = doc["content"]
            truncated = len(content) > DOC_CONTENT_LIMIT
            if truncated:
                content = content[:DOC_CONTENT_LIMIT]
            prompt += f"\n--- Document {i}: {doc['name']} ---\n{content}\n"
            if truncated:
                prompt += "[...document truncated for length]\n"

    return prompt


# ── Session persistence ────────────────────────────────────────────────────────

def load_sessions() -> dict:
    path = Path(SESSIONS_FILE)
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}


def save_sessions(sessions: dict) -> None:
    with open(SESSIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(sessions, f, indent=2, ensure_ascii=False)


def new_session(paper_title: str, section: str) -> tuple[str, dict]:
    session_id = str(uuid.uuid4())[:8]
    session = {
        "session_id": session_id,
        "timestamp": datetime.now().isoformat(),
        "last_modified": datetime.now().isoformat(),
        "paper_title": paper_title,
        "section": section,
        "conversation_history": [],
        "current_draft": "",
        "drafts": [],
        "iteration_notes": [],
        "documents": [],   # [{"name": str, "source": str, "content": str, "added_at": str}]
    }
    return session_id, session


def touch_session(session: dict) -> None:
    """Update last_modified timestamp."""
    session["last_modified"] = datetime.now().isoformat()


# ── Claude API call ────────────────────────────────────────────────────────────

def call_claude(conversation_history: list, session: dict, client: anthropic.Anthropic) -> str:
    response = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        system=build_system_prompt(session),
        messages=conversation_history,
    )
    return response.content[0].text


# ── UI: Sidebar ────────────────────────────────────────────────────────────────

def render_sidebar(sessions: dict) -> None:
    with st.sidebar:
        st.title("Paper Agent")
        if st.button("＋ New Session", use_container_width=True, type="primary"):
            st.session_state.current_session_id = None
            st.rerun()

        st.divider()
        st.subheader("Sessions")

        if not sessions:
            st.caption("No sessions yet.")
            return

        for sid, s in sorted(sessions.items(), key=lambda x: x[1].get("last_modified", x[1]["timestamp"]), reverse=True):
            title = s["paper_title"]
            title_short = (title[:22] + "…") if len(title) > 25 else title
            date = s.get("last_modified", s["timestamp"])[:10]
            n_versions = len(s.get("drafts", []))
            n_docs = len(s.get("documents", []))
            meta_parts = [s["section"]]
            if n_versions:
                meta_parts.append(f"{n_versions}v")
            if n_docs:
                meta_parts.append(f"{n_docs} doc{'s' if n_docs > 1 else ''}")
            meta = " · ".join(meta_parts)
            label = f"{title_short}\n{meta} · {date}"
            is_active = st.session_state.get("current_session_id") == sid
            if st.button(label, key=f"btn_{sid}", use_container_width=True,
                         type="primary" if is_active else "secondary"):
                st.session_state.current_session_id = sid
                st.rerun()


# ── UI: New session form ───────────────────────────────────────────────────────

def render_new_session_form() -> None:
    st.title("Research Paper Writing Agent")
    st.markdown(
        "A style-aware writing partner that drafts and iterates on research paper sections "
        "in your voice. Sessions are saved automatically — pick up exactly where you left off."
    )
    st.divider()

    with st.form("new_session_form"):
        paper_title = st.text_input(
            "Paper Title",
            placeholder="e.g., Scaling Laws for Neural Language Models",
        )
        section = st.selectbox("Section to Draft", SECTIONS)
        submitted = st.form_submit_button("Start Session →", type="primary")

        if submitted:
            if not paper_title.strip():
                st.error("Enter a paper title to continue.")
            else:
                sid, session = new_session(paper_title.strip(), section)
                st.session_state.sessions[sid] = session
                save_sessions(st.session_state.sessions)
                st.session_state.current_session_id = sid
                st.rerun()


# ── UI: Document panel ─────────────────────────────────────────────────────────

def render_document_panel(session: dict) -> None:
    """Render the document attachment section."""
    docs = session.setdefault("documents", [])
    n = len(docs)
    label = f"Reference Documents ({n} attached)" if n else "Reference Documents"

    with st.expander(label, expanded=(n > 0)):
        # Show attached docs
        if docs:
            for i, doc in enumerate(docs):
                col_name, col_remove = st.columns([5, 1])
                with col_name:
                    chars = len(doc["content"])
                    st.markdown(f"**{doc['name']}** — {chars:,} chars · added {doc['added_at'][:10]}")
                with col_remove:
                    if st.button("✕", key=f"remove_doc_{i}", help="Remove this document"):
                        docs.pop(i)
                        touch_session(session)
                        save_sessions(st.session_state.sessions)
                        st.rerun()
            st.divider()

        # Add by URL (Google Docs, any web page, direct PDF link)
        st.markdown("**Add from URL** (Google Docs, Google Drive, any web page or PDF link)")
        col_url, col_add_url = st.columns([4, 1])
        with col_url:
            url_input = st.text_input(
                "url_input",
                placeholder="https://docs.google.com/document/d/... or any URL",
                label_visibility="collapsed",
                key="doc_url_input",
            )
        with col_add_url:
            if st.button("Add →", key="add_url_btn", use_container_width=True):
                if not url_input.strip():
                    st.error("Paste a URL first.")
                else:
                    with st.spinner("Fetching…"):
                        try:
                            name, content = fetch_url_content(url_input.strip())
                            if not content.strip():
                                st.error("Could not extract text from that URL.")
                            else:
                                docs.append({
                                    "name": name,
                                    "source": url_input.strip(),
                                    "content": content,
                                    "added_at": datetime.now().isoformat(),
                                })
                                touch_session(session)
                                save_sessions(st.session_state.sessions)
                                st.success(f"Added: {name} ({len(content):,} chars)")
                                st.rerun()
                        except Exception as e:
                            st.error(f"Failed to fetch URL: {e}")

        # Add by file upload (PDF)
        st.markdown("**Upload a file** (PDF)")
        uploaded_file = st.file_uploader(
            "upload_pdf",
            type=["pdf"],
            label_visibility="collapsed",
            key="doc_file_uploader",
        )
        if uploaded_file is not None:
            # Check if already added (by name)
            already_added = any(d["name"] == uploaded_file.name for d in docs)
            if already_added:
                st.caption(f"✓ {uploaded_file.name} is already attached.")
            else:
                if st.button(f"Add \"{uploaded_file.name}\"", key="confirm_upload_btn"):
                    with st.spinner("Extracting text…"):
                        try:
                            if not PDF_SUPPORT:
                                st.error("pypdf not installed. Run: pip install pypdf")
                            else:
                                content = extract_pdf_text(uploaded_file.read())
                                if not content.strip():
                                    st.error("Could not extract text from this PDF.")
                                else:
                                    docs.append({
                                        "name": uploaded_file.name,
                                        "source": "uploaded file",
                                        "content": content,
                                        "added_at": datetime.now().isoformat(),
                                    })
                                    touch_session(session)
                                    save_sessions(st.session_state.sessions)
                                    st.success(f"Added: {uploaded_file.name} ({len(content):,} chars)")
                                    st.rerun()
                        except Exception as e:
                            st.error(f"Failed to read PDF: {e}")

        if docs:
            st.caption(
                "Documents are saved to this session and automatically included as context "
                "on every generation and revision."
            )


# ── UI: Active session ─────────────────────────────────────────────────────────

def render_active_session(session: dict, session_id: str, client: anthropic.Anthropic) -> None:
    has_draft = bool(session["current_draft"])
    current_version = session["drafts"][-1]["version"] if session["drafts"] else ""
    n_turns = len(session["conversation_history"]) // 2
    last_saved = session.get("last_modified", session["timestamp"])[:16].replace("T", " ")

    # ── Header ──
    st.markdown(f"## {session['paper_title']}")
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.markdown(f"**Section:** {session['section']}")
    with col2: st.markdown(f"**Session:** `{session_id}`")
    with col3: st.markdown(f"**Turns:** {n_turns}")
    with col4: st.markdown(f"**Saved:** {last_saved}")
    st.divider()

    # ── Two-column layout ──
    left, right = st.columns([1, 1], gap="large")

    with left:
        if not has_draft:
            st.subheader("Your Notes")
            notes = st.text_area(
                "notes_input",
                height=320,
                placeholder=(
                    "Paste raw notes, bullet points, or a rough outline:\n\n"
                    "• Main claim: model X outperforms baselines by 4.2% BLEU\n"
                    "• Why: larger receptive field, no recurrence bottleneck\n"
                    "• Caveat: tested on English-German only\n"
                    "• Include comparison table reference"
                ),
                label_visibility="collapsed",
            )
            if st.button("Generate Draft →", type="primary", use_container_width=True):
                if not notes.strip():
                    st.error("Paste some notes first.")
                else:
                    n_docs = len(session.get("documents", []))
                    doc_note = f" (using {n_docs} attached document{'s' if n_docs > 1 else ''})" if n_docs else ""
                    user_msg = (
                        f"Draft the {session['section']} section for my paper "
                        f"titled '{session['paper_title']}'.\n\nNotes:\n\n{notes}"
                    )
                    session["conversation_history"].append({"role": "user", "content": user_msg})
                    with st.spinner(f"Drafting{doc_note}…"):
                        draft = call_claude(session["conversation_history"], session, client)
                    session["conversation_history"].append({"role": "assistant", "content": draft})
                    session["current_draft"] = draft
                    session["drafts"].append({"version": "V1", "content": draft})
                    touch_session(session)
                    save_sessions(st.session_state.sessions)
                    st.rerun()

        else:
            st.subheader(f"Revise {current_version}")
            feedback = st.text_area(
                "feedback_input",
                height=280,
                placeholder=(
                    "Describe what to change:\n\n"
                    '"The second paragraph hedges too much—make it direct."\n'
                    '"Cut the last sentence, it repeats the point above."\n'
                    '"Add a concrete example to the third claim."\n'
                    '"The methods description is too vague—be more specific."'
                ),
                label_visibility="collapsed",
            )
            col_revise, col_reset = st.columns([2, 1])
            with col_revise:
                if st.button("Revise Draft →", type="primary", use_container_width=True):
                    if not feedback.strip():
                        st.error("Enter feedback first.")
                    else:
                        user_msg = f"Revise the draft based on this feedback:\n\n{feedback}"
                        session["conversation_history"].append({"role": "user", "content": user_msg})
                        with st.spinner("Revising…"):
                            revised = call_claude(session["conversation_history"], session, client)
                        session["conversation_history"].append({"role": "assistant", "content": revised})
                        session["current_draft"] = revised
                        session["iteration_notes"].append(feedback)
                        new_ver = f"V{len(session['drafts']) + 1}"
                        session["drafts"].append({"version": new_ver, "content": revised})
                        touch_session(session)
                        save_sessions(st.session_state.sessions)
                        st.rerun()
            with col_reset:
                if st.button("Start Over", use_container_width=True):
                    session["current_draft"] = ""
                    session["conversation_history"] = []
                    session["drafts"] = []
                    session["iteration_notes"] = []
                    touch_session(session)
                    save_sessions(st.session_state.sessions)
                    st.rerun()

    with right:
        if has_draft:
            st.subheader(f"Current Draft — {current_version}")
            st.code(session["current_draft"], language=None)
        else:
            st.subheader("Draft")
            st.info("Your draft will appear here after generation.")

    # ── Document panel ──
    st.divider()
    render_document_panel(session)

    # ── Iteration history ──
    if len(session["drafts"]) > 1:
        with st.expander(f"Iteration History — {len(session['drafts'])} versions"):
            for item in reversed(session["drafts"][:-1]):
                idx = session["drafts"].index(item)
                st.markdown(f"**{item['version']}**")
                if idx > 0 and idx - 1 < len(session["iteration_notes"]):
                    st.caption(f"_{session['iteration_notes'][idx - 1]}_")
                st.code(item["content"], language=None)
                st.divider()


# ── Entry point ────────────────────────────────────────────────────────────────

def main() -> None:
    st.set_page_config(
        page_title="Paper Writing Agent",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    try:
        client = anthropic.Anthropic()
    except Exception:
        st.error(
            "**ANTHROPIC_API_KEY not set.**\n\n"
            "In PowerShell:\n```\n$env:ANTHROPIC_API_KEY = 'sk-ant-...'\nstreamlit run app.py\n```"
        )
        st.stop()

    if "sessions" not in st.session_state:
        st.session_state.sessions = load_sessions()
    if "current_session_id" not in st.session_state:
        st.session_state.current_session_id = None

    render_sidebar(st.session_state.sessions)

    current_id = st.session_state.current_session_id
    if current_id is None or current_id not in st.session_state.sessions:
        render_new_session_form()
    else:
        render_active_session(st.session_state.sessions[current_id], current_id, client)


if __name__ == "__main__":
    main()
