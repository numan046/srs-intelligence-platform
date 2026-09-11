"""
================================================================================
 AI-Powered SRS Requirement Intelligence, Requirement Visualizer & Project
 Planning Platform
================================================================================
 Single-file Streamlit application.

 Run in Google Colab (see bottom-of-file instructions) or locally with:
     streamlit run app.py
================================================================================
"""

import os
import io
import json
import time
import uuid
import shutil
import re
import traceback
from datetime import datetime
from pathlib import Path

import streamlit as st
import bcrypt

import fitz  # PyMuPDF
import docx  # python-docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_openai import OpenAIEmbeddings

from openai import OpenAI

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
)


# ==============================================================================
# CONFIGURATION
# ==============================================================================

APP_TITLE = "SRS Intelligence Platform"
APP_ICON = "◆"

DATA_DIR = "data"
USERS_FILE = os.path.join(DATA_DIR, "users.json")
HISTORY_DIR = os.path.join(DATA_DIR, "history")
VECTORS_DIR = os.path.join(DATA_DIR, "vectors")
UPLOADS_DIR = os.path.join(DATA_DIR, "uploads")
REPORTS_DIR = os.path.join(DATA_DIR, "reports")

OPENAI_CHAT_MODEL = "gpt-4o-mini"
OPENAI_EMBEDDING_MODEL = "text-embedding-3-small"

MIN_PASSWORD_LENGTH = 8
CHUNK_SIZE = 1200
CHUNK_OVERLAP = 200

ANALYSIS_CATEGORIES = [
    "business_goals",
    "functional_requirements",
    "non_functional_requirements",
    "constraints",
    "stakeholders",
    "assumptions",
    "dependencies",
    "risks",
    "scope",
    "missing_requirements",
    "requirement_quality_issues",
    "suggested_improvements",
    "requirement_priority",
    "timeline_estimate",
    "team_recommendation",
    "project_complexity",
    "requirement_categories",
    "validation_report",
]


# ==============================================================================
# DIRECTORY / STORAGE BOOTSTRAP
# ==============================================================================

def ensure_directories():
    """Create the full data/ folder skeleton if it does not already exist."""
    for d in [DATA_DIR, HISTORY_DIR, VECTORS_DIR, UPLOADS_DIR, REPORTS_DIR]:
        os.makedirs(d, exist_ok=True)
    if not os.path.exists(USERS_FILE):
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump([], f, indent=2)


def user_history_dir(user_id):
    p = os.path.join(HISTORY_DIR, user_id)
    os.makedirs(p, exist_ok=True)
    return p


def user_vectors_dir(user_id):
    p = os.path.join(VECTORS_DIR, user_id)
    os.makedirs(p, exist_ok=True)
    return p


def user_uploads_dir(user_id):
    p = os.path.join(UPLOADS_DIR, user_id)
    os.makedirs(p, exist_ok=True)
    return p


def user_reports_dir(user_id):
    p = os.path.join(REPORTS_DIR, user_id)
    os.makedirs(p, exist_ok=True)
    return p


def safe_read_json(path, default):
    """Read a JSON file, returning `default` if it is missing, empty, or corrupt."""
    try:
        if not os.path.exists(path):
            return default
        with open(path, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return default
            return json.loads(content)
    except (json.JSONDecodeError, OSError):
        return default


def safe_write_json(path, data):
    """Atomically write JSON to disk (write to temp file, then replace)."""
    try:
        tmp_path = f"{path}.tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        shutil.move(tmp_path, path)
        return True
    except OSError:
        return False


# ==============================================================================
# AUTHENTICATION
# ==============================================================================

def load_users():
    return safe_read_json(USERS_FILE, [])


def save_users(users):
    return safe_write_json(USERS_FILE, users)


def find_user_by_email(email):
    email = email.strip().lower()
    for u in load_users():
        if u["email"].strip().lower() == email:
            return u
    return None


def find_user_by_id(user_id):
    for u in load_users():
        if u["id"] == user_id:
            return u
    return None


def hash_password(plain_password):
    return bcrypt.hashpw(plain_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(plain_password, password_hash):
    try:
        return bcrypt.checkpw(plain_password.encode("utf-8"), password_hash.encode("utf-8"))
    except (ValueError, TypeError):
        return False


def validate_signup_inputs(full_name, email, password, confirm_password):
    """Returns (is_valid, error_message)."""
    if not full_name or len(full_name.strip()) < 2:
        return False, "Please enter your full name."
    if not email or "@" not in email or "." not in email.split("@")[-1]:
        return False, "Please enter a valid email address."
    if len(password) < MIN_PASSWORD_LENGTH:
        return False, f"Password must be at least {MIN_PASSWORD_LENGTH} characters long."
    if password != confirm_password:
        return False, "Passwords do not match."
    if find_user_by_email(email):
        return False, "An account with this email already exists. Please log in instead."
    return True, ""


def signup_user(full_name, email, password):
    users = load_users()
    new_user = {
        "id": str(uuid.uuid4()),
        "full_name": full_name.strip(),
        "email": email.strip().lower(),
        "password_hash": hash_password(password),
        "created_at": datetime.utcnow().isoformat(),
    }
    users.append(new_user)
    save_users(users)
    # Pre-create this user's private folder tree so isolation is structural,
    # not just logical.
    user_history_dir(new_user["id"])
    user_vectors_dir(new_user["id"])
    user_uploads_dir(new_user["id"])
    user_reports_dir(new_user["id"])
    return new_user


def login_user(email, password):
    """Returns (user_dict_or_None, error_message)."""
    user = find_user_by_email(email)
    if not user:
        return None, "No account found with this email."
    if not verify_password(password, user["password_hash"]):
        return None, "Incorrect password."
    return user, ""


def init_session_state():
    defaults = {
        "authenticated": False,
        "user": None,
        "active_tab": "Upload SRS",
        "active_doc_id": None,
        "chat_sessions": {},   # doc_id -> list of {role, content}
        "openai_warning_shown": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def logout():
    st.session_state.authenticated = False
    st.session_state.user = None
    st.session_state.active_doc_id = None
    st.session_state.chat_sessions = {}
    st.session_state.active_tab = "Upload SRS"


def current_user_id():
    if st.session_state.get("user"):
        return st.session_state["user"]["id"]
    return None


# ==============================================================================
# DOCUMENT PARSING
# ==============================================================================

class DocumentParseError(Exception):
    pass


def extract_text_from_pdf(file_bytes):
    """Returns (full_text, page_count). Raises DocumentParseError on failure."""
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        raise DocumentParseError(f"Could not open PDF — the file may be corrupt or password protected ({e}).")
    if doc.page_count == 0:
        raise DocumentParseError("The PDF appears to have no pages.")
    pages_text = []
    for page in doc:
        try:
            pages_text.append(page.get_text("text"))
        except Exception:
            pages_text.append("")
    doc.close()
    full_text = "\n".join(pages_text)
    if not full_text.strip():
        raise DocumentParseError(
            "No extractable text was found in this PDF. It may be a scanned/image-only "
            "document without OCR support."
        )
    return full_text, len(pages_text)


def extract_text_from_docx(file_bytes):
    """Returns (full_text, page_count_estimate). Raises DocumentParseError on failure."""
    try:
        document = docx.Document(io.BytesIO(file_bytes))
    except Exception as e:
        raise DocumentParseError(f"Could not open DOCX — the file may be corrupt ({e}).")
    parts = []
    for para in document.paragraphs:
        if para.text:
            parts.append(para.text)
    for table in document.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text for cell in row.cells if cell.text)
            if row_text.strip():
                parts.append(row_text)
    full_text = "\n".join(parts)
    if not full_text.strip():
        raise DocumentParseError("No extractable text was found in this DOCX file.")
    # DOCX has no native page concept; estimate ~500 words per page.
    word_count = len(full_text.split())
    estimated_pages = max(1, round(word_count / 500))
    return full_text, estimated_pages


def extract_text_from_txt(file_bytes):
    """Returns (full_text, page_count_estimate). Raises DocumentParseError on failure."""
    try:
        full_text = file_bytes.decode("utf-8")
    except UnicodeDecodeError:
        try:
            full_text = file_bytes.decode("latin-1")
        except Exception as e:
            raise DocumentParseError(f"Could not decode text file ({e}).")
    if not full_text.strip():
        raise DocumentParseError("The uploaded TXT file is empty.")
    word_count = len(full_text.split())
    estimated_pages = max(1, round(word_count / 500))
    return full_text, estimated_pages


def parse_uploaded_document(uploaded_file):
    """
    Dispatch to the right parser based on file extension.
    Returns dict: {text, page_count, word_count, filename, extension}
    Raises DocumentParseError on any failure.
    """
    if uploaded_file is None:
        raise DocumentParseError("No file was uploaded.")

    filename = uploaded_file.name
    extension = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    file_bytes = uploaded_file.getvalue()

    if not file_bytes:
        raise DocumentParseError("The uploaded file is empty.")

    if extension == "pdf":
        text, page_count = extract_text_from_pdf(file_bytes)
    elif extension == "docx":
        text, page_count = extract_text_from_docx(file_bytes)
    elif extension == "txt":
        text, page_count = extract_text_from_txt(file_bytes)
    else:
        raise DocumentParseError(
            f"Unsupported file type '.{extension}'. Please upload a PDF, DOCX, or TXT file."
        )

    cleaned = clean_text(text)
    word_count = len(cleaned.split())

    return {
        "text": cleaned,
        "raw_bytes": file_bytes,
        "page_count": page_count,
        "word_count": word_count,
        "filename": filename,
        "extension": extension,
    }


def clean_text(text):
    """Normalize whitespace and strip control characters while preserving structure."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = "\n".join(line.strip() for line in text.split("\n"))
    text = re.sub(r"[^\x09\x0A\x20-\x7E\u00A0-\uFFFF]", "", text)
    return text.strip()


def chunk_text(text, chunk_size=CHUNK_SIZE, chunk_overlap=CHUNK_OVERLAP):
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=chunk_size,
        chunk_overlap=chunk_overlap,
        separators=["\n\n", "\n", ". ", " ", ""],
    )
    chunks = splitter.split_text(text)
    return [c for c in chunks if c.strip()]


# ==============================================================================
# OPENAI CLIENT HELPERS
# ==============================================================================

def get_openai_api_key():
    """Reads the API key from Streamlit secrets first, then environment variables."""
    try:
        if "OPENAI_API_KEY" in st.secrets and st.secrets["OPENAI_API_KEY"]:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("OPENAI_API_KEY", "")


def get_openai_client():
    api_key = get_openai_api_key()
    if not api_key:
        return None
    try:
        return OpenAI(api_key=api_key)
    except Exception:
        return None


def openai_key_missing_banner():
    st.error(
        "️ No OpenAI API key configured. Add it to Streamlit secrets as "
        "`OPENAI_API_KEY = \"sk-...\"` (Settings → Secrets in Colab/Streamlit Cloud, "
        "or a `.streamlit/secrets.toml` file) and reload the app."
    )


def call_chat_completion(messages, temperature=0.2, max_tokens=3000, json_mode=False):
    """
    Thin wrapper around the OpenAI chat completion endpoint with consistent
    error handling. Returns (text_or_None, error_message_or_None).
    """
    client = get_openai_client()
    if client is None:
        return None, "OpenAI API key is not configured."
    try:
        kwargs = dict(
            model=OPENAI_CHAT_MODEL,
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens,
        )
        if json_mode:
            kwargs["response_format"] = {"type": "json_object"}
        response = client.chat.completions.create(**kwargs)
        return response.choices[0].message.content, None
    except Exception as e:
        return None, f"OpenAI request failed: {e}"


def extract_json_from_text(text):
    """
    Best-effort extraction of a JSON object from a model response that may
    include markdown code fences or stray prose around the JSON payload.
    """
    if text is None:
        return None
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{.*\}", cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            return None
    return None


# ==============================================================================
# EMBEDDINGS / FAISS VECTOR STORE
# ==============================================================================

def get_embeddings_model():
    api_key = get_openai_api_key()
    if not api_key:
        return None
    try:
        return OpenAIEmbeddings(model=OPENAI_EMBEDDING_MODEL, openai_api_key=api_key)
    except Exception:
        return None


def build_and_save_vectorstore(user_id, doc_id, chunks):
    """
    Embeds the chunks and persists a FAISS index to disk under this user's
    private vectors directory. Returns (success_bool, error_message_or_None).
    """
    embeddings = get_embeddings_model()
    if embeddings is None:
        return False, "OpenAI API key is not configured; cannot generate embeddings."
    try:
        store = FAISS.from_texts(chunks, embeddings)
        target_dir = os.path.join(user_vectors_dir(user_id), doc_id)
        os.makedirs(target_dir, exist_ok=True)
        store.save_local(target_dir)
        return True, None
    except Exception as e:
        return False, f"Failed to build vector store: {e}"


def load_vectorstore(user_id, doc_id):
    """Loads a previously-saved FAISS index for this user's document, or None."""
    embeddings = get_embeddings_model()
    if embeddings is None:
        return None
    target_dir = os.path.join(user_vectors_dir(user_id), doc_id)
    if not os.path.isdir(target_dir):
        return None
    try:
        return FAISS.load_local(target_dir, embeddings, allow_dangerous_deserialization=True)
    except Exception:
        return None


def retrieve_relevant_chunks(user_id, doc_id, query, k=5):
    store = load_vectorstore(user_id, doc_id)
    if store is None:
        return []
    try:
        docs = store.similarity_search(query, k=k)
        return [d.page_content for d in docs]
    except Exception:
        return []


# ==============================================================================
# AI REQUIREMENT ENGINE — 18-CATEGORY ANALYSIS
# ==============================================================================

ANALYSIS_SYSTEM_PROMPT = """You are a senior requirements engineering consultant with deep expertise \
in software requirements specifications (SRS), business analysis, and project planning. You analyze \
SRS documents with rigor and produce structured, actionable output. You ONLY use information that is \
present in or reasonably inferable from the supplied document text. You never invent specific facts \
(like company names, exact dates, or budgets) that are not in the document. Where information is \
genuinely absent, say so explicitly rather than fabricating it.

You must respond with a single valid JSON object and nothing else — no markdown fences, no commentary \
before or after it."""

ANALYSIS_USER_PROMPT_TEMPLATE = """Analyze the following SRS document excerpt(s) and produce a complete \
requirement engineering analysis as a single JSON object with EXACTLY these top-level keys:

1. "business_goals": array of strings — the business objectives the system aims to achieve.
2. "functional_requirements": array of objects {{"id": "FR-1", "description": "...", "source_hint": "..."}}.
3. "non_functional_requirements": array of objects {{"id": "NFR-1", "category": "Performance|Security|Usability|Reliability|Scalability|Other", "description": "..."}}.
4. "constraints": array of strings — technical, business, regulatory, or resource constraints.
5. "stakeholders": array of objects {{"role": "...", "interest": "..."}}.
6. "assumptions": array of strings.
7. "dependencies": array of strings — internal or external dependencies.
8. "risks": array of objects {{"risk": "...", "impact": "Low|Medium|High", "mitigation": "..."}}.
9. "scope": object {{"in_scope": [..strings..], "out_of_scope": [..strings..]}}.
10. "missing_requirements": array of strings — important requirement areas the document does NOT cover \
(e.g. error handling, accessibility, audit logging) but a complete SRS normally would.
11. "requirement_quality_issues": array of strings — high-level quality problems noticed across the document \
(vagueness, missing acceptance criteria, etc.) — keep this list summary-level; detailed per-requirement \
validation happens separately.
12. "suggested_improvements": array of strings — concrete suggestions to strengthen the SRS.
13. "requirement_priority": array of objects {{"requirement_id": "FR-1", "priority": "Must Have|Should Have|Could Have|Won't Have"}} using MoSCoW.
14. "timeline_estimate": object {{"estimate": "e.g. 3-4 months", "reasoning": "..."}}.
15. "team_recommendation": array of objects {{"role": "e.g. Backend Engineer", "count": 1, "reasoning": "..."}}.
16. "project_complexity": object {{"level": "Low|Medium|High", "reasoning": "..."}}.
17. "requirement_categories": object mapping category name -> array of requirement ids/descriptions \
(e.g. {{"Authentication": [...], "Reporting": [...]}}).
18. "validation_report": string — a short paragraph summarizing overall SRS quality and completeness.

If the document does not contain enough information for a given key, return an empty array/object or a \
string explicitly stating "Not specified in the uploaded SRS." — never fabricate specifics.

DOCUMENT EXCERPTS:
---
{document_text}
---

Respond with ONLY the JSON object."""


def truncate_for_prompt(text, max_chars=18000):
    """Keep prompt sizes sane; very large SRS docs get sampled from start/middle/end."""
    if len(text) <= max_chars:
        return text
    third = max_chars // 3
    return (
        text[:third]
        + "\n\n...[middle of document omitted for length]...\n\n"
        + text[len(text) // 2: len(text) // 2 + third]
        + "\n\n...[more content omitted for length]...\n\n"
        + text[-third:]
    )


def run_full_analysis(document_text):
    """
    Calls OpenAI to produce the full 18-category requirement analysis.
    Returns (analysis_dict_or_None, error_message_or_None).
    """
    prompt = ANALYSIS_USER_PROMPT_TEMPLATE.format(document_text=truncate_for_prompt(document_text))
    messages = [
        {"role": "system", "content": ANALYSIS_SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    raw, err = call_chat_completion(messages, temperature=0.1, max_tokens=4000, json_mode=True)
    if err:
        return None, err
    parsed = extract_json_from_text(raw)
    if parsed is None:
        return None, "The AI response could not be parsed as valid JSON. Please try again."
    # Ensure every expected key exists so the UI never KeyErrors.
    for key in ANALYSIS_CATEGORIES:
        if key not in parsed:
            parsed[key] = [] if key not in ("scope", "timeline_estimate", "project_complexity",
                                             "requirement_categories") else {}
    return parsed, None


# ==============================================================================
# REQUIREMENT VALIDATION ENGINE
# ==============================================================================

VALIDATION_SYSTEM_PROMPT = """You are a meticulous requirements quality auditor. You examine individual \
requirement statements and flag concrete quality defects. You respond with a single valid JSON object \
only — no markdown fences, no commentary."""

VALIDATION_USER_PROMPT_TEMPLATE = """Review the following list of requirements extracted from an SRS \
document. For EACH requirement that has a quality defect, identify the defect type and produce a fix. \
Valid defect types are exactly: "Ambiguous", "Incomplete", "Unmeasurable", "Not Testable", "Conflicting".

A requirement may have more than one defect — list it once per defect with the most significant defect \
type, or combine reasoning if multiple apply. Requirements with no defects should be omitted entirely.

Return a JSON object with this exact shape:
{{
  "issues": [
    {{
      "requirement_id": "FR-1",
      "original_text": "...",
      "issue_type": "Ambiguous|Incomplete|Unmeasurable|Not Testable|Conflicting",
      "explanation": "why this is a problem",
      "corrected_version": "a rewritten, fixed requirement statement",
      "confidence_score": 0.0
    }}
  ]
}}

confidence_score must be a number between 0 and 1 representing how confident you are that this is a \
genuine defect.

REQUIREMENTS TO REVIEW:
---
{requirements_block}
---

Respond with ONLY the JSON object."""


def build_requirements_block(analysis):
    """Flattens functional + non-functional requirements into a numbered text block for validation."""
    lines = []
    for fr in analysis.get("functional_requirements", []):
        if isinstance(fr, dict):
            lines.append(f"{fr.get('id', 'FR-?')}: {fr.get('description', '')}")
        else:
            lines.append(str(fr))
    for nfr in analysis.get("non_functional_requirements", []):
        if isinstance(nfr, dict):
            lines.append(f"{nfr.get('id', 'NFR-?')}: {nfr.get('description', '')}")
        else:
            lines.append(str(nfr))
    return "\n".join(lines)


def run_requirement_validation(analysis):
    """
    Calls OpenAI to detect ambiguous/incomplete/unmeasurable/not-testable/conflicting requirements.
    Returns (issues_list, error_message_or_None).
    """
    requirements_block = build_requirements_block(analysis)
    if not requirements_block.strip():
        return [], None

    prompt = VALIDATION_USER_PROMPT_TEMPLATE.format(requirements_block=requirements_block)
    messages = [
        {"role": "system", "content": VALIDATION_SYSTEM_PROMPT},
        {"role": "user", "content": prompt},
    ]
    raw, err = call_chat_completion(messages, temperature=0.1, max_tokens=3000, json_mode=True)
    if err:
        return [], err
    parsed = extract_json_from_text(raw)
    if parsed is None or "issues" not in parsed:
        return [], "The validation response could not be parsed as valid JSON."

    issues = parsed["issues"]
    cleaned_issues = []
    for issue in issues:
        if not isinstance(issue, dict):
            continue
        try:
            score = float(issue.get("confidence_score", 0.5))
        except (TypeError, ValueError):
            score = 0.5
        cleaned_issues.append({
            "requirement_id": issue.get("requirement_id", "N/A"),
            "original_text": issue.get("original_text", ""),
            "issue_type": issue.get("issue_type", "Ambiguous"),
            "explanation": issue.get("explanation", ""),
            "corrected_version": issue.get("corrected_version", ""),
            "confidence_score": max(0.0, min(1.0, score)),
        })
    return cleaned_issues, None


# ==============================================================================
# RAG CHAT ENGINE
# ==============================================================================

CHAT_SYSTEM_PROMPT = """You are an assistant that answers questions STRICTLY using the provided excerpts \
from a user's uploaded SRS document. Do not use outside knowledge and do not speculate beyond what the \
excerpts say. If the excerpts do not contain the information needed to answer, respond with EXACTLY this \
sentence and nothing else: "Information not found in uploaded SRS."

When you can answer, be precise, cite specific requirement language where helpful, and keep the answer \
focused."""


def answer_question_with_rag(user_id, doc_id, question, chat_history):
    """
    Retrieves relevant chunks from the document's FAISS store and asks OpenAI to answer
    using only that context. Returns (answer_text, error_message_or_None).
    """
    relevant_chunks = retrieve_relevant_chunks(user_id, doc_id, question, k=5)
    if not relevant_chunks:
        context_block = "(no relevant content retrieved)"
    else:
        context_block = "\n\n---\n\n".join(relevant_chunks)

    history_messages = []
    for turn in chat_history[-6:]:
        history_messages.append({"role": turn["role"], "content": turn["content"]})

    messages = [{"role": "system", "content": CHAT_SYSTEM_PROMPT}]
    messages.extend(history_messages)
    messages.append({
        "role": "user",
        "content": f"DOCUMENT EXCERPTS:\n---\n{context_block}\n---\n\nQUESTION: {question}",
    })

    answer, err = call_chat_completion(messages, temperature=0.1, max_tokens=1200, json_mode=False)
    if err:
        return None, err
    if not answer or not answer.strip():
        return "Information not found in uploaded SRS.", None
    return answer.strip(), None


# ==============================================================================
# HISTORY MANAGEMENT
# ==============================================================================

def analysis_file_path(user_id, doc_id):
    return os.path.join(user_history_dir(user_id), f"analysis_{doc_id}.json")


def save_analysis_record(user_id, record):
    path = analysis_file_path(user_id, record["doc_id"])
    return safe_write_json(path, record)


# Keys introduced by the newer planning/intelligence modules. Older history
# records saved before these modules existed won't have them — this keeps
# the UI from ever KeyError-ing on a record from a previous app version.
NEW_MODULE_DEFAULTS = {
    "roadmap": None,
    "cost_estimation": None,
    "tech_recommendation": None,
    "architecture": None,
    "output_prediction": None,
    "coverage_score": None,
    "srs_improvements": None,
    "smart_project": None,
    "use_cases": None,
    "risk_feasibility": None,
    "team_roles": None,
}

# Human-readable labels for the badges shown in Analysis History (which
# modules have been generated for a given record). Keep in sync with
# NEW_MODULE_DEFAULTS above.
MODULE_BADGE_LABELS = {
    "roadmap": "🗺️ Roadmap",
    "cost_estimation": "💰 Cost",
    "tech_recommendation": "🧰 Tech Stack",
    "architecture": "🏗️ Architecture",
    "output_prediction": "🔮 Output Prediction",
    "coverage_score": "📈 Coverage",
    "srs_improvements": "✨ SRS Improver",
    "smart_project": "🧩 Project Plan",
    "use_cases": "🧾 Use Cases",
    "risk_feasibility": "⚠️ Risk & Feasibility",
    "team_roles": "👥 Team Roles",
}


def ensure_record_defaults(record):
    """Backfills any new-module keys missing from an older history record."""
    if record is None:
        return None
    for key, default in NEW_MODULE_DEFAULTS.items():
        if key not in record:
            record[key] = default
    record.setdefault("chat_history", [])
    record.setdefault("validation_issues", [])
    return record


def load_analysis_record(user_id, doc_id):
    path = analysis_file_path(user_id, doc_id)
    record = safe_read_json(path, None)
    return ensure_record_defaults(record)


def list_user_history(user_id):
    """Returns all analysis records for this user, most recent first."""
    folder = user_history_dir(user_id)
    records = []
    try:
        for fname in os.listdir(folder):
            if fname.startswith("analysis_") and fname.endswith(".json"):
                record = safe_read_json(os.path.join(folder, fname), None)
                if record:
                    records.append(ensure_record_defaults(record))
    except OSError:
        pass
    records.sort(key=lambda r: r.get("upload_date", ""), reverse=True)
    return records


def update_chat_history(user_id, doc_id, role, content):
    record = load_analysis_record(user_id, doc_id)
    if record is None:
        return
    record.setdefault("chat_history", [])
    record["chat_history"].append({
        "role": role,
        "content": content,
        "timestamp": datetime.utcnow().isoformat(),
    })
    save_analysis_record(user_id, record)


def delete_analysis_record(user_id, doc_id):
    path = analysis_file_path(user_id, doc_id)
    try:
        if os.path.exists(path):
            os.remove(path)
        vec_dir = os.path.join(user_vectors_dir(user_id), doc_id)
        if os.path.isdir(vec_dir):
            shutil.rmtree(vec_dir)
        upload_path = os.path.join(user_uploads_dir(user_id), doc_id)
        if os.path.isdir(upload_path):
            shutil.rmtree(upload_path)
        return True
    except OSError:
        return False


def save_module_result(user_id, doc_id, module_key, value):
    """Persists a single generated-module's output (e.g. 'roadmap') onto its history record."""
    record = load_analysis_record(user_id, doc_id)
    if record is None:
        return False
    record[module_key] = value
    return save_analysis_record(user_id, record)


# ==============================================================================
# REPORT EXPORT
# ==============================================================================

def generate_report_json(record):
    """Builds a clean, downloadable JSON report from a stored analysis record."""
    report = {
        "report_generated_at": datetime.utcnow().isoformat(),
        "document": {
            "filename": record.get("filename"),
            "upload_date": record.get("upload_date"),
            "page_count": record.get("page_count"),
            "word_count": record.get("word_count"),
        },
        "analysis": record.get("analysis", {}),
        "validation_issues": record.get("validation_issues", []),
        "chat_history": record.get("chat_history", []),
        "roadmap": record.get("roadmap"),
        "cost_estimation": record.get("cost_estimation"),
        "tech_recommendation": record.get("tech_recommendation"),
        "architecture": record.get("architecture"),
        "output_prediction": record.get("output_prediction"),
        "coverage_score": record.get("coverage_score"),
        "srs_improvements": record.get("srs_improvements"),
        "smart_project": record.get("smart_project"),
        "use_cases": record.get("use_cases"),
        "risk_feasibility": record.get("risk_feasibility"),
        "team_roles": record.get("team_roles"),
    }
    return json.dumps(report, indent=2, ensure_ascii=False)


# ==============================================================================
# NEW MODULE ENGINES — Developer Roadmap, Cost, Tech Stack, Architecture,
# Output Predictor, Coverage Score, SRS Improver, Smart Project Generator
# ==============================================================================

def summarize_analysis_for_prompt(analysis, max_chars=7000):
    """Compact JSON view of the core analysis, reused as context across the new modules."""
    compact = {
        "business_goals": analysis.get("business_goals", []),
        "functional_requirements": analysis.get("functional_requirements", []),
        "non_functional_requirements": analysis.get("non_functional_requirements", []),
        "constraints": analysis.get("constraints", []),
        "stakeholders": analysis.get("stakeholders", []),
        "risks": analysis.get("risks", []),
        "scope": analysis.get("scope", {}),
        "requirement_priority": analysis.get("requirement_priority", []),
        "project_complexity": analysis.get("project_complexity", {}),
        "team_recommendation": analysis.get("team_recommendation", []),
        "timeline_estimate": analysis.get("timeline_estimate", {}),
    }
    text = json.dumps(compact, ensure_ascii=False)
    if len(text) > max_chars:
        text = text[:max_chars] + "...[truncated]"
    return text


def run_module_json_generation(system_prompt, user_prompt, default_shape, max_tokens=3500):
    """
    Shared call pattern for every new-module generator: call OpenAI in JSON mode,
    parse the response, and backfill any missing keys against `default_shape` so
    the UI never KeyErrors. Returns (data_dict, error_message_or_None).
    """
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]
    raw, err = call_chat_completion(messages, temperature=0.2, max_tokens=max_tokens, json_mode=True)
    if err:
        return None, err
    parsed = extract_json_from_text(raw)
    if parsed is None:
        return None, "The AI response could not be parsed as valid JSON. Please try again."
    for key, default_value in default_shape.items():
        if key not in parsed:
            parsed[key] = default_value
    return parsed, None


# ---- MODULE 1 & 2: Developer Execution Roadmap + Daily/Weekly/Monthly Planner ----

ROADMAP_SYSTEM_PROMPT = """You are a senior technical program manager who turns SRS analyses into \
concrete, realistic execution plans for a software development team. Respond with a single valid JSON \
object only — no markdown fences, no commentary."""

ROADMAP_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, produce a complete \
developer execution roadmap as a JSON object with EXACTLY these keys:

1. "phases": array of objects, one per project phase (e.g. Planning, UI/UX, Backend, AI/RAG Integration, \
Testing, Deployment), each shaped {{"phase_name": "Phase 1: Planning", "tasks": ["..."], "duration": "1-2 weeks", \
"expected_output": "Approved SRS", "dependencies": "...", "owner": "Business Analyst"}}.
2. "daily_plan": array of EXACTLY 10 objects covering the first two working weeks, shaped \
{{"day_label": "Day 1", "tasks": ["..."], "hours": 8, "deliverables": ["..."], "meeting_agenda": "...", \
"expected_output": "..."}}.
3. "weekly_plan": array of objects, one per week for the full estimated timeline, shaped \
{{"week_label": "Week 1", "focus": "...", "monday": "...", "tuesday": "...", "wednesday": "...", \
"thursday": "...", "friday": "...", "saturday": "..."}}.
4. "monthly_milestones": array of objects shaped {{"month_label": "Month 1", "milestone": "...", \
"success_criteria": "..."}}.

Base phase count, week count, and month count on the project's actual estimated complexity and timeline. \
Be concrete and specific to this project — avoid generic filler.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

ROADMAP_DEFAULT_SHAPE = {"phases": [], "daily_plan": [], "weekly_plan": [], "monthly_milestones": []}


def generate_roadmap(analysis):
    prompt = ROADMAP_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(ROADMAP_SYSTEM_PROMPT, prompt, ROADMAP_DEFAULT_SHAPE, max_tokens=4000)


# ---- MODULE 3: Cost Estimation Engine ----

COST_SYSTEM_PROMPT = """You are a software project cost estimation specialist. You produce realistic, \
ranged cost estimates (Low/Medium/High) in USD based on project scope, complexity, and team needs. You are \
explicit that these are rough planning estimates, not quotes. Respond with a single valid JSON object only \
— no markdown fences, no commentary."""

COST_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, estimate project costs as a \
JSON object with EXACTLY these keys, each an object with "low", "medium", "high" numeric USD values plus a \
short "notes" string (except where noted):

1. "development_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "..."}}
2. "cloud_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "..."}} (monthly cloud hosting)
3. "openai_usage_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "..."}} (monthly OpenAI API usage)
4. "storage_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "..."}} (monthly)
5. "deployment_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "..."}} (one-time)
6. "maintenance_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "..."}} (monthly ongoing)
7. "total_estimate": {{"low": 0, "medium": 0, "high": 0, "notes": "one-time development + deployment total"}}
8. "monthly_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "recurring monthly total: cloud + openai + storage + maintenance"}}
9. "yearly_cost": {{"low": 0, "medium": 0, "high": 0, "notes": "monthly_cost x 12"}}
10. "assumptions": array of strings listing the key assumptions behind these numbers (team rates, region, etc.).

All numeric values must be plain numbers (no currency symbols, no strings).

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

COST_DEFAULT_SHAPE = {
    "development_cost": {}, "cloud_cost": {}, "openai_usage_cost": {}, "storage_cost": {},
    "deployment_cost": {}, "maintenance_cost": {}, "total_estimate": {}, "monthly_cost": {},
    "yearly_cost": {}, "assumptions": [],
}


def generate_cost_estimation(analysis):
    prompt = COST_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(COST_SYSTEM_PROMPT, prompt, COST_DEFAULT_SHAPE, max_tokens=2500)


# ---- MODULE 4: Technology Recommendation Engine ----

TECH_SYSTEM_PROMPT = """You are a pragmatic solutions architect who recommends concrete technology \
choices for software projects, always explaining the reasoning, advantages, and tradeoffs. Respond with a \
single valid JSON object only — no markdown fences, no commentary."""

TECH_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, recommend a technology stack \
as a JSON object with EXACTLY these keys, each shaped \
{{"technology": "...", "reason": "...", "advantages": ["..."], "tradeoffs": ["..."]}}:

1. "frontend" — frontend language/framework
2. "backend" — backend language/framework
3. "database" — primary database technology
4. "ai_stack" — AI/ML stack (models, RAG tooling, vector store)
5. "cloud" — cloud provider
6. "deployment" — deployment/hosting approach
7. "security" — security approach (auth, encryption, etc.)

Base every recommendation on what THIS specific project actually needs according to the analysis — not \
generic defaults.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

TECH_DEFAULT_SHAPE = {
    "frontend": {}, "backend": {}, "database": {}, "ai_stack": {},
    "cloud": {}, "deployment": {}, "security": {},
}


def generate_tech_recommendation(analysis):
    prompt = TECH_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(TECH_SYSTEM_PROMPT, prompt, TECH_DEFAULT_SHAPE, max_tokens=2500)


# ---- MODULE 5: Automatic System Architect ----

ARCHITECTURE_SYSTEM_PROMPT = """You are a software systems architect. You design clean, sensible system \
architectures grounded in the actual requirements given — not generic templates. Respond with a single \
valid JSON object only — no markdown fences, no commentary."""

ARCHITECTURE_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, design a system \
architecture as a JSON object with EXACTLY these keys:

1. "architecture_style": short string naming the overall style (e.g. "Layered Monolith with RAG Service").
2. "modules": array of objects {{"name": "...", "responsibility": "..."}}.
3. "folder_structure": a single string containing a plain-text folder tree (use indentation and "/" for \
folders, newlines between entries).
4. "data_flow": array of strings, each one ordered step in how data moves through the system end to end.
5. "api_suggestions": array of objects {{"method": "GET|POST|PUT|DELETE", "endpoint": "/api/...", "purpose": "..."}}.
6. "microservices_recommendation": {{"recommended": true or false, "reasoning": "..."}}.
7. "monolith_recommendation": {{"recommended": true or false, "reasoning": "..."}}.

Exactly one of microservices_recommendation/monolith_recommendation should have "recommended": true, based \
on actual project complexity and team size implied by the analysis.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

ARCHITECTURE_DEFAULT_SHAPE = {
    "architecture_style": "Not specified in the uploaded SRS.", "modules": [],
    "folder_structure": "", "data_flow": [], "api_suggestions": [],
    "microservices_recommendation": {}, "monolith_recommendation": {},
}


def generate_architecture(analysis):
    prompt = ARCHITECTURE_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(ARCHITECTURE_SYSTEM_PROMPT, prompt, ARCHITECTURE_DEFAULT_SHAPE, max_tokens=3000)


# ---- MODULE 7: Project Output Predictor ----

OUTPUT_PREDICTOR_SYSTEM_PROMPT = """You are a delivery lead who predicts what a software project will \
concretely ship, based on its requirements. Respond with a single valid JSON object only — no markdown \
fences, no commentary."""

OUTPUT_PREDICTOR_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, predict the \
project's eventual output as a JSON object with EXACTLY these keys:

1. "final_deliverables": array of strings.
2. "screens": array of strings naming the application screens/pages expected.
3. "modules": array of strings naming the functional modules expected.
4. "estimated_features": array of strings naming concrete features.
5. "complexity": {{"level": "Low|Medium|High", "reasoning": "..."}}.
6. "deployment_readiness": {{"status": "Not Ready|Partially Ready|Ready", "reasoning": "..."}}.
7. "team_size": {{"recommended_size": 0, "breakdown": [{{"role": "...", "count": 0}}]}}.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

OUTPUT_PREDICTOR_DEFAULT_SHAPE = {
    "final_deliverables": [], "screens": [], "modules": [], "estimated_features": [],
    "complexity": {}, "deployment_readiness": {}, "team_size": {},
}


def generate_output_prediction(analysis):
    prompt = OUTPUT_PREDICTOR_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(OUTPUT_PREDICTOR_SYSTEM_PROMPT, prompt, OUTPUT_PREDICTOR_DEFAULT_SHAPE, max_tokens=2200)


# ---- MODULE 8: Requirement Coverage Score ----

COVERAGE_SYSTEM_PROMPT = """You are a requirements quality scorer. You assign objective-feeling percentage \
scores (0-100) for how complete, clear, testable, traceable, and risky a set of requirements is, based on \
the analysis and any already-detected validation issues. Respond with a single valid JSON object only — no \
markdown fences, no commentary."""

COVERAGE_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary and the list of already-detected \
quality issues below, calculate requirement coverage scores as a JSON object with EXACTLY these keys (all \
percentages are numbers from 0 to 100):

1. "completeness_pct": how much of a typical SRS's expected content is present.
2. "clarity_pct": how clearly worded the requirements are (inverse of ambiguity).
3. "testability_pct": how testable/measurable the requirements are.
4. "traceability_pct": how well requirements trace to business goals and stakeholders.
5. "risk_pct": overall project risk level (higher = riskier).
6. "missing_pct": how much important content appears to be missing.
7. "overall_score": a single 0-100 composite score.
8. "summary": a short paragraph explaining the scores.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

DETECTED QUALITY ISSUES ({issue_count} found):
{issues_block}

Respond with ONLY the JSON object."""

COVERAGE_DEFAULT_SHAPE = {
    "completeness_pct": 0, "clarity_pct": 0, "testability_pct": 0, "traceability_pct": 0,
    "risk_pct": 0, "missing_pct": 0, "overall_score": 0, "summary": "Not specified in the uploaded SRS.",
}


def generate_coverage_score(analysis, validation_issues):
    issues_block = "\n".join(
        f"- [{i.get('issue_type')}] {i.get('requirement_id')}: {i.get('explanation')}"
        for i in (validation_issues or [])
    ) or "None detected."
    prompt = COVERAGE_USER_PROMPT_TEMPLATE.format(
        analysis_summary=summarize_analysis_for_prompt(analysis),
        issue_count=len(validation_issues or []),
        issues_block=issues_block,
    )
    return run_module_json_generation(COVERAGE_SYSTEM_PROMPT, prompt, COVERAGE_DEFAULT_SHAPE, max_tokens=1500)


# ---- MODULE 11: SRS Quality Improver ----

SRS_IMPROVER_SYSTEM_PROMPT = """You are an expert requirements editor. You rewrite weak requirements into \
strong, unambiguous, testable ones, and assemble a clean improved SRS section. Respond with a single valid \
JSON object only — no markdown fences, no commentary."""

SRS_IMPROVER_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary and detected quality issues \
below, produce an SRS improvement package as a JSON object with EXACTLY these keys:

1. "improvements": array of objects {{"original": "...", "issue": "...", "suggested": "...", "reason": "..."}} \
— one entry per weak requirement found in the detected issues (or, if none were detected, identify the \
weakest 3-5 requirements yourself).
2. "improved_srs_document": a single string containing a clean, well-organized improved SRS section \
(functional + non-functional requirements only) using the corrected wording, formatted as readable \
plain text with line breaks between requirements.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

DETECTED QUALITY ISSUES:
{issues_block}

Respond with ONLY the JSON object."""

SRS_IMPROVER_DEFAULT_SHAPE = {"improvements": [], "improved_srs_document": ""}


def generate_srs_improvements(analysis, validation_issues):
    issues_block = "\n".join(
        f"- [{i.get('issue_type')}] {i.get('requirement_id')}: {i.get('original_text')} -- {i.get('explanation')}"
        for i in (validation_issues or [])
    ) or "None detected — please identify the weakest requirements yourself."
    prompt = SRS_IMPROVER_USER_PROMPT_TEMPLATE.format(
        analysis_summary=summarize_analysis_for_prompt(analysis),
        issues_block=issues_block,
    )
    return run_module_json_generation(SRS_IMPROVER_SYSTEM_PROMPT, prompt, SRS_IMPROVER_DEFAULT_SHAPE, max_tokens=3000)


# ---- MODULE 12: Smart Project Generator (Epics, Stories, Sprints, Kanban, Backlog) ----

SMART_PROJECT_SYSTEM_PROMPT = """You are an agile delivery lead who converts requirements into a ready-to-run \
agile project plan: epics, user stories with acceptance criteria, a sprint plan, a kanban snapshot, and a \
prioritized backlog. Respond with a single valid JSON object only — no markdown fences, no commentary."""

SMART_PROJECT_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, generate an agile \
project plan as a JSON object with EXACTLY these keys:

1. "epics": array of objects {{"id": "EPIC-1", "title": "...", "description": "..."}}.
2. "user_stories": array of objects {{"id": "US-1", "epic_id": "EPIC-1", \
"story": "As a [role], I want [capability] so that [benefit]", "acceptance_criteria": ["..."], \
"story_points": 1}}.
3. "sprint_plan": array of objects {{"sprint": "Sprint 1", "goal": "...", "story_ids": ["US-1", "US-2"], \
"duration": "2 weeks"}}.
4. "kanban": object {{"backlog": ["US-..."], "to_do": ["US-..."], "in_progress": ["US-..."], "done": ["US-..."]}} \
representing a sensible starting snapshot (most items in backlog/to_do since the project is just starting).
5. "backlog": array of objects {{"id": "US-1", "priority": "High|Medium|Low"}} representing the prioritized \
full backlog order.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

SMART_PROJECT_DEFAULT_SHAPE = {
    "epics": [], "user_stories": [], "sprint_plan": [],
    "kanban": {"backlog": [], "to_do": [], "in_progress": [], "done": []}, "backlog": [],
}


def generate_smart_project(analysis):
    prompt = SMART_PROJECT_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(SMART_PROJECT_SYSTEM_PROMPT, prompt, SMART_PROJECT_DEFAULT_SHAPE, max_tokens=4000)


# ---- MODULE 13: Use Case Generator ----

USE_CASE_SYSTEM_PROMPT = """You are a senior business analyst who writes complete, professional UML-style \
use case specifications from SRS requirement analyses. You ground every use case in the actual \
requirements, actors, and stakeholders provided in the analysis — you do not invent unrelated \
functionality. Respond with a single valid JSON object only — no markdown fences, no commentary."""

USE_CASE_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, generate a complete use \
case specification as a JSON object with EXACTLY these keys:

1. "primary_actors": array of strings naming the primary actors (the user/system roles that initiate use cases).
2. "secondary_actors": array of strings naming secondary/supporting actors (external systems, APIs, services).
3. "use_cases": array of objects, one per significant functional requirement or user-facing capability, each shaped:
{{
  "id": "UC-1",
  "title": "...",
  "primary_actor": "...",
  "secondary_actors": ["..."],
  "trigger": "the event that starts this use case",
  "inputs": ["..."],
  "outputs": ["..."],
  "preconditions": ["..."],
  "postconditions": ["..."],
  "main_flow": ["step 1...", "step 2...", "..."],
  "alternative_flow": ["..."],
  "exception_flow": ["..."],
  "business_rules": ["..."],
  "success_scenario": "a short paragraph describing the successful end state",
  "failure_scenario": "a short paragraph describing what happens on failure"
}}

Generate one use case per significant functional requirement or user-facing capability found in the \
analysis — do not invent capabilities outside it. If the analysis is sparse, generate at least 3 reasonable \
core use cases grounded in the business goals and stakeholders given.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

USE_CASE_DEFAULT_SHAPE = {"primary_actors": [], "secondary_actors": [], "use_cases": []}


def build_use_case_diagram_data(data):
    """Derives UML-style actor/use-case/relationship data directly from the generated use cases,
    so the diagram is always consistent with the detailed use case content (no separate LLM call)."""
    use_cases = data.get("use_cases", []) or []
    actors = set(a for a in (data.get("primary_actors", []) or []) if a)
    actors.update(a for a in (data.get("secondary_actors", []) or []) if a)
    relationships = []
    diagram_use_cases = []
    for idx, uc in enumerate(use_cases):
        if not isinstance(uc, dict):
            continue
        uc_id = uc.get("id") or f"UC-{idx + 1}"
        diagram_use_cases.append({"id": uc_id, "title": uc.get("title", "")})
        primary_actor = uc.get("primary_actor")
        if primary_actor:
            actors.add(primary_actor)
            relationships.append({"actor": primary_actor, "use_case": uc_id, "type": "primary"})
        for sa in uc.get("secondary_actors", []) or []:
            if sa:
                actors.add(sa)
                relationships.append({"actor": sa, "use_case": uc_id, "type": "secondary"})
    return {
        "actors": sorted(actors),
        "use_cases": diagram_use_cases,
        "relationships": relationships,
    }


def build_use_case_table(data):
    """Flattens the generated use cases into simple row dicts for tabular display/export."""
    rows = []
    for uc in data.get("use_cases", []) or []:
        if not isinstance(uc, dict):
            continue
        rows.append({
            "id": uc.get("id", "—"),
            "title": uc.get("title", "—"),
            "primary_actor": uc.get("primary_actor", "—"),
            "secondary_actors": ", ".join(uc.get("secondary_actors", []) or []) or "—",
            "trigger": uc.get("trigger", "—"),
        })
    return rows


def generate_use_cases(analysis):
    prompt = USE_CASE_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    data, err = run_module_json_generation(USE_CASE_SYSTEM_PROMPT, prompt, USE_CASE_DEFAULT_SHAPE, max_tokens=4000)
    if err or data is None:
        return data, err
    data["diagram_data"] = build_use_case_diagram_data(data)
    data["use_case_table"] = build_use_case_table(data)
    return data, None


# ---- MODULE 14: Risk & Feasibility Analyzer ----

RISK_SYSTEM_PROMPT = """You are a seasoned project risk and feasibility analyst. You assess software \
projects across technical, budget, schedule, operational, security, maintainability, and scalability risk \
dimensions, and produce an overall feasibility and complexity assessment. You base every judgment on the \
actual analysis provided — not generic boilerplate. Respond with a single valid JSON object only — no \
markdown fences, no commentary."""

RISK_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, produce a complete risk & \
feasibility assessment as a JSON object with EXACTLY these keys. Each of the seven risk keys must be an \
object shaped {{"probability": "Low|Medium|High", "impact": "Low|Medium|High", "severity": "Low|Medium|High", \
"description": "...", "mitigation": "..."}}:

1. "technical_risk"
2. "budget_risk"
3. "schedule_risk"
4. "operational_risk"
5. "security_risk"
6. "maintainability_risk"
7. "scalability_risk"
8. "feasibility_score": number from 0 to 100 (higher = more feasible).
9. "complexity_score": number from 0 to 100 (higher = more complex).
10. "recommendation": a short paragraph giving an overall go / no-go / conditional-go recommendation with reasoning.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

RISK_DEFAULT_SHAPE = {
    "technical_risk": {}, "budget_risk": {}, "schedule_risk": {}, "operational_risk": {},
    "security_risk": {}, "maintainability_risk": {}, "scalability_risk": {},
    "feasibility_score": 0, "complexity_score": 0,
    "recommendation": "Not specified in the uploaded SRS.",
}

RISK_CATEGORIES = [
    ("technical_risk", "⚙️ Technical Risk"), ("budget_risk", "💰 Budget Risk"),
    ("schedule_risk", "🗓️ Schedule Risk"), ("operational_risk", "🔧 Operational Risk"),
    ("security_risk", "🔐 Security Risk"), ("maintainability_risk", "🛠️ Maintainability Risk"),
    ("scalability_risk", "📈 Scalability Risk"),
]


def generate_risk_feasibility(analysis):
    prompt = RISK_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    return run_module_json_generation(RISK_SYSTEM_PROMPT, prompt, RISK_DEFAULT_SHAPE, max_tokens=3000)


# ---- MODULE 15: Team Role Generator ----

TEAM_ROLES = [
    ("ai_engineer", "🤖 AI Engineer"), ("backend_developer", "⚙️ Backend Developer"),
    ("frontend_developer", "🖥️ Frontend Developer"), ("qa_engineer", "🧪 QA Engineer"),
    ("project_manager", "🧭 Project Manager"), ("devops", "🚀 DevOps"),
    ("ui_designer", "🎨 UI Designer"),
]

TEAM_ROLE_SYSTEM_PROMPT = """You are an experienced engineering resource planner who builds concrete \
staffing plans for software projects: responsibilities, weekly hours, costs, deliverables, skills, and \
timeline per role, plus an overall resourcing summary. Respond with a single valid JSON object only — no \
markdown fences, no commentary."""

TEAM_ROLE_USER_PROMPT_TEMPLATE = """Using the requirement analysis summary below, build a team staffing \
plan as a JSON object with EXACTLY these keys:

1. "roles": object with EXACTLY these 7 keys, each shaped \
{{"needed": true or false, "responsibilities": ["..."], "hours_per_week": 0, "estimated_cost": 0, \
"deliverables": ["..."], "skills": ["..."], "timeline": "e.g. Weeks 1-8"}}:
   - "ai_engineer"
   - "backend_developer"
   - "frontend_developer"
   - "qa_engineer"
   - "project_manager"
   - "devops"
   - "ui_designer"
   Set "needed": false (with zero/empty values) only if this specific project genuinely has no use for \
that role. "estimated_cost" is the total estimated USD cost for that role across the whole project, as a \
plain number (no currency symbols).
2. "resource_planning": a short paragraph summarizing the overall staffing approach and reasoning.
3. "workload_distribution": array of objects {{"role": "...", "percentage": 0}} — each role's share of \
total project effort, summing to roughly 100.

REQUIREMENT ANALYSIS SUMMARY:
---
{analysis_summary}
---

Respond with ONLY the JSON object."""

TEAM_ROLE_DEFAULT_SHAPE = {
    "roles": {}, "resource_planning": "Not specified in the uploaded SRS.", "workload_distribution": [],
}


def compute_team_aggregates(data):
    """Computes team_size, team_cost, and a hours-per-person breakdown directly from the per-role data,
    so these summary numbers always stay consistent with the detailed role figures."""
    roles = data.get("roles", {}) or {}
    hours_per_person = []
    total_cost = 0.0
    team_size = 0
    for key, label in TEAM_ROLES:
        role = roles.get(key, {}) or {}
        if not isinstance(role, dict):
            role = {}
        needed = role.get("needed", True)
        try:
            hpw = float(role.get("hours_per_week", 0) or 0)
        except (TypeError, ValueError):
            hpw = 0.0
        try:
            cost = float(role.get("estimated_cost", 0) or 0)
        except (TypeError, ValueError):
            cost = 0.0
        if needed and (hpw > 0 or cost > 0):
            team_size += 1
        hours_per_person.append({"role": label, "hours_per_week": hpw})
        total_cost += cost
    return {"team_size": team_size, "team_cost": total_cost, "hours_per_person": hours_per_person}


def generate_team_roles(analysis):
    prompt = TEAM_ROLE_USER_PROMPT_TEMPLATE.format(analysis_summary=summarize_analysis_for_prompt(analysis))
    data, err = run_module_json_generation(TEAM_ROLE_SYSTEM_PROMPT, prompt, TEAM_ROLE_DEFAULT_SHAPE, max_tokens=3500)
    if err or data is None:
        return data, err
    data["team_planning"] = compute_team_aggregates(data)
    return data, None


# ==============================================================================
# UI — GLOBAL STYLE
# ==============================================================================

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
    /* Premium Deep Space Palette */
    --bg-deep: #0a0e27;
    --bg-primary: #111638;
    --bg-secondary: #1a1f4e;
    --bg-tertiary: #252b6e;
    --bg-card: rgba(26, 31, 78, 0.4);
    --bg-card-hover: rgba(37, 43, 110, 0.6);
    --bg-elevated: rgba(26, 31, 78, 0.8);
    
    /* Premium Accents - Cyan & Purple */
    --accent-primary: #06b6d4;
    --accent-secondary: #8b5cf6;
    --accent-tertiary: #3b82f6;
    --accent-soft: rgba(6, 182, 212, 0.15);
    --accent-gradient: linear-gradient(135deg, #06b6d4 0%, #8b5cf6 50%, #3b82f6 100%);
    --accent-gradient-reverse: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 50%, #06b6d4 100%);
    --accent-gradient-subtle: linear-gradient(135deg, rgba(6, 182, 212, 0.1) 0%, rgba(139, 92, 246, 0.05) 100%);
    --accent-glow: 0 0 20px rgba(6, 182, 212, 0.5);
    
    /* Success Accent - Emerald */
    --success-accent: #10b981;
    --success-soft: rgba(16, 185, 129, 0.15);
    
    /* Borders */
    --border-primary: rgba(139, 92, 246, 0.2);
    --border-secondary: rgba(6, 182, 212, 0.1);
    --border-accent: rgba(6, 182, 212, 0.4);
    
    /* Typography */
    --text-primary: #f8fafc;
    --text-secondary: #cbd5e1;
    --text-tertiary: #94a3b8;
    --text-muted: #64748b;
    
    /* Status Colors */
    --success: #10b981;
    --success-soft: rgba(16, 185, 129, 0.15);
    --warning: #f59e0b;
    --warning-soft: rgba(245, 158, 11, 0.15);
    --danger: #ef4444;
    --danger-soft: rgba(239, 68, 68, 0.15);
    --info: #06b6d4;
    --info-soft: rgba(6, 182, 212, 0.15);
    
    /* Spacing & Radius */
    --radius-xs: 6px;
    --radius-sm: 10px;
    --radius-md: 14px;
    --radius-lg: 20px;
    --radius-xl: 28px;
    --radius-full: 9999px;
    
    /* Premium Shadows */
    --shadow-sm: 0 2px 8px rgba(0, 0, 0, 0.4);
    --shadow-md: 0 8px 16px rgba(0, 0, 0, 0.5);
    --shadow-lg: 0 16px 32px rgba(0, 0, 0, 0.6);
    --shadow-xl: 0 24px 48px rgba(0, 0, 0, 0.7);
    --shadow-accent: 0 8px 24px rgba(6, 182, 212, 0.3);
    --shadow-glow: 0 0 30px rgba(139, 92, 246, 0.4);
    
    /* Transitions */
    --transition-fast: 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    --transition-smooth: 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    --transition-slow: 0.6s cubic-bezier(0.4, 0, 0.2, 1);
}

/* Base Reset & Typography */
* {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
}

html, body, .stApp {
    background: var(--bg-deep);
    color: var(--text-primary);
    position: relative;
}

/* Animated Gradient Mesh Background */
.stApp::before {
    content: '';
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: 
        radial-gradient(circle at 20% 30%, rgba(6, 182, 212, 0.08) 0%, transparent 50%),
        radial-gradient(circle at 80% 70%, rgba(139, 92, 246, 0.08) 0%, transparent 50%),
        radial-gradient(circle at 50% 50%, rgba(59, 130, 246, 0.05) 0%, transparent 50%);
    animation: meshGradient 20s ease infinite;
    z-index: -1;
    pointer-events: none;
}

@keyframes meshGradient {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: 0.8; transform: scale(1.1); }
}

/* Hide Streamlit chrome */
#MainMenu, footer, header[data-testid="stHeader"] {
    background: transparent;
}

/* ================================================================
   SIDEBAR - Premium Glass Navigation
   ================================================================ */
section[data-testid="stSidebar"] {
    background: rgba(17, 22, 56, 0.6);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border-right: 1px solid var(--border-primary);
    box-shadow: 4px 0 24px rgba(0, 0, 0, 0.3);
}

section[data-testid="stSidebar"] > div {
    padding-top: 0;
}

.sidebar-brand {
    position: sticky;
    top: 0;
    z-index: 10;
    padding: 1.25rem 1.25rem;
    margin-bottom: 0;
    border-bottom: 1px solid var(--border-primary);
    background: linear-gradient(135deg, rgba(6, 182, 212, 0.1) 0%, rgba(139, 92, 246, 0.05) 100%);
    backdrop-filter: blur(10px);
}

.sidebar-brand .brand-title {
    font-weight: 800;
    font-size: 1.2rem;
    letter-spacing: -0.02em;
    background: var(--accent-gradient);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
    line-height: 1.2;
    animation: gradientShift 3s ease infinite;
    background-size: 200% 200%;
}

@keyframes gradientShift {
    0%, 100% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
}

.sidebar-brand .brand-sub {
    color: var(--text-tertiary);
    font-size: 0.65rem;
    margin-top: 0.35rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    font-weight: 600;
}

.user-chip {
    background: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: var(--radius-md);
    padding: 0.85rem 1rem;
    margin: 0.85rem 0;
    backdrop-filter: blur(10px);
    transition: all var(--transition-smooth);
    box-shadow: var(--shadow-sm);
}

.user-chip:hover {
    border-color: var(--border-accent);
    background: var(--bg-card-hover);
    box-shadow: var(--shadow-accent);
    transform: translateY(-2px);
}

.user-chip .u-name {
    font-weight: 600;
    font-size: 0.9rem;
    color: var(--text-primary);
    letter-spacing: -0.01em;
}

.user-chip .u-email {
    color: var(--text-tertiary);
    font-size: 0.72rem;
    margin-top: 0.2rem;
}

.sidebar-section-label {
    color: var(--text-tertiary);
    font-size: 0.62rem;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    font-weight: 700;
    margin: 1rem 0 0.5rem 0.3rem;
    padding-left: 0.6rem;
    border-left: 3px solid var(--accent-gradient);
}

section[data-testid="stSidebar"] div.stButton > button {
    background: transparent !important;
    color: var(--text-secondary) !important;
    text-align: left !important;
    justify-content: flex-start !important;
    border: 1px solid transparent !important;
    box-shadow: none !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
    padding: 0.6rem 0.9rem !important;
    margin: 0.15rem 0 !important;
    border-radius: var(--radius-sm) !important;
    transition: all var(--transition-fast) !important;
    letter-spacing: -0.01em;
    position: relative;
    overflow: hidden;
}

section[data-testid="stSidebar"] div.stButton > button::before {
    content: '';
    position: absolute;
    top: 0;
    left: -100%;
    width: 100%;
    height: 100%;
    background: linear-gradient(90deg, transparent, rgba(6, 182, 212, 0.1), transparent);
    transition: left 0.4s ease;
}

section[data-testid="stSidebar"] div.stButton > button:hover::before {
    left: 100%;
}

section[data-testid="stSidebar"] div.stButton > button:hover {
    background: var(--bg-card-hover) !important;
    border-color: var(--border-accent) !important;
    color: var(--accent-primary) !important;
    transform: translateX(3px);
    box-shadow: 0 4px 12px rgba(6, 182, 212, 0.2);
}

section[data-testid="stSidebar"] div.stButton > button:active {
    transform: translateX(1px);
}

/* ================================================================
   TYPOGRAPHY - Premium Hierarchy
   ================================================================ */
h1, h2, h3, h4, h5, h6 {
    color: var(--text-primary) !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em;
    line-height: 1.25;
}

h1 { 
    font-size: 2rem !important;
    background: var(--accent-gradient);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
}
h2 { font-size: 1.65rem !important; }
h3 { font-size: 1.35rem !important; }
h4 { font-size: 1.15rem !important; }

p, span, label, .stMarkdown {
    color: var(--text-primary);
    line-height: 1.6;
}

.stCaption, [data-testid="stCaptionContainer"] {
    color: var(--text-tertiary) !important;
    font-size: 0.82rem;
}

/* ================================================================
   HERO SECTION - Premium Animated
   ================================================================ */
.app-hero {
    position: relative;
    padding: 2rem 2rem;
    border-radius: var(--radius-xl);
    background: linear-gradient(135deg, rgba(6, 182, 212, 0.1) 0%, rgba(139, 92, 246, 0.08) 50%, rgba(59, 130, 246, 0.1) 100%);
    border: 1px solid var(--border-accent);
    margin-bottom: 1.5rem;
    overflow: hidden;
    backdrop-filter: blur(20px);
    box-shadow: var(--shadow-glow);
    animation: heroGlow 4s ease-in-out infinite;
}

@keyframes heroGlow {
    0%, 100% { box-shadow: 0 0 30px rgba(139, 92, 246, 0.3); }
    50% { box-shadow: 0 0 50px rgba(6, 182, 212, 0.4); }
}

.app-hero::before {
    content: "";
    position: absolute;
    top: -50%;
    right: -20%;
    width: 400px;
    height: 400px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(6, 182, 212, 0.15) 0%, transparent 70%);
    filter: blur(60px);
    animation: float 6s ease-in-out infinite;
}

.app-hero::after {
    content: "";
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    height: 2px;
    background: var(--accent-gradient);
    animation: shimmer 3s linear infinite;
}

@keyframes float {
    0%, 100% { transform: translate(0, 0) scale(1); }
    50% { transform: translate(-30px, 30px) scale(1.1); }
}

@keyframes shimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

.app-hero h1 {
    margin: 0;
    font-size: 2rem;
    position: relative;
    z-index: 1;
    letter-spacing: -0.02em;
    animation: fadeInUp 0.8s ease;
}

.app-hero p {
    color: var(--text-secondary);
    margin: 0.5rem 0 0 0;
    position: relative;
    z-index: 1;
    font-size: 1rem;
    max-width: 600px;
    animation: fadeInUp 0.8s ease 0.2s backwards;
}

@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

/* ================================================================
   CARDS - Premium Glassmorphism
   ================================================================ */
.metric-card, .content-card, .glass-card, .history-item, .module-card {
    background: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: var(--radius-lg);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    transition: all var(--transition-smooth);
    box-shadow: var(--shadow-md);
    position: relative;
    overflow: hidden;
}

.metric-card::before, .content-card::before, .module-card::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 3px;
    background: var(--accent-gradient);
    opacity: 0;
    transition: opacity var(--transition-smooth);
}

.metric-card:hover::before, .content-card:hover::before, .module-card:hover::before {
    opacity: 1;
}

.metric-card:hover, .content-card:hover, .glass-card:hover, .module-card:hover {
    border-color: var(--border-accent);
    background: var(--bg-card-hover);
    box-shadow: var(--shadow-glow);
    transform: translateY(-4px) scale(1.02);
}

.metric-card {
    padding: 1.25rem;
    text-align: left;
    position: relative;
    overflow: hidden;
}

.metric-card::after {
    content: '';
    position: absolute;
    top: -50%;
    right: -50%;
    width: 200%;
    height: 200%;
    background: radial-gradient(circle, rgba(6, 182, 212, 0.1) 0%, transparent 70%);
    opacity: 0;
    transition: opacity var(--transition-smooth);
    pointer-events: none;
}

.metric-card:hover::after {
    opacity: 1;
}

.metric-card .metric-icon {
    font-size: 1.5rem;
    opacity: 0.9;
    margin-bottom: 0.5rem;
    filter: drop-shadow(0 0 8px rgba(6, 182, 212, 0.5));
}

.metric-card .metric-label {
    color: var(--text-tertiary);
    font-size: 0.68rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-top: 0.5rem;
    font-weight: 700;
}

.metric-card .metric-value {
    color: var(--text-primary);
    font-size: 2rem;
    font-weight: 800;
    margin-top: 0.3rem;
    letter-spacing: -0.02em;
    background: var(--accent-gradient);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
}

.metric-card .metric-delta {
    font-size: 0.75rem;
    margin-top: 0.4rem;
    font-weight: 600;
}

.metric-delta.up { color: var(--success); }
.metric-delta.down { color: var(--danger); }

.content-card {
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}

.content-card h4 {
    margin-top: 0 !important;
    font-size: 1.1rem !important;
    color: var(--accent-primary) !important;
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-weight: 700 !important;
}

.module-card {
    padding: 1.5rem 1.75rem;
    margin-bottom: 1rem;
}

.module-card h3 {
    margin-top: 0;
    font-size: 1.25rem;
    font-weight: 700;
}

.module-card .module-tagline {
    color: var(--text-secondary);
    font-size: 0.88rem;
    margin-top: -0.2rem;
}

/* ================================================================
   BADGES - Premium
   ================================================================ */
.badge {
    display: inline-block;
    padding: 0.25rem 0.75rem;
    border-radius: var(--radius-full);
    font-size: 0.68rem;
    font-weight: 700;
    margin-right: 0.3rem;
    letter-spacing: 0.03em;
    border: 1px solid transparent;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
    transition: all var(--transition-fast);
}

.badge:hover {
    transform: scale(1.1);
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
}

.badge-high {
    background-color: var(--danger-soft);
    color: var(--danger);
    border-color: rgba(231, 76, 60, 0.2);
}

.badge-medium {
    background-color: var(--warning-soft);
    color: var(--warning);
    border-color: rgba(243, 156, 18, 0.2);
}

.badge-low {
    background-color: var(--success-soft);
    color: var(--success);
    border-color: rgba(46, 204, 113, 0.2);
}

.badge-info {
    background-color: var(--info-soft);
    color: var(--info);
    border-color: rgba(52, 152, 219, 0.2);
}

.badge-neutral {
    background-color: var(--accent-soft);
    color: var(--accent-primary);
    border-color: rgba(212, 175, 55, 0.2);
}

/* ================================================================
   HISTORY LIST
   ================================================================ */
.history-item {
    padding: 1rem 1.25rem;
    margin-bottom: 0.75rem;
    transition: all var(--transition-smooth);
}

.history-item:hover {
    transform: translateX(4px);
}

/* ================================================================
   BUTTONS - Premium with Glow
   ================================================================ */
div.stButton > button, .stDownloadButton > button {
    background: var(--accent-gradient) !important;
    color: #fff !important;
    border: none !important;
    border-radius: var(--radius-sm) !important;
    font-weight: 700 !important;
    padding: 0.7rem 1.5rem !important;
    transition: all var(--transition-smooth) !important;
    box-shadow: 0 4px 15px rgba(6, 182, 212, 0.4) !important;
    letter-spacing: -0.01em;
    font-size: 0.9rem !important;
    position: relative;
    overflow: hidden;
}

div.stButton > button::before,
.stDownloadButton > button::before {
    content: '';
    position: absolute;
    top: 50%;
    left: 50%;
    width: 0;
    height: 0;
    border-radius: 50%;
    background: rgba(255, 255, 255, 0.3);
    transform: translate(-50%, -50%);
    transition: width 0.6s ease, height 0.6s ease;
}

div.stButton > button:hover::before,
.stDownloadButton > button:hover::before {
    width: 300px;
    height: 300px;
}

div.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-3px) scale(1.05) !important;
    box-shadow: 0 8px 25px rgba(6, 182, 212, 0.6), 0 0 30px rgba(139, 92, 246, 0.4) !important;
    color: #fff !important;
}

div.stButton > button:active {
    transform: translateY(-1px) scale(1.02) !important;
}

/* Secondary button style for sidebar */
section[data-testid="stSidebar"] div.stButton > button {
    background: transparent !important;
    color: var(--text-secondary) !important;
    box-shadow: none !important;
}

section[data-testid="stSidebar"] div.stButton > button:hover {
    background: var(--bg-card-hover) !important;
    color: var(--accent-primary) !important;
    transform: translateX(3px) !important;
    box-shadow: 0 4px 12px rgba(6, 182, 212, 0.2) !important;
}

/* ================================================================
   CHAT
   ================================================================ */
[data-testid="stChatMessage"] {
    background: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: var(--radius-md);
    backdrop-filter: blur(20px);
}

/* ================================================================
   TABS - Premium with Animation
   ================================================================ */
.stTabs [data-baseweb="tab-list"] {
    gap: 3px;
    border-bottom: 2px solid var(--border-primary);
    position: relative;
}

.stTabs [data-baseweb="tab-list"]::after {
    content: '';
    position: absolute;
    bottom: -2px;
    left: 0;
    width: 100%;
    height: 2px;
    background: var(--accent-gradient);
    transform: scaleX(0);
    transform-origin: left;
    transition: transform 0.4s ease;
}

.stTabs [data-baseweb="tab"] {
    color: var(--text-secondary);
    font-weight: 600;
    border-radius: var(--radius-sm) var(--radius-sm) 0 0;
    padding: 0.7rem 1.25rem;
    transition: all var(--transition-smooth);
    font-size: 0.88rem;
    position: relative;
}

.stTabs [data-baseweb="tab"]::before {
    content: '';
    position: absolute;
    bottom: 0;
    left: 50%;
    width: 0;
    height: 3px;
    background: var(--accent-gradient);
    transform: translateX(-50%);
    transition: width 0.3s ease;
}

.stTabs [data-baseweb="tab"]:hover::before {
    width: 80%;
}

.stTabs [aria-selected="true"] {
    color: var(--accent-primary) !important;
    background: var(--accent-soft);
}

.stTabs [aria-selected="true"]::before {
    width: 100% !important;
}

/* ================================================================
   EXPANDERS
   ================================================================ */
[data-testid="stExpander"] {
    background: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: var(--radius-md);
    backdrop-filter: blur(10px);
}

/* ================================================================
   INPUTS
   ================================================================ */
input, textarea, .stSelectbox div[data-baseweb="select"] > div {
    background-color: var(--bg-card) !important;
    border-radius: var(--radius-sm) !important;
    border: 1px solid var(--border-primary) !important;
    color: var(--text-primary) !important;
    transition: all var(--transition-fast);
}

input:focus, textarea:focus {
    border-color: var(--accent-primary) !important;
    box-shadow: 0 0 0 2px var(--accent-soft);
}

/* ================================================================
   PROGRESS BARS
   ================================================================ */
.stProgress > div > div > div > div {
    background: var(--accent-gradient) !important;
}

.score-row {
    margin-bottom: 0.75rem;
}

.score-row .score-top {
    display: flex;
    justify-content: space-between;
    font-size: 0.8rem;
    margin-bottom: 0.3rem;
}

.score-row .score-top .score-label {
    color: var(--text-secondary);
    font-weight: 500;
}

.score-row .score-top .score-value {
    color: var(--accent-primary);
    font-weight: 700;
}

.score-track {
    height: 5px;
    border-radius: var(--radius-full);
    background: var(--bg-tertiary);
    overflow: hidden;
}

.score-fill {
    height: 100%;
    border-radius: var(--radius-full);
    background: var(--accent-gradient);
    transition: width var(--transition-slow);
}

.score-fill.risk {
    background: linear-gradient(135deg, var(--danger), var(--warning));
}

/* ================================================================
   TABLES - Premium Modern
   ================================================================ */
.modern-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.88rem;
    margin: 0.5rem 0 1rem 0;
    border-radius: var(--radius-md);
    overflow: hidden;
    box-shadow: var(--shadow-sm);
}

.modern-table th {
    text-align: left;
    color: var(--text-tertiary);
    font-weight: 700;
    text-transform: uppercase;
    font-size: 0.68rem;
    letter-spacing: 0.08em;
    padding: 0.75rem 1rem;
    border-bottom: 2px solid var(--border-accent);
    background: var(--bg-card);
}

.modern-table td {
    padding: 0.75rem 1rem;
    border-bottom: 1px solid var(--border-secondary);
    color: var(--text-primary);
    vertical-align: top;
    transition: all var(--transition-fast);
}

.modern-table tr {
    transition: all var(--transition-fast);
}

.modern-table tr:hover td {
    background: var(--bg-card-hover);
    transform: scale(1.01);
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
}

/* ================================================================
   KANBAN - Premium Cards
   ================================================================ */
.kanban-col {
    background: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: var(--radius-lg);
    padding: 1rem;
    min-height: 140px;
    backdrop-filter: blur(10px);
    box-shadow: var(--shadow-sm);
    transition: all var(--transition-smooth);
}

.kanban-col:hover {
    border-color: var(--border-accent);
    box-shadow: var(--shadow-md);
}

.kanban-col h5 {
    color: var(--text-secondary);
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin: 0 0 0.75rem 0;
    font-weight: 700;
}

.kanban-card {
    background: var(--bg-tertiary);
    border: 1px solid var(--border-secondary);
    border-radius: var(--radius-sm);
    padding: 0.75rem 1rem;
    font-size: 0.85rem;
    margin-bottom: 0.6rem;
    transition: all var(--transition-smooth);
    box-shadow: var(--shadow-sm);
}

.kanban-card:hover {
    border-color: var(--border-accent);
    transform: translateY(-3px) scale(1.02);
    box-shadow: var(--shadow-md);
}

/* ================================================================
   TIMELINE - Premium
   ================================================================ */
.timeline-item {
    border-left: 3px solid var(--accent-gradient);
    padding-left: 1.25rem;
    margin-bottom: 1rem;
    position: relative;
    transition: all var(--transition-smooth);
}

.timeline-item:hover {
    transform: translateX(4px);
}

.timeline-item::before {
    content: "";
    position: absolute;
    left: -7px;
    top: 4px;
    width: 11px;
    height: 11px;
    border-radius: 50%;
    background: var(--accent-gradient);
    box-shadow: 0 0 0 4px var(--bg-deep), 0 0 15px rgba(6, 182, 212, 0.6);
    animation: pulse 2s ease-in-out infinite;
}

@keyframes pulse {
    0%, 100% { box-shadow: 0 0 0 4px var(--bg-deep), 0 0 15px rgba(6, 182, 212, 0.6); }
    50% { box-shadow: 0 0 0 4px var(--bg-deep), 0 0 25px rgba(6, 182, 212, 0.9); }
}

.timeline-item .t-title {
    font-weight: 700;
    color: var(--text-primary);
    font-size: 0.95rem;
}

.timeline-item .t-meta {
    color: var(--text-tertiary);
    font-size: 0.82rem;
    margin-top: 0.2rem;
}

/* ================================================================
   ANIMATIONS
   ================================================================ */
@keyframes fadeSlideUp {
    from {
        opacity: 0;
        transform: translateY(8px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

@keyframes shimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

/* ================================================================
   SCROLLBAR - Premium
   ================================================================ */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-thumb {
    background: var(--accent-gradient);
    border-radius: var(--radius-full);
    box-shadow: 0 0 10px rgba(6, 182, 212, 0.5);
}

::-webkit-scrollbar-thumb:hover {
    background: var(--accent-gradient-reverse);
    box-shadow: 0 0 15px rgba(139, 92, 246, 0.6);
}

::-webkit-scrollbar-track {
    background: var(--bg-primary);
    border-radius: var(--radius-full);
}

/* ================================================================
   UI READABILITY FIX
   ================================================================ */

/* Text inputs */
input, textarea,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input,
[data-testid="stTimeInput"] input,
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea {
    background-color: var(--bg-card) !important;
    color: var(--text-primary) !important;
    border: 1px solid var(--border-primary) !important;
    caret-color: var(--text-primary) !important;
}

input::placeholder, textarea::placeholder {
    color: var(--text-tertiary) !important;
    opacity: 1;
}

/* Select */
[data-baseweb="select"] > div {
    background-color: var(--bg-card) !important;
    border-color: var(--border-primary) !important;
}

[data-baseweb="select"] * {
    color: var(--text-primary) !important;
    fill: var(--text-primary) !important;
}

[data-baseweb="tag"] {
    background: var(--accent-primary) !important;
}

[data-baseweb="tag"] * {
    color: #000 !important;
}

/* Select dropdown */
[data-baseweb="popover"] [data-baseweb="menu"],
[data-baseweb="popover"] ul[role="listbox"] {
    background-color: var(--bg-elevated) !important;
    border: 1px solid var(--border-primary) !important;
    backdrop-filter: blur(20px);
}

[data-baseweb="popover"] li, [role="option"] {
    background-color: transparent !important;
    color: var(--text-primary) !important;
}

[data-baseweb="popover"] li *, [role="option"] * {
    color: var(--text-primary) !important;
}

[role="option"]:hover, li[aria-selected="true"] {
    background-color: var(--accent-soft) !important;
}

/* Checkboxes & radios */
[data-testid="stCheckbox"] label *, [data-testid="stRadio"] label *,
[data-testid="stCheckbox"] p, [data-testid="stRadio"] p {
    color: var(--text-primary) !important;
}

/* Sliders */
[data-testid="stSlider"] label, [data-testid="stSlider"] label *,
[data-testid="stSlider"] div[data-testid="stTickBarMin"],
[data-testid="stSlider"] div[data-testid="stTickBarMax"] {
    color: var(--text-primary) !important;
}

/* File uploader */
[data-testid="stFileUploaderDropzone"] {
    background-color: var(--bg-card) !important;
    border: 1px dashed var(--border-primary) !important;
    border-radius: var(--radius-md) !important;
}

[data-testid="stFileUploaderDropzone"] * {
    color: var(--text-primary) !important;
}

[data-testid="stFileUploaderDropzone"] button {
    background: var(--accent-gradient) !important;
    color: #000 !important;
    border: none !important;
}

[data-testid="stFileUploaderDropzone"] button * {
    color: #000 !important;
}

/* Expanders */
[data-testid="stExpander"] summary, [data-testid="stExpander"] summary * {
    color: var(--text-primary) !important;
}

[data-testid="stExpander"] svg {
    fill: var(--text-secondary) !important;
}

/* Code blocks */
[data-testid="stCodeBlock"], [data-testid="stCodeBlock"] pre, pre, code {
    background-color: var(--bg-secondary) !important;
    color: var(--text-primary) !important;
    border-radius: var(--radius-sm);
    font-family: 'JetBrains Mono', monospace !important;
}

[data-testid="stCodeBlock"] * {
    color: var(--text-primary) !important;
}

/* JSON viewer */
[data-testid="stJson"] {
    background-color: var(--bg-secondary) !important;
    border: 1px solid var(--border-secondary);
    border-radius: var(--radius-sm);
    padding: 0.6rem;
}

[data-testid="stJson"] * {
    color: var(--text-primary) !important;
}

/* Dataframes */
[data-testid="stDataFrame"], [data-testid="stTable"] {
    background-color: var(--bg-card) !important;
    border-radius: var(--radius-sm);
}

[data-testid="stDataFrame"] *, [data-testid="stTable"] * {
    color: var(--text-primary) !important;
}

/* Alerts */
[data-testid="stAlert"] {
    border-radius: var(--radius-md) !important;
    border: 1px solid var(--border-primary) !important;
    background-color: var(--bg-card) !important;
}

[data-testid="stAlert"] p, [data-testid="stAlert"] span, [data-testid="stAlert"] div,
[data-testid="stAlert"] li {
    color: var(--text-primary) !important;
}

[data-testid="stAlertContentError"] {
    background-color: var(--danger-soft) !important;
    border-color: rgba(231, 76, 60, 0.3) !important;
}

[data-testid="stAlertContentWarning"] {
    background-color: var(--warning-soft) !important;
    border-color: rgba(243, 156, 18, 0.3) !important;
}

[data-testid="stAlertContentSuccess"] {
    background-color: var(--success-soft) !important;
    border-color: rgba(46, 204, 113, 0.3) !important;
}

[data-testid="stAlertContentInfo"] {
    background-color: var(--info-soft) !important;
    border-color: rgba(52, 152, 219, 0.3) !important;
}

/* Forms */
[data-testid="stForm"] {
    background: transparent;
    border: none;
}

/* Card content readability */
.content-card > p, .content-card > h4, .content-card > div,
.metric-card .metric-label, .metric-card .metric-value,
.module-card .module-tagline,
.history-item > div {
    color: var(--text-primary);
}

/* ================================================================
   MODERN HERO SECTION
   ================================================================ */
.premium-hero {
    position: relative;
    overflow: hidden;
    padding: 1.5rem 1.5rem;
}

.hero-glow {
    position: absolute;
    top: -50%;
    right: -20%;
    width: 300px;
    height: 300px;
    background: radial-gradient(circle, rgba(59, 130, 246, 0.1) 0%, transparent 70%);
    filter: blur(40px);
}

.hero-content {
    position: relative;
    z-index: 2;
}

.hero-title {
    font-size: 1.75rem !important;
    font-weight: 700 !important;
    margin: 0 0 0.4rem 0 !important;
    background: linear-gradient(135deg, #fff 0%, #3b82f6 100%);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
}

.hero-subtitle {
    font-size: 0.95rem !important;
    color: var(--text-secondary) !important;
    margin: 0 !important;
}

.hero-decoration {
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    pointer-events: none;
    z-index: 1;
}

.deco-circle {
    position: absolute;
    border-radius: 50%;
    background: var(--accent-gradient);
    opacity: 0.08;
}

.deco-1 {
    width: 80px;
    height: 80px;
    top: 10%;
    right: 10%;
}

.deco-2 {
    width: 50px;
    height: 50px;
    bottom: 20%;
    right: 20%;
}

.deco-3 {
    width: 30px;
    height: 30px;
    top: 40%;
    right: 5%;
}

/* ================================================================
   ENHANCED METRIC CARDS
   ================================================================ */
.metric-card {
    position: relative;
    overflow: hidden;
}

.metric-card::after {
    content: '';
    position: absolute;
    top: 0;
    left: -100%;
    width: 100%;
    height: 100%;
    background: linear-gradient(90deg, transparent, rgba(59, 130, 246, 0.08), transparent);
    transition: left 0.5s ease;
}

.metric-card:hover::after {
    left: 100%;
}

.metric-value {
    position: relative;
    z-index: 1;
}

/* ================================================================
   MODERN BUTTONS
   ================================================================ */
.stButton > button, button[kind="primary"] {
    position: relative;
    overflow: hidden;
    transition: all 0.2s ease;
}

.stButton > button::before,
button[kind="primary"]::before {
    content: '';
    position: absolute;
    top: 50%;
    left: 50%;
    width: 0;
    height: 0;
    border-radius: 50%;
    background: rgba(255, 255, 255, 0.15);
    transform: translate(-50%, -50%);
    transition: width 0.4s ease, height 0.4s ease;
}

.stButton > button:hover::before,
button[kind="primary"]:hover::before {
    width: 200px;
    height: 200px;
}

/* ================================================================
   LOADING STATES
   ================================================================ */
@keyframes skeleton-loading {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

.skeleton {
    background: linear-gradient(90deg, var(--bg-card) 25%, var(--bg-card-hover) 50%, var(--bg-card) 75%);
    background-size: 200% 100%;
    animation: skeleton-loading 1.5s infinite;
    border-radius: var(--radius-sm);
}

/* ================================================================
   VISIBILITY ANIMATIONS
   ================================================================ */
.metric-card,
.content-card,
.module-card {
    opacity: 1;
    transform: translateY(0);
    transition: all 0.25s ease;
}

.metric-card.visible,
.content-card.visible,
.module-card.visible {
    opacity: 1;
    transform: translateY(0);
}

/* Staggered animation delays */
.metric-card:nth-child(1) { transition-delay: 0.1s; }
.metric-card:nth-child(2) { transition-delay: 0.2s; }
.metric-card:nth-child(3) { transition-delay: 0.3s; }
.metric-card:nth-child(4) { transition-delay: 0.4s; }

.content-card:nth-child(1) { transition-delay: 0.1s; }
.content-card:nth-child(2) { transition-delay: 0.2s; }
.content-card:nth-child(3) { transition-delay: 0.3s; }

/* ================================================================
   ENHANCED FILE UPLOADER
   ================================================================ */
[data-testid="stFileUploaderDropzone"] {
    position: relative;
    overflow: hidden;
    transition: all 0.3s ease;
}

[data-testid="stFileUploaderDropzone"]::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: var(--accent-gradient);
    opacity: 0;
    transition: opacity 0.3s ease;
    z-index: 0;
}

[data-testid="stFileUploaderDropzone"]:hover::before {
    opacity: 0.05;
}

[data-testid="stFileUploaderDropzone"] > * {
    position: relative;
    z-index: 1;
}

/* ================================================================
   PREMIUM TABLES
   ================================================================ */
.modern-table {
    border-radius: var(--radius-md);
    overflow: hidden;
}

.modern-table thead {
    background: var(--bg-tertiary);
}

.modern-table tbody tr {
    transition: all 0.2s ease;
}

.modern-table tbody tr:hover {
    background: var(--accent-soft);
    transform: scale(1.01);
}

/* ================================================================
   CHAT INTERFACE ENHANCEMENTS
   ================================================================ */
[data-testid="stChatMessage"] {
    transition: all 0.3s ease;
}

[data-testid="stChatMessage"]:hover {
    transform: translateX(5px);
    border-color: var(--border-accent);
}

/* ================================================================
   EXPANDER ENHANCEMENTS
   ================================================================ */
[data-testid="stExpander"] {
    transition: all 0.3s ease;
}

[data-testid="stExpander"]:hover {
    border-color: var(--border-accent);
    box-shadow: var(--shadow-gold);
}

/* ================================================================
   TAB ENHANCEMENTS
   ================================================================ */
.stTabs [data-baseweb="tab"] {
    position: relative;
    transition: all 0.3s ease;
}

.stTabs [data-baseweb="tab"]::after {
    content: '';
    position: absolute;
    bottom: 0;
    left: 50%;
    width: 0;
    height: 2px;
    background: var(--accent-gradient);
    transition: all 0.3s ease;
    transform: translateX(-50%);
}

.stTabs [aria-selected="true"]::after {
    width: 80%;
}

/* ================================================================
   SCROLL ANIMATIONS
   ================================================================ */
@keyframes slideInFromLeft {
    from {
        opacity: 0;
        transform: translateX(-30px);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

@keyframes slideInFromRight {
    from {
        opacity: 0;
        transform: translateX(30px);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

@keyframes zoomIn {
    from {
        opacity: 0;
        transform: scale(0.9);
    }
    to {
        opacity: 1;
        transform: scale(1);
    }
}

.slide-in-left {
    animation: slideInFromLeft 0.6s ease backwards;
}

.slide-in-right {
    animation: slideInFromRight 0.6s ease backwards;
}

.zoom-in {
    animation: zoomIn 0.5s ease backwards;
}

/* ================================================================
   GLOW EFFECTS
   ================================================================ */
.glow-gold {
    box-shadow: 0 0 20px rgba(212, 175, 55, 0.3);
}

.glow-blue {
    box-shadow: 0 0 20px rgba(52, 152, 219, 0.3);
}

/* ================================================================
   GRADIENT BORDERS
   ================================================================ */
.gradient-border {
    position: relative;
    background: var(--bg-card);
    border-radius: var(--radius-md);
}

.gradient-border::before {
    content: '';
    position: absolute;
    top: -2px;
    left: -2px;
    right: -2px;
    bottom: -2px;
    background: var(--accent-gradient);
    border-radius: var(--radius-md);
    z-index: -1;
    opacity: 0;
    transition: opacity 0.3s ease;
}

.gradient-border:hover::before {
    opacity: 1;
}
</style>

<script>
// Premium UI Enhancements
(function() {
    'use strict';
    
    // Wait for DOM to be ready
    function init() {
        // Add smooth scroll behavior
        document.documentElement.style.scrollBehavior = 'smooth';
        
        // Animate metric cards on load
        const metricCards = document.querySelectorAll('.metric-card');
        metricCards.forEach((card, index) => {
            card.style.opacity = '0';
            card.style.transform = 'translateY(20px)';
            setTimeout(() => {
                card.style.transition = 'all 0.6s cubic-bezier(0.4, 0, 0.2, 1)';
                card.style.opacity = '1';
                card.style.transform = 'translateY(0)';
            }, index * 100);
        });
        
        // Add hover effects to buttons
        const buttons = document.querySelectorAll('button[kind="primary"], .stButton > button');
        buttons.forEach(btn => {
            btn.addEventListener('mouseenter', function() {
                this.style.transform = 'translateY(-2px)';
            });
            btn.addEventListener('mouseleave', function() {
                this.style.transform = 'translateY(0)';
            });
        });
        
        // Animate progress bars
        const progressBars = document.querySelectorAll('.stProgress > div > div > div > div');
        progressBars.forEach(bar => {
            const width = bar.style.width;
            bar.style.width = '0%';
            setTimeout(() => {
                bar.style.transition = 'width 1s cubic-bezier(0.4, 0, 0.2, 1)';
                bar.style.width = width;
            }, 300);
        });
        
        // Add loading animation to file uploader
        const fileUploader = document.querySelector('[data-testid="stFileUploader"]');
        if (fileUploader) {
            fileUploader.addEventListener('dragover', function() {
                this.style.borderColor = 'var(--accent-primary)';
                this.style.background = 'var(--accent-soft)';
            });
            fileUploader.addEventListener('dragleave', function() {
                this.style.borderColor = 'var(--border-primary)';
                this.style.background = 'var(--bg-card)';
            });
        }
        
        // Sidebar nav items animation
        const navItems = document.querySelectorAll('.nav-item, section[data-testid="stSidebar"] button');
        navItems.forEach((item, index) => {
            item.style.opacity = '0';
            item.style.transform = 'translateX(-10px)';
            setTimeout(() => {
                item.style.transition = 'all 0.4s cubic-bezier(0.4, 0, 0.2, 1)';
                item.style.opacity = '1';
                item.style.transform = 'translateX(0)';
            }, 500 + (index * 50));
        });
        
        // Add ripple effect to buttons
        function createRipple(event) {
            const button = event.currentTarget;
            const circle = document.createElement('span');
            const diameter = Math.max(button.clientWidth, button.clientHeight);
            const radius = diameter / 2;
            
            circle.style.width = circle.style.height = `${diameter}px`;
            circle.style.left = `${event.clientX - button.getBoundingClientRect().left - radius}px`;
            circle.style.top = `${event.clientY - button.getBoundingClientRect().top - radius}px`;
            circle.classList.add('ripple');
            
            const ripple = button.querySelector('.ripple');
            if (ripple) {
                ripple.remove();
            }
            
            button.appendChild(circle);
        }
        
        buttons.forEach(btn => {
            btn.addEventListener('click', createRipple);
        });
        
        // Auto-hide alerts after 5 seconds
        const alerts = document.querySelectorAll('[data-testid="stAlert"]');
        alerts.forEach(alert => {
            setTimeout(() => {
                alert.style.transition = 'all 0.5s ease';
                alert.style.opacity = '0';
                alert.style.transform = 'translateY(-10px)';
                setTimeout(() => alert.remove(), 500);
            }, 5000);
        });
        
        // Add parallax effect to hero section
        const hero = document.querySelector('.app-hero');
        if (hero) {
            window.addEventListener('scroll', () => {
                const scrolled = window.pageYOffset;
                hero.style.transform = `translateY(${scrolled * 0.3}px)`;
            });
        }
        
        console.log('✨ Premium UI enhancements loaded');
    }
    
    // Run when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
    
    // Re-run on Streamlit reruns
    const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
            if (mutation.addedNodes.length > 0) {
                setTimeout(init, 100);
            }
        });
    });
    
    observer.observe(document.body, { childList: true, subtree: true });
})();
</script>

<style>
/* Ripple Effect */
.ripple {
    position: absolute;
    border-radius: 50%;
    background: rgba(212, 175, 55, 0.4);
    transform: scale(0);
    animation: ripple-animation 0.6s ease-out;
    pointer-events: none;
}

@keyframes ripple-animation {
    to {
        transform: scale(4);
        opacity: 0;
    }
}

/* Enhanced Button Styles */
button[kind="primary"], .stButton > button {
    position: relative;
    overflow: hidden;
}

/* Smooth transitions for all interactive elements */
a, button, input, textarea, select {
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

/* Enhanced focus states */
input:focus, textarea:focus, select:focus {
    outline: none;
    box-shadow: 0 0 0 3px var(--accent-soft);
}

/* Card entrance animation */
@keyframes cardEntrance {
    from {
        opacity: 0;
        transform: translateY(30px) scale(0.95);
    }
    to {
        opacity: 1;
        transform: translateY(0) scale(1);
    }
}

.metric-card, .content-card, .module-card {
    animation: cardEntrance 0.6s cubic-bezier(0.4, 0, 0.2, 1) backwards;
}

.metric-card:nth-child(1) { animation-delay: 0.1s; }
.metric-card:nth-child(2) { animation-delay: 0.2s; }
.metric-card:nth-child(3) { animation-delay: 0.3s; }
.metric-card:nth-child(4) { animation-delay: 0.4s; }

/* Shimmer effect for loading states */
@keyframes shimmer {
    0% { background-position: -1000px 0; }
    100% { background-position: 1000px 0; }
}

.loading-shimmer {
    background: linear-gradient(90deg, var(--bg-card) 0%, var(--bg-card-hover) 50%, var(--bg-card) 100%);
    background-size: 1000px 100%;
    animation: shimmer 2s infinite;
}

/* Enhanced scrollbar */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: var(--bg-secondary);
}

::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg, var(--accent-primary), var(--accent-tertiary));
    border-radius: 10px;
}

::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(180deg, var(--accent-secondary), var(--accent-primary));
}

/* Glassmorphism enhancement */
.glass-effect {
    background: rgba(255, 255, 255, 0.05);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid rgba(255, 255, 255, 0.1);
}

/* Premium shadow effects */
.shadow-premium {
    box-shadow: 0 10px 40px rgba(0, 0, 0, 0.4), 0 0 20px rgba(212, 175, 55, 0.1);
}

/* Gradient text */
.gradient-text {
    background: var(--accent-gradient);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
}

/* Pulse animation for status indicators */
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.5; }
}

.status-pulse {
    animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
}

/* Slide in animation */
@keyframes slideInLeft {
    from {
        opacity: 0;
        transform: translateX(-30px);
    }
    to {
        opacity: 1;
        transform: translateX(0);
    }
}

.slide-in-left {
    animation: slideInLeft 0.5s cubic-bezier(0.4, 0, 0.2, 1) backwards;
}

/* Fade in animation */
@keyframes fadeIn {
    from { opacity: 0; }
    to { opacity: 1; }
}

.fade-in {
    animation: fadeIn 0.6s ease backwards;
}

/* Scale in animation */
@keyframes scaleIn {
    from {
        opacity: 0;
        transform: scale(0.9);
    }
    to {
        opacity: 1;
        transform: scale(1);
    }
}

.scale-in {
    animation: scaleIn 0.4s cubic-bezier(0.4, 0, 0.2, 1) backwards;
}
</style>
"""


def inject_global_css():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    # Add premium JavaScript enhancements
    st.markdown("""
    <script>
    // Premium interactions
    document.addEventListener('DOMContentLoaded', function() {
        // Add smooth reveal animations
        const observerOptions = {
            threshold: 0.1,
            rootMargin: '0px 0px -50px 0px'
        };
        
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                }
            });
        }, observerOptions);
        
        document.querySelectorAll('.metric-card, .content-card, .module-card').forEach(el => {
            observer.observe(el);
        });
    });
    </script>
    """, unsafe_allow_html=True)


def render_hero(title, subtitle):
    st.markdown(
        f'''
        <div class="app-hero premium-hero">
            <div class="hero-glow"></div>
            <div class="hero-content">
                <h1 class="hero-title">{title}</h1>
                <p class="hero-subtitle">{subtitle}</p>
            </div>
            <div class="hero-decoration">
                <div class="deco-circle deco-1"></div>
                <div class="deco-circle deco-2"></div>
                <div class="deco-circle deco-3"></div>
            </div>
        </div>
        ''',
        unsafe_allow_html=True,
    )


def render_metric_card(col, label, value):
    col.markdown(
        f'<div class="metric-card"><div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div></div>',
        unsafe_allow_html=True,
    )


def priority_badge_class(priority_text):
    text = (priority_text or "").lower()
    if "must" in text or "high" in text:
        return "badge-high"
    if "should" in text or "medium" in text:
        return "badge-medium"
    if "could" in text or "low" in text:
        return "badge-low"
    return "badge-neutral"


def render_metric_card_icon(col, icon, label, value, delta=None):
    """Like render_metric_card but with an icon and an optional delta line."""
    delta_html = ""
    if delta:
        delta_class = "up" if str(delta).strip().startswith(("+", "▲")) else "down"
        delta_html = f'<div class="metric-delta {delta_class}">{delta}</div>'
    col.markdown(
        f'<div class="metric-card"><div class="metric-icon">{icon}</div>'
        f'<div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>{delta_html}</div>',
        unsafe_allow_html=True,
    )


def render_score_bar(label, value_pct, risk_style=False):
    """Renders a labeled percentage progress bar using the premium score-track styling."""
    try:
        pct = max(0, min(100, float(value_pct)))
    except (TypeError, ValueError):
        pct = 0
    fill_class = "score-fill risk" if risk_style else "score-fill"
    st.markdown(
        f'<div class="score-row"><div class="score-top">'
        f'<span class="score-label">{label}</span><span class="score-value">{pct:.0f}%</span></div>'
        f'<div class="score-track"><div class="{fill_class}" style="width:{pct:.0f}%;"></div></div></div>',
        unsafe_allow_html=True,
    )


def render_module_header(icon, title, tagline):
    st.markdown(
        f'<div class="module-card"><h3>{icon} {title}</h3>'
        f'<p class="module-tagline">{tagline}</p></div>',
        unsafe_allow_html=True,
    )


def render_modern_table(headers, rows):
    """Renders a list-of-lists as a styled HTML table (used for roadmaps, costs, backlog, etc.)."""
    if not rows:
        st.caption("Not specified in the uploaded SRS.")
        return
    head_html = "".join(f"<th>{h}</th>" for h in headers)
    body_html = ""
    for row in rows:
        cells = "".join(f"<td>{c}</td>" for c in row)
        body_html += f"<tr>{cells}</tr>"
    st.markdown(
        f'<table class="modern-table"><thead><tr>{head_html}</tr></thead>'
        f'<tbody>{body_html}</tbody></table>',
        unsafe_allow_html=True,
    )


def render_kanban_board(kanban):
    columns_def = [("backlog", "📋 Backlog"), ("to_do", "🗒️ To Do"),
                   ("in_progress", "⚙️ In Progress"), ("done", "✅ Done")]
    cols = st.columns(4)
    for col, (key, label) in zip(cols, columns_def):
        items = kanban.get(key, []) if isinstance(kanban, dict) else []
        cards_html = "".join(f'<div class="kanban-card">{item}</div>' for item in items) or \
            '<div class="kanban-card" style="opacity:0.5;">Empty</div>'
        col.markdown(
            f'<div class="kanban-col"><h5>{label}</h5>{cards_html}</div>',
            unsafe_allow_html=True,
        )


def render_timeline(items):
    """items: list of (title, meta) tuples."""
    if not items:
        st.caption("Not specified in the uploaded SRS.")
        return
    html_parts = []
    for title, meta in items:
        html_parts.append(
            f'<div class="timeline-item"><div class="t-title">{title}</div>'
            f'<div class="t-meta">{meta}</div></div>'
        )
    st.markdown('<div class="content-card">' + "".join(html_parts) + '</div>', unsafe_allow_html=True)


# ==============================================================================
# UI — AUTH PAGES
# ==============================================================================

def render_auth_page():
    inject_global_css()
    render_hero(
        f"{APP_ICON} {APP_TITLE}",
        "Enterprise-grade requirement engineering intelligence. Upload, analyze, validate, and export.",
    )

    _, center, _ = st.columns([1, 1.4, 1])
    with center:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        tab_login, tab_signup = st.tabs(["🔐 Log In", "🆕 Sign Up"])

        with tab_login:
            with st.form("login_form", clear_on_submit=False):
                email = st.text_input("Email", key="login_email")
                password = st.text_input("Password", type="password", key="login_password")
                submitted = st.form_submit_button("Log In", use_container_width=True)
            if submitted:
                if not email or not password:
                    st.error("Please enter both email and password.")
                else:
                    with st.spinner("Verifying credentials..."):
                        user, err = login_user(email, password)
                    if user:
                        st.session_state.authenticated = True
                        st.session_state.user = {
                            "id": user["id"],
                            "full_name": user["full_name"],
                            "email": user["email"],
                            "created_at": user["created_at"],
                        }
                        st.success(f"Welcome back, {user['full_name']}!")
                        st.rerun()
                    else:
                        st.error(err)

        with tab_signup:
            with st.form("signup_form", clear_on_submit=False):
                full_name = st.text_input("Full Name", key="signup_name")
                email_su = st.text_input("Email", key="signup_email")
                password_su = st.text_input("Password", type="password", key="signup_password")
                confirm_su = st.text_input("Confirm Password", type="password", key="signup_confirm")
                st.caption(f"Password must be at least {MIN_PASSWORD_LENGTH} characters.")
                submitted_su = st.form_submit_button("Create Account", use_container_width=True)
            if submitted_su:
                is_valid, error_message = validate_signup_inputs(full_name, email_su, password_su, confirm_su)
                if not is_valid:
                    st.error(error_message)
                else:
                    with st.spinner("Creating your account..."):
                        new_user = signup_user(full_name, email_su, password_su)
                    st.success("Account created successfully! Please log in using the Log In tab.")

        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# UI — SIDEBAR
# ==============================================================================

NAV_GROUPS = [
    ("Workspace", [
        ("Upload SRS", "◈"),
        ("Analysis History", ""),
        ("AI Chat", "◉"),
    ]),
    ("Planning & Delivery", [
        ("Developer Roadmap", "🗺️"),
        ("Cost Estimation", "💰"),
        ("Tech Stack Advisor", "🧰"),
        ("System Architect", "🏗️"),
        ("Output Predictor", "🔮"),
    ]),
    ("Quality & Insight", [
        ("Coverage Score", "📈"),
        ("SRS Improver", "✨"),
    ]),
    ("Use Cases, Risk & Team", [
        ("Use Case Generator", ""),
        ("Risk & Feasibility", ""),
        ("Team Roles", "◸"),
    ]),
    ("Execution Tools", [
        ("Project Generator", "🧩"),
        ("Management Dashboard", "🧭"),
        ("Export Center", "⬇️"),
    ]),
    ("Account", [
        ("Reports", "📊"),
        ("Profile", "👤"),
    ]),
]


def render_sidebar():
    user = st.session_state.user
    with st.sidebar:
        st.markdown(
            f'<div class="sidebar-brand">'
            f'<div class="brand-title">{APP_ICON} {APP_TITLE}</div>'
            f'<div class="brand-sub">Enterprise Requirement Intelligence</div></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<div class="user-chip"><div class="u-name">{user["full_name"]}</div>'
            f'<div class="u-email">{user["email"]}</div></div>',
            unsafe_allow_html=True,
        )

        for section_label, items in NAV_GROUPS:
            st.markdown(f'<div class="sidebar-section-label">{section_label}</div>', unsafe_allow_html=True)
            for tab_name, icon in items:
                label = f"{icon}  {tab_name}"
                if st.button(label, key=f"nav_{tab_name}", use_container_width=True):
                    st.session_state.active_tab = tab_name
                    st.rerun()

        st.markdown("---")
        if not get_openai_api_key():
            st.warning("OpenAI API key not set.", icon="⚠️")
        else:
            st.success("OpenAI connected.", icon="✅")

        if st.button("🚪 Log Out", use_container_width=True):
            logout()
            st.rerun()


# ==============================================================================
# UI — DASHBOARD HEADER / METRICS
# ==============================================================================

def render_dashboard_header():
    user = st.session_state.user
    history = list_user_history(user["id"])
    total_uploads = len(history)
    total_analyses = sum(1 for r in history if r.get("analysis"))
    modules_generated = sum(
        1 for r in history for k in NEW_MODULE_DEFAULTS if r.get(k)
    )
    recent = history[0]["filename"] if history else "—"

    render_hero(
        f"Welcome back, {user['full_name'].split(' ')[0]}",
        "Your enterprise requirement intelligence workspace.",
    )
    c1, c2, c3, c4 = st.columns(4)
    render_metric_card_icon(c1, "◈", "Total Uploads", total_uploads)
    render_metric_card_icon(c2, "◉", "Total Analyses", total_analyses)
    render_metric_card_icon(c3, "◫", "Planning Modules Generated", modules_generated)
    render_metric_card_icon(c4, "", "Most Recent File", recent)
    st.markdown("<br>", unsafe_allow_html=True)


# ==============================================================================
# UI — TAB 1: UPLOAD SRS
# ==============================================================================

def process_uploaded_document(uploaded_file, user_id):
    """Runs the full pipeline: parse -> clean -> chunk -> embed -> analyze -> validate -> save."""
    progress = st.progress(0, text="Extracting text...")

    try:
        parsed = parse_uploaded_document(uploaded_file)
    except DocumentParseError as e:
        progress.empty()
        st.error(f"❌ {e}")
        return

    progress.progress(15, text="Cleaning and chunking text...")
    chunks = chunk_text(parsed["text"])
    if not chunks:
        progress.empty()
        st.error("❌ No usable text could be extracted after cleaning. Please try a different file.")
        return

    doc_id = str(uuid.uuid4())[:12]

    progress.progress(35, text="Generating embeddings and building vector store...")
    ok, err = build_and_save_vectorstore(user_id, doc_id, chunks)
    if not ok:
        progress.empty()
        st.error(f"❌ {err}")
        return

    progress.progress(55, text="Running AI requirement analysis (this can take a moment)...")
    analysis, err = run_full_analysis(parsed["text"])
    if err:
        progress.empty()
        st.error(f"❌ {err}")
        return

    progress.progress(80, text="Validating requirement quality...")
    issues, val_err = run_requirement_validation(analysis)
    if val_err:
        st.warning(f"Validation step had an issue: {val_err}")

    progress.progress(95, text="Saving to your history...")
    record = {
        "doc_id": doc_id,
        "filename": parsed["filename"],
        "upload_date": datetime.utcnow().isoformat(),
        "page_count": parsed["page_count"],
        "word_count": parsed["word_count"],
        "analysis": analysis,
        "validation_issues": issues,
        "chat_history": [],
    }
    save_analysis_record(user_id, record)

    # Persist the raw file bytes for traceability (kept inside the user's private folder).
    try:
        upload_path = os.path.join(user_uploads_dir(user_id), doc_id)
        os.makedirs(upload_path, exist_ok=True)
        with open(os.path.join(upload_path, parsed["filename"]), "wb") as f:
            f.write(parsed["raw_bytes"])
    except OSError:
        pass

    progress.progress(100, text="Done!")
    time.sleep(0.4)
    progress.empty()

    st.session_state.active_doc_id = doc_id
    st.success(f"✅ Analysis complete for **{parsed['filename']}**")
    render_analysis_summary(record)


def render_analysis_summary(record):
    st.markdown("#### Document Snapshot")
    c1, c2, c3 = st.columns(3)
    render_metric_card(c1, "Filename", record["filename"])
    render_metric_card(c2, "Pages", record["page_count"])
    render_metric_card(c3, "Word Count", record["word_count"])
    st.markdown("<br>", unsafe_allow_html=True)
    render_full_analysis(record["analysis"], record.get("validation_issues", []))


def render_upload_tab():
    st.markdown("### 📤 Upload SRS Document")
    st.caption("Supported formats: PDF, DOCX, TXT")

    if not get_openai_api_key():
        openai_key_missing_banner()

    uploaded_file = st.file_uploader(
        "Drop your SRS file here",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=False,
    )

    if uploaded_file is not None:
        if st.button("🚀 Analyze Document", use_container_width=True):
            process_uploaded_document(uploaded_file, current_user_id())


# ==============================================================================
# UI — ANALYSIS RENDERING (shared by Upload tab and History tab)
# ==============================================================================

def render_list_card(title, items, empty_text="Not specified in the uploaded SRS."):
    st.markdown(f'<div class="content-card"><h4>{title}</h4>', unsafe_allow_html=True)
    if not items:
        st.caption(empty_text)
    else:
        for item in items:
            if isinstance(item, dict):
                st.markdown("- " + " — ".join(str(v) for v in item.values() if v))
            else:
                st.markdown(f"- {item}")
    st.markdown("</div>", unsafe_allow_html=True)


def render_full_analysis(analysis, validation_issues=None):
    validation_issues = validation_issues or []

    col1, col2 = st.columns(2)
    with col1:
        render_list_card("🎯 Business Goals", analysis.get("business_goals", []))
        render_list_card("⚙️ Functional Requirements", analysis.get("functional_requirements", []))
        render_list_card("🛡️ Non-Functional Requirements", analysis.get("non_functional_requirements", []))
        render_list_card("🚧 Constraints", analysis.get("constraints", []))
        render_list_card("👥 Stakeholders", analysis.get("stakeholders", []))
        render_list_card("🧩 Assumptions", analysis.get("assumptions", []))
        render_list_card("🔗 Dependencies", analysis.get("dependencies", []))
        render_list_card("⚠️ Risks", analysis.get("risks", []))
        render_list_card("❓ Missing Requirements", analysis.get("missing_requirements", []))

    with col2:
        scope = analysis.get("scope", {})
        st.markdown('<div class="content-card"><h4>🗺️ Scope</h4>', unsafe_allow_html=True)
        st.markdown("**In Scope:**")
        for s in scope.get("in_scope", []) or ["Not specified in the uploaded SRS."]:
            st.markdown(f"- {s}")
        st.markdown("**Out of Scope:**")
        for s in scope.get("out_of_scope", []) or ["Not specified in the uploaded SRS."]:
            st.markdown(f"- {s}")
        st.markdown("</div>", unsafe_allow_html=True)

        render_list_card("🧹 Requirement Quality Issues (Summary)", analysis.get("requirement_quality_issues", []))
        render_list_card("💡 Suggested Improvements", analysis.get("suggested_improvements", []))

        st.markdown('<div class="content-card"><h4>📌 Requirement Priority (MoSCoW)</h4>', unsafe_allow_html=True)
        priorities = analysis.get("requirement_priority", [])
        if not priorities:
            st.caption("Not specified in the uploaded SRS.")
        for p in priorities:
            if isinstance(p, dict):
                badge_class = priority_badge_class(p.get("priority", ""))
                st.markdown(
                    f'<span class="badge {badge_class}">{p.get("priority", "N/A")}</span> '
                    f'{p.get("requirement_id", "")}',
                    unsafe_allow_html=True,
                )
        st.markdown("</div>", unsafe_allow_html=True)

        timeline = analysis.get("timeline_estimate", {})
        st.markdown(
            f'<div class="content-card"><h4>⏱️ Timeline Estimate</h4>'
            f'<p><strong>{timeline.get("estimate", "Not specified")}</strong></p>'
            f'<p style="color:#9aa3b5;">{timeline.get("reasoning", "")}</p></div>',
            unsafe_allow_html=True,
        )

        render_list_card("🧑‍💻 Team Recommendation", analysis.get("team_recommendation", []))

        complexity = analysis.get("project_complexity", {})
        badge_class = priority_badge_class(complexity.get("level", ""))
        st.markdown(
            f'<div class="content-card"><h4>📈 Project Complexity</h4>'
            f'<span class="badge {badge_class}">{complexity.get("level", "Unknown")}</span>'
            f'<p style="color:#9aa3b5;margin-top:0.5rem;">{complexity.get("reasoning", "")}</p></div>',
            unsafe_allow_html=True,
        )

        categories = analysis.get("requirement_categories", {})
        st.markdown('<div class="content-card"><h4>🗂️ Requirement Categories</h4>', unsafe_allow_html=True)
        if not categories:
            st.caption("Not specified in the uploaded SRS.")
        for cat_name, cat_items in categories.items():
            st.markdown(f"**{cat_name}**")
            for ci in cat_items:
                st.markdown(f"- {ci}")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        f'<div class="content-card"><h4>📋 Validation Report (Overall)</h4>'
        f'<p>{analysis.get("validation_report", "Not specified in the uploaded SRS.")}</p></div>',
        unsafe_allow_html=True,
    )

    render_validation_issues_section(validation_issues)


def render_validation_issues_section(validation_issues):
    st.markdown("#### 🔍 Detailed Requirement Validation")
    if not validation_issues:
        st.caption("No quality defects detected, or validation has not been run yet.")
        return
    for issue in validation_issues:
        badge_class = {
            "Ambiguous": "badge-medium",
            "Incomplete": "badge-medium",
            "Unmeasurable": "badge-high",
            "Not Testable": "badge-high",
            "Conflicting": "badge-high",
        }.get(issue.get("issue_type"), "badge-neutral")
        confidence_pct = round(issue.get("confidence_score", 0.5) * 100)
        st.markdown(
            f'<div class="content-card">'
            f'<span class="badge {badge_class}">{issue.get("issue_type")}</span> '
            f'<strong>{issue.get("requirement_id")}</strong> '
            f'<span style="color:#9aa3b5;font-size:0.8rem;">confidence {confidence_pct}%</span>'
            f'<p style="margin-top:0.5rem;"><em>Original:</em> {issue.get("original_text")}</p>'
            f'<p><em>Issue:</em> {issue.get("explanation")}</p>'
            f'<p><em>Suggested Fix:</em> {issue.get("corrected_version")}</p>'
            f'</div>',
            unsafe_allow_html=True,
        )


# ==============================================================================
# UI — TAB 2: ANALYSIS HISTORY
# ==============================================================================

def render_history_tab():
    st.markdown("### 🗂️ Analysis History")
    user_id = current_user_id()
    history = list_user_history(user_id)

    if not history:
        st.info("You haven't analyzed any documents yet. Head to the Upload SRS tab to get started.")
        return

    for record in history:
        with st.container():
            st.markdown('<div class="history-item">', unsafe_allow_html=True)
            c1, c2, c3 = st.columns([3, 2, 1])
            with c1:
                st.markdown(f"**{record['filename']}**")
                st.caption(f"Uploaded {record['upload_date'][:19].replace('T', ' ')} UTC")
            with c2:
                st.caption(f"{record['page_count']} pages · {record['word_count']} words")
                st.caption(f"{len(record.get('validation_issues', []))} quality issues flagged")
                generated = [MODULE_BADGE_LABELS[k] for k in NEW_MODULE_DEFAULTS if record.get(k)]
                if generated:
                    st.markdown(
                        " ".join(f'<span class="badge badge-info">{g}</span>' for g in generated),
                        unsafe_allow_html=True,
                    )
            with c3:
                if st.button("Open", key=f"open_{record['doc_id']}"):
                    st.session_state.active_doc_id = record["doc_id"]
                if st.button("Delete", key=f"delete_{record['doc_id']}"):
                    delete_analysis_record(user_id, record["doc_id"])
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.active_doc_id:
        record = load_analysis_record(user_id, st.session_state.active_doc_id)
        if record:
            st.markdown("---")
            st.markdown(f"## 📄 {record['filename']}")
            render_full_analysis(record["analysis"], record.get("validation_issues", []))


# ==============================================================================
# UI — TAB 3: AI CHAT (RAG)
# ==============================================================================

def render_chat_tab():
    st.markdown("### 💬 AI Chat — Ask Your Documents")
    user_id = current_user_id()
    history = list_user_history(user_id)

    if not history:
        st.info("Upload a document first, then come back here to chat with it.")
        return

    if not get_openai_api_key():
        openai_key_missing_banner()

    doc_options = {f"{r['filename']} ({r['upload_date'][:10]})": r["doc_id"] for r in history}
    default_label = None
    if st.session_state.active_doc_id:
        for label, did in doc_options.items():
            if did == st.session_state.active_doc_id:
                default_label = label
    selected_label = st.selectbox(
        "Choose a document to chat with",
        list(doc_options.keys()),
        index=list(doc_options.keys()).index(default_label) if default_label else 0,
    )
    doc_id = doc_options[selected_label]
    st.session_state.active_doc_id = doc_id

    record = load_analysis_record(user_id, doc_id)
    chat_history = record.get("chat_history", []) if record else []

    chat_container = st.container(height=420)
    with chat_container:
        for turn in chat_history:
            with st.chat_message("user" if turn["role"] == "user" else "assistant"):
                st.markdown(turn["content"])

    question = st.chat_input("Ask something about this SRS, e.g. 'What are the business goals?'")
    if question:
        update_chat_history(user_id, doc_id, "user", question)
        with st.spinner("Thinking..."):
            answer, err = answer_question_with_rag(user_id, doc_id, question, chat_history)
        if err:
            st.error(f"❌ {err}")
        else:
            update_chat_history(user_id, doc_id, "assistant", answer)
        st.rerun()

    st.caption(
        "💡 Try: \"What business goals exist?\" · \"What requirement is missing?\" · "
        "\"Generate use cases\""
    )


# ==============================================================================
# UI — TAB 4: REPORTS
# ==============================================================================

def render_reports_tab():
    st.markdown("### 📊 Reports")
    user_id = current_user_id()
    history = list_user_history(user_id)

    if not history:
        st.info("No analyses available yet. Upload an SRS document to generate your first report.")
        return

    doc_options = {f"{r['filename']} ({r['upload_date'][:10]})": r["doc_id"] for r in history}
    selected_label = st.selectbox("Select a document report to export", list(doc_options.keys()))
    doc_id = doc_options[selected_label]
    record = load_analysis_record(user_id, doc_id)

    if record:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        st.markdown(f"**{record['filename']}**")
        st.caption(f"{record['page_count']} pages · {record['word_count']} words · "
                   f"{len(record.get('validation_issues', []))} quality issues flagged")
        generated = [MODULE_BADGE_LABELS[k] for k in NEW_MODULE_DEFAULTS if record.get(k)]
        if generated:
            st.markdown(
                " ".join(f'<span class="badge badge-info">{g}</span>' for g in generated),
                unsafe_allow_html=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        report_json = generate_report_json(record)
        st.download_button(
            label="⬇️ Download JSON Report",
            data=report_json,
            file_name=f"report_{record['filename'].rsplit('.', 1)[0]}_{doc_id}.json",
            mime="application/json",
            use_container_width=True,
        )

        # Persist a copy in the user's private reports folder for traceability.
        try:
            report_path = os.path.join(user_reports_dir(user_id), f"report_{doc_id}.json")
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(report_json)
        except OSError:
            pass

        with st.expander("Preview Report JSON"):
            st.json(json.loads(report_json))


# ==============================================================================
# UI — TAB 5: PROFILE
# ==============================================================================

def render_profile_tab():
    st.markdown("### 👤 Profile")
    user = st.session_state.user
    history = list_user_history(user["id"])

    st.markdown(
        f'<div class="content-card">'
        f'<h4>Account Details</h4>'
        f'<p><strong>Full Name:</strong> {user["full_name"]}</p>'
        f'<p><strong>Email:</strong> {user["email"]}</p>'
        f'<p><strong>Member Since:</strong> {user["created_at"][:10]}</p>'
        f'<p><strong>Total Documents Analyzed:</strong> {len(history)}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown("#### 🔒 Change Password")
    with st.form("change_password_form"):
        current_pw = st.text_input("Current Password", type="password")
        new_pw = st.text_input("New Password", type="password")
        confirm_pw = st.text_input("Confirm New Password", type="password")
        submitted = st.form_submit_button("Update Password")

    if submitted:
        full_user_record = find_user_by_id(user["id"])
        if not full_user_record or not verify_password(current_pw, full_user_record["password_hash"]):
            st.error("Current password is incorrect.")
        elif len(new_pw) < MIN_PASSWORD_LENGTH:
            st.error(f"New password must be at least {MIN_PASSWORD_LENGTH} characters long.")
        elif new_pw != confirm_pw:
            st.error("New passwords do not match.")
        else:
            users = load_users()
            for u in users:
                if u["id"] == user["id"]:
                    u["password_hash"] = hash_password(new_pw)
            save_users(users)
            st.success("Password updated successfully.")


# ==============================================================================
# SHARED HELPERS FOR THE NEW PLANNING / INTELLIGENCE MODULES
# ==============================================================================

def fmt_currency(value):
    try:
        return f"${float(value):,.0f}"
    except (TypeError, ValueError):
        return str(value) if value not in (None, "") else "—"


def priority_html_badge(priority):
    badge_class = priority_badge_class(priority)
    return f'<span class="badge {badge_class}">{priority}</span>'


def get_selected_document(user_id, key_suffix=""):
    """
    Shared document picker used by every new module tab. Returns (doc_id, record),
    or (None, None) if the user has nothing uploaded yet.
    """
    history = list_user_history(user_id)
    if not history:
        st.info("Upload a document in the Upload SRS tab first, then come back here.")
        return None, None

    doc_options = {f"{r['filename']} ({r['upload_date'][:10]})": r["doc_id"] for r in history}
    labels = list(doc_options.keys())
    default_label = None
    if st.session_state.active_doc_id:
        for label, did in doc_options.items():
            if did == st.session_state.active_doc_id:
                default_label = label
    selected_label = st.selectbox(
        "Select a document",
        labels,
        index=labels.index(default_label) if default_label else 0,
        key=f"doc_select_{key_suffix}",
    )
    doc_id = doc_options[selected_label]
    st.session_state.active_doc_id = doc_id
    record = load_analysis_record(user_id, doc_id)
    return doc_id, record


# ==============================================================================
# NEW MODULE 1 & 2 — DEVELOPER EXECUTION ROADMAP + DAILY/WEEKLY/MONTHLY PLANNER
# ==============================================================================

def render_roadmap_tab():
    render_module_header("🗺️", "Developer Execution Roadmap",
                          "Phase-by-phase delivery plan with daily, weekly, and monthly breakdowns.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "roadmap")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    roadmap = record.get("roadmap")
    btn_label = "🔁 Regenerate Roadmap" if roadmap else "🚀 Generate Roadmap"
    if st.button(btn_label, key="gen_roadmap"):
        with st.spinner("Building phases, daily plan, weekly plan, and milestones..."):
            data, err = generate_roadmap(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "roadmap", data)
            st.success("Roadmap generated.")
            st.rerun()

    if not roadmap:
        st.caption("No roadmap generated yet for this document.")
        return

    render_roadmap_progress_tracker(user_id, doc_id, record, roadmap)
    render_roadmap_content(roadmap)


def render_roadmap_progress_tracker(user_id, doc_id, record, roadmap):
    phases = roadmap.get("phases", [])
    if not phases:
        return
    progress = record.get("roadmap_progress") or {"completed": [False] * len(phases)}
    completed = progress.get("completed", [False] * len(phases))
    if len(completed) != len(phases):
        completed = (completed + [False] * len(phases))[:len(phases)]

    st.markdown("#### 📊 Progress Tracker")
    done_count = sum(1 for c in completed if c)
    pct = (done_count / len(phases)) * 100 if phases else 0
    render_score_bar(f"{done_count} / {len(phases)} phases complete", pct)

    with st.expander("✅ Update phase completion"):
        with st.form("roadmap_progress_form"):
            new_completed = []
            for i, p in enumerate(phases):
                checked = st.checkbox(
                    p.get("phase_name", f"Phase {i + 1}"),
                    value=completed[i] if i < len(completed) else False,
                    key=f"phase_chk_{doc_id}_{i}",
                )
                new_completed.append(checked)
            if st.form_submit_button("Save Progress"):
                record["roadmap_progress"] = {"completed": new_completed}
                save_analysis_record(user_id, record)
                st.success("Progress updated.")
                st.rerun()


def render_roadmap_content(roadmap):
    sub_phases, sub_daily, sub_weekly, sub_monthly = st.tabs(
        ["📌 Phases", "🗓️ Daily Plan", "📆 Weekly Plan", "🏁 Monthly Milestones"]
    )

    with sub_phases:
        phases = roadmap.get("phases", [])
        if not phases:
            st.caption("Not specified in the uploaded SRS.")
        for idx, p in enumerate(phases, 1):
            phase_name = p.get("phase_name", f"Phase {idx}")
            tasks = p.get("tasks", []) or ["—"]
            duration = p.get("duration", "—")
            owner = p.get("owner", "—")
            dependencies = p.get("dependencies", "—")
            expected_output = p.get("expected_output", "—")
            
            # Create structured phase card
            st.markdown(
                f'''<div class="content-card">
                    <h4>🔹 {phase_name}</h4>
                    <div style="margin: 0.5rem 0;">
                        <strong style="color: var(--text-secondary); font-size: 0.8rem;">Tasks:</strong>
                        <ul style="margin: 0.3rem 0 0.5rem 1.2rem; padding: 0; color: var(--text-primary);">
                            {"".join(f"<li style='margin-bottom: 0.2rem;'>{t}</li>" for t in tasks)}
                        </ul>
                    </div>
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 0.75rem; padding: 0.5rem 0; border-top: 1px solid var(--border-secondary);">
                        <div>
                            <div style="font-size: 0.65rem; color: var(--text-tertiary); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.2rem;">Duration</div>
                            <div style="font-size: 0.85rem; font-weight: 500; color: var(--text-primary);">{duration}</div>
                        </div>
                        <div>
                            <div style="font-size: 0.65rem; color: var(--text-tertiary); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.2rem;">Owner</div>
                            <div style="font-size: 0.85rem; font-weight: 500; color: var(--text-primary);">{owner}</div>
                        </div>
                        <div>
                            <div style="font-size: 0.65rem; color: var(--text-tertiary); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.2rem;">Dependencies</div>
                            <div style="font-size: 0.85rem; font-weight: 500; color: var(--text-primary);">{dependencies}</div>
                        </div>
                    </div>
                    <div style="padding-top: 0.5rem; border-top: 1px solid var(--border-secondary); margin-top: 0.5rem;">
                        <span style="font-size: 0.75rem; color: var(--text-tertiary);">Expected Output: </span>
                        <span style="font-size: 0.85rem; color: var(--text-primary); font-weight: 500;">{expected_output}</span>
                    </div>
                </div>''',
                unsafe_allow_html=True,
            )

    with sub_daily:
        daily = roadmap.get("daily_plan", [])
        if not daily:
            st.caption("Not specified in the uploaded SRS.")
        for d in daily:
            with st.expander(f"📅 {d.get('day_label', 'Day')} — {d.get('hours', '—')}h"):
                st.markdown("**Tasks:**")
                for t in d.get("tasks", []) or ["—"]:
                    st.markdown(f"- {t}")
                st.markdown("**Deliverables:**")
                for dl in d.get("deliverables", []) or ["—"]:
                    st.markdown(f"- {dl}")
                st.markdown(f"**Meeting Agenda:** {d.get('meeting_agenda', '—')}")
                st.markdown(f"**Expected Output:** {d.get('expected_output', '—')}")

    with sub_weekly:
        weekly = roadmap.get("weekly_plan", [])
        rows = [
            [w.get("week_label", "—"), w.get("focus", "—"), w.get("monday", "—"), w.get("tuesday", "—"),
             w.get("wednesday", "—"), w.get("thursday", "—"), w.get("friday", "—"), w.get("saturday", "—")]
            for w in weekly
        ]
        render_modern_table(["Week", "Focus", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], rows)

    with sub_monthly:
        items = [
            (m.get("month_label", "Month"), f"{m.get('milestone', '')} — Success: {m.get('success_criteria', '')}")
            for m in roadmap.get("monthly_milestones", [])
        ]
        render_timeline(items)


# ==============================================================================
# NEW MODULE 3 — COST ESTIMATION ENGINE
# ==============================================================================

COST_LINE_ITEMS = [
    ("Development", "development_cost"), ("Cloud Hosting (Monthly)", "cloud_cost"),
    ("OpenAI Usage (Monthly)", "openai_usage_cost"), ("Storage (Monthly)", "storage_cost"),
    ("Deployment (One-time)", "deployment_cost"), ("Maintenance (Monthly)", "maintenance_cost"),
    ("Total Estimate", "total_estimate"), ("Monthly Recurring Cost", "monthly_cost"),
    ("Yearly Recurring Cost", "yearly_cost"),
]


def render_cost_tab():
    render_module_header("💰", "Cost Estimation Engine",
                          "Development, cloud, OpenAI usage, storage, deployment & maintenance — Low / Medium / High.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "cost")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    cost = record.get("cost_estimation")
    btn_label = "🔁 Regenerate Estimate" if cost else "🚀 Estimate Costs"
    if st.button(btn_label, key="gen_cost"):
        with st.spinner("Estimating development, cloud, and operating costs..."):
            data, err = generate_cost_estimation(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "cost_estimation", data)
            st.success("Cost estimate generated.")
            st.rerun()

    if not cost:
        st.caption("No cost estimate generated yet for this document.")
        return
    render_cost_content(cost)


def render_cost_content(cost):
    rows = []
    for label, key in COST_LINE_ITEMS:
        item = cost.get(key, {}) or {}
        rows.append([
            label, fmt_currency(item.get("low")), fmt_currency(item.get("medium")),
            fmt_currency(item.get("high")), item.get("notes", "—"),
        ])
    render_modern_table(["Cost Item", "Low", "Medium", "High", "Notes"], rows)

    total = cost.get("total_estimate", {}) or {}
    monthly = cost.get("monthly_cost", {}) or {}
    yearly = cost.get("yearly_cost", {}) or {}
    c1, c2, c3 = st.columns(3)
    render_metric_card_icon(c1, "🧾", "Total One-Time (Medium)", fmt_currency(total.get("medium")))
    render_metric_card_icon(c2, "📅", "Monthly Recurring (Medium)", fmt_currency(monthly.get("medium")))
    render_metric_card_icon(c3, "🗓️", "Yearly Recurring (Medium)", fmt_currency(yearly.get("medium")))

    st.markdown("#### Assumptions")
    for a in cost.get("assumptions", []) or ["Not specified in the uploaded SRS."]:
        st.markdown(f"- {a}")


# ==============================================================================
# NEW MODULE 4 — TECHNOLOGY RECOMMENDATION ENGINE
# ==============================================================================

TECH_LAYERS = [
    ("frontend", "🖥️ Frontend"), ("backend", "⚙️ Backend"), ("database", "🗄️ Database"),
    ("ai_stack", "🤖 AI Stack"), ("cloud", "☁️ Cloud"), ("deployment", "🚀 Deployment"),
    ("security", "🔐 Security"),
]


def render_tech_tab():
    render_module_header("🧰", "Technology Recommendation Engine",
                          "AI-recommended frontend, backend, database, AI stack, cloud, deployment & security.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "tech")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    tech = record.get("tech_recommendation")
    btn_label = "🔁 Regenerate Stack" if tech else "🚀 Recommend Stack"
    if st.button(btn_label, key="gen_tech"):
        with st.spinner("Matching technology choices to your requirements..."):
            data, err = generate_tech_recommendation(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "tech_recommendation", data)
            st.success("Technology stack recommended.")
            st.rerun()

    if not tech:
        st.caption("No technology recommendation generated yet for this document.")
        return
    render_tech_content(tech)


def render_tech_content(tech):
    cols = st.columns(2)
    for idx, (key, label) in enumerate(TECH_LAYERS):
        item = tech.get(key, {}) or {}
        with cols[idx % 2]:
            st.markdown(
                f'<div class="content-card"><h4>{label}: {item.get("technology", "Not specified")}</h4>',
                unsafe_allow_html=True,
            )
            st.markdown(f"**Why:** {item.get('reason', '—')}")
            st.markdown("**Advantages:**")
            for a in item.get("advantages", []) or ["—"]:
                st.markdown(f"- {a}")
            st.markdown("**Tradeoffs:**")
            for t in item.get("tradeoffs", []) or ["—"]:
                st.markdown(f"- {t}")
            st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# NEW MODULE 5 — AUTOMATIC SYSTEM ARCHITECT
# ==============================================================================

def render_architecture_tab():
    render_module_header("🏗️", "Automatic System Architect",
                          "Architecture style, modules, folder structure, data flow, and API suggestions.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "arch")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    arch = record.get("architecture")
    btn_label = "🔁 Regenerate Architecture" if arch else "🚀 Design Architecture"
    if st.button(btn_label, key="gen_arch"):
        with st.spinner("Designing the system architecture..."):
            data, err = generate_architecture(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "architecture", data)
            st.success("Architecture generated.")
            st.rerun()

    if not arch:
        st.caption("No architecture generated yet for this document.")
        return
    render_architecture_content(arch)


def render_architecture_content(arch):
    st.markdown(
        f'<div class="content-card"><h4>🏛️ Architecture Style</h4>'
        f'<p>{arch.get("architecture_style", "Not specified in the uploaded SRS.")}</p></div>',
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        render_list_card("🧩 Modules", arch.get("modules", []))
        st.markdown('<div class="content-card"><h4>📁 Folder Structure</h4>', unsafe_allow_html=True)
        st.code(arch.get("folder_structure", "Not specified in the uploaded SRS."), language="text")
        st.markdown("</div>", unsafe_allow_html=True)
    with col2:
        render_list_card("🔄 Data Flow", arch.get("data_flow", []))
        st.markdown('<div class="content-card"><h4>🔌 API Suggestions</h4>', unsafe_allow_html=True)
        rows = [[a.get("method", "—"), a.get("endpoint", "—"), a.get("purpose", "—")]
                for a in arch.get("api_suggestions", [])]
        render_modern_table(["Method", "Endpoint", "Purpose"], rows)
        st.markdown("</div>", unsafe_allow_html=True)

    micro = arch.get("microservices_recommendation", {}) or {}
    mono = arch.get("monolith_recommendation", {}) or {}
    c1, c2 = st.columns(2)
    with c1:
        badge = "badge-low" if micro.get("recommended") else "badge-neutral"
        verdict = "Recommended" if micro.get("recommended") else "Not Recommended"
        st.markdown(
            f'<div class="content-card"><h4>🧬 Microservices</h4>'
            f'<span class="badge {badge}">{verdict}</span>'
            f'<p style="margin-top:0.6rem;">{micro.get("reasoning", "—")}</p></div>',
            unsafe_allow_html=True,
        )
    with c2:
        badge = "badge-low" if mono.get("recommended") else "badge-neutral"
        verdict = "Recommended" if mono.get("recommended") else "Not Recommended"
        st.markdown(
            f'<div class="content-card"><h4>🧱 Monolith</h4>'
            f'<span class="badge {badge}">{verdict}</span>'
            f'<p style="margin-top:0.6rem;">{mono.get("reasoning", "—")}</p></div>',
            unsafe_allow_html=True,
        )


# ==============================================================================
# NEW MODULE 7 — PROJECT OUTPUT PREDICTOR
# ==============================================================================

def render_output_predictor_tab():
    render_module_header("🔮", "Project Output Predictor",
                          "Predicted deliverables, screens, modules, features, complexity & team size.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "predict")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    pred = record.get("output_prediction")
    btn_label = "🔁 Regenerate Prediction" if pred else "🚀 Predict Output"
    if st.button(btn_label, key="gen_predict"):
        with st.spinner("Predicting project deliverables..."):
            data, err = generate_output_prediction(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "output_prediction", data)
            st.success("Prediction generated.")
            st.rerun()

    if not pred:
        st.caption("No prediction generated yet for this document.")
        return
    render_output_prediction_content(pred)


def render_output_prediction_content(pred):
    col1, col2 = st.columns(2)
    with col1:
        render_list_card("📦 Final Deliverables", pred.get("final_deliverables", []))
        render_list_card("🖼️ Screens", pred.get("screens", []))
        render_list_card("🧩 Modules", pred.get("modules", []))
    with col2:
        render_list_card("⚡ Estimated Features", pred.get("estimated_features", []))

        complexity = pred.get("complexity", {}) or {}
        badge = priority_badge_class(complexity.get("level", ""))
        st.markdown(
            f'<div class="content-card"><h4>📈 Complexity</h4>'
            f'<span class="badge {badge}">{complexity.get("level", "Unknown")}</span>'
            f'<p style="margin-top:0.6rem;">{complexity.get("reasoning", "—")}</p></div>',
            unsafe_allow_html=True,
        )

        readiness = pred.get("deployment_readiness", {}) or {}
        st.markdown(
            f'<div class="content-card"><h4>🚦 Deployment Readiness</h4>'
            f'<span class="badge badge-info">{readiness.get("status", "Unknown")}</span>'
            f'<p style="margin-top:0.6rem;">{readiness.get("reasoning", "—")}</p></div>',
            unsafe_allow_html=True,
        )

        team = pred.get("team_size", {}) or {}
        st.markdown(
            f'<div class="content-card"><h4>👥 Recommended Team Size</h4>'
            f'<p style="font-size:1.6rem;font-weight:800;">{team.get("recommended_size", "—")}</p>',
            unsafe_allow_html=True,
        )
        for b in team.get("breakdown", []) or []:
            if isinstance(b, dict):
                st.markdown(f"- {b.get('role', '—')}: {b.get('count', '—')}")
        st.markdown("</div>", unsafe_allow_html=True)


# ==============================================================================
# NEW MODULE 8 — REQUIREMENT COVERAGE SCORE
# ==============================================================================

def render_coverage_tab():
    render_module_header("📈", "Requirement Coverage Score",
                          "Completeness, clarity, testability, traceability, risk & missing-content scoring.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "coverage")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    score = record.get("coverage_score")
    btn_label = "🔁 Recalculate Score" if score else "🚀 Calculate Coverage Score"
    if st.button(btn_label, key="gen_coverage"):
        with st.spinner("Scoring requirement coverage..."):
            data, err = generate_coverage_score(record["analysis"], record.get("validation_issues", []))
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "coverage_score", data)
            st.success("Coverage score calculated.")
            st.rerun()

    if not score:
        st.caption("No coverage score generated yet for this document.")
        return
    render_coverage_score_content(score)


def render_coverage_score_content(score):
    st.markdown('<div class="content-card"><h4>📋 Scorecard</h4>', unsafe_allow_html=True)
    render_score_bar("Completeness", score.get("completeness_pct", 0))
    render_score_bar("Clarity", score.get("clarity_pct", 0))
    render_score_bar("Testability", score.get("testability_pct", 0))
    render_score_bar("Traceability", score.get("traceability_pct", 0))
    render_score_bar("Risk", score.get("risk_pct", 0), risk_style=True)
    render_score_bar("Missing Content", score.get("missing_pct", 0), risk_style=True)
    st.markdown("</div>", unsafe_allow_html=True)

    overall = score.get("overall_score", 0)
    c1, _ = st.columns([1, 3])
    overall_display = f"{overall:.0f}/100" if isinstance(overall, (int, float)) else str(overall)
    render_metric_card_icon(c1, "🏆", "Overall Score", overall_display)

    st.markdown(
        f'<div class="content-card"><h4>📝 Summary</h4>'
        f'<p>{score.get("summary", "Not specified in the uploaded SRS.")}</p></div>',
        unsafe_allow_html=True,
    )


# ==============================================================================
# NEW MODULE 11 — SRS QUALITY IMPROVER
# ==============================================================================

def render_srs_improver_tab():
    render_module_header("✨", "SRS Quality Improver",
                          "Rewrites weak requirements and assembles a clean, improved SRS section.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "improver")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    improvements = record.get("srs_improvements")
    btn_label = "🔁 Regenerate Improvements" if improvements else "🚀 Improve SRS"
    if st.button(btn_label, key="gen_improve"):
        with st.spinner("Rewriting weak requirements..."):
            data, err = generate_srs_improvements(record["analysis"], record.get("validation_issues", []))
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "srs_improvements", data)
            st.success("SRS improvements generated.")
            st.rerun()

    if not improvements:
        st.caption("No SRS improvements generated yet for this document.")
        return
    render_srs_improvements_content(improvements, record["filename"])


def render_srs_improvements_content(improvements, filename):
    rows = [
        [item.get("original", "—"), item.get("issue", "—"), item.get("suggested", "—"), item.get("reason", "—")]
        for item in improvements.get("improvements", [])
    ]
    st.markdown("#### Before & After")
    render_modern_table(["Original Requirement", "Issue", "Suggested Requirement", "Reason"], rows)

    st.markdown("#### 📄 Improved SRS Document")
    doc_text = improvements.get("improved_srs_document", "Not specified in the uploaded SRS.")
    doc_html = doc_text.replace("\n", "<br>")
    st.markdown(f'<div class="content-card">{doc_html}</div>', unsafe_allow_html=True)

    base_name = filename.rsplit(".", 1)[0] if "." in filename else filename
    st.download_button(
        "⬇️ Download Improved SRS (.txt)", data=doc_text,
        file_name=f"improved_{base_name}.txt", mime="text/plain", key="dl_improved_srs",
    )


# ==============================================================================
# NEW MODULE 12 — SMART PROJECT GENERATOR
# ==============================================================================

def render_project_generator_tab():
    render_module_header("🧩", "Smart Project Generator",
                          "Epics, user stories, sprint plan, kanban board & prioritized backlog.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "smart")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    smart = record.get("smart_project")
    btn_label = "🔁 Regenerate Project Plan" if smart else "🚀 Generate Project Plan"
    if st.button(btn_label, key="gen_smart"):
        with st.spinner("Generating epics, stories, sprints, and a backlog..."):
            data, err = generate_smart_project(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "smart_project", data)
            st.success("Project plan generated.")
            st.rerun()

    if not smart:
        st.caption("No project plan generated yet for this document.")
        return
    render_smart_project_content(smart)


def render_smart_project_content(smart):
    sub_epics, sub_stories, sub_sprints, sub_kanban, sub_backlog = st.tabs(
        ["🎯 Epics", "📝 User Stories", "🏁 Sprint Plan", "🗂️ Kanban", "📚 Backlog"]
    )
    story_lookup = {s.get("id"): s for s in smart.get("user_stories", []) if isinstance(s, dict)}

    with sub_epics:
        rows = [[e.get("id", "—"), e.get("title", "—"), e.get("description", "—")]
                for e in smart.get("epics", [])]
        render_modern_table(["ID", "Title", "Description"], rows)

    with sub_stories:
        stories = smart.get("user_stories", [])
        if not stories:
            st.caption("Not specified in the uploaded SRS.")
        for s in stories:
            preview = (s.get("story", "") or "")[:70]
            with st.expander(f"{s.get('id', 'US')} — {preview}..."):
                st.markdown(f"**Epic:** {s.get('epic_id', '—')}")
                st.markdown(f"**Story:** {s.get('story', '—')}")
                st.markdown("**Acceptance Criteria:**")
                for ac in s.get("acceptance_criteria", []) or ["—"]:
                    st.markdown(f"- {ac}")
                st.markdown(f"**Story Points:** {s.get('story_points', '—')}")

    with sub_sprints:
        sprints = smart.get("sprint_plan", [])
        if not sprints:
            st.caption("Not specified in the uploaded SRS.")
        for sp in sprints:
            story_ids = ", ".join(sp.get("story_ids", []) or [])
            st.markdown(
                f'<div class="content-card"><h4>🏁 {sp.get("sprint", "Sprint")}</h4>'
                f'<p><strong>Goal:</strong> {sp.get("goal", "—")}</p>'
                f'<p><strong>Duration:</strong> {sp.get("duration", "—")}</p>'
                f'<p><strong>Stories:</strong> {story_ids or "—"}</p></div>',
                unsafe_allow_html=True,
            )

    with sub_kanban:
        render_kanban_board(smart.get("kanban", {}))

    with sub_backlog:
        rows = []
        for b in smart.get("backlog", []):
            story = story_lookup.get(b.get("id"), {})
            story_text = (story.get("story", "—") or "—")[:80]
            rows.append([b.get("id", "—"), priority_html_badge(b.get("priority", "—")), story_text])
        render_modern_table(["ID", "Priority", "Story"], rows)


# ==============================================================================
# NEW MODULE 13 — USE CASE GENERATOR
# ==============================================================================

def render_use_case_tab():
    render_module_header("🧾", "Use Case Generator",
                          "Actors, triggers, flows, business rules, scenarios, a use case diagram & table.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "usecase")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    use_cases = record.get("use_cases")
    btn_label = "🔁 Regenerate Use Cases" if use_cases else "🚀 Generate Use Cases"
    if st.button(btn_label, key="gen_usecases"):
        with st.spinner("Deriving actors, flows, and scenarios from your requirements..."):
            data, err = generate_use_cases(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "use_cases", data)
            st.success("Use cases generated.")
            st.rerun()

    if not use_cases:
        st.caption("No use cases generated yet for this document.")
        return
    render_use_case_content(use_cases)


def render_use_case_content(data):
    c1, c2, c3 = st.columns(3)
    render_metric_card_icon(c1, "📋", "Use Cases", len(data.get("use_cases", []) or []))
    render_metric_card_icon(c2, "🧑", "Primary Actors", len(data.get("primary_actors", []) or []))
    render_metric_card_icon(c3, "🔌", "Secondary Actors", len(data.get("secondary_actors", []) or []))

    sub_diagram, sub_table, sub_detail, sub_actors = st.tabs(
        ["🗺️ Diagram", "📊 Use Case Table", "🔎 Detailed Flows", "👥 Actors"]
    )

    with sub_diagram:
        render_use_case_diagram(data.get("diagram_data", {}))

    with sub_table:
        rows = [[r.get("id", "—"), r.get("title", "—"), r.get("primary_actor", "—"),
                 r.get("secondary_actors", "—"), r.get("trigger", "—")]
                for r in data.get("use_case_table", [])]
        render_modern_table(["ID", "Title", "Primary Actor", "Secondary Actors", "Trigger"], rows)

    with sub_detail:
        use_cases = data.get("use_cases", [])
        if not use_cases:
            st.caption("Not specified in the uploaded SRS.")
        for uc in use_cases:
            if not isinstance(uc, dict):
                continue
            with st.expander(f"{uc.get('id', 'UC')} — {uc.get('title', 'Use Case')}"):
                c1, c2 = st.columns(2)
                c1.markdown(f"**Primary Actor:** {uc.get('primary_actor', '—')}")
                c2.markdown(f"**Secondary Actors:** {', '.join(uc.get('secondary_actors', []) or ['—'])}")
                st.markdown(f"**Trigger:** {uc.get('trigger', '—')}")

                cc1, cc2 = st.columns(2)
                with cc1:
                    st.markdown("**Inputs:**")
                    for i in uc.get("inputs", []) or ["—"]:
                        st.markdown(f"- {i}")
                    st.markdown("**Preconditions:**")
                    for i in uc.get("preconditions", []) or ["—"]:
                        st.markdown(f"- {i}")
                with cc2:
                    st.markdown("**Outputs:**")
                    for i in uc.get("outputs", []) or ["—"]:
                        st.markdown(f"- {i}")
                    st.markdown("**Postconditions:**")
                    for i in uc.get("postconditions", []) or ["—"]:
                        st.markdown(f"- {i}")

                st.markdown("**Main Flow:**")
                for i, step in enumerate(uc.get("main_flow", []) or ["—"], 1):
                    st.markdown(f"{i}. {step}")
                st.markdown("**Alternative Flow:**")
                for i in uc.get("alternative_flow", []) or ["—"]:
                    st.markdown(f"- {i}")
                st.markdown("**Exception Flow:**")
                for i in uc.get("exception_flow", []) or ["—"]:
                    st.markdown(f"- {i}")
                st.markdown("**Business Rules:**")
                for i in uc.get("business_rules", []) or ["—"]:
                    st.markdown(f"- {i}")

                st.markdown(
                    f'<div class="content-card" style="margin-top:0.5rem;">'
                    f'<p><span class="badge badge-low">Success</span> {uc.get("success_scenario", "—")}</p>'
                    f'<p><span class="badge badge-high">Failure</span> {uc.get("failure_scenario", "—")}</p>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    with sub_actors:
        col1, col2 = st.columns(2)
        with col1:
            render_list_card("🧑 Primary Actors", data.get("primary_actors", []))
        with col2:
            render_list_card("🔌 Secondary Actors", data.get("secondary_actors", []))


def render_use_case_diagram(diagram_data):
    """Renders a lightweight inline SVG UML-style use case diagram from derived diagram data."""
    use_cases = diagram_data.get("use_cases", []) if diagram_data else []
    if not use_cases:
        st.caption("Not specified in the uploaded SRS.")
        return

    actors = diagram_data.get("actors", [])
    relationships = diagram_data.get("relationships", [])
    primary_actor_names = {r["actor"] for r in relationships if r.get("type") == "primary"}
    left_actors = [a for a in actors if a in primary_actor_names] or actors[:max(1, len(actors) // 2)] or ["Actor"]
    right_actors = [a for a in actors if a not in left_actors]

    width = 760
    row_h = 60
    height = max(260, len(use_cases) * row_h + 80)
    boundary_x1, boundary_x2 = 220, width - 220
    uc_cx = (boundary_x1 + boundary_x2) // 2
    uc_w, uc_h = boundary_x2 - boundary_x1 - 30, 44

    def positions(names, x):
        n = len(names) or 1
        step = height / (n + 1)
        return {name: (x, int(step * (i + 1))) for i, name in enumerate(names)}

    left_pos = positions(left_actors, 60)
    right_pos = positions(right_actors, width - 60) if right_actors else {}

    uc_pos, start_y = {}, 50
    for i, uc in enumerate(use_cases):
        uc_pos[uc.get("id") or f"UC-{i + 1}"] = (uc_cx, start_y + i * row_h + row_h // 2)

    def esc(s):
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    parts = [f'<svg viewBox="0 0 {width} {height}" xmlns="http://www.w3.org/2000/svg" '
             f'style="width:100%;height:auto;display:block;">']
    parts.append(f'<rect x="0" y="0" width="{width}" height="{height}" fill="#11142a" rx="14"/>')
    parts.append(f'<rect x="{boundary_x1}" y="16" width="{boundary_x2 - boundary_x1}" height="{height - 32}" '
                 f'rx="14" fill="none" stroke="#7c6cf6" stroke-width="1.5" stroke-dasharray="5 4"/>')

    for rel in relationships:
        a_pos = left_pos.get(rel.get("actor")) or right_pos.get(rel.get("actor"))
        u_pos = uc_pos.get(rel.get("use_case"))
        if a_pos and u_pos:
            stroke = "#4fd1c5" if rel.get("type") == "primary" else "rgba(255,255,255,0.35)"
            end_x = u_pos[0] - uc_w // 2 if a_pos[0] < uc_cx else u_pos[0] + uc_w // 2
            parts.append(f'<line x1="{a_pos[0]}" y1="{a_pos[1]}" x2="{end_x}" y2="{u_pos[1]}" '
                         f'stroke="{stroke}" stroke-width="1.4"/>')

    def draw_actor(name, pos, anchor):
        x, y = pos
        label_x = x + (16 if anchor == "start" else -16)
        return (
            f'<circle cx="{x}" cy="{y - 18}" r="7" fill="#f3f4f8"/>'
            f'<line x1="{x}" y1="{y - 11}" x2="{x}" y2="{y + 8}" stroke="#f3f4f8" stroke-width="2"/>'
            f'<line x1="{x - 10}" y1="{y - 2}" x2="{x + 10}" y2="{y - 2}" stroke="#f3f4f8" stroke-width="2"/>'
            f'<line x1="{x}" y1="{y + 8}" x2="{x - 9}" y2="{y + 22}" stroke="#f3f4f8" stroke-width="2"/>'
            f'<line x1="{x}" y1="{y + 8}" x2="{x + 9}" y2="{y + 22}" stroke="#f3f4f8" stroke-width="2"/>'
            f'<text x="{label_x}" y="{y + 5}" fill="#f3f4f8" font-size="11" text-anchor="{anchor}" '
            f'font-family="Inter, sans-serif">{esc(name)[:18]}</text>'
        )

    for name, pos in left_pos.items():
        parts.append(draw_actor(name, pos, "start"))
    for name, pos in right_pos.items():
        parts.append(draw_actor(name, pos, "end"))

    for i, uc in enumerate(use_cases):
        cx, cy = uc_pos[uc.get("id") or f"UC-{i + 1}"]
        title = esc(uc.get("title") or uc.get("id") or "")[:32]
        parts.append(
            f'<ellipse cx="{cx}" cy="{cy}" rx="{uc_w // 2}" ry="{uc_h // 2}" '
            f'fill="rgba(124,108,246,0.18)" stroke="#a499ff" stroke-width="1.4"/>'
            f'<text x="{cx}" y="{cy + 4}" fill="#f3f4f8" font-size="11.5" text-anchor="middle" '
            f'font-family="Inter, sans-serif">{title}</text>'
        )

    parts.append("</svg>")
    st.markdown('<div class="content-card">' + "".join(parts) + "</div>", unsafe_allow_html=True)
    st.caption("Teal lines = primary actor · grey lines = secondary actor.")


# ==============================================================================
# NEW MODULE 14 — RISK & FEASIBILITY ANALYZER
# ==============================================================================

def render_risk_tab():
    render_module_header("⚠️", "Risk & Feasibility Analyzer",
                          "Technical, budget, schedule, operational, security, maintainability & "
                          "scalability risk — plus feasibility, complexity & recommendation.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "risk")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    risk = record.get("risk_feasibility")
    btn_label = "🔁 Regenerate Assessment" if risk else "🚀 Analyze Risk & Feasibility"
    if st.button(btn_label, key="gen_risk"):
        with st.spinner("Assessing risk dimensions and overall feasibility..."):
            data, err = generate_risk_feasibility(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "risk_feasibility", data)
            st.success("Risk & feasibility assessment generated.")
            st.rerun()

    if not risk:
        st.caption("No risk & feasibility assessment generated yet for this document.")
        return
    render_risk_content(risk)


def risk_level_badge_class(level_text):
    text = (level_text or "").lower()
    if "high" in text:
        return "badge-high"
    if "medium" in text:
        return "badge-medium"
    if "low" in text:
        return "badge-low"
    return "badge-neutral"


def render_risk_content(risk):
    feasibility = risk.get("feasibility_score", 0)
    complexity = risk.get("complexity_score", 0)
    c1, c2 = st.columns(2)
    with c1:
        render_score_bar("Feasibility Score", feasibility)
    with c2:
        render_score_bar("Complexity Score", complexity, risk_style=True)

    st.markdown("#### Risk Dimensions")
    cols = st.columns(2)
    for idx, (key, label) in enumerate(RISK_CATEGORIES):
        item = risk.get(key, {}) or {}
        with cols[idx % 2]:
            prob_badge = risk_level_badge_class(item.get("probability", ""))
            impact_badge = risk_level_badge_class(item.get("impact", ""))
            sev_badge = risk_level_badge_class(item.get("severity", ""))
            st.markdown(
                f'<div class="content-card"><h4>{label}</h4>'
                f'<span class="badge {prob_badge}">Probability: {item.get("probability", "—")}</span> '
                f'<span class="badge {impact_badge}">Impact: {item.get("impact", "—")}</span> '
                f'<span class="badge {sev_badge}">Severity: {item.get("severity", "—")}</span>'
                f'<p style="margin-top:0.6rem;">{item.get("description", "Not specified in the uploaded SRS.")}</p>'
                f'<p><em>Mitigation:</em> {item.get("mitigation", "—")}</p></div>',
                unsafe_allow_html=True,
            )

    st.markdown(
        f'<div class="content-card"><h4>🧭 Overall Recommendation</h4>'
        f'<p>{risk.get("recommendation", "Not specified in the uploaded SRS.")}</p></div>',
        unsafe_allow_html=True,
    )


# ==============================================================================
# NEW MODULE 15 — TEAM ROLE GENERATOR
# ==============================================================================

def render_team_tab():
    render_module_header("👥", "Team Role Generator",
                          "Role-by-role responsibilities, hours, cost, skills & timeline — plus "
                          "resource planning, workload distribution, team size & total cost.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "team")
    if record is None:
        return
    if not get_openai_api_key():
        openai_key_missing_banner()

    team = record.get("team_roles")
    btn_label = "🔁 Regenerate Team Plan" if team else "🚀 Generate Team Plan"
    if st.button(btn_label, key="gen_team"):
        with st.spinner("Planning roles, hours, costs, and workload distribution..."):
            data, err = generate_team_roles(record["analysis"])
        if err:
            st.error(f"❌ {err}")
        else:
            save_module_result(user_id, doc_id, "team_roles", data)
            st.success("Team plan generated.")
            st.rerun()

    if not team:
        st.caption("No team plan generated yet for this document.")
        return
    render_team_content(team)


def render_team_content(team):
    planning = team.get("team_planning", {}) or {}
    c1, c2, c3 = st.columns(3)
    render_metric_card_icon(c1, "👥", "Team Size", planning.get("team_size", "—"))
    render_metric_card_icon(c2, "💵", "Total Team Cost", fmt_currency(planning.get("team_cost")))
    render_metric_card_icon(c3, "📋", "Roles Defined", len(team.get("roles", {}) or {}))

    st.markdown("#### Role Breakdown")
    roles = team.get("roles", {}) or {}
    cols = st.columns(2)
    for idx, (key, label) in enumerate(TEAM_ROLES):
        role = roles.get(key, {}) or {}
        with cols[idx % 2]:
            needed = role.get("needed", True)
            badge = '<span class="badge badge-low">Needed</span>' if needed else \
                    '<span class="badge badge-neutral">Not Needed</span>'
            st.markdown(
                f'<div class="content-card"><h4>{label}</h4>{badge}'
                f'<p style="margin-top:0.6rem;"><strong>Hours/Week:</strong> {role.get("hours_per_week", "—")} '
                f'&nbsp;·&nbsp; <strong>Cost:</strong> {fmt_currency(role.get("estimated_cost"))} '
                f'&nbsp;·&nbsp; <strong>Timeline:</strong> {role.get("timeline", "—")}</p>',
                unsafe_allow_html=True,
            )
            st.markdown("**Responsibilities:**")
            for r in role.get("responsibilities", []) or ["—"]:
                st.markdown(f"- {r}")
            st.markdown("**Deliverables:**")
            for d in role.get("deliverables", []) or ["—"]:
                st.markdown(f"- {d}")
            st.markdown("**Skills:**")
            for s in role.get("skills", []) or ["—"]:
                st.markdown(f"- {s}")
            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("#### ⏱️ Hours Per Person")
    rows = [[h.get("role", "—"), h.get("hours_per_week", "—")] for h in planning.get("hours_per_person", [])]
    render_modern_table(["Role", "Hours / Week"], rows)

    st.markdown("#### 📊 Workload Distribution")
    workload = team.get("workload_distribution", []) or []
    if not workload:
        st.caption("Not specified in the uploaded SRS.")
    for w in workload:
        if isinstance(w, dict):
            render_score_bar(w.get("role", "—"), w.get("percentage", 0))

    st.markdown(
        f'<div class="content-card"><h4>🧭 Resource Planning</h4>'
        f'<p>{team.get("resource_planning", "Not specified in the uploaded SRS.")}</p></div>',
        unsafe_allow_html=True,
    )


# ==============================================================================
# NEW MODULE 10 — MANAGEMENT DASHBOARD
# ==============================================================================

def render_management_dashboard_tab():
    render_module_header("🧭", "Management Dashboard",
                          "Project health, budget, progress, risk, team capacity & deadline — at a glance.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "mgmt")
    if record is None:
        return

    cost = record.get("cost_estimation")
    coverage = record.get("coverage_score")
    prediction = record.get("output_prediction")
    roadmap = record.get("roadmap")
    use_cases = record.get("use_cases")
    risk = record.get("risk_feasibility")
    team = record.get("team_roles")

    missing_labels = []
    if not cost:
        missing_labels.append("Cost Estimation")
    if not coverage:
        missing_labels.append("Coverage Score")
    if not prediction:
        missing_labels.append("Output Predictor")
    if not roadmap:
        missing_labels.append("Developer Roadmap")
    if not use_cases:
        missing_labels.append("Use Case Generator")
    if not risk:
        missing_labels.append("Risk & Feasibility")
    if not team:
        missing_labels.append("Team Roles")
    if missing_labels:
        st.info(
            "This dashboard pulls live data from other modules. For a complete picture, generate: "
            + ", ".join(missing_labels)
        )

    health_score = coverage.get("overall_score") if coverage else None
    budget_val = (cost.get("total_estimate", {}) or {}).get("medium") if cost else None
    risk_val = coverage.get("risk_pct") if coverage else None
    team_val = (prediction.get("team_size", {}) or {}).get("recommended_size") if prediction else None

    c1, c2, c3, c4 = st.columns(4)
    render_metric_card_icon(c1, "❤️", "Project Health",
                             f"{health_score:.0f}/100" if isinstance(health_score, (int, float)) else "—")
    render_metric_card_icon(c2, "💰", "Budget (Medium Est.)",
                             fmt_currency(budget_val) if budget_val is not None else "—")
    render_metric_card_icon(c3, "⚠️", "Risk Level",
                             f"{risk_val:.0f}%" if isinstance(risk_val, (int, float)) else "—")
    render_metric_card_icon(c4, "👥", "Team Capacity Needed", team_val if team_val is not None else "—")

    st.markdown("#### 📊 Progress Snapshot")
    phases = roadmap.get("phases", []) if roadmap else []
    if phases:
        progress_data = record.get("roadmap_progress", {}) or {}
        completed = progress_data.get("completed", [])
        done_count = sum(1 for c in completed if c) if completed else 0
        pct = (done_count / len(phases)) * 100 if phases else 0
        render_score_bar(f"Roadmap Phases Completed ({done_count}/{len(phases)})", pct)
    else:
        st.caption("Generate a Developer Roadmap to see phase progress here.")

    deadline_text = (record.get("analysis", {}).get("timeline_estimate", {}) or {}).get(
        "estimate", "Not specified in the uploaded SRS."
    )
    st.markdown(
        f'<div class="content-card"><h4>⏰ Deadline / Timeline Estimate</h4><p>{deadline_text}</p></div>',
        unsafe_allow_html=True,
    )

    st.markdown("#### 🧾 Use Cases · ⚠️ Risk & Feasibility · 👥 Team Plan")
    uc_count = len(use_cases.get("use_cases", [])) if use_cases else None
    feasibility_val = risk.get("feasibility_score") if risk else None
    complexity_val = risk.get("complexity_score") if risk else None
    team_planning = team.get("team_planning", {}) if team else {}
    team_size_val = team_planning.get("team_size") if team else None
    team_cost_val = team_planning.get("team_cost") if team else None

    d1, d2, d3, d4, d5 = st.columns(5)
    render_metric_card_icon(d1, "📋", "Use Cases", uc_count if uc_count is not None else "—")
    render_metric_card_icon(d2, "🛡️", "Feasibility",
                             f"{feasibility_val:.0f}/100" if isinstance(feasibility_val, (int, float)) else "—")
    render_metric_card_icon(d3, "🧮", "Complexity",
                             f"{complexity_val:.0f}/100" if isinstance(complexity_val, (int, float)) else "—")
    render_metric_card_icon(d4, "👥", "Team Size", team_size_val if team_size_val is not None else "—")
    render_metric_card_icon(d5, "💵", "Team Cost",
                             fmt_currency(team_cost_val) if team_cost_val is not None else "—")


# ==============================================================================
# NEW MODULE 9 — EXPORT CENTER (PDF, DOCX, JSON, Developer Plan, Timeline, Architecture)
# ==============================================================================

def build_docx_report(record):
    """Builds a formatted Word document from a stored analysis record. Returns bytes."""
    document = docx.Document()
    title = document.add_heading(f"SRS Intelligence Report — {record.get('filename', 'Document')}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x4F, 0x46, 0xE5)
        run.font.size = Pt(22)
    document.add_paragraph(f"Generated: {datetime.utcnow().isoformat()}Z")
    meta_p = document.add_paragraph()
    meta_p.add_run(f"Pages: {record.get('page_count', '—')}    Word Count: {record.get('word_count', '—')}")

    analysis = record.get("analysis", {}) or {}

    def add_list_section(heading_text, items):
        document.add_heading(heading_text, level=1)
        if not items:
            document.add_paragraph("Not specified in the uploaded SRS.")
            return
        for item in items:
            if isinstance(item, dict):
                text = " — ".join(str(v) for v in item.values() if v)
            else:
                text = str(item)
            document.add_paragraph(text, style="List Bullet")

    add_list_section("Business Goals", analysis.get("business_goals", []))
    add_list_section("Functional Requirements", analysis.get("functional_requirements", []))
    add_list_section("Non-Functional Requirements", analysis.get("non_functional_requirements", []))
    add_list_section("Constraints", analysis.get("constraints", []))
    add_list_section("Stakeholders", analysis.get("stakeholders", []))
    add_list_section("Risks", analysis.get("risks", []))
    add_list_section("Missing Requirements", analysis.get("missing_requirements", []))
    add_list_section("Suggested Improvements", analysis.get("suggested_improvements", []))

    document.add_heading("Validation Report", level=1)
    document.add_paragraph(analysis.get("validation_report", "Not specified in the uploaded SRS."))

    issues = record.get("validation_issues", []) or []
    if issues:
        document.add_heading("Detected Requirement Issues", level=1)
        for i in issues:
            p = document.add_paragraph()
            p.add_run(f"[{i.get('issue_type')}] {i.get('requirement_id')}: ").bold = True
            p.add_run(i.get("explanation", ""))
            document.add_paragraph(f"Suggested fix: {i.get('corrected_version', '')}", style="List Bullet")

    roadmap = record.get("roadmap")
    if roadmap:
        document.add_heading("Developer Roadmap", level=1)
        for p in roadmap.get("phases", []):
            document.add_heading(p.get("phase_name", "Phase"), level=2)
            document.add_paragraph(f"Duration: {p.get('duration', '—')}   Owner: {p.get('owner', '—')}")
            for t in p.get("tasks", []) or []:
                document.add_paragraph(t, style="List Bullet")
            document.add_paragraph(f"Expected Output: {p.get('expected_output', '—')}")

    cost = record.get("cost_estimation")
    if cost:
        document.add_heading("Cost Estimation", level=1)
        table = document.add_table(rows=1, cols=4)
        table.style = "Light Grid Accent 1"
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = "Item", "Low", "Medium", "High"
        for label, key in COST_LINE_ITEMS:
            item = cost.get(key, {}) or {}
            row = table.add_row().cells
            row[0].text = label
            row[1].text = fmt_currency(item.get("low"))
            row[2].text = fmt_currency(item.get("medium"))
            row[3].text = fmt_currency(item.get("high"))

    tech = record.get("tech_recommendation")
    if tech:
        document.add_heading("Technology Recommendation", level=1)
        for key, label in TECH_LAYERS:
            item = tech.get(key, {}) or {}
            document.add_heading(f"{label}: {item.get('technology', '—')}", level=2)
            document.add_paragraph(item.get("reason", ""))

    architecture = record.get("architecture")
    if architecture:
        document.add_heading("System Architecture", level=1)
        document.add_paragraph(f"Style: {architecture.get('architecture_style', '—')}")
        for m in architecture.get("modules", []):
            if isinstance(m, dict):
                document.add_paragraph(f"{m.get('name', '')}: {m.get('responsibility', '')}", style="List Bullet")

    use_cases = record.get("use_cases")
    if use_cases:
        document.add_heading("Use Cases", level=1)
        for uc in use_cases.get("use_cases", []) or []:
            if not isinstance(uc, dict):
                continue
            document.add_heading(f"{uc.get('id', 'UC')}: {uc.get('title', '')}", level=2)
            document.add_paragraph(
                f"Primary Actor: {uc.get('primary_actor', '—')}    "
                f"Secondary Actors: {', '.join(uc.get('secondary_actors', []) or ['—'])}"
            )
            document.add_paragraph(f"Trigger: {uc.get('trigger', '—')}")
            document.add_paragraph("Main Flow:")
            for step in uc.get("main_flow", []) or []:
                document.add_paragraph(step, style="List Bullet")
            document.add_paragraph(f"Success Scenario: {uc.get('success_scenario', '—')}")
            document.add_paragraph(f"Failure Scenario: {uc.get('failure_scenario', '—')}")

    risk = record.get("risk_feasibility")
    if risk:
        document.add_heading("Risk & Feasibility Analysis", level=1)
        document.add_paragraph(
            f"Feasibility Score: {risk.get('feasibility_score', '—')}/100    "
            f"Complexity Score: {risk.get('complexity_score', '—')}/100"
        )
        for key, label in RISK_CATEGORIES:
            item = risk.get(key, {}) or {}
            document.add_heading(label, level=2)
            document.add_paragraph(
                f"Probability: {item.get('probability', '—')}   Impact: {item.get('impact', '—')}   "
                f"Severity: {item.get('severity', '—')}"
            )
            document.add_paragraph(item.get("description", ""))
            document.add_paragraph(f"Mitigation: {item.get('mitigation', '—')}")
        document.add_heading("Recommendation", level=2)
        document.add_paragraph(risk.get("recommendation", "Not specified in the uploaded SRS."))

    team = record.get("team_roles")
    if team:
        document.add_heading("Team Roles & Resource Planning", level=1)
        planning = team.get("team_planning", {}) or {}
        document.add_paragraph(
            f"Team Size: {planning.get('team_size', '—')}    Team Cost: {fmt_currency(planning.get('team_cost'))}"
        )
        for key, label in TEAM_ROLES:
            role = team.get("roles", {}).get(key, {}) or {}
            document.add_heading(label, level=2)
            document.add_paragraph(
                f"Hours/Week: {role.get('hours_per_week', '—')}   "
                f"Cost: {fmt_currency(role.get('estimated_cost'))}   Timeline: {role.get('timeline', '—')}"
            )
            for r in role.get("responsibilities", []) or []:
                document.add_paragraph(r, style="List Bullet")
        document.add_heading("Resource Planning", level=2)
        document.add_paragraph(team.get("resource_planning", "Not specified in the uploaded SRS."))

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_pdf_report(record):
    """Builds a formatted PDF report from a stored analysis record using reportlab. Returns bytes."""
    buffer = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.7 * inch, bottomMargin=0.7 * inch)
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    heading_style = styles["Heading2"]
    body_style = styles["Normal"]
    bullet_style = ParagraphStyle("BulletStyle", parent=body_style, leftIndent=14, bulletIndent=4)

    story = [
        Paragraph(f"SRS Intelligence Report — {record.get('filename', 'Document')}", title_style),
        Paragraph(f"Generated: {datetime.utcnow().isoformat()}Z", body_style),
        Spacer(1, 12),
    ]

    analysis = record.get("analysis", {}) or {}

    def add_section(heading_text, items):
        story.append(Paragraph(heading_text, heading_style))
        if not items:
            story.append(Paragraph("Not specified in the uploaded SRS.", body_style))
        else:
            for item in items:
                if isinstance(item, dict):
                    text = " — ".join(str(v) for v in item.values() if v)
                else:
                    text = str(item)
                story.append(Paragraph(f"• {text}", bullet_style))
        story.append(Spacer(1, 10))

    add_section("Business Goals", analysis.get("business_goals", []))
    add_section("Functional Requirements", analysis.get("functional_requirements", []))
    add_section("Non-Functional Requirements", analysis.get("non_functional_requirements", []))
    add_section("Constraints", analysis.get("constraints", []))
    add_section("Risks", analysis.get("risks", []))
    add_section("Missing Requirements", analysis.get("missing_requirements", []))

    story.append(Paragraph("Validation Report", heading_style))
    story.append(Paragraph(analysis.get("validation_report", "Not specified in the uploaded SRS."), body_style))
    story.append(Spacer(1, 10))

    cost = record.get("cost_estimation")
    if cost:
        story.append(Paragraph("Cost Estimation", heading_style))
        table_data = [["Item", "Low", "Medium", "High"]]
        for label, key in COST_LINE_ITEMS:
            item = cost.get(key, {}) or {}
            table_data.append([label, fmt_currency(item.get("low")), fmt_currency(item.get("medium")),
                                fmt_currency(item.get("high"))])
        cost_table = Table(table_data, hAlign="LEFT")
        cost_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7c6cf6")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f3fa")]),
        ]))
        story.append(cost_table)
        story.append(Spacer(1, 10))

    roadmap = record.get("roadmap")
    if roadmap and roadmap.get("phases"):
        story.append(PageBreak())
        story.append(Paragraph("Developer Roadmap", heading_style))
        for p in roadmap.get("phases", []):
            story.append(Paragraph(p.get("phase_name", "Phase"), styles["Heading3"]))
            story.append(Paragraph(f"Duration: {p.get('duration', '—')} | Owner: {p.get('owner', '—')}", body_style))
            for t in p.get("tasks", []) or []:
                story.append(Paragraph(f"• {t}", bullet_style))
            story.append(Spacer(1, 6))

    use_cases = record.get("use_cases")
    if use_cases and use_cases.get("use_cases"):
        story.append(PageBreak())
        story.append(Paragraph("Use Cases", heading_style))
        for uc in use_cases.get("use_cases", []) or []:
            if not isinstance(uc, dict):
                continue
            story.append(Paragraph(f"{uc.get('id', 'UC')}: {uc.get('title', '')}", styles["Heading3"]))
            story.append(Paragraph(
                f"Primary Actor: {uc.get('primary_actor', '—')} | Trigger: {uc.get('trigger', '—')}", body_style
            ))
            for step in uc.get("main_flow", []) or []:
                story.append(Paragraph(f"• {step}", bullet_style))
            story.append(Spacer(1, 6))

    risk = record.get("risk_feasibility")
    if risk:
        story.append(PageBreak())
        story.append(Paragraph("Risk & Feasibility Analysis", heading_style))
        story.append(Paragraph(
            f"Feasibility Score: {risk.get('feasibility_score', '—')}/100 | "
            f"Complexity Score: {risk.get('complexity_score', '—')}/100", body_style
        ))
        risk_table_data = [["Risk", "Probability", "Impact", "Severity"]]
        for key, label in RISK_CATEGORIES:
            item = risk.get(key, {}) or {}
            risk_table_data.append([label, item.get("probability", "—"), item.get("impact", "—"),
                                     item.get("severity", "—")])
        risk_table = Table(risk_table_data, hAlign="LEFT")
        risk_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7c6cf6")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f3fa")]),
        ]))
        story.append(risk_table)
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"Recommendation: {risk.get('recommendation', '—')}", body_style))

    team = record.get("team_roles")
    if team:
        story.append(PageBreak())
        story.append(Paragraph("Team Roles & Resource Planning", heading_style))
        planning = team.get("team_planning", {}) or {}
        story.append(Paragraph(
            f"Team Size: {planning.get('team_size', '—')} | "
            f"Team Cost: {fmt_currency(planning.get('team_cost'))}", body_style
        ))
        team_table_data = [["Role", "Hours/Week", "Cost", "Timeline"]]
        for key, label in TEAM_ROLES:
            role = team.get("roles", {}).get(key, {}) or {}
            team_table_data.append([label, str(role.get("hours_per_week", "—")),
                                     fmt_currency(role.get("estimated_cost")), role.get("timeline", "—")])
        team_table = Table(team_table_data, hAlign="LEFT")
        team_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#7c6cf6")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#cccccc")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f3fa")]),
        ]))
        story.append(team_table)
        story.append(Spacer(1, 10))

    pdf_doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def render_export_center_tab():
    render_module_header("⬇️", "Export Center",
                          "Download this analysis as PDF, Word, full JSON, or focused planning exports.")
    user_id = current_user_id()
    doc_id, record = get_selected_document(user_id, "export")
    if record is None:
        return

    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.markdown(f"**{record['filename']}**")
    st.caption(f"{record['page_count']} pages · {record['word_count']} words")
    st.markdown("</div>", unsafe_allow_html=True)

    pdf_key = f"_export_pdf_bytes_{doc_id}"
    docx_key = f"_export_docx_bytes_{doc_id}"

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📄 Build PDF Report", use_container_width=True, key="build_pdf"):
            with st.spinner("Building PDF..."):
                try:
                    st.session_state[pdf_key] = build_pdf_report(record)
                except Exception as e:
                    st.error(f"❌ Could not build PDF: {e}")
        if st.session_state.get(pdf_key):
            st.download_button(
                "⬇️ Download PDF", data=st.session_state[pdf_key],
                file_name=f"report_{doc_id}.pdf", mime="application/pdf",
                use_container_width=True, key="dl_pdf",
            )
    with col2:
        if st.button("📝 Build Word Document", use_container_width=True, key="build_docx"):
            with st.spinner("Building DOCX..."):
                try:
                    st.session_state[docx_key] = build_docx_report(record)
                except Exception as e:
                    st.error(f"❌ Could not build DOCX: {e}")
        if st.session_state.get(docx_key):
            st.download_button(
                "⬇️ Download Word", data=st.session_state[docx_key],
                file_name=f"report_{doc_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="dl_docx",
            )
    with col3:
        report_json = generate_report_json(record)
        st.download_button(
            "⬇️ Download Full JSON", data=report_json,
            file_name=f"report_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_json_full",
        )

    st.markdown("#### 🎯 Focused Exports")
    c1, c2, c3 = st.columns(3)
    with c1:
        dev_plan = record.get("roadmap") or {}
        st.download_button(
            "⬇️ Developer Plan (JSON)", data=json.dumps(dev_plan, indent=2, ensure_ascii=False),
            file_name=f"developer_plan_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_devplan", disabled=not dev_plan,
        )
    with c2:
        timeline_export = {
            "weekly_plan": (record.get("roadmap") or {}).get("weekly_plan", []),
            "monthly_milestones": (record.get("roadmap") or {}).get("monthly_milestones", []),
            "timeline_estimate": record.get("analysis", {}).get("timeline_estimate", {}),
        }
        st.download_button(
            "⬇️ Timeline (JSON)", data=json.dumps(timeline_export, indent=2, ensure_ascii=False),
            file_name=f"timeline_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_timeline",
        )
    with c3:
        arch_export = record.get("architecture") or {}
        st.download_button(
            "⬇️ Architecture (JSON)", data=json.dumps(arch_export, indent=2, ensure_ascii=False),
            file_name=f"architecture_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_arch", disabled=not arch_export,
        )

    if not record.get("roadmap"):
        st.caption("💡 Generate a Developer Roadmap first to enable the Developer Plan export.")
    if not record.get("architecture"):
        st.caption("💡 Generate a System Architecture first to enable the Architecture export.")

    st.markdown("#### 🧾 Use Cases · ⚠️ Risk & Team Exports")
    e1, e2, e3 = st.columns(3)
    with e1:
        use_cases_export = record.get("use_cases") or {}
        st.download_button(
            "⬇️ Use Cases (JSON)", data=json.dumps(use_cases_export, indent=2, ensure_ascii=False),
            file_name=f"use_cases_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_usecases", disabled=not use_cases_export,
        )
    with e2:
        risk_export = record.get("risk_feasibility") or {}
        st.download_button(
            "⬇️ Risk & Feasibility (JSON)", data=json.dumps(risk_export, indent=2, ensure_ascii=False),
            file_name=f"risk_feasibility_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_risk", disabled=not risk_export,
        )
    with e3:
        team_export = record.get("team_roles") or {}
        st.download_button(
            "⬇️ Team Roles (JSON)", data=json.dumps(team_export, indent=2, ensure_ascii=False),
            file_name=f"team_roles_{doc_id}.json", mime="application/json",
            use_container_width=True, key="dl_team", disabled=not team_export,
        )
    if not record.get("use_cases") or not record.get("risk_feasibility") or not record.get("team_roles"):
        st.caption("💡 Generate Use Cases, Risk & Feasibility, and Team Roles to enable their exports above.")


# ==============================================================================
# UI — DASHBOARD ROUTER
# ==============================================================================

def render_dashboard():
    inject_global_css()
    render_sidebar()
    render_dashboard_header()

    tab = st.session_state.active_tab
    try:
        if tab == "Upload SRS":
            render_upload_tab()
        elif tab == "Analysis History":
            render_history_tab()
        elif tab == "AI Chat":
            render_chat_tab()
        elif tab == "Developer Roadmap":
            render_roadmap_tab()
        elif tab == "Cost Estimation":
            render_cost_tab()
        elif tab == "Tech Stack Advisor":
            render_tech_tab()
        elif tab == "System Architect":
            render_architecture_tab()
        elif tab == "Output Predictor":
            render_output_predictor_tab()
        elif tab == "Coverage Score":
            render_coverage_tab()
        elif tab == "SRS Improver":
            render_srs_improver_tab()
        elif tab == "Use Case Generator":
            render_use_case_tab()
        elif tab == "Risk & Feasibility":
            render_risk_tab()
        elif tab == "Team Roles":
            render_team_tab()
        elif tab == "Project Generator":
            render_project_generator_tab()
        elif tab == "Management Dashboard":
            render_management_dashboard_tab()
        elif tab == "Export Center":
            render_export_center_tab()
        elif tab == "Reports":
            render_reports_tab()
        elif tab == "Profile":
            render_profile_tab()
        else:
            render_upload_tab()
    except Exception as e:
        st.error(
            "❌ Something went wrong while rendering this section. "
            "Your data has not been lost — try switching tabs or reloading."
        )
        with st.expander("Technical details"):
            st.code(traceback.format_exc())


# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon=APP_ICON,
        layout="wide",
        initial_sidebar_state="expanded",
    )
    ensure_directories()
    init_session_state()

    if not st.session_state.authenticated or not st.session_state.user:
        render_auth_page()
    else:
        # Defensive re-check: make sure the session's user still exists
        # (e.g. users.json was reset) before granting dashboard access.
        if not find_user_by_id(st.session_state.user["id"]):
            logout()
            st.rerun()
        render_dashboard()


if __name__ == "__main__":
    main()


# ==============================================================================
# GOOGLE COLAB DEPLOYMENT INSTRUCTIONS
# ==============================================================================
#
# Run the following in separate Colab cells (after saving this file as app.py
# in your Colab working directory):
#
# --- Cell 1: install dependencies ---
#   !pip install -q streamlit bcrypt PyMuPDF python-docx langchain langchain-community \
#       langchain-openai faiss-cpu openai pyngrok reportlab
#
# --- Cell 2: provide your OpenAI key via Streamlit secrets ---
#   import os
#   os.makedirs(".streamlit", exist_ok=True)
#   with open(".streamlit/secrets.toml", "w") as f:
#       f.write('OPENAI_API_KEY = "sk-your-key-here"\n')
#
# --- Cell 3: authenticate ngrok (sign up free at https://ngrok.com to get a token) ---
#   from pyngrok import ngrok
#   ngrok.set_auth_token("your-ngrok-auth-token-here")
#
# --- Cell 4: launch Streamlit and expose it via ngrok ---
#   import subprocess, time
#   from pyngrok import ngrok
#
#   ngrok.kill()
#   proc = subprocess.Popen([
#       "streamlit", "run", "app.py",
#       "--server.port", "8501",
#       "--server.headless", "true",
#   ])
#   time.sleep(6)
#   public_url = ngrok.connect(8501, "http")
#   print("Your app is live at:", public_url)
#
# --- Cell 5 (optional): stop the app later ---
#   ngrok.kill()
#   proc.terminate()
#
# ==============================================================================
