"""
Phase 5 -- Streamlit Resume Generator App + Role-Fit Analyzer (Side-Panel Layout)
Imports create_resume() from Phase 1 (resume_generator.py) without modification.
DOCX generation is identical to Phase 2 (app.py). Phase 5 adds a side-panel canvas
for the simplified 3-section role-fit analysis (Strengths / Weaknesses / Suggestions)
that appears to the right of the upload form.
"""

import io
import json
import os
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI

# Phase 1 import -- resume_generator.py is NOT modified
from resume_generator import create_resume

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
load_dotenv(Path(__file__).resolve().parent / ".env")

MODEL = "gpt-4o-mini"


def _resolve_api_key() -> str:
    """Return the active OpenAI key: UI-entered key > .env fallback."""
    return (
        st.session_state.get("user_api_key", "").strip()
        or os.getenv("OPENAI_API_KEY", "")
    )

# ---------------------------------------------------------------------------
# WorkPace-inspired theme CSS
# ---------------------------------------------------------------------------
WORKPACE_CSS = """
<style>
    /* ----- global background ----- */
    .stApp {
        background-color: #EAF4FB;
    }

    /* ----- header area ----- */
    header[data-testid="stHeader"] {
        background-color: #EAF4FB;
    }

    /* ----- card container ----- */
    .wp-card {
        background: #ffffff;
        border-radius: 12px;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
        padding: 2rem;
        margin-bottom: 1.5rem;
    }

    /* ----- brand header ----- */
    .wp-brand {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0A3D62;
        margin-bottom: 0.25rem;
    }
    .wp-brand span {
        color: #0077B6;
    }

    .wp-subtitle {
        color: #555;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }

    /* ----- buttons ----- */
    .stButton > button {
        background-color: #0A3D62 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        transition: background 0.2s;
    }
    .stButton > button:hover {
        background-color: #0077B6 !important;
    }

    /* ----- download button ----- */
    .stDownloadButton > button {
        background-color: #0077B6 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 2rem !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }
    .stDownloadButton > button:hover {
        background-color: #0A3D62 !important;
    }

    /* ----- file uploader zone ----- */
    [data-testid="stFileUploader"] {
        border: 2px dashed #0077B6;
        border-radius: 12px;
        padding: 1rem;
        background-color: #ffffff;
    }
    [data-testid="stFileUploaderDropzone"] {
        background-color: #F0F7FC !important;
        border: none !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzone"] span,
    [data-testid="stFileUploaderDropzone"] small,
    [data-testid="stFileUploaderDropzone"] p,
    [data-testid="stFileUploaderDropzone"] div {
        color: #0A3D62 !important;
    }
    [data-testid="stFileUploaderDropzone"] button {
        color: #0A3D62 !important;
        border-color: #0077B6 !important;
        background-color: #ffffff !important;
    }
    [data-testid="stFileUploaderDropzone"] svg {
        fill: #0077B6 !important;
        stroke: #0077B6 !important;
    }

    /* ----- spinner ----- */
    .stSpinner > div {
        border-top-color: #0077B6 !important;
    }

    /* ----- success / error ----- */
    .stAlert {
        border-radius: 8px;
    }

    /* ----- widget labels (file uploader, text area, etc.) ----- */
    [data-testid="stWidgetLabel"] p {
        color: #0A3D62 !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
    }

    /* ----- bordered containers (upload card, processing card) ----- */
    [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08) !important;
        border: none !important;
        padding: 0.5rem !important;
    }

    /* ----- text area white background ----- */
    [data-testid="stTextArea"] textarea {
        background-color: #ffffff !important;
    }

    /* ----- Phase 5: analysis canvas (right panel) ----- */
    .analysis-canvas {
        background: #FFFFFF;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        padding: 1.5rem 1.75rem;
    }

    /* Score row */
    .canvas-score-row {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        padding-bottom: 0.9rem;
        margin-bottom: 1rem;
        border-bottom: 1px solid #E2E8F0;
    }
    .canvas-score-label {
        font-size: 1.25rem;
        font-weight: 700;
        color: #1E293B;
    }
    .score-badge {
        display: inline-block;
        padding: 0.2rem 0.75rem;
        border-radius: 999px;
        font-size: 1rem;
        font-weight: 700;
    }
    .score-badge-high { background: #ECFDF5; color: #065F46; }
    .score-badge-mid  { background: #FEF9EC; color: #D97706; }
    .score-badge-low  { background: #FEF2F2; color: #DC2626; }

    /* Summary strip */
    .fit-summary {
        background: #F1F5F9;
        border-left: 3px solid #94A3B8;
        border-radius: 4px;
        padding: 0.6rem 0.9rem;
        color: #475569;
        font-size: 0.9rem;
        font-style: italic;
        margin-bottom: 1.1rem;
    }

    /* Section blocks */
    .fit-block {
        border-radius: 8px;
        padding: 0.9rem 1.15rem;
        margin-bottom: 0.9rem;
        border-left: 4px solid;
    }
    .fit-block-strengths   { background: #ECFDF5; border-color: #10B981; }
    .fit-block-weaknesses  { background: #FEF2F2; border-color: #EF4444; }
    .fit-block-suggestions { background: #EFF6FF; border-color: #3B82F6; }

    .fit-block-title { font-size: 0.95rem; font-weight: 700; margin-bottom: 0.5rem; }
    .fit-title-strength   { color: #065F46; }
    .fit-title-weakness   { color: #7F1D1D; }
    .fit-title-suggestion { color: #1E3A8A; }

    .fit-block-item { font-size: 0.88rem; line-height: 1.65; margin-bottom: 0.25rem; }
    .fit-item-strength   { color: #065F46; }
    .fit-item-weakness   { color: #7F1D1D; }
    .fit-item-suggestion { color: #1E3A8A; }
</style>
"""

# ---------------------------------------------------------------------------
# AI system prompt  (derived from the Custom GPT prompt, outputs JSON)
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """\
You are a Resume Content Extractor.

You receive raw resume text (from a .docx, .pdf, or .txt upload). Your job is \
to parse and organise it into a structured JSON object that matches the exact \
schema below. Return ONLY valid JSON -- no markdown, no explanation, no extra keys.

### Core Rules (apply to every section)

- Use ONLY information explicitly present in the uploaded resume. Never invent, \
infer, or add any detail not stated by the candidate.
- Preserve the original depth and richness of the resume. Do NOT compress, \
summarise, or drop content. If the resume has 5 bullets for a role, output 5 bullets. \
If a bullet is 15 words long, keep it at roughly that length.
- Write in polished, professional resume style: active voice, strong verbs, \
specific numbers and tools preserved exactly as written.

### Content Rules

1. **name** -- The candidate's full name.

2. **summary** -- A list of 6-8 bullet strings derived from the resume content. \
Each bullet must start with a bold label followed by a colon:
   - Always include a "Domain:" bullet listing the candidate's domain expertise areas \
     (e.g. "Domain: Ed-Tech, FinTech, Marketing Analytics").
   - Choose remaining labels dynamically based on what is prominent in the resume. \
     Use labels such as: Summary, SQL, Excel, Tools, Data Competencies, BI Tools, \
     Collaboration & Leadership, Problem-Solving, Infrastructure & Automation, \
     Data Visualization, Data Engineering, Cloud, etc.
   - Each bullet must be a single, specific, impact-focused sentence (10-20 words). \
     Reflect actual skills, tools, and achievements from the resume -- do not pad \
     or genericise.

3. **experience** -- A list of objects, one per role:
   - "role": Job title ONLY -- strip company name, location, and dates.
   - "tech": Comma-separated list of technologies/tools used in that role, \
     taken directly from the resume text for that role. Do not add any technology not specifically mentioned in that role \
   - "bullets": Preserve ALL duty and accomplishment bullets from the original resume \
     for this role. Do not drop or merge bullets. Reword the bullet points if required using active voice and strong verbs \
     Keep each bullet at its original length and specificity (typically 10-20 words).

4. **projects** -- A list of objects, one per project:
   - "title": Project name ONLY -- strip technology tags, locations, and dates.
   - "bullets": Preserve ALL detail bullets from the original resume for this project. \
     Do not drop or merge bullets. Each bullet should capture technologies used, \
     outcomes achieved, and responsibilities held, at the original level of detail.

5. **education** -- A list of objects, one per degree. If the resume lists multiple \
degrees (e.g. a Bachelor's AND a Master's), you MUST include a separate object for \
EACH degree. Never combine or drop degrees. Each object has:
   - "degree": Degree name.
   - "college": Institution name and graduation year (if present).

6. **skills** -- A list of strings, each in the format "Category: item1, item2, item3". \
Group by meaningful categories based on resume content (e.g. Programming, \
Web Technologies, Databases, Tools, Methodologies, Cloud, BI Tools, etc.). \
Include all skills mentioned in the resume.

### JSON Schema

{
  "name": "string",
  "summary": ["Label: text", "..."],
  "experience": [
    {
      "role": "string",
      "tech": "string",
      "bullets": ["string"]
    }
  ],
  "projects": [
    {
      "title": "string",
      "bullets": ["string"]
    }
  ],
  "education": [
    {
      "degree": "string",
      "college": "string"
    },
    {
      "degree": "string (include ALL degrees)",
      "college": "string"
    }
  ],
  "skills": ["Category: items", "..."],
  "certifications": ["string or empty list"]
}

Return ONLY the JSON object. No markdown fences, no commentary.\
"""

# ---------------------------------------------------------------------------
# Phase 3: Role-fit rating system prompt
# ---------------------------------------------------------------------------
RATING_SYSTEM_PROMPT = """\
You are a Resume-to-Job Fit Analyst.

You will receive two pieces of text:
1. **RESUME** -- the full text of a candidate's resume.
2. **JOB CONTEXT** -- a free-form description provided by the user. It may \
include a job title, project description, required technologies, tools, \
responsibilities, team structure, domain, or any combination of these.

Your task is to evaluate how well the resume matches the job context and return \
a structured JSON assessment. Return ONLY valid JSON -- no markdown, no \
explanation, no extra keys.

### Evaluation Criteria

Score the resume on a 1-10 scale considering ALL of the following:
- Skills match: Do the candidate's skills, tools, and technologies align with \
  what the job context requires?
- Experience relevance: Does the work history show relevant responsibilities, \
  domains, or project types?
- Project alignment: Do projects use similar technologies or solve similar \
  problems?
- Keyword coverage: Are the specific tools, frameworks, methodologies, and \
  domain terms from the job context present in the resume?
- Seniority fit: Does the experience level match what the role requires?

### Output Rules

1. **strengths** -- List 3-6 specific things that make the resume a good fit. \
Reference actual resume content (e.g. "3 years of Python experience aligns \
with the required Python proficiency"). Each item is a plain string.

2. **weaknesses** -- List every significant gap or shortcoming. This includes:
   - Missing skills, tools, or technologies mentioned in the job context but \
     absent from the resume.
   - Experience gaps -- areas where the candidate's experience is thin, \
     irrelevant, or too junior relative to the role.
   - Domain mismatches or missing certifications.
   Each item is a plain string that clearly states what is missing or weak.

3. **suggestions** -- List 3-6 concrete, actionable steps the candidate should \
take to better align their resume with this specific job context. Focus on \
content changes: which sections to strengthen, what keywords to add, which \
experience to highlight or reframe. Do NOT suggest formatting or layout changes. \
Each item is a plain string.

### JSON Schema

{
  "overall_score": <integer 1-10>,
  "score_out_of": 10,
  "summary": "<1-2 sentence overall verdict>",
  "strengths": ["<string>", "..."],
  "weaknesses": ["<string>", "..."],
  "suggestions": ["<string>", "..."]
}

Return ONLY the JSON object. No markdown fences, no commentary.\
"""

# ---------------------------------------------------------------------------
# Text extraction helpers  (inline, no separate module)
# ---------------------------------------------------------------------------

def extract_text_docx(file_bytes: bytes) -> str:
    """Extract text from a .docx file using python-docx."""
    from docx import Document as DocxDocument

    with io.BytesIO(file_bytes) as buf:
        doc = DocxDocument(buf)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def extract_text_pdf(file_bytes: bytes) -> str:
    """Extract text from a PDF using pdfplumber."""
    import pdfplumber

    text_parts = []
    with io.BytesIO(file_bytes) as buf:
        with pdfplumber.open(buf) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(page_text)
    return "\n".join(text_parts)

def extract_text_txt(file_bytes: bytes) -> str:
    """Decode a plain-text file."""
    return file_bytes.decode("utf-8", errors="replace")

EXTRACTORS = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": extract_text_docx,
    "application/pdf": extract_text_pdf,
    "text/plain": extract_text_txt,
}

def extract_text(uploaded_file) -> str:
    """Route to the correct extractor based on MIME type."""
    mime = uploaded_file.type
    file_bytes = uploaded_file.getvalue()

    # Fallback by extension if MIME is generic
    if mime not in EXTRACTORS:
        name_lower = uploaded_file.name.lower()
        if name_lower.endswith(".docx"):
            return extract_text_docx(file_bytes)
        if name_lower.endswith(".pdf"):
            return extract_text_pdf(file_bytes)
        if name_lower.endswith(".txt"):
            return extract_text_txt(file_bytes)
        raise ValueError(f"Unsupported file type: {mime} ({uploaded_file.name})")

    return EXTRACTORS[mime](file_bytes)

# ---------------------------------------------------------------------------
# OpenAI calls
# ---------------------------------------------------------------------------
REQUIRED_KEYS = {"name", "summary", "experience", "projects", "education", "skills"}

def call_openai(resume_text: str, api_key: str = "") -> dict:
    """Send resume text to the model and return the parsed JSON dict."""
    key = api_key or _resolve_api_key()
    if not key:
        raise RuntimeError("Please enter your OpenAI API key above.")

    client = OpenAI(api_key=key)

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": resume_text},
        ],
        response_format={"type": "json_object"},
        temperature=0.3,
    )

    raw = response.choices[0].message.content
    data = json.loads(raw)

    # Validate required keys
    missing = REQUIRED_KEYS - set(data.keys())
    if missing:
        raise ValueError(f"AI response missing required keys: {missing}")

    # Ensure certifications key exists (may be absent)
    data.setdefault("certifications", [])

    return data

RATING_REQUIRED_KEYS = {
    "overall_score", "score_out_of", "summary",
    "strengths", "weaknesses", "suggestions",
}

def call_openai_rating(resume_text: str, job_context: str, api_key: str = "") -> dict:
    """Rate the resume against a free-form job context description."""
    key = api_key or _resolve_api_key()
    if not key:
        raise RuntimeError("Please enter your OpenAI API key above.")

    user_message = (
        "=== RESUME ===\n"
        f"{resume_text}\n\n"
        "=== JOB CONTEXT ===\n"
        f"{job_context}"
    )

    client = OpenAI(api_key=key)

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": RATING_SYSTEM_PROMPT},
            {"role": "user", "content": user_message},
        ],
        response_format={"type": "json_object"},
        temperature=0.4,
    )

    raw = response.choices[0].message.content
    data = json.loads(raw)

    missing = RATING_REQUIRED_KEYS - set(data.keys())
    if missing:
        raise ValueError(f"Rating response missing required keys: {missing}")

    return data

# ---------------------------------------------------------------------------
# DOCX generation wrapper  (writes to BytesIO buffer)
# ---------------------------------------------------------------------------

def generate_resume_bytes(data: dict) -> bytes:
    """Call Phase 1's create_resume() and return the .docx bytes."""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        create_resume(data, output_path=tmp_path)
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

# ---------------------------------------------------------------------------
# Phase 3: Rating display helpers
# ---------------------------------------------------------------------------

def _score_badge_class(score: int) -> str:
    """Return badge CSS class based on score."""
    if score >= 8:
        return "score-badge score-badge-high"
    if score >= 5:
        return "score-badge score-badge-mid"
    return "score-badge score-badge-low"

def display_rating_canvas(rating: dict):
    """Render the 3-section role-fit analysis as left-border accent blocks."""
    score   = rating.get("overall_score", 0)
    out_of  = rating.get("score_out_of", 10)
    badge   = _score_badge_class(score)
    summary = rating.get("summary", "")

    parts = ['<div class="analysis-canvas">']

    # Score row with badge pill
    parts.append(
        f'<div class="canvas-score-row">'
        f'<span class="canvas-score-label">Score</span>'
        f'<span class="{badge}">{score}/{out_of}</span>'
        f'</div>'
    )

    # Summary strip
    if summary:
        parts.append(f'<div class="fit-summary">{summary}</div>')

    # Strengths block
    strengths = rating.get("strengths", [])
    if strengths:
        items = "".join(
            f'<div class="fit-block-item fit-item-strength">&bull; {s}</div>'
            for s in strengths
        )
        parts.append(
            f'<div class="fit-block fit-block-strengths">'
            f'<div class="fit-block-title fit-title-strength">Strength</div>'
            f'{items}</div>'
        )

    # Weaknesses block
    weaknesses = rating.get("weaknesses", [])
    if weaknesses:
        items = "".join(
            f'<div class="fit-block-item fit-item-weakness">&bull; {w}</div>'
            for w in weaknesses
        )
        parts.append(
            f'<div class="fit-block fit-block-weaknesses">'
            f'<div class="fit-block-title fit-title-weakness">Weaknesses</div>'
            f'{items}</div>'
        )

    # Suggestions block
    suggestions = rating.get("suggestions", [])
    if suggestions:
        items = "".join(
            f'<div class="fit-block-item fit-item-suggestion">&bull; {s}</div>'
            for s in suggestions
        )
        parts.append(
            f'<div class="fit-block fit-block-suggestions">'
            f'<div class="fit-block-title fit-title-suggestion">Suggestions</div>'
            f'{items}</div>'
        )

    parts.append('</div>')
    st.markdown("".join(parts), unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="Resume Generator",
        page_icon="briefcase",
        layout="wide",
    )

    # Inject WorkPace CSS
    st.markdown(WORKPACE_CSS, unsafe_allow_html=True)

    # Decide layout: centered single column by default,
    # left + right panel when analysis data exists.
    has_rating = "rating_data" in st.session_state

    if has_rating:
        left_col, right_col = st.columns([4, 5])
    else:
        # Simulate a centered column inside wide layout
        _pad_l, left_col, _pad_r = st.columns([1, 3, 1])
        right_col = None

    # Track whether we need to rerun after column blocks close
    needs_rerun = False

    # ---- Left column: brand + upload + download ----
    with left_col:
        # ---- Brand header ----
        st.markdown(
            '<div class="wp-card">'
            '  <div class="wp-brand">Resume<span>Gen</span></div>'
            '  <div class="wp-subtitle">'
            "    Upload your resume and get a professionally formatted Word document "
            "    powered by AI."
            "  </div>"
            "</div>",
            unsafe_allow_html=True,
        )

        # ---- API key input ----
        api_key_input = st.text_input(
            "OpenAI / GPT API Key",
            type="password",
            placeholder="sk-...",
            help=(
                "Enter your OpenAI API key. It is used only for this session "
                "to process your resume via the GPT API and is not stored."
            ),
        ).strip()
        if api_key_input:
            st.session_state["user_api_key"] = api_key_input

        # ---- Upload section ----
        with st.container(border=True):
            uploaded_file = st.file_uploader(
                "Drop your resume here",
                type=["docx", "pdf", "txt"],
                help="Supported formats: .docx, .pdf, .txt",
            )

            job_context = st.text_area(
                "Job Description / Role Context (optional)",
                placeholder=(
                    "Describe the role, project, required technologies, responsibilities, "
                    "and any specific skills needed. Leave blank to skip the role-fit analysis."
                ),
                height=120,
            )

            # Phase 4: three independent action buttons
            btn1, btn2, btn3 = st.columns(3)
            with btn1:
                gen_clicked = st.button(
                    "Build Resume",
                    disabled=uploaded_file is None,
                    use_container_width=True,
                )
            with btn2:
                fit_clicked = st.button(
                    "Analyse Fit",
                    disabled=uploaded_file is None,
                    use_container_width=True,
                )
            with btn3:
                both_clicked = st.button(
                    "Build & Analyse",
                    disabled=uploaded_file is None,
                    use_container_width=True,
                )

        # ---- Processing (runs on button click, stores results in session) ----
        if (gen_clicked or fit_clicked or both_clicked) and uploaded_file is not None:
            # Validate API key before any OpenAI work
            active_key = _resolve_api_key()
            if not active_key:
                st.error("Please enter your OpenAI API key above before proceeding.")
            else:
                # Step 1 -- extract text (shared by all paths)
                with st.spinner("Extracting text from your resume..."):
                    try:
                        resume_text = extract_text(uploaded_file)
                    except Exception as exc:
                        st.error(f"Could not read the uploaded file: {exc}")
                        return

                if not resume_text.strip():
                    st.error("The uploaded file appears to be empty.")
                    return

                # Step 2 -- Resume generation (only when requested)
                if gen_clicked or both_clicked:
                    with st.spinner("AI is analysing your resume..."):
                        try:
                            resume_data = call_openai(resume_text, api_key=active_key)
                        except Exception as exc:
                            st.error(f"AI processing failed: {exc}")
                            return

                    with st.spinner("Building your formatted resume..."):
                        try:
                            docx_bytes = generate_resume_bytes(resume_data)
                        except Exception as exc:
                            st.error(f"DOCX generation failed: {exc}")
                            return

                    safe_name = resume_data.get("name", "resume").replace(" ", "_")
                    st.session_state["docx_bytes"] = docx_bytes
                    st.session_state["docx_filename"] = f"{safe_name}_Resume.docx"

                # Step 3 -- Role-fit analysis (only when requested)
                if fit_clicked or both_clicked:
                    if not job_context.strip():
                        st.warning(
                            "Please enter a job description or role context to run "
                            "the fit analysis."
                        )
                    else:
                        with st.spinner("Analysing role fit..."):
                            try:
                                rating_data = call_openai_rating(
                                    resume_text, job_context.strip(),
                                    api_key=active_key,
                                )
                                st.session_state["rating_data"] = rating_data
                                needs_rerun = True
                            except Exception as exc:
                                st.error(f"Role-fit analysis failed: {exc}")
                                st.session_state.pop("rating_data", None)

        # ---- Download button (persists via session state) ----
        if "docx_bytes" in st.session_state:
            with st.container(border=True):
                st.success("Your resume is ready!")
                st.download_button(
                    label="Download Resume",
                    data=st.session_state["docx_bytes"],
                    file_name=st.session_state["docx_filename"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

    # ---- Right column: analysis canvas ----
    if has_rating and right_col is not None:
        with right_col:
            display_rating_canvas(st.session_state["rating_data"])

    # Rerun AFTER all column blocks are closed to avoid ASGI crash
    # (RerunException inside a with-column context breaks some Streamlit versions)
    if needs_rerun:
        st.rerun()

if __name__ == "__main__":
    main()

