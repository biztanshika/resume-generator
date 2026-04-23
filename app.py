"""
Phase 2 -- Streamlit Resume Generator App
Imports create_resume() from Phase 1 (resume_generator.py) without modification.
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


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
MODEL = "gpt-4o"


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
    }


    /* ----- spinner ----- */
    .stSpinner > div {
        border-top-color: #0077B6 !important;
    }


    /* ----- success / error ----- */
    .stAlert {
        border-radius: 8px;
    }
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
     taken directly from the resume text for that role.
   - "bullets": Preserve ALL duty and accomplishment bullets from the original resume \
     for this role. Do not drop or merge bullets. Keep each bullet at its original \
     length and specificity (typically 10-20 words). Use active voice and strong verbs.

4. **projects** -- A list of objects, one per project:
   - "title": Project name ONLY -- strip technology tags, locations, and dates.
   - "bullets": Preserve ALL detail bullets from the original resume for this project. \
     Do not drop or merge bullets. Each bullet should capture technologies used, \
     outcomes achieved, and responsibilities held, at the original level of detail.

5. **education** -- A single object:
   - "degree": Degree name.
   - "college": Institution name and graduation year (if present).

6. **skills** -- A list of strings, each in the format "Category: item1, item2, item3". \
Group by meaningful categories based on resume content (e.g. Programming, \
Web Technologies, Databases, Tools, Methodologies, Cloud, BI Tools, etc.). \
Include all skills mentioned in the resume.

7. **certifications** -- A list of strings, one per certification. \
If no certifications are found, return [].

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
  "education": {
    "degree": "string",
    "college": "string"
  },
  "skills": ["Category: items", "..."],
  "certifications": ["string or empty list"]
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
# OpenAI call
# ---------------------------------------------------------------------------
REQUIRED_KEYS = {"name", "summary", "experience", "projects", "education", "skills"}




def call_openai(resume_text: str) -> dict:
    """Send resume text to gpt-4.5-preview and return the parsed JSON dict."""
    if not OPENAI_API_KEY:
        raise RuntimeError(
            "OPENAI_API_KEY is not set. Add it to the .env file in the project root."
        )


    client = OpenAI(api_key=OPENAI_API_KEY)


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
# Streamlit UI
# ---------------------------------------------------------------------------


def main():
    st.set_page_config(
        page_title="Resume Generator",
        page_icon="briefcase",
        layout="centered",
    )


    # Inject WorkPace CSS
    st.markdown(WORKPACE_CSS, unsafe_allow_html=True)


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


    # ---- Upload section ----
    st.markdown('<div class="wp-card">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Drop your resume here",
        type=["docx", "pdf", "txt"],
        help="Supported formats: .docx, .pdf, .txt",
    )


    generate_clicked = st.button("Generate Resume", disabled=uploaded_file is None)
    st.markdown("</div>", unsafe_allow_html=True)


    # ---- Processing ----
    if generate_clicked and uploaded_file is not None:
        st.markdown('<div class="wp-card">', unsafe_allow_html=True)


        # Step 1 -- extract text
        with st.spinner("Extracting text from your resume..."):
            try:
                resume_text = extract_text(uploaded_file)
            except Exception as exc:
                st.error(f"Could not read the uploaded file: {exc}")
                st.markdown("</div>", unsafe_allow_html=True)
                return


        if not resume_text.strip():
            st.error("The uploaded file appears to be empty.")
            st.markdown("</div>", unsafe_allow_html=True)
            return


        # Step 2 -- AI parsing
        with st.spinner("AI is analysing your resume..."):
            try:
                resume_data = call_openai(resume_text)
            except Exception as exc:
                st.error(f"AI processing failed: {exc}")
                st.markdown("</div>", unsafe_allow_html=True)
                return


        # Step 3 -- generate DOCX
        with st.spinner("Building your formatted resume..."):
            try:
                docx_bytes = generate_resume_bytes(resume_data)
            except Exception as exc:
                st.error(f"DOCX generation failed: {exc}")
                st.markdown("</div>", unsafe_allow_html=True)
                return


        st.success("Your resume is ready!")


        # Derive download filename from the candidate name
        safe_name = resume_data.get("name", "resume").replace(" ", "_")
        filename = f"{safe_name}_Resume.docx"


        st.download_button(
            label="Download Resume",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


        st.markdown("</div>", unsafe_allow_html=True)




if __name__ == "__main__":
    main()



