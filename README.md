# Resume Generator


An AI-powered resume formatting and analysis tool. Upload any resume and the app
extracts its content, uses GPT to rewrite and structure it, then outputs a
professionally formatted Microsoft Word document (.docx) matching the company template.
It also supports **bulk resume screening** — score and rank multiple candidates
against a job description in one go, with CSV/Excel export.


---


## Local Setup


### Prerequisites
- Python 3.10 or later ([python.org](https://www.python.org/downloads/))
- Git ([git-scm.com](https://git-scm.com/))
- An OpenAI API key (`sk-...`) — entered directly in the app at runtime


### 1. Get the code


Clone the repository (or download and unzip it):


```bash
git clone https://github.com/your-username/your-repo-name.git
cd "Res Generator"
```


### 2. Create a virtual environment


```bash
python -m venv .venv
```


Activate it:


| Platform | Shell | Command |
|---|---|---|
| Windows | PowerShell | `.\.venv\Scripts\Activate.ps1` |
| Windows | CMD | `.\.venv\Scripts\activate.bat` |
| macOS / Linux | bash/zsh | `source .venv/bin/activate` |


> If PowerShell blocks the script, run this first:
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned`


### 3. Install dependencies


```bash
pip install -r requirements.txt
```


This installs: `streamlit`, `openai`, `pdfplumber`, `python-docx`, `lxml`,
`python-dotenv`, `openpyxl`.


### 4. Run the app


```bash
# With venv activated:
streamlit run app_v6.py


# Or using the venv Python directly (Windows):
.\.venv\Scripts\python.exe -m streamlit run app_v6.py
```


The app opens at **http://localhost:8501**.


### 5. Enter your API key


Paste your OpenAI `sk-...` key into the **"OpenAI / GPT API Key"** field at
the top of the page. The key is used only for the current browser session and
is never saved to disk.


**Optional local fallback:** Create a `.env` file in the project root:
```
OPENAI_API_KEY=sk-your-key-here
```
The UI-entered key always takes priority over the `.env` value.  
**Never commit `.env` to version control** — it is already listed in `.gitignore`.


### 6. Verify the install (optional)


```bash
.\.venv\Scripts\python.exe -c "from app_v6 import _extract_job_context, _render_bulk_results; print('All imports OK')"
```


---


## Streamlit Cloud Deployment


1. Push the repo to GitHub (ensure `requirements.txt` is committed).
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **Create app** → **Yup, I have an app**.
4. Set: Repository → your repo, Branch → `main`, Main file path → `app_v6.py`.
5. Click **Advanced settings** and set the Python version to match your local version (`python --version`).
6. Leave the **Secrets** field empty — the app uses BYOK (users supply their own API key at runtime).
7. Click **Deploy**. First build takes 2–5 minutes.


---


## Project Structure


```
Res Generator/
|-- app_v6.py                 Main Streamlit app (Phase 5 -- current version)
|-- resume_generator.py       Phase 1 -- DOCX generation engine (DO NOT MODIFY)
|-- requirements.txt          Python dependencies (used by pip and Streamlit Cloud)
|-- requirements_phase2.txt   Legacy deps file (superseded by requirements.txt)
|-- .env.example              Template showing required environment variables
|-- .gitignore                Excludes .env, .venv, __pycache__, generated .docx files
|-- vba_code_example .txt     Reference VBA macro (source of truth for formatting)
|-- assets/                   Icon XML templates and EMF image used in DOCX output
|   |-- Experience_0.xml      Briefcase icon anchor for Experience section
|   |-- Projects_0.xml        Briefcase icon anchor for Projects section
|   |-- Education_0.xml       Briefcase icon anchor for Education section
|   |-- Skills_0.xml          Lightbulb icon anchor for Skills section
|   |-- image1.emf            Icon image embedded in generated DOCX
|-- .venv/                    Python virtual environment (not committed)
```


---


## Architecture Overview


The project is split into two isolated phases:


```
[User uploads resume(s) + optional job description]
        |
        v
  app_v6.py (Phase 5 -- current)
        |-- _resolve_api_key()       UI key > .env fallback (BYOK)
        |-- extract_text()           strip text from .docx / .pdf / .txt
        |-- _extract_job_context()   extract JD text from typed input or uploaded file
        |
        |-- Single-resume path:
        |       |-- call_openai()            rewrite resume --> JSON
        |       |-- call_openai_rating()     score resume against JD
        |       |-- generate_resume_bytes()  --> st.download_button
        |
        |-- Bulk-analysis path (multiple files):
        |       |-- _analyse_one()           worker (runs in ThreadPoolExecutor)
        |       |-- _render_bulk_results()   ranked expandable cards
        |       |-- _build_csv_bytes()       CSV export
        |       |-- _build_excel_bytes()     styled Excel export
        |
        v
  resume_generator.py (Phase 1 -- frozen)
        |-- create_resume(data)   build the DOCX from the JSON dict
                v
        [returns .docx bytes] --> st.download_button
```


**Phase 1 (`resume_generator.py`) is a stable, frozen module.  
Phase 5 (`app_v6.py`) is the user-facing layer that drives it.**


---


## Phase 1 -- DOCX Generation Engine


**File:** `resume_generator.py`  
**Rule: This file must never be modified once working. Phase 2 imports from it.**


### What it does
Replicates the company VBA macro (`vba_code_example .txt`) in pure Python using
`python-docx`. All naming, structure, and formatting matches the VBA original.


### Key functions (mirrors VBA Sub names)


| Python function | VBA equivalent |
|---|---|
| `create_resume(data, output_path)` | `RecreateFinalResume_CleanAndSimple` |
| `build_experience_section(doc, experiences)` | `BuildExperienceSection` |
| `build_projects_section(doc, projects)` | `BuildProjectsSection` |
| `insert_section_header(doc, title)` | `InsertSectionHeader` |
| `insert_role_title(doc, title)` | `InsertRoleTitle` |
| `insert_technologies_line(doc, tech_text)` | `InsertTechnologiesLine` |
| `insert_experience_bullet_list(doc, items)` | `InsertExperienceBulletList` |
| `insert_project_with_number(doc, idx, title, bullets)` | `InsertProjectWithNumber` |
| `insert_project_bullet_list(doc, items)` | `InsertProjectBulletList` |
| `insert_summary_inline(doc, items)` | `InsertSummaryInline` |
| `insert_skills_list(doc, items)` | `InsertSkillsList` |
| `insert_education_entry(doc, degree, college)` | `InsertEducationEntry` |
| `add_header_line(doc)` | VBA Shapes.AddLine |


### Data dict schema (input to `create_resume`)


```python
{
    "name": "string",
    "summary": [
        "Summary: High-level overview...",
        "Domain: FinTech, Ed-Tech",
        "SQL: ...",
        # 6-8 items total, label before colon is rendered bold
    ],
    "experience": [
        {
            "role": "Data Engineer",            # job title only, no company/dates
            "tech": "Python, SQL, Tableau",     # comma-separated string
            "bullets": ["Accomplished X...", ...]
        }
    ],
    "projects": [
        {
            "title": "Inventory Management System",   # name only, no tech tags/dates
            "bullets": ["Built X using Y...", ...]
        }
    ],
    "education": {
        "degree": "Bachelor of Science in Computer Science",
        "college": "University of California, Los Angeles (UCLA), 2017"
    },
    "skills": [
        "Programming: Python, JavaScript, Java",   # label before colon is bold
        "Databases: MySQL, MongoDB"
    ],
    "certifications": []   # empty list if none; same format as skills if present
}
```


### Formatting constants
- Font: Century Gothic throughout
- Name: 28pt Bold
- Section headers: 16pt Bold, colour `#0066CC` (RGB 0, 102, 204)
- Role titles: 12pt Bold
- Technologies line: 11pt, label bold
- Bullet text: 11pt
- Education degree: 14pt, colour `#0066CC`
- Blue header line: 2pt stroke, drawn at 60pt from page top via DrawingML anchor
- Margins: 1 inch all sides


### Icon injection
Each section header paragraph has a floating icon injected after the text run.
Icon coordinates are stored as DrawingML anchor XML in `assets/*.xml`.
The EMF image (`assets/image1.emf`) is embedded as a document relationship.


---


## Phase 5 -- Streamlit Web Application


**File:** `app_v6.py`


### What it does
1. Presents a WorkPace-themed web UI
2. Accepts a single or multiple uploaded resumes (.docx, .pdf, .txt)
3. Accepts a job description typed in or uploaded as a file (.pdf, .docx, .txt)
4. **Single-resume mode:** rewrites the resume via GPT and generates a formatted DOCX download
5. **Bulk mode:** scores and ranks all uploaded resumes against the JD in parallel (up to 10 workers), with expandable result cards and CSV/Excel export
6. API key entered by the user in the UI at runtime (BYOK) -- no credentials stored anywhere


### Key functions


| Function | Purpose |
|---|---|
| `_resolve_api_key()` | Returns UI-entered key, falls back to `.env` |
| `extract_text(uploaded_file)` | Routes to correct extractor by MIME type |
| `extract_text_docx(bytes)` | Extracts paragraphs from .docx via python-docx |
| `extract_text_pdf(bytes)` | Extracts text pages from PDF via pdfplumber |
| `extract_text_txt(bytes)` | Decodes plain text |
| `_extract_job_context(uploaded_jd)` | Extracts JD text from an uploaded file |
| `call_openai(resume_text, api_key)` | Calls gpt-4o-mini, validates + returns JSON dict |
| `call_openai_rating(resume_text, job_context, api_key)` | Scores resume against JD, returns score + verdict |
| `_analyse_one(filename, file_bytes, mime, job_context, api_key)` | Thread worker for one resume |
| `_build_csv_bytes(results)` | UTF-8 CSV with Rank / File / Score / Verdict / Top items |
| `_build_excel_bytes(results)` | Styled openpyxl workbook, returns `None` if not installed |
| `_render_bulk_results(results)` | Ranked expandable cards with download buttons |
| `generate_resume_bytes(data)` | Wraps `create_resume()` to return bytes via temp file |
| `main()` | Streamlit page layout and UI logic |


### AI system prompt rules (derived from original Custom GPT prompt)
The system prompt instructs the model to extract and structure:
- **name**: candidate's full name
- **summary**: 6-8 bullets, always includes a `Domain:` bullet; all labels bold
- **experience**: job title only (company/location/dates stripped); tech as CSV string; duty bullets
- **projects**: project name only (tech tags/dates stripped); detail bullets
- **education**: degree + institution + year
- **skills**: `"Category: items"` format per bullet
- **certifications**: optional; empty list `[]` if absent


`response_format={"type": "json_object"}` is set on the API call to enforce
valid JSON output and eliminate any parsing risk.


### Theme
WorkPace-inspired colour palette injected via a single `st.markdown` CSS block:


| Element | Value |
|---|---|
| Page background | `#EAF4FB` |
| Primary accent | `#0077B6` |
| CTA / generate button | `#0A3D62` (dark navy) |
| Cards | White, `border-radius: 12px`, `box-shadow` |
| Upload zone | Dashed `#0077B6` border |


---


## Setup Instructions


See the **[Local Setup](#local-setup)** section at the top of this file for the
full step-by-step instructions.


---


## Running Phase 1 Standalone


`resume_generator.py` can be run directly to regenerate `sample_resume.docx`
from the hardcoded `SAMPLE_DATA` constant -- useful for verifying the template
without needing the Streamlit app or an API key:


```bash
python resume_generator.py
```


Output: `sample_resume.docx` in the same folder.


---


## Development Guidelines


### The Phase Boundary Rule
`resume_generator.py` is **frozen** (Phase 1). It is the single source of truth
for all DOCX formatting. Changes to the visual output of the resume must only
ever be made here, and those changes are automatically picked up by `app_v6.py`
because it imports `create_resume` at runtime.


Do not copy formatting logic into `app_v6.py`. Do not duplicate functions across files.


### Adding a new section to the resume
1. Add a new helper function in `resume_generator.py` (e.g. `insert_awards_list`)
2. Call it from `create_resume()` in the appropriate position
3. Add the new key to the data dict schema and the AI system prompt in `app_v6.py`
4. Update the `REQUIRED_KEYS` set in `app_v6.py` if the new key is mandatory
5. Test by running `resume_generator.py` directly with sample data first


### Adding a new icon to a section
1. Capture the desired icon positioning by placing it manually in a Word document
2. Unzip the .docx, copy the `<wp:anchor>` XML from `word/document.xml`
3. Save it as `assets/<SectionName>_0.xml`
4. Add `"SectionName": "SectionName_0.xml"` to `SECTION_ICON_MAP` in `resume_generator.py`


### Changing the AI model
Only one line needs to change in `app_v6.py`:


```python
MODEL = "gpt-4o-mini"   # change this value
```


Ensure the new model supports `response_format={"type": "json_object"}`. If it
does not, remove that parameter and add `"Return ONLY valid JSON."` to the end
of the system prompt.


### Changing font or colours
All constants are at the top of `resume_generator.py`:


```python
FONT_NAME = "Century Gothic"
BLUE = RGBColor(0x00, 0x66, 0xCC)
LINE_WEIGHT_PT = 2
```


### Adding a new input format
Add an extractor function in `app_v6.py` and register it in the `EXTRACTORS` dict:


```python
def extract_text_rtf(file_bytes: bytes) -> str:
    ...


EXTRACTORS["application/rtf"] = extract_text_rtf
```


Also add the extension to `st.file_uploader(..., type=["docx", "pdf", "txt", "rtf"])`.


### Streamlit theme changes
All CSS lives in the `WORKPACE_CSS` string constant at the top of `app_v6.py`.
It is injected once via `st.markdown(WORKPACE_CSS, unsafe_allow_html=True)`.
Edit it there -- no separate CSS file is needed.


### What NOT to do
- Do not hardcode candidate data anywhere in `app_v6.py`
- Do not modify `assets/*.xml` files manually -- they contain precise EMU coordinates
- Do not add new Python files for single-use utilities; keep logic in `app_v6.py` and `resume_generator.py`
- Do not create temporary scripts and leave them in the folder


---


## Troubleshooting


| Problem | Likely cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named streamlit` | Running system Python instead of venv | Use `.\\.venv\Scripts\python.exe -m streamlit run app_v6.py` |
| `Please enter your OpenAI API key` | No key entered in the UI and no `.env` fallback | Paste your `sk-...` key into the API key field at the top of the app |
| `AI response missing required keys` | Model returned incomplete JSON | Retry; if persistent, check the model name in `MODEL` constant |
| Generated DOCX has no icons | `assets/` folder missing or moved | Ensure `assets/` is in the same directory as `resume_generator.py` |
| PDF extraction returns empty text | Scanned/image-based PDF | Text extraction only works on machine-readable PDFs; use .txt or .docx instead |
| Port 8501 already in use | Another Streamlit instance running | Add `--server.port 8502` to the run command |
| Bulk analysis button is greyed out | Only one file uploaded | Upload two or more resumes to enable bulk mode |
| Excel export button missing | `openpyxl` not installed | Run `pip install openpyxl` in the venv |


---


## Environment Variables Reference


| Variable | Required | Description |
|---|---|---|
| `OPENAI_API_KEY` | No (optional fallback) | OpenAI secret key, starting with `sk-`. Only needed for local dev if you prefer not to enter the key in the UI each time. The UI-entered key always takes precedence. |


---


## Hosting (Streamlit Cloud)


The app is designed for **Bring Your Own Key (BYOK)** deployment:


1. Push the repo to GitHub (ensure `.env` is in `.gitignore`).
2. Deploy on [Streamlit Community Cloud](https://streamlit.io/cloud).
3. Each visitor pastes their own OpenAI key -- no shared secret is needed.


**Risk model:** BYOK avoids shared-credit abuse because the app never stores
or funds API calls itself. However, Streamlit Cloud is server-side Python, so
resume text and the user's API key are processed on the hosted server during
the request. For a portfolio / MVP demo this is acceptable. For a production
app handling sensitive data, move to a backend with authentication, rate
limiting, and managed secrets.


---


## Dependencies


| Package | Version pinned | Purpose |
|---|---|---|
| `streamlit` | no | Web UI framework |
| `openai` | `>=1.0` | OpenAI Python SDK |
| `pdfplumber` | no | PDF text extraction |
| `python-dotenv` | no | `.env` file loading |
| `python-docx` | no | DOCX reading (extraction) and writing (generation) |
| `lxml` | no | Low-level XML manipulation for DrawingML shapes |
