# Resume Generator


An AI-powered resume formatting tool. Upload any resume, and the app extracts its
content, uses GPT to rewrite and structure it, then outputs a professionally
formatted Microsoft Word document (.docx) matching the company template.


---


## Project Structure


```
Res Generator/
|-- resume_generator.py       Phase 1 -- DOCX generation engine (DO NOT MODIFY)
|-- app.py                    Phase 2 -- Streamlit web application
|-- vba_code_example .txt     Reference VBA macro (source of truth for formatting)
|-- sample_resume.docx        Test output generated from resume_generator.py
|-- .env                      Your API key (NOT committed to version control)
|-- .env.example              Template showing required environment variables
|-- requirements_phase2.txt   Python dependencies for Phase 2
|-- assets/                   Icon XML templates and EMF image used in DOCX output
|   |-- Experience_0.xml      Briefcase icon anchor for Experience section
|   |-- Projects_0.xml        Briefcase icon anchor for Projects section
|   |-- Education_0.xml       Briefcase icon anchor for Education section
|   |-- Skills_0.xml          Lightbulb icon anchor for Skills section
|   |-- image1.emf            Icon image embedded in generated DOCX
|-- .venv/                    Python virtual environment
```


---


## Architecture Overview


The project is split into two isolated phases:


```
[User uploads resume]
        |
        v
  app.py (Phase 2)
        |-- extract_text()         strip text from .docx / .pdf / .txt
        |-- call_openai()          send to gpt-4.5-preview, receive JSON
        |-- generate_resume_bytes()
                |
                v
        resume_generator.py (Phase 1)
                |-- create_resume(data)   build the DOCX from the JSON dict
                        |-- add_header_line()
                        |-- insert_name()
                        |-- insert_summary_inline()
                        |-- build_experience_section()
                        |-- build_projects_section()
                        |-- insert_section_header()  + icon injection
                        |-- insert_education_entry()
                        |-- insert_skills_list()
                        v
                [returns .docx bytes] --> st.download_button
```


**Phase 1 (`resume_generator.py`) is a stable, frozen module.  
Phase 2 (`app.py`) is the user-facing layer that drives it.**


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


## Phase 2 -- Streamlit Web Application


**File:** `app.py`


### What it does
1. Presents a WorkPace-themed web UI
2. Accepts uploaded resumes (.docx, .pdf, .txt)
3. Extracts raw text from the upload
4. Sends text to `gpt-4.5-preview` with a structured system prompt
5. Receives a JSON object matching the Phase 1 data schema
6. Calls `create_resume()` from Phase 1 to generate the DOCX
7. Serves the DOCX as a download


### Key functions


| Function | Purpose |
|---|---|
| `extract_text(uploaded_file)` | Routes to correct extractor by MIME type |
| `extract_text_docx(bytes)` | Extracts paragraphs from .docx via python-docx |
| `extract_text_pdf(bytes)` | Extracts text pages from PDF via pdfplumber |
| `extract_text_txt(bytes)` | Decodes plain text |
| `call_openai(resume_text)` | Calls gpt-4.5-preview, validates + returns JSON dict |
| `generate_resume_bytes(data)` | Wraps create_resume() to return bytes via temp file |
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


### Requirements
- Python 3.10 or later
- An OpenAI API key with access to `gpt-4.5-preview`
- Microsoft Word (to open generated .docx files)


### 1. Clone / copy the folder
Place the entire `Res Generator` folder on the target machine. Keep all files
intact -- do not move `assets/` or rename any files.


### 2. Create a virtual environment


```bash
cd "Res Generator"
python -m venv .venv
```


Activate it:
- Windows: `.\.venv\Scripts\Activate.ps1` (PowerShell) or `.\.venv\Scripts\activate.bat` (CMD)
- macOS/Linux: `source .venv/bin/activate`


### 3. Install dependencies


```bash
pip install -r requirements_phase2.txt
```


This installs: `streamlit`, `openai`, `pdfplumber`, `python-docx`, `lxml`, `python-dotenv`.


### 4. Set your API key


Copy `.env.example` to `.env`:


```bash
copy .env.example .env       # Windows
cp .env.example .env         # macOS/Linux
```


Open `.env` and replace the placeholder:


```
OPENAI_API_KEY=sk-your-actual-key-here
```


**Never commit `.env` to version control.**


### 5. Run the app


```bash
.\.venv\Scripts\python.exe -m streamlit run app.py   # Windows venv
# or, if venv is already activated:
streamlit run app.py
```


The app opens at `http://localhost:8501`.


### 6. Verify it works (optional smoke test)


From the activated venv:


```bash
python -c "from app import extract_text_txt, call_openai, generate_resume_bytes; print('Imports OK')"
```


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
ever be made here, and those changes are automatically picked up by `app.py`
because it imports `create_resume` at runtime.


Do not copy formatting logic into `app.py`. Do not duplicate functions across files.


### Adding a new section to the resume
1. Add a new helper function in `resume_generator.py` (e.g. `insert_awards_list`)
2. Call it from `create_resume()` in the appropriate position
3. Add the new key to the data dict schema and the AI system prompt in `app.py`
4. Update the `REQUIRED_KEYS` set in `app.py` if the new key is mandatory
5. Test by running `resume_generator.py` directly with sample data first


### Adding a new icon to a section
1. Capture the desired icon positioning by placing it manually in a Word document
2. Unzip the .docx, copy the `<wp:anchor>` XML from `word/document.xml`
3. Save it as `assets/<SectionName>_0.xml`
4. Add `"SectionName": "SectionName_0.xml"` to `SECTION_ICON_MAP` in `resume_generator.py`


### Changing the AI model
Only one line needs to change in `app.py`:


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
Add an extractor function in `app.py` and register it in the `EXTRACTORS` dict:


```python
def extract_text_rtf(file_bytes: bytes) -> str:
    ...


EXTRACTORS["application/rtf"] = extract_text_rtf
```


Also add the extension to `st.file_uploader(..., type=["docx", "pdf", "txt", "rtf"])`.


### Streamlit theme changes
All CSS lives in the `WORKPACE_CSS` string constant at the top of `app.py`.
It is injected once via `st.markdown(WORKPACE_CSS, unsafe_allow_html=True)`.
Edit it there -- no separate CSS file is needed.


### What NOT to do
- Do not hardcode candidate data anywhere in `app.py`
- Do not modify `assets/*.xml` files manually -- they contain precise EMU coordinates
- Do not add new Python files for single-use utilities; keep logic in the two main files
- Do not create temporary scripts and leave them in the folder


---


## Troubleshooting


| Problem | Likely cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named streamlit` | Running system Python instead of venv | Use `.\.venv\Scripts\python.exe -m streamlit run app.py` |
| `OPENAI_API_KEY is not set` | `.env` file missing or key is placeholder | Create `.env` from `.env.example`, add real key |
| `AI response missing required keys` | Model returned incomplete JSON | Retry; if persistent, check the model name in `MODEL` constant |
| Generated DOCX has no icons | `assets/` folder missing or moved | Ensure `assets/` is in the same directory as `resume_generator.py` |
| PDF extraction returns empty text | Scanned/image-based PDF | Text extraction only works on machine-readable PDFs; use .txt or .docx instead |
| Port 8501 already in use | Another Streamlit instance running | Add `--server.port 8502` to the run command |


---


## Environment Variables Reference


| Variable | Required | Description |
|---|---|---|
| `OPENAI_API_KEY` | Yes | OpenAI secret key, starting with `sk-` |


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



